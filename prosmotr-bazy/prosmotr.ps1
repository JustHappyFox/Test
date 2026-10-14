# Просмотр базы FoxPro (таблицы .dbf) - ТОЛЬКО ЧТЕНИЕ.
# Таблицы открываются только на чтение, в базу ничего не записывается.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @([System.Windows.Forms.Form].Assembly.Location, [System.Drawing.Point].Assembly.Location)

$code = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

// Код написан на C# 2.0, чтобы собираться и в старом PowerShell 2.0 (Windows 7).

public class DbfField {
    public string Name;
    public char Type;
    public int Offset;
    public int Length;
}

// Решает, оставить ли прочитанную запись (для поиска по всем таблицам без загрузки их целиком)
public delegate bool RowFilter(object[] values);

// Сообщает, сколько записей уже прочитано (чтобы окно не "зависало" на больших таблицах)
public delegate void LoadProgress(int done, int expected);

// Таблица DBF (FoxPro / dBase), прочитанная в память. Файл открывается только на чтение.
public class DbfTable {
    public string FilePath;
    public List<DbfField> Fields = new List<DbfField>();
    public List<object[]> Rows = new List<object[]>();
    public int Deleted;
    public int Total;   // сколько неудалённых записей прочитано

    const int MaxMemoFile = 64 * 1024 * 1024;

    public static DbfTable Load(string path, bool otherEncoding) {
        return Load(path, otherEncoding, null, null);
    }

    public static DbfTable Load(string path, bool otherEncoding, RowFilter filter, LoadProgress progress) {
        DbfTable t = new DbfTable();
        t.FilePath = path;
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16)) {
            byte[] head = ReadBytes(fs, 32);
            int count = BitConverter.ToInt32(head, 4);
            int headerLen = BitConverter.ToUInt16(head, 8);
            int recordLen = BitConverter.ToUInt16(head, 10);
            if (headerLen < 33 || recordLen < 1 || count < 0)
                throw new InvalidDataException("Это не таблица DBF.");

            // базы FoxPro для DOS хранят текст в кодировке 866, отмеченные как Windows - в 1251
            bool windows = head[29] == 0xC9;
            if (otherEncoding) windows = !windows;
            Encoding enc = Encoding.GetEncoding(windows ? 1251 : 866);

            byte[] desc = ReadBytes(fs, headerLen - 32);
            int offset = 1;
            for (int p = 0; p + 32 <= desc.Length && desc[p] != 0x0D; p += 32) {
                DbfField f = new DbfField();
                int n = 0;
                while (n < 11 && desc[p + n] != 0) n++;
                f.Name = Encoding.ASCII.GetString(desc, p, n).Trim();
                f.Type = char.ToUpperInvariant((char)desc[p + 11]);
                f.Length = desc[p + 16];
                if (f.Type == 'C') f.Length += desc[p + 17] << 8;   // длинные текстовые поля (Clipper)
                f.Offset = offset;
                offset += f.Length;
                if (f.Type != '0') t.Fields.Add(f);                  // '0' - служебное поле VFP
            }
            if (offset > recordLen)
                throw new InvalidDataException("Не удалось разобрать структуру таблицы.");

            byte[] memo = null;
            int blockSize = 0;
            foreach (DbfField f in t.Fields) {
                if (f.Type == 'M') { memo = LoadMemo(path, out blockSize); break; }
            }

            fs.Seek(headerLen, SeekOrigin.Begin);
            byte[] rec = new byte[recordLen];
            for (int i = 0; i < count; i++) {
                if (progress != null && i > 0 && i % 5000 == 0) progress(i, count);
                if (!ReadFull(fs, rec) || rec[0] == 0x1A) break;
                if (rec[0] == (byte)'*') { t.Deleted++; continue; }
                object[] values = new object[t.Fields.Count];
                for (int j = 0; j < values.Length; j++)
                    values[j] = Parse(rec, t.Fields[j], enc, memo, blockSize);
                t.Total++;
                if (filter == null || filter(values)) t.Rows.Add(values);
            }
        }
        return t;
    }

    // Текст записи для поиска: все поля подряд, маленькими буквами
    public static string SearchText(object[] values) {
        StringBuilder sb = new StringBuilder();
        foreach (object v in values) sb.Append(Format(v, true)).Append('\u0001');
        return sb.ToString().ToLowerInvariant();
    }

    public static string[] SplitWords(string query) {
        return query.ToLowerInvariant().Split(new char[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
    }

    // Все слова должны встретиться в тексте записи
    public static bool Matches(string text, string[] words) {
        foreach (string w in words) {
            if (!MatchesWord(text, w)) return false;
        }
        return true;
    }

    static bool MatchesWord(string text, string w) {
        if (!IsDigits(w)) return text.IndexOf(w, StringComparison.Ordinal) >= 0;
        string trimmed = w.TrimStart('0');
        // короткое число - дом или квартира - ищем только целиком: "18" не найдёт "118" и "180"
        if (w.Length <= 4)
            return ContainsNumber(text, w) || (trimmed.Length > 0 && trimmed.Length < w.Length && ContainsNumber(text, trimmed));
        // длинное - лицевой счёт, телефон: можно часть; с нулями впереди ищем и без них
        if (text.IndexOf(w, StringComparison.Ordinal) >= 0) return true;
        return trimmed.Length > 0 && trimmed.Length < w.Length && text.IndexOf(trimmed, StringComparison.Ordinal) >= 0;
    }

    static bool ContainsNumber(string text, string n) {
        int i = text.IndexOf(n, StringComparison.Ordinal);
        while (i >= 0) {
            bool before = i == 0 || text[i - 1] < '0' || text[i - 1] > '9';
            int end = i + n.Length;
            bool after = end >= text.Length || text[end] < '0' || text[end] > '9';
            if (before && after) return true;
            i = text.IndexOf(n, i + 1, StringComparison.Ordinal);
        }
        return false;
    }

    static bool IsDigits(string s) {
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return true;
    }

    static object Parse(byte[] rec, DbfField f, Encoding enc, byte[] memo, int blockSize) {
        switch (f.Type) {
            case 'C':
                return enc.GetString(rec, f.Offset, f.Length).TrimEnd(' ', '\0');
            case 'N':
            case 'F': {
                string s = Encoding.ASCII.GetString(rec, f.Offset, f.Length).Trim(' ', '\0');
                if (s.Length == 0) return null;
                decimal d;
                if (decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
                return s;
            }
            case 'D': {
                string s = Encoding.ASCII.GetString(rec, f.Offset, Math.Min(8, f.Length)).Trim(' ', '\0');
                DateTime d;
                if (s.Length == 8 && DateTime.TryParseExact(s, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
                return null;
            }
            case 'L': {
                char c = (char)rec[f.Offset];
                if ("TtYy".IndexOf(c) >= 0) return "Да";
                if ("FfNn".IndexOf(c) >= 0) return "Нет";
                return null;
            }
            case 'M': {
                int block = MemoBlock(rec, f);
                if (block <= 0) return null;
                if (memo == null) return "(текст)";
                return ReadMemo(memo, block, blockSize, enc);
            }
            case 'G':
            case 'P':
                return MemoBlock(rec, f) > 0 ? "(объект)" : null;
            case 'I':
                return f.Length == 4 ? (object)(decimal)BitConverter.ToInt32(rec, f.Offset) : null;
            case 'B':
                return f.Length == 8 ? (object)BitConverter.ToDouble(rec, f.Offset) : null;
            case 'Y':
                return f.Length == 8 ? (object)(BitConverter.ToInt64(rec, f.Offset) / 10000m) : null;
            case 'T': {
                if (f.Length != 8) return null;
                int day = BitConverter.ToInt32(rec, f.Offset);
                int ms = BitConverter.ToInt32(rec, f.Offset + 4);
                if (day <= 1721426) return null;   // 1721426 - юлианский день 01.01.0001
                try { return new DateTime(1, 1, 1).AddDays(day - 1721426).AddMilliseconds(ms); }
                catch (ArgumentOutOfRangeException) { return null; }
            }
            default:
                return enc.GetString(rec, f.Offset, f.Length).Trim(' ', '\0');
        }
    }

    static int MemoBlock(byte[] rec, DbfField f) {
        if (f.Length == 4) return BitConverter.ToInt32(rec, f.Offset);
        string s = Encoding.ASCII.GetString(rec, f.Offset, f.Length).Trim(' ', '\0');
        int b;
        return int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out b) ? b : 0;
    }

    // Файл .FPT рядом с таблицей; если его нет или он слишком большой - текст не читаем
    static byte[] LoadMemo(string path, out int blockSize) {
        blockSize = 0;
        string memoPath = Path.ChangeExtension(path, ".FPT");
        if (!File.Exists(memoPath)) memoPath = Path.ChangeExtension(path, ".fpt");
        if (!File.Exists(memoPath)) return null;
        try {
            using (FileStream fs = new FileStream(memoPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                if (fs.Length < 512 || fs.Length > MaxMemoFile) return null;
                byte[] data = ReadBytes(fs, (int)fs.Length);
                blockSize = (data[6] << 8) | data[7];
                if (blockSize <= 0) blockSize = 64;
                return data;
            }
        } catch (IOException) {
            return null;
        }
    }

    static string ReadMemo(byte[] memo, int block, int blockSize, Encoding enc) {
        long pos = (long)block * blockSize;
        if (pos + 8 > memo.Length) return "";
        int p = (int)pos;
        int type = BigEndian(memo, p);
        int len = BigEndian(memo, p + 4);
        if (type != 1) return "(объект)";
        if (len <= 0) return "";
        if (p + 8 + (long)len > memo.Length) len = memo.Length - p - 8;
        return enc.GetString(memo, p + 8, len).TrimEnd('\0', ' ', '\r', '\n');
    }

    static int BigEndian(byte[] b, int p) {
        return (b[p] << 24) | (b[p + 1] << 16) | (b[p + 2] << 8) | b[p + 3];
    }

    static byte[] ReadBytes(Stream s, int n) {
        byte[] b = new byte[n];
        if (!ReadFull(s, b)) throw new InvalidDataException("Файл повреждён или обрезан.");
        return b;
    }

    static bool ReadFull(Stream s, byte[] b) {
        int got = 0;
        while (got < b.Length) {
            int r = s.Read(b, got, b.Length - got);
            if (r <= 0) return false;
            got += r;
        }
        return true;
    }

    // Значение для показа: в таблице длинный текст обрезается, в карточке - целиком
    public static string Format(object v, bool full) {
        if (v == null) return "";
        if (v is DateTime) {
            DateTime d = (DateTime)v;
            return d.TimeOfDay == TimeSpan.Zero ? d.ToString("dd.MM.yyyy") : d.ToString("dd.MM.yyyy HH:mm:ss");
        }
        string s = v as string;
        if (s == null) return Convert.ToString(v, CultureInfo.CurrentCulture);
        if (full) return s;
        s = s.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
        return s.Length > 200 ? s.Substring(0, 200) + "..." : s;
    }

    public static int CompareValues(object a, object b) {
        if (a == null) return b == null ? 0 : 1;
        if (b == null) return -1;
        if (a is decimal && b is decimal) return ((decimal)a).CompareTo((decimal)b);
        if (a is double && b is double) return ((double)a).CompareTo((double)b);
        if (a is DateTime && b is DateTime) return ((DateTime)a).CompareTo((DateTime)b);
        return string.Compare(Format(a, true), Format(b, true), StringComparison.CurrentCultureIgnoreCase);
    }

    // Выгрузка в формате "Таблица XML 2003": Excel открывает его сам и не теряет нули в начале номеров
    public static void WriteSpreadsheet(DbfTable t, IList<int> rows, int count, string file) {
        using (StreamWriter w = new StreamWriter(file, false, new UTF8Encoding(false))) {
            w.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            w.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            w.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            w.WriteLine("<Styles><Style ss:ID=\"h\"><Font ss:Bold=\"1\"/></Style><Style ss:ID=\"d\"><NumberFormat ss:Format=\"Short Date\"/></Style></Styles>");
            w.WriteLine("<Worksheet ss:Name=\"" + Xml(Path.GetFileNameWithoutExtension(t.FilePath)) + "\"><Table>");
            w.Write("<Row>");
            foreach (DbfField f in t.Fields)
                w.Write("<Cell ss:StyleID=\"h\"><Data ss:Type=\"String\">" + Xml(f.Name) + "</Data></Cell>");
            w.WriteLine("</Row>");
            for (int i = 0; i < count; i++) {
                w.Write("<Row>");
                foreach (object v in t.Rows[rows[i]]) {
                    if (v == null)
                        w.Write("<Cell/>");
                    else if (v is decimal || (v is double && !double.IsNaN((double)v) && !double.IsInfinity((double)v)))
                        w.Write("<Cell><Data ss:Type=\"Number\">" + Convert.ToString(v, CultureInfo.InvariantCulture) + "</Data></Cell>");
                    else if (v is DateTime)
                        w.Write("<Cell ss:StyleID=\"d\"><Data ss:Type=\"DateTime\">" + ((DateTime)v).ToString("yyyy-MM-dd'T'HH:mm:ss.fff", CultureInfo.InvariantCulture) + "</Data></Cell>");
                    else
                        w.Write("<Cell><Data ss:Type=\"String\">" + Xml(Format(v, true)) + "</Data></Cell>");
                }
                w.WriteLine("</Row>");
            }
            w.WriteLine("</Table></Worksheet></Workbook>");
        }
    }

    static string Xml(string s) {
        StringBuilder sb = new StringBuilder(s.Length);
        foreach (char c in s) {
            switch (c) {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\n': sb.Append("&#10;"); break;
                case '\t': sb.Append(c); break;
                default:
                    if (c >= ' ') sb.Append(c);
                    break;
            }
        }
        return sb.ToString();
    }
}

public class DbfViewer : Form {
    readonly string settingsFile;
    string folder;
    readonly TextBox folderBox = new TextBox();
    readonly ComboBox tableBox = new ComboBox();
    readonly CheckBox encodingBox = new CheckBox();
    internal readonly TextBox searchBox = new TextBox();
    readonly ComboBox fieldBox = new ComboBox();
    readonly DataGridView grid = new DataGridView();
    readonly Label status = new Label();
    readonly Timer searchTimer = new Timer();

    DbfTable table;
    string[] rowText;
    readonly List<int> view = new List<int>();
    int sortColumn = -1;
    bool sortDescending;
    bool suspend;
    bool resizeColumns;
    bool loading;
    readonly FlowLayoutPanel bar1 = Bar();
    readonly FlowLayoutPanel bar2 = Bar();

    // больше стольки записей таблица целиком не показывается - только найденное
    const int ShowAllLimit = 5000;

    public static void Run(string settingsFile) {
        Application.EnableVisualStyles();
        Application.Run(new DbfViewer(settingsFile));
    }

    public DbfViewer(string settingsFile) {
        this.settingsFile = settingsFile;
        Text = "Просмотр базы (только чтение)";
        Size = new Size(1100, 700);
        StartPosition = FormStartPosition.CenterScreen;

        folderBox.Width = 420;
        folderBox.KeyDown += new KeyEventHandler(OnFolderKey);
        Button openButton = new Button();
        openButton.Text = "Открыть";
        openButton.AutoSize = true;
        openButton.Click += delegate { OpenTypedFolder(); };
        Button folderButton = new Button();
        folderButton.Text = "Обзор...";
        folderButton.AutoSize = true;
        folderButton.Click += delegate { ChooseFolder(); };
        bar1.Controls.AddRange(new Control[] { Caption("Папка:"), folderBox, openButton, folderButton });

        tableBox.DropDownStyle = ComboBoxStyle.DropDownList;
        tableBox.Width = 180;
        tableBox.SelectedIndexChanged += delegate { LoadSelected(); };
        Button reload = new Button();
        reload.Text = "Обновить";
        reload.AutoSize = true;
        reload.Click += delegate { LoadSelected(); };
        searchBox.Width = 260;
        searchBox.TextChanged += delegate { searchTimer.Stop(); searchTimer.Start(); };
        searchBox.KeyDown += new KeyEventHandler(OnSearchKey);
        fieldBox.DropDownStyle = ComboBoxStyle.DropDownList;
        fieldBox.Width = 140;
        fieldBox.SelectedIndexChanged += delegate { ApplyFilter(); };
        Button everywhere = new Button();
        everywhere.Text = "Во всех таблицах";
        everywhere.AutoSize = true;
        everywhere.Click += delegate { SearchAllTables(); };
        Button excel = new Button();
        excel.Text = "В Excel";
        excel.AutoSize = true;
        excel.Click += delegate { ExportToExcel(); };
        encodingBox.Text = "Другая кодировка";
        encodingBox.AutoSize = true;
        encodingBox.Margin = new Padding(12, 6, 3, 3);
        encodingBox.CheckedChanged += delegate { LoadSelected(); };
        bar2.Controls.AddRange(new Control[] {
            Caption("Таблица:"), tableBox, reload,
            Caption("Поиск:"), searchBox, everywhere, Caption("в поле:"), fieldBox, excel, encodingBox });

        grid.Dock = DockStyle.Fill;
        grid.VirtualMode = true;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.CellValueNeeded += new DataGridViewCellValueEventHandler(OnCellValueNeeded);
        grid.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(OnHeaderClick);
        grid.CellDoubleClick += new DataGridViewCellEventHandler(OnCellDoubleClick);
        grid.KeyDown += new KeyEventHandler(OnGridKey);

        status.Dock = DockStyle.Bottom;
        status.Height = 22;
        status.TextAlign = ContentAlignment.MiddleLeft;

        searchTimer.Interval = 300;
        searchTimer.Tick += delegate { searchTimer.Stop(); ApplyFilter(); };

        Controls.Add(grid);
        Controls.Add(status);
        Controls.Add(bar2);
        Controls.Add(bar1);

        string lastTable = ReadSetting(1);
        Shown += delegate {
            Cursor = Cursors.WaitCursor;
            status.Text = "Ищу папку базы...";
            status.Refresh();
            folder = FindFolder(ReadSetting(0));
            Cursor = Cursors.Default;
            FillTables(lastTable);
            searchBox.Focus();
        };
    }

    // Батник базы подключает \\fs-dos\communal как диск O:, но только пока база запущена,
    // поэтому пробуем и диск, и прямой сетевой путь
    string FindFolder(string saved) {
        string[] candidates = { saved, @"O:\BASAFOX", @"\\fs-dos\communal\BASAFOX" };
        foreach (string c in candidates) {
            try { if (c != null && Directory.Exists(c)) return c; }
            catch (Exception) { }
        }
        return saved != null ? saved : @"\\fs-dos\communal\BASAFOX";
    }

    void OnFolderKey(object sender, KeyEventArgs e) {
        if (e.KeyCode != Keys.Enter) return;
        e.Handled = true;
        e.SuppressKeyPress = true;
        OpenTypedFolder();
    }

    void OpenTypedFolder() {
        string p = folderBox.Text.Trim();
        if (p.Length == 0) return;
        folder = p;
        SaveSettings();
        FillTables(null);
    }

    static FlowLayoutPanel Bar() {
        FlowLayoutPanel p = new FlowLayoutPanel();
        p.Dock = DockStyle.Top;
        p.AutoSize = true;
        p.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        p.Padding = new Padding(4, 4, 4, 0);
        return p;
    }

    static Label Caption(string text) {
        Label l = new Label();
        l.Text = text;
        l.AutoSize = true;
        l.Margin = new Padding(10, 8, 0, 3);
        return l;
    }

    void ChooseFolder() {
        FolderBrowserDialog d = new FolderBrowserDialog();
        d.Description = "Папка с таблицами базы (файлы .dbf)";
        if (folder != null && Directory.Exists(folder)) d.SelectedPath = folder;
        if (d.ShowDialog(this) != DialogResult.OK) return;
        folder = d.SelectedPath;
        SaveSettings();
        FillTables(null);
    }

    void FillTables(string select) {
        folderBox.Text = folder;
        suspend = true;
        tableBox.Items.Clear();
        suspend = false;
        string[] files = new string[0];
        try { files = Directory.GetFiles(folder, "*.dbf"); }
        catch (Exception ex) { Message("Не удалось открыть папку " + folder + "\n\n" + ex.Message + "\n\nПроверьте путь в поле \"Папка\" и нажмите \"Открыть\"."); }
        Array.Sort(files, StringComparer.OrdinalIgnoreCase);
        int index = -1;
        foreach (string f in files) {
            string name = Path.GetFileName(f);
            if (select != null && string.Compare(name, select, StringComparison.OrdinalIgnoreCase) == 0) index = tableBox.Items.Count;
            tableBox.Items.Add(name);
        }
        table = null;
        BuildColumns();
        if (files.Length == 0) status.Text = "В этой папке нет таблиц .dbf. Впишите другой путь или нажмите \"Обзор...\".";
        else if (index >= 0) tableBox.SelectedIndex = index;
        else status.Text = "Выберите таблицу в списке.";
    }

    void LoadSelected() {
        if (suspend || loading || tableBox.SelectedItem == null) return;
        string name = (string)tableBox.SelectedItem;
        table = null;
        BuildColumns();
        SetBusy(true);
        status.Text = "Загрузка " + name + "...";
        status.Refresh();
        try {
            table = DbfTable.Load(Path.Combine(folder, name), encodingBox.Checked, null, delegate(int done, int expected) {
                if (IsDisposed) throw new OperationCanceledException();
                status.Text = "Загрузка " + name + ": прочитано " + done + " из " + expected + " записей...";
                Application.DoEvents();
            });
        } catch (OperationCanceledException) {
            return;
        } catch (IOException ex) {
            table = null;
            Message("Не удалось открыть таблицу (возможно, она сейчас занята):\n" + ex.Message);
        } catch (Exception ex) {
            table = null;
            Message("Не удалось прочитать таблицу:\n" + ex.Message);
        } finally {
            if (!IsDisposed) SetBusy(false);
        }
        SaveSettings();
        BuildColumns();
        ApplyFilter();
        searchBox.Focus();
    }

    void SetBusy(bool busy) {
        loading = busy;
        bar1.Enabled = !busy;
        bar2.Enabled = !busy;
        grid.Enabled = !busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    void BuildColumns() {
        suspend = true;
        grid.RowCount = 0;
        grid.Columns.Clear();
        fieldBox.Items.Clear();
        fieldBox.Items.Add("все поля");
        fieldBox.SelectedIndex = 0;
        sortColumn = -1;
        rowText = null;
        view.Clear();
        if (table != null) {
            foreach (DbfField f in table.Fields) {
                DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
                c.HeaderText = f.Name;
                c.SortMode = DataGridViewColumnSortMode.Programmatic;
                if ("NFIBY".IndexOf(f.Type) >= 0) c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                grid.Columns.Add(c);
                fieldBox.Items.Add(f.Name);
            }
            rowText = new string[table.Rows.Count];
            for (int i = 0; i < rowText.Length; i++) rowText[i] = DbfTable.SearchText(table.Rows[i]);
            resizeColumns = true;
        }
        suspend = false;
    }

    internal void ApplyFilter() {
        if (suspend || loading) return;
        view.Clear();
        string[] words = DbfTable.SplitWords(searchBox.Text);
        if (table != null && (words.Length > 0 || table.Rows.Count <= ShowAllLimit)) {
            int field = fieldBox.SelectedIndex - 1;   // -1 - искать во всех полях
            for (int i = 0; i < table.Rows.Count; i++) {
                string text = field < 0 ? rowText[i] : DbfTable.Format(table.Rows[i][field], true).ToLowerInvariant();
                if (DbfTable.Matches(text, words)) view.Add(i);
            }
            SortView();
        }
        grid.RowCount = 0;
        grid.RowCount = view.Count;
        grid.Invalidate();
        if (resizeColumns && grid.Columns.Count > 0 && view.Count > 0) {   // ширина - по первым найденным строкам
            resizeColumns = false;
            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            foreach (DataGridViewColumn c in grid.Columns) if (c.Width > 300) c.Width = 300;
        }
        UpdateStatus();
    }

    // Поиск по всем таблицам папки: таблицы читаются по очереди, в памяти остаются только совпадения
    internal void SearchAllTables() {
        string query = searchBox.Text.Trim();
        string[] words = DbfTable.SplitWords(query);
        if (words.Length == 0) {
            Message("Впишите в поле \"Поиск\", что искать: адрес, номер, фамилию...");
            searchBox.Focus();
            return;
        }
        if (tableBox.Items.Count == 0) {
            Message("В папке нет таблиц .dbf.");
            return;
        }
        searchTimer.Stop();
        ScanResults results = new ScanResults(this, query);
        results.Show(this);
        Enabled = false;
        try {
            int count = tableBox.Items.Count;
            for (int i = 0; i < count && !results.Cancelled; i++) {
                string name = (string)tableBox.Items[i];
                results.SetProgress("Проверяю " + name + " (" + (i + 1) + " из " + count + ")...");
                ScanTable(results, name, words);
            }
        } finally {
            Enabled = true;
            results.Finish();
        }
    }

    void ScanTable(ScanResults results, string name, string[] words) {
        try {
            DbfTable t = DbfTable.Load(Path.Combine(folder, name), encodingBox.Checked,
                delegate(object[] values) { return DbfTable.Matches(DbfTable.SearchText(values), words); },
                delegate(int done, int expected) {
                    Application.DoEvents();
                    if (results.Cancelled || IsDisposed) throw new OperationCanceledException();
                });
            if (t.Rows.Count > 0) results.AddResult(name, t.Rows.Count, t.Total);
        } catch (OperationCanceledException) {
        } catch (Exception ex) {
            results.AddError(name, ex.Message);
        }
        Application.DoEvents();
    }

    // Открыть таблицу из результатов поиска сразу с тем же поиском
    internal void OpenTable(string name, string query) {
        int index = -1;
        for (int i = 0; i < tableBox.Items.Count; i++) {
            if (string.Compare((string)tableBox.Items[i], name, StringComparison.OrdinalIgnoreCase) == 0) { index = i; break; }
        }
        if (index < 0) return;
        searchBox.Text = query;
        searchTimer.Stop();
        if (tableBox.SelectedIndex == index) ApplyFilter();
        else tableBox.SelectedIndex = index;
        Activate();
    }

    void SortView() {
        if (sortColumn < 0 || table == null || sortColumn >= table.Fields.Count) return;
        view.Sort(new Comparison<int>(CompareRows));
    }

    int CompareRows(int a, int b) {
        int r = DbfTable.CompareValues(table.Rows[a][sortColumn], table.Rows[b][sortColumn]);
        if (sortDescending) r = -r;
        return r != 0 ? r : a.CompareTo(b);
    }

    void UpdateStatus() {
        if (table == null) { status.Text = ""; return; }
        if (DbfTable.SplitWords(searchBox.Text).Length == 0 && table.Rows.Count > ShowAllLimit) {
            status.Text = "В таблице " + table.Rows.Count + " записей. Впишите в \"Поиск\" адрес, номер или фамилию - покажу только подходящие. Например: заводская 18";
            return;
        }
        status.Text = "Показано записей: " + view.Count + " из " + table.Rows.Count
            + (table.Deleted > 0 ? " (удалённые записи не показываются: " + table.Deleted + ")" : "")
            + ".   Двойной щелчок или Enter по строке - карточка целиком.";
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        if (table == null || e.RowIndex >= view.Count || e.ColumnIndex >= table.Fields.Count) return;
        e.Value = DbfTable.Format(table.Rows[view[e.RowIndex]][e.ColumnIndex], false);
    }

    void OnHeaderClick(object sender, DataGridViewCellMouseEventArgs e) {
        if (table == null || e.ColumnIndex < 0) return;
        if (e.ColumnIndex == sortColumn) {
            sortDescending = !sortDescending;
        } else {
            if (sortColumn >= 0 && sortColumn < grid.Columns.Count)
                grid.Columns[sortColumn].HeaderCell.SortGlyphDirection = SortOrder.None;
            sortColumn = e.ColumnIndex;
            sortDescending = false;
        }
        grid.Columns[sortColumn].HeaderCell.SortGlyphDirection = sortDescending ? SortOrder.Descending : SortOrder.Ascending;
        Cursor = Cursors.WaitCursor;
        SortView();
        Cursor = Cursors.Default;
        grid.Invalidate();
    }

    void OnCellDoubleClick(object sender, DataGridViewCellEventArgs e) {
        if (e.RowIndex >= 0) ShowCard(e.RowIndex);
    }

    void OnGridKey(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter && grid.CurrentCell != null) {
            e.Handled = true;
            e.SuppressKeyPress = true;
            ShowCard(grid.CurrentCell.RowIndex);
        }
    }

    void OnSearchKey(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) {
            e.Handled = true;
            e.SuppressKeyPress = true;
            searchTimer.Stop();
            if (e.Control) SearchAllTables();   // Ctrl+Enter - искать во всех таблицах
            else ApplyFilter();
        } else if (e.KeyCode == Keys.Down && grid.RowCount > 0) {
            e.Handled = true;
            grid.Focus();
        }
    }

    internal void ShowCard(int viewRow) {
        if (table == null || viewRow < 0 || viewRow >= view.Count) return;
        object[] values = table.Rows[view[viewRow]];
        Form card = new Form();
        card.Text = "Карточка - " + Path.GetFileName(table.FilePath);
        card.Size = new Size(650, 600);
        card.StartPosition = FormStartPosition.CenterScreen;
        DataGridView g = new DataGridView();
        g.Dock = DockStyle.Fill;
        g.ReadOnly = true;
        g.AllowUserToAddRows = false;
        g.AllowUserToDeleteRows = false;
        g.RowHeadersVisible = false;
        g.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        g.ColumnCount = 2;
        g.Columns[0].HeaderText = "Поле";
        g.Columns[1].HeaderText = "Значение";
        g.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        g.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        g.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        g.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        for (int j = 0; j < values.Length; j++)
            g.Rows.Add(table.Fields[j].Name, DbfTable.Format(values[j], true));
        card.Controls.Add(g);
        card.Show(this);
    }

    void ExportToExcel() {
        if (table == null || view.Count == 0) { Message("Нечего выгружать: таблица пуста или ничего не найдено."); return; }
        const int limit = 65000;
        if (view.Count > limit) {
            string q = "Найдено " + view.Count + " записей, в Excel будут выгружены первые " + limit
                + ".\nМожно сначала сузить список поиском. Продолжить?";
            if (MessageBox.Show(this, q, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        }
        string name = Path.GetFileNameWithoutExtension(table.FilePath);
        string file = Path.Combine(Path.GetTempPath(), name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xml");
        Cursor = Cursors.WaitCursor;
        try {
            DbfTable.WriteSpreadsheet(table, view, Math.Min(view.Count, limit), file);
        } catch (Exception ex) {
            Message("Не удалось сохранить файл для Excel:\n" + ex.Message);
            return;
        } finally {
            Cursor = Cursors.Default;
        }
        try {
            Process.Start("excel.exe", "\"" + file + "\"");
        } catch {
            try { Process.Start(file); }
            catch (Exception ex) { Message("Файл сохранён, но Excel не открылся сам:\n" + file + "\n\n" + ex.Message); }
        }
    }

    string ReadSetting(int line) {
        try {
            if (!File.Exists(settingsFile)) return null;
            string[] lines = File.ReadAllLines(settingsFile, Encoding.UTF8);
            return line < lines.Length && lines[line].Trim().Length > 0 ? lines[line].Trim() : null;
        } catch (Exception) {
            return null;
        }
    }

    void SaveSettings() {
        try {
            string current = tableBox.SelectedItem as string;
            File.WriteAllLines(settingsFile, new string[] { folder, current == null ? "" : current }, Encoding.UTF8);
        } catch (Exception) {
            // папка скрипта может быть только для чтения - тогда просто не запоминаем
        }
    }

    void Message(string text) {
        if (IsDisposed) return;
        MessageBox.Show(this, text, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}

// Окно с результатами поиска по всем таблицам: где и сколько нашлось
public class ScanResults : Form {
    readonly DbfViewer viewer;
    readonly string query;
    readonly ListView list = new ListView();
    readonly Label info = new Label();
    readonly Button stop = new Button();
    public bool Cancelled;
    bool scanning = true;
    int found;

    public ScanResults(DbfViewer viewer, string query) {
        this.viewer = viewer;
        this.query = query;
        Text = "Поиск во всех таблицах: " + query;
        Size = new Size(560, 500);
        StartPosition = FormStartPosition.CenterScreen;

        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.Columns.Add("Таблица", 180);
        list.Columns.Add("Найдено", 90, HorizontalAlignment.Right);
        list.Columns.Add("Всего записей", 260);
        list.ItemActivate += delegate { OpenSelected(); };

        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 34;
        info.Dock = DockStyle.Fill;
        info.TextAlign = ContentAlignment.MiddleLeft;
        stop.Text = "Стоп";
        stop.Dock = DockStyle.Right;
        stop.Width = 90;
        stop.Click += delegate { Cancelled = true; };
        bottom.Controls.Add(info);
        bottom.Controls.Add(stop);

        Controls.Add(list);
        Controls.Add(bottom);
        FormClosing += delegate { Cancelled = true; };
    }

    public void SetProgress(string text) {
        if (IsDisposed) return;
        info.Text = text;
        info.Refresh();
    }

    public void AddResult(string table, int hits, int total) {
        if (IsDisposed) return;
        list.Items.Add(new ListViewItem(new string[] { table, hits.ToString(), total.ToString() }));
        found++;
    }

    public void AddError(string table, string error) {
        if (IsDisposed) return;
        ListViewItem item = new ListViewItem(new string[] { table, "не открылась", error });
        item.ForeColor = Color.Gray;
        item.Tag = "error";
        list.Items.Add(item);
    }

    public void Finish() {
        scanning = false;
        if (IsDisposed) return;
        stop.Enabled = false;
        info.Text = (Cancelled ? "Поиск остановлен. " : "Поиск закончен. ")
            + (found == 0 ? "Ничего не найдено." : "Таблиц с совпадениями: " + found + ". Двойной щелчок - открыть таблицу с найденным.");
    }

    void OpenSelected() {
        if (scanning) { info.Text = "Дождитесь конца поиска или нажмите \"Стоп\"."; return; }
        if (list.SelectedItems.Count == 0 || list.SelectedItems[0].Tag != null) return;
        viewer.OpenTable(list.SelectedItems[0].Text, query);
    }
}
'@

# Перед запуском Windows собирает программу и пишет при этом временные файлы в папку TEMP.
# Если туда записать не получилось (антивирус, переполненная папка), пробуем другие папки.
$tempDirs = @((Join-Path $here 'tmp'), 'D:\temp')
if ($env:PUBLIC) { $tempDirs += (Join-Path $env:PUBLIC 'prosmotr-tmp') }
$attempt = 0
while ($true) {
    try {
        Add-Type -ReferencedAssemblies $refs -TypeDefinition $code
        break
    } catch {
        if ($attempt -ge $tempDirs.Count) {
            [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n$_`n`nПопробуйте запустить zapusk.bat ещё раз.", 'Просмотр базы - ошибка')
            exit
        }
        $dir = $tempDirs[$attempt]
        $attempt++
        try {
            if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
            $env:TMP = $dir
            $env:TEMP = $dir
        } catch { }
        Start-Sleep -Seconds 1
    }
}

try {
    [DbfViewer]::Run((Join-Path $here 'nastroiki.txt'))
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("$_", 'Просмотр базы - ошибка')
}
