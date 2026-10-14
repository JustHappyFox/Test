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

// Сообщает, сколько записей уже прочитано (чтобы окно не "зависало" на больших таблицах)
public delegate void LoadProgress(int done, int expected);

// Слово поиска, переведённое в байты кодировки таблицы
public class SearchWord {
    public byte[] Bytes;
    public byte[] Trimmed;      // то же число без нулей впереди, или null
    public bool WholeNumber;    // короткое число (дом, квартира) - ищется только целиком
}

// Примечания (memo) читаются из файла .FPT только когда нужны - для карточки и выгрузки
public class MemoReader : IDisposable {
    readonly FileStream fs;
    readonly int blockSize;
    readonly Encoding enc;

    MemoReader(FileStream fs, int blockSize, Encoding enc) {
        this.fs = fs;
        this.blockSize = blockSize;
        this.enc = enc;
    }

    public static MemoReader Open(string dbfPath, Encoding enc) {
        string path = Path.ChangeExtension(dbfPath, ".FPT");
        if (!File.Exists(path)) path = Path.ChangeExtension(dbfPath, ".fpt");
        if (!File.Exists(path)) return null;
        FileStream fs = null;
        try {
            fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            byte[] h = new byte[8];
            if (DbfTable.ReadUpTo(fs, h, 8) < 8) { fs.Close(); return null; }
            int size = (h[6] << 8) | h[7];
            return new MemoReader(fs, size > 0 ? size : 64, enc);
        } catch (IOException) {
            if (fs != null) fs.Close();
            return null;
        }
    }

    public string Read(int block) {
        long pos = (long)block * blockSize;
        if (pos + 8 > fs.Length) return "";
        fs.Seek(pos, SeekOrigin.Begin);
        byte[] h = new byte[8];
        if (DbfTable.ReadUpTo(fs, h, 8) < 8) return "";
        int type = (h[0] << 24) | (h[1] << 16) | (h[2] << 8) | h[3];
        long len = (uint)((h[4] << 24) | (h[5] << 16) | (h[6] << 8) | h[7]);
        if (type != 1) return "(объект)";
        len = Math.Min(len, Math.Min(fs.Length - pos - 8, 1 << 20));   // не больше 1 МБ на одно примечание
        if (len <= 0) return "";
        byte[] b = new byte[len];
        int got = DbfTable.ReadUpTo(fs, b, b.Length);
        return enc.GetString(b, 0, got).TrimEnd('\0', ' ', '\r', '\n');
    }

    public void Dispose() {
        fs.Close();
    }
}

// Таблица DBF (FoxPro / dBase). Файл открывается только на чтение.
// Записи хранятся в памяти как есть, байтами, и разбираются на поля только для показа -
// так таблица на 100 000 абонентов читается быстро и занимает в памяти немного.
public class DbfTable {
    public string FilePath;
    public List<DbfField> Fields = new List<DbfField>();
    public Encoding Enc;
    public int Count;     // сколько записей в памяти
    public int Total;     // сколько неудалённых записей в файле
    public int Deleted;

    const int BlockRecords = 4096;   // записи хранятся кусками, без одного огромного массива
    int recordLength;
    int lowerLength;      // запись в поисковой копии: поля через разделитель, чтобы "18" и "12" не слились в "1812"
    int[] lowerOffsets;
    readonly List<byte[]> data = new List<byte[]>();    // записи как в файле
    readonly List<byte[]> lower = new List<byte[]>();   // они же маленькими буквами, только текст - для поиска
    readonly byte[] lowerMap = new byte[256];

    public static DbfTable Load(string path, bool otherEncoding) {
        return Load(path, otherEncoding, null, null);
    }

    // query - если задан, в памяти остаются только подходящие записи (поиск во всех таблицах)
    public static DbfTable Load(string path, bool otherEncoding, string query, LoadProgress progress) {
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
            t.Enc = Encoding.GetEncoding(windows ? 1251 : 866);
            t.recordLength = recordLen;
            t.BuildLowerMap();

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

            // поисковая копия: только текстовые поля (двоичные числа и ссылки на примечания - пропуск),
            // после каждого поля - разделитель
            int fieldCount = t.Fields.Count;
            bool[] text = new bool[fieldCount];
            t.lowerOffsets = new int[fieldCount];
            int pos = 0;
            for (int j = 0; j < fieldCount; j++) {
                t.lowerOffsets[j] = pos;
                pos += t.Fields[j].Length + 1;
                text[j] = "CNFDL".IndexOf(t.Fields[j].Type) >= 0;
            }
            t.lowerLength = Math.Max(pos, 1);
            SearchWord[] filter = query == null ? null : t.PrepareWords(query);

            fs.Seek(headerLen, SeekOrigin.Begin);
            const int chunkRecords = 2048;
            byte[] chunk = new byte[chunkRecords * recordLen];
            byte[] low = new byte[t.lowerLength];
            int index = 0;
            bool end = false;
            while (index < count && !end) {
                int want = Math.Min(chunkRecords, count - index);
                int got = ReadUpTo(fs, chunk, want * recordLen) / recordLen;
                for (int r = 0; r < got; r++) {
                    int start = r * recordLen;
                    byte flag = chunk[start];
                    if (flag == 0x1A) { end = true; break; }          // конец файла
                    if (flag == (byte)'*') { t.Deleted++; continue; }  // удалённая запись
                    t.Total++;
                    for (int j = 0; j < fieldCount; j++) {
                        DbfField f = t.Fields[j];
                        int src = start + f.Offset, dst = t.lowerOffsets[j];
                        if (text[j]) for (int k = 0; k < f.Length; k++) low[dst + k] = t.lowerMap[chunk[src + k]];
                        else for (int k = 0; k < f.Length; k++) low[dst + k] = 1;
                        low[dst + f.Length] = 1;
                    }
                    if (filter == null || MatchesAll(low, 0, t.lowerLength, filter)) t.Append(chunk, start, low);
                }
                index += got;
                if (got < want) break;   // файл короче, чем написано в его заголовке
                if (progress != null) progress(index, count);
            }
        }
        return t;
    }

    void BuildLowerMap() {
        byte[] one = new byte[1];
        for (int b = 0; b < 256; b++) {
            lowerMap[b] = (byte)b;
            one[0] = (byte)b;
            string s = Enc.GetString(one);
            string l = s.ToLowerInvariant();
            if (l == s) continue;
            byte[] lb = Enc.GetBytes(l);
            if (lb.Length == 1 && Enc.GetString(lb) == l) lowerMap[b] = lb[0];
        }
    }

    void Append(byte[] src, int start, byte[] low) {
        int block = Count / BlockRecords;
        int slot = Count % BlockRecords;
        if (block == data.Count) {
            data.Add(new byte[BlockRecords * recordLength]);
            lower.Add(new byte[BlockRecords * lowerLength]);
        }
        Buffer.BlockCopy(src, start, data[block], slot * recordLength, recordLength);
        Buffer.BlockCopy(low, 0, lower[block], slot * lowerLength, lowerLength);
        Count++;
    }

    public object GetValue(int row, int field, MemoReader memo) {
        byte[] buf = data[row / BlockRecords];
        int start = (row % BlockRecords) * recordLength;
        return Parse(buf, start, Fields[field], Enc, memo);
    }

    public object[] GetRow(int row, MemoReader memo) {
        object[] values = new object[Fields.Count];
        for (int j = 0; j < values.Length; j++) values[j] = GetValue(row, j, memo);
        return values;
    }

    public MemoReader OpenMemo() {
        foreach (DbfField f in Fields) if (f.Type == 'M') return MemoReader.Open(FilePath, Enc);
        return null;
    }

    // ---- поиск ----

    public static string[] SplitWords(string query) {
        return query.ToLowerInvariant().Split(new char[] { ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
    }

    public SearchWord[] PrepareWords(string query) {
        string[] words = SplitWords(query);
        SearchWord[] result = new SearchWord[words.Length];
        for (int i = 0; i < words.Length; i++) {
            string w = words[i];
            // дата 03.03.2021 хранится в таблице как 20210303
            if (w.Length == 10 && w[2] == '.' && w[5] == '.' && IsDigits(w.Remove(5, 1).Remove(2, 1)))
                w = w.Substring(6, 4) + w.Substring(3, 2) + w.Substring(0, 2);
            SearchWord sw = new SearchWord();
            sw.Bytes = ToLowerBytes(w);
            if (IsDigits(w)) {
                sw.WholeNumber = w.Length <= 4;
                string trimmed = w.TrimStart('0');
                if (trimmed.Length > 0 && trimmed.Length < w.Length) sw.Trimmed = ToLowerBytes(trimmed);
            }
            result[i] = sw;
        }
        return result;
    }

    byte[] ToLowerBytes(string s) {
        byte[] b = Enc.GetBytes(s);
        for (int i = 0; i < b.Length; i++) b[i] = lowerMap[b[i]];
        return b;
    }

    // field = -1 - искать во всей записи, иначе только в этом поле
    public bool Matches(int row, SearchWord[] words, int field) {
        if (words.Length == 0) return true;
        byte[] buf = lower[row / BlockRecords];
        int start = (row % BlockRecords) * lowerLength;
        if (field < 0) return MatchesAll(buf, start, lowerLength, words);
        return MatchesAll(buf, start + lowerOffsets[field], Fields[field].Length, words);
    }

    static bool MatchesAll(byte[] buf, int from, int length, SearchWord[] words) {
        int end = from + length;
        foreach (SearchWord w in words) {
            bool ok = w.WholeNumber
                ? ContainsNumber(buf, from, end, w.Bytes) || (w.Trimmed != null && ContainsNumber(buf, from, end, w.Trimmed))
                : IndexOf(buf, from, end, w.Bytes) >= 0 || (w.Trimmed != null && IndexOf(buf, from, end, w.Trimmed) >= 0);
            if (!ok) return false;
        }
        return true;
    }

    static int IndexOf(byte[] buf, int from, int end, byte[] pattern) {
        int last = end - pattern.Length;
        byte first = pattern[0];
        for (int i = from; i <= last; i++) {
            if (buf[i] != first) continue;
            int k = 1;
            while (k < pattern.Length && buf[i + k] == pattern[k]) k++;
            if (k == pattern.Length) return i;
        }
        return -1;
    }

    // число целиком: "18" находит "18" и "18а", но не "118", не "180" и не копейки в "61.18"
    static bool ContainsNumber(byte[] buf, int from, int end, byte[] number) {
        int i = IndexOf(buf, from, end, number);
        while (i >= 0) {
            int after = i + number.Length;
            bool before = i == from || !(IsDigit(buf[i - 1]) || (IsPoint(buf[i - 1]) && i - 1 > from && IsDigit(buf[i - 2])));
            bool behind = after >= end || !(IsDigit(buf[after]) || (IsPoint(buf[after]) && after + 1 < end && IsDigit(buf[after + 1])));
            if (before && behind) return true;
            i = IndexOf(buf, i + 1, end, number);
        }
        return false;
    }

    static bool IsPoint(byte b) {
        return b == (byte)'.' || b == (byte)',';
    }

    static bool IsDigit(byte b) {
        return b >= (byte)'0' && b <= (byte)'9';
    }

    static bool IsDigits(string s) {
        if (s.Length == 0) return false;
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return true;
    }

    // ---- разбор полей ----

    static object Parse(byte[] buf, int start, DbfField f, Encoding enc, MemoReader memo) {
        int at = start + f.Offset;
        switch (f.Type) {
            case 'C':
                return enc.GetString(buf, at, f.Length).TrimEnd(' ', '\0');
            case 'N':
            case 'F': {
                string s = Encoding.ASCII.GetString(buf, at, f.Length).Trim(' ', '\0');
                if (s.Length == 0) return null;
                decimal d;
                if (decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
                return s;
            }
            case 'D': {
                string s = Encoding.ASCII.GetString(buf, at, Math.Min(8, f.Length)).Trim(' ', '\0');
                DateTime d;
                if (s.Length == 8 && DateTime.TryParseExact(s, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
                return null;
            }
            case 'L': {
                char c = (char)buf[at];
                if ("TtYy".IndexOf(c) >= 0) return "Да";
                if ("FfNn".IndexOf(c) >= 0) return "Нет";
                return null;
            }
            case 'M': {
                int block = MemoBlock(buf, at, f.Length);
                if (block <= 0) return null;
                return memo == null ? "(текст)" : memo.Read(block);
            }
            case 'G':
            case 'P':
                return MemoBlock(buf, at, f.Length) > 0 ? "(объект)" : null;
            case 'I':
                return f.Length == 4 ? (object)(decimal)BitConverter.ToInt32(buf, at) : null;
            case 'B':
                return f.Length == 8 ? (object)BitConverter.ToDouble(buf, at) : null;
            case 'Y':
                return f.Length == 8 ? (object)(BitConverter.ToInt64(buf, at) / 10000m) : null;
            case 'T': {
                if (f.Length != 8) return null;
                int day = BitConverter.ToInt32(buf, at);
                int ms = BitConverter.ToInt32(buf, at + 4);
                if (day <= 1721426) return null;   // 1721426 - юлианский день 01.01.0001
                try { return new DateTime(1, 1, 1).AddDays(day - 1721426).AddMilliseconds(ms); }
                catch (ArgumentOutOfRangeException) { return null; }
            }
            default:
                return enc.GetString(buf, at, f.Length).Trim(' ', '\0');
        }
    }

    static int MemoBlock(byte[] buf, int at, int length) {
        if (length == 4) return BitConverter.ToInt32(buf, at);
        string s = Encoding.ASCII.GetString(buf, at, length).Trim(' ', '\0');
        int b;
        return int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out b) ? b : 0;
    }

    static byte[] ReadBytes(Stream s, int n) {
        byte[] b = new byte[n];
        if (ReadUpTo(s, b, n) < n) throw new InvalidDataException("Файл повреждён или обрезан.");
        return b;
    }

    internal static int ReadUpTo(Stream s, byte[] b, int n) {
        int got = 0;
        while (got < n) {
            int r = s.Read(b, got, n - got);
            if (r <= 0) break;
            got += r;
        }
        return got;
    }

    // ---- показ и выгрузка ----

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
        using (MemoReader memo = t.OpenMemo())
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
                foreach (object v in t.GetRow(rows[i], memo)) {
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
    object[] sortKeys;   // значения столбца сортировки, пока идёт сортировка
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
            resizeColumns = true;
        }
        suspend = false;
    }

    internal void ApplyFilter() {
        if (suspend || loading) return;
        view.Clear();
        if (table != null) {
            SearchWord[] words = table.PrepareWords(searchBox.Text);
            if (words.Length > 0 || table.Count <= ShowAllLimit) {
                int field = fieldBox.SelectedIndex - 1;   // -1 - искать во всех полях
                for (int i = 0; i < table.Count; i++) {
                    if (table.Matches(i, words, field)) view.Add(i);
                }
                SortView();
            }
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
                ScanTable(results, name, query);
            }
        } finally {
            Enabled = true;
            results.Finish();
        }
    }

    void ScanTable(ScanResults results, string name, string query) {
        try {
            DbfTable t = DbfTable.Load(Path.Combine(folder, name), encodingBox.Checked, query,
                delegate(int done, int expected) {
                    Application.DoEvents();
                    if (results.Cancelled || IsDisposed) throw new OperationCanceledException();
                });
            if (t.Count > 0) results.AddResult(name, t.Count, t.Total);
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
        if (sortColumn < 0 || table == null || sortColumn >= table.Fields.Count || view.Count < 2) return;
        sortKeys = new object[table.Count];
        foreach (int row in view) sortKeys[row] = table.GetValue(row, sortColumn, null);
        view.Sort(new Comparison<int>(CompareRows));
        sortKeys = null;
    }

    int CompareRows(int a, int b) {
        int r = DbfTable.CompareValues(sortKeys[a], sortKeys[b]);
        if (sortDescending) r = -r;
        return r != 0 ? r : a.CompareTo(b);
    }

    void UpdateStatus() {
        if (table == null) { status.Text = ""; return; }
        if (DbfTable.SplitWords(searchBox.Text).Length == 0 && table.Count > ShowAllLimit) {
            status.Text = "В таблице " + table.Count + " записей. Впишите в \"Поиск\" адрес, номер или фамилию - покажу только подходящие. Например: заводская 18";
            return;
        }
        status.Text = "Показано записей: " + view.Count + " из " + table.Count
            + (table.Deleted > 0 ? " (удалённые записи не показываются: " + table.Deleted + ")" : "")
            + ".   Двойной щелчок или Enter по строке - карточка целиком.";
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        if (table == null || e.RowIndex >= view.Count || e.ColumnIndex >= table.Fields.Count) return;
        e.Value = DbfTable.Format(table.GetValue(view[e.RowIndex], e.ColumnIndex, null), false);
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
        object[] values;
        using (MemoReader memo = table.OpenMemo()) values = table.GetRow(view[viewRow], memo);
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
