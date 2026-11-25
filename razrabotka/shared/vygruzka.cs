// ==== ОБЩИЙ БЛОК ВЫГРУЗОК (исходник - razrabotka/shared/vygruzka.cs; здесь не править - вставляется скриптом inject.py) ====
// Источник данных для «Поиска» и «Договора»: база по-старому (O:\BASAFOX, индексы FoxPro) или Excel-выгрузки (быт, юр, наряды).
// Настройки - общие для меню и программ: %APPDATA%\Programmy-Energosbyt\istochnik.txt (меню в папки программ не пишет).
// Выгрузка быта один раз переводится в свою быструю копию (таблица DBF + справочник улиц) - это «свои индексы»:
// %LOCALAPPDATA%\Programmy-Energosbyt\vygruzka\. Заменили файл выгрузки - «Обновить индексы» в настройках меню.
// Только чтение выгрузок: сами файлы Excel не меняются.

public delegate void VygProgress(string text);

public static class Istochnik {
    public const string Base = "база", ExcelMode = "excel";

    static string Profile(Environment.SpecialFolder f) {
        string p = "";
        try { p = Environment.GetFolderPath(f); } catch { }
        if (p == null || p.Length == 0) p = Path.GetTempPath();
        return Path.Combine(p, "Programmy-Energosbyt");
    }

    public static string SettingsFile { get { return Path.Combine(Profile(Environment.SpecialFolder.ApplicationData), "istochnik.txt"); } }
    public static string CacheFolder { get { return Path.Combine(Profile(Environment.SpecialFolder.LocalApplicationData), "vygruzka"); } }
    public static string BytTable { get { return Path.Combine(CacheFolder, "byt.dbf"); } }
    public static string BytStreets { get { return Path.Combine(CacheFolder, "byt_ulicy.txt"); } }
    public static string BytInfo { get { return Path.Combine(CacheFolder, "byt_info.txt"); } }

    // настройки: Режим=база|excel, Быт=, Юр=, Наряды=
    public static Dictionary<string, string> Read() {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try {
            if (File.Exists(SettingsFile))
                foreach (string line in File.ReadAllLines(SettingsFile, Encoding.UTF8)) {
                    int eq = line.IndexOf('=');
                    if (eq > 0 && !line.StartsWith("#")) d[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
                }
        } catch { }
        return d;
    }

    static string Get(Dictionary<string, string> d, string k) {
        string v;
        return d.TryGetValue(k, out v) && v != null ? v : "";
    }

    public static bool UseExcel { get { return Get(Read(), "Режим") == ExcelMode; } }
    public static string BytPath { get { return Get(Read(), "Быт"); } }
    public static string JurPath { get { return Get(Read(), "Юр"); } }
    public static string NaryadyPath { get { return Get(Read(), "Наряды"); } }

    public static void Save(bool excel, string byt, string jur, string naryady) {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsFile));
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("# Откуда «Поиск» и «Договор» берут данные. Меняется в меню программ: «Настройки…»");
        sb.AppendLine("Режим=" + (excel ? ExcelMode : Base));
        sb.AppendLine("Быт=" + byt.Trim());
        sb.AppendLine("Юр=" + jur.Trim());
        sb.AppendLine("Наряды=" + naryady.Trim());
        string tmp = SettingsFile + ".tmp";
        File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(true));
        if (File.Exists(SettingsFile)) File.Delete(SettingsFile);
        File.Move(tmp, SettingsFile);
    }

    // какой выгрузки не хватает: null - всё в порядке, иначе - что сказать пользователю
    public static string Missing(string kind, string path) {
        if (path == null || path.Trim().Length == 0)
            return "Не указан путь к выгрузке " + kind + ". Укажите его в меню программ: «Настройки…».";
        if (!File.Exists(path.Trim()))
            return "Нет файла выгрузки " + kind + ": " + path.Trim() + ". Проверьте путь в меню программ: «Настройки…».";
        return null;
    }

    // своя копия выгрузки быта: null - готова и свежая; иначе - что не так
    public static string BytCacheProblem() {
        string path = BytPath;
        string miss = Missing("быта", path);
        if (miss != null) return miss;
        if (!File.Exists(BytTable) || !File.Exists(BytStreets) || !File.Exists(BytInfo))
            return "Индексы выгрузки быта ещё не созданы. Меню программ: «Настройки…» -> «Обновить индексы».";
        Dictionary<string, string> info = ReadInfo(BytInfo);
        if (string.Compare(Get(info, "Файл"), Path.GetFullPath(path.Trim()), StringComparison.OrdinalIgnoreCase) != 0)
            return "Индексы созданы по другому файлу выгрузки быта. Меню программ: «Настройки…» -> «Обновить индексы».";
        if (Get(info, "Версия") != VygruzkaByt.Version)
            return "Индексы выгрузки быта созданы прежней версией программы. Меню программ: «Настройки…» -> «Обновить индексы».";
        try {
            FileInfo fi = new FileInfo(path.Trim());
            if (Get(info, "Размер") != fi.Length.ToString() || Get(info, "Изменён") != fi.LastWriteTimeUtc.Ticks.ToString())
                return "Файл выгрузки быта заменён после создания индексов. Меню программ: «Настройки…» -> «Обновить индексы».";
        } catch (Exception ex) { return "Выгрузка быта недоступна: " + ex.Message; }
        return null;
    }

    public static Dictionary<string, string> ReadInfo(string file) {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try {
            foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                int eq = line.IndexOf('=');
                if (eq > 0) d[line.Substring(0, eq).Trim()] = line.Substring(eq + 1);
            }
        } catch { }
        return d;
    }

    // «от 14.01.2026 10:02 (152 340 абонентов)» - для надписей в окнах
    public static string BytAbout() {
        Dictionary<string, string> info = ReadInfo(BytInfo);
        string when = Get(info, "ФайлДата"), rows = Get(info, "Строк");
        return "выгрузка быта" + (when.Length > 0 ? " от " + when : "") + (rows.Length > 0 ? " (" + rows + " записей)" : "");
    }
}

// Вид выгрузки: имя (для надписей), ключевой столбец (строки без него пропускаются), файлы своей копии
public class VygKind {
    public string Code, Name;
    public string[] Keys;
    public string Table, Info, Streets;

    public static VygKind Byt { get { return Make("byt", "быта", new string[] { "НОМЕР" }, true); } }
    public static VygKind Nar { get { return Make("nar", "нарядов", new string[] { "Н_З", "НОМЕР_НЗ", "N_NAR" }, false); } }
    public static VygKind Jur { get { return Make("jur", "юр лиц", new string[] { "KP" }, false); } }

    static VygKind Make(string code, string name, string[] keys, bool streets) {
        VygKind k = new VygKind();
        k.Code = code;
        k.Name = name;
        k.Keys = keys;
        k.Table = Path.Combine(Istochnik.CacheFolder, code + ".dbf");
        k.Info = Path.Combine(Istochnik.CacheFolder, code + "_info.txt");
        k.Streets = streets ? Path.Combine(Istochnik.CacheFolder, code + "_ulicy.txt") : null;
        return k;
    }

    public string Path0 { get { return Code == "byt" ? Istochnik.BytPath : Code == "nar" ? Istochnik.NaryadyPath : Istochnik.JurPath; } }

    // своя копия выгрузки по файлу path: null - готова и свежая; иначе - что не так (юр база в .dbf копии не требует)
    public string Problem(string path) {
        string miss = Istochnik.Missing(Name, path);
        if (miss != null) return miss;
        path = path.Trim();
        if (Code == "jur" && path.EndsWith(".dbf", StringComparison.OrdinalIgnoreCase)) return null;
        if (!File.Exists(Table) || !File.Exists(Info) || (Streets != null && !File.Exists(Streets)))
            return "Индексы выгрузки " + Name + " ещё не созданы. Меню программ: «Настройки…» -> «Обновить индексы».";
        Dictionary<string, string> info = Istochnik.ReadInfo(Info);
        string v;
        if (string.Compare(info.TryGetValue("Файл", out v) ? v : "", Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase) != 0)
            return "Индексы созданы по другому файлу выгрузки " + Name + ". Меню программ: «Настройки…» -> «Обновить индексы».";
        if ((info.TryGetValue("Версия", out v) ? v : "") != VygruzkaByt.Version)
            return "Индексы выгрузки " + Name + " созданы прежней версией программы. Меню программ: «Настройки…» -> «Обновить индексы».";
        try {
            FileInfo fi = new FileInfo(path);
            if ((info.TryGetValue("Размер", out v) ? v : "") != fi.Length.ToString() || (info.TryGetValue("Изменён", out v) ? v : "") != fi.LastWriteTimeUtc.Ticks.ToString())
                return "Файл выгрузки " + Name + " заменён после создания индексов. Меню программ: «Настройки…» -> «Обновить индексы».";
        } catch (Exception ex) { return "Выгрузка " + Name + " недоступна: " + ex.Message; }
        return null;
    }

    public string Problem() { return Problem(Path0); }

    // нужна ли своя копия (юр база в .dbf - читается как есть)
    public bool NeedsCopy(string path) { return !(Code == "jur" && (path ?? "").Trim().EndsWith(".dbf", StringComparison.OrdinalIgnoreCase)); }

    // таблица для чтения: своя копия или сам .dbf юр базы
    public string TableFor(string path) { return NeedsCopy(path) ? Table : path.Trim(); }

    // «выгрузка нарядов от 14.01.2026 10:02 (152 340 записей)»
    public string About() {
        Dictionary<string, string> info = Istochnik.ReadInfo(Info);
        string when, rows;
        info.TryGetValue("ФайлДата", out when);
        info.TryGetValue("Строк", out rows);
        if (!NeedsCopy(Path0)) { try { when = File.GetLastWriteTime(Path0.Trim()).ToString("dd.MM.yyyy HH:mm"); rows = null; } catch { } }
        return "выгрузка " + Name + (when != null && when.Length > 0 ? " от " + when : "") + (rows != null && rows.Length > 0 ? " (" + rows + " записей)" : "");
    }
}

// Чтение .xlsx без Excel и сторонних программ: zip (свой разбор + DeflateStream из .NET 2.0) и XML листа потоком
public class XlsxBook : IDisposable {
    class Entry { public string Name; public int Method; public long Offset, Packed, Size; }
    readonly FileStream fs;
    readonly List<Entry> entries = new List<Entry>();
    public string[] Shared = new string[0];

    public XlsxBook(string path) {
        fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16);
        try { ReadDirectory(); }
        catch { fs.Close(); throw; }
    }

    public void Dispose() { fs.Close(); }

    static int U16(byte[] b, int p) { return b[p] | (b[p + 1] << 8); }
    static long U32(byte[] b, int p) { return (uint)(b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)); }

    void ReadDirectory() {
        long len = fs.Length;
        int tail = (int)Math.Min(len, 66000);
        byte[] t = new byte[tail];
        fs.Seek(len - tail, SeekOrigin.Begin);
        ReadFull(fs, t, tail);
        int eocd = -1;
        for (int i = tail - 22; i >= 0; i--)
            if (t[i] == 0x50 && t[i + 1] == 0x4B && t[i + 2] == 5 && t[i + 3] == 6) { eocd = i; break; }
        if (eocd < 0) throw new InvalidDataException("Файл Excel .xlsx повреждён (не zip). Сохраните выгрузку в Excel заново.");
        int count = U16(t, eocd + 10);
        long dirSize = U32(t, eocd + 12), dirOffset = U32(t, eocd + 16);
        byte[] d = new byte[dirSize];
        fs.Seek(dirOffset, SeekOrigin.Begin);
        ReadFull(fs, d, (int)dirSize);
        int p = 0;
        for (int i = 0; i < count && p + 46 <= d.Length; i++) {
            if (U32(d, p) != 0x02014b50) break;
            Entry e = new Entry();
            e.Method = U16(d, p + 10);
            e.Packed = U32(d, p + 20);
            e.Size = U32(d, p + 24);
            int nl = U16(d, p + 28), xl = U16(d, p + 30), cl = U16(d, p + 32);
            e.Offset = U32(d, p + 42);
            e.Name = Encoding.UTF8.GetString(d, p + 46, nl).Replace('\\', '/');
            entries.Add(e);
            p += 46 + nl + xl + cl;
        }
    }

    static void ReadFull(Stream s, byte[] b, int n) {
        int got = 0;
        while (got < n) {
            int r = s.Read(b, got, n - got);
            if (r <= 0) throw new EndOfStreamException();
            got += r;
        }
    }

    public bool Has(string name) { return Find(name) != null; }

    Entry Find(string name) {
        foreach (Entry e in entries) if (string.Compare(e.Name, name, StringComparison.OrdinalIgnoreCase) == 0) return e;
        return null;
    }

    // содержимое файла внутри книги - потоком (распаковка на лету)
    public Stream Open(string name) {
        Entry e = Find(name);
        if (e == null) throw new FileNotFoundException("В книге нет " + name);
        byte[] h = new byte[30];
        fs.Seek(e.Offset, SeekOrigin.Begin);
        ReadFull(fs, h, 30);
        long start = e.Offset + 30 + U16(h, 26) + U16(h, 28);
        byte[] packed = new byte[e.Packed];
        fs.Seek(start, SeekOrigin.Begin);
        ReadFull(fs, packed, (int)e.Packed);
        Stream raw = new MemoryStream(packed);
        if (e.Method == 0) return raw;
        if (e.Method == 8) return new System.IO.Compression.DeflateStream(raw, System.IO.Compression.CompressionMode.Decompress);
        throw new InvalidDataException("Неизвестное сжатие в книге (" + e.Method + ")");
    }

    static XmlReader Xml(Stream s) {
        XmlReaderSettings st = new XmlReaderSettings();
        st.IgnoreWhitespace = false;
        return XmlReader.Create(s, st);
    }

    // первый лист книги (как его открывает Excel): workbook.xml -> rels -> путь
    public string FirstSheet() {
        string rid = null;
        if (Has("xl/workbook.xml"))
            using (XmlReader r = Xml(Open("xl/workbook.xml")))
                while (r.Read())
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "sheet") {
                        for (int i = 0; i < r.AttributeCount; i++) {
                            r.MoveToAttribute(i);
                            if (r.LocalName == "id") rid = r.Value;
                        }
                        break;
                    }
        if (rid != null && Has("xl/_rels/workbook.xml.rels"))
            using (XmlReader r = Xml(Open("xl/_rels/workbook.xml.rels")))
                while (r.Read())
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship" && r.GetAttribute("Id") == rid) {
                        string target = r.GetAttribute("Target") ?? "";
                        if (target.StartsWith("/")) return target.Substring(1);
                        return "xl/" + target;
                    }
        foreach (Entry e in entries) if (e.Name.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && e.Name.EndsWith(".xml")) return e.Name;
        throw new InvalidDataException("В книге нет листов.");
    }

    public void LoadShared() {
        if (!Has("xl/sharedStrings.xml")) return;
        List<string> list = new List<string>();
        using (XmlReader r = Xml(Open("xl/sharedStrings.xml"))) {
            StringBuilder cur = null;
            int phonetic = 0;
            bool moved = false;     // ReadElementContentAsString уже передвинул на следующий узел
            while (moved || r.Read()) {
                moved = false;
                if (r.NodeType == XmlNodeType.Element) {
                    if (r.LocalName == "si") { cur = new StringBuilder(); if (r.IsEmptyElement) { list.Add(""); cur = null; } }
                    else if (r.LocalName == "rPh") { if (!r.IsEmptyElement) phonetic++; }
                    else if (r.LocalName == "t" && cur != null && phonetic == 0 && !r.IsEmptyElement) { cur.Append(r.ReadElementContentAsString()); moved = true; }
                } else if (r.NodeType == XmlNodeType.EndElement) {
                    if (r.LocalName == "si" && cur != null) { list.Add(cur.ToString()); cur = null; }
                    else if (r.LocalName == "rPh") phonetic--;
                }
            }
        }
        Shared = list.ToArray();
    }

    public delegate void Row(int number, string[] cells, bool[] numeric);

    // строки листа по порядку: ячейки - текстом как записаны (числа - как в файле, даты - числом Excel); пустые строки пропускаются
    public void ReadRows(string sheet, int width, Row onRow) {
        using (XmlReader r = Xml(Open(sheet))) {
            string[] cells = null;
            bool[] num = null;
            int rowNo = 0, col = -1;
            bool any = false;
            while (r.Read()) {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "row") {
                    string rn = r.GetAttribute("r");
                    int n;
                    rowNo = rn != null && int.TryParse(rn, out n) ? n : rowNo + 1;
                    cells = new string[width];
                    num = new bool[width];
                    col = -1;
                    any = false;
                    if (r.IsEmptyElement) cells = null;
                } else if (r.NodeType == XmlNodeType.Element && r.LocalName == "c" && cells != null) {
                    string refName = r.GetAttribute("r"), type = r.GetAttribute("t") ?? "";
                    col = refName != null ? ColumnOf(refName) : col + 1;
                    if (r.IsEmptyElement) continue;
                    string value = null;
                    int depth = r.Depth;
                    StringBuilder inline = null;
                    bool moved = false;     // ReadElementContentAsString уже передвинул на следующий узел
                    while (moved || r.Read()) {
                        moved = false;
                        if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) break;    // </c>
                        if (r.NodeType != XmlNodeType.Element || r.IsEmptyElement) continue;
                        if (r.LocalName == "v") { value = r.ReadElementContentAsString(); moved = true; }
                        else if (r.LocalName == "t" && type == "inlineStr") {
                            if (inline == null) inline = new StringBuilder();
                            inline.Append(r.ReadElementContentAsString());
                            moved = true;
                        }
                    }
                    if (inline != null) value = inline.ToString();
                    if (value == null || col < 0 || col >= width) continue;
                    bool isNum = false;
                    if (type == "s") {
                        int i;
                        value = int.TryParse(value, out i) && i >= 0 && i < Shared.Length ? Shared[i] : "";
                    } else if (type == "b") value = value == "1" ? "1" : "0";
                    else if (type == "e") value = "";
                    else if (type == "" || type == "n") isNum = true;
                    value = value.Trim();
                    if (value.Length == 0) continue;
                    cells[col] = value;
                    num[col] = isNum;
                    any = true;
                } else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "row") {
                    if (cells != null && any) onRow(rowNo, cells, num);
                    cells = null;
                }
            }
        }
    }

    // «AB12» -> 27 (с нуля)
    public static int ColumnOf(string cellRef) {
        int c = 0;
        foreach (char ch in cellRef) {
            char u = char.ToUpperInvariant(ch);
            if (u < 'A' || u > 'Z') break;
            c = c * 26 + (u - 'A' + 1);
        }
        return c - 1;
    }

    // ширина листа по строке заголовка: сколько столбцов (первая непустая строка)
    public string[] Header(string sheet, out int headerRow) {
        string[] found = null;
        int hr = 0;
        try {
            ReadRows(sheet, 400, delegate(int n, string[] cells, bool[] numeric) {
                if (found != null) throw new StopReading();
                int last = -1;
                for (int i = 0; i < cells.Length; i++) if (cells[i] != null) last = i;
                found = new string[last + 1];
                Array.Copy(cells, found, last + 1);
                hr = n;
                throw new StopReading();
            });
        } catch (StopReading) { }
        headerRow = hr;
        return found ?? new string[0];
    }

    public class StopReading : Exception { }
}

// Выгрузка быта -> своя копия: byt.dbf (поля как в таблице абонентов базы) + byt_ulicy.txt (код улицы -> н.п., улица, сельсовет)
// Книга выгрузки любого вида (смотрится по содержимому, не по расширению): таблица .dbf (25.11); .xlsx; .xls - Excel 97-2003, Excel 5.0/95 и
// Excel 2-4 (так пишет FoxPro «COPY TO ... TYPE XLS»); «xls», который на деле HTML-таблица, XML-таблица Excel 2003 или текст
// с табуляцией / точкой с запятой. Строки всех листов подряд, номер сквозной (пустые строки не считаются); первая строка -
// заголовок; на следующих листах та же строка заголовка пропускается (xls - не больше 65536 строк на листе, выгрузку делят).
public abstract class VygBook : IDisposable {
    public delegate void Row(int number, string[] cells, bool[] numeric);
    protected delegate void SheetRow(int sheet, string[] cells, bool[] numeric);
    string[] header;

    public static VygBook Open(string path) {
        byte[] h = new byte[8];
        int got;
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) got = fs.Read(h, 0, 8);
        if (got >= 2 && h[0] == 0x50 && h[1] == 0x4B) return new XlsxVygBook(path);
        if (got == 8 && h[0] == 0xD0 && h[1] == 0xCF && h[2] == 0x11 && h[3] == 0xE0 && h[4] == 0xA1 && h[5] == 0xB1 && h[6] == 0x1A && h[7] == 0xE1)
            return new XlsVygBook(path, true);
        int t = got >= 4 ? h[0] | (h[1] << 8) : 0, l = got >= 4 ? h[2] | (h[3] << 8) : 0;
        if ((t == 0x0009 || t == 0x0209 || t == 0x0409 || t == 0x0809) && l >= 4 && l <= 20) return new XlsVygBook(path, false);
        if (DbfVygBook.Is(path)) return new DbfVygBook(path);
        string kind = TextVygBook.Sniff(path);
        if (kind == "xml") return new XmlVygBook(path);
        if (kind == "html") return new HtmlVygBook(path);
        return new TextVygBook(path);
    }

    public virtual void Dispose() { }
    public virtual void Prepare() { }
    protected abstract void Sheets(int width, SheetRow onRow);

    // первая непустая строка - заголовок (номер 1)
    public string[] Header(out int headerRow) {
        string[] found = null;
        try {
            Sheets(400, delegate(int sheet, string[] cells, bool[] num) {
                int last = -1;
                for (int i = 0; i < cells.Length; i++) if (cells[i] != null) last = i;
                found = new string[last + 1];
                Array.Copy(cells, found, last + 1);
                throw new XlsxBook.StopReading();
            });
        } catch (XlsxBook.StopReading) { }
        header = found ?? new string[0];
        headerRow = 1;
        return header;
    }

    public void ReadRows(int width, Row onRow) {
        int n = 0, last = -1;
        Sheets(width, delegate(int sheet, string[] cells, bool[] num) {
            bool first = sheet != last;
            last = sheet;
            if (first && n > 0 && header != null && Same(header, cells)) return;    // заголовок следующего листа
            onRow(++n, cells, num);
        });
    }

    static bool Same(string[] h, string[] cells) {
        int any = 0;
        for (int i = 0; i < h.Length; i++) {
            if (h[i] == null) continue;
            if (i >= cells.Length || cells[i] == null || cells[i] != h[i]) return false;
            any++;
        }
        return any > 0;
    }

    // строка листа собирается из ячеек; в cells - только непустое (обрезанное)
    protected sealed class RowBuf {
        readonly int width;
        public string[] Cells;
        public bool[] Num;
        public bool Any;
        public RowBuf(int width) { this.width = width; Reset(); }
        public void Reset() { Cells = new string[width]; Num = new bool[width]; Any = false; }
        public void Set(int col, string v, bool num) {
            if (v == null || col < 0 || col >= width) return;
            v = v.Trim();
            if (v.Length == 0) return;
            Cells[col] = v;
            Num[col] = num;
            Any = true;
        }
    }

    protected static string Num(double d) { return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture); }
}

// Таблица DBF (dBase, FoxPro, Clipper) как книга выгрузки: строка 1 - имена полей (как их показывает Excel), дальше записи.
// 25.11: база - на виртуальной машине без Excel, её сценарий выгрузки (TEMP.DBF -> Excel -> .xls) не срабатывает; TEMP.DBF
// читается как есть. Кодировка - по метке таблицы (866 / 1251), без метки - по самим буквам. Удалённые записи пропускаются.
public class DbfVygBook : VygBook {
    readonly string path;
    public DbfVygBook(string path) { this.path = path; }

    // похоже ли на таблицу DBF: версия, дата правки, длины заголовка и записи, первое поле с допустимым типом
    public static bool Is(string path) {
        try {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                byte[] h = new byte[64];
                if (fs.Read(h, 0, 64) < 64) return false;
                if (Array.IndexOf(new int[] { 0x02, 0x03, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8B, 0xCB, 0xF5, 0xFB }, (int)h[0]) < 0) return false;
                if (h[2] < 1 || h[2] > 12 || h[3] < 1 || h[3] > 31) return false;
                int headerLen = h[8] | (h[9] << 8), recLen = h[10] | (h[11] << 8);
                if (headerLen < 65 || recLen < 2 || headerLen > fs.Length) return false;
                return "CNDLMFBIYTG0@+OVQW".IndexOf((char)h[32 + 11]) >= 0 && h[32] >= 0x20;
            }
        } catch { return false; }
    }

    sealed class Fld { public char Type; public int Off, Len, Dec; }

    static Encoding EncodingOf(byte ldid, byte[] sample) {
        if (ldid == 0x26 || ldid == 0x65 || ldid == 0x01 || ldid == 0x64) return Encoding.GetEncoding(866);
        if (ldid == 0x57 || ldid == 0xC9 || ldid == 0x03) return Encoding.GetEncoding(1251);
        // без метки: в 866 русские буквы - 0x80-0xAF и 0xE0-0xEF, в 1251 - 0xC0-0xFF (0xC0-0xDF в 866 - рамки, в данных их нет)
        int dos = 0, win = 0;
        foreach (byte b in sample) { if (b >= 0x80 && b <= 0xAF) dos++; else if (b >= 0xC0 && b <= 0xDF) win++; }
        return Encoding.GetEncoding(win > dos ? 1251 : 866);
    }

    protected override void Sheets(int width, SheetRow onRow) {
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16)) {
            byte[] h = new byte[32];
            fs.Read(h, 0, 32);
            int count = BitConverter.ToInt32(h, 4), headerLen = h[8] | (h[9] << 8), recLen = h[10] | (h[11] << 8);
            byte[] desc = new byte[headerLen - 32];
            fs.Read(desc, 0, desc.Length);
            List<Fld> fields = new List<Fld>();
            int off = 1;
            for (int p = 0; p + 32 <= desc.Length && desc[p] != 0x0D; p += 32) {
                Fld f = new Fld();
                f.Type = (char)desc[p + 11];
                f.Len = desc[p + 16];
                f.Dec = desc[p + 17];
                if (f.Type == 'C' && f.Dec > 0) f.Len += desc[p + 17] << 8;   // длинные символьные поля Clipper/FoxPro
                f.Off = off;
                off += f.Len;
                fields.Add(f);
            }
            // кодировка: метка (байт 29) или проба первых записей
            long avail = Math.Max(0, (fs.Length - headerLen) / Math.Max(1, recLen));
            if (count <= 0 || count > avail) count = (int)avail;
            byte[] sample = new byte[(int)Math.Min((long)recLen * Math.Min(count, 2000), 1 << 20)];
            fs.Position = headerLen;
            int sg = fs.Read(sample, 0, sample.Length);
            Encoding enc = EncodingOf(h[29], sample.Length == sg ? sample : Resize(sample, sg));
            string[] head = new string[width];
            bool[] none = new bool[width];
            for (int i = 0; i < fields.Count && i < width; i++) {
                int p = i * 32, n = 0;
                while (n < 11 && desc[p + n] != 0) n++;
                head[i] = enc.GetString(desc, p, n).Trim();
            }
            onRow(0, head, none);
            fs.Position = headerLen;
            byte[] rec = new byte[recLen];
            RowBuf row = new RowBuf(width);
            for (int r = 0; r < count; r++) {
                int got = 0;
                while (got < recLen) { int k = fs.Read(rec, got, recLen - got); if (k <= 0) break; got += k; }
                if (got < recLen || rec[0] == 0x1A) break;
                if (rec[0] == (byte)'*') continue;
                row.Reset();
                for (int i = 0; i < fields.Count && i < width; i++) {
                    Fld f = fields[i];
                    if (f.Off + f.Len > recLen) break;
                    switch (f.Type) {
                        case 'C': case 'V': row.Set(i, enc.GetString(rec, f.Off, f.Len), false); break;
                        case 'D': {
                            string d = Encoding.ASCII.GetString(rec, f.Off, f.Len).Trim();
                            if (d.Length == 8 && d != "00000000") row.Set(i, d.Substring(6, 2) + "." + d.Substring(4, 2) + "." + d.Substring(0, 4), false);
                            break;
                        }
                        case 'N': case 'F': {
                            string v = Encoding.ASCII.GetString(rec, f.Off, f.Len).Trim();
                            double x;
                            if (double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x)) row.Set(i, Num(x), true);
                            break;
                        }
                        case 'L': {
                            char c = (char)rec[f.Off];
                            if (c == 'T' || c == 't' || c == 'Y' || c == 'y') row.Set(i, "1", true);
                            else if (c == 'F' || c == 'f' || c == 'N' || c == 'n') row.Set(i, "0", true);
                            break;
                        }
                        case 'I': if (f.Len == 4) row.Set(i, BitConverter.ToInt32(rec, f.Off).ToString(), true); break;
                        case 'B': case 'O': if (f.Len == 8) row.Set(i, Num(BitConverter.ToDouble(rec, f.Off)), true); break;
                        case 'Y': if (f.Len == 8) row.Set(i, Num(BitConverter.ToInt64(rec, f.Off) / 10000.0), true); break;
                    }
                }
                if (row.Any) onRow(0, row.Cells, row.Num);
            }
        }
    }

    static byte[] Resize(byte[] a, int n) { byte[] r = new byte[Math.Max(0, n)]; Array.Copy(a, r, r.Length); return r; }
}

public class XlsxVygBook : VygBook {
    readonly XlsxBook book;
    string sheet;
    public XlsxVygBook(string path) { book = new XlsxBook(path); }
    public override void Dispose() { book.Dispose(); }
    public override void Prepare() { book.LoadShared(); sheet = book.FirstSheet(); }
    protected override void Sheets(int width, SheetRow onRow) {
        if (sheet == null) Prepare();
        book.ReadRows(sheet, width, delegate(int n, string[] cells, bool[] num) { onRow(0, cells, num); });
    }
}

// .xls: двоичный Excel (BIFF2-8). Поток книги - из составного файла OLE (Excel 5-2003) или весь файл (Excel 2-4)
public class XlsVygBook : VygBook {
    readonly byte[] wb;
    int biff = 8;
    Encoding enc = Encoding.GetEncoding(1251);
    string[] sst = new string[0];
    bool prepared;

    public XlsVygBook(string path, bool ole) {
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16))
            wb = ole ? OleStream(fs) : ReadAll(fs);
    }

    static int U16(byte[] b, int p) { return b[p] | (b[p + 1] << 8); }
    static int I32(byte[] b, int p) { return b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24); }

    static byte[] ReadAll(FileStream fs) {
        if (fs.Length > int.MaxValue) throw new InvalidDataException("Файл выгрузки слишком большой.");
        byte[] b = new byte[fs.Length];
        Read(fs, 0, b, 0, b.Length);
        return b;
    }

    static void Read(FileStream fs, long at, byte[] b, int off, int n) {
        fs.Seek(at, SeekOrigin.Begin);
        while (n > 0) {
            int r = fs.Read(b, off, n);
            if (r <= 0) throw new InvalidDataException("Файл .xls обрезан (не дочитался).");
            off += r;
            n -= r;
        }
    }

    // составной файл OLE: поток «Workbook» (Excel 97-2003) или «Book» (Excel 5.0/95)
    static byte[] OleStream(FileStream fs) {
        byte[] h = new byte[512];
        Read(fs, 0, h, 0, 512);
        int ss = 1 << U16(h, 0x1E), mss = 1 << U16(h, 0x20);
        int nFat = I32(h, 0x2C), dirStart = I32(h, 0x30), cutoff = I32(h, 0x38), miniStart = I32(h, 0x3C), difat = I32(h, 0x44), nDifat = I32(h, 0x48);
        List<int> fatSectors = new List<int>();
        for (int i = 0; i < 109 && fatSectors.Count < nFat; i++) fatSectors.Add(I32(h, 0x4C + i * 4));
        byte[] sec = new byte[ss];
        for (int k = 0; k < nDifat && difat >= 0 && fatSectors.Count < nFat; k++) {
            Read(fs, (long)(difat + 1) * ss, sec, 0, ss);
            for (int i = 0; i < ss / 4 - 1 && fatSectors.Count < nFat; i++) fatSectors.Add(I32(sec, i * 4));
            difat = I32(sec, ss - 4);
        }
        int[] fat = new int[fatSectors.Count * (ss / 4)];
        for (int k = 0; k < fatSectors.Count; k++) {
            Read(fs, (long)(fatSectors[k] + 1) * ss, sec, 0, ss);
            for (int i = 0; i < ss / 4; i++) fat[k * (ss / 4) + i] = I32(sec, i * 4);
        }
        byte[] dir = Chain(fs, fat, dirStart, ss, -1);
        int found = -1;
        for (int p = 0; p + 128 <= dir.Length; p += 128) {
            int nl = U16(dir, p + 64);
            string name = nl >= 2 ? Encoding.Unicode.GetString(dir, p, Math.Min(64, nl) - 2) : "";
            if (dir[p + 66] == 2 && (string.Compare(name, "Workbook", true) == 0 || string.Compare(name, "Book", true) == 0)) { found = p; if (string.Compare(name, "Workbook", true) == 0) break; }
        }
        if (found < 0) throw new InvalidDataException("Это не книга Excel: в файле .xls нет листов (поток Workbook).");
        int start = I32(dir, found + 116), size = I32(dir, found + 120);
        if (size >= cutoff) return Chain(fs, fat, start, ss, size);
        // маленький поток - в «мини-потоке» корневого элемента
        byte[] mini = Chain(fs, fat, I32(dir, 116), ss, I32(dir, 120));
        byte[] mfRaw = Chain(fs, fat, miniStart, ss, -1);
        byte[] r = new byte[size];
        int s = start, got = 0, guard = 0;
        while (got < size && s >= 0 && guard++ < 10000000) {
            int n = Math.Min(mss, size - got);
            Buffer.BlockCopy(mini, s * mss, r, got, n);
            got += n;
            s = s * 4 + 4 <= mfRaw.Length ? I32(mfRaw, s * 4) : -2;
        }
        return r;
    }

    static byte[] Chain(FileStream fs, int[] fat, int start, int ss, int size) {
        List<int> chain = new List<int>();
        for (int s = start; s >= 0 && s < fat.Length && chain.Count <= fat.Length; s = fat[s]) chain.Add(s);
        long total = (long)chain.Count * ss;
        int n = size >= 0 ? (int)Math.Min(size, total) : (int)total;
        byte[] r = new byte[n];
        for (int i = 0, got = 0; i < chain.Count && got < n; i++) {
            int k = Math.Min(ss, n - got);
            Read(fs, (long)(chain[i] + 1) * ss, r, got, k);
            got += k;
        }
        return r;
    }

    // кодировка строк Excel 2-5 (Excel 97-2003 - Юникод): по записи CODEPAGE; нет её или «западная» - по буквам (866 или 1251)
    public override void Prepare() {
        if (prepared) return;
        prepared = true;
        int cp = -1;
        long c866 = 0, c1251 = 0;
        for (int p = 0; p + 4 <= wb.Length; ) {
            int t = U16(wb, p), l = U16(wb, p + 2), d = p + 4;
            if (d + l > wb.Length) break;
            if (t == 0x0809 || t == 0x0409 || t == 0x0209 || t == 0x0009) {
                if (t == 0x0809) biff = l >= 2 && U16(wb, d) == 0x0600 ? 8 : 5;
                else biff = t == 0x0009 ? 2 : t == 0x0209 ? 3 : 4;
            } else if (t == 0x0042 && l >= 2) cp = U16(wb, d);
            else if (t == 0x00FC && biff == 8) sst = ReadSst(p);
            else if (biff < 8) {
                int s = -1, n = 0;
                if (t == 0x0004 && l >= 8) { s = d + 8; n = wb[d + 7]; }
                else if ((t == 0x0204 || t == 0x00D6) && l >= 8) { s = d + 8; n = U16(wb, d + 6); }
                for (int i = 0; s >= 0 && i < n && s + i < d + l; i++) {
                    byte b = wb[s + i];
                    if (b >= 0x80 && b <= 0xAF) c866++;
                    else if ((b >= 0xC0 && b <= 0xDF) || b >= 0xF0) c1251++;
                }
            }
            p = d + l;
        }
        if (cp == 866 || cp == 1251) enc = Encoding.GetEncoding(cp);
        else if (cp > 0 && cp != 367 && cp != 1252 && cp != 32768 && cp != 32769 && cp != 1200) { try { enc = Encoding.GetEncoding(cp); } catch { } }
        else if (c866 > c1251) enc = Encoding.GetEncoding(866);
    }

    // таблица строк (SST) с продолжениями (CONTINUE): строка может переходить в следующую запись
    string[] ReadSst(int p0) {
        List<int> st = new List<int>(), en = new List<int>();
        int l0 = U16(wb, p0 + 2);
        st.Add(p0 + 4 + 8);
        en.Add(p0 + 4 + l0);
        int unique = I32(wb, p0 + 8);
        for (int p = p0 + 4 + l0; p + 4 <= wb.Length && U16(wb, p) == 0x003C; ) {
            int l = U16(wb, p + 2);
            st.Add(p + 4);
            en.Add(Math.Min(wb.Length, p + 4 + l));
            p += 4 + l;
        }
        List<string> r = new List<string>(Math.Max(0, Math.Min(unique, 1 << 20)));
        int seg = 0, pos = st[0];
        StringBuilder sb = new StringBuilder();
        try {
            while (r.Count < unique) {
                if (pos >= en[seg]) { if (++seg >= st.Count) break; pos = st[seg]; }
                int cch = Byte2(ref seg, ref pos, st, en);
                int flags = Byte1(ref seg, ref pos, st, en);
                int runs = (flags & 8) != 0 ? Byte2(ref seg, ref pos, st, en) : 0;
                int ext = 0;
                if ((flags & 4) != 0) { ext = Byte2(ref seg, ref pos, st, en); ext |= Byte2(ref seg, ref pos, st, en) << 16; }
                bool high = (flags & 1) != 0;
                sb.Length = 0;
                for (int i = 0; i < cch; i++) {
                    if (pos >= en[seg]) {
                        if (++seg >= st.Count) break;
                        pos = st[seg];
                        high = (wb[pos++] & 1) != 0;      // в продолжении - заново признак «два байта на букву»
                    }
                    if (high) { sb.Append((char)U16(wb, pos)); pos += 2; } else sb.Append((char)wb[pos++]);
                }
                r.Add(sb.ToString());
                Skip(runs * 4 + ext, ref seg, ref pos, st, en);
            }
        } catch (IndexOutOfRangeException) { } catch (ArgumentOutOfRangeException) { }     // обрезанная таблица строк - что прочиталось
        return r.ToArray();
    }

    int Byte1(ref int seg, ref int pos, List<int> st, List<int> en) {
        if (pos >= en[seg]) { seg++; pos = st[seg]; }
        return wb[pos++];
    }

    int Byte2(ref int seg, ref int pos, List<int> st, List<int> en) {
        int a = Byte1(ref seg, ref pos, st, en);
        return a | (Byte1(ref seg, ref pos, st, en) << 8);
    }

    void Skip(int n, ref int seg, ref int pos, List<int> st, List<int> en) {
        while (n > 0 && seg < st.Count) {
            int k = Math.Min(n, en[seg] - pos);
            pos += k;
            n -= k;
            if (n > 0) { if (seg + 1 >= st.Count) { seg++; break; } seg++; pos = st[seg]; }
        }
    }

    static double Rk(int rk) {
        double d = (rk & 2) != 0 ? (double)(rk >> 2) : BitConverter.Int64BitsToDouble(((long)(rk & unchecked((int)0xFFFFFFFC))) << 32);
        return (rk & 1) != 0 ? d / 100 : d;
    }

    // строка ячейки: Excel 97-2003 - длина 2 байта, признак, буквы (1 или 2 байта); Excel 2-5 - длина и байты в кодировке книги
    string Text(int p, int end, bool shortLen) {
        if (biff == 8) {
            int cch = shortLen ? wb[p] : U16(wb, p);
            p += shortLen ? 1 : 2;
            bool high = (wb[p++] & 1) != 0;
            int n = Math.Min(cch, (end - p) / (high ? 2 : 1));
            return high ? Encoding.Unicode.GetString(wb, p, n * 2) : Encoding.GetEncoding(28591).GetString(wb, p, n);
        }
        int len = shortLen ? wb[p] : U16(wb, p);
        p += shortLen ? 1 : 2;
        return enc.GetString(wb, p, Math.Max(0, Math.Min(len, end - p)));
    }

    protected override void Sheets(int width, SheetRow onRow) {
        Prepare();
        RowBuf row = new RowBuf(width);
        int sheet = -1, depth = 0, curRow = -1, fRow = -1, fCol = -1;
        bool inSheet = false;
        for (int p = 0; p + 4 <= wb.Length; ) {
            int t = U16(wb, p), l = U16(wb, p + 2), d = p + 4, e = d + l;
            if (e > wb.Length) break;
            p = e;
            if (t == 0x0809 || t == 0x0409 || t == 0x0209 || t == 0x0009) {
                depth++;
                int kind = l >= 4 ? U16(wb, d + 2) : 0x10;
                if (depth == 1) { inSheet = kind == 0x0010 || (t != 0x0809 && kind != 0x0020 && kind != 0x0040); if (inSheet) sheet++; curRow = -1; }
                continue;
            }
            if (t == 0x000A) {
                if (depth == 1 && inSheet && row.Any) { onRow(sheet, row.Cells, row.Num); row.Reset(); }
                if (depth > 0) depth--;
                if (depth == 0) inSheet = false;
                continue;
            }
            if (!inSheet || depth != 1 || l < 4) continue;
            int r = U16(wb, d), c = U16(wb, d + 2);
            string v = null;
            bool num = false;
            switch (t) {
                case 0x00FD: if (l >= 10) { int i = I32(wb, d + 6); v = i >= 0 && i < sst.Length ? sst[i] : null; } break;   // LABELSST
                case 0x0204: case 0x00D6: if (l >= 8) v = Text(d + 6, e, false); break;                                    // LABEL, RSTRING
                case 0x0004: if (l >= 8) v = enc.GetString(wb, d + 8, Math.Min(wb[d + 7], e - d - 8)); break;                // LABEL (Excel 2)
                case 0x0203: if (l >= 14) { v = Num(BitConverter.ToDouble(wb, d + 6)); num = true; } break;                 // NUMBER
                case 0x0003: if (l >= 15) { v = Num(BitConverter.ToDouble(wb, d + 7)); num = true; } break;                 // NUMBER (Excel 2)
                case 0x0002: if (l >= 9) { v = U16(wb, d + 7).ToString(); num = true; } break;                               // INTEGER (Excel 2)
                case 0x027E: if (l >= 10) { v = Num(Rk(I32(wb, d + 6))); num = true; } break;                                // RK
                case 0x0205: if (l >= 8 && wb[d + 7] == 0) v = wb[d + 6] != 0 ? "1" : "0"; break;                           // BOOLERR
                case 0x00BD:                                                                                                    // MULRK
                    for (int q = d + 4, k = c; q + 6 <= e - 2; q += 6, k++) {
                        if (r != curRow) { if (row.Any) onRow(sheet, row.Cells, row.Num); row.Reset(); curRow = r; }
                        row.Set(k, Num(Rk(I32(wb, q + 2))), true);
                    }
                    continue;
                case 0x0006: case 0x0206: case 0x0406: {                                                                       // FORMULA - готовый результат
                    int o = biff == 2 ? d + 7 : d + 6;
                    if (o + 8 > e) break;
                    if (wb[o + 6] == 0xFF && wb[o + 7] == 0xFF) {
                        if (wb[o] == 0) { fRow = r; fCol = c; }        // текст - в следующей записи STRING
                        else if (wb[o] == 1) v = wb[o + 2] != 0 ? "1" : "0";
                    } else { v = Num(BitConverter.ToDouble(wb, o)); num = true; }
                    break;
                }
                case 0x0207: case 0x0007:                                                                                       // STRING (текст формулы)
                    if (fRow >= 0) { r = fRow; c = fCol; v = t == 0x0007 ? enc.GetString(wb, d + 1, Math.Min(wb[d], e - d - 1)) : Text(d, e, false); fRow = -1; }
                    break;
                default: continue;
            }
            if (v == null) continue;
            if (r != curRow) { if (row.Any) onRow(sheet, row.Cells, row.Num); row.Reset(); curRow = r; }
            row.Set(c, v, num);
        }
        if (row.Any) onRow(sheet, row.Cells, row.Num);
    }
}

// текстовые «xls»: определение вида и кодировки
public class TextVygBook : VygBook {
    protected readonly string path;
    public TextVygBook(string path) { this.path = path; }

    static byte[] Head(string path, int n) {
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
            byte[] b = new byte[Math.Min(n, fs.Length)];
            int got = 0;
            while (got < b.Length) { int r = fs.Read(b, got, b.Length - got); if (r <= 0) break; got += r; }
            return b;
        }
    }

    // «xml» (таблица Excel 2003), «html» или «text»
    public static string Sniff(string path) {
        byte[] b = Head(path, 65536);
        string s = (b.Length > 1 && b[1] == 0 ? Encoding.Unicode.GetString(b) : Encoding.GetEncoding(28591).GetString(b)).ToLowerInvariant();
        if (s.Contains("urn:schemas-microsoft-com:office:spreadsheet") && s.Contains("<workbook")) return "xml";
        if (s.Contains("<table") || s.Contains("<html") || s.Contains("<!doctype")) return "html";
        return "text";
    }

    // кодировка: метка BOM, charset=..., иначе UTF-8, если весь начальный кусок - верный UTF-8, иначе 1251
    protected Encoding Detect() {
        byte[] b = Head(path, 1 << 20);
        if (b.Length >= 2 && ((b[0] == 0xFF && b[1] == 0xFE) || (b[0] == 0xFE && b[1] == 0xFF))) return b[0] == 0xFF ? Encoding.Unicode : Encoding.BigEndianUnicode;
        if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8;
        string a = Encoding.GetEncoding(28591).GetString(b, 0, Math.Min(b.Length, 8192)).ToLowerInvariant();
        int cs = a.IndexOf("charset=");
        if (cs < 0) cs = a.IndexOf("encoding=");
        if (cs >= 0) {
            string rest = a.Substring(a.IndexOf('=', cs) + 1).TrimStart('"', '\'', ' ');
            StringBuilder name = new StringBuilder();
            foreach (char ch in rest) { if (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_') name.Append(ch); else break; }
            try { if (name.Length > 0) return Encoding.GetEncoding(name.ToString()); } catch { }
        }
        int cut = b.Length;
        if (cut == 1 << 20) {               // кусок файла: последнюю (может быть, разрезанную) букву UTF-8 не проверять
            while (cut > 0 && b.Length - cut < 3 && (b[cut - 1] & 0xC0) == 0x80) cut--;
            if (cut > 0 && b[cut - 1] >= 0xC0) cut--;
        }
        try { new UTF8Encoding(false, true).GetString(b, 0, cut); return Encoding.UTF8; } catch { return Encoding.GetEncoding(1251); }
    }

    protected override void Sheets(int width, SheetRow onRow) {
        using (StreamReader sr = new StreamReader(path, Detect(), true)) {
            char sep = '\0';
            string line;
            while ((line = sr.ReadLine()) != null) {
                if (sep == '\0') { if (line.Trim().Length == 0) continue; sep = line.IndexOf('\t') >= 0 ? '\t' : line.IndexOf(';') >= 0 ? ';' : ','; }
                RowBuf row = new RowBuf(width);
                List<string> cells = Split(line, sep);
                for (int i = 0; i < cells.Count; i++) row.Set(i, cells[i], false);
                if (row.Any) onRow(0, row.Cells, row.Num);
            }
        }
    }

    static List<string> Split(string line, char sep) {
        List<string> r = new List<string>();
        StringBuilder sb = new StringBuilder();
        bool q = false;
        for (int i = 0; i < line.Length; i++) {
            char ch = line[i];
            if (q) {
                if (ch == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; } else q = false; }
                else sb.Append(ch);
            } else if (ch == '"' && sb.Length == 0) q = true;
            else if (ch == sep) { r.Add(sb.ToString()); sb.Length = 0; }
            else sb.Append(ch);
        }
        r.Add(sb.ToString());
        return r;
    }
}

// HTML-таблица (часто сохраняют под именем .xls): <tr> - строка, <td>/<th> - ячейка (colspan - пустые следом), каждая <table> - лист
public class HtmlVygBook : TextVygBook {
    public HtmlVygBook(string path) : base(path) { }

    protected override void Sheets(int width, SheetRow onRow) {
        using (StreamReader sr = new StreamReader(path, Detect(), true)) {
            RowBuf row = null;
            StringBuilder cell = null, tag = new StringBuilder();
            int col = 0, span = 1, sheet = -1;
            bool skip = false;
            int ch;
            while ((ch = sr.Read()) >= 0) {
                if (ch != '<') { if (cell != null && !skip) cell.Append((char)ch); continue; }
                tag.Length = 0;
                char quote = '\0';
                while ((ch = sr.Read()) >= 0) {
                    if (quote != '\0') { if (ch == quote) quote = '\0'; }
                    else if (ch == '"' || ch == '\'') quote = (char)ch;
                    else if (ch == '>') break;
                    tag.Append((char)ch);
                    if (tag.Length == 3 && tag.ToString() == "!--") {          // комментарий - до -->
                        int a = 0, b = 0;
                        while ((ch = sr.Read()) >= 0) { if (ch == '>' && a == '-' && b == '-') break; a = b; b = ch; }
                        tag.Length = 0;
                        break;
                    }
                }
                if (tag.Length == 0) continue;
                string t = tag.ToString();
                bool end = t.StartsWith("/");
                string name = t.TrimStart('/');
                int sp = 0;
                while (sp < name.Length && (char.IsLetterOrDigit(name[sp]))) sp++;
                string low = name.Substring(0, sp).ToLowerInvariant();
                if (low == "style" || low == "script") { skip = !end; continue; }
                if (low == "table" && !end) { sheet++; continue; }
                if (low == "br" && cell != null) { cell.Append(' '); continue; }
                if (low == "tr") {
                    if (cell != null) { Put(row, ref col, cell, span); cell = null; }
                    if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num);
                    row = end ? null : new RowBuf(width);
                    col = 0;
                    continue;
                }
                if (low == "td" || low == "th") {
                    if (cell != null) { Put(row, ref col, cell, span); cell = null; }
                    if (!end) {
                        if (row == null) { row = new RowBuf(width); col = 0; }
                        cell = new StringBuilder();
                        span = 1;
                        string lt = t.ToLowerInvariant();
                        int cs = lt.IndexOf("colspan");
                        if (cs >= 0) {
                            string digits = "";
                            for (int i = cs + 7; i < lt.Length && digits.Length < 3; i++) { if (char.IsDigit(lt[i])) digits += lt[i]; else if (digits.Length > 0) break; }
                            if (digits.Length > 0) span = Math.Max(1, Math.Min(200, int.Parse(digits)));
                        }
                    }
                    continue;
                }
                if (low == "table" && end) {
                    if (cell != null) { Put(row, ref col, cell, span); cell = null; }
                    if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num);
                    row = null;
                }
            }
            if (cell != null) Put(row, ref col, cell, span);
            if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num);
        }
    }

    static void Put(RowBuf row, ref int col, StringBuilder cell, int span) {
        if (row != null) row.Set(col, Entities(cell.ToString()), false);
        col += span;
    }

    // &amp; &lt; &gt; &quot; &nbsp; &#1234; &#x4D2; ; пробелы и переводы строк - один пробел
    static string Entities(string s) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '&') {
                int semi = s.IndexOf(';', i);
                if (semi > i && semi - i <= 10) {
                    string e = s.Substring(i + 1, semi - i - 1).ToLowerInvariant();
                    string r = null;
                    if (e == "amp") r = "&"; else if (e == "lt") r = "<"; else if (e == "gt") r = ">"; else if (e == "quot") r = "\"";
                    else if (e == "nbsp") r = " "; else if (e == "apos") r = "'";
                    else if (e.StartsWith("#")) {
                        int code;
                        bool ok = e.StartsWith("#x") ? int.TryParse(e.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out code) : int.TryParse(e.Substring(1), out code);
                        if (ok && code > 0 && code < 0xFFFF) r = ((char)code).ToString();
                    }
                    if (r != null) { sb.Append(r); i = semi; continue; }
                }
            }
            sb.Append(char.IsWhiteSpace(c) ? ' ' : c);
        }
        string t = sb.ToString();
        while (t.Contains("  ")) t = t.Replace("  ", " ");
        return t;
    }
}

// XML-таблица Excel 2003 (Workbook/Worksheet/Table/Row/Cell/Data; ss:Index - номер столбца или строки с 1)
public class XmlVygBook : TextVygBook {
    public XmlVygBook(string path) : base(path) { }

    protected override void Sheets(int width, SheetRow onRow) {
        XmlReaderSettings st = new XmlReaderSettings();
        st.IgnoreWhitespace = true;
        using (XmlReader r = XmlReader.Create(path, st)) {
            RowBuf row = null;
            int sheet = -1, col = 0;
            string type = "";
            bool moved = false;
            while (moved || r.Read()) {
                moved = false;
                if (r.NodeType == XmlNodeType.Element) {
                    string n = r.LocalName;
                    if (n == "Worksheet") sheet++;
                    else if (n == "Row") {
                        if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num);
                        row = new RowBuf(width);
                        col = 0;
                        if (r.IsEmptyElement) row = null;
                    } else if (n == "Cell" && row != null) {
                        string idx = r.GetAttribute("Index", "urn:schemas-microsoft-com:office:spreadsheet");
                        int k;
                        if (idx != null && int.TryParse(idx, out k) && k > 0) col = k - 1;
                        string across = r.GetAttribute("MergeAcross", "urn:schemas-microsoft-com:office:spreadsheet");
                        int m = 0;
                        if (across != null) int.TryParse(across, out m);
                        if (r.IsEmptyElement) col += 1 + m;
                        else cellSpan = 1 + m;
                    } else if (n == "Data" && row != null && !r.IsEmptyElement) {
                        type = r.GetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet") ?? "";
                        string v = r.ReadElementContentAsString();
                        moved = true;
                        if (type == "DateTime") { int dot = v.IndexOf('.'); if (dot > 0) v = v.Substring(0, dot); }
                        row.Set(col, v, type == "Number");
                    }
                } else if (r.NodeType == XmlNodeType.EndElement) {
                    if (r.LocalName == "Cell") { col += cellSpan; cellSpan = 1; }
                    else if (r.LocalName == "Row") { if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num); row = null; }
                }
            }
            if (row != null && row.Any) onRow(Math.Max(0, sheet), row.Cells, row.Num);
        }
    }

    int cellSpan = 1;
}

public static class VygruzkaByt {
    public const string Version = "1";

    // столбец выгрузки -> поле копии (как в ABONENT.DBF, чтобы «Поиск» и «Договор» понимали их как раньше); D - дата
    static readonly string[] Known = {
        "НОМЕР", "NAB", "ФИО", "FIO", "NNP", "NP", "EXP_4", "EXP_4", "УЛИЦА", "ULICA", "ДОМ", "DOM", "КОР", "KOR", "КВ", "KV",
        "ТЕЛ", "TEL", "КОД_ЛГ", "KOD_LG", "GIL", "GIL", "LGI", "LGI", "СК", "SKID", "ТАРИФ", "KTAR", "КНИГА", "KNIGA",
        "ТИП_СЧ", "TIPCH", "РАЗР", "RAS", "N_СЧЕТЧИКА", "NCH", "KVP", "KVP", "ГОД", "GODPR", "НАЧ_ПОК", "NPOK", "ДАТА_ПР", "DPR:D",
        "ПОК", "POK", "ОПЛ_ПОК", "OPOK", "ДАТА_ЗАМ", "DZA:D", "ПОК_СН_СЧ", "PZS", "N_ЗАМ_СЧ", "NSC", "ТИП_ЗАМ_СЧ", "TSC",
        "IDSS", "IDSS", "N_F", "N_F", "ПР_ДИФ", "PR_DIF", "ПР_ВОДА_ГА", "PR_VODA", "КОМ_КВ", "KOM_KV", "СОБСТВ", "SOBSTV",
        "АСКУЭ", "ASKUE", "МЕСТО_ПУ", "MESTO_PU", "ГОД_ЗАМ_ПУ", "GOD_ZPU", "КВАР_ЗАМ_П", "KV_ZPU", "М_ТЕЛ", "M_TEL",
        "СОГЛ_СМС", "SOGL_SMS", "СЕЛЬСОВЕТ", "SELSOVET", "U", "U", "ЭЛ_ПУ", "EL_PU", "МП_ИНТ", "MP_INT", "ДАТА_НЗ", "DATA_NZ:D",
        "РАБОТА", "RABOTA", "ДАТА_ИЗВ", "DATA_IZV:D", "ВИД_ИЗВ", "VID_IZV", "ДАТА_ДОГ", "DATA_DOG:D", "Д_ПЕРЕОФ", "D_PEREOF:D",
        "КАТ_ГР", "KAT_GR", "ПР_ПЛ", "PR_PL",
        // наряды
        "NAB", "NAB", "Н_З", "NZ", "ДАТА_ОФ", "DOF:D", "ФИО_ПР", "FIO_PR", "НОМЕР_СЧ", "NCH", "РАЗ", "RAS", "ГОД_П", "GODPR", "ПОК_ПР", "POK_PR",
        "ВИД_АСКУЭ", "VID_ASKUE", "ДАТА_ВЫП", "DVIP:D", "ФИО_ВЫП", "FIO_VIP", "ПРИМЕЧАНИЕ", "PRIM" };

    public class Column {
        public int Source;           // номер столбца в выгрузке
        public string Header;        // как подписан в выгрузке
        public string Field;         // поле копии
        public bool Date;
        public bool Pad = true;      // NAB - лицевой: дополнить нулями до 9 (у юр базы NAB - код абонента в договоре, не трогать)
        public int Length = 1;
    }

    static string Norm(string h) { return (h ?? "").Trim().ToUpperInvariant().Replace('Ё', 'Е').Replace(' ', '_'); }

    static string Translit(string h) {
        string ru = "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
        string[] la = { "A", "B", "V", "G", "D", "E", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "H", "C", "CH", "SH", "SC", "", "Y", "", "E", "JU", "JA" };
        StringBuilder sb = new StringBuilder();
        foreach (char c in Norm(h)) {
            int i = ru.IndexOf(c);
            if (i >= 0) sb.Append(la[i]);
            else if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '_') sb.Append(c);
        }
        string s = sb.ToString();
        if (s.Length == 0 || char.IsDigit(s[0])) s = "F" + s;
        return s.Length > 10 ? s.Substring(0, 10) : s;
    }

    public static List<Column> MapColumns(string[] header) {
        List<Column> cols = new List<Column>();
        Dictionary<string, bool> used = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        // служебные поля копии - чтобы незнакомый столбец их не занял
        used["UL"] = true;
        for (int i = 0; i < header.Length; i++) {
            string h = Norm(header[i]);
            if (h.Length == 0) continue;
            string field = null;
            for (int k = 0; k + 1 < Known.Length; k += 2) if (Known[k] == h) { field = Known[k + 1]; break; }
            bool date = false;
            if (field != null && field.EndsWith(":D")) { date = true; field = field.Substring(0, field.Length - 2); }
            if (field == null) {
                field = Translit(h);
                date = h.StartsWith("ДАТА") || h.StartsWith("Д_");
            }
            string f = field;
            for (int n = 2; used.ContainsKey(f); n++) f = (field.Length > 8 ? field.Substring(0, 8) : field) + n;
            used[f] = true;
            Column c = new Column();
            c.Source = i;
            c.Header = (header[i] ?? "").Trim();
            c.Field = f;
            c.Date = date;
            cols.Add(c);
        }
        return cols;
    }

    // число из файла: «45678123» / «4.5E-2» -> текст без лишнего (целые - без «.0»)
    public static string Number(string v) {
        double d;
        if (!double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d)) return v;
        if (Math.Abs(d) < 1e15 && Math.Abs(d - Math.Round(d)) < 1e-9) return ((long)Math.Round(d)).ToString();
        return d.ToString("0.##########", System.Globalization.CultureInfo.InvariantCulture);
    }

    // дата Excel (число дней от 30.12.1899) или текст «14.01.2026» -> «20260114» (как дата в DBF); не дата - пусто
    public static string Date(string v, bool numeric) {
        double d;
        if (numeric && double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d) && d > 1 && d < 2958466)
            return new DateTime(1899, 12, 30).AddDays(Math.Floor(d)).ToString("yyyyMMdd");
        DateTime dt;
        string[] formats = { "dd.MM.yyyy", "d.M.yyyy", "yyyy-MM-dd", "dd.MM.yyyy H:mm:ss", "yyyy-MM-ddTHH:mm:ss", "dd.MM.yy" };
        if (DateTime.TryParseExact(v.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt))
            return dt.ToString("yyyyMMdd");
        return "";
    }

    // значение ячейки для копии
    public static string Value(Column c, string v, bool numeric) {
        if (v == null) return "";
        if (c.Date) return Date(v, numeric);
        if (numeric) v = Number(v);
        if (c.Field == "NAB" && c.Pad) {
            string digits = "";
            foreach (char ch in v) if (ch >= '0' && ch <= '9') digits += ch;
            if (digits.Length > 0 && digits.Length < 9) return digits.PadLeft(9, '0');
            return digits.Length > 0 ? digits : v;
        }
        return v.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
    }

    static string Key(string ss, string np, string ul) {
        return (ss + "|" + np + "|" + ul).ToLowerInvariant().Replace('ё', 'е');
    }

    public static string Build(string xlsx, VygProgress progress) { return Build(VygKind.Byt, xlsx, progress); }

    // юр база в .dbf (читается как есть, своя копия не нужна): прочитать заголовок - сколько строк, есть ли нужные поля
    public static string CheckDbf(string path, out bool ok) {
        ok = false;
        try {
            byte[] h = new byte[32];
            List<string> fields = new List<string>();
            long count;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                if (fs.Read(h, 0, 32) < 32) return "Юр база " + Path.GetFileName(path) + " - не таблица DBF (файл слишком короткий).";
                count = (uint)(h[4] | (h[5] << 8) | (h[6] << 16) | (h[7] << 24));
                int headerLen = h[8] | (h[9] << 8);
                byte[] f = new byte[32];
                for (int p = 32; p + 32 <= headerLen && fs.Read(f, 0, 32) == 32 && f[0] != 0x0D; p += 32)
                    fields.Add(Encoding.ASCII.GetString(f, 0, 11).Split('\0')[0].Trim().ToUpperInvariant());
            }
            if (fields.Count == 0) return "Юр база " + Path.GetFileName(path) + " - не таблица DBF (нет полей).";
            if (!fields.Contains("KP") || !fields.Contains("NAB"))
                return "В юр базе " + Path.GetFileName(path) + " нет полей KP и NAB - нужна таблица точек учёта FTU (KP, NP, NAB, NAB_NAIM, NAIM, NST, TIPSCH...).";
            ok = true;
            return "Юр база (DBF, читается как есть): " + count.ToString("#,0") + " строк, от " + File.GetLastWriteTime(path).ToString("dd.MM.yyyy HH:mm") + ".";
        } catch (Exception ex) { return "Юр база " + Path.GetFileName(path) + " не читается: " + ex.Message; }
    }

    // Пересоздать копию выгрузки (быт, наряды, юр лица). Пишется во временные файлы, готовое заменяет старое (прервалось - старое
    // цело). Строки без ключевого столбца (у быта - НОМЕР, у нарядов - Н_З, у юр - KP) пропускаются. Возвращает итог для окна.
    public static string Build(VygKind kind, string xlsx, VygProgress progress) {
        DateTime t0 = DateTime.Now;
        string folder = Istochnik.CacheFolder;
        Directory.CreateDirectory(folder);
        FileInfo src = new FileInfo(xlsx);
        Encoding enc = Encoding.GetEncoding(1251);
        List<Column> cols;
        int headerRow, rows = 0, skipped = 0;
        int ss = -1, np = -1, ul = -1;
        Dictionary<string, int> streetCodes = new Dictionary<string, int>();
        List<string[]> streets = new List<string[]>();
        using (VygBook book = VygBook.Open(xlsx)) {
            if (progress != null) progress("Читаю выгрузку " + kind.Name + ": тексты...");
            book.Prepare();
            string[] header = book.Header(out headerRow);
            cols = MapColumns(header);
            if (kind.Code == "jur") foreach (Column c in cols) c.Pad = false;
            Column nab = null;
            foreach (Column c in cols) {
                if (nab == null && Array.IndexOf(kind.Keys, Norm(c.Header)) >= 0) nab = c;
                if (c.Field == "SELSOVET") ss = c.Source;
                if (c.Field == "NP") np = c.Source;
                if (c.Field == "ULICA") ul = c.Source;
            }
            if (nab == null)
                throw new InvalidDataException("В первой строке выгрузки " + kind.Name + " нет столбца «" + kind.Keys[0] + "». Заголовки: " + string.Join(", ", header));
            if (kind.Streets == null) { ss = np = ul = -1; }
            int width = header.Length;
            // проход 1: длины полей, коды улиц
            int seen = 0;
            book.ReadRows(width, delegate(int n, string[] cells, bool[] num) {
                if (n <= headerRow) return;
                if (cells[nab.Source] == null) { skipped++; return; }
                foreach (Column c in cols) {
                    string v = Value(c, cells[c.Source], num[c.Source]);
                    int len = enc.GetByteCount(v);
                    if (len > c.Length) c.Length = Math.Min(254, len);
                }
                string k = Key(ss >= 0 ? cells[ss] ?? "" : "", np >= 0 ? cells[np] ?? "" : "", ul >= 0 ? cells[ul] ?? "" : "");
                if (k != "||" && !streetCodes.ContainsKey(k)) {
                    streetCodes[k] = streets.Count + 1;
                    streets.Add(new string[] { (streets.Count + 1).ToString(), np >= 0 ? cells[np] ?? "" : "", ul >= 0 ? cells[ul] ?? "" : "", ss >= 0 ? cells[ss] ?? "" : "" });
                }
                if (++seen % 20000 == 0 && progress != null) progress("Читаю выгрузку " + kind.Name + ": " + seen.ToString("#,0") + " записей...");
            });
            foreach (Column c in cols) if (c.Date) c.Length = 8;
            // проход 2: запись копии
            string tmpTable = kind.Table + ".tmp";
            int ulLen = Math.Max(1, streets.Count.ToString().Length);
            using (FileStream fs = new FileStream(tmpTable, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16)) {
                List<string> names = new List<string>();
                List<char> types = new List<char>();
                List<int> lens = new List<int>();
                names.Add("UL"); types.Add('C'); lens.Add(ulLen);
                foreach (Column c in cols) { names.Add(c.Field); types.Add(c.Date ? 'D' : 'C'); lens.Add(c.Length); }
                int recLen = 1;
                foreach (int l in lens) recLen += l;
                WriteHeader(fs, names, types, lens, 0, recLen);
                byte[] rec = new byte[recLen];
                int written = 0;
                book.ReadRows(width, delegate(int n, string[] cells, bool[] num) {
                    if (n <= headerRow || cells[nab.Source] == null) return;
                    for (int i = 0; i < rec.Length; i++) rec[i] = 0x20;
                    int p = 1;
                    string k = Key(ss >= 0 ? cells[ss] ?? "" : "", np >= 0 ? cells[np] ?? "" : "", ul >= 0 ? cells[ul] ?? "" : "");
                    int code;
                    if (streetCodes.TryGetValue(k, out code)) Put(rec, p, ulLen, enc, code.ToString().PadLeft(ulLen));
                    p += ulLen;
                    foreach (Column c in cols) {
                        Put(rec, p, c.Length, enc, Value(c, cells[c.Source], num[c.Source]));
                        p += c.Length;
                    }
                    fs.Write(rec, 0, rec.Length);
                    written++;
                    if (written % 20000 == 0 && progress != null) progress("Создаю индексы выгрузки " + kind.Name + ": " + written.ToString("#,0") + " из " + seen.ToString("#,0") + "...");
                });
                fs.WriteByte(0x1A);
                rows = written;
                fs.Seek(0, SeekOrigin.Begin);
                WriteHeader(fs, names, types, lens, rows, recLen);
            }
            // справочник улиц (у быта)
            string tmpStreets = kind.Streets == null ? null : kind.Streets + ".tmp";
            if (tmpStreets != null) {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("# код\tнаселённый пункт\tулица\tсельсовет");
                foreach (string[] s in streets) sb.AppendLine(string.Join("\t", s));
                File.WriteAllText(tmpStreets, sb.ToString(), new UTF8Encoding(true));
            }
            // сведения о копии (без данных абонентов)
            StringBuilder info = new StringBuilder();
            info.AppendLine("Версия=" + Version);
            info.AppendLine("Файл=" + src.FullName);
            info.AppendLine("Размер=" + src.Length);
            info.AppendLine("Изменён=" + src.LastWriteTimeUtc.Ticks);
            info.AppendLine("ФайлДата=" + src.LastWriteTime.ToString("dd.MM.yyyy HH:mm"));
            info.AppendLine("Создана=" + DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            info.AppendLine("Строк=" + rows);
            info.AppendLine("Улиц=" + streets.Count);
            List<string> map = new List<string>();
            foreach (Column c in cols) map.Add(c.Field + ":" + c.Header);
            info.AppendLine("Столбцы=" + string.Join(";", map.ToArray()));
            string tmpInfo = kind.Info + ".tmp";
            File.WriteAllText(tmpInfo, info.ToString(), new UTF8Encoding(true));
            Replace(tmpTable, kind.Table);
            if (tmpStreets != null) Replace(tmpStreets, kind.Streets);
            Replace(tmpInfo, kind.Info);
        }
        return "Индексы выгрузки " + kind.Name + " созданы: " + rows.ToString("#,0") + " записей" + (kind.Streets != null ? ", улиц " + streets.Count.ToString("#,0") : "")
            + (skipped > 0 ? ", строк без «" + kind.Keys[0] + "» пропущено " + skipped.ToString("#,0") : "") + ", "
            + (DateTime.Now - t0).TotalSeconds.ToString("0") + " с.";
    }

    static void Replace(string tmp, string target) {
        if (File.Exists(target)) File.Delete(target);
        File.Move(tmp, target);
    }

    static void Put(byte[] rec, int pos, int len, Encoding enc, string v) {
        byte[] b = enc.GetBytes(v);
        Buffer.BlockCopy(b, 0, rec, pos, Math.Min(len, b.Length));
    }

    // заголовок DBF (FoxPro, кодировка 1251 - метка 0xC9)
    static void WriteHeader(Stream s, List<string> names, List<char> types, List<int> lens, int count, int recLen) {
        int headerLen = 32 + 32 * names.Count + 1;
        byte[] h = new byte[headerLen];
        DateTime now = DateTime.Now;
        h[0] = 0x03;
        h[1] = (byte)(now.Year % 100);
        h[2] = (byte)now.Month;
        h[3] = (byte)now.Day;
        BitConverter.GetBytes(count).CopyTo(h, 4);
        BitConverter.GetBytes((ushort)headerLen).CopyTo(h, 8);
        BitConverter.GetBytes((ushort)recLen).CopyTo(h, 10);
        h[29] = 0xC9;
        int offset = 1;
        for (int i = 0; i < names.Count; i++) {
            int p = 32 + 32 * i;
            byte[] n = Encoding.ASCII.GetBytes(names[i]);
            Buffer.BlockCopy(n, 0, h, p, Math.Min(10, n.Length));
            h[p + 11] = (byte)types[i];
            BitConverter.GetBytes(offset).CopyTo(h, p + 12);
            h[p + 16] = (byte)lens[i];
            offset += lens[i];
        }
        h[headerLen - 1] = 0x0D;
        s.Write(h, 0, h.Length);
    }

    // справочник улиц копии: код -> {н.п., улица, сельсовет}
    public static Dictionary<string, string[]> ReadStreets(string file) {
        Dictionary<string, string[]> d = new Dictionary<string, string[]>();
        foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
            if (line.StartsWith("#")) continue;
            string[] p = line.Split('\t');
            if (p.Length >= 4) d[p[0].Trim()] = new string[] { p[1], p[2], p[3] };
        }
        return d;
    }

    // подписи полей копии: поле -> заголовок выгрузки
    public static Dictionary<string, string> Labels() { return Labels(VygKind.Byt); }

    public static Dictionary<string, string> Labels(VygKind kind) {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string cols;
        Istochnik.ReadInfo(kind.Info).TryGetValue("Столбцы", out cols);
        if (cols != null)
            foreach (string pair in cols.Split(';')) {
                int c = pair.IndexOf(':');
                if (c > 0) d[pair.Substring(0, c)] = pair.Substring(c + 1);
            }
        return d;
    }
}
// ==== КОНЕЦ ОБЩЕГО БЛОКА ВЫГРУЗОК ====
