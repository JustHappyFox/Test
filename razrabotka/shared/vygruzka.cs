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
// ---- 26.11: юр база (выгрузка юр лиц, таблица точек учёта FTU) - для поиска юр (меню программ) и справки юр («Договор») ----
// KP договор, NAB карточка, NP наименование договора, NAB_NAIM абонент (объект с адресом), NAIM точка учёта, NST № счётчика,
// TIPSCH, GODEN поверка, AMPER, VOLT, FAZA, KTT, RAEN. Читается любая выгрузка (dbf, xls, xlsx) через VygBook, один раз.
public class JurRow {
    public string Kp = "", Nab = "", Contract = "", Abon = "", Point = "", Meter = "", Type = "", Poverka = "", Amper = "", Volt = "",
        Faza = "", Ktt = "", Raen = "";
    public string Hay = "";       // для поиска по словам: всё строчными
    public Dictionary<string, string> Fields = new Dictionary<string, string>();   // все непустые поля строки (имя поля -> значение)

    public string Card { get { return Kp + (Nab.Length > 0 ? "/" + Nab : ""); } }
}

// карточка договора (договор + карточка) с её точками учёта
public class JurCard {
    public string Kp = "", Nab = "", Contract = "", Abon = "", Raen = "";
    public readonly List<JurRow> Points = new List<JurRow>();
    public string Hay = "";
    public string Card { get { return Kp + (Nab.Length > 0 ? "/" + Nab : ""); } }
}

public static class JurIndex {
    static List<JurRow> rows;
    static string stamp = "";

    // все строки юр базы (заменили файл выгрузки - читается заново); нет выгрузки - исключение с понятным текстом
    public static List<JurRow> All() {
        VygKind k = VygKind.Jur;
        string path = Istochnik.JurPath;
        string problem = k.Problem(path);
        if (problem != null) throw new Exception(problem);
        string table = k.TableFor(path);
        string st = "";
        try { FileInfo fi = new FileInfo(table); st = table + "|" + fi.Length + "|" + fi.LastWriteTimeUtc.Ticks; } catch { }
        if (rows != null && st == stamp) return rows;
        List<JurRow> r = new List<JurRow>();
        using (VygBook book = VygBook.Open(table)) {
            book.Prepare();
            int headerRow;
            string[] h = book.Header(out headerRow);
            Dictionary<string, int> col = new Dictionary<string, int>();
            for (int i = 0; i < h.Length; i++) if (h[i] != null && !col.ContainsKey(h[i].Trim().ToUpperInvariant())) col[h[i].Trim().ToUpperInvariant()] = i;
            if (!col.ContainsKey("KP")) throw new Exception("в юр базе " + Path.GetFileName(table) + " нет поля KP - нужна таблица точек учёта FTU (KP, NP, NAB, NAB_NAIM, NAIM, NST...)");
            book.ReadRows(Math.Max(h.Length, 1), delegate(int n, string[] cells, bool[] num) {
                if (n <= headerRow) return;
                JurRow x = new JurRow();
                x.Kp = Num(Cell(cells, col, "KP")); x.Nab = Num(Cell(cells, col, "NAB")); x.Contract = Quotes(Cell(cells, col, "NP"));
                x.Abon = Cell(cells, col, "NAB_NAIM"); x.Point = Cell(cells, col, "NAIM"); x.Meter = Num(Cell(cells, col, "NST"));
                x.Type = Cell(cells, col, "TIPSCH"); x.Poverka = Cell(cells, col, "GODEN"); x.Amper = Cell(cells, col, "AMPER");
                x.Volt = Num(Cell(cells, col, "VOLT")); x.Faza = Num(Cell(cells, col, "FAZA")); x.Ktt = Num(Cell(cells, col, "KTT")); x.Raen = Cell(cells, col, "RAEN");
                if (x.Poverka == "-") x.Poverka = "";
                if (x.Kp.Length == 0 && x.Meter.Length == 0) return;
                for (int i = 0; i < h.Length && i < cells.Length; i++)
                    if (h[i] != null && cells[i] != null && cells[i].Trim().Length > 0) x.Fields[h[i].Trim().ToUpperInvariant()] = cells[i].Trim();
                x.Hay = Low(x.Contract + " " + x.Abon + " " + x.Point + " " + x.Raen + " " + x.Type + " " + x.Meter + " " + x.Kp + "/" + x.Nab);
                r.Add(x);
            });
        }
        rows = r;
        stamp = st;
        return r;
    }

    static string Cell(string[] cells, Dictionary<string, int> col, string name) {
        int i;
        return col.TryGetValue(name, out i) && i < cells.Length && cells[i] != null ? cells[i].Trim() : "";
    }

    // «11118.00» -> «11118»
    static string Num(string s) {
        s = (s ?? "").Trim();
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(s, @"^(\d+)[.,]0*$");
        return m.Success ? m.Groups[1].Value : s;
    }

    // ООО "Тест" -> ООО «Тест»
    public static string Quotes(string s) {
        s = (s ?? "").Trim();
        StringBuilder b = new StringBuilder();
        bool open = true;
        foreach (char c in s) {
            if (c == '"') { b.Append(open ? '«' : '»'); open = !open; } else b.Append(c);
        }
        return b.ToString();
    }

    static string Low(string s) {
        StringBuilder b = new StringBuilder();
        foreach (char c in (s ?? "").ToLowerInvariant()) b.Append(c == 'ё' ? 'е' : char.IsLetterOrDigit(c) || c == '/' ? c : ' ');
        return " " + b.ToString() + " ";
    }

    static bool SameNum(string a, string b) { return a.TrimStart('0') == b.TrimStart('0') && a.Length > 0 && b.Length > 0; }

    public static List<JurRow> Card(string kp, string nab) {
        List<JurRow> r = new List<JurRow>();
        foreach (JurRow x in All()) if (SameNum(x.Kp, kp) && (nab.Length == 0 || SameNum(x.Nab, nab))) r.Add(x);
        return r;
    }

    // № счётчика: точно (без ведущих нулей), нет - по части номера (от 4 цифр)
    public static List<JurRow> ByMeter(string no) {
        string d = System.Text.RegularExpressions.Regex.Replace(no ?? "", @"\D", "");
        List<JurRow> r = new List<JurRow>();
        if (d.Length == 0) return r;
        foreach (JurRow x in All()) if (SameNum(x.Meter, d)) r.Add(x);
        if (r.Count == 0 && d.Length >= 4) foreach (JurRow x in All()) if (x.Meter.Contains(d)) r.Add(x);
        return r;
    }

    // поиск: num - договор, «договор/карточка» (или через пробел, дефис) или № счётчика; text - слова (название, адрес, точка учёта)
    public static List<JurRow> Find(string num, string text, int max) {
        List<JurRow> r = new List<JurRow>();
        num = (num ?? "").Trim();
        string kp = "", nab = "", digits = "";
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(num, @"^(\d{1,7})\s*[/\\\-\s]\s*(\d{1,6})$");
        if (m.Success) { kp = m.Groups[1].Value; nab = m.Groups[2].Value; }
        else if (System.Text.RegularExpressions.Regex.IsMatch(num, @"^\d+$")) digits = num;
        else if (num.Length > 0) text = (num + " " + text).Trim();
        List<string> words = new List<string>();
        foreach (string w in Low(text).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) words.Add(w);
        if (kp.Length == 0 && digits.Length == 0 && words.Count == 0) return r;
        string part = digits.TrimStart('0').Length > 0 ? digits.TrimStart('0') : digits;
        foreach (JurRow x in All()) {
            if (kp.Length > 0 && !(SameNum(x.Kp, kp) && SameNum(x.Nab, nab))) continue;
            // одни цифры: договор с таким номером или счётчик с этими цифрами (от 4)
            if (digits.Length > 0 && !SameNum(x.Kp, digits) && !(digits.Length >= 4 && x.Meter.Contains(part))) continue;
            bool all = true;
            foreach (string w in words) if (x.Hay.IndexOf(w, StringComparison.Ordinal) < 0) { all = false; break; }
            if (!all) continue;
            r.Add(x);
            if (r.Count >= max) break;
        }
        return r;
    }

    // 27.11: поиск карточками - договор: все его карточки; договор/карточка - она; № счётчика - карточки с этим счётчиком;
    // слова (адрес, название) - карточки, где есть все слова (в наименовании, абоненте, районе или точках учёта)
    public static List<JurCard> Cards(string num, string text, int max) {
        List<JurCard> r = new List<JurCard>();
        num = (num ?? "").Trim();
        string kp = "", nab = "", digits = "";
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(num, @"^(\d{1,7})\s*[/\\\-\s]\s*(\d{1,6})$");
        if (m.Success) { kp = m.Groups[1].Value; nab = m.Groups[2].Value; }
        else if (System.Text.RegularExpressions.Regex.IsMatch(num, @"^\d+$")) digits = num;
        else if (num.Length > 0) text = (num + " " + text).Trim();
        List<string> words = new List<string>();
        foreach (string w in Low(text).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) words.Add(w);
        if (kp.Length == 0 && digits.Length == 0 && words.Count == 0) return r;
        // цифры: есть такой договор - его карточки, нет - счётчик (от 4 цифр)
        bool isContract = false;
        if (digits.Length > 0) foreach (JurRow x in All()) if (SameNum(x.Kp, digits)) { isContract = true; break; }
        string part = digits.TrimStart('0').Length > 0 ? digits.TrimStart('0') : digits;
        Dictionary<string, JurCard> byKey = new Dictionary<string, JurCard>();
        List<JurCard> order = new List<JurCard>();
        Dictionary<string, bool> meterHit = new Dictionary<string, bool>();
        foreach (JurRow x in All()) {
            if (kp.Length > 0 && !(SameNum(x.Kp, kp) && SameNum(x.Nab, nab))) continue;
            if (digits.Length > 0 && isContract && !SameNum(x.Kp, digits)) continue;
            string key = x.Kp + "/" + x.Nab;
            JurCard c;
            if (!byKey.TryGetValue(key, out c)) {
                c = new JurCard(); c.Kp = x.Kp; c.Nab = x.Nab; c.Contract = x.Contract; c.Abon = x.Abon; c.Raen = x.Raen;
                byKey[key] = c; order.Add(c);
            }
            c.Points.Add(x);
            c.Hay += x.Hay;
            if (digits.Length > 0 && !isContract && digits.Length >= 4 && x.Meter.Contains(part)) meterHit[key] = true;
        }
        foreach (JurCard c in order) {
            if (digits.Length > 0 && !isContract && !meterHit.ContainsKey(c.Kp + "/" + c.Nab)) continue;
            bool all = true;
            foreach (string w in words) if (c.Hay.IndexOf(w, StringComparison.Ordinal) < 0) { all = false; break; }
            if (!all) continue;
            r.Add(c);
            if (r.Count >= max) break;
        }
        r.Sort(delegate(JurCard a, JurCard b) {
            long ka, kb, na, nb;
            long.TryParse(a.Kp, out ka); long.TryParse(b.Kp, out kb); long.TryParse(a.Nab, out na); long.TryParse(b.Nab, out nb);
            return ka != kb ? ka.CompareTo(kb) : na.CompareTo(nb);
        });
        return r;
    }

    // названия полей юр базы для «данных счётчика» (незнакомые - как в базе)
    public static readonly string[][] FieldNames = {
        new string[] { "KP", "Договор" }, new string[] { "NAB", "Карточка" }, new string[] { "NP", "Наименование договора" },
        new string[] { "NAB_NAIM", "Абонент (объект, адрес)" }, new string[] { "NAIM", "Точка учёта" }, new string[] { "NST", "№ счётчика" },
        new string[] { "TIPSCH", "Тип счётчика" }, new string[] { "VID", "Вид счётчика" }, new string[] { "FAZA", "Фазность" },
        new string[] { "AMPER", "Ток, А" }, new string[] { "VOLT", "Напряжение, В" }, new string[] { "GODEN", "Поверка (год-квартал)" },
        new string[] { "SHKALA", "Разрядность (шкала)" }, new string[] { "KTT", "Коэффициент ТТ" }, new string[] { "KTU", "KTU" },
        new string[] { "KPOTR", "KPOTR" }, new string[] { "NT1", "NT1" }, new string[] { "NT2", "NT2" }, new string[] { "NT3", "NT3" },
        new string[] { "NAPR", "NAPR" }, new string[] { "ASKUE", "АСКУЭ" }, new string[] { "BAL", "Баланс" }, new string[] { "NAIM_RES", "РЭС" },
        new string[] { "KOD_TP", "Код ТП" }, new string[] { "NAIM_TP", "ТП" }, new string[] { "ADR_TP", "Адрес ТП" }, new string[] { "RAEN", "Район" },
        new string[] { "GRUP", "GRUP" }, new string[] { "UROV", "UROV" } };

    // данные счётчика: известные поля по порядку, потом остальные
    public static List<string[]> Details(JurRow x) {
        List<string[]> r = new List<string[]>();
        Dictionary<string, bool> done = new Dictionary<string, bool>();
        foreach (string[] f in FieldNames) {
            done[f[0]] = true;
            string v;
            if (x.Fields.TryGetValue(f[0], out v) && v.Length > 0) r.Add(new string[] { f[1], f[0] == "NP" ? Quotes(v) : v });
            if (f[0] == "NAB_NAIM") { string a = AddressOf(x); if (a.Length > 0) r.Add(new string[] { "Адрес (собран)", a }); }
        }
        foreach (KeyValuePair<string, string> kv in x.Fields) if (!done.ContainsKey(kv.Key)) r.Add(new string[] { kv.Key, kv.Value });
        return r;
    }

    // «Нераспроданные квартиры п.Макетино,ул.Полевая,34» -> «п. Макетино, ул. Полевая, 34»; район - спереди
    public static string AddressOf(JurRow x) {
        string a = x.Abon ?? "";
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(a, @"(?:^|[\s,])((?:п|д|аг|г|гп|г\.п|пос|дер|ул|пр|пер)\.\s*\S.*)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        string tail = m.Success ? m.Groups[1].Value : "";
        tail = System.Text.RegularExpressions.Regex.Replace(tail, @"\b(п|д|аг|г|гп|пос|дер|ул|пр|пер)\.\s*", "$1. ");
        tail = System.Text.RegularExpressions.Regex.Replace(tail, @"\s*,\s*", ", ").Trim().TrimEnd(',', '.');
        string r = x.Raen.Trim();
        if (r.Length > 0) r = System.Text.RegularExpressions.Regex.Replace(r, @"\s+район$", " р-н", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return (r.Length > 0 && tail.Length > 0 && tail.IndexOf(r, StringComparison.OrdinalIgnoreCase) < 0 ? r + ", " : "") + tail;
    }

    // строка прибора учёта для справки: точка|тип|№|U|I|показания|пов. кв.|пов. год|Т/Т|К расч
    public static string PuLine(JurRow x, string zeros) {
        string u = x.Volt == "400" || x.Faza == "3" ? "3х230/400" : x.Volt.Length > 0 ? x.Volt : "";
        string i = System.Text.RegularExpressions.Regex.Replace(x.Amper ?? "", @"^\s*(\d+(?:[.,]\d+)?)\s*[-/]\s*(\d+(?:[.,]\d+)?)\s*$", "$1($2)");
        string kv = "", god = "";
        System.Text.RegularExpressions.Match p = System.Text.RegularExpressions.Regex.Match(x.Poverka ?? "", @"(\d{4})\D*(\d)?");
        if (p.Success) { god = p.Groups[1].Value; kv = p.Groups[2].Value; }
        double k;
        bool tt = double.TryParse(x.Ktt, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out k) && k > 1;
        string tts = tt ? (k * 5).ToString(System.Globalization.CultureInfo.InvariantCulture) + "/5" : "-";
        return string.Join("|", new string[] { x.Point, x.Type, x.Meter, u, i, zeros, kv, god, tts, tt ? x.Ktt : "1" });
    }
}


// 27.11: окно поиска по юр базе - одно для меню программ и «Договора». Список карточек (договор - все его карточки, адрес -
// только нужные) -> двойной щелчок: точки учёта карточки -> двойной щелчок: данные счётчика. «← Назад» / Backspace - уровень вверх.
// Оформление (тема, шрифты, кнопки) - программа после создания (Look / St), здесь - только устройство окна.
public delegate void JurCardAction(JurCard c);

public class JurBrowser : Form {
    readonly TextBox num = new TextBox(), words = new TextBox();
    readonly ListView list = new ListView();
    readonly Timer delay = new Timer();
    public readonly Label TitleLabel = new Label(), Caption1 = new Label(), Caption2 = new Label(), PathLabel = new Label(), Info = new Label();
    public readonly Button BackButton = new Button(), CopyButton = new Button(), ActionButton = new Button(), CloseButton = new Button();
    public JurCardAction CardAction;          // «Справка юр по карточке» («Договор»); null - кнопки нет
    public Color HintColor = Color.Gray, GoodColor = Color.Green, BadColor = Color.Firebrick;
    int level;                                // 0 - карточки, 1 - точки учёта карточки, 2 - данные счётчика
    Panel topPanel;
    List<JurCard> cards = new List<JurCard>();
    JurCard card;
    JurRow point;

    public JurBrowser() {
        Text = "Поиск юр - по юр базе";
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1240, wa.Width), Math.Min(780, wa.Height));
        MinimumSize = new Size(Math.Min(860, wa.Width), Math.Min(480, wa.Height));
        KeyPreview = true;
        Panel top = new Panel();
        top.Dock = DockStyle.Top;
        top.Height = 120;
        TitleLabel.Text = "Поиск юр";
        TitleLabel.Font = new Font("Segoe UI", 16f, FontStyle.Bold);
        TitleLabel.AutoSize = true;
        TitleLabel.Location = new Point(18, 6);
        Caption1.Text = "Договор, договор/карточка (11118/26) или № счётчика";
        Caption2.Text = "Адрес, название, точка учёта - слова через пробел";
        num.Font = words.Font = new Font("Segoe UI", 11f);
        PathLabel.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
        PathLabel.AutoEllipsis = true;
        BackButton.Text = "← Назад";
        BackButton.Click += delegate { Up(); };
        top.Controls.AddRange(new Control[] { TitleLabel, Caption1, Caption2, num, words, BackButton, PathLabel });
        topPanel = top;
        top.Resize += delegate { PlaceTop(); };
        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.GridLines = true;
        list.DoubleClick += delegate { Down(); };
        list.Resize += delegate { FillLast(); };
        list.KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; e.Handled = true; Down(); }
            else if (e.KeyCode == Keys.Back) { e.SuppressKeyPress = true; e.Handled = true; Up(); }
        };
        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 58;
        CopyButton.Text = "Копировать";
        CopyButton.Click += delegate { Copy(); };
        ActionButton.Text = "Справка юр по карточке";
        ActionButton.Click += delegate { JurCard c = card ?? Selected(); if (c != null && CardAction != null) CardAction(c); };
        CloseButton.Text = "Закрыть";
        CloseButton.Click += delegate { Close(); };
        Info.AutoEllipsis = true;
        bottom.Controls.AddRange(new Control[] { CopyButton, ActionButton, CloseButton, Info });
        bottom.Resize += delegate {
            CloseButton.SetBounds(bottom.ClientSize.Width - 18 - 110, 11, 110, 36);
            ActionButton.SetBounds(CloseButton.Left - 10 - 220, 11, 220, 36);
            int left = ActionButton.Visible ? ActionButton.Left : CloseButton.Left;
            CopyButton.SetBounds(left - 10 - 140, 11, 140, 36);
            Info.SetBounds(18, 11, Math.Max(10, CopyButton.Left - 32), 36);
        };
        Controls.Add(list);
        Controls.Add(top);
        Controls.Add(bottom);
        delay.Interval = 350;
        delay.Tick += delegate { delay.Stop(); Search(); };
        num.TextChanged += delegate { delay.Stop(); delay.Start(); };
        words.TextChanged += delegate { delay.Stop(); delay.Start(); };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; if (level > 0) Up(); else Close(); }
            else if (e.KeyCode == Keys.Enter && (ActiveControl == num || ActiveControl == words)) {
                e.SuppressKeyPress = true; delay.Stop(); Search();
                if (list.Items.Count > 0) { list.Focus(); list.Items[0].Selected = true; list.Items[0].Focused = true; }
            }
        };
        Load += delegate { ActionButton.Visible = CardAction != null; ShowLevel(); };
        Shown += delegate { Hint(); num.Focus(); };
    }

    void PlaceTop() {
        Panel top = topPanel;
        int w = (top.ClientSize.Width - 56) / 2;
        Caption1.SetBounds(18, 42, w, 18);
        num.SetBounds(18, 60, w, 26);
        Caption2.SetBounds(38 + w, 42, w, 18);
        words.SetBounds(38 + w, 60, w, 26);
        BackButton.SetBounds(18, 90, 110, 28);
        int x = BackButton.Visible ? 138 : 18;
        PathLabel.SetBounds(x, 94, Math.Max(10, top.ClientSize.Width - x - 18), 22);
    }

    public string Num { get { return num.Text; } set { num.Text = value; } }
    public string Words { get { return words.Text; } set { words.Text = value; } }
    public int Level { get { return level; } }
    public int Rows { get { return list.Items.Count; } }
    public string PathText { get { return PathLabel.Text; } }
    public string InfoText { get { return Info.Text; } }
    public void Select(int i) { if (i < list.Items.Count) { list.SelectedItems.Clear(); list.Items[i].Selected = true; list.Items[i].Focused = true; } }
    public string Cell(int row, int col) { return row < list.Items.Count && col < list.Items[row].SubItems.Count ? list.Items[row].SubItems[col].Text : ""; }

    void Say(string text, Color c) { Info.Text = text; Info.ForeColor = c; }

    void Hint() {
        try { Say("В юр базе точек учёта: " + JurIndex.All().Count + ". Впишите договор, № счётчика или адрес - список карточек сразу.", HintColor); }
        catch (Exception ex) { Say("Юр база не открылась: " + ex.Message, BadColor); }
    }

    public void Search() {
        level = 0; card = null; point = null;
        cards = new List<JurCard>();
        if (num.Text.Trim().Length == 0 && words.Text.Trim().Length == 0) { ShowLevel(); Hint(); return; }
        try { cards = JurIndex.Cards(num.Text, words.Text, 500); }
        catch (Exception ex) { ShowLevel(); Say("Юр база не открылась: " + ex.Message, BadColor); return; }
        ShowLevel();
        // номер счётчика или договор/карточка - одна карточка: сразу её точки учёта (счётчик - выделен)
        string n = num.Text.Trim();
        if (cards.Count == 1 && words.Text.Trim().Length == 0 && n.Length > 0) {
            card = cards[0];
            level = 1;
            ShowLevel();
            string d = System.Text.RegularExpressions.Regex.Replace(n, @"\D", "");
            for (int i = 0; i < card.Points.Count; i++)
                if (d.Length >= 4 && card.Points[i].Meter.Contains(d.TrimStart('0').Length > 0 ? d.TrimStart('0') : d)) { Select(i); list.EnsureVisible(i); break; }
        }
    }

    JurCard Selected() {
        if (level != 0) return card;
        ListViewItem it = list.SelectedItems.Count > 0 ? list.SelectedItems[0] : list.Items.Count == 1 ? list.Items[0] : null;
        return it == null ? null : it.Tag as JurCard;
    }

    public void Down() {
        ListViewItem it = list.SelectedItems.Count > 0 ? list.SelectedItems[0] : null;
        if (it == null) return;
        if (level == 0 && it.Tag is JurCard) { card = (JurCard)it.Tag; level = 1; ShowLevel(); if (list.Items.Count > 0) Select(0); }
        else if (level == 1 && it.Tag is JurRow) { point = (JurRow)it.Tag; level = 2; ShowLevel(); }
    }

    public void Up() {
        if (level == 2) { level = 1; JurRow was = point; point = null; ShowLevel(); if (card != null) Select(card.Points.IndexOf(was)); }
        else if (level == 1) { JurCard was = card; level = 0; card = null; ShowLevel(); Select(cards.IndexOf(was)); }
        list.Focus();
    }

    void Columns(string[] caps, int[] ws) {
        list.Columns.Clear();
        for (int i = 0; i < caps.Length; i++) list.Columns.Add(caps[i], ws[i]);
    }

    // последний столбец - до правого края (иначе справа от шапки - пустая светлая полоса)
    void FillLast() {
        if (list.Columns.Count == 0) return;
        int used = 0;
        for (int i = 0; i < list.Columns.Count - 1; i++) used += list.Columns[i].Width;
        int room = list.ClientSize.Width - used - 2;
        ColumnHeader last = list.Columns[list.Columns.Count - 1];
        if (room > last.Width) last.Width = room;
    }

    void ShowLevel() {
        list.BeginUpdate();
        try {
            list.Items.Clear();
            if (level == 0) {
                Columns(new string[] { "Договор", "Карт.", "Наименование договора", "Абонент (объект, адрес)", "Точек учёта", "Район" }, new int[] { 80, 60, 260, 470, 100, 140 });
                foreach (JurCard c in cards) {
                    ListViewItem it = new ListViewItem(new string[] { c.Kp, c.Nab, c.Contract, c.Abon, c.Points.Count.ToString(), c.Raen });
                    it.Tag = c;
                    list.Items.Add(it);
                }
                Dictionary<string, bool> dogs = new Dictionary<string, bool>();
                foreach (JurCard c in cards) dogs[c.Kp] = true;
                PathLabel.Text = cards.Count == 0 ? "" : "Карточки: " + cards.Count + ", договоров: " + dogs.Count;
                if (num.Text.Trim().Length > 0 || words.Text.Trim().Length > 0)
                    Say(cards.Count == 0 ? "Ничего не нашлось - проверьте номер или уберите лишнее слово." :
                        "Двойной щелчок (или Enter) по карточке - её точки учёта." + (cards.Count >= 500 ? " Показаны первые 500 карточек." : ""), cards.Count == 0 ? BadColor : HintColor);
            } else if (level == 1 && card != null) {
                Columns(new string[] { "№", "Точка учёта", "Тип счётчика", "№ счётчика", "Поверка", "I, А", "U, В", "Ктт" }, new int[] { 40, 260, 200, 130, 90, 80, 70, 60 });
                for (int i = 0; i < card.Points.Count; i++) {
                    JurRow x = card.Points[i];
                    ListViewItem it = new ListViewItem(new string[] { (i + 1).ToString(), x.Point, x.Type, x.Meter, x.Poverka, x.Amper, x.Volt, x.Ktt });
                    it.Tag = x;
                    list.Items.Add(it);
                }
                PathLabel.Text = "Договор " + card.Card + " - " + card.Contract + " - " + card.Abon;
                Say("Точек учёта: " + card.Points.Count + ". Двойной щелчок (или Enter) по точке - данные счётчика; Backspace - к карточкам.", HintColor);
            } else if (level == 2 && point != null) {
                Columns(new string[] { "Что", "Значение" }, new int[] { 260, 760 });
                foreach (string[] d in JurIndex.Details(point)) list.Items.Add(new ListViewItem(d));
                PathLabel.Text = "Договор " + point.Card + " - " + point.Point + " - счётчик № " + point.Meter;
                Say("Данные счётчика из юр базы. Backspace - к точкам учёта карточки.", HintColor);
            }
        } finally { list.EndUpdate(); }
        BackButton.Visible = level > 0;
        ActionButton.Visible = CardAction != null;
        if (topPanel != null) PlaceTop();
        FillLast();
    }

    void Copy() {
        StringBuilder sb = new StringBuilder();
        if (level == 2) foreach (ListViewItem it in list.Items) sb.Append(it.SubItems[0].Text).Append(": ").Append(it.SubItems[1].Text).Append("\r\n");
        else if (list.SelectedItems.Count > 0) {
            ListViewItem it = list.SelectedItems[0];
            for (int i = 0; i < it.SubItems.Count; i++) sb.Append(i > 0 ? "; " : "").Append(list.Columns[i].Text).Append(": ").Append(it.SubItems[i].Text);
        }
        if (sb.Length == 0) return;
        try { Clipboard.SetText(sb.ToString()); Say("Скопировано в буфер обмена.", GoodColor); } catch { }
    }
}
// ==== КОНЕЦ ОБЩЕГО БЛОКА ВЫГРУЗОК ====
