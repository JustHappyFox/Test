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
        if (eocd < 0) throw new InvalidDataException("Это не файл Excel .xlsx (не zip). Сохраните выгрузку в Excel как «Книга Excel (*.xlsx)».");
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
        "КАТ_ГР", "KAT_GR", "ПР_ПЛ", "PR_PL" };

    public class Column {
        public int Source;           // номер столбца в выгрузке
        public string Header;        // как подписан в выгрузке
        public string Field;         // поле копии
        public bool Date;
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
        if (c.Field == "NAB") {
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

    // Пересоздать копию выгрузки быта. Пишется во временные файлы, готовое заменяет старое (прервалось - старое цело).
    // Возвращает итог для окна.
    public static string Build(string xlsx, VygProgress progress) {
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
        using (XlsxBook book = new XlsxBook(xlsx)) {
            if (progress != null) progress("Читаю выгрузку быта: тексты...");
            book.LoadShared();
            string sheet = book.FirstSheet();
            string[] header = book.Header(sheet, out headerRow);
            cols = MapColumns(header);
            Column nab = null;
            foreach (Column c in cols) {
                if (c.Field == "NAB") nab = c;
                if (c.Field == "SELSOVET") ss = c.Source;
                if (c.Field == "NP") np = c.Source;
                if (c.Field == "ULICA") ul = c.Source;
            }
            if (nab == null)
                throw new InvalidDataException("В первой строке выгрузки нет столбца «НОМЕР» (лицевой). Заголовки: " + string.Join(", ", header));
            int width = header.Length;
            // проход 1: длины полей, коды улиц
            int seen = 0;
            book.ReadRows(sheet, width, delegate(int n, string[] cells, bool[] num) {
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
                if (++seen % 20000 == 0 && progress != null) progress("Читаю выгрузку быта: " + seen.ToString("#,0") + " записей...");
            });
            foreach (Column c in cols) if (c.Date) c.Length = 8;
            // проход 2: запись копии
            string tmpTable = Istochnik.BytTable + ".tmp";
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
                book.ReadRows(sheet, width, delegate(int n, string[] cells, bool[] num) {
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
                    if (written % 20000 == 0 && progress != null) progress("Создаю индексы выгрузки быта: " + written.ToString("#,0") + " из " + seen.ToString("#,0") + "...");
                });
                fs.WriteByte(0x1A);
                rows = written;
                fs.Seek(0, SeekOrigin.Begin);
                WriteHeader(fs, names, types, lens, rows, recLen);
            }
            // справочник улиц
            string tmpStreets = Istochnik.BytStreets + ".tmp";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("# код\tнаселённый пункт\tулица\tсельсовет");
            foreach (string[] s in streets) sb.AppendLine(string.Join("\t", s));
            File.WriteAllText(tmpStreets, sb.ToString(), new UTF8Encoding(true));
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
            string tmpInfo = Istochnik.BytInfo + ".tmp";
            File.WriteAllText(tmpInfo, info.ToString(), new UTF8Encoding(true));
            Replace(tmpTable, Istochnik.BytTable);
            Replace(tmpStreets, Istochnik.BytStreets);
            Replace(tmpInfo, Istochnik.BytInfo);
        }
        return "Индексы выгрузки быта созданы: " + rows.ToString("#,0") + " записей, улиц " + streets.Count.ToString("#,0")
            + (skipped > 0 ? ", строк без лицевого пропущено " + skipped.ToString("#,0") : "") + ", "
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
    public static Dictionary<string, string> Labels() {
        Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string cols;
        Istochnik.ReadInfo(Istochnik.BytInfo).TryGetValue("Столбцы", out cols);
        if (cols != null)
            foreach (string pair in cols.Split(';')) {
                int c = pair.IndexOf(':');
                if (c > 0) d[pair.Substring(0, c)] = pair.Substring(c + 1);
            }
        return d;
    }
}
// ==== КОНЕЦ ОБЩЕГО БЛОКА ВЫГРУЗОК ====
