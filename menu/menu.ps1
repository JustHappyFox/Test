# Меню программ: одно окно, из которого запускаются «Поиск абонента», «Оформление договора» и «Фото с телефона».
# Запуск: zapusk.vbs (без чёрного окна) или zapusk.bat. Ярлык на рабочий стол - в самом меню.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @([System.Windows.Forms.Form].Assembly.Location, [System.Drawing.Point].Assembly.Location, [System.Xml.XmlDocument].Assembly.Location)

$code = @'
using System.Xml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

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
// Книга выгрузки любого вида (смотрится по содержимому, не по расширению): .xlsx; .xls - Excel 97-2003, Excel 5.0/95 и
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

// Меню программ: одно окно (и один ярлык на рабочем столе), из которого запускаются «Поиск абонента», «Оформление договора»
// и «Фото с телефона». Меню только запускает их zapusk.vbs - в папках программ ничего не пишет. Где какая программа -
// находит само (рядом с собой, на рабочем столе, в «Документах», на дисках) или указывается вручную; запоминается в nastroiki.txt
public static class MenuApp {
    public static string Here = "";

    public static void Run(string here, string launch) {
        Here = here;
        try { Application.EnableVisualStyles(); } catch { }
        try { Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException); } catch { }
        Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e) { Fail(e.Exception); };
        Application.Run(new MenuForm(here));
    }

    // ошибка - в oshibki.txt рядом с меню и окно (номера от 6 цифр скрыты)
    public static void Fail(Exception ex) {
        string text = System.Text.RegularExpressions.Regex.Replace(ex.ToString(), @"\d{6,}", "######");
        try {
            File.AppendAllText(Path.Combine(Here, "oshibki.txt"), "=== " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " ===\r\nWindows " + Environment.OSVersion.Version
                + ", .NET " + Environment.Version + "\r\n" + text + "\r\n\r\n", new UTF8Encoding(true));
        } catch { }
        MessageBox.Show("Ошибка: " + ex.Message + "\n\nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", "Программы - ошибка");
    }
}

// цвета - общей темы программ (Look: тёмная / светлая, как в «Договоре»)
public static class Pal {
    public static Color Back { get { return Look.Back; } }
    public static Color Tile { get { return Look.Surface; } }
    public static Color TileHover { get { return Look.Hover; } }
    public static Color Border { get { return Look.Border; } }
    public static Color Text { get { return Look.Text; } }
    public static Color Hint { get { return Look.Hint; } }
    public static Color Bad { get { return Look.Bad; } }
    public static Color Good { get { return Look.Good; } }
    public static Color Link { get { return Look.Link; } }
    public static Color Warn { get { return Look.Warn; } }
    public static readonly Font Title = new Font("Segoe UI", 17f, FontStyle.Bold), Big = new Font("Segoe UI", 13.5f, FontStyle.Bold),
        Base = new Font("Segoe UI", 10f), Small = new Font("Segoe UI", 9f), Number = new Font("Segoe UI", 22f, FontStyle.Bold);
}

// программа меню: скрипт, по которому узнаётся её папка, и файлы, которые появляются, когда ею пользуются
public class Prog {
    public string Key, Title, About, Script;
    public string[] Work;            // рабочие файлы/папки: такая папка - настоящая, а не свежераспакованная копия
    public Color Mark;
    public string Dir = "";          // где лежит (пусто - не найдена)
    public bool Auto;                // найдена сама

    public static Prog[] All() {
        Prog a = new Prog(), b = new Prog(), c = new Prog();
        a.Key = "Поиск"; a.Title = "Поиск абонента"; a.Script = "prosmotr.ps1"; a.Mark = Color.FromArgb(99, 160, 235);
        a.About = "лицевой, адрес, счётчик, Ф.И.О.; карточка, наряды, обходы, оплаты";
        a.Work = new string[] { "poisk.txt", "indeks_fio.dat" };
        b.Key = "Договор"; b.Title = "Оформление договора"; b.Script = "dogovor.ps1"; b.Mark = Color.FromArgb(112, 200, 142);
        b.About = "договор, справка, наряд на подключение, отчёты пломб";
        b.Work = new string[] { "nastroiki.txt", "Договоры", "Справки" };
        c.Key = "Фото"; c.Title = "Фото с телефона"; c.Script = "foto.ps1"; c.Mark = Color.FromArgb(230, 170, 90);
        c.About = "фото, скрины, видео с телефона - в папки год\\месяц\\число";
        c.Work = new string[] { "nastroiki.txt", "telefony.txt" };
        return new Prog[] { a, b, c };
    }

    public bool Fits(string dir) {
        try { return dir != null && dir.Length > 1 && File.Exists(Path.Combine(dir, Script)); } catch { return false; }
    }

    // чем больше рабочих файлов и чем свежее - тем вероятнее, что это та папка, где работают
    public long Score(string dir) {
        int work = 0;
        DateTime t = DateTime.MinValue;
        foreach (string w in Work) {
            try {
                string p = Path.Combine(dir, w);
                DateTime m = DateTime.MinValue;
                if (File.Exists(p)) m = File.GetLastWriteTime(p);
                else if (Directory.Exists(p)) m = Directory.GetLastWriteTime(p);
                else continue;
                work++;
                if (m > t) t = m;
            } catch { }
        }
        try { DateTime s = File.GetLastWriteTime(Path.Combine(dir, Script)); if (work == 0 && s > t) t = s; } catch { }
        return work * 100000000000000L + Math.Max(0L, t.Ticks / 10000000L);
    }
}

// Поиск папок программ: папка меню и соседние, на ступень выше, рабочий стол, «Документы», корни дисков - по две ступени вглубь
public static class Finder {
    public static void Find(string here, Prog[] progs) {
        List<string> roots = new List<string>();
        try {
            string full = Path.GetFullPath(here);
            roots.Add(full);
            string up = Path.GetDirectoryName(full);
            if (up != null) { roots.Add(up); string up2 = Path.GetDirectoryName(up); if (up2 != null) roots.Add(up2); }
        } catch { }
        try { roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)); } catch { }
        try { roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.Personal)); } catch { }
        try { foreach (DriveInfo d in DriveInfo.GetDrives()) if (d.DriveType == DriveType.Fixed && d.IsReady) roots.Add(d.RootDirectory.FullName); } catch { }
        long[] best = new long[progs.Length];
        string[] found = new string[progs.Length];
        Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        int budget = 6000;
        foreach (string r in roots) if (r != null && r.Length > 0) Walk(r, 2, progs, found, best, seen, ref budget);
        for (int i = 0; i < progs.Length; i++) {
            if (found[i] == null) continue;
            progs[i].Dir = found[i];
            progs[i].Auto = true;
        }
    }

    static void Walk(string dir, int depth, Prog[] progs, string[] found, long[] best, Dictionary<string, bool> seen, ref int budget) {
        if (budget-- <= 0 || seen.ContainsKey(dir)) return;
        seen[dir] = true;
        for (int i = 0; i < progs.Length; i++) {
            if (progs[i].Dir.Length > 0 || !progs[i].Fits(dir)) continue;
            long s = progs[i].Score(dir);
            if (found[i] == null || s > best[i]) { found[i] = dir; best[i] = s; }
        }
        if (depth == 0) return;
        string[] subs;
        try { subs = Directory.GetDirectories(dir); } catch { return; }
        foreach (string sub in subs) {
            string n = Path.GetFileName(sub).ToLowerInvariant();
            if (n.StartsWith("$") || n.StartsWith(".") || n == "windows" || n == "program files" || n == "program files (x86)" || n == "programdata"
                || n == "system volume information" || n == "recovery" || n == "appdata" || n == "proc" || n == "sys"
                || n == "договоры" || n == "справки" || n == "наряды" || n == "пломбы" || n == "kesh-programmy") continue;
            try { if ((File.GetAttributes(sub) & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue; } catch { continue; }
            Walk(sub, depth - 1, progs, found, best, seen, ref budget);
        }
    }
}

// настройки меню: «ключ=значение» в nastroiki.txt
public class MenuSettings {
    readonly string file;
    public readonly Dictionary<string, string> Values = new Dictionary<string, string>();

    public MenuSettings(string file) {
        this.file = file;
        try {
            if (File.Exists(file))
                foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                    int e = line.IndexOf('=');
                    if (e > 0 && !line.StartsWith("#")) Values[line.Substring(0, e).Trim()] = line.Substring(e + 1).Trim();
                }
        } catch { }
    }

    public string Get(string k) { string v; return Values.TryGetValue(k, out v) ? v : ""; }

    public void Save() {
        try {
            List<string> lines = new List<string>();
            lines.Add("# Меню программ: где какая программа. Файл можно удалить - меню найдёт программы заново.");
            foreach (KeyValuePair<string, string> kv in Values) lines.Add(kv.Key + "=" + kv.Value);
            File.WriteAllLines(file, lines.ToArray(), new UTF8Encoding(true));
        } catch { }
    }
}

// плитка программы: номер, название, что умеет, где лежит; щелчок - запуск
public class ProgTile : Control {
    public Prog P;
    public int Index;
    public string State = "";
    public Color StateColor = Pal.Hint;
    bool hover;
    public readonly LinkLabel Folder = new LinkLabel();

    public ProgTile(Prog p, int index) {
        P = p;
        Index = index;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        Cursor = Cursors.Hand;
        BackColor = Pal.Tile;
        Folder.Text = "папка…";
        Folder.Font = Pal.Small;
        Folder.AutoSize = true;
        Folder.LinkColor = Folder.ActiveLinkColor = Pal.Link;
        Folder.BackColor = Color.Transparent;
        Folder.LinkBehavior = LinkBehavior.HoverUnderline;
        Controls.Add(Folder);
        Resize += delegate { Folder.Location = new Point(Width - Folder.Width - 14, Height - Folder.Height - 10); };
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.Clear(hover ? Pal.TileHover : Pal.Tile);
        using (Pen pen = new Pen(Pal.Border)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        using (SolidBrush b = new SolidBrush(P.Mark)) g.FillRectangle(b, 0, 0, 6, Height);
        TextRenderer.DrawText(g, (Index + 1).ToString(), Pal.Number, new Point(20, (Height - 40) / 2), P.Mark);
        TextRenderer.DrawText(g, P.Title, Pal.Big, new Point(62, 10), Pal.Text);
        TextRenderer.DrawText(g, P.About, Pal.Base, new Point(64, 40), Pal.Hint);
        Rectangle st = new Rectangle(64, Height - 28, Width - 64 - Folder.Width - 30, 20);
        TextRenderer.DrawText(g, State, Pal.Small, st, StateColor, TextFormatFlags.Left | TextFormatFlags.PathEllipsis | TextFormatFlags.SingleLine);
    }
}

public class MenuForm : Form {
    readonly string here;
    readonly MenuSettings settings;
    readonly Prog[] progs = Prog.All();
    readonly List<ProgTile> tiles = new List<ProgTile>();
    readonly CheckBox closeAfter = new CheckBox();
    readonly Label status = new Label();
    readonly LinkLabel shortcut = new LinkLabel(), again = new LinkLabel();
    LinkLabel theme;
    bool searching;

    public delegate void Starter(string file, string args);
    public static Starter Start;          // для проверки: вместо запуска - сюда

    public MenuForm(string here) {
        this.here = here;
        settings = new MenuSettings(Path.Combine(here, "nastroiki.txt"));
        // тема (nastroiki.txt «ТёмнаяТема»; нет записи - тёмная) - до создания элементов
        Look.Apply(settings.Get("ТёмнаяТема") != "0");
        Text = "Программы";
        Font = Pal.Base;
        BackColor = Pal.Back;
        ForeColor = Pal.Text;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(640, 486);
        KeyPreview = true;
        try { Icon = MenuIcon.Make(); } catch { }
        Label title = new Label();
        title.Text = "Программы";
        title.Font = Pal.Title;
        title.ForeColor = Pal.Text;
        title.SetBounds(28, 18, 400, 36);
        Label sub = new Label();
        sub.Text = "щёлкните программу (или нажмите 1, 2, 3)";
        sub.Font = Pal.Small;
        sub.ForeColor = Pal.Hint;
        sub.SetBounds(30, 54, 500, 20);
        Controls.Add(title);
        Controls.Add(sub);
        // «Настройки…» - откуда «Поиск» и «Договор» берут данные: база или Excel-выгрузки (быт, юр, наряды)
        Button data = new Button();
        data.Text = "Настройки…";
        data.Font = Pal.Base;
        data.SetBounds(ClientSize.Width - 28 - 130, 22, 130, 32);
        data.Click += delegate { OpenDataSettings(); };
        Controls.Add(data);
        Look.Button(data);
        for (int i = 0; i < progs.Length; i++) {
            ProgTile t = new ProgTile(progs[i], i);
            t.SetBounds(28, 84 + i * 104, 584, 92);
            int k = i;
            t.Click += delegate { Launch(k); };
            t.Folder.LinkClicked += delegate { PickFolder(k); };
            tiles.Add(t);
            Controls.Add(t);
        }
        closeAfter.Text = "Закрывать меню после запуска программы";
        closeAfter.ForeColor = Pal.Text;
        closeAfter.Font = Pal.Small;
        closeAfter.SetBounds(28, 396, 330, 24);
        closeAfter.Checked = settings.Get("Закрывать") != "0";
        closeAfter.CheckedChanged += delegate { settings.Values["Закрывать"] = closeAfter.Checked ? "1" : "0"; settings.Save(); };
        Controls.Add(closeAfter);
        foreach (LinkLabel l in new LinkLabel[] { shortcut, again }) {
            l.Font = Pal.Small;
            l.AutoSize = true;
            l.LinkColor = l.ActiveLinkColor = Pal.Link;
            l.LinkBehavior = LinkBehavior.HoverUnderline;
            Controls.Add(l);
        }
        shortcut.LinkClicked += delegate { MakeShortcut(); };
        theme = Look.Switch(delegate(bool dark) { settings.Values["ТёмнаяТема"] = dark ? "1" : "0"; settings.Save(); ThemePlace(); });
        Controls.Add(theme);
        again.Text = "Искать программы заново";
        again.LinkClicked += delegate { Search(true); };
        status.Font = Pal.Small;
        status.ForeColor = Pal.Hint;
        status.SetBounds(28, 454, 584, 22);
        Controls.Add(status);
        ShortcutText();
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) Close();
            int n = e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D3 ? e.KeyCode - Keys.D1 : e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad3 ? e.KeyCode - Keys.NumPad1 : -1;
            if (n >= 0) { e.SuppressKeyPress = true; Launch(n); }
        };
        // запомненные папки; нет или программы там уже нет - поискать
        bool missing = false;
        foreach (Prog p in progs) {
            string d = settings.Get(p.Key);
            if (p.Fits(d)) p.Dir = d;
            else missing = true;
        }
        Show2();
        Shown += delegate { if (missing) Search(false); };
    }

    void Place(LinkLabel l, int x, int y) { l.Location = new Point(x, y); }

    public DataSettingsForm DataForm;     // для проверки

    public void OpenDataSettings() {
        using (DataSettingsForm f = new DataSettingsForm()) {
            DataForm = f;
            f.ShowDialog(this);
            DataForm = null;
        }
    }

    void ShortcutText() {
        shortcut.Text = File.Exists(ShortcutPath) ? "Ярлык «Программы» на рабочем столе есть - обновить" : "Создать ярлык «Программы» на рабочем столе";
        Place(shortcut, 30, 426);
        ThemePlace();
        Place(again, ClientSize.Width - 28 - TextRenderer.MeasureText(again.Text, Pal.Small).Width - 4, 400);
    }

    void ThemePlace() { if (theme != null) Place(theme, ClientSize.Width - 28 - TextRenderer.MeasureText(theme.Text, Look.Small).Width - 4, 426); }

    // плитки: где лежит программа
    void Show2() {
        for (int i = 0; i < progs.Length; i++) {
            Prog p = progs[i];
            ProgTile t = tiles[i];
            if (p.Dir.Length > 0) { t.State = "папка: " + p.Dir + (p.Auto ? " (найдена сама)" : ""); t.StateColor = Pal.Hint; }
            else if (searching) { t.State = "ищу папку программы..."; t.StateColor = Pal.Hint; }
            else { t.State = "не найдена - укажите папку (где " + p.Script + "): «папка…»"; t.StateColor = Pal.Bad; }
            t.Invalidate();
        }
    }

    // поиск папок программ - в фоне (диски - секунды)
    public void Search(bool all) {
        if (searching) return;
        searching = true;
        if (all) foreach (Prog p in progs) { p.Dir = ""; p.Auto = false; }
        status.Text = "Ищу программы: рядом с меню, на рабочем столе, в «Документах», на дисках...";
        status.ForeColor = Pal.Hint;
        Show2();
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            try { Finder.Find(here, progs); } catch { }
            try {
                BeginInvoke((MethodInvoker)delegate {
                    searching = false;
                    int n = 0;
                    foreach (Prog p in progs) if (p.Dir.Length > 0) { settings.Values[p.Key] = p.Dir; n++; }
                    settings.Save();
                    status.Text = n == progs.Length ? "Все программы найдены." : "Найдено программ: " + n + " из " + progs.Length + " - для остальных укажите папку («папка…»).";
                    status.ForeColor = n == progs.Length ? Pal.Good : Pal.Bad;
                    Show2();
                });
            } catch { }
        });
        th.IsBackground = true;
        th.Start();
    }

    void PickFolder(int i) {
        Prog p = progs[i];
        using (FolderBrowserDialog d = new FolderBrowserDialog()) {
            d.Description = "Папка программы «" + p.Title + "» (где " + p.Script + " и zapusk.vbs)";
            d.ShowNewFolderButton = false;
            try { if (Directory.Exists(p.Dir)) d.SelectedPath = p.Dir; } catch { }
            if (d.ShowDialog(this) != DialogResult.OK) return;
            SetFolder(i, d.SelectedPath);
        }
    }

    public bool SetFolder(int i, string dir) {
        Prog p = progs[i];
        if (!p.Fits(dir)) {
            status.Text = "В папке " + dir + " нет " + p.Script + " - это не «" + p.Title + "».";
            status.ForeColor = Pal.Bad;
            return false;
        }
        p.Dir = dir;
        p.Auto = false;
        settings.Values[p.Key] = dir;
        settings.Save();
        status.Text = "«" + p.Title + "» - папка запомнена.";
        status.ForeColor = Pal.Good;
        Show2();
        return true;
    }

    // запуск: zapusk.vbs программы (без чёрного окна), нет его - сам скрипт через PowerShell
    public void Launch(int i) {
        if (i < 0 || i >= progs.Length) return;
        Prog p = progs[i];
        if (p.Dir.Length == 0 || !p.Fits(p.Dir)) {
            status.Text = searching ? "Ещё ищу, где «" + p.Title + "», - секунду..." : "«" + p.Title + "» не найдена - укажите папку («папка…»).";
            status.ForeColor = Pal.Bad;
            if (!searching) { p.Dir = ""; Show2(); }
            return;
        }
        string vbs = Path.Combine(p.Dir, "zapusk.vbs"), file, args;
        if (File.Exists(vbs)) { file = "wscript.exe"; args = "\"" + vbs + "\""; }
        else { file = "powershell.exe"; args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + Path.Combine(p.Dir, p.Script) + "\""; }
        try {
            if (Start != null) Start(file, args);
            else {
                ProcessStartInfo si = new ProcessStartInfo(file, args);
                si.WorkingDirectory = p.Dir;
                si.UseShellExecute = false;
                Process.Start(si);
            }
        } catch (Exception ex) {
            status.Text = "«" + p.Title + "» не запустилась: " + ex.Message;
            status.ForeColor = Pal.Bad;
            return;
        }
        status.Text = "Запускаю «" + p.Title + "»...";
        status.ForeColor = Pal.Good;
        if (closeAfter.Checked) {
            Timer t = new Timer();
            t.Interval = 700;
            t.Tick += delegate { t.Stop(); Close(); };
            t.Start();
        }
    }

    public static string ShortcutPath {
        get {
            string desk = "";
            try { desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); } catch { }
            return Path.Combine(desk, "Программы.lnk");
        }
    }

    // ярлык на рабочем столе: wscript.exe "<папка меню>\zapusk.vbs", значок - menu.ico (рисуется сам)
    public void MakeShortcut() {
        try {
            string ico = Path.Combine(here, "menu.ico");
            try { MenuIcon.Save(ico); } catch { ico = ""; }
            Type t = Type.GetTypeFromProgID("WScript.Shell");
            if (t == null) throw new Exception("в Windows нет WScript.Shell");
            object sh = Activator.CreateInstance(t);
            object lnk = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, sh, new object[] { ShortcutPath });
            Type lt = lnk.GetType();
            string wscript = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "wscript.exe");
            lt.InvokeMember("TargetPath", BindingFlags.SetProperty, null, lnk, new object[] { wscript });
            lt.InvokeMember("Arguments", BindingFlags.SetProperty, null, lnk, new object[] { "\"" + Path.Combine(here, "zapusk.vbs") + "\"" });
            lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, lnk, new object[] { here });
            lt.InvokeMember("Description", BindingFlags.SetProperty, null, lnk, new object[] { "Поиск абонента, Оформление договора, Фото с телефона" });
            if (ico.Length > 0) lt.InvokeMember("IconLocation", BindingFlags.SetProperty, null, lnk, new object[] { ico + ",0" });
            lt.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
            status.Text = "Ярлык «Программы» - на рабочем столе.";
            status.ForeColor = Pal.Good;
        } catch (Exception ex) {
            Exception e = ex.InnerException ?? ex;
            status.Text = "Ярлык не создался: " + e.Message;
            status.ForeColor = Pal.Bad;
        }
        ShortcutText();
    }

    // для проверки
    public string Dump() {
        StringBuilder sb = new StringBuilder();
        foreach (ProgTile t in tiles) sb.Append("  ").Append(t.Index + 1).Append(". ").Append(t.P.Title).Append(": ").Append(t.State).Append("\n");
        sb.Append("  внизу: ").Append(status.Text).Append("\n");
        return sb.ToString();
    }
    public bool Busy { get { return searching; } }
}

// значок меню: синий квадрат со скруглёнными углами и тремя полосками разных цветов (три программы)
// «Настройки…» меню: откуда «Поиск» и «Договор» берут данные - база (как раньше) или Excel-выгрузки (быт, юр, наряды);
// пути к выгрузкам; «Обновить индексы» - перевести выгрузку в свою быструю копию (после замены файла выгрузки).
// Настройки - в профиле пользователя (Istochnik.SettingsFile), общие для всех программ.
public class DataSettingsForm : Form {
    readonly RadioButton useBase = new RadioButton(), useExcel = new RadioButton();
    readonly TextBox[] paths = { new TextBox(), new TextBox(), new TextBox() };
    readonly Label[] notes = { new Label(), new Label(), new Label() };
    readonly Label indexState = new Label(), progress = new Label();
    readonly Button rebuild = new Button(), save = new Button(), cancel = new Button();
    static readonly string[] Kinds = { "быта (физ лица)", "юр лиц", "нарядов" };
    bool building;
    public string LastResult = "";          // для проверки: чем кончилось «Обновить индексы»

    public DataSettingsForm() {
        Text = "Настройки: откуда брать данные";
        Font = Pal.Base;
        BackColor = Pal.Back;
        ForeColor = Pal.Text;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        // маленький экран (нетбук 1024x600): строки плотнее, окно ниже - кнопки «Сохранить» и «Отмена» видны
        bool c = Screen.PrimaryScreen.WorkingArea.Height < 700;
        ClientSize = new Size(700, c ? 460 : 590);
        Label title = new Label();
        title.Text = "Откуда «Поиск» и «Договор» берут данные";
        title.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
        title.SetBounds(24, c ? 10 : 16, 650, 26);
        Controls.Add(title);
        useBase.Text = "База, как раньше (O:\\BASAFOX, база юр лиц, наряды - по индексам самих баз)";
        useExcel.Text = "Excel-выгрузки (файлы ниже)";
        useBase.SetBounds(26, c ? 40 : 50, 650, 24);
        useExcel.SetBounds(26, c ? 64 : 76, 650, 24);
        Controls.Add(useBase);
        Controls.Add(useExcel);
        int y = c ? 94 : 112;
        for (int i = 0; i < 3; i++) {
            Label cap = new Label();
            cap.Text = "Выгрузка " + Kinds[i] + (i == 1 ? " (таблица .dbf, .xls или .xlsx)" : " (.xls или .xlsx)");
            cap.ForeColor = Pal.Hint;
            cap.Font = Pal.Small;
            cap.SetBounds(24, y, 400, 18);
            paths[i].SetBounds(24, y + 20, 540, 24);
            Button pick = new Button();
            pick.Text = "Выбрать…";
            pick.SetBounds(574, y + 18, 100, 28);
            int k = i;
            pick.Click += delegate { Pick(k); };
            paths[i].TextChanged += delegate { ShowState(); };
            notes[i].Font = Pal.Small;
            notes[i].SetBounds(24, y + 46, 650, 18);
            Controls.Add(cap);
            Controls.Add(paths[i]);
            Controls.Add(pick);
            Controls.Add(notes[i]);
            y += c ? 66 : 74;
        }
        indexState.SetBounds(24, y + 2, 650, c ? 52 : 60);
        rebuild.Text = "Обновить индексы";
        rebuild.SetBounds(24, y + (c ? 56 : 68), 170, 32);
        rebuild.Click += delegate { Rebuild(); };
        progress.Font = Pal.Small;
        progress.ForeColor = Pal.Hint;
        progress.SetBounds(204, y + (c ? 56 : 68), 470, c ? 50 : 64);
        Controls.Add(indexState);
        Controls.Add(rebuild);
        Controls.Add(progress);
        save.Text = "Сохранить";
        cancel.Text = "Отмена";
        save.SetBounds(ClientSize.Width - 24 - 120 - 10 - 110, ClientSize.Height - 46, 120, 32);
        cancel.SetBounds(ClientSize.Width - 24 - 110, ClientSize.Height - 46, 110, 32);
        save.Click += delegate { if (Save()) { DialogResult = DialogResult.OK; Close(); } };
        cancel.Click += delegate { Close(); };
        Controls.Add(save);
        Controls.Add(cancel);
        AcceptButton = save;
        CancelButton = cancel;
        useBase.CheckedChanged += delegate { ShowState(); };
        useExcel.Checked = Istochnik.UseExcel;
        useBase.Checked = !useExcel.Checked;
        paths[0].Text = Istochnik.BytPath;
        paths[1].Text = Istochnik.JurPath;
        paths[2].Text = Istochnik.NaryadyPath;
        Look.Dress(this);
        Look.MainButton(save);
        ShowState();
        FormClosing += delegate(object o, FormClosingEventArgs e) {
            if (building) { e.Cancel = true; progress.Text = "Подождите: индексы ещё создаются."; }
        };
    }

    static readonly VygKind[] KindsOf = { VygKind.Byt, VygKind.Jur, VygKind.Nar };   // в порядке полей: быт, юр, наряды

    void Pick(int i) {
        using (OpenFileDialog d = new OpenFileDialog()) {
            d.Title = "Выгрузка " + Kinds[i];
            d.Filter = i == 1 ? "Таблица DBF или книга Excel (*.dbf;*.xls;*.xlsx)|*.dbf;*.xls;*.xlsx|Все файлы (*.*)|*.*" : "Книга Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
            try { if (File.Exists(paths[i].Text.Trim())) d.InitialDirectory = Path.GetDirectoryName(paths[i].Text.Trim()); } catch { }
            if (d.ShowDialog(this) == DialogResult.OK) paths[i].Text = d.FileName;
        }
    }

    public void SetPaths(string byt, string jur, string nar, bool excel) {
        paths[0].Text = byt; paths[1].Text = jur; paths[2].Text = nar;
        useExcel.Checked = excel; useBase.Checked = !excel;
    }

    // под каждым путём - есть ли файл; ниже - состояние индексов каждой выгрузки (по полям окна, ещё не сохранённым)
    void ShowState() {
        bool excel = useExcel.Checked;
        for (int i = 0; i < 3; i++) {
            string p = paths[i].Text.Trim();
            string miss = Istochnik.Missing(Kinds[i], p);
            if (p.Length == 0) { notes[i].Text = excel ? "не указана - " + (i == 0 ? "поиск абонентов работать не будет" : i == 1 ? "перевода из юр в физ не будет" : "нарядов не будет") : "не указана"; notes[i].ForeColor = excel ? Pal.Bad : Pal.Hint; }
            else if (miss != null) { notes[i].Text = "файла нет: " + p; notes[i].ForeColor = Pal.Bad; }
            else {
                string info = "";
                try { FileInfo fi = new FileInfo(p); info = "файл есть: " + (fi.Length / 1048576.0).ToString("0.0") + " МБ, изменён " + fi.LastWriteTime.ToString("dd.MM.yyyy HH:mm"); } catch { }
                if (!KindsOf[i].NeedsCopy(p)) info += " - таблица DBF, индексы не нужны";
                notes[i].Text = info;
                notes[i].ForeColor = Pal.Good;
            }
        }
        List<string> lines = new List<string>();
        bool bad = false;
        for (int i = 0; i < 3; i++) {
            string p = paths[i].Text.Trim();
            if (p.Length == 0 || !KindsOf[i].NeedsCopy(p)) continue;
            string problem = KindsOf[i].Problem(p);
            if (problem != null) { bad = true; lines.Add(problem.Replace(" Меню программ: «Настройки…» -> «Обновить индексы».", " - «Обновить индексы».").Replace(" Укажите его в меню программ: «Настройки…».", "").Replace(" Проверьте путь в меню программ: «Настройки…».", "")); }
            else lines.Add("Индексы готовы: " + KindsOf[i].About() + ".");
        }
        if (lines.Count == 0) lines.Add("Индексы: выгрузки не указаны.");
        indexState.Text = string.Join("\n", lines.ToArray());
        indexState.ForeColor = !bad ? Pal.Good : excel ? Pal.Bad : Pal.Hint;
        rebuild.Enabled = !building && ToBuild().Count > 0;
    }

    // какие выгрузки переводить в свою копию: указан файл, он есть, копия нужна (не .dbf юр базы)
    List<int> ToBuild() {
        List<int> r = new List<int>();
        for (int i = 0; i < 3; i++) {
            string p = paths[i].Text.Trim();
            if (Istochnik.Missing(Kinds[i], p) == null && KindsOf[i].NeedsCopy(p)) r.Add(i);
        }
        return r;
    }

    bool AnyProblem() {
        foreach (int i in ToBuild()) if (KindsOf[i].Problem(paths[i].Text.Trim()) != null) return true;
        return false;
    }

    bool Save() {
        bool excel = useExcel.Checked;
        try { Istochnik.Save(excel, paths[0].Text, paths[1].Text, paths[2].Text); }
        catch (Exception ex) { MessageBox.Show(this, "Настройки не сохранились:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
        if (excel && AnyProblem()
            && MessageBox.Show(this, "Индексы выгрузок не готовы:\n" + indexState.Text + "\n\nСоздать их сейчас? (Без них «Поиск» и «Договор» не найдут данные.)", Text,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
            Rebuild();
            return false;      // окно остаётся, пока индексы создаются; потом - «Сохранить» ещё раз или закрыть
        }
        return true;
    }

    // перевести выгрузки (быт, наряды, юр в .xlsx) в свои копии - по очереди, в фоне, окно показывает ход
    public void Rebuild() {
        List<int> todo = ToBuild();
        if (building || todo.Count == 0) return;
        building = true;
        rebuild.Enabled = save.Enabled = false;
        progress.ForeColor = Pal.Hint;
        progress.Text = "Создаю индексы выгрузок (один раз после замены файлов; большая выгрузка - около минуты)...";
        string[] ps = new string[3];
        for (int i = 0; i < 3; i++) ps[i] = paths[i].Text.Trim();
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            List<string> results = new List<string>();
            bool ok = true;
            foreach (int i in todo) {
                try {
                    results.Add(VygruzkaByt.Build(KindsOf[i], ps[i], delegate(string text) {
                        try { BeginInvoke((MethodInvoker)delegate { progress.Text = text; }); } catch { }
                    }));
                } catch (Exception ex) { results.Add("Индексы выгрузки " + Kinds[i] + " не созданы: " + ex.Message); ok = false; }
            }
            string result = string.Join("\n", results.ToArray());
            try {
                BeginInvoke((MethodInvoker)delegate {
                    building = false;
                    save.Enabled = true;
                    LastResult = result;
                    progress.Text = result;
                    progress.ForeColor = ok ? Pal.Good : Pal.Bad;
                    ShowState();
                });
            } catch { }
        });
        th.IsBackground = true;
        th.Start();
    }

    public bool Building { get { return building; } }
}

public static class MenuIcon {
    static Bitmap Draw(int size) {
        Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp)) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            float r = size * 0.22f, m = size * 0.04f, w = size - 2 * m;
            using (GraphicsPath p = new GraphicsPath()) {
                p.AddArc(m, m, r * 2, r * 2, 180, 90);
                p.AddArc(m + w - r * 2, m, r * 2, r * 2, 270, 90);
                p.AddArc(m + w - r * 2, m + w - r * 2, r * 2, r * 2, 0, 90);
                p.AddArc(m, m + w - r * 2, r * 2, r * 2, 90, 90);
                p.CloseFigure();
                using (SolidBrush b = new SolidBrush(Color.FromArgb(44, 98, 165))) g.FillPath(b, p);
            }
            Color[] bars = { Color.FromArgb(150, 200, 255), Color.FromArgb(140, 225, 165), Color.FromArgb(250, 195, 110) };
            float h = size * 0.13f, x = size * 0.22f, bw = size * 0.56f;
            for (int i = 0; i < 3; i++)
                using (SolidBrush b = new SolidBrush(bars[i])) g.FillRectangle(b, x, size * 0.24f + i * size * 0.2f, bw, h);
        }
        return bmp;
    }

    public static Icon Make() {
        using (MemoryStream ms = new MemoryStream(Bytes())) return new Icon(ms);
    }

    public static void Save(string file) { File.WriteAllBytes(file, Bytes()); }

    // файл .ico: значки 48, 32 и 16 точек, 32 бита (BMP внутри)
    public static byte[] Bytes() {
        int[] sizes = { 48, 32, 16 };
        List<byte[]> images = new List<byte[]>();
        foreach (int s in sizes) images.Add(Image(Draw(s)));
        MemoryStream ms = new MemoryStream();
        BinaryWriter w = new BinaryWriter(ms);
        w.Write((short)0);
        w.Write((short)1);
        w.Write((short)sizes.Length);
        int offset = 6 + 16 * sizes.Length;
        for (int i = 0; i < sizes.Length; i++) {
            w.Write((byte)sizes[i]);
            w.Write((byte)sizes[i]);
            w.Write((byte)0);
            w.Write((byte)0);
            w.Write((short)1);
            w.Write((short)32);
            w.Write(images[i].Length);
            w.Write(offset);
            offset += images[i].Length;
        }
        foreach (byte[] img in images) w.Write(img);
        w.Flush();
        return ms.ToArray();
    }

    static byte[] Image(Bitmap bmp) {
        int s = bmp.Width;
        MemoryStream ms = new MemoryStream();
        BinaryWriter w = new BinaryWriter(ms);
        w.Write(40);
        w.Write(s);
        w.Write(s * 2);           // высота - вместе с маской
        w.Write((short)1);
        w.Write((short)32);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        for (int y = s - 1; y >= 0; y--)
            for (int x = 0; x < s; x++) {
                Color c = bmp.GetPixel(x, y);
                w.Write(c.B);
                w.Write(c.G);
                w.Write(c.R);
                w.Write(c.A);
            }
        int row = ((s + 31) / 32) * 4;      // маска - всё видно (прозрачность - в альфе)
        w.Write(new byte[row * s]);
        bmp.Dispose();
        w.Flush();
        return ms.ToArray();
    }
}

// Общее оформление программ - как у «Оформления договора»: тёмная тема по умолчанию (глаза меньше устают), светлая - по выбору
// (переключатель в окне, меняется при следующем запуске). Цвета ставятся один раз при запуске (Look.Apply), до создания окон;
// Look.Dress(окно) перекрашивает окно и всё, что в нём (и что добавится потом): системное светлое - в цвета темы, кнопки -
// плоские с рамкой. Цвета, заданные программой нарочно (например, синие окна как в программе базы), не трогаются
public static class Look {
    public static bool Dark = true;
    public static System.Drawing.Color Back, Surface, Input, Header, Border, InputBorder, Text, Caption, Hint, Accent, AccentHover, Hover, Pressed,
        Link, Good, Bad, Warn, Info, Strong, RowGood, RowWarn, RowBad, RowOther, SelGood, SelBad, SelOther, RowPick;
    public static readonly System.Drawing.Font Base = new System.Drawing.Font("Segoe UI", 9f), Small = new System.Drawing.Font("Segoe UI", 8.5f),
        Bold = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold), Title = new System.Drawing.Font("Segoe UI", 15f, System.Drawing.FontStyle.Bold);

    static System.Drawing.Color C(int r, int g, int b) { return System.Drawing.Color.FromArgb(r, g, b); }

    static Look() { Apply(true); }

    public static void Apply(bool dark) {
        Dark = dark;
        if (dark) {
            Back = C(30, 33, 38); Surface = C(40, 44, 51); Input = C(48, 53, 61); Header = C(44, 48, 55);
            Border = C(66, 72, 82); InputBorder = C(98, 106, 118);
            Text = C(222, 226, 232); Caption = C(172, 180, 192); Hint = C(138, 146, 158);
            Accent = C(44, 98, 165); AccentHover = C(54, 114, 188); Hover = C(56, 63, 74); Pressed = C(66, 76, 92);
            Link = C(118, 172, 236); Good = C(112, 200, 142); Bad = C(240, 120, 120); Warn = C(232, 164, 72); Info = C(130, 182, 240);
            Strong = C(205, 211, 220);
            RowGood = C(38, 72, 50); RowWarn = C(84, 64, 34); RowBad = C(88, 42, 46); RowOther = C(52, 56, 62);
            SelGood = C(46, 120, 70); SelBad = C(150, 50, 56); SelOther = C(84, 90, 100); RowPick = C(38, 60, 92);
        } else {
            Back = C(242, 244, 247); Surface = System.Drawing.Color.White; Input = System.Drawing.Color.White; Header = C(232, 236, 241);
            Border = C(221, 226, 233); InputBorder = C(196, 204, 214);
            Text = C(31, 36, 43); Caption = C(58, 66, 78); Hint = C(128, 137, 150);
            Accent = C(35, 99, 178); AccentHover = C(28, 84, 153); Hover = C(236, 242, 250); Pressed = C(222, 233, 247);
            Link = C(35, 99, 178); Good = C(46, 140, 87); Bad = C(200, 55, 55); Warn = C(190, 105, 0); Info = C(0, 70, 140);
            Strong = System.Drawing.Color.Black;
            RowGood = C(214, 241, 219); RowWarn = C(255, 228, 190); RowBad = C(255, 218, 218); RowOther = C(236, 236, 236);
            SelGood = C(40, 140, 60); SelBad = C(190, 40, 40); SelOther = C(110, 110, 110); RowPick = C(226, 240, 255);
        }
    }

    public static string SwitchText { get { return Dark ? "Тема: тёмная - сделать светлой" : "Тема: светлая - сделать тёмной"; } }
    public static string SwitchedText { get { return Dark ? "будет светлая - после перезапуска" : "будет тёмная - после перезапуска"; } }

    public delegate void SaveTheme(bool dark);

    // ссылка «Тема: ...» - щелчок запоминает другую тему (save), она будет при следующем запуске
    public static System.Windows.Forms.LinkLabel Switch(SaveTheme save) {
        System.Windows.Forms.LinkLabel l = new System.Windows.Forms.LinkLabel();
        l.Text = SwitchText;
        l.Font = Small;
        l.AutoSize = true;
        l.LinkColor = l.ActiveLinkColor = Link;
        l.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
        bool next = !Dark;
        l.LinkClicked += delegate {
            save(next);
            l.Text = next == Dark ? SwitchText : SwitchedText;
            next = !next;
        };
        return l;
    }

    static bool Same(System.Drawing.Color a, System.Drawing.Color b) { return a.ToArgb() == b.ToArgb(); }

    // светлое системное (его и красим)
    static bool Plain(System.Drawing.Color c) {
        return c.A == 255 && (Same(c, System.Drawing.Color.White) || Same(c, System.Drawing.SystemColors.Window) || Same(c, System.Drawing.SystemColors.Control)
            || Same(c, System.Drawing.SystemColors.ControlLight) || Same(c, System.Drawing.SystemColors.ButtonFace) || Same(c, System.Drawing.SystemColors.AppWorkspace));
    }

    static bool PlainText(System.Drawing.Color c) {
        return Same(c, System.Drawing.Color.Black) || Same(c, System.Drawing.SystemColors.ControlText) || Same(c, System.Drawing.SystemColors.WindowText);
    }

    // главная кнопка окна (синяя): AcceptButton окна или Tag «главная»
    public static void MainButton(System.Windows.Forms.Button b) { b.Tag = "главная"; Button(b); }

    public static void Button(System.Windows.Forms.Button b) {
        System.Windows.Forms.Form f = b.FindForm();
        bool main = "главная".Equals(b.Tag) || f != null && f.AcceptButton == b;
        bool custom = !main && !Plain(b.BackColor) && b.BackColor.A == 255;     // цвет задан программой - оставить
        b.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
        b.UseVisualStyleBackColor = false;
        b.Cursor = System.Windows.Forms.Cursors.Hand;
        if (main) {
            b.BackColor = Accent;
            b.ForeColor = System.Drawing.Color.White;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = b.FlatAppearance.MouseDownBackColor = AccentHover;
        } else if (!custom) {
            b.BackColor = Surface;
            b.ForeColor = Text;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = InputBorder;
            b.FlatAppearance.MouseOverBackColor = Hover;
            b.FlatAppearance.MouseDownBackColor = Pressed;
        }
    }

    public static void Grid(System.Windows.Forms.DataGridView g) {
        g.BackgroundColor = Surface;
        g.GridColor = Border;
        g.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        g.DefaultCellStyle.BackColor = Input;
        g.DefaultCellStyle.ForeColor = Text;
        g.DefaultCellStyle.SelectionBackColor = Accent;
        g.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
        g.EnableHeadersVisualStyles = false;
        g.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
        g.ColumnHeadersDefaultCellStyle.BackColor = Header;
        g.ColumnHeadersDefaultCellStyle.ForeColor = Text;
        g.ColumnHeadersDefaultCellStyle.SelectionBackColor = Header;
        g.ColumnHeadersDefaultCellStyle.SelectionForeColor = Text;
        g.RowHeadersDefaultCellStyle.BackColor = Header;
        g.RowHeadersDefaultCellStyle.ForeColor = Text;
    }

    // окно и всё в нём (и что добавится потом)
    public static void Dress(System.Windows.Forms.Control c) {
        if (c == null || "свой цвет".Equals(c.Tag)) return;
        if (c is System.Windows.Forms.Form) {
            c.BackColor = Back;
            c.ForeColor = Text;
        } else if (c is System.Windows.Forms.TextBoxBase || c is System.Windows.Forms.ComboBox || c is System.Windows.Forms.ListBox
            || c is System.Windows.Forms.ListView || c is System.Windows.Forms.NumericUpDown || c is System.Windows.Forms.TreeView) {
            if (Plain(c.BackColor)) c.BackColor = c is System.Windows.Forms.TextBoxBase && ((System.Windows.Forms.TextBoxBase)c).ReadOnly ? Surface : Input;
            if (PlainText(c.ForeColor)) c.ForeColor = Text;
            System.Windows.Forms.TextBox t = c as System.Windows.Forms.TextBox;
            if (t != null && t.BorderStyle == System.Windows.Forms.BorderStyle.Fixed3D) t.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            System.Windows.Forms.ComboBox cb = c as System.Windows.Forms.ComboBox;
            if (cb != null && Dark) cb.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            System.Windows.Forms.ListView lv = c as System.Windows.Forms.ListView;
            if (lv != null && Dark) lv.GridLines = false;           // линии сетки Windows рисует светлыми - на тёмном режут глаз
        } else if (c is System.Windows.Forms.DataGridView) {
            Grid((System.Windows.Forms.DataGridView)c);
        } else if (c is System.Windows.Forms.Button) {
            Button((System.Windows.Forms.Button)c);
        } else if (c is System.Windows.Forms.LinkLabel) {
            System.Windows.Forms.LinkLabel l = (System.Windows.Forms.LinkLabel)c;
            l.LinkColor = Link;
            l.ActiveLinkColor = AccentHover;
            l.VisitedLinkColor = Link;
            if (PlainText(l.ForeColor)) l.ForeColor = Text;
        } else if (c is System.Windows.Forms.DateTimePicker) {
            System.Windows.Forms.DateTimePicker d = (System.Windows.Forms.DateTimePicker)c;
            d.CalendarMonthBackground = Input;
            d.CalendarForeColor = Text;
            d.CalendarTitleBackColor = Accent;
            d.CalendarTitleForeColor = System.Drawing.Color.White;
            d.CalendarTrailingForeColor = Hint;
        } else if (c is System.Windows.Forms.MonthCalendar) {
            System.Windows.Forms.MonthCalendar m = (System.Windows.Forms.MonthCalendar)c;
            // календарь Windows в стиле системы свои цвета не берёт - в тёмной теме рисуем его без стиля системы
            if (Dark) { if (m.IsHandleCreated) Plain(m); else m.HandleCreated += delegate { Plain(m); }; }
            m.BackColor = Input;
            m.ForeColor = Text;
            m.TitleBackColor = Accent;
            m.TitleForeColor = System.Drawing.Color.White;
            m.TrailingForeColor = Hint;
        } else if (c is System.Windows.Forms.ProgressBar) {
        } else {
            // подписи, галочки, панели, разделители
            if (!(c is System.Windows.Forms.Label) && !(c is System.Windows.Forms.CheckBox) && !(c is System.Windows.Forms.RadioButton) && Plain(c.BackColor)) c.BackColor = Back;
            if (PlainText(c.ForeColor)) c.ForeColor = Text;
            else if (Same(c.ForeColor, System.Drawing.SystemColors.GrayText)) c.ForeColor = Hint;
        }
        foreach (System.Windows.Forms.Control x in c.Controls) Dress(x);
        if (!hooked.ContainsKey(c) && (c is System.Windows.Forms.ContainerControl || c is System.Windows.Forms.Panel || c is System.Windows.Forms.GroupBox
            || c is System.Windows.Forms.SplitterPanel || c is System.Windows.Forms.TabPage)) {
            hooked[c] = true;
            System.Windows.Forms.Control me = c;
            c.ControlAdded += delegate(object s, System.Windows.Forms.ControlEventArgs e) { Dress(e.Control); };
            c.Disposed += delegate { hooked.Remove(me); };
        }
    }

    [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    static extern int SetWindowTheme(System.IntPtr hwnd, string app, string idList);

    static void Plain(System.Windows.Forms.Control c) {
        if (System.Environment.OSVersion.Platform != System.PlatformID.Win32NT) return;
        try { SetWindowTheme(c.Handle, "", ""); } catch { }
    }

    static readonly System.Collections.Generic.Dictionary<System.Windows.Forms.Control, bool> hooked = new System.Collections.Generic.Dictionary<System.Windows.Forms.Control, bool>();

    // окно - в теме при показе (и шрифт программы)
    public static void Wear(System.Windows.Forms.Form f) {
        f.Load += delegate { Dress(f); };
    }
}
'@
# Программа собирается из этого скрипта ОДИН раз после обновления и сохраняется готовой (файл menu_*.dll
# в папке профиля пользователя). Следующие запуски берут готовую - без сборки, за секунды.
# Изменился скрипт - у файла другое имя, программа соберётся заново.
$t0 = Get-Date
$stamp = '' + $code.Length
try {
    $sha = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider
    $stamp += '_' + ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($code)), 0, 6) -replace '-', '')
} catch {
    $stamp += '_' + ($code.GetHashCode() -band 0x7FFFFFFF)
}
$name = "menu_$stamp.dll"
$cacheDirs = @()
if ($env:LOCALAPPDATA) { $cacheDirs += (Join-Path $env:LOCALAPPDATA 'menu-programm') }
if ($env:PUBLIC) { $cacheDirs += (Join-Path $env:PUBLIC 'menu-programm') }
$cacheDirs += (Join-Path $here 'kesh-programmy')

function Test-Ready { return ($null -ne ('MenuApp' -as [type])) }

$launch = $null
foreach ($d in $cacheDirs) {
    $dll = Join-Path $d $name
    if ((-not (Test-Ready)) -and (Test-Path $dll)) {
        try {
            [void][Reflection.Assembly]::LoadFrom($dll)
            if (Test-Ready) { $launch = 'готовая программа ' + $dll + ', ' + ((Get-Date) - $t0).TotalSeconds.ToString('0.0') + ' с' }
        } catch { }
    }
}

# Предупреждения компилятора (не ошибки) не должны мешать запуску: -IgnoreWarnings, если PowerShell его знает
$quiet = @{}
try { if ((Get-Command Add-Type).Parameters.ContainsKey('IgnoreWarnings')) { $quiet['IgnoreWarnings'] = $true } } catch { }

# Окошко на время сборки: иначе несколько минут ничего не видно
$splash = $null
$splashText = $null
function Show-Splash {
    try {
        $script:splash = New-Object System.Windows.Forms.Form
        $splash.Text = 'Программы'
        $splash.FormBorderStyle = 'FixedDialog'
        $splash.ControlBox = $false
        $splash.StartPosition = 'CenterScreen'
        $splash.ClientSize = New-Object System.Drawing.Size(460, 110)
        $splash.TopMost = $true
        $script:splashText = New-Object System.Windows.Forms.Label
        $splashText.Dock = 'Fill'
        $splashText.TextAlign = 'MiddleCenter'
        $splashText.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
        $splash.Controls.Add($splashText)
        Update-Splash 0
        $splash.Show()
        [System.Windows.Forms.Application]::DoEvents()
    } catch { $script:splash = $null }
}
function Update-Splash($seconds) {
    if ($splashText -eq $null) { return }
    $splashText.Text = "Программа обновилась - Windows собирает её.`nЭто только один раз после обновления.`n`nПрошло: " + [int]$seconds + ' с'
}

# Сборка. Временные файлы компилятора - в своей папке: папку TEMP Windows компилятор на некоторых
# компьютерах открыть не может ("cannot open ...\Temp\RES....tmp for writing").
# Собираем в фоне, а окошко тем временем показывает, сколько прошло.
function Build-Program($dll, $tmp) {
    if (Test-Path $tmp) { Get-ChildItem -Path $tmp -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
    else { [void](New-Item -ItemType Directory -Path $tmp) }
    $oldTmp = $env:TMP
    $oldTemp = $env:TEMP
    $env:TMP = $tmp
    $env:TEMP = $tmp
    $ps = $null
    try {
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript({
            param($code, $refs, $dll, $quiet)
            $ErrorActionPreference = 'Stop'
            if ($dll) { Add-Type @quiet -ReferencedAssemblies $refs -TypeDefinition $code -OutputAssembly $dll -OutputType Library }
            else { Add-Type @quiet -ReferencedAssemblies $refs -TypeDefinition $code }
        }).AddArgument($code).AddArgument($refs).AddArgument($dll).AddArgument($quiet)
        $h = $ps.BeginInvoke()
        while (-not $h.IsCompleted) {
            if ($splash -ne $null) {
                Update-Splash ((Get-Date) - $t0).TotalSeconds
                [System.Windows.Forms.Application]::DoEvents()
            }
            Start-Sleep -Milliseconds 200
        }
        try { [void]$ps.EndInvoke($h) } catch {
            $ex = $_.Exception
            if ($ex.InnerException -ne $null) { $ex = $ex.InnerException }
            throw $ex
        }
        if ($ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
    } finally {
        if ($ps -ne $null) { $ps.Dispose() }
        $env:TMP = $oldTmp
        $env:TEMP = $oldTemp
        Get-ChildItem -Path $tmp -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($dll -and -not (Test-Ready)) { [void][Reflection.Assembly]::LoadFrom($dll) }
}

$errors = @()
if (-not (Test-Ready)) {
    Show-Splash
    foreach ($d in $cacheDirs) {
        try {
            if (-not (Test-Path $d)) { [void](New-Item -ItemType Directory -Path $d) }
            # прежние версии программы больше не нужны
            Get-ChildItem -Path $d -Filter 'menu_*.dll' -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne $name } | Remove-Item -ErrorAction SilentlyContinue
            $dll = Join-Path $d $name
            # файл есть, но не загрузился (прошлая сборка прервалась) - собираем заново
            if (Test-Path $dll) { Remove-Item $dll -Force -ErrorAction SilentlyContinue }
            Build-Program $dll (Join-Path $d 'tmp')
            if (Test-Ready) { $launch = 'собрана и сохранена (' + $dll + ') за ' + ((Get-Date) - $t0).TotalSeconds.ToString('0') + ' с'; break }
        } catch {
            $errors += "$d : $_"
            if ($dll -and (Test-Path $dll)) { Remove-Item $dll -Force -ErrorAction SilentlyContinue }
            # ошибка в самом коде - другая папка не поможет
            if ("$_" -match 'error CS\d+') { break }
        }
    }
    # сохранить не получилось нигде - собираем в памяти (временные файлы - рядом со скриптом)
    if ((-not (Test-Ready)) -and -not ("$errors" -match 'error CS\d+')) {
        try {
            Build-Program $null (Join-Path $here 'tmp')
            if (Test-Ready) { $launch = 'собрана в памяти (сохранить не удалось) за ' + ((Get-Date) - $t0).TotalSeconds.ToString('0') + ' с; ' + ($errors -join '; ') }
        } catch { $errors += "в памяти: $_" }
    }
    if ($splash -ne $null) { $splash.Close(); $splash.Dispose() }
}

# ошибка запуска - в oshibki.txt рядом с программой (что за Windows, PowerShell, .NET, памяти) - его можно прислать
function Write-Crash($what, $text) {
    try {
        $mem = ''
        try { $os = Get-WmiObject Win32_OperatingSystem; $mem = ', памяти свободно ' + [int]($os.FreePhysicalMemory / 1024) + ' из ' + [int]($os.TotalVisibleMemorySize / 1024) + ' МБ' } catch { }
        $info = 'Windows ' + [Environment]::OSVersion.Version + ', PowerShell ' + $PSVersionTable.PSVersion + ', .NET ' + [Environment]::Version + ', ' + ([IntPtr]::Size * 8) + '-битный' + $mem
        $text = [regex]::Replace("$text", '\d{6,}', '######')
        $line = '=== ' + (Get-Date).ToString('dd.MM.yyyy HH:mm:ss') + ' (' + $what + ") ===`r`n" + $info + "`r`n" + $text + "`r`n`r`n"
        [IO.File]::AppendAllText((Join-Path $here 'oshibki.txt'), $line, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}

if (-not (Test-Ready)) {
    Write-Crash 'сборка' ($errors -join "`r`n`r`n")
    [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n" + ($errors -join "`n`n") + "`n`nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", 'Программы - ошибка')
    exit
}

try {
    [MenuApp]::Run($here, $launch)
} catch {
    $ex = $_.Exception
    while ($ex.InnerException -ne $null) { $ex = $ex.InnerException }
    Write-Crash 'работа' $ex.ToString()
    [void][System.Windows.Forms.MessageBox]::Show("Ошибка: " + $ex.Message + "`n`nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", 'Программы - ошибка')
}
