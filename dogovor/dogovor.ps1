# Оформление договора электроснабжения: форма для заполнения -> договор в Word по шаблону shablon.docx.
# Запуск: zapusk.vbs (без чёрного окна) или zapusk.bat.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @([System.Windows.Forms.Form].Assembly.Location, [System.Drawing.Point].Assembly.Location, [System.Xml.XmlDocument].Assembly.Location)

$code = @'
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System;

// .docx - это ZIP-архив. Читаем и пишем его сами (Deflate есть в любом .NET), без Word и сторонних программ.
public class ZipEntry {
    public string Name;
    public int Method;          // 0 - без сжатия, 8 - deflate
    public uint Crc;
    public int Size;
    public byte[] Raw;          // данные как в архиве (сжатые)
    public ushort Time, Date;

    public byte[] GetData() {
        if (Method == 0) return Raw;
        using (MemoryStream src = new MemoryStream(Raw))
        using (DeflateStream z = new DeflateStream(src, CompressionMode.Decompress))
        using (MemoryStream dst = new MemoryStream()) {
            byte[] buf = new byte[65536];
            int n;
            while ((n = z.Read(buf, 0, buf.Length)) > 0) dst.Write(buf, 0, n);
            return dst.ToArray();
        }
    }

    public void SetData(byte[] data) {
        Size = data.Length;
        Crc = Crc32.Compute(data);
        using (MemoryStream dst = new MemoryStream()) {
            using (DeflateStream z = new DeflateStream(dst, CompressionMode.Compress, true)) z.Write(data, 0, data.Length);
            Raw = dst.ToArray();
        }
        Method = 8;
    }
}

public static class Crc32 {
    static uint[] table;

    public static uint Compute(byte[] data) {
        if (table == null) {
            uint[] t = new uint[256];
            for (uint i = 0; i < 256; i++) {
                uint c = i;
                for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                t[i] = c;
            }
            table = t;
        }
        uint crc = 0xFFFFFFFFu;
        foreach (byte b in data) crc = table[(crc ^ b) & 0xFF] ^ (crc >> 8);
        return crc ^ 0xFFFFFFFFu;
    }
}

public static class Zip {
    public static List<ZipEntry> Read(string path) {
        byte[] d = File.ReadAllBytes(path);
        int eocd = -1;
        for (int i = d.Length - 22; i >= Math.Max(0, d.Length - 65557); i--)
            if (U32(d, i) == 0x06054b50) { eocd = i; break; }
        if (eocd < 0) throw new InvalidDataException("Файл не похож на документ Word (.docx).");
        int count = U16(d, eocd + 10);
        int pos = (int)U32(d, eocd + 16);
        List<ZipEntry> list = new List<ZipEntry>();
        for (int k = 0; k < count; k++) {
            if (U32(d, pos) != 0x02014b50) throw new InvalidDataException("Повреждён архив .docx.");
            ZipEntry e = new ZipEntry();
            e.Method = U16(d, pos + 10);
            e.Time = (ushort)U16(d, pos + 12);
            e.Date = (ushort)U16(d, pos + 14);
            e.Crc = U32(d, pos + 16);
            int csize = (int)U32(d, pos + 20);
            e.Size = (int)U32(d, pos + 24);
            int nl = U16(d, pos + 28), el = U16(d, pos + 30), cl = U16(d, pos + 32);
            int local = (int)U32(d, pos + 42);
            e.Name = Encoding.UTF8.GetString(d, pos + 46, nl);
            int data = local + 30 + U16(d, local + 26) + U16(d, local + 28);
            e.Raw = new byte[csize];
            Buffer.BlockCopy(d, data, e.Raw, 0, csize);
            list.Add(e);
            pos += 46 + nl + el + cl;
        }
        return list;
    }

    public static void Write(string path, List<ZipEntry> entries) {
        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (BinaryWriter w = new BinaryWriter(fs)) {
            List<int> offsets = new List<int>();
            foreach (ZipEntry e in entries) {
                byte[] name = Encoding.UTF8.GetBytes(e.Name);
                offsets.Add((int)fs.Position);
                w.Write(0x04034b50u); w.Write((ushort)20); w.Write((ushort)0x0800); w.Write((ushort)e.Method);
                w.Write(e.Time); w.Write(e.Date); w.Write(e.Crc); w.Write(e.Raw.Length); w.Write(e.Size);
                w.Write((ushort)name.Length); w.Write((ushort)0); w.Write(name); w.Write(e.Raw);
            }
            int cd = (int)fs.Position;
            for (int i = 0; i < entries.Count; i++) {
                ZipEntry e = entries[i];
                byte[] name = Encoding.UTF8.GetBytes(e.Name);
                w.Write(0x02014b50u); w.Write((ushort)20); w.Write((ushort)20); w.Write((ushort)0x0800); w.Write((ushort)e.Method);
                w.Write(e.Time); w.Write(e.Date); w.Write(e.Crc); w.Write(e.Raw.Length); w.Write(e.Size);
                w.Write((ushort)name.Length); w.Write((ushort)0); w.Write((ushort)0); w.Write((ushort)0); w.Write((ushort)0);
                w.Write(0u); w.Write(offsets[i]); w.Write(name);
            }
            int cdSize = (int)fs.Position - cd;
            w.Write(0x06054b50u); w.Write((ushort)0); w.Write((ushort)0);
            w.Write((ushort)entries.Count); w.Write((ushort)entries.Count); w.Write(cdSize); w.Write(cd); w.Write((ushort)0);
        }
    }

    static int U16(byte[] d, int p) { return d[p] | (d[p + 1] << 8); }
    static uint U32(byte[] d, int p) { return (uint)(d[p] | (d[p + 1] << 8) | (d[p + 2] << 16) | (d[p + 3] << 24)); }
}

// Шаблон договора: поля «ИМЯ» заменяются текстом, «□ИМЯ» - флажком ☑ или ☐ (символы Wingdings, как в бланке),
// «¶ИМЯ» - текстом, а если он пустой, абзац с полем убирается.
// Поле может быть разбито Word-ом на несколько кусков текста - собираем его по всему абзацу.
public static class DocxTemplate {
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly Regex Field = new Regex("«([□¶]?)([A-ZА-ЯЁ0-9_]+)»");

    // values: имя поля -> текст ("\n" - перенос строки); для флажков "1" - отмечен. Возвращает поля, которых не было в values
    public static List<string> Fill(string template, string output, Dictionary<string, string> values) {
        List<ZipEntry> entries = Zip.Read(template);
        List<string> missing = new List<string>();
        foreach (ZipEntry e in entries) {
            if (e.Name != "word/document.xml" && !e.Name.StartsWith("word/header") && !e.Name.StartsWith("word/footer")) continue;
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            using (MemoryStream ms = new MemoryStream(e.GetData())) doc.Load(ms);
            XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", W);
            bool changed = false;
            List<XmlNode> paras = new List<XmlNode>();
            foreach (XmlNode p in doc.SelectNodes("//w:p", ns)) paras.Add(p);
            foreach (XmlNode p in paras)
                if (FillParagraph(doc, p, values, missing)) changed = true;
            if (!changed) continue;
            using (MemoryStream ms = new MemoryStream()) {
                XmlWriterSettings ws = new XmlWriterSettings();
                ws.Encoding = new UTF8Encoding(false);
                using (XmlWriter xw = XmlWriter.Create(ms, ws)) doc.Save(xw);
                e.SetData(ms.ToArray());
            }
        }
        Zip.Write(output, entries);
        return missing;
    }

    // ближайший абзац узла (текст в надписях - свой абзац внутри абзаца)
    static XmlNode OwnParagraph(XmlNode n) {
        for (XmlNode x = n.ParentNode; x != null; x = x.ParentNode)
            if (x.LocalName == "p" && x.NamespaceURI == W) return x;
        return null;
    }

    static List<XmlNode> Texts(XmlNode p) {
        List<XmlNode> ts = new List<XmlNode>();
        foreach (XmlNode t in ((XmlElement)p).GetElementsByTagName("t", W))
            if (OwnParagraph(t) == p && t.ParentNode.LocalName == "r") ts.Add(t);
        return ts;
    }

    static bool FillParagraph(XmlDocument doc, XmlNode p, Dictionary<string, string> values, List<string> missing) {
        List<XmlNode> ts = Texts(p);
        if (ts.Count == 0) return false;
        StringBuilder sb = new StringBuilder();
        foreach (XmlNode t in ts) sb.Append(t.InnerText);
        MatchCollection ms = Field.Matches(sb.ToString());
        if (ms.Count == 0) return false;
        bool hasPara = false;
        // с конца: текст перед более ранними полями не меняется, их позиции те же
        for (int m = ms.Count - 1; m >= 0; m--) {
            Match match = ms[m];
            if (m < ms.Count - 1) ts = Texts(p);
            int start = match.Index, end = match.Index + match.Length, off = 0;
            XmlNode first = null;
            int firstOffset = 0;
            foreach (XmlNode t in ts) {
                string s = t.InnerText;
                int a = off, b = off + s.Length;
                off = b;
                if (b <= start || a >= end) continue;
                int ls = Math.Max(start, a) - a, le = Math.Min(end, b) - a;
                if (first == null) {
                    first = t;
                    firstOffset = ls;
                    SetText(t, s.Substring(0, ls) + match.Value + s.Substring(le));
                } else SetText(t, s.Substring(0, ls) + s.Substring(le));
            }
            string name = match.Groups[2].Value;
            bool box = match.Groups[1].Value == "□";
            if (match.Groups[1].Value == "¶") hasPara = true;
            string value;
            if (!values.TryGetValue((box ? "□" : "") + name, out value)) { value = ""; if (!missing.Contains(name)) missing.Add(name); }
            Put(doc, first, firstOffset, match.Length, value, box);
        }
        // «¶ИМЯ» - поле в своём абзаце: пустое значение - абзаца нет совсем
        if (hasPara && p.ParentNode != null) {
            StringBuilder rest = new StringBuilder();
            foreach (XmlNode t in Texts(p)) rest.Append(t.InnerText);
            if (rest.ToString().Trim().Length == 0 && ((XmlElement)p).GetElementsByTagName("br", W).Count == 0) p.ParentNode.RemoveChild(p);
        }
        return true;
    }

    static void SetText(XmlNode t, string text) {
        t.InnerText = text;
        ((XmlElement)t).SetAttribute("space", "http://www.w3.org/XML/1998/namespace", "preserve");
    }

    // вместо куска текста t[offset, offset+length) - значение: текст (с переносами строк) или флажок
    static void Put(XmlDocument doc, XmlNode t, int offset, int length, string value, bool box) {
        XmlNode run = t.ParentNode;
        string text = t.InnerText;
        string before = text.Substring(0, offset), after = text.Substring(offset + length);
        if (box) {
            // флажок - отдельным куском: иначе шрифт Wingdings может достаться и соседнему тексту
            XmlNode rPr = null;
            foreach (XmlNode c in run.ChildNodes) if (c.LocalName == "rPr") rPr = c;
            XmlNode head = NewRun(doc, rPr), mark = NewRun(doc, rPr), tail = NewRun(doc, rPr);
            bool past = false;
            foreach (XmlNode c in new List<XmlNode>(Children(run))) {
                if (c == rPr) continue;
                if (c == t) { past = true; continue; }
                (past ? tail : head).AppendChild(c);
            }
            if (before.Length > 0) head.AppendChild(MakeText(doc, before));
            if (after.Length > 0) tail.InsertAfter(MakeText(doc, after), rPr == null ? null : tail.FirstChild);
            XmlElement sym = doc.CreateElement("w", "sym", W);
            sym.SetAttribute("font", W, "Wingdings");
            sym.SetAttribute("char", W, value == "1" ? "F0FE" : "F0A8");
            mark.AppendChild(sym);
            XmlNode parent = run.ParentNode;
            foreach (XmlNode r in new XmlNode[] { head, mark, tail })
                if (r.ChildNodes.Count > (rPr == null ? 0 : 1)) parent.InsertBefore(r, run);
            parent.RemoveChild(run);
            return;
        }
        List<XmlNode> parts = new List<XmlNode>();
        if (before.Length > 0) parts.Add(MakeText(doc, before));
        string[] lines = value.Replace("\r\n", "\n").Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) parts.Add(doc.CreateElement("w", "br", W));
            if (lines[i].Length > 0) parts.Add(MakeText(doc, lines[i]));
        }
        if (after.Length > 0) parts.Add(MakeText(doc, after));
        foreach (XmlNode n in parts) run.InsertBefore(n, t);
        run.RemoveChild(t);
    }

    static List<XmlNode> Children(XmlNode n) {
        List<XmlNode> list = new List<XmlNode>();
        foreach (XmlNode c in n.ChildNodes) list.Add(c);
        return list;
    }

    static XmlNode NewRun(XmlDocument doc, XmlNode rPr) {
        XmlElement r = doc.CreateElement("w", "r", W);
        if (rPr != null) r.AppendChild(rPr.CloneNode(true));
        return r;
    }

    static XmlNode MakeText(XmlDocument doc, string s) {
        XmlElement t = doc.CreateElement("w", "t", W);
        t.SetAttribute("space", "http://www.w3.org/XML/1998/namespace", "preserve");
        t.InnerText = s;
        return t;
    }

    // все поля шаблона - для проверки
    public static List<string> Fields(string template) {
        List<string> names = new List<string>();
        foreach (ZipEntry e in Zip.Read(template)) {
            if (e.Name != "word/document.xml") continue;
            XmlDocument doc = new XmlDocument();
            using (MemoryStream ms = new MemoryStream(e.GetData())) doc.Load(ms);
            StringBuilder sb = new StringBuilder();
            foreach (XmlNode p in doc.GetElementsByTagName("p", W)) {
                sb.Length = 0;
                foreach (XmlNode t in ((XmlElement)p).GetElementsByTagName("t", W)) if (OwnParagraph(t) == p) sb.Append(t.InnerText);
                foreach (Match m in Field.Matches(sb.ToString())) {
                    string n = (m.Groups[1].Value == "□" ? "□" : "") + m.Groups[2].Value;
                    if (!names.Contains(n)) names.Add(n);
                }
            }
        }
        return names;
    }
}

// Справочники (spravochniki.txt), запомненные значения (nastroiki.txt) и ранее введённые варианты (istoriya.txt).
// Все файлы - простой текст рядом с программой, их можно править Блокнотом.
public class Lists {
    public readonly string File;
    public DateTime Stamp;
    Dictionary<string, List<string>> lists = new Dictionary<string, List<string>>();
    List<string> lines = new List<string>();

    public Lists(string file) {
        File = file;
        if (!System.IO.File.Exists(file)) System.IO.File.WriteAllText(file, Default, new UTF8Encoding(true));
        Load();
        try { AddMissing(); } catch { }
    }

    // в справочник от прошлой версии дописываются списки, которых в нём ещё нет (правки пользователя не трогаются)
    void AddMissing() {
        StringBuilder add = new StringBuilder();
        bool take = false;
        foreach (string raw in Default.Replace("\r\n", "\n").Split('\n')) {
            string s = raw.Trim();
            if (s.StartsWith("[") && s.EndsWith("]")) take = !lists.ContainsKey(Key(s.Substring(1, s.Length - 2)));
            if (take) add.Append(raw).Append("\r\n");
        }
        if (add.Length == 0) return;
        System.IO.File.AppendAllText(File, "\r\n" + add.ToString().TrimEnd() + "\r\n", new UTF8Encoding(false));
        Load();
    }

    // счётчики из справочника: "CE208BY S53 145.1 JA.UKVFLZ" -> модель CE208BY, корпус S53, модификация 145.1 JA.UKVFLZ
    public List<MeterType> Meters() {
        List<MeterType> r = new List<MeterType>();
        foreach (string s in Get("Счётчики")) r.Add(MeterType.Parse(s));
        return r;
    }

    public void Load() {
        lists.Clear();
        lines = new List<string>(System.IO.File.ReadAllLines(File, Encoding.UTF8));
        Stamp = System.IO.File.GetLastWriteTime(File);
        List<string> cur = null;
        foreach (string raw in lines) {
            string s = raw.Trim();
            if (s.Length == 0 || s.StartsWith("#")) continue;
            if (s.StartsWith("[") && s.EndsWith("]")) {
                string name = Key(s.Substring(1, s.Length - 2));
                if (!lists.TryGetValue(name, out cur)) { cur = new List<string>(); lists[name] = cur; }
                continue;
            }
            if (cur != null && !cur.Contains(s)) cur.Add(s);
        }
    }

    public bool Changed() {
        try { return System.IO.File.GetLastWriteTime(File) != Stamp; } catch { return false; }
    }

    static string Key(string name) { return name.Trim().ToLowerInvariant().Replace('ё', 'е'); }

    public List<string> Get(string name) {
        List<string> l;
        return lists.TryGetValue(Key(name), out l) ? new List<string>(l) : new List<string>();
    }

    // подписанты: "Как в списке | должность | ФИО (кого) | доверенность | подпись"
    public List<string[]> Signers() {
        List<string[]> r = new List<string[]>();
        foreach (string s in Get("Подписанты")) {
            string[] p = s.Split('|');
            string[] v = new string[5];
            for (int i = 0; i < 5; i++) v[i] = i < p.Length ? p[i].Trim() : "";
            if (v[0].Length > 0) r.Add(v);
        }
        return r;
    }

    // запомнить подписанта: строка с тем же кратким именем заменяется, иначе добавляется в конец списка
    public void SaveSigner(string[] v) {
        string line = string.Join(" | ", v);
        int section = -1, last = -1, same = -1;
        for (int i = 0; i < lines.Count; i++) {
            string s = lines[i].Trim();
            if (s.StartsWith("[") && s.EndsWith("]")) {
                if (section >= 0 && last >= 0) break;
                if (Key(s.Substring(1, s.Length - 2)) == Key("Подписанты")) { section = i; last = i; }
                continue;
            }
            if (section < 0) continue;
            if (s.Length > 0) last = i;
            if (s.Length > 0 && !s.StartsWith("#") && Key(s.Split('|')[0]) == Key(v[0])) same = i;
        }
        if (same >= 0) lines[same] = line;
        else if (section >= 0) lines.Insert(last + 1, line);
        else { lines.Add(""); lines.Add("[Подписанты]"); lines.Add(line); }
        System.IO.File.WriteAllLines(File, lines.ToArray(), new UTF8Encoding(true));
        Load();
    }

    public const string Default =
"# Справочники программы «Оформление договора».\r\n" +
"# Можно править Блокнотом: каждый список начинается строкой [Название], ниже - по одному варианту в строке.\r\n" +
"# Строки, начинающиеся с #, - пояснения, программа их не читает. После сохранения программа подхватит изменения сама.\r\n" +
"\r\n" +
"[Подписанты]\r\n" +
"# Как в списке | должность (в лице кого) | фамилия имя отчество (кого) | доверенность | подпись\r\n" +
"Зенченко Д.И. | начальника группы по работе с потребителями Острошицкого производственно-сбытового участка отделения по сбыту электрической энергии Минского района филиала Энергосбыт РУП «МИНСКЭНЕРГО» | Зенченко Дмитрия Ивановича | № 3/1832 от 11.09.2026 | Д.И.Зенченко\r\n" +
"\r\n" +
"[Оформил]\r\n" +
"Гаврош В.А.\r\nКолотило А.С.\r\nГалай А.В.\r\nНюнько В.С.\r\nВолкович В.Н.\r\nХилевич Р.А.\r\nМихалевич Д.А.\r\nШевчик Я.Е.\r\nПоляков И.А.\r\nЦеховой А.Н.\r\nЗеленевская Е.В.\r\nГоряев С.А.\r\nУстинович И.И.\r\nСинегуб Д.С.\r\n" +
"\r\n" +
"[Объект электроснабжения]\r\n" +
"Жилой дом\r\nСтройплощадка жилого дома\r\nЛичное подсобное хозяйство\r\nКвартира\r\nГараж\r\nЖилой дом (1-ая очередь)\r\nКвартира в блокированном жилом доме\r\n" +
"\r\n" +
"[Тип населённого пункта]\r\n" +
"д.\r\nаг.\r\nг.\r\nг.п.\r\nп.\r\nСТ\r\n" +
"\r\n" +
"[Тип улицы]\r\n" +
"ул.\r\nпер.\r\nпр-т\r\nпр.\r\nтуп.\r\nмкр.\r\n" +
"\r\n" +
"[Сельсовет]\r\n" +
"Боровлянский\r\nКолодищанский\r\nОстрошицкий\r\nПапернянский\r\nШершунский\r\nЮзуфовский\r\n" +
"\r\n" +
"[Кем выдан допуск]\r\n" +
"филиал по Минску и Минской области ГУ «Госэнергогазнадзор»\r\n" +
"\r\n" +
"[Прочие условия]\r\n" +
"# пункт 28.1 договора; «отсутствуют» пишется без кавычек\r\n" +
"Потребитель обязуется использовать электроприборы бытового назначения только заводского изготовления, не превышать разрешенную к использованию мощность, указанную в Приложении, не применять электроприемники, влияющие на качество электроэнергии у других потребителей, не подключать без разрешения Исполнителя к своим электрическим сетям других потребителей\r\n" +
"После завершения строительства Потребитель обязуется получить документы на электроснабжение жилого дома (иного капитального строения) и не позднее 10 дней с даты получения документов обратиться к Исполнителю для перезаключения договора\r\n" +
"В случае осуществления предпринимательской деятельности или сдачи нежилого капитального строения в аренду юридическому лицу или индивидуальному предпринимателю Потребитель обязан в 10-дневный срок обратиться к Исполнителю с документами для перезаключения договора\r\n" +
"отсутствуют\r\n" +
"\r\n" +
"[Тарифная группа]\r\n" +
"(23) стройплощадка\r\n" +
"(23) Хозпостройка\r\n" +
"(23) \"Жилой дом, НЕ ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\"\r\n" +
"(26) \"Жилой дом, ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\"\r\n" +
"(30) \"Эл.эн. для нужд отопления и горячего водоснабжения в ж\\д, не оборудованных сис-ми централизованного тепло- и газоснабжения, при отсутствии отдельного прибора учета расхода эл.эн. для нужд отопления и горячего водоснабжения\"\r\n" +
"(31) \"Эл.эн. для нужд отопления и горячего водоснабжения в ж\\д, не оборудованных сис-ми централизованного тепло- и газоснабжения, при наличии отдельного прибора учета расхода эл.эн. для нужд отопления и горячего водоснабжения\"\r\n" +
"(23) отсутствует регистрация \"Жилой дом, НЕ ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\"\r\n" +
"(26) отсутствует регистрация \"Жилой дом, ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\"\r\n" +
"(30) отсутствует регистрация \"Эл.эн. для нужд отопления и горячего водоснабжения в ж\\д, не оборудованных сис-ми централизованного тепло- и газоснабжения, при отсутствии отдельного прибора учета расхода эл.эн. для нужд отопления и горячего водоснабжения\"\r\n" +
"(31) отсутствует регистрация \"Эл.эн. для нужд отопления и горячего водоснабжения в ж\\д, не оборудованных сис-ми централизованного тепло- и газоснабжения, при наличии отдельного прибора учета расхода эл.эн. для нужд отопления и горячего водоснабжения\"\r\n" +
"\r\n" +
"[Цель использования]\r\n" +
"Бытовое потребление\r\nОтопление и горячее водоснабжение\r\nБытовое потребление, отопление и горячее водоснабжение\r\n" +
"\r\n" +
"[Счётчики]\r\n" +
"# Полное обозначение, как на счётчике. В программе выбирается по частям: модель, корпус, модификация.\r\n" +
"# Однофазные СЕ 208\r\n" +
"CE208BY S53 145.1 JA.UKVFLZ\r\nCE208BY S53 145.1 JR.UKVFLZ\r\nCE208BY S53 145.1 JP.UKVFLZ\r\nCE208BY S53 145.1 JPR.UKVFLZ\r\n" +
"CE208BY S53 145.1 JR.QUKVFLZ\r\nCE208BY S53 145.1 JPR.QUKVFLZ\r\nCE208BY S53 146.1 JR.UKVFLZ\r\nCE208BY S53 146.1 JPR.UKVFLZ\r\n" +
"CE208BY S53 146.1 JR.QUKVFLZ\r\nCE208BY S53 146.1 JPR.QUKVFLZ\r\nCE208BY S7 145.1 JPR.UKVFLZ\r\nCE208BY S7 146.1 JPR.UKVFLZ\r\n" +
"CE208BY S7 149.1 JPR.QUKVFLZ\r\nCE208BY C1.145.2.JR.QUKVFZ\r\nCE208BY C1.146.2.JR.QUKVFZ\r\nCE208BY C1.145.2.JPR.QUKVFZ\r\n" +
"CE208BY C1.146.2.JPR.QUKVFZ\r\nCE208BY C2 145.2.JPR.QUKVFZ\r\nCE208BY C2 146.2.JPR.QUKVFZ\r\n" +
"# Трёхфазные СЕ 318\r\n" +
"CE318BY R32 043 JA.UVFL\r\nCE318BY R32 043 JR.UVFL\r\nCE318BY R32 146 JA.UVFL\r\nCE318BY R32 146 JR.UVFL\r\n" +
"CE318BY R32 146 JP.UVFL\r\nCE318BY R32 146 JPR.UVFL\r\nCE318BY S31 043 JPR.UVFL\r\nCE318BY S31 146 JA.UVFL\r\n" +
"CE318BY S31 146 JR.UVFL\r\nCE318BY S31 146 JPR.UVFL\r\nCE318BY S31 146 JR.QUVFL\r\nCE318BY S31 146 JPR.QUVFL\r\n" +
"CE318BY S34 146 JPR.UVFL\r\nCE318BY S34 146 JPR.QUVFL\r\nCE318BY S35 043 JPR.UVFL\r\nCE318BY S35 146 JPR.UVFL\r\n" +
"CE318BY S35 149.JPR.QUVFL\r\nCE318BY S37 146 JPR.QUVFL\r\n" +
"\r\n" +
"[Место установки]\r\n" +
"ЩУЭ на границе участка\r\nза границей земельного участка\r\nв доме\r\nв гараже\r\nна лестничной клетке\r\n" +
"\r\n" +
"[Принадлежность]\r\n" +
"Потребитель\r\nРУП «Минскэнерго»\r\n" +
"\r\n" +
"[Тип модема]\r\n" +
"PLC\r\nрадио\r\nPLC и радио\r\nотсутствует\r\n" +
"\r\n" +
"[Пломба госповерителя]\r\n" +
"леска\r\nпроволока\r\n" +
"\r\n" +
"[Электроприборы]\r\n" +
"# электронагревательные приборы: в договор - «Конвекторы - 2 шт. 1,5 кВт; Котёл - 1 шт. 6 кВт»\r\n" +
"Конвекторы\r\nКотёл\r\nБойлер\r\nЭл. плита\r\nТёплый пол\r\nОбогреватель\r\nТепловой насос\r\n" +
"\r\n" +
"[Вводной аппарат]\r\n" +
"отсутствует\r\nВА\r\nВН\r\n" +
"\r\n" +
"[Аппарат защиты]\r\n" +
"ВА\r\nВН\r\nдиф. автомат\r\nУЗО\r\n" +
"\r\n" +
"[Ток аппарата, А]\r\n" +
"# номиналы автоматов\r\n" +
"6\r\n10\r\n16\r\n20\r\n25\r\n32\r\n40\r\n50\r\n63\r\n80\r\n100\r\n120\r\n";
}

// Запомненные значения (последний подписант, кто оформил и т.п.) и ранее введённые варианты для подсказок
public class Memory {
    readonly string file;
    public readonly Dictionary<string, string> Values = new Dictionary<string, string>();
    public readonly Dictionary<string, List<string>> History = new Dictionary<string, List<string>>();
    const int Keep = 60;

    public Memory(string file) {
        this.file = file;
        try {
            if (!System.IO.File.Exists(file)) return;
            foreach (string line in System.IO.File.ReadAllLines(file, Encoding.UTF8)) {
                int eq = line.IndexOf('=');
                if (eq <= 0 || line.StartsWith("#")) continue;
                string k = line.Substring(0, eq).Trim(), v = Unescape(line.Substring(eq + 1));
                if (k.StartsWith("история.")) {
                    string h = k.Substring(8);
                    if (!History.ContainsKey(h)) History[h] = new List<string>();
                    if (v.Length > 0) History[h].Add(v);
                } else Values[k] = v;
            }
        } catch { }
    }

    public string Get(string key, string def) {
        string v;
        return Values.TryGetValue(key, out v) ? v : def;
    }

    public List<string> Hist(string key) {
        List<string> l;
        return History.TryGetValue(key, out l) ? new List<string>(l) : new List<string>();
    }

    // новое значение - в начало списка
    public void Remember(string key, string value) {
        value = (value ?? "").Trim();
        if (value.Length == 0 || value.IndexOf('\n') >= 0) return;
        List<string> l;
        if (!History.TryGetValue(key, out l)) { l = new List<string>(); History[key] = l; }
        int i = l.FindIndex(delegate(string s) { return string.Compare(s, value, StringComparison.OrdinalIgnoreCase) == 0; });
        if (i >= 0) l.RemoveAt(i);
        l.Insert(0, value);
        if (l.Count > Keep) l.RemoveRange(Keep, l.Count - Keep);
    }

    public void Save() {
        try {
            StringBuilder sb = new StringBuilder();
            sb.Append("# Запомненные значения программы «Оформление договора». Файл можно удалить - заполнится заново.\r\n");
            foreach (KeyValuePair<string, string> kv in Values) sb.Append(kv.Key).Append('=').Append(Escape(kv.Value)).Append("\r\n");
            foreach (KeyValuePair<string, List<string>> kv in History)
                foreach (string v in kv.Value) sb.Append("история.").Append(kv.Key).Append('=').Append(Escape(v)).Append("\r\n");
            System.IO.File.WriteAllText(file, sb.ToString(), new UTF8Encoding(true));
        } catch { }
    }

    public static string Escape(string s) {
        return (s ?? "").Replace("\\", "\\\\").Replace("\r\n", "\n").Replace("\n", "\\n");
    }

    public static string Unescape(string s) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            if (s[i] == '\\' && i + 1 < s.Length) {
                i++;
                sb.Append(s[i] == 'n' ? "\r\n" : s[i].ToString());
            } else sb.Append(s[i]);
        }
        return sb.ToString();
    }
}

// Тип счётчика по частям: модель, корпус, модификация (как в справочнике [Счётчики])
public class MeterType {
    public string Model = "", Case = "", Mod = "", Full = "";

    public static MeterType Parse(string full) {
        MeterType m = new MeterType();
        m.Full = full.Trim();
        int sp = m.Full.IndexOf(' ');
        if (sp < 0) { m.Model = m.Full; return m; }
        m.Model = m.Full.Substring(0, sp);
        string rest = m.Full.Substring(sp + 1).Trim();
        int k = rest.IndexOfAny(new char[] { ' ', '.' });
        m.Case = k < 0 ? rest : rest.Substring(0, k);
        m.Mod = k < 0 ? "" : rest.Substring(k + 1).Trim();
        return m;
    }

    public static bool Same(string a, string b) { return string.Compare(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase) == 0; }

    // СЕ 1xx, 2xx - однофазные, СЕ 3xx - трёхфазные
    public static string Phase(string model) {
        string m = model.Trim().ToUpper().Replace('С', 'C').Replace('Е', 'E');
        if (!m.StartsWith("CE") || m.Length < 3) return "";
        if (m[2] == '1' || m[2] == '2') return "1 ф";
        if (m[2] == '3') return "3 ф";
        return "";
    }

    // интерфейсы в модификации: P - PLC-модем, R - радиомодем ("JPR.UVFL" - оба)
    public static string Modem(string mod) {
        int j = mod.IndexOf('J');
        if (j < 0) return "";
        int e = mod.IndexOf('.', j);
        string code = e < 0 ? mod.Substring(j + 1) : mod.Substring(j + 1, e - j - 1);
        bool plc = code.IndexOf('P') >= 0, radio = code.IndexOf('R') >= 0;
        return plc && radio ? "PLC и радио" : plc ? "PLC" : radio ? "радио" : "";
    }
}

// Оформление: цвета, шрифты и элементы формы
public static class St {
    public static readonly Color Back = Color.FromArgb(242, 244, 247);
    public static readonly Color Card = Color.White;
    public static readonly Color Border = Color.FromArgb(221, 226, 233);
    public static readonly Color InputBorder = Color.FromArgb(196, 204, 214);
    public static readonly Color Nav = Color.FromArgb(28, 51, 84);
    public static readonly Color NavHover = Color.FromArgb(39, 67, 106);
    public static readonly Color NavActive = Color.FromArgb(48, 83, 130);
    public static readonly Color NavText = Color.FromArgb(196, 211, 230);
    public static readonly Color NavMark = Color.FromArgb(99, 179, 237);
    public static readonly Color Accent = Color.FromArgb(35, 99, 178);
    public static readonly Color AccentHover = Color.FromArgb(28, 84, 153);
    public static readonly Color Caption = Color.FromArgb(92, 102, 117);
    public static readonly Color Text = Color.FromArgb(31, 36, 43);
    public static readonly Color Hint = Color.FromArgb(128, 137, 150);
    public static readonly Color Missing = Color.FromArgb(253, 232, 232);
    public static readonly Color MissingBorder = Color.FromArgb(220, 90, 90);
    public static readonly Color Bad = Color.FromArgb(200, 55, 55);
    public static readonly Color Good = Color.FromArgb(46, 140, 87);
    public static readonly Color Box = Color.FromArgb(248, 250, 252);
    public static readonly Color Disabled = Color.FromArgb(236, 239, 243);

    public static readonly Font Base = new Font("Segoe UI", 10f);
    public static readonly Font Small = new Font("Segoe UI", 9f);
    public static readonly Font Bold = new Font("Segoe UI", 10f, FontStyle.Bold);
    public static readonly Font Title = new Font("Segoe UI", 12.5f, FontStyle.Bold);
    public static readonly Font Big = new Font("Segoe UI", 16f, FontStyle.Bold);
    public static readonly Font NavFont = new Font("Segoe UI", 10.5f);
}

public static class Ui {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);

    static bool Windows { get { return Environment.OSVersion.Platform == PlatformID.Win32NT; } }

    // серая подсказка в пустом поле (Windows Vista и новее)
    public static void Cue(Control c, string text) {
        if (!Windows) return;
        try {
            if (c is TextBox && !((TextBox)c).Multiline) SendMessage(c.Handle, 0x1501, (IntPtr)1, text);
            else if (c is ComboBox) SendMessage(c.Handle, 0x1703, IntPtr.Zero, text);
        } catch { }
    }

    public static int InputHeight {
        get {
            TextBox t = new TextBox();
            t.Font = St.Base;
            t.BorderStyle = BorderStyle.FixedSingle;
            int h = t.PreferredHeight;
            t.Dispose();
            return Math.Max(h, 24) + 2;
        }
    }

    public static TextBox Text(int lines) {
        TextBox t = new TextBox();
        t.Font = St.Base;
        t.BorderStyle = BorderStyle.FixedSingle;
        t.ForeColor = St.Text;
        if (lines > 1) {
            t.Multiline = true;
            t.ScrollBars = ScrollBars.Vertical;
            t.AcceptsReturn = true;
            t.Tag = lines;
        }
        return t;
    }

    public static Label Caption(string text) {
        Label l = new Label();
        l.AutoSize = false;
        l.Text = text;
        l.Font = St.Small;
        l.ForeColor = St.Caption;
        l.BackColor = Color.Transparent;
        l.TextAlign = ContentAlignment.BottomLeft;
        l.AutoEllipsis = true;
        return l;
    }

    public static Label Note(string text) {
        Label l = new Label();
        l.AutoSize = false;
        l.Text = text;
        l.Font = St.Small;
        l.ForeColor = St.Hint;
        l.BackColor = Color.Transparent;
        return l;
    }

    public static Button Button(string text, bool accent) {
        Button b = new Button();
        b.Text = text;
        b.FlatStyle = FlatStyle.Flat;
        b.Cursor = Cursors.Hand;
        b.UseVisualStyleBackColor = false;
        if (accent) {
            b.Font = St.Bold;
            b.BackColor = St.Accent;
            b.ForeColor = Color.White;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = St.AccentHover;
            b.FlatAppearance.MouseDownBackColor = St.AccentHover;
        } else {
            b.Font = St.Base;
            b.BackColor = Color.White;
            b.ForeColor = St.Text;
            b.FlatAppearance.BorderColor = St.InputBorder;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(236, 242, 250);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(222, 233, 247);
        }
        return b;
    }

    public static LinkLabel Link(string text) {
        LinkLabel l = new LinkLabel();
        l.Text = text;
        l.AutoSize = true;
        l.Font = St.Small;
        l.LinkColor = St.Accent;
        l.ActiveLinkColor = St.AccentHover;
        l.LinkBehavior = LinkBehavior.HoverUnderline;
        l.BackColor = Color.Transparent;
        return l;
    }

    // высота элемента в сетке
    public static int HeightOf(Control c, int width) {
        if (c is TextBox && ((TextBox)c).Multiline) {
            int lines = c.Tag is int ? (int)c.Tag : 3;
            return lines * c.Font.Height + 10;
        }
        if (c is Label && !(c is LinkLabel)) {
            if (c.Text.Length == 0) return 0;
            Size s = TextRenderer.MeasureText(c.Text, c.Font, new Size(Math.Max(width, 10), int.MaxValue), TextFormatFlags.WordBreak);
            return s.Height + 4;
        }
        if (c is CheckBox) return Math.Max(24, c.Font.Height + 8);
        if (c is Heaters) return ((Heaters)c).NeededHeight;
        if (c is ComboBox) return Math.Max(((ComboBox)c).PreferredHeight, c.Height);
        return InputHeightCached;
    }

    static int inputHeight;
    public static int InputHeightCached { get { if (inputHeight == 0) inputHeight = InputHeight; return inputHeight; } }

    // колесо мыши прокручивает форму, а не меняет значение в списке
    public static void Wheel(Control c, MouseEventArgs e) {
        HandledMouseEventArgs h = e as HandledMouseEventArgs;
        if (h != null) h.Handled = true;
        for (Control p = c.Parent; p != null; p = p.Parent) {
            ScrollableControl s = p as ScrollableControl;
            if (s != null && s.AutoScroll) {
                Point pos = s.AutoScrollPosition;
                s.AutoScrollPosition = new Point(-pos.X, Math.Max(0, -pos.Y - e.Delta));
                return;
            }
        }
    }

    public static GraphicsPath Round(Rectangle r, int radius) {
        GraphicsPath p = new GraphicsPath();
        int d = radius * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

// Выпадающий список: колесо мыши не меняет выбор; длинные варианты переносятся на несколько строк
public class Combo : ComboBox {
    public bool Wrap;

    public Combo(bool editable, bool wrap) {
        Font = St.Base;
        DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList;
        FlatStyle = FlatStyle.Flat;
        ForeColor = St.Text;
        MaxDropDownItems = 14;
        Wrap = wrap;
        if (editable) {
            AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            AutoCompleteSource = AutoCompleteSource.ListItems;
        }
        if (wrap) {
            DrawMode = DrawMode.OwnerDrawVariable;
            DropDownHeight = 420;
        }
    }

    static readonly ToolTip tip = new ToolTip();

    // длинный текст: показываем начало, целиком - во всплывающей подсказке
    void ShowStart() {
        if (!Wrap || DropDownStyle != ComboBoxStyle.DropDown) return;
        tip.SetToolTip(this, Text.Length > 60 ? Text : "");
        if (!IsHandleCreated) return;
        BeginInvoke((MethodInvoker)delegate { try { Select(0, 0); } catch { } });
    }

    protected override void OnSelectedIndexChanged(EventArgs e) {
        base.OnSelectedIndexChanged(e);
        ShowStart();
    }

    protected override void OnMouseWheel(MouseEventArgs e) {
        if (DroppedDown) base.OnMouseWheel(e);
        else Ui.Wheel(this, e);
    }

    protected override void OnDropDown(EventArgs e) {
        // ширина списка - по самому длинному варианту, но не шире экрана
        int w = Width;
        if (!Wrap) {
            foreach (object o in Items) w = Math.Max(w, TextRenderer.MeasureText(o.ToString(), Font).Width + 30);
            w = Math.Min(w, Screen.FromControl(this).WorkingArea.Width - 40);
        } else w = Math.Max(Width, 560);
        DropDownWidth = w;
        base.OnDropDown(e);
    }

    int TextWidth { get { return Math.Max(Width, 560) - 14; } }

    protected override void OnMeasureItem(MeasureItemEventArgs e) {
        if (e.Index >= 0 && e.Index < Items.Count) {
            Size s = TextRenderer.MeasureText(e.Graphics, Items[e.Index].ToString(), Font, new Size(TextWidth, int.MaxValue), TextFormatFlags.WordBreak);
            e.ItemHeight = s.Height + 8;
        }
        base.OnMeasureItem(e);
    }

    protected override void OnDrawItem(DrawItemEventArgs e) {
        if (!Wrap) { base.OnDrawItem(e); return; }
        if (e.Index < 0) return;
        bool sel = (e.State & DrawItemState.Selected) != 0;
        using (SolidBrush b = new SolidBrush(sel ? St.Accent : Color.White)) e.Graphics.FillRectangle(b, e.Bounds);
        Rectangle r = new Rectangle(e.Bounds.X + 5, e.Bounds.Y + 4, e.Bounds.Width - 10, e.Bounds.Height - 6);
        TextRenderer.DrawText(e.Graphics, Items[e.Index].ToString(), Font, r, sel ? Color.White : St.Text, TextFormatFlags.WordBreak | TextFormatFlags.Left | TextFormatFlags.Top);
        if (!sel) using (Pen p = new Pen(St.Border)) e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
    }

    public void SetItems(List<string> items) {
        string text = Text;
        BeginUpdate();
        Items.Clear();
        foreach (string s in items) if (s.Length > 0 && !Items.Contains(s)) Items.Add(s);
        EndUpdate();
        if (DropDownStyle == ComboBoxStyle.DropDownList) {
            int i = Items.IndexOf(text);
            SelectedIndex = i;
        } else Text = text;
    }

    public string Value {
        get { return Text.Trim(); }
        set {
            string v = value ?? "";
            if (DropDownStyle == ComboBoxStyle.DropDownList) {
                int i = Items.IndexOf(v);
                if (i < 0 && v.Length > 0) i = Items.Add(v);
                SelectedIndex = i;
            } else { Text = v; ShowStart(); }
        }
    }
}

// Выбор из нескольких вариантов кнопками: [ да | нет ], [ 1 ф | 3 ф ] ...
public class Seg : Panel {
    readonly RadioButton[] buttons;
    public event EventHandler Changed;
    bool missing;

    public Seg(params string[] options) {
        BackColor = St.InputBorder;
        Padding = new Padding(1);
        buttons = new RadioButton[options.Length];
        for (int i = 0; i < options.Length; i++) {
            RadioButton b = new RadioButton();
            b.Appearance = Appearance.Button;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.CheckedBackColor = St.Accent;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 240, 250);
            b.BackColor = Color.White;
            b.ForeColor = St.Text;
            b.Text = options[i];
            b.TextAlign = ContentAlignment.MiddleCenter;
            b.Font = St.Base;
            b.Cursor = Cursors.Hand;
            b.UseVisualStyleBackColor = false;
            b.CheckedChanged += delegate(object s, EventArgs e) {
                RadioButton r = (RadioButton)s;
                r.ForeColor = r.Checked ? Color.White : St.Text;
                r.BackColor = r.Checked ? St.Accent : Color.White;
                if (r.Checked && Changed != null) Changed(this, EventArgs.Empty);
            };
            buttons[i] = b;
            Controls.Add(b);
        }
    }

    protected override void OnLayout(LayoutEventArgs e) {
        base.OnLayout(e);
        if (buttons == null) return;
        int n = buttons.Length, w = ClientSize.Width - 2, h = ClientSize.Height - 2;
        for (int i = 0; i < n; i++) {
            int x0 = 1 + w * i / n, x1 = 1 + w * (i + 1) / n;
            if (buttons[i] != null) buttons[i].SetBounds(x0, 1, x1 - x0 - (i < n - 1 ? 1 : 0), h);
        }
    }

    public bool Missing {
        get { return missing; }
        set { missing = value; BackColor = value ? St.MissingBorder : St.InputBorder; }
    }

    public int Index {
        get { for (int i = 0; i < buttons.Length; i++) if (buttons[i].Checked) return i; return -1; }
        set {
            for (int i = 0; i < buttons.Length; i++) buttons[i].Checked = i == value;
            if (value < 0 && Changed != null) Changed(this, EventArgs.Empty);
        }
    }

    public string Value {
        get { int i = Index; return i < 0 ? "" : buttons[i].Text; }
        set {
            int k = -1;
            for (int i = 0; i < buttons.Length; i++) if (buttons[i].Text == value) k = i;
            Index = k;
        }
    }

    public RadioButton Button(int i) { return buttons[i]; }
}

// Белая карточка раздела с рамкой
public class CardPanel : Panel {
    public CardPanel() {
        BackColor = St.Card;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
    }
    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        using (Pen p = new Pen(St.Border)) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
    }
}

// Светлая рамка внутри карточки (счётчик 1, счётчик 2)
public class BoxPanel : Panel {
    public string Title = "";
    public BoxPanel() {
        BackColor = St.Box;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
    }
    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        using (Pen p = new Pen(St.Border)) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
        TextRenderer.DrawText(e.Graphics, Title, St.Bold, new Point(14, 10), Enabled ? St.Text : St.Hint);
    }
}

// Ячейка сетки: подпись над полем; поле может занимать несколько колонок
public class Cell {
    public string Caption;
    public Label Cap;
    public Control Ctl;
    public Grid Sub;
    public int Span = 1;

    public static Cell F(string caption, Control c, int span) {
        Cell x = new Cell();
        x.Caption = caption;
        x.Ctl = c;
        x.Span = span;
        return x;
    }
    public static Cell G(Grid g, int span) {
        Cell x = new Cell();
        x.Sub = g;
        x.Span = span;
        return x;
    }
    public static Cell Empty(int span) {
        Cell x = new Cell();
        x.Span = span;
        return x;
    }
}

public class Grid {
    public readonly Control Host;
    public readonly int Cols;
    public int PadX = 20, PadTop = 0, PadBottom = 20, Gap = 14, RowGap = 12;
    readonly List<Cell[]> rows = new List<Cell[]>();

    public Grid(Control host, int cols) { Host = host; Cols = cols; }

    public void Row(params Cell[] cells) {
        foreach (Cell c in cells) {
            if (c.Caption != null) { c.Cap = Ui.Caption(c.Caption); Host.Controls.Add(c.Cap); }
            if (c.Ctl != null) Host.Controls.Add(c.Ctl);
            if (c.Sub != null) Host.Controls.Add(c.Sub.Host);
        }
        rows.Add(cells);
    }

    public IEnumerable<Cell> Cells() {
        foreach (Cell[] r in rows) foreach (Cell c in r) yield return c;
    }

    public int Layout(int width) {
        int capH = St.Small.Height + 2;
        int y = PadTop;
        float colW = (width - 2 * PadX - Gap * (Cols - 1)) / (float)Cols;
        foreach (Cell[] row in rows) {
            int col = 0, rowH = 0;
            bool anyCap = false;
            foreach (Cell c in row) if (c.Caption != null) anyCap = true;
            foreach (Cell c in row) {
                int cx = PadX + (int)Math.Round(col * (colW + Gap));
                int cw = (int)Math.Round(c.Span * colW + (c.Span - 1) * Gap);
                int cy = y;
                // поля одной строки - на одном уровне, даже если у какого-то нет подписи
                if (c.Cap != null) c.Cap.SetBounds(cx, cy, cw, capH);
                if (anyCap && c.Sub == null) cy += capH + 3;
                int h = 0;
                if (c.Sub != null) { h = c.Sub.Layout(cw); c.Sub.Host.SetBounds(cx, cy, cw, h); }
                else if (c.Ctl != null) { h = Ui.HeightOf(c.Ctl, cw); c.Ctl.SetBounds(cx, cy, cw, h); }
                rowH = Math.Max(rowH, cy + h - y);
                col += c.Span;
            }
            y += rowH + RowGap;
        }
        return y - (rows.Count > 0 ? RowGap : 0) + PadBottom;
    }
}

// Раздел формы: заголовок, справа от него - дополнительные элементы, ниже - сетка полей
public class Card {
    public readonly string Name;
    public readonly CardPanel Panel = new CardPanel();
    public readonly Label Title = new Label();
    public readonly Label Sub = new Label();
    public readonly Grid Grid;
    public readonly List<Control> Extra = new List<Control>();

    public Card(string name, string subtitle, int cols) {
        Name = name;
        Title.Text = name;
        Title.Font = St.Title;
        Title.ForeColor = St.Text;
        Title.AutoSize = true;
        Title.BackColor = Color.Transparent;
        Sub.Text = subtitle ?? "";
        Sub.Font = St.Small;
        Sub.ForeColor = St.Hint;
        Sub.AutoSize = true;
        Sub.BackColor = Color.Transparent;
        Panel.Controls.Add(Title);
        Panel.Controls.Add(Sub);
        Grid = new Grid(Panel, cols);
        Grid.PadTop = 58;
    }

    public void AddExtra(Control c) {
        Extra.Add(c);
        Panel.Controls.Add(c);
    }

    public int Layout(int width) {
        Panel.Width = width;
        Title.Location = new Point(18, 14);
        Sub.Location = new Point(Title.Right + 10, 14 + Title.Height - Sub.Height - 2);
        int x = width - 20;
        for (int i = Extra.Count - 1; i >= 0; i--) {
            Control c = Extra[i];
            x -= c.Width;
            c.Location = new Point(x, 12 + Math.Max(0, (Title.Height - c.Height) / 2));
            x -= 12;
        }
        int h = Grid.Layout(width);
        Panel.Height = h;
        return h;
    }
}

// Пункт меню слева: название раздела и сколько в нём не заполнено
public class NavItem : Control {
    public bool Active, Small;
    public int Count = -1;
    bool hover;

    public NavItem(string text, bool small) {
        Text = text;
        Small = small;
        Font = small ? St.Small : St.NavFont;
        Height = small ? 32 : 40;
        Cursor = Cursors.Hand;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        Color back = Active ? St.NavActive : hover ? St.NavHover : St.Nav;
        using (SolidBrush b = new SolidBrush(back)) g.FillRectangle(b, ClientRectangle);
        if (Active) using (SolidBrush b = new SolidBrush(St.NavMark)) g.FillRectangle(b, 0, 0, 4, Height);
        Color fore = Active || hover ? Color.White : St.NavText;
        Rectangle tr = new Rectangle(22, 0, Width - 60, Height);
        TextRenderer.DrawText(g, Text, Font, tr, fore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
        if (Count > 0) {
            string s = Count.ToString();
            Size ts = TextRenderer.MeasureText(s, St.Small);
            int w = Math.Max(20, ts.Width + 6), h = 20;
            Rectangle r = new Rectangle(Width - 18 - w, (Height - h) / 2, w, h);
            using (GraphicsPath p = Ui.Round(r, 9)) using (SolidBrush b = new SolidBrush(Color.FromArgb(214, 76, 76))) g.FillPath(b, p);
            TextRenderer.DrawText(g, s, St.Small, r, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        } else if (Count == 0) {
            int cx = Width - 28, cy = Height / 2;
            using (Pen p = new Pen(Color.FromArgb(120, 200, 150), 2f)) {
                g.DrawLine(p, cx - 5, cy, cx - 1, cy + 4);
                g.DrawLine(p, cx - 1, cy + 4, cx + 6, cy - 4);
            }
        }
    }
}

// Всплывающее окошко под полем: не забирает фокус у поля (можно дальше печатать), закрывается само
public class Popup : Form {
    public Popup() {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = St.InputBorder;
        Padding = new Padding(1);
        KeyPreview = true;
        KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };
        Deactivate += delegate { BeginInvoke((MethodInvoker)delegate { if (!IsDisposed) Close(); }); };
    }

    protected override bool ShowWithoutActivation { get { return true; } }

    // под элементом, но не за краем экрана
    public void ShowUnder(Control c) {
        Rectangle wa = Screen.FromControl(c).WorkingArea;
        Point p = c.PointToScreen(new Point(0, c.Height + 2));
        if (p.Y + Height > wa.Bottom) p.Y = c.PointToScreen(Point.Empty).Y - Height - 2;
        if (p.X + Width > wa.Right) p.X = wa.Right - Width;
        Location = p;
        Show(c.FindForm());
    }
}

// Дата: поле + кнопка-календарик; календарик открывается и по щелчку в поле
public class DateBox : Panel {
    readonly TextBox box = new TextBox();
    readonly Button pick = new Button();
    DateTime value = DateTime.Today;
    Popup popup;
    public event EventHandler ValueChanged;

    public DateBox() {
        BackColor = Color.Transparent;
        box.Font = St.Base;
        box.BorderStyle = BorderStyle.FixedSingle;
        box.ForeColor = St.Text;
        box.Text = Text2;
        box.Leave += delegate { Parse(); };
        box.KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Down && e.Alt) { ShowCalendar(); e.Handled = true; }
            if (e.KeyCode == Keys.Escape && popup != null) popup.Close();
        };
        box.MouseUp += delegate { if (popup == null || popup.IsDisposed) ShowCalendar(); };
        pick.FlatStyle = FlatStyle.Flat;
        pick.FlatAppearance.BorderColor = St.InputBorder;
        pick.BackColor = Color.White;
        pick.UseVisualStyleBackColor = false;
        pick.Cursor = Cursors.Hand;
        pick.TabStop = false;
        pick.Paint += delegate(object s, PaintEventArgs e) { DrawIcon(e.Graphics, pick.ClientRectangle, pick.Enabled); };
        pick.Click += delegate { if (popup != null && !popup.IsDisposed) popup.Close(); else ShowCalendar(); };
        Controls.Add(box);
        Controls.Add(pick);
    }

    static void DrawIcon(Graphics g, Rectangle r, bool on) {
        Color c = on ? St.Accent : St.Hint;
        int w = 14, h = 13, x = r.X + (r.Width - w) / 2, y = r.Y + (r.Height - h) / 2;
        using (Pen p = new Pen(c, 1.4f)) {
            g.DrawRectangle(p, x, y + 1, w, h - 1);
            g.DrawLine(p, x, y + 4, x + w, y + 4);
            g.DrawLine(p, x + 3, y - 1, x + 3, y + 2);
            g.DrawLine(p, x + w - 3, y - 1, x + w - 3, y + 2);
        }
        using (SolidBrush b = new SolidBrush(c))
            for (int row = 0; row < 2; row++)
                for (int col = 0; col < 3; col++) g.FillRectangle(b, x + 2 + col * 4, y + 6 + row * 3, 2, 2);
    }

    protected override void OnLayout(LayoutEventArgs e) {
        base.OnLayout(e);
        int bw = Math.Max(28, box.Height);
        box.SetBounds(0, 0, Math.Max(10, Width - bw - 2), Height);
        pick.SetBounds(Width - bw, 0, bw, box.Height);
    }

    protected override void OnMouseWheel(MouseEventArgs e) { Ui.Wheel(this, e); }

    public DateTime Value {
        get { return value; }
        set {
            bool changed = value.Date != this.value.Date;
            this.value = value.Date;
            box.Text = Text2;
            if (changed && ValueChanged != null) ValueChanged(this, EventArgs.Empty);
        }
    }

    public string Text2 { get { return value.ToString("dd.MM.yyyy"); } }

    // набранная руками дата: 1.2.26, 01.02.2026, 01022026
    void Parse() {
        string s = box.Text.Trim().Replace(',', '.').Replace('/', '.').Replace('-', '.').Replace(' ', '.');
        if (s.IndexOf('.') < 0 && s.Length == 8) s = s.Substring(0, 2) + "." + s.Substring(2, 2) + "." + s.Substring(4);
        DateTime d;
        string[] formats = { "d.M.yyyy", "d.M.yy", "d.M" };
        if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) Value = d;
        else box.Text = Text2;
    }

    void ShowCalendar() {
        if (!Enabled) return;
        Parse();
        popup = new Popup();
        MonthCalendar cal = new MonthCalendar();
        cal.MaxSelectionCount = 1;
        cal.ShowTodayCircle = true;
        cal.SetDate(value);
        cal.Location = new Point(1, 1);
        popup.Controls.Add(cal);
        popup.ClientSize = new Size(cal.Width + 2, cal.Height + 2);
        cal.DateSelected += delegate(object s, DateRangeEventArgs e) {
            Value = e.Start;
            popup.Close();
            box.Focus();
            box.SelectAll();
        };
        popup.ShowUnder(this);
    }

    public new bool Focus() { return box.Focus(); }
}

// Электронагревательные приборы: строки «вид | шт | мощность», «+» добавляет строку.
// В договор - одной строкой через «; »: «Конвекторы - 2 шт. 1,5 кВт; Котёл - 1 шт. 6 кВт»
public class Heaters : Panel {
    public const string Own = "свое";
    class Row {
        public Combo Kind, Count;
        public Button Power, Remove;
        public string PowerValue = "";
    }
    readonly List<Row> rows = new List<Row>();
    readonly Button add = new Button();
    List<string> kinds = new List<string>();
    bool quiet;
    public event EventHandler Changed;      // поменялось значение
    public event EventHandler Resized;      // строк стало больше или меньше - перестроить форму

    static readonly string[] Counts = BuildCounts();
    static string[] BuildCounts() {
        List<string> l = new List<string>();
        for (int i = 1; i <= 15; i++) l.Add(i.ToString());
        l.Add(Own);
        return l.ToArray();
    }

    public Heaters() {
        BackColor = Color.Transparent;
        add.Text = "+";
        add.Font = St.Bold;
        add.FlatStyle = FlatStyle.Flat;
        add.FlatAppearance.BorderColor = St.Accent;
        add.ForeColor = St.Accent;
        add.BackColor = Color.White;
        add.UseVisualStyleBackColor = false;
        add.Cursor = Cursors.Hand;
        add.Click += delegate { AddRow(); Relayout(); rows[rows.Count - 1].Kind.Focus(); };
        Controls.Add(add);
        AddRow();
    }

    int RowH { get { return Ui.InputHeightCached; } }
    public int NeededHeight { get { return rows.Count * (RowH + 6) - 6; } }

    public void SetKinds(List<string> list) {
        kinds = new List<string>(list);
        if (!kinds.Contains(Own)) kinds.Add(Own);
        quiet = true;
        try { foreach (Row r in rows) r.Kind.SetItems(kinds); } finally { quiet = false; }
    }

    Row AddRow() {
        Row r = new Row();
        r.Kind = new Combo(true, false);
        r.Kind.SetItems(kinds);
        r.Count = new Combo(true, false);
        r.Count.SetItems(new List<string>(Counts));
        r.Power = new Button();
        r.Power.FlatStyle = FlatStyle.Flat;
        r.Power.FlatAppearance.BorderColor = St.InputBorder;
        r.Power.BackColor = Color.White;
        r.Power.UseVisualStyleBackColor = false;
        r.Power.Font = St.Base;
        r.Power.TextAlign = ContentAlignment.MiddleLeft;
        r.Power.Cursor = Cursors.Hand;
        r.Remove = new Button();
        r.Remove.Text = "×";
        r.Remove.Font = St.Bold;
        r.Remove.FlatStyle = FlatStyle.Flat;
        r.Remove.FlatAppearance.BorderColor = St.InputBorder;
        r.Remove.ForeColor = St.Hint;
        r.Remove.BackColor = Color.White;
        r.Remove.UseVisualStyleBackColor = false;
        r.Remove.Cursor = Cursors.Hand;
        r.Remove.TabStop = false;
        ShowPower(r);
        Row row = r;
        r.Kind.TextChanged += delegate { KindChanged(row); };
        r.Kind.SelectedIndexChanged += delegate { KindChanged(row); };
        r.Count.TextChanged += delegate { CountChanged(row); };
        r.Count.SelectedIndexChanged += delegate { CountChanged(row); };
        r.Power.Click += delegate { PickPower(row); };
        r.Remove.Click += delegate { RemoveRow(row); };
        rows.Add(r);
        Controls.Add(r.Kind);
        Controls.Add(r.Count);
        Controls.Add(r.Power);
        Controls.Add(r.Remove);
        return r;
    }

    void RemoveRow(Row r) {
        if (rows.Count == 1) {
            quiet = true;
            try { r.Kind.Text = ""; r.Count.Text = ""; r.PowerValue = ""; ShowPower(r); } finally { quiet = false; }
        } else {
            rows.Remove(r);
            foreach (Control c in new Control[] { r.Kind, r.Count, r.Power, r.Remove }) { Controls.Remove(c); c.Dispose(); }
            Relayout();
        }
        Fire();
    }

    void Relayout() {
        PerformLayout();
        if (Resized != null) Resized(this, EventArgs.Empty);
    }

    // «свое» - поле очищается для своего значения
    static bool TakeOwn(Combo c) {
        if (c.Text.Trim() != Own) return false;
        c.BeginInvoke((MethodInvoker)delegate { c.Text = ""; c.Focus(); });
        return true;
    }

    void KindChanged(Row r) {
        if (quiet) return;
        if (TakeOwn(r.Kind)) return;
        // выбрали вид - одна штука, если количество ещё не указано
        if (r.Kind.Text.Trim().Length > 0 && r.Count.Text.Trim().Length == 0 && r.Kind.SelectedIndex >= 0) {
            quiet = true;
            try { r.Count.Text = "1"; } finally { quiet = false; }
        }
        Fire();
    }

    void CountChanged(Row r) {
        if (quiet) return;
        if (TakeOwn(r.Count)) return;
        Fire();
    }

    void Fire() { if (!quiet && Changed != null) Changed(this, EventArgs.Empty); }

    void ShowPower(Row r) {
        r.Power.Text = r.PowerValue.Length > 0 ? r.PowerValue + " кВт" : "мощность...";
        r.Power.ForeColor = r.PowerValue.Length > 0 ? St.Text : St.Hint;
    }

    // мощность - окошко с вариантами 0,5 ... 25 кВт и полем для своего значения
    void PickPower(Row r) {
        Popup p = new Popup();
        p.BackColor = Color.White;
        p.Padding = new Padding(0);
        p.Paint += delegate(object s, PaintEventArgs e) { using (Pen pen = new Pen(St.InputBorder)) e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1); };
        int cols = 10, bw = 52, bh = 30, pad = 8;
        Label cap = Ui.Caption("Мощность одного прибора, кВт");
        cap.SetBounds(pad, 6, cols * bw, 20);
        p.Controls.Add(cap);
        for (int i = 1; i <= 50; i++) {
            string v = Kw(i * 0.5);
            Button b = new Button();
            b.Text = v;
            b.Font = St.Base;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderColor = St.Border;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 240, 250);
            b.UseVisualStyleBackColor = false;
            b.BackColor = v == r.PowerValue ? St.Accent : Color.White;
            b.ForeColor = v == r.PowerValue ? Color.White : St.Text;
            b.Cursor = Cursors.Hand;
            b.SetBounds(pad + ((i - 1) % cols) * bw, 30 + ((i - 1) / cols) * bh, bw - 2, bh - 2);
            b.Click += delegate { SetPower(r, v); p.Close(); };
            p.Controls.Add(b);
        }
        int y = 30 + 5 * bh + 6;
        Label own = Ui.Caption("свое:");
        own.TextAlign = ContentAlignment.MiddleLeft;
        own.SetBounds(pad, y, 46, 26);
        TextBox t = Ui.Text(1);
        t.SetBounds(pad + 48, y + 1, 120, 26);
        Button ok = Ui.Button("OK", true);
        ok.SetBounds(pad + 176, y, 60, 28);
        ok.Click += delegate { SetPower(r, Clean(t.Text)); p.Close(); };
        t.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SetPower(r, Clean(t.Text)); p.Close(); } };
        Button none = Ui.Button("без мощности", false);
        none.Font = St.Small;
        none.SetBounds(pad + cols * bw - 120, y, 118, 28);
        none.Click += delegate { SetPower(r, ""); p.Close(); };
        p.Controls.Add(own);
        p.Controls.Add(t);
        p.Controls.Add(ok);
        p.Controls.Add(none);
        p.ClientSize = new Size(pad * 2 + cols * bw - 2, y + 28 + pad);
        p.ShowUnder(r.Power);
        p.Activate();
        t.Focus();
    }

    static string Clean(string s) {
        s = s.Replace("кВт", "").Replace("квт", "").Trim().Replace('.', ',');
        double d;
        if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d > 0) return Kw(d);
        return s;
    }

    static string Kw(double d) { return d.ToString("0.##", CultureInfo.InvariantCulture).Replace('.', ','); }

    void SetPower(Row r, string v) {
        r.PowerValue = v;
        ShowPower(r);
        Fire();
    }

    protected override void OnLayout(LayoutEventArgs e) {
        base.OnLayout(e);
        if (rows == null || add == null) return;
        int h = RowH, gap = 6, w = ClientSize.Width;
        int bw = h, cw = 70, pw = 110;
        int kw = Math.Max(60, w - cw - pw - bw * 2 - gap * 4);
        for (int i = 0; i < rows.Count; i++) {
            Row r = rows[i];
            int y = i * (h + gap), x = 0;
            r.Kind.SetBounds(x, y, kw, h); x += kw + gap;
            r.Count.SetBounds(x, y, cw, h); x += cw + gap;
            r.Power.SetBounds(x, y, pw, h); x += pw + gap;
            r.Remove.SetBounds(x, y, bw, h); x += bw + gap;
            if (i == rows.Count - 1) add.SetBounds(x, y, bw, h);
        }
    }

    // в договор: «Конвекторы - 2 шт. 1,5 кВт; Котёл - 6 кВт»
    public string Value {
        get {
            List<string> parts = new List<string>();
            foreach (Row r in rows) {
                string kind = Plain(r.Kind.Text), n = Plain(r.Count.Text), p = r.PowerValue;
                if (kind.Length == 0 && n.Length == 0 && p.Length == 0) continue;
                List<string> d = new List<string>();
                if (n.Length > 0) d.Add(n + " шт.");
                if (p.Length > 0) d.Add(p + " кВт");
                parts.Add(kind + (d.Count > 0 ? (kind.Length > 0 ? " - " : "") + string.Join(" ", d.ToArray()) : ""));
            }
            return string.Join("; ", parts.ToArray());
        }
        set {
            quiet = true;
            try {
                while (rows.Count > 1) { Row r = rows[rows.Count - 1]; rows.Remove(r); foreach (Control c in new Control[] { r.Kind, r.Count, r.Power, r.Remove }) { Controls.Remove(c); c.Dispose(); } }
                Row first = rows[0];
                first.Kind.Text = ""; first.Count.Text = ""; first.PowerValue = ""; ShowPower(first);
                string s = (value ?? "").Replace("\r\n", "; ").Replace("\n", "; ");
                bool used = false;
                foreach (string part in s.Split(';')) {
                    string t = part.Trim();
                    if (t.Length == 0) continue;
                    Row r = used ? AddRow() : first;
                    used = true;
                    Match m = Pattern.Match(t);
                    if (m.Success) {
                        r.Kind.Text = m.Groups["k"].Value.Trim();
                        r.Count.Text = m.Groups["n"].Value;
                        r.PowerValue = m.Groups["p"].Value;
                    } else r.Kind.Text = t;
                    ShowPower(r);
                }
            } finally { quiet = false; }
            Relayout();
        }
    }

    static string Plain(string s) { s = s.Trim(); return s == Own ? "" : s; }

    static readonly Regex Pattern = new Regex(@"^(?<k>.*?)\s*-?\s*(?:(?<n>\d+)\s*шт\.?)?\s*-?\s*(?:(?<p>\d+(?:[.,]\d+)?)\s*кВт)?$", RegexOptions.IgnoreCase);
}

// Движок чтения базы FoxPro - из программы «Поиск абонента» (только чтение)
// Код написан на C# 2.0, чтобы собираться и в старом PowerShell 2.0 (Windows 7).

public class DbfField {
    public string Name;
    public char Type;
    public int Offset;
    public int Length;
    public int Decimals;   // знаков после запятой у числовых полей
}

// Сообщает, сколько записей уже прочитано (чтобы окно не "зависало" на больших таблицах)
public delegate void LoadProgress(int done, int expected);

// Решает по записи как она есть в файле, оставлять ли её (например, только улицы своих сельсоветов)
public delegate bool RecordFilter(byte[] record, int start);

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
    public int Expected;  // сколько записей написано в заголовке файла
    public long BytesRead;
    public double Seconds;
    public List<int> RecNos = new List<int>();   // номер каждой записи в файле (с 1)
    public bool Truncated;                       // подходящих записей было больше, прочитана часть
    byte[] headerBytes;
    int headerLength;

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
        return LoadCore(path, otherEncoding, query, progress, false, null, 0, 0);
    }

    // maxKeep - больше стольких подходящих записей не набирать (перебор с очень общим запросом)
    public static DbfTable Load(string path, string query, RecordFilter keep, LoadProgress progress, int maxKeep) {
        return LoadCore(path, false, query, progress, false, keep, 0, maxKeep);
    }

    // keep - какие записи оставлять; maxRecords > 0 - прочитать только начало файла
    public static DbfTable Load(string path, RecordFilter keep, LoadProgress progress, int maxRecords) {
        return LoadCore(path, false, null, progress, false, keep, maxRecords, 0);
    }

    // Только заголовок: список полей и число записей, без чтения самих записей
    public static DbfTable LoadHeader(string path) {
        return LoadCore(path, false, null, null, true, null, 0, 0);
    }

    static DbfTable LoadCore(string path, bool otherEncoding, string query, LoadProgress progress, bool headerOnly,
                             RecordFilter keep, int maxRecords, int maxKeep) {
        DateTime started = DateTime.Now;
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
                else f.Decimals = desc[p + 17];
                f.Offset = offset;
                offset += f.Length;
                if (f.Type != '0') t.Fields.Add(f);                  // '0' - служебное поле VFP
            }
            if (offset > recordLen)
                throw new InvalidDataException("Не удалось разобрать структуру таблицы.");
            t.Expected = count;
            t.headerLength = headerLen;
            t.headerBytes = new byte[headerLen];
            Buffer.BlockCopy(head, 0, t.headerBytes, 0, 32);
            Buffer.BlockCopy(desc, 0, t.headerBytes, 32, desc.Length);
            if (headerOnly) return t;

            t.PrepareLower();
            SearchWord[] filter = query == null ? null : t.PrepareWords(query);

            fs.Seek(headerLen, SeekOrigin.Begin);
            const int chunkRecords = 2048;
            byte[] chunk = new byte[chunkRecords * recordLen];
            byte[] low = new byte[t.lowerLength];
            int index = 0;
            int limit = maxRecords > 0 ? Math.Min(count, maxRecords) : count;
            bool end = false;
            while (index < limit && !end) {
                int want = Math.Min(chunkRecords, limit - index);
                int got = ReadUpTo(fs, chunk, want * recordLen) / recordLen;
                t.BytesRead += (long)got * recordLen;
                for (int r = 0; r < got; r++) {
                    int start = r * recordLen;
                    byte flag = chunk[start];
                    if (flag == 0x1A) { end = true; break; }          // конец файла
                    if (flag == (byte)'*') { t.Deleted++; continue; }  // удалённая запись
                    t.Total++;
                    if (keep != null && !keep(chunk, start)) continue;
                    t.MakeLower(chunk, start, low);
                    if (filter == null || MatchesAll(low, 0, t.lowerLength, filter)) t.Append(chunk, start, low, index + r + 1);
                    if (maxKeep > 0 && t.Count >= maxKeep) { t.Truncated = true; end = true; break; }
                }
                index += got;
                if (got < want) break;   // файл короче, чем написано в его заголовке
                if (progress != null) progress(index, limit);
            }
        }
        t.Seconds = (DateTime.Now - started).TotalSeconds;
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

    // поисковая копия: только текстовые поля (двоичные числа и ссылки на примечания - пропуск),
    // после каждого поля - разделитель
    bool[] lowerText;

    void PrepareLower() {
        int fieldCount = Fields.Count;
        lowerText = new bool[fieldCount];
        lowerOffsets = new int[fieldCount];
        int pos = 0;
        for (int j = 0; j < fieldCount; j++) {
            lowerOffsets[j] = pos;
            pos += Fields[j].Length + 1;
            lowerText[j] = "CNFDL".IndexOf(Fields[j].Type) >= 0;
        }
        lowerLength = Math.Max(pos, 1);
    }

    void MakeLower(byte[] src, int start, byte[] low) {
        for (int j = 0; j < Fields.Count; j++) {
            DbfField f = Fields[j];
            int from = start + f.Offset, dst = lowerOffsets[j];
            if (lowerText[j]) for (int k = 0; k < f.Length; k++) low[dst + k] = lowerMap[src[from + k]];
            else for (int k = 0; k < f.Length; k++) low[dst + k] = 1;
            low[dst + f.Length] = 1;
        }
    }

    public int RecordLength { get { return recordLength; } }
    public int HeaderLength { get { return headerLength; } }

    // Пустая таблица той же структуры - в неё складываются записи, прочитанные по номерам
    public static DbfTable CreateLike(DbfTable h) {
        DbfTable t = new DbfTable();
        t.FilePath = h.FilePath;
        t.Fields = h.Fields;
        t.Enc = h.Enc;
        t.recordLength = h.recordLength;
        t.headerLength = h.headerLength;
        t.headerBytes = h.headerBytes;
        t.Expected = h.Expected;
        t.BuildLowerMap();
        t.PrepareLower();
        return t;
    }

    public void AddRecord(byte[] src, int start, int recNo) {
        if (lowerOffsets == null) PrepareLower();
        byte[] low = new byte[lowerLength];
        MakeLower(src, start, low);
        Append(src, start, low, recNo);
        Total++;
    }

    // запись как в файле: буфер и место в нём
    public byte[] RecordBuffer(int row, out int start) {
        start = (row % BlockRecords) * recordLength;
        return data[row / BlockRecords];
    }

    void Append(byte[] src, int start, byte[] low, int recNo) {
        RecNos.Add(recNo);
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
        for (int i = 0; i < words.Length; i++) result[i] = PrepareWord(words[i]);
        return result;
    }

    public SearchWord PrepareWord(string word) {
        string w = word.ToLowerInvariant();
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
        return sw;
    }

    public int FieldIndex(string name) {
        for (int i = 0; i < Fields.Count; i++)
            if (string.Compare(Fields[i].Name, name, StringComparison.OrdinalIgnoreCase) == 0) return i;
        return -1;
    }

    // слово встречается в поле; field = -1 - где угодно в записи
    public bool FieldContains(int row, int field, SearchWord w) {
        byte[] buf = lower[row / BlockRecords];
        int start = (row % BlockRecords) * lowerLength;
        if (field < 0) return MatchWord(buf, start, start + lowerLength, w);
        int from = start + lowerOffsets[field];
        return MatchWord(buf, from, from + Fields[field].Length, w);
    }

    // поле целиком (без пробелов по краям) равно слову: дом "18" - это 18, а не 118 и не 18а
    public bool FieldEquals(int row, int field, SearchWord w) {
        byte[] buf = lower[row / BlockRecords];
        int from = (row % BlockRecords) * lowerLength + lowerOffsets[field];
        int end = from + Fields[field].Length;
        while (from < end && buf[from] <= 32) from++;
        while (end > from && buf[end - 1] <= 32) end--;
        if (end - from != w.Bytes.Length) return false;
        for (int k = 0; k < w.Bytes.Length; k++) if (buf[from + k] != w.Bytes[k]) return false;
        return true;
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
        foreach (SearchWord w in words) if (!MatchWord(buf, from, from + length, w)) return false;
        return true;
    }

    static bool MatchWord(byte[] buf, int from, int end, SearchWord w) {
        if (w.WholeNumber)
            return ContainsNumber(buf, from, end, w.Bytes) || (w.Trimmed != null && ContainsNumber(buf, from, end, w.Trimmed));
        return IndexOf(buf, from, end, w.Bytes) >= 0 || (w.Trimmed != null && IndexOf(buf, from, end, w.Trimmed) >= 0);
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

    internal static string Xml(string s) {
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

// ---- выражения FoxPro: по ним построены индексы (ключ и условие FOR) ----

public class FoxValue {
    public char Type;      // C - текст, N - число, D - дата (номер дня), L - да/нет
    public byte[] Text;
    public double Num;
    public bool Log;

    public static FoxValue Chars(byte[] b) { FoxValue v = new FoxValue(); v.Type = 'C'; v.Text = b; return v; }
    public static FoxValue Number(double d) { FoxValue v = new FoxValue(); v.Type = 'N'; v.Num = d; return v; }
    public static FoxValue Date(double d) { FoxValue v = new FoxValue(); v.Type = 'D'; v.Num = d; return v; }
    public static FoxValue Logical(bool b) { FoxValue v = new FoxValue(); v.Type = 'L'; v.Log = b; return v; }
}

// Запись, на которой считается выражение
public class FoxContext {
    public DbfTable Table;
    public byte[] Rec;
    public int Start;
    public bool CyrillicUpper = true;   // UPPER() переводит и русские буквы (как в вашей FoxPro - проверяется по индексу)
}

public class FoxExpr {
    // f - поле, v - готовое значение, c - функция, + - = # < > a (<=) b (>=) e (==) $ & (и) | (или) ! (не) m (минус числа)
    public char Op;
    public string Name;
    public int Field = -1;
    public FoxValue Value;
    public FoxExpr[] Args = new FoxExpr[0];

    static readonly string[] Functions = {
        "UPPER", "LOWER", "STR", "DTOS", "DTOC", "LEFT", "RIGHT", "SUBSTR", "ALLTRIM", "TRIM", "RTRIM", "LTRIM",
        "PADL", "PADR", "VAL", "INT", "EMPTY", "DELETED", "IIF", "SPACE" };

    public static FoxExpr Parse(string text, DbfTable t) {
        Parser p = new Parser(text, t);
        FoxExpr e = p.Or();
        p.Space();
        if (p.pos < text.Length) throw new FormatException("непонятно: " + text.Substring(p.pos));
        return e;
    }

    // поля, из которых считается выражение
    public void CollectFields(List<int> list) {
        if (Op == 'f' && !list.Contains(Field)) list.Add(Field);
        foreach (FoxExpr a in Args) a.CollectFields(list);
    }

    // слагаемые верхнего уровня: STR(UL,5)+DOM+KV -> STR(UL,5), DOM, KV
    public List<FoxExpr> Terms() {
        List<FoxExpr> list = new List<FoxExpr>();
        if (Op == '+') { list.AddRange(Args[0].Terms()); list.AddRange(Args[1].Terms()); }
        else list.Add(this);
        return list;
    }

    // начало значения поля даёт начало ключа: FIO, UPPER(FIO), LEFT(FIO,20) - да; STR(), ALLTRIM() - нет
    public bool PrefixSafe() {
        if (Op == 'f') return true;
        if (Op != 'c') return false;
        if (Name == "UPPER" || Name == "LOWER" || Name == "PADR") return Args.Length > 0 && Args[0].PrefixSafe();
        if (Name == "LEFT") return Args.Length == 2 && Args[0].PrefixSafe();
        if (Name == "SUBSTR") return Args.Length >= 2 && Args[1].Op == 'v' && Args[1].Value.Num == 1 && Args[0].PrefixSafe();
        return false;
    }

    public bool UsesCase() {
        if (Op == 'c' && (Name == "UPPER" || Name == "LOWER")) return true;
        foreach (FoxExpr a in Args) if (a.UsesCase()) return true;
        return false;
    }

    // ---- вычисление ----

    public FoxValue Eval(FoxContext c) {
        switch (Op) {
            case 'f': return FieldValue(c, Field);
            case 'v': return Value;
            case 'c': return Call(c);
            case 'm': return FoxValue.Number(-Num(Args[0].Eval(c)));
            case '!': return FoxValue.Logical(!Bool(Args[0].Eval(c)));
            case '&': return FoxValue.Logical(Bool(Args[0].Eval(c)) && Bool(Args[1].Eval(c)));
            case '|': return FoxValue.Logical(Bool(Args[0].Eval(c)) || Bool(Args[1].Eval(c)));
        }
        FoxValue a = Args[0].Eval(c), b = Args[1].Eval(c);
        switch (Op) {
            case '+':
                if (a.Type == 'C' && b.Type == 'C') return FoxValue.Chars(Concat(a.Text, b.Text));
                if (a.Type == 'D' && b.Type == 'N') return FoxValue.Date(a.Num + b.Num);
                return FoxValue.Number(Num(a) + Num(b));
            case '-':
                if (a.Type == 'C' && b.Type == 'C') {
                    int end = a.Text.Length;
                    while (end > 0 && a.Text[end - 1] == ' ') end--;
                    byte[] r = new byte[a.Text.Length + b.Text.Length];
                    Buffer.BlockCopy(a.Text, 0, r, 0, end);
                    Buffer.BlockCopy(b.Text, 0, r, end, b.Text.Length);
                    for (int i = end + b.Text.Length; i < r.Length; i++) r[i] = (byte)' ';
                    return FoxValue.Chars(r);
                }
                if (a.Type == 'D' && b.Type == 'D') return FoxValue.Number(a.Num - b.Num);
                if (a.Type == 'D') return FoxValue.Date(a.Num - Num(b));
                return FoxValue.Number(Num(a) - Num(b));
            case '$':
                return FoxValue.Logical(Find(Str(b), Str(a)) >= 0);
        }
        int cmp = Compare(a, b, Op == '=' || Op == '#');
        switch (Op) {
            case '=': case 'e': return FoxValue.Logical(cmp == 0);
            case '#': return FoxValue.Logical(cmp != 0);
            case '<': return FoxValue.Logical(cmp < 0);
            case '>': return FoxValue.Logical(cmp > 0);
            case 'a': return FoxValue.Logical(cmp <= 0);
            default: return FoxValue.Logical(cmp >= 0);
        }
    }

    // "=" для текста - как в FoxPro при SET EXACT OFF: сравнение по длине правой части
    static int Compare(FoxValue a, FoxValue b, bool loose) {
        if (a.Type == 'C' && b.Type == 'C') {
            byte[] x = a.Text, y = b.Text;
            int n = loose ? y.Length : Math.Max(x.Length, y.Length);
            for (int i = 0; i < n; i++) {
                int p = i < x.Length ? x[i] : 32, q = i < y.Length ? y[i] : 32;
                if (p != q) return p < q ? -1 : 1;
            }
            return 0;
        }
        if (a.Type == 'L' && b.Type == 'L') return a.Log == b.Log ? 0 : a.Log ? 1 : -1;
        return Num(a).CompareTo(Num(b));
    }

    static double Num(FoxValue v) {
        if (v.Type == 'N' || v.Type == 'D') return v.Num;
        throw new NotSupportedException("ожидалось число");
    }

    static bool Bool(FoxValue v) {
        if (v.Type == 'L') return v.Log;
        throw new NotSupportedException("ожидалось да/нет");
    }

    static byte[] Str(FoxValue v) {
        if (v.Type == 'C') return v.Text;
        throw new NotSupportedException("ожидался текст");
    }

    static int Int(FoxValue v) {
        return (int)Num(v);
    }

    static byte[] Concat(byte[] a, byte[] b) {
        byte[] r = new byte[a.Length + b.Length];
        Buffer.BlockCopy(a, 0, r, 0, a.Length);
        Buffer.BlockCopy(b, 0, r, a.Length, b.Length);
        return r;
    }

    static byte[] Sub(byte[] s, int from, int len) {
        if (from < 0) from = 0;
        if (from > s.Length) from = s.Length;
        if (len > s.Length - from) len = s.Length - from;
        if (len < 0) len = 0;
        byte[] r = new byte[len];
        Buffer.BlockCopy(s, from, r, 0, len);
        return r;
    }

    static int Find(byte[] where, byte[] what) {
        for (int i = 0; i + what.Length <= where.Length; i++) {
            int k = 0;
            while (k < what.Length && where[i + k] == what[k]) k++;
            if (k == what.Length) return i;
        }
        return -1;
    }

    static FoxValue FieldValue(FoxContext c, int field) {
        DbfField f = c.Table.Fields[field];
        int at = c.Start + f.Offset;
        switch (f.Type) {
            case 'C': return FoxValue.Chars(Sub(c.Rec, at, f.Length));
            case 'N':
            case 'F': {
                string s = Encoding.ASCII.GetString(c.Rec, at, f.Length).Trim(' ', '\0');
                double d;
                if (s.Length == 0 || !double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) d = 0;
                return FoxValue.Number(d);
            }
            case 'D': {
                string s = Encoding.ASCII.GetString(c.Rec, at, Math.Min(8, f.Length)).Trim(' ', '\0');
                DateTime d;
                if (s.Length == 8 && DateTime.TryParseExact(s, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
                    return FoxValue.Date(Julian(d));
                return FoxValue.Date(0);
            }
            case 'L': return FoxValue.Logical("TtYy".IndexOf((char)c.Rec[at]) >= 0);
        }
        throw new NotSupportedException("поле " + f.Name + " типа " + f.Type);
    }

    public static double Julian(DateTime d) {
        return (d.Date - new DateTime(1, 1, 1)).Days + 1721426;
    }

    static string Dtos(double julian) {
        if (julian <= 0) return "        ";
        try { return new DateTime(1, 1, 1).AddDays(julian - 1721426).ToString("yyyyMMdd", CultureInfo.InvariantCulture); }
        catch (ArgumentOutOfRangeException) { return "        "; }
    }

    static readonly Dictionary<string, byte[]> caseMaps = new Dictionary<string, byte[]>();

    // таблица замены букв: UPPER/LOWER в кодировке таблицы; cyrillic = false - только латиница
    static byte[] CaseMap(Encoding enc, bool upper, bool cyrillic) {
        string key = enc.CodePage + (upper ? "U" : "L") + (cyrillic ? "R" : "");
        lock (caseMaps) {
            byte[] map;
            if (caseMaps.TryGetValue(key, out map)) return map;
            map = new byte[256];
            byte[] one = new byte[1];
            for (int b = 0; b < 256; b++) {
                map[b] = (byte)b;
                if (b < 128) {
                    if (upper && b >= 'a' && b <= 'z') map[b] = (byte)(b - 32);
                    if (!upper && b >= 'A' && b <= 'Z') map[b] = (byte)(b + 32);
                    continue;
                }
                if (!cyrillic) continue;
                one[0] = (byte)b;
                string s = enc.GetString(one);
                string t = upper ? s.ToUpperInvariant() : s.ToLowerInvariant();
                if (t == s) continue;
                byte[] tb = enc.GetBytes(t);
                if (tb.Length == 1 && enc.GetString(tb) == t) map[b] = tb[0];
            }
            caseMaps[key] = map;
            return map;
        }
    }

    FoxValue Call(FoxContext c) {
        FoxValue[] a = new FoxValue[Args.Length];
        if (Name != "IIF") for (int i = 0; i < a.Length; i++) a[i] = Args[i].Eval(c);
        switch (Name) {
            case "UPPER":
            case "LOWER": {
                byte[] s = Str(a[0]), map = CaseMap(c.Table.Enc, Name == "UPPER", c.CyrillicUpper);
                byte[] r = new byte[s.Length];
                for (int i = 0; i < s.Length; i++) r[i] = map[s[i]];
                return FoxValue.Chars(r);
            }
            case "STR": {
                double v = Num(a[0]);
                int len = a.Length > 1 ? Int(a[1]) : 10, dec = a.Length > 2 ? Int(a[2]) : 0;
                if (len <= 0) len = 10;
                string s = null;
                for (int d = Math.Max(dec, 0); d >= 0; d--) {
                    s = Math.Round(v, d, MidpointRounding.AwayFromZero).ToString("F" + d, CultureInfo.InvariantCulture);
                    if (s == "-0" || (s.StartsWith("-0.") && s.Trim('-', '0', '.').Length == 0)) s = s.Substring(1);
                    if (s.Length <= len) break;
                }
                if (s.Length > len) s = new string('*', len);
                return FoxValue.Chars(Encoding.ASCII.GetBytes(s.PadLeft(len)));
            }
            case "DTOS": return FoxValue.Chars(Encoding.ASCII.GetBytes(Dtos(Num(a[0]))));
            case "DTOC":
                if (a.Length > 1 && Int(a[1]) == 1) return FoxValue.Chars(Encoding.ASCII.GetBytes(Dtos(Num(a[0]))));
                throw new NotSupportedException("DTOC без ,1 зависит от настроек даты");
            case "LEFT": return FoxValue.Chars(Sub(Str(a[0]), 0, Int(a[1])));
            case "RIGHT": {
                byte[] s = Str(a[0]);
                int n = Math.Min(Math.Max(Int(a[1]), 0), s.Length);
                return FoxValue.Chars(Sub(s, s.Length - n, n));
            }
            case "SUBSTR": {
                byte[] s = Str(a[0]);
                int from = Int(a[1]) - 1;
                return FoxValue.Chars(Sub(s, from, a.Length > 2 ? Int(a[2]) : s.Length));
            }
            case "ALLTRIM":
            case "TRIM":
            case "RTRIM":
            case "LTRIM": {
                byte[] s = Str(a[0]);
                int from = 0, end = s.Length;
                if (Name != "LTRIM") while (end > 0 && (s[end - 1] == ' ' || s[end - 1] == 0)) end--;
                if (Name == "ALLTRIM" || Name == "LTRIM") while (from < end && s[from] == ' ') from++;
                return FoxValue.Chars(Sub(s, from, end - from));
            }
            case "PADL":
            case "PADR": {
                byte[] s = Str(a[0]);
                int n = Math.Max(Int(a[1]), 0);
                byte fill = a.Length > 2 && Str(a[2]).Length > 0 ? Str(a[2])[0] : (byte)' ';
                if (s.Length >= n) return FoxValue.Chars(Sub(s, 0, n));
                byte[] r = new byte[n];
                for (int i = 0; i < n; i++) r[i] = fill;
                Buffer.BlockCopy(s, 0, r, Name == "PADL" ? n - s.Length : 0, s.Length);
                return FoxValue.Chars(r);
            }
            case "VAL": {
                string s = Encoding.ASCII.GetString(Str(a[0])).Trim();
                int end = 0;
                while (end < s.Length && (char.IsDigit(s[end]) || s[end] == '.' || (end == 0 && (s[end] == '-' || s[end] == '+')))) end++;
                double d;
                return FoxValue.Number(double.TryParse(s.Substring(0, end), NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0);
            }
            case "INT": return FoxValue.Number(Math.Truncate(Num(a[0])));
            case "SPACE": return FoxValue.Chars(Encoding.ASCII.GetBytes(new string(' ', Math.Max(Int(a[0]), 0))));
            case "EMPTY": {
                FoxValue v = a[0];
                if (v.Type == 'C') { foreach (byte b in v.Text) if (b != ' ' && b != 0) return FoxValue.Logical(false); return FoxValue.Logical(true); }
                if (v.Type == 'L') return FoxValue.Logical(!v.Log);
                return FoxValue.Logical(v.Num == 0);
            }
            case "DELETED": return FoxValue.Logical(c.Rec[c.Start] == (byte)'*');
            case "IIF": return Bool(Args[0].Eval(c)) ? Args[1].Eval(c) : Args[2].Eval(c);
        }
        throw new NotSupportedException("функция " + Name);
    }

    // Ключ индекса из значения: текст дополняется пробелами до длины ключа,
    // число и дата - 8 байт, упорядоченных так, чтобы сравнение байтов давало порядок чисел
    public static byte[] KeyBytes(FoxValue v, int keyLen) {
        byte[] k;
        switch (v.Type) {
            case 'C':
                k = new byte[keyLen];
                for (int i = 0; i < keyLen; i++) k[i] = i < v.Text.Length ? v.Text[i] : (byte)' ';
                return k;
            case 'L':
                return new byte[] { (byte)(v.Log ? 'T' : 'F') };
        }
        if (keyLen == 4) {
            uint u = (uint)((int)v.Num) + 0x80000000u;
            return new byte[] { (byte)(u >> 24), (byte)(u >> 16), (byte)(u >> 8), (byte)u };
        }
        double d = v.Num == 0 ? 0.0 : v.Num;   // без "минус нуля"
        byte[] le = BitConverter.GetBytes(d);
        if (!BitConverter.IsLittleEndian) Array.Reverse(le);
        k = new byte[8];
        for (int i = 0; i < 8; i++) k[i] = le[7 - i];
        if (d >= 0) k[0] ^= 0x80;
        else for (int i = 0; i < 8; i++) k[i] ^= 0xFF;
        return k;
    }

    // ---- разбор ----

    class Parser {
        readonly string s;
        readonly DbfTable t;
        public int pos;

        public Parser(string s, DbfTable t) { this.s = s; this.t = t; }

        public void Space() { while (pos < s.Length && char.IsWhiteSpace(s[pos])) pos++; }

        bool Word(string w) {
            Space();
            if (string.Compare(s, pos, w, 0, w.Length, StringComparison.OrdinalIgnoreCase) != 0) return false;
            int end = pos + w.Length;
            if (char.IsLetter(w[w.Length - 1]) && end < s.Length && (char.IsLetterOrDigit(s[end]) || s[end] == '_')) return false;
            pos = end;
            return true;
        }

        bool Sym(string w) {
            Space();
            if (string.CompareOrdinal(s, pos, w, 0, w.Length) != 0) return false;
            pos += w.Length;
            return true;
        }

        static FoxExpr Node(char op, params FoxExpr[] args) {
            FoxExpr e = new FoxExpr();
            e.Op = op;
            e.Args = args;
            return e;
        }

        public FoxExpr Or() {
            FoxExpr e = And();
            while (Word(".OR.") || Word("OR")) e = Node('|', e, And());
            return e;
        }

        FoxExpr And() {
            FoxExpr e = Not();
            while (Word(".AND.") || Word("AND")) e = Node('&', e, Not());
            return e;
        }

        FoxExpr Not() {
            if (Word(".NOT.") || Word("NOT")) return Node('!', Not());
            Space();
            if (pos < s.Length && s[pos] == '!' && !(pos + 1 < s.Length && s[pos + 1] == '=')) { pos++; return Node('!', Not()); }
            return Cmp();
        }

        FoxExpr Cmp() {
            FoxExpr e = Add();
            string[] ops = { "==", "<>", "!=", "<=", ">=", "=", "#", "<", ">", "$" };
            char[] codes = { 'e', '#', '#', 'a', 'b', '=', '#', '<', '>', '$' };
            for (int i = 0; i < ops.Length; i++) if (Sym(ops[i])) return Node(codes[i], e, Add());
            return e;
        }

        FoxExpr Add() {
            FoxExpr e = Unary();
            while (true) {
                if (Sym("+")) e = Node('+', e, Unary());
                else if (Sym("-") ) e = Node('-', e, Unary());
                else return e;
            }
        }

        FoxExpr Unary() {
            if (Sym("-")) return Node('m', Unary());
            return Primary();
        }

        FoxExpr Primary() {
            Space();
            if (pos >= s.Length) throw new FormatException("выражение оборвалось");
            char ch = s[pos];
            if (ch == '(') {
                pos++;
                FoxExpr e = Or();
                if (!Sym(")")) throw new FormatException("нет закрывающей скобки");
                return e;
            }
            if (ch == '\'' || ch == '"' || ch == '[') {
                char close = ch == '[' ? ']' : ch;
                int end = s.IndexOf(close, pos + 1);
                if (end < 0) throw new FormatException("строка без конца");
                FoxExpr e = new FoxExpr();
                e.Op = 'v';
                e.Value = FoxValue.Chars(t.Enc.GetBytes(s.Substring(pos + 1, end - pos - 1)));
                pos = end + 1;
                return e;
            }
            if (Word(".T.") || Word(".F.")) {
                FoxExpr e = new FoxExpr();
                e.Op = 'v';
                e.Value = FoxValue.Logical(char.ToUpperInvariant(s[pos - 2]) == 'T');
                return e;
            }
            if (char.IsDigit(ch) || (ch == '.' && pos + 1 < s.Length && char.IsDigit(s[pos + 1]))) {
                int start = pos;
                while (pos < s.Length && (char.IsDigit(s[pos]) || s[pos] == '.')) pos++;
                FoxExpr e = new FoxExpr();
                e.Op = 'v';
                e.Value = FoxValue.Number(double.Parse(s.Substring(start, pos - start), CultureInfo.InvariantCulture));
                return e;
            }
            if (!(char.IsLetter(ch) || ch == '_')) throw new FormatException("непонятный знак '" + ch + "'");
            string name = Ident();
            // псевдоним таблицы: ABONENT.NAB или ABONENT->NAB
            Space();
            if (pos + 1 < s.Length && s[pos] == '-' && s[pos + 1] == '>') { pos += 2; Space(); name = Ident(); }
            else if (pos + 1 < s.Length && s[pos] == '.' && char.IsLetter(s[pos + 1]) && t.FieldIndex(name) < 0) {
                int save = pos;
                pos++;
                string inner = Ident();
                if (pos < s.Length && s[pos] == '.') pos = save;   // это .AND. и подобное
                else name = inner;
            }
            Space();
            if (pos < s.Length && s[pos] == '(') {
                pos++;
                FoxExpr e = new FoxExpr();
                e.Op = 'c';
                e.Name = FunctionName(name);
                List<FoxExpr> args = new List<FoxExpr>();
                if (!Sym(")")) {
                    do args.Add(Or()); while (Sym(","));
                    if (!Sym(")")) throw new FormatException("нет закрывающей скобки у " + name);
                }
                e.Args = args.ToArray();
                return e;
            }
            int f = t.FieldIndex(name);
            if (f < 0) throw new FormatException("нет такого поля: " + name);
            FoxExpr fe = new FoxExpr();
            fe.Op = 'f';
            fe.Field = f;
            fe.Name = t.Fields[f].Name;
            return fe;
        }

        string Ident() {
            int start = pos;
            while (pos < s.Length && (char.IsLetterOrDigit(s[pos]) || s[pos] == '_')) pos++;
            if (pos == start) throw new FormatException("ожидалось имя");
            return s.Substring(start, pos - start).ToUpperInvariant();
        }

        // FoxPro понимает имена функций, сокращённые до 4 букв: SUBS, ALLT, UPPE
        static string FunctionName(string name) {
            foreach (string f in Functions) {
                if (f == name) return f;
                if (name.Length >= 4 && f.StartsWith(name)) return f;
            }
            throw new NotSupportedException("функция " + name);
        }
    }
}

// ---- индексы FoxPro: .CDX (несколько тегов в файле) и .IDX (один индекс в файле) ----

public class IndexTag {
    public string FilePath, FileName, Name;
    public string Expr = "", ForText;
    public int KeyLen, Options;
    public bool Compact, Structural, Descending;
    public long Root;
    public FoxExpr Key, ForExpr;
    public char KeyType = '?';
    public List<FoxExpr> Terms = new List<FoxExpr>();
    public List<string> TermFields = new List<string>();   // поле каждого слагаемого ключа (null - не одно поле)
    public string Problem;          // почему индекс не годится (null - годится)
    public int State;               // 0 - ещё не проверен, 1 - проверен и годится, -1 - не годится
    public string Check = "";
    public bool CyrillicUpper = true;
    public int MinKeys = -1;        // обычный .IDX: в узлах первые ключи (1) или последние (0); -1 - ещё не смотрели
    public string Raw = "";         // служебные байты заголовка (порядок сортировки и т.п.) - для отчёта
    public string Collation = "";   // порядок сортировки FoxPro (SET COLLATE): пусто или MACHINE - байты как есть
    public bool Collated;           // ключ - не буквы, а "веса" сортировки (RUSSIAN, GENERAL...)
    public string Learned = "";     // как изучены веса по этому индексу - для отчёта
    public readonly object Lock = new object();

    public string Label { get { return Name == null ? FileName : FileName + ":" + Name; } }
    public string FirstField { get { return TermFields.Count > 0 ? TermFields[0] : null; } }
}

public class IndexNode {
    public bool Leaf;
    public long Right = -1;
    public List<byte[]> Keys = new List<byte[]>();
    public List<long> Values = new List<long>();   // в листе - номер записи, иначе - где лежит следующий узел
}

// Чтение индексного файла кусками по 512 байт (узлами), без загрузки целиком
public class IndexReader : IDisposable {
    readonly FileStream fs;
    readonly Dictionary<long, byte[]> pages = new Dictionary<long, byte[]>();
    public int Reads;

    public IndexReader(string path) {
        fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.RandomAccess);
    }

    public long Length { get { return fs.Length; } }

    public byte[] Read(long offset, int size) {
        if (offset < 0 || offset + size > fs.Length) throw new InvalidDataException("ссылка за концом файла индекса");
        byte[] b = new byte[size];
        fs.Seek(offset, SeekOrigin.Begin);
        if (DbfTable.ReadUpTo(fs, b, size) < size) throw new InvalidDataException("файл индекса обрезан");
        Reads++;
        return b;
    }

    // узел индекса (512 байт); прочитанные узлы запоминаются до конца поиска
    byte[] Page(long offset) {
        byte[] b;
        if (pages.TryGetValue(offset, out b)) return b;
        if (offset <= 0) throw new InvalidDataException("неверная ссылка в индексе");
        b = Read(offset, 512);
        pages[offset] = b;
        return b;
    }

    public void Dispose() { fs.Close(); }

    static int LE16(byte[] b, int at) { return b[at] | (b[at + 1] << 8); }
    static long LE32(byte[] b, int at) { return (uint)(b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)); }
    static long BE32(byte[] b, int at) { return (uint)((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]); }

    static string CString(byte[] b, int at, int max, Encoding enc) {
        int end = at;
        while (end < b.Length && end < at + max && b[end] != 0) end++;
        return enc.GetString(b, at, end - at).Trim();
    }

    public IndexNode Node(IndexTag tag, long offset) {
        byte[] p = Page(offset);
        IndexNode node = new IndexNode();
        int attr = LE16(p, 0), n = LE16(p, 2);
        node.Leaf = (attr & 2) != 0;
        long right = LE32(p, 8);
        node.Right = right == 0xFFFFFFFFL || right == 0 ? -1 : right;
        int keyLen = tag.KeyLen;
        if (tag.Compact && node.Leaf) {
            // хвост ключа: у текста - пробелы, у чисел и у "весов" сортировки (RUSSIAN, GENERAL) - нулевые байты
            DecodeLeaf(p, n, keyLen, tag.KeyType == 'C' && !tag.Collated ? (byte)' ' : (byte)0, node);
            return node;
        }
        int size = keyLen + (tag.Compact ? 8 : 4);
        if (12 + n * size > 512) throw new InvalidDataException("узел индекса повреждён");
        for (int i = 0; i < n; i++) {
            int at = 12 + i * size;
            byte[] k = new byte[keyLen];
            Buffer.BlockCopy(p, at, k, 0, keyLen);
            node.Keys.Add(k);
            // компактный: номер записи и ссылка на узел; обычный .IDX: одно число (в листе - запись, иначе - узел)
            node.Values.Add(tag.Compact ? BE32(p, at + keyLen + 4) : BE32(p, at + keyLen));
        }
        return node;
    }

    // Лист компактного индекса: в начале - номера записей и счётчики сжатия, сами ключи - с конца узла.
    // Ключ = первые dup байт предыдущего ключа + свои байты + trail байт-заполнителей.
    static void DecodeLeaf(byte[] p, int n, int keyLen, byte trail, IndexNode node) {
        long recMask = LE32(p, 14);
        int dupMask = p[18], trlMask = p[19], recBits = p[20], dupBits = p[21], req = p[23];
        if (req < 1 || req > 8 || 24 + n * req > 512) throw new InvalidDataException("лист индекса повреждён");
        int pos = 512;
        byte[] prev = new byte[keyLen];
        for (int i = 0; i < n; i++) {
            ulong v = 0;
            for (int b = 0; b < req; b++) v |= (ulong)p[24 + i * req + b] << (8 * b);
            long rec = (long)(v & (ulong)recMask);
            int dup = (int)((v >> recBits) & (ulong)dupMask);
            int trl = (int)((v >> (recBits + dupBits)) & (ulong)trlMask);
            if (i == 0) dup = 0;
            int len = keyLen - dup - trl;
            if (len < 0 || dup > keyLen || pos - len < 24 + n * req) throw new InvalidDataException("лист индекса повреждён");
            byte[] k = new byte[keyLen];
            Buffer.BlockCopy(prev, 0, k, 0, dup);
            pos -= len;
            Buffer.BlockCopy(p, pos, k, dup, len);
            for (int j = keyLen - trl; j < keyLen; j++) k[j] = trail;
            node.Keys.Add(k);
            node.Values.Add(rec);
            prev = k;
        }
    }

    // ключ из индекса сравнивается с искомым только по длине искомого (так ищется и начало)
    public static int Compare(byte[] entry, byte[] key) {
        int n = Math.Min(entry.Length, key.Length);
        for (int i = 0; i < n; i++) if (entry[i] != key[i]) return entry[i] < key[i] ? -1 : 1;
        return key.Length > entry.Length ? -1 : 0;
    }

    // Все записи, у которых ключ начинается с key (или равен ему, если key полной длины)
    public List<int> Seek(IndexTag tag, byte[] key, int limit, out bool more) {
        more = false;
        List<int> found = new List<int>();
        if (!tag.Compact && tag.MinKeys < 0) DetectMinKeys(tag);
        IndexNode node = Node(tag, tag.Root);
        for (int depth = 0; !node.Leaf; depth++) {
            if (depth > 40 || node.Keys.Count == 0) return found;
            int pick;
            if (tag.MinKeys == 1) {
                // в узле первые ключи потомков: берём последнего потомка, чей первый ключ меньше искомого
                pick = 0;
                for (int i = 1; i < node.Keys.Count; i++) if (Compare(node.Keys[i], key) < 0) pick = i; else break;
            } else {
                // в узле последние ключи потомков: первый потомок, чей последний ключ не меньше искомого
                pick = node.Keys.Count - 1;
                for (int i = 0; i < node.Keys.Count; i++) if (Compare(node.Keys[i], key) >= 0) { pick = i; break; }
            }
            node = Node(tag, node.Values[pick]);
        }
        for (int guard = 0; guard < 1000000; guard++) {
            for (int i = 0; i < node.Keys.Count; i++) {
                int c = Compare(node.Keys[i], key);
                if (c < 0) continue;
                if (c > 0) return found;
                if (found.Count >= limit) { more = true; return found; }
                found.Add((int)node.Values[i]);
            }
            if (node.Right < 0) return found;
            node = Node(tag, node.Right);
        }
        return found;
    }

    // Самый левый лист - для обхода индекса по порядку
    public IndexNode First(IndexTag tag) {
        if (!tag.Compact && tag.MinKeys < 0) DetectMinKeys(tag);
        IndexNode node = Node(tag, tag.Root);
        for (int depth = 0; !node.Leaf && node.Values.Count > 0 && depth < 40; depth++) node = Node(tag, node.Values[0]);
        return node;
    }

    void DetectMinKeys(IndexTag tag) {
        tag.MinKeys = 0;
        IndexNode root = Node(tag, tag.Root);
        if (root.Leaf || root.Keys.Count == 0) return;
        IndexNode child = Node(tag, root.Values[0]);
        if (child.Keys.Count == 0) return;
        if (Compare(root.Keys[0], child.Keys[child.Keys.Count - 1]) == 0) return;
        if (Compare(root.Keys[0], child.Keys[0]) == 0) tag.MinKeys = 1;
    }

    // ---- заголовки: список тегов файла ----

    public static List<IndexTag> ReadTags(string path, Encoding enc) {
        return ReadTags(path, enc, string.Compare(Path.GetExtension(path), ".cdx", StringComparison.OrdinalIgnoreCase) == 0);
    }

    // cdx - файл со списком тегов (.CDX); иначе - один индекс (.IDX). Для временных файлов (.TMP) пробуем оба
    public static List<IndexTag> ReadTags(string path, Encoding enc, bool cdx) {
        List<IndexTag> tags = new List<IndexTag>();
        using (IndexReader r = new IndexReader(path)) {
            if (r.Length < 1024) throw new InvalidDataException("файл слишком короткий для индекса");
            byte[] h = r.Read(0, 1024);
            if (cdx) {
                // в начале .CDX - список тегов: такой же индекс, где ключ - имя тега, а "номер записи" - где лежит его заголовок
                IndexTag dir = new IndexTag();
                dir.Compact = true;
                dir.KeyType = 'C';
                dir.KeyLen = LE16(h, 12);
                dir.Root = LE32(h, 0);
                if (dir.KeyLen < 1 || dir.KeyLen > 64) throw new InvalidDataException("не похоже на индекс FoxPro");
                IndexNode node = r.First(dir);
                for (int guard = 0; guard < 1000; guard++) {
                    for (int i = 0; i < node.Keys.Count; i++) {
                        IndexTag t = ReadCompactHeader(r, node.Values[i], enc);
                        t.Name = enc.GetString(node.Keys[i]).Trim(' ', '\0');
                        t.FilePath = path;
                        t.FileName = Path.GetFileName(path);
                        tags.Add(t);
                    }
                    if (node.Right < 0) break;
                    node = r.Node(dir, node.Right);
                }
            } else if ((h[14] & 0x20) != 0) {
                IndexTag t = ReadCompactHeader(r, 0, enc);
                t.FilePath = path;
                t.FileName = Path.GetFileName(path);
                tags.Add(t);
            } else {
                // обычный (некомпактный) .IDX: заголовок 512 байт, выражение - с 16-го байта
                IndexTag t = new IndexTag();
                t.FilePath = path;
                t.FileName = Path.GetFileName(path);
                t.Root = LE32(h, 0);
                t.KeyLen = LE16(h, 12);
                t.Options = h[14];
                t.Expr = CString(h, 16, 220, enc);
                if ((t.Options & 8) != 0) t.ForText = CString(h, 236, 220, enc);
                tags.Add(t);
            }
        }
        return tags;
    }

    static IndexTag ReadCompactHeader(IndexReader r, long at, Encoding enc) {
        byte[] h = r.Read(at, 1024);
        IndexTag t = new IndexTag();
        t.Compact = true;
        StringBuilder raw = new StringBuilder("сигн. " + h[15].ToString("X2") + ", байты 480-511:");
        for (int i = 480; i < 512; i++) raw.Append(" " + h[i].ToString("X2"));
        t.Raw = raw.ToString();
        // имя порядка сортировки (RUSSIAN, GENERAL, MACHINE) - в байтах 494-501
        StringBuilder coll = new StringBuilder();
        for (int i = 494; i < 502 && h[i] >= (byte)'A' && h[i] <= (byte)'Z'; i++) coll.Append((char)h[i]);
        t.Collation = coll.ToString();
        t.Root = LE32(h, 0);
        t.KeyLen = LE16(h, 12);
        t.Options = h[14];
        t.Descending = LE16(h, 502) != 0;
        int forPos = LE16(h, 504), forLen = LE16(h, 506), keyPos = LE16(h, 508), keyLen = LE16(h, 510);
        if (keyPos == 0 && keyLen != 0 && forPos == 0 && forLen != 0) forPos = keyLen;
        if (keyPos < 512) t.Expr = CString(h, 512 + keyPos, keyLen == 0 ? 512 - keyPos : keyLen, enc);
        if ((t.Options & 8) != 0 && forPos < 512 && h[512 + forPos] != 0)
            t.ForText = CString(h, 512 + forPos, forLen == 0 ? 512 - forPos : forLen, enc);
        return t;
    }
}

// Итог одного поиска
public class SearchResult {
    public string Text;
    public SearchMode Mode;
    public DbfTable Table;
    public AbonentQuery Query;
    public List<int> Rows = new List<int>();
    public int Hidden;          // нашлись, но в невыбранных сельсоветах
    public bool NoIndex;        // нет подходящего индекса - можно только перебором
    public bool TooMany;        // нашлось слишком много - показана часть
    public bool Scanned;        // искали перебором всей базы
    public string Error;
    public bool Partial;        // частичное совпадение (иначе - только точное)
    public AddressMatch Address; // как понят адрес (точный поиск адреса)
    public string NameText, AddressText;   // поиск по ФИО с адресом ("Иванов Боровляны"): что ФИО, что адрес
    public string Note = "";    // пояснение для строки состояния
    public string How = "";     // как искали - для отчёта
    public double Seconds;
    public Dictionary<int, int> Tariff = new Dictionary<int, int>();      // строка -> номер тарифа (T1, T2...)
    public Dictionary<int, string> Power = new Dictionary<int, string>(); // строка -> мощность (доп. сведения)
    public Dictionary<int, bool> Added = new Dictionary<int, bool>();     // строки, которые не искали: другие тарифы того же счётчика
    public Dictionary<int, string> Group = new Dictionary<int, string>(); // строка -> её группа (рамка): тарифы одного счётчика, абоненты одного адреса
    public Dictionary<int, string> Related = new Dictionary<int, string>(); // добавленные по адресу или ФИО -> почему
    public Dictionary<int, bool> Orange = new Dictionary<int, bool>();      // на адресе больше одного счётчика - оранжевым

    public string GroupOf(int row) {
        string g;
        return Group.TryGetValue(row, out g) ? g : null;
    }

    public int Searched { get { return Rows.Count - Added.Count; } }

    public string TariffText(int row) {
        int n;
        return Tariff.TryGetValue(row, out n) ? "T" + n : "";
    }

    public string PowerText(int row) {
        string v;
        return Power.TryGetValue(row, out v) ? v : "";
    }
}

public delegate bool CancelCheck();

// ---- сведения из соседних таблиц ----
// Тарифы: у многотарифного счётчика на каждый тариф (T1, T2, T3) свой лицевой; связка лежит в ABDIF_T.DBF:
// основной лицевой (MAINNAB), лицевой тарифа (NAB) и номер тарифа.
// Дополнительные сведения (в программе базы - Ctrl+D): DOP.DBF / DOGOVOR.DBF, оттуда же мощность.
public class ExtraInfo {
    // тарифы
    string tariffPath;
    DateTime tariffTime, tariffChecked;
    long tariffSize;
    readonly Dictionary<string, int> tariffOf = new Dictionary<string, int>();   // лицевой -> номер тарифа
    readonly Dictionary<string, string> mainOf = new Dictionary<string, string>(); // лицевой -> основной лицевой
    readonly Dictionary<string, List<KeyValuePair<int, string>>> groups = new Dictionary<string, List<KeyValuePair<int, string>>>(); // основной -> (тариф, лицевой)
    int mainField = -1, nabField = -1, numField = -1;
    readonly object sync = new object();
    public string TariffNote = "";

    // дополнительные сведения
    public readonly List<LiveBase> Tables = new List<LiveBase>();
    public LiveBase PowerBase;
    public int PowerField = -1;
    public string PowerNote = "";

    public static string Nab(string s) {
        string t = s.Trim();
        string r = t.TrimStart('0');
        return r.Length == 0 && t.Length > 0 ? "0" : r;
    }

    public static ExtraInfo Open(string folder) {
        ExtraInfo x = new ExtraInfo();
        string tariff = Find(folder, "ABDIF_T.DBF");
        if (tariff != null) {
            x.tariffPath = tariff;
            try { x.LoadTariffs(true); }
            catch (Exception ex) { x.TariffNote = "таблица тарифов не прочиталась: " + ex.Message; x.tariffPath = null; }
        } else x.TariffNote = "таблицы тарифов ABDIF_T.DBF нет";
        Report.Line("Тарифы (T1, T2, T3): " + x.TariffNote);

        foreach (string name in new string[] { "DOP.DBF", "DOGOVOR.DBF" }) {
            string path = Find(folder, name);
            if (path == null) { Report.Line("Доп. сведения: таблицы " + name + " нет"); continue; }
            try {
                LiveBase b = LiveBase.Open(path);
                if (b.Header.FieldIndex("NAB") < 0) { Report.Line("Доп. сведения: в " + name + " нет поля NAB - не связать с абонентом"); continue; }
                x.Tables.Add(b);
                Report.Line("Доп. сведения " + name + ": записей " + b.Header.Expected + ", индексы: "
                    + (b.Tags.Count == 0 ? "нет" : TagList(b)) + ". Поля:");
                Report.Fields(b.Sample, "    ");
            } catch (Exception ex) {
                Report.Line("Доп. сведения: " + name + " не открылась - " + ex.Message);
            }
        }
        x.FindPower();
        Report.Line("Мощность: " + x.PowerNote);
        return x;
    }

    static string TagList(LiveBase b) {
        List<string> l = new List<string>();
        foreach (IndexTag t in b.Tags) l.Add(t.Name + " (" + t.Expr + ")");
        return string.Join(", ", l.ToArray());
    }

    static string Find(string folder, string name) {
        foreach (string f in Directory.GetFiles(folder))
            if (string.Compare(Path.GetFileName(f), name, StringComparison.OrdinalIgnoreCase) == 0) return f;
        return null;
    }

    // поле мощности: по имени (MOSH..., MOCH..., POW...), число с одним знаком после запятой
    void FindPower() {
        int best = 0;
        foreach (LiveBase b in Tables) {
            for (int i = 0; i < b.Header.Fields.Count; i++) {
                DbfField f = b.Header.Fields[i];
                if (f.Type != 'N' && f.Type != 'F') continue;
                string n = f.Name.ToUpperInvariant();
                int score = 0;
                if (n.Contains("MOSH") || n.Contains("MOSC") || n.Contains("MOCH") || n.Contains("MOSN") || n.Contains("MOWN") || n.Contains("MOW")) score = 10;
                else if (n.StartsWith("MOS") || n.StartsWith("MOЩ") || n.StartsWith("MO")) score = 7;
                else if (n.Contains("POW") || n.Contains("PMAX") || n.Contains("PUST") || n.Contains("PRAZ")) score = 6;
                if (f.Decimals == 1) score += 3;
                else if (f.Decimals > 0) score += 1;
                if (score > best) { best = score; PowerBase = b; PowerField = i; }
            }
        }
        if (PowerBase != null && best >= 6) {
            PowerNote = "поле " + Path.GetFileName(PowerBase.FilePath) + "." + PowerBase.Header.Fields[PowerField].Name;
            return;
        }
        // по имени не нашлось: единственное число с одним знаком после запятой в DOP
        PowerBase = null;
        PowerField = -1;
        foreach (LiveBase b in Tables) {
            int found = -1, count = 0;
            for (int i = 0; i < b.Header.Fields.Count; i++)
                if ((b.Header.Fields[i].Type == 'N' || b.Header.Fields[i].Type == 'F') && b.Header.Fields[i].Decimals == 1) { found = i; count++; }
            if (count == 1) { PowerBase = b; PowerField = found; PowerNote = "поле " + Path.GetFileName(b.FilePath) + "." + b.Header.Fields[found].Name + " (угадано: единственное число вида 0.0)"; return; }
        }
        PowerNote = "поле мощности не найдено - пришлите отчёт (в нём список полей таблиц доп. сведений)";
    }

    // ---- тарифы ----

    void LoadTariffs(bool first) {
        FileInfo fi = new FileInfo(tariffPath);
        DbfTable t = DbfTable.Load(tariffPath, false);
        mainField = t.FieldIndex("MAINNAB");
        nabField = t.FieldIndex("NAB");
        if (first) {
            numField = -1;
            // номер тарифа - числовое поле, где только 1..9
            for (int j = 0; j < t.Fields.Count && numField < 0; j++) {
                if (j == mainField || j == nabField) continue;
                DbfField f = t.Fields[j];
                if (f.Type != 'N' && !(f.Type == 'C' && f.Length <= 2)) continue;
                int filled = 0, ok = 0;
                for (int r = 0; r < t.Count && r < 2000; r++) {
                    string v = DbfTable.Format(t.GetValue(r, j, null), true).Trim();
                    if (v.Length == 0) continue;
                    filled++;
                    int n;
                    if (int.TryParse(v, out n) && n >= 1 && n <= 9) ok++;
                }
                if (filled > 0 && ok == filled) numField = j;
            }
        }
        Dictionary<string, int> tariff = new Dictionary<string, int>();
        Dictionary<string, string> main = new Dictionary<string, string>();
        Dictionary<string, List<KeyValuePair<int, string>>> group = new Dictionary<string, List<KeyValuePair<int, string>>>();
        if (nabField >= 0 && numField >= 0) {
            for (int r = 0; r < t.Count; r++) {
                string raw = DbfTable.Format(t.GetValue(r, nabField, null), true).Trim();
                string nab = Nab(raw);
                int n;
                if (nab.Length == 0 || !int.TryParse(DbfTable.Format(t.GetValue(r, numField, null), true).Trim(), out n)) continue;
                tariff[nab] = n;
                string m = mainField >= 0 ? Nab(DbfTable.Format(t.GetValue(r, mainField, null), true)) : "";
                if (m.Length == 0) continue;
                main[nab] = m;
                List<KeyValuePair<int, string>> list;
                if (!group.TryGetValue(m, out list)) { list = new List<KeyValuePair<int, string>>(); group[m] = list; }
                list.Add(new KeyValuePair<int, string>(n, raw));
            }
        }
        foreach (List<KeyValuePair<int, string>> list in group.Values)
            list.Sort(delegate(KeyValuePair<int, string> a, KeyValuePair<int, string> b) { return a.Key.CompareTo(b.Key); });
        lock (sync) {
            tariffOf.Clear();
            mainOf.Clear();
            groups.Clear();
            foreach (KeyValuePair<string, int> kv in tariff) tariffOf[kv.Key] = kv.Value;
            foreach (KeyValuePair<string, string> kv in main) mainOf[kv.Key] = kv.Value;
            foreach (KeyValuePair<string, List<KeyValuePair<int, string>>> kv in group) groups[kv.Key] = kv.Value;
            tariffTime = fi.LastWriteTime;
            tariffSize = fi.Length;
            tariffChecked = DateTime.Now;
        }
        if (first) {
            TariffNote = Path.GetFileName(tariffPath) + ", записей " + t.Count + ", лицевых с тарифом " + tariff.Count
                + "; поля: основной - " + (mainField < 0 ? "нет" : t.Fields[mainField].Name) + ", лицевой - " + (nabField < 0 ? "нет" : t.Fields[nabField].Name)
                + ", номер тарифа - " + (numField < 0 ? "не найден" : t.Fields[numField].Name);
            Report.Line("Поля таблицы тарифов:");
            Report.Fields(t, "    ");
        }
    }

    // не реже раза в минуту смотрим, не поменялась ли таблица тарифов (данные - свежие)
    void FreshTariffs() {
        if (tariffPath == null) return;
        bool check;
        lock (sync) check = DateTime.Now - tariffChecked > TimeSpan.FromSeconds(60);
        if (!check) return;
        try {
            FileInfo fi = new FileInfo(tariffPath);
            bool changed;
            lock (sync) { changed = fi.LastWriteTime != tariffTime || fi.Length != tariffSize; tariffChecked = DateTime.Now; }
            if (changed) LoadTariffs(false);
        } catch (Exception) { }
    }

    // номер тарифа лицевого (0 - не тарифный) и основной лицевой
    public int Tariff(string nab, out string main) {
        FreshTariffs();
        string k = Nab(nab);
        lock (sync) {
            int n;
            if (!tariffOf.TryGetValue(k, out n)) { main = null; return 0; }
            if (!mainOf.TryGetValue(k, out main)) main = null;
            return n;
        }
    }

    // все лицевые одного многотарифного счётчика по основному лицевому: (номер тарифа, лицевой), по порядку тарифов
    public List<KeyValuePair<int, string>> Group(string main) {
        lock (sync) {
            List<KeyValuePair<int, string>> list;
            return main != null && groups.TryGetValue(Nab(main), out list) ? new List<KeyValuePair<int, string>>(list) : new List<KeyValuePair<int, string>>();
        }
    }

    // ---- мощность и карточка ----

    public Dictionary<string, string> Power(List<string> nabs) {
        Dictionary<string, string> result = new Dictionary<string, string>();
        if (PowerBase == null || nabs.Count == 0) return result;
        DbfTable t;
        Dictionary<string, int> rows = PowerBase.LookupNabs(nabs, out t);
        if (t == null) return result;
        foreach (KeyValuePair<string, int> kv in rows) {
            int start;
            byte[] buf = t.RecordBuffer(kv.Value, out start);
            DbfField f = t.Fields[PowerField];
            result[kv.Key] = Encoding.ASCII.GetString(buf, start + f.Offset, f.Length).Trim();
        }
        return result;
    }

    // для карточки: все поля из таблиц доп. сведений (пары "подпись, значение")
    // запись абонента в каждой таблице доп. сведений (DOP, DOGOVOR)
    public List<ExtraRecord> Records(string nab) {
        List<ExtraRecord> list = new List<ExtraRecord>();
        foreach (LiveBase b in Tables) {
            ExtraRecord e = new ExtraRecord();
            e.Name = Path.GetFileName(b.FilePath);
            list.Add(e);
            DbfTable t;
            Dictionary<string, int> rows;
            try { rows = b.LookupNabs(new List<string>(new string[] { nab }), out t); }
            catch (Exception ex) { e.Problem = "не прочиталось: " + ex.Message; continue; }
            int row;
            if (t == null) { e.Problem = "нет индекса по номеру абонента"; continue; }
            if (!rows.TryGetValue(Nab(nab), out row)) { e.Problem = "записи нет"; continue; }
            e.Table = t;
            using (MemoReader memo = t.OpenMemo()) e.Values = t.GetRow(row, memo);
        }
        return list;
    }

    public List<string> Details(string nab) {
        List<string> pairs = new List<string>();
        foreach (ExtraRecord e in Records(nab)) {
            pairs.Add("--- Доп. сведения (" + e.Name + ")");
            if (e.Problem != null) { pairs.Add(e.Problem); continue; }
            pairs.Add("");
            for (int j = 0; j < e.Values.Length; j++) {
                if (string.Compare(e.Table.Fields[j].Name, "NAB", StringComparison.OrdinalIgnoreCase) == 0) continue;
                pairs.Add(FieldNames.Title(e.Table.Fields[j].Name));
                pairs.Add(DbfTable.Format(e.Values[j], true));
            }
        }
        return pairs;
    }
}

// Запись абонента в таблице доп. сведений
public class ExtraRecord {
    public string Name;        // таблица (DOP.DBF)
    public DbfTable Table;
    public object[] Values;
    public string Problem;     // не прочиталось, записи нет...
}

// Свой индекс по ФИО: файл indeks_fio.dat в папке программы (индекса по ФИО у таблицы абонентов нет).
// В нём только ФИО (маленькими буквами) и номер записи - не копия базы. Строится один раз: таблица читается
// в фоне, программа в это время работает; прервали (закрыли программу) - продолжится с того же места. Дальше:
//  - новые абоненты (FoxPro дописывает их в конец таблицы) дочитываются перед каждым поиском - секунды;
//  - удалённые пропускаются: найденные записи всё равно читаются из базы;
//  - исправленные ФИО подхватываются тихой пересборкой раз в сутки (старый индекс в это время работает).
public class NameIndex {
    const string Magic = "POISK-FIO-1";
    const int BlockShift = 22, BlockSize = 1 << BlockShift;   // ФИО лежат подряд в блоках по 4 МБ
    readonly List<byte[]> blocks = new List<byte[]>();
    readonly List<int> blockLength = new List<int>();
    int[] start = new int[1 << 16];     // где ФИО записи: номер блока и место в нём
    byte[] length = new byte[1 << 16];
    int[] recNo = new int[1 << 16];
    public int Count;                   // сколько ФИО в индексе
    public int Scanned;                 // сколько записей таблицы просмотрено
    public int Total;                   // записей в таблице, когда смотрели последний раз
    public DateTime Built = DateTime.Now;   // когда начата сборка (раз в сутки - заново)
    public double Seconds;
    readonly int field, fieldLength, headerLength, recordLength, codePage;
    readonly string table;
    readonly Encoding enc;
    readonly byte[] fold = new byte[256];   // маленькие буквы, ё -> е

    public NameIndex(DbfTable h, string tablePath) {
        int f = h.FieldIndex("FIO");
        field = h.Fields[f].Offset;
        fieldLength = h.Fields[f].Length;
        headerLength = h.HeaderLength;
        recordLength = h.RecordLength;
        enc = h.Enc;
        codePage = enc.CodePage;
        table = Path.GetFileName(tablePath).ToUpperInvariant();
        byte[] one = new byte[1];
        for (int b = 0; b < 256; b++) {
            fold[b] = (byte)b;
            one[0] = (byte)b;
            string s = enc.GetString(one);
            string l = s.ToLowerInvariant().Replace('ё', 'е');
            if (l == s) continue;
            byte[] lb = enc.GetBytes(l);
            if (lb.Length == 1 && enc.GetString(lb) == l) fold[b] = lb[0];
        }
    }

    public static int RecordCount(FileStream fs) {
        byte[] h = new byte[8];
        fs.Seek(0, SeekOrigin.Begin);
        if (DbfTable.ReadUpTo(fs, h, 8) < 8) return 0;
        return BitConverter.ToInt32(h, 4);
    }

    // дочитать таблицу с Scanned до конца; progress может прервать (исключением);
    // file - сохранять туда по ходу (каждые 100 тысяч записей), чтобы прерванная сборка продолжилась
    public void Continue(string dbfPath, LoadProgress progress, string file) {
        DateTime started = DateTime.Now;
        try {
            using (FileStream fs = new FileStream(dbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16, FileOptions.SequentialScan)) {
                Total = RecordCount(fs);
                Read(fs, Scanned, Total, progress, file);
            }
        } finally {
            Seconds += (DateTime.Now - started).TotalSeconds;
        }
    }

    // перед поиском: дочитать новые записи; false - таблица стала короче (её пересобрали), индекс устарел
    public bool Update(FileStream fs, out int added) {
        int count = RecordCount(fs);
        added = 0;
        if (count < Scanned) return false;
        Total = count;
        if (count > Scanned) { int was = Count; Read(fs, Scanned, count, null, null); added = Count - was; }
        return true;
    }

    void Read(FileStream fs, int from, int count, LoadProgress progress, string file) {
        const int chunkRecords = 4096;
        byte[] buf = new byte[chunkRecords * recordLength];
        fs.Seek(headerLength + (long)from * recordLength, SeekOrigin.Begin);
        int index = from, saved = from;
        while (index < count) {
            int want = Math.Min(chunkRecords, count - index);
            int got = DbfTable.ReadUpTo(fs, buf, want * recordLength) / recordLength;
            bool end = false;
            for (int r = 0; r < got; r++) {
                int p = r * recordLength;
                if (buf[p] == 0x1A) { got = r; end = true; break; }   // конец файла
                if (buf[p] != (byte)'*') Add(buf, p + field, index + r + 1);
            }
            index += got;
            Scanned = index;
            if (end || got < want) break;
            if (file != null && index - saved >= 100000) { TrySave(file); saved = index; }
            if (progress != null) progress(index, count);
        }
    }

    void Add(byte[] src, int from, int rec) {
        int n = fieldLength;
        while (n > 0 && (src[from + n - 1] == (byte)' ' || src[from + n - 1] == 0)) n--;
        if (n > 255) n = 255;
        if (blocks.Count == 0 || blockLength[blocks.Count - 1] + n > BlockSize) { blocks.Add(new byte[BlockSize]); blockLength.Add(0); }
        int bi = blocks.Count - 1, used = blockLength[bi];
        byte[] block = blocks[bi];
        for (int k = 0; k < n; k++) block[used + k] = fold[src[from + k]];
        if (Count == start.Length) {
            int size = Math.Max(16, Count * 2);
            Array.Resize(ref start, size);
            Array.Resize(ref length, size);
            Array.Resize(ref recNo, size);
        }
        start[Count] = (bi << BlockShift) | used;
        length[Count] = (byte)n;
        recNo[Count] = rec;
        Count++;
        blockLength[bi] = used + n;
    }

    // номера записей, где ФИО подходит: точно (с фамилии, словами целиком) или частично (все слова - части ФИО)
    public List<int> Find(List<string> words, bool partial) {
        List<int> found = new List<int>();
        if (words.Count == 0) return found;
        byte[] first = enc.GetBytes(words[0]);
        for (int k = 0; k < first.Length; k++) first[k] = fold[first[k]];
        for (int i = 0; i < Count; i++) {
            byte[] block = blocks[start[i] >> BlockShift];
            int s = start[i] & (BlockSize - 1), n = length[i];
            if (!Contains(block, s, n, first)) continue;
            string fio = enc.GetString(block, s, n);
            if (partial ? AbonentQuery.ContainsAll(fio, words) : AbonentQuery.NameMatches(fio, words)) found.Add(recNo[i]);
        }
        return found;
    }

    static bool Contains(byte[] data, int from, int n, byte[] what) {
        int last = from + n - what.Length;
        for (int p = from; p <= last; p++) {
            if (data[p] != what[0]) continue;
            int k = 1;
            while (k < what.Length && data[p + k] == what[k]) k++;
            if (k == what.Length) return true;
        }
        return false;
    }

    public string Describe() {
        return "свой индекс ФИО: " + Count + " ФИО, просмотрено записей " + Scanned + " из " + Total + ", сборка начата " + Built.ToString("dd.MM.yyyy HH:mm");
    }

    void TrySave(string file) {
        try { Save(file); } catch (Exception) { }
    }

    // в файл: сначала во временный, потом заменить - прерванная запись не портит прежний индекс
    public void Save(string file) {
        string tmp = file + ".tmp";
        using (FileStream fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16))
        using (BinaryWriter w = new BinaryWriter(fs, Encoding.UTF8)) {
            w.Write(Magic);
            w.Write(table);
            w.Write(headerLength);
            w.Write(recordLength);
            w.Write(field);
            w.Write(fieldLength);
            w.Write(codePage);
            w.Write(Scanned);
            w.Write(Total);
            w.Write(Built.Ticks);
            w.Write(Count);
            w.Write(blocks.Count);
            for (int i = 0; i < blocks.Count; i++) { w.Write(blockLength[i]); w.Write(blocks[i], 0, blockLength[i]); }
            WriteInts(w, start, Count);
            w.Write(length, 0, Count);
            WriteInts(w, recNo, Count);
        }
        if (File.Exists(file)) File.Delete(file);
        File.Move(tmp, file);
    }

    // из файла; null - файл от другой таблицы или другой структуры (тогда строим заново)
    public static NameIndex Load(string file, DbfTable h, string tablePath) {
        using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16))
        using (BinaryReader r = new BinaryReader(fs, Encoding.UTF8)) {
            if (r.ReadString() != Magic) return null;
            NameIndex n = new NameIndex(h, tablePath);
            if (r.ReadString() != n.table || r.ReadInt32() != n.headerLength || r.ReadInt32() != n.recordLength
                || r.ReadInt32() != n.field || r.ReadInt32() != n.fieldLength || r.ReadInt32() != n.codePage) return null;
            n.Scanned = r.ReadInt32();
            n.Total = r.ReadInt32();
            n.Built = new DateTime(r.ReadInt64());
            int count = r.ReadInt32();
            int nb = r.ReadInt32();
            for (int i = 0; i < nb; i++) {
                int len = r.ReadInt32();
                byte[] b = new byte[BlockSize];
                if (r.Read(b, 0, len) != len) return null;
                n.blocks.Add(b);
                n.blockLength.Add(len);
            }
            n.start = ReadInts(r, count);
            n.length = r.ReadBytes(count);
            n.recNo = ReadInts(r, count);
            if (n.start == null || n.recNo == null || n.length.Length != count) return null;
            n.Count = count;
            return n;
        }
    }

    static void WriteInts(BinaryWriter w, int[] a, int n) {
        byte[] b = new byte[n * 4];
        Buffer.BlockCopy(a, 0, b, 0, n * 4);
        w.Write(b);
    }

    static int[] ReadInts(BinaryReader r, int n) {
        byte[] b = r.ReadBytes(n * 4);
        if (b.Length != n * 4) return null;
        int[] a = new int[Math.Max(16, n)];
        Buffer.BlockCopy(b, 0, a, 0, n * 4);
        return a;
    }
}

// База абонентов "вживую": при запуске читаются только заголовок, начало таблицы и список индексов.
// Каждый поиск идёт по индексу FoxPro прямо в файлах базы - данные всегда свежие.
public class LiveBase {
    public string FilePath;
    public DbfTable Header;      // поля и кодировка
    public DbfTable Sample;      // первые записи - чтобы понять, как записаны номера и ФИО
    public StreetDirectory Dir;
    public List<IndexTag> Tags = new List<IndexTag>();
    public List<string> IndexFiles = new List<string>();
    public int RecordCount;
    public ExtraInfo Extra;      // тарифы и доп. сведения из соседних таблиц

    public const int FetchLimit = 300;    // больше записей за один поиск не показываем (адрес - до 5000: потом отбор по дому)
    const int SeekLimit = 5000;
    const int NameByAddressLimit = 30000;   // ФИО с адресом: сколько абонентов этого места читаем, чтобы найти среди них ФИО

    static readonly string[] AccountFields = { "NAB", "LS", "LIC", "LSCH", "LS_ERIP" };

    public static LiveBase Open(string path) {
        LiveBase b = new LiveBase();
        b.FilePath = path;
        b.Header = DbfTable.LoadHeader(path);
        b.RecordCount = b.Header.Expected;
        b.Sample = DbfTable.Load(path, null, null, 500);
        b.FindIndexes();
        return b;
    }

    // Индексы рядом с таблицей: свой .CDX (с тем же именем, FoxPro открывает его сам) и прочие .CDX/.IDX,
    // если их выражения состоят из полей этой таблицы
    void FindIndexes() {
        string folder = Path.GetDirectoryName(FilePath);
        string own = Path.GetFileNameWithoutExtension(FilePath);
        List<string> files = new List<string>();
        Dictionary<string, bool> tables = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        string[] all = Directory.GetFiles(folder);
        foreach (string f in all) if (string.Compare(Path.GetExtension(f), ".dbf", StringComparison.OrdinalIgnoreCase) == 0) tables[Path.GetFileNameWithoutExtension(f)] = true;
        foreach (string f in all) {
            string ext = Path.GetExtension(f).ToLowerInvariant();
            if (ext != ".cdx" && ext != ".idx") continue;
            IndexFiles.Add(Path.GetFileName(f));
            // .CDX другой таблицы (ARC13.CDX у ARC13.DBF) - её собственный индекс, не наш
            string name = Path.GetFileNameWithoutExtension(f);
            if (ext == ".cdx" && string.Compare(name, own, StringComparison.OrdinalIgnoreCase) != 0 && tables.ContainsKey(name)) continue;
            files.Add(f);
        }
        files.Sort(delegate(string a, string b) {
            bool oa = string.Compare(Path.GetFileNameWithoutExtension(a), own, StringComparison.OrdinalIgnoreCase) == 0;
            bool ob = string.Compare(Path.GetFileNameWithoutExtension(b), own, StringComparison.OrdinalIgnoreCase) == 0;
            if (oa != ob) return oa ? -1 : 1;
            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        });
        foreach (string f in files) {
            bool structural = string.Compare(Path.GetFileNameWithoutExtension(f), own, StringComparison.OrdinalIgnoreCase) == 0
                && string.Compare(Path.GetExtension(f), ".cdx", StringComparison.OrdinalIgnoreCase) == 0;
            List<IndexTag> tags;
            try {
                tags = IndexReader.ReadTags(f, Header.Enc);
            } catch (Exception ex) {
                if (structural) Report.Line("    " + Path.GetFileName(f) + ": не прочитался - " + ex.Message);
                continue;
            }
            foreach (IndexTag t in tags) {
                t.Structural = structural;
                Prepare(t);
                // чужой индекс (другой таблицы) - выражение не из наших полей; такие не показываем
                if (!structural && t.Key == null) continue;
                Tags.Add(t);
            }
        }
    }

    void Prepare(IndexTag t) {
        try {
            t.Key = FoxExpr.Parse(t.Expr, Header);
        } catch (Exception ex) {
            t.Problem = "выражение не разобрать (" + ex.Message + ")";
            return;
        }
        if (t.ForText != null && t.ForText.Length > 0) {
            try { t.ForExpr = FoxExpr.Parse(t.ForText, Header); }
            catch (Exception ex) { t.Problem = "условие FOR не разобрать (" + ex.Message + ")"; return; }
        }
        if ((t.Options & 1) != 0) { t.Problem = "UNIQUE - в индексе только первая запись с каждым ключом"; return; }
        try {
            byte[] blank = BlankRecord();
            FoxContext c = Context(blank, t);
            t.KeyType = t.Key.Eval(c).Type;
            if (t.ForExpr != null && t.ForExpr.Eval(c).Type != 'L') { t.Problem = "условие FOR - не да/нет"; return; }
        } catch (Exception ex) {
            t.Problem = "выражение не посчитать (" + ex.Message + ")";
            return;
        }
        t.Collated = t.KeyType == 'C' && t.Collation.Length > 0 && t.Collation != "MACHINE";
        if ((t.KeyType == 'N' || t.KeyType == 'D') && t.KeyLen != 8 && t.KeyLen != 4) { t.Problem = "непонятная длина числового ключа " + t.KeyLen; return; }
        if (t.KeyLen < 1 || t.KeyLen > 240) { t.Problem = "непонятная длина ключа " + t.KeyLen; return; }
        t.Terms = t.KeyType == 'C' ? t.Key.Terms() : new List<FoxExpr>(new FoxExpr[] { t.Key });
        foreach (FoxExpr term in t.Terms) {
            List<int> fields = new List<int>();
            term.CollectFields(fields);
            t.TermFields.Add(fields.Count == 1 ? Header.Fields[fields[0]].Name.ToUpperInvariant() : null);
        }
    }

    public byte[] BlankRecord() {
        byte[] r = new byte[Header.RecordLength];
        for (int i = 0; i < r.Length; i++) r[i] = (byte)' ';
        return r;
    }

    FoxContext Context(byte[] rec, IndexTag t) {
        FoxContext c = new FoxContext();
        c.Table = Header;
        c.Rec = rec;
        c.CyrillicUpper = t == null || t.CyrillicUpper;
        return c;
    }

    // ---- чтение записей по номерам ----

    public FileStream OpenTable() {
        return new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.RandomAccess);
    }

    public int CurrentCount(FileStream fs) {
        byte[] h = new byte[8];
        fs.Seek(0, SeekOrigin.Begin);
        if (DbfTable.ReadUpTo(fs, h, 8) == 8) RecordCount = BitConverter.ToInt32(h, 4);
        return RecordCount;
    }

    public byte[] ReadRecord(FileStream fs, int recNo) {
        long at = Header.HeaderLength + (long)(recNo - 1) * Header.RecordLength;
        if (recNo < 1 || at + Header.RecordLength > fs.Length) return null;
        byte[] r = new byte[Header.RecordLength];
        fs.Seek(at, SeekOrigin.Begin);
        return DbfTable.ReadUpTo(fs, r, r.Length) == r.Length ? r : null;
    }

    // Записи из разных мест таблицы (по одной через равные промежутки)
    public DbfTable SpreadSample(int n) {
        DbfTable part = DbfTable.CreateLike(Header);
        part.FilePath = FilePath;
        using (FileStream fs = OpenTable()) {
            int count = CurrentCount(fs);
            for (int i = 1; i <= n && count > 0; i++) {
                int recNo = (int)Math.Max(1, (long)count * i / (n + 1));
                byte[] rec = ReadRecord(fs, recNo);
                if (rec != null && rec[0] != (byte)'*') part.AddRecord(rec, 0, recNo);
            }
        }
        return part;
    }

    // Записи по номерам; соседние читаются одним куском. Удалённые пропускаются.
    public DbfTable Fetch(FileStream fs, List<int> recNos, int limit, out int deleted) {
        deleted = 0;
        DbfTable part = DbfTable.CreateLike(Header);
        part.FilePath = FilePath;
        List<int> list = new List<int>(recNos);
        list.Sort();
        int len = Header.RecordLength;
        byte[] buf = new byte[65536 + len];
        int i = 0;
        while (i < list.Count && part.Count < limit) {
            int first = list[i], j = i;
            while (j + 1 < list.Count && (long)(list[j + 1] - first + 1) * len <= 65536 && list[j + 1] - list[j] <= 16) j++;
            int last = list[j];
            fs.Seek(Header.HeaderLength + (long)(first - 1) * len, SeekOrigin.Begin);
            int got = DbfTable.ReadUpTo(fs, buf, (last - first + 1) * len);
            for (int k = i; k <= j && part.Count < limit; k++) {
                if (k > i && list[k] == list[k - 1]) continue;
                int pos = (list[k] - first) * len;
                if (pos + len > got) continue;
                if (buf[pos] == (byte)'*') { deleted++; continue; }
                part.AddRecord(buf, pos, list[k]);
            }
            i = j + 1;
        }
        return part;
    }

    // ---- проверка индекса: ключи в нём такие же, как считает программа? ----

    // Берём записи из разных мест таблицы, считаем их ключ и ищем его в индексе - запись должна найтись.
    // Так заодно видно, что индекс от этой таблицы и что он не устарел.
    public bool Verify(IndexTag t) {
        lock (t.Lock) {
            if (t.State != 0) return t.State > 0;
            if (t.Problem != null) { t.State = -1; t.Check = t.Problem; return false; }
            if (t.Collated) {
                Learn(t);
                if (t.Problem != null) { t.State = -1; t.Check = t.Problem + "; " + t.Learned; return false; }
            }
            bool[] variants = t.Key.UsesCase() && !t.Collated ? new bool[] { true, false } : new bool[] { true };
            string last = "";
            foreach (bool cyr in variants) {
                t.CyrillicUpper = cyr;
                string note;
                int result = TryVerify(t, out note);
                last = note;
                if (result > 0) {
                    t.State = 1;
                    t.Check = "годится: " + note + (t.Key.UsesCase() ? (cyr ? " (UPPER - и русские буквы)" : " (UPPER - только латиница)") : "");
                    return true;
                }
                if (result == 0) { t.State = 0; t.Check = "не проверен: " + note; t.CyrillicUpper = true; return false; }
            }
            t.CyrillicUpper = true;
            t.State = -1;
            t.Check = "не годится: " + last;
            return false;
        }
    }

    // 1 - годится, -1 - нет, 0 - не удалось проверить (ошибка чтения)
    int TryVerify(IndexTag t, out string note) {
        int tried = 0, ok = 0;
        string sample = "";
        try {
            using (FileStream fs = OpenTable())
            using (IndexReader r = new IndexReader(t.FilePath)) {
                int count = CurrentCount(fs);
                if (count == 0) { note = "таблица пуста"; return 0; }
                for (int step = 1; step <= 40 && tried < 8; step++) {
                    int recNo = (int)Math.Max(1, Math.Min(count, (long)count * step / 41));
                    byte[] rec = ReadRecord(fs, recNo);
                    if (rec == null || rec[0] == (byte)'*') continue;
                    FoxContext c = Context(rec, t);
                    if (t.ForExpr != null && !t.ForExpr.Eval(c).Log) continue;
                    byte[] key = t.Collated ? CollatedKey(t, t.Key.Eval(c).Text, true) : FoxExpr.KeyBytes(t.Key.Eval(c), t.KeyLen);
                    if (key == null || (t.KeyType == 'C' && IsBlank(key))) continue;   // пустой ключ есть у многих - не показательно
                    bool more;
                    List<int> found = r.Seek(t, key, SeekLimit, out more);
                    if (more && !found.Contains(recNo)) continue;      // слишком много одинаковых - не показательно
                    tried++;
                    if (found.Contains(recNo)) ok++;
                    else if (sample.Length == 0) sample = "; ключ записи: " + (t.Collated ? "длина " + key.Length : Report.Shape(Header.Enc.GetString(key), false))
                        + ", нашлось по нему записей: " + found.Count;
                }
                note = "нашлись " + ok + " из " + tried + " проверочных записей (чтений индекса " + r.Reads + ")" + sample;
            }
        } catch (Exception ex) {
            note = "ошибка при проверке: " + ex.Message;
            return 0;
        }
        if (tried < 2) { note = "мало записей для проверки (" + tried + ")"; return ok == tried && tried > 0 ? 1 : 0; }
        return ok >= tried * 0.75 ? 1 : -1;
    }

    // ---- порядок сортировки RUSSIAN / GENERAL ----
    // При SET COLLATE TO "RUSSIAN" (или GENERAL) FoxPro кладёт в индекс не сами символы, а их "веса":
    // по байту на символ (большие и маленькие буквы - одинаково), затем добавки для букв с особыми знаками,
    // остаток ключа - нулевые байты. Ключ вдвое длиннее поля. Таблица весов RUSSIAN нигде не опубликована,
    // поэтому программа узнаёт её по самому индексу: берёт ключи из нескольких мест индекса, читает их записи
    // и сопоставляет символ за символом.

    // веса GENERAL для латиницы, цифр и знаков (для RUSSIAN они, как правило, те же - проверяется по индексу)
    static readonly int[] GeneralAscii = {
        16,16,16,16,16,16,16,16,16,17,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,16,
        17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,86,87,88,89,90,91,92,93,94,95,33,34,35,36,37,38,
        39,96,97,98,100,102,103,104,105,106,107,108,109,111,112,114,115,116,117,118,119,120,122,123,124,125,126,40,41,42,43,44,
        45,96,97,98,100,102,103,104,105,106,107,108,109,111,112,114,115,116,117,118,119,120,122,123,124,125,126,46,47,48,49,16 };

    class Weights {
        public int[] Head = new int[256];
        public bool LikeGeneral = true;   // все изученные латинские буквы, цифры и знаки совпали с GENERAL
        public int Known;
        public Weights() { for (int i = 0; i < 256; i++) Head[i] = -1; }
    }

    readonly Dictionary<string, Weights> weights = new Dictionary<string, Weights>();

    Weights WeightsOf(string collation) {
        lock (weights) {
            Weights w;
            if (!weights.TryGetValue(collation, out w)) { w = new Weights(); weights[collation] = w; }
            return w;
        }
    }

    // вес символа: изученный; для латиницы и цифр - как в GENERAL, если изученное с ним не расходится; иначе -1
    int Head(Weights w, byte c) {
        int h = w.Head[c];
        if (h >= 0) return h;
        return c < 128 && w.LikeGeneral ? GeneralAscii[c] : -1;
    }

    // Начало ключа по значению: веса символов до первого неизвестного (тогда ищется по более короткому началу)
    byte[] CollatedKey(IndexTag t, byte[] raw, bool trimEnd) {
        if (raw == null) return null;
        Weights w = WeightsOf(t.Collation);
        int n = raw.Length;
        if (trimEnd) while (n > 0 && raw[n - 1] == (byte)' ') n--;
        List<byte> key = new List<byte>();
        lock (w) {
            for (int i = 0; i < n && key.Count < t.KeyLen; i++) {
                int h = Head(w, raw[i]);
                if (h < 0) break;
                key.Add((byte)h);
            }
        }
        return key.Count == 0 ? null : key.ToArray();
    }

    IndexNode Descend(IndexReader r, IndexTag t, bool last) {
        IndexNode node = r.Node(t, t.Root);
        for (int depth = 0; !node.Leaf && node.Values.Count > 0 && depth < 40; depth++)
            node = r.Node(t, node.Values[last ? node.Values.Count - 1 : node.Values.Count / 2]);
        return node;
    }

    void Learn(IndexTag t) {
        if (t.Learned.Length > 0) return;
        Weights w = WeightsOf(t.Collation);
        int samples = 0, shorter = 0, conflicts = 0, learned = 0;
        try {
            using (FileStream fs = OpenTable())
            using (IndexReader r = new IndexReader(t.FilePath)) {
                List<IndexNode> leaves = new List<IndexNode>();
                leaves.Add(r.First(t));
                leaves.Add(Descend(r, t, false));
                leaves.Add(Descend(r, t, true));
                foreach (IndexNode leaf in leaves) {
                    for (int i = 0, used = 0; i < leaf.Keys.Count && used < 25; i++) {
                        byte[] rec = ReadRecord(fs, (int)leaf.Values[i]);
                        if (rec == null) continue;
                        FoxValue v = t.Key.Eval(Context(rec, t));
                        if (v.Type != 'C') continue;
                        int n = v.Text.Length;
                        while (n > 0 && v.Text[n - 1] == (byte)' ') n--;
                        if (n == 0) continue;
                        byte[] k = leaf.Keys[i];
                        int m = k.Length;
                        while (m > 0 && k[m - 1] == 0) m--;
                        used++;
                        samples++;
                        if (m < n) { shorter++; continue; }
                        bool conflict = false;
                        lock (w) {
                            for (int j = 0; j < n; j++) {
                                byte c = v.Text[j];
                                if (w.Head[c] < 0) { w.Head[c] = k[j]; learned++; w.Known++; }
                                else if (w.Head[c] != k[j]) conflict = true;
                            }
                        }
                        if (conflict) conflicts++;
                    }
                }
            }
        } catch (Exception ex) {
            t.Learned = "веса не изучены: " + ex.Message;
            t.Problem = "не удалось изучить сортировку " + t.Collation;
            return;
        }
        lock (w) {
            for (int c = 0; c < 128; c++) if (w.Head[c] >= 0 && w.Head[c] != GeneralAscii[c]) w.LikeGeneral = false;
        }
        t.Learned = "сортировка " + t.Collation + ": изучено по " + samples + " ключам, новых символов " + learned
            + (shorter > 0 ? ", ключ короче значения: " + shorter : "") + (conflicts > 0 ? ", противоречий: " + conflicts : "");
        if (samples >= 5 && shorter + conflicts > samples / 5)
            t.Problem = "ключи устроены не так, как ожидалось (сортировка " + t.Collation + ")";
    }

    // для отчёта: какие веса узнаны (цифры, пробел, сколько русских букв)
    public string WeightsReport() {
        StringBuilder sb = new StringBuilder();
        lock (weights) {
            foreach (KeyValuePair<string, Weights> kv in weights) {
                Weights w = kv.Value;
                StringBuilder digits = new StringBuilder();
                for (int c = '0'; c <= '9'; c++) digits.Append((char)c + "=" + (w.Head[c] < 0 ? "?" : w.Head[c].ToString()) + " ");
                int cyr = 0;
                for (int c = 128; c < 256; c++) if (w.Head[c] >= 0) cyr++;
                sb.Append("Сортировка " + kv.Key + ": узнано символов " + w.Known + "; цифры: " + digits + "пробел=" + (w.Head[32] < 0 ? "?" : w.Head[32].ToString())
                    + "; русских и прочих букв: " + cyr + "; латиница и цифры как в GENERAL: " + (w.LikeGeneral ? "да" : "нет") + ". ");
            }
        }
        return sb.ToString();
    }

    static bool IsBlank(byte[] k) {
        foreach (byte b in k) if (b != ' ' && b != 0) return false;
        return true;
    }

    // Подходящий проверенный индекс, ключ которого начинается с поля field
    public IndexTag TagFor(string field) {
        return TagFor(field, false);
    }

    // preferLong - лучше индекс из нескольких полей (улица + дом: сразу находится нужный дом)
    public IndexTag TagFor(string field, bool preferLong) {
        List<IndexTag> list = new List<IndexTag>();
        foreach (IndexTag t in Tags) if (t.Problem == null && string.Compare(t.FirstField, field, StringComparison.OrdinalIgnoreCase) == 0) list.Add(t);
        list.Sort(delegate(IndexTag a, IndexTag b) {
            if (a.Structural != b.Structural) return a.Structural ? -1 : 1;
            if ((a.ForExpr == null) != (b.ForExpr == null)) return a.ForExpr == null ? -1 : 1;
            if (preferLong && a.Terms.Count != b.Terms.Count) return b.Terms.Count.CompareTo(a.Terms.Count);
            return a.Expr.Length.CompareTo(b.Expr.Length);
        });
        foreach (IndexTag t in list) if (Verify(t)) return t;
        return null;
    }

    // проверить все индексы по нужным полям сразу после запуска (итог - в отчёт)
    public void VerifyMain() {
        foreach (string f in new string[] { "NAB", "LS_ERIP", "NCH", "NSC", "FIO" }) TagFor(f);
        TagFor("UL", true);
    }

    // ---- ключи для поиска ----

    void Put(byte[] rec, string fieldName, string value) {
        int i = Header.FieldIndex(fieldName);
        if (i < 0) throw new ArgumentException(fieldName);
        DbfField f = Header.Fields[i];
        byte[] b;
        switch (f.Type) {
            case 'N':
            case 'F': {
                double d;
                if (!double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) throw new FormatException("не число");
                string s = d.ToString("F" + f.Decimals, CultureInfo.InvariantCulture);
                if (s.Length > f.Length) throw new FormatException("не помещается");
                b = Encoding.ASCII.GetBytes(s.PadLeft(f.Length));
                break;
            }
            default:
                b = Header.Enc.GetBytes(value);
                break;
        }
        for (int k = 0; k < f.Length; k++) rec[f.Offset + k] = k < b.Length ? b[k] : (byte)' ';
    }

    // Ключ (или его начало) по значениям первых полей ключа. prefixLast - последнее значение может быть началом поля.
    public byte[] KeyFor(IndexTag t, string[] values, bool prefixLast) {
        if (values.Length == 0 || values.Length > t.Terms.Count) return null;
        byte[] rec = BlankRecord();
        try {
            for (int i = 0; i < values.Length; i++) {
                if (t.TermFields[i] == null) return null;
                Put(rec, t.TermFields[i], values[i]);
            }
            FoxContext c = Context(rec, t);
            if (t.KeyType != 'C') {
                if (prefixLast || values.Length != 1) return null;
                return FoxExpr.KeyBytes(t.Key.Eval(c), t.KeyLen);
            }
            if (t.Collated) {
                // веса сортировки: значения подряд, без хвостовых пробелов; ищется всегда начало ключа,
                // точное совпадение потом сверяется по самим записям
                List<byte> raw = new List<byte>();
                for (int i = 0; i < values.Length; i++) {
                    FoxValue v = t.Terms[i].Eval(c);
                    if (v.Type != 'C' || (prefixLast && i == values.Length - 1 && !t.Terms[i].PrefixSafe())) return null;
                    raw.AddRange(v.Text);
                }
                return CollatedKey(t, raw.ToArray(), true);
            }
            List<byte> key = new List<byte>();
            for (int i = 0; i < values.Length; i++) {
                FoxValue v = t.Terms[i].Eval(c);
                if (v.Type != 'C') return null;
                int len = v.Text.Length;
                if (i == values.Length - 1 && prefixLast) {
                    if (!t.Terms[i].PrefixSafe()) return null;
                    while (len > 0 && v.Text[len - 1] == ' ') len--;
                }
                for (int k = 0; k < len; k++) key.Add(v.Text[k]);
            }
            if (key.Count == 0) return null;
            if (key.Count > t.KeyLen) key.RemoveRange(t.KeyLen, key.Count - t.KeyLen);
            return key.ToArray();
        } catch (FormatException) {
            return null;
        }
    }

    // ---- поиск ----

    class Seeker {
        public readonly Dictionary<int, bool> Found = new Dictionary<int, bool>();
        public bool More;
        public int Seeks, Reads;
        public readonly List<string> Tags = new List<string>();
        readonly Dictionary<string, IndexReader> readers = new Dictionary<string, IndexReader>();

        public int Seek(IndexTag t, byte[] key) {
            if (key == null) return 0;
            IndexReader r;
            if (!readers.TryGetValue(t.FilePath, out r)) { r = new IndexReader(t.FilePath); readers[t.FilePath] = r; }
            if (!Tags.Contains(t.Label)) Tags.Add(t.Label);
            Seeks++;
            int before = r.Reads;
            bool more;
            List<int> list = r.Seek(t, key, SeekLimit, out more);
            Reads += r.Reads - before;
            if (more) More = true;
            int added = 0;
            foreach (int n in list) if (!Found.ContainsKey(n)) { Found[n] = true; added++; }
            return list.Count;
        }

        public void Close() { foreach (IndexReader r in readers.Values) r.Dispose(); }
    }

    public SearchResult Search(string text, SearchMode mode, Dictionary<string, bool> selsovety, bool partial) {
        return Search(text, mode, selsovety, FetchLimit, partial);
    }

    // ---- свой индекс по ФИО (в базе индекса по ФИО нет) - см. NameIndex ----
    NameIndex names;                       // готовый - по нему ищем
    readonly object namesLock = new object();
    public string NamesFile;               // файл индекса (в папке программы)

    public bool NamesNeeded { get { return Header.FieldIndex("FIO") >= 0 && TagFor("FIO") == null; } }
    public bool NamesReady { get { return names != null; } }

    // Файл индекса при запуске. Возвращает индекс, который надо (до)строить в фоне, или null - готов, строить не надо
    public NameIndex OpenNames() {
        NameIndex n = null;
        try { if (NamesFile != null && File.Exists(NamesFile)) n = NameIndex.Load(NamesFile, Header, FilePath); }
        catch (Exception ex) { Report.Line("Свой индекс ФИО не прочитался (" + ex.Message + ") - строю заново"); n = null; }
        if (n == null) { Report.Line("Свой индекс ФИО: ещё нет - строю в фоне (" + NamesFile + ")"); return new NameIndex(Header, FilePath); }
        int count;
        using (FileStream fs = OpenTable()) count = NameIndex.RecordCount(fs);
        if (count < n.Scanned) { Report.Line("Свой индекс ФИО: таблицу пересобрали (записей стало меньше) - строю заново"); return new NameIndex(Header, FilePath); }
        n.Total = count;
        if (count - n.Scanned > 20000) { Report.Line("Свой индекс ФИО: недостроен (" + n.Scanned + " из " + count + ") - достраиваю в фоне"); return n; }
        lock (namesLock) names = n;
        Report.Line("Свой индекс ФИО: " + n.Describe());
        if ((DateTime.Now - n.Built).TotalHours >= 20) {
            Report.Line("Свой индекс ФИО собран больше суток назад - пересобираю в фоне (чтобы подхватить исправленные ФИО), пока ищем по старому");
            return new NameIndex(Header, FilePath);
        }
        return null;
    }

    // (до)строить в фоне; готовый становится рабочим и сохраняется. savePartial - сохранять по ходу
    // (первая сборка: прервали - продолжится; при пересборке - нет, чтобы не испортить рабочий файл)
    public void BuildNames(NameIndex n, LoadProgress progress, bool savePartial) {
        n.Continue(FilePath, progress, savePartial ? NamesFile : null);
        lock (namesLock) {
            names = n;
            try { if (NamesFile != null) n.Save(NamesFile); }
            catch (Exception ex) { Report.Line("Свой индекс ФИО не сохранился: " + ex.Message); }
        }
        Report.Line(DateTime.Now.ToString("HH:mm:ss") + "  Свой индекс ФИО построен за " + n.Seconds.ToString("0") + " с: " + n.Describe());
    }

    // ФИО по своему индексу: номера записей (перед этим дочитываются новые записи таблицы); null - индекс устарел
    List<int> FindNames(FileStream fs, string name, bool partial, out string how) {
        List<string> words = partial ? new List<string>(DbfTable.SplitWords(name.Replace('ё', 'е').Replace('Ё', 'Е'))) : AbonentQuery.PersonWords(name);
        lock (namesLock) {
            int added;
            if (!names.Update(fs, out added)) { names = null; how = "свой индекс ФИО устарел (таблицу пересобрали)"; return null; }
            List<int> found = names.Find(words, partial);
            how = names.Describe() + (added > 0 ? ", дочитано новых записей: " + added : "") + "; совпало " + found.Count;
            return found;
        }
    }

    // ФИО (и, если есть, адрес "Иванов Боровляны") из найденных по своему индексу; null - так не получилось
    SearchResult SearchNames(SearchResult res, string text, bool withAddress, Dictionary<string, bool> selsovety, int limit, DateTime started) {
        string name = text, addr = null;
        if (withAddress) {
            if (!SplitNameAddress(text, out name, out addr)) return null;
            AddressMatch m = Dir.Match(addr);
            res.Address = m;
            if (m.Problem != null) { res.Note = m.Problem; return Finish(res, null, selsovety, started); }
        }
        using (FileStream fs = OpenTable()) {
            string how;
            List<int> recNos = FindNames(fs, name, res.Partial, out how);
            if (recNos == null) { res.NoIndex = true; res.How = how; return Finish(res, null, selsovety, started); }
            // с адресом, а по ФИО слишком много (частая фамилия) - лучше взять абонентов этого места по индексу адреса
            if (addr != null && recNos.Count > limit) return null;
            if (recNos.Count > limit) res.TooMany = true;
            int deleted;
            DbfTable part = Fetch(fs, recNos, limit, out deleted);
            res.NameText = name;
            if (addr != null) res.AddressText = addr;
            res.How = how + ", прочитано записей " + part.Count + (deleted > 0 ? ", удалённых пропущено " + deleted : "");
            return Finish(res, part, selsovety, started);
        }
    }

    // Ошибка чтения бывает, если программа базы как раз дописывает индекс - тогда пробуем ещё раз.
    // partial - частичное совпадение (начало номера, похожие слова); иначе - только точное
    public SearchResult Search(string text, SearchMode mode, Dictionary<string, bool> selsovety, int limit, bool partial) {
        SearchResult r = SearchOnce(text, mode, selsovety, limit, partial, false);
        if (r.Error != null) {
            System.Threading.Thread.Sleep(300);
            SearchResult again = SearchOnce(text, mode, selsovety, limit, partial, false);
            again.How += "; (со второй попытки, первая: " + r.Error + ")";
            return again;
        }
        // по индексу ФИО целиком не нашлось - может, в конце адрес ("Иванов Боровляны")
        if (mode == SearchMode.Name && r.Rows.Count == 0 && r.Hidden == 0 && r.AddressText == null && !r.NoIndex && (TagFor("FIO") != null || names != null)) {
            SearchResult a = SearchOnce(text, mode, selsovety, limit, partial, true);
            if (a.Error == null && (a.Rows.Count > 0 || a.Hidden > 0 || a.NameText != null)) return a;
        }
        return r;
    }

    // withAddress - ФИО с адресом ("Иванов Боровляны"), даже если есть индекс по ФИО
    SearchResult SearchOnce(string text, SearchMode mode, Dictionary<string, bool> selsovety, int limit, bool partial, bool withAddress) {
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Partial = partial;
        if (mode == SearchMode.Name && names != null && TagFor("FIO") == null) {
            try {
                SearchResult r = SearchNames(res, text, withAddress, selsovety, limit, started);
                if (r != null) return r;
                res = new SearchResult();
                res.Text = text;
                res.Mode = mode;
                res.Partial = partial;
            } catch (Exception ex) {
                res.Error = ex.Message;
                res.How += " ОШИБКА: " + ex.GetType().Name + ": " + ex.Message;
                return Finish(res, null, selsovety, started);
            }
        }
        Seeker s = new Seeker();
        try {
            bool ok;
            switch (mode) {
                case SearchMode.Account: ok = SeekAccount(text, s, res, partial); break;
                case SearchMode.Meter: ok = SeekMeter(text, s, res, partial); break;
                case SearchMode.Name:
                    if (withAddress) ok = SeekNameByAddress(text, s, res, partial);
                    else {
                        ok = SeekName(text, s, res, partial);
                        if (!ok && res.Note.Length == 0) ok = SeekNameByAddress(text, s, res, partial);
                    }
                    break;
                default: ok = partial || Dir == null ? SeekAddress(text, s, res) : SeekAddressExact(text, s, res); break;
            }
            if (!ok) { res.NoIndex = res.Error == null && res.Note.Length == 0; return Finish(res, null, selsovety, started); }
            List<int> recNos = new List<int>(s.Found.Keys);
            if (mode == SearchMode.Address) limit = Math.Max(limit, SeekLimit);
            if (res.NameText != null) limit = Math.Max(limit, NameByAddressLimit);
            if (recNos.Count > limit || s.More) res.TooMany = true;
            int deleted;
            DbfTable part;
            using (FileStream fs = OpenTable()) part = Fetch(fs, recNos, limit, out deleted);
            res.How = "индекс " + string.Join(", ", s.Tags.ToArray()) + ": поисков " + s.Seeks + ", чтений индекса " + s.Reads
                + ", ключей нашлось " + recNos.Count + (s.More ? "+" : "") + ", прочитано записей " + part.Count
                + (deleted > 0 ? ", удалённых пропущено " + deleted : "") + (res.How.Length > 0 ? "; " + res.How : "");
            return Finish(res, part, selsovety, started);
        } catch (Exception ex) {
            res.Error = ex.Message;
            res.How += " ОШИБКА: " + ex.GetType().Name + ": " + ex.Message;
            return Finish(res, null, selsovety, started);
        } finally {
            s.Close();
        }
    }

    // Перебор всей таблицы - если нужного индекса нет (долго, но находит всё)
    public SearchResult Scan(string text, SearchMode mode, Dictionary<string, bool> selsovety, LoadProgress progress, bool partial) {
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Scanned = true;
        res.Partial = partial;
        DbfTable part;
        if (mode == SearchMode.Address && Dir != null && Header.FieldIndex("UL") >= 0) {
            Dictionary<string, bool> codes = new Dictionary<string, bool>();
            if (partial) {
                int best;
                foreach (string c in Dir.BestStreets(AddressParts.Parse(text).Words, null, out best)) codes[c] = true;
            } else {
                res.Address = Dir.Match(text);
                if (res.Address.Problem != null) { res.Note = res.Address.Problem; return Finish(res, null, selsovety, started); }
                foreach (string c in res.Address.Codes) codes[c] = true;
            }
            DbfField ul = Header.Fields[Header.FieldIndex("UL")];
            part = DbfTable.Load(FilePath, null, delegate(byte[] rec, int start) {
                return codes.ContainsKey(StreetDirectory.NormalizeCode(Encoding.ASCII.GetString(rec, start + ul.Offset, ul.Length)));
            }, progress, 20000);
        } else {
            // предварительный отбор: запись должна содержать номер (или фамилию); точность проверяется потом
            string words = text;
            if (mode == SearchMode.Account) { words = Digits(text).TrimStart('0'); if (words.Length == 0) words = Digits(text); }
            else if (mode == SearchMode.Meter) words = LongestDigits(text).TrimStart('0');
            else if (mode == SearchMode.Name && !partial) { string[] w = DbfTable.SplitWords(text.Replace('.', ' ')); words = w.Length > 0 ? w[0] : ""; }
            part = DbfTable.Load(FilePath, words, null, progress, 20000);
        }
        res.TooMany = part.Truncated;
        double mb = part.BytesRead / 1048576.0;
        res.How = "перебор всей таблицы: " + mb.ToString("0") + " МБ за " + part.Seconds.ToString("0") + " с, записей " + part.Total + ", подошло " + part.Count;
        return Finish(res, part, selsovety, started);
    }

    SearchResult Finish(SearchResult res, DbfTable part, Dictionary<string, bool> selsovety, DateTime started) {
        if (part == null) part = DbfTable.CreateLike(Header);
        part.FilePath = FilePath;
        res.Table = part;
        res.Query = new AbonentQuery(part, Dir);
        res.Query.Partial = res.Partial;
        res.Query.Address = res.Address;
        if (part.Count > 0 && res.Note.Length == 0) {
            List<int> rows = res.Query.Find(res.NameText ?? res.Text, res.Mode);
            // ФИО с адресом: из найденных по ФИО - только по этому адресу (дом, квартира - если вписаны)
            if (res.NameText != null && res.AddressText != null) {
                Dictionary<int, bool> here = new Dictionary<int, bool>();
                foreach (int row in res.Query.Find(res.AddressText, SearchMode.Address)) here[row] = true;
                rows = rows.FindAll(delegate(int row) { return here.ContainsKey(row); });
            }
            int ul = part.FieldIndex("UL");
            foreach (int row in rows) {
                if (Dir != null && ul >= 0 && selsovety != null
                    && !Dir.InSelsovety(DbfTable.Format(part.GetValue(row, ul, null), true), selsovety)) { res.Hidden++; continue; }
                res.Rows.Add(row);
            }
            if (res.Rows.Count > 0) AddExtras(res);
        }
        res.Seconds = (DateTime.Now - started).TotalSeconds;
        return res;
    }

    // Тариф (T1, T2...) и мощность для найденных; лицевые одного счётчика - рядом, по порядку тарифов
    // Дополнительный поиск к найденному: другие тарифы того же счётчика (ABDIF_T), другие счётчики на том же
    // адресе и абоненты с тем же ФИО в том же доме. Найденное и добавленное собирается в группы (рамка).
    void AddExtras(SearchResult res) {
        int nab = res.Table.FieldIndex("NAB");
        if (nab < 0) return;
        Dictionary<int, string> mainOf = new Dictionary<int, string>();
        if (Extra != null) {
            foreach (int row in res.Rows) {
                string v = DbfTable.Format(res.Table.GetValue(row, nab, null), true);
                string main;
                int t = Extra.Tariff(v, out main);
                if (t > 0) { res.Tariff[row] = t; mainOf[row] = main == null ? ExtraInfo.Nab(v) : main; }
            }
            if (res.Tariff.Count > 0) AddSiblings(res, nab, mainOf);
        }
        try { AddRelated(res, nab, mainOf); }
        catch (Exception ex) { res.How += "; поиск по адресу/ФИО не удался: " + ex.Message; }
        Arrange(res, mainOf);
        if (Extra == null) return;
        List<string> nabs = new List<string>();
        foreach (int row in res.Rows) if (nabs.Count < 100) nabs.Add(DbfTable.Format(res.Table.GetValue(row, nab, null), true));
        try {
            Dictionary<string, string> power = Extra.Power(nabs);
            foreach (int row in res.Rows) {
                string p;
                if (power.TryGetValue(ExtraInfo.Nab(DbfTable.Format(res.Table.GetValue(row, nab, null), true)), out p)) res.Power[row] = p;
            }
        } catch (Exception ex) {
            res.How += "; мощность не прочиталась: " + ex.Message;
        }
    }

    string Field(DbfTable t, int row, int field) {
        return field < 0 ? "" : DbfTable.Format(t.GetValue(row, field, null), true).Trim();
    }

    // адрес целиком (улица, дом, корпус, квартира) - для сравнения; пусто, если дома нет
    string AddrKey(DbfTable t, int row) {
        string dom = AddressMatch.Num(Field(t, row, t.FieldIndex("DOM")));
        if (dom.Length == 0) return "";
        return StreetDirectory.NormalizeCode(Field(t, row, t.FieldIndex("UL"))) + "|" + dom + "|"
            + Part(Field(t, row, t.FieldIndex("KOR"))) + "|" + Part(Field(t, row, t.FieldIndex("KV")));
    }

    // корпус/квартира "0" - то же, что пусто (частный дом)
    static string Part(string v) {
        string n = AddressMatch.Num(v);
        return n == "0" ? "" : n;
    }

    string HouseKey(DbfTable t, int row) {
        string dom = AddressMatch.Num(Field(t, row, t.FieldIndex("DOM")));
        return dom.Length == 0 ? "" : StreetDirectory.NormalizeCode(Field(t, row, t.FieldIndex("UL"))) + "|" + dom;
    }

    static string FioKey(string fio) {
        StringBuilder sb = new StringBuilder();
        foreach (string w in fio.ToLowerInvariant().Replace('ё', 'е').Split(new char[] { ' ', '.', ',', '\t' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (sb.Length > 0) sb.Append(' ');
            sb.Append(w);
        }
        return sb.ToString();
    }

    // Другие абоненты того же дома: с тем же адресом (второй счётчик) или с тем же ФИО. Дом читается по индексу
    // улица+дом; ФИО ищется только в доме найденного (индекса по ФИО в базе нет).
    void AddRelated(SearchResult res, int nab, Dictionary<int, string> mainOf) {
        DbfTable tb = res.Table;
        int ulF = tb.FieldIndex("UL"), domF = tb.FieldIndex("DOM"), fioF = tb.FieldIndex("FIO");
        if (ulF < 0 || domF < 0) return;
        IndexTag t = TagFor("UL", true);
        if (t == null || t.Terms.Count < 2 || t.TermFields[1] != "DOM") return;
        List<int> searched = new List<int>();
        foreach (int row in res.Rows) if (!res.Added.ContainsKey(row)) searched.Add(row);
        // дома найденных (не больше 30 - иначе поиск по общему адресу читал бы слишком много)
        Dictionary<string, List<int>> houses = new Dictionary<string, List<int>>();
        foreach (int row in searched) {
            string h = HouseKey(tb, row);
            if (h.Length == 0) continue;
            List<int> list;
            if (!houses.TryGetValue(h, out list)) { if (houses.Count >= 30) continue; list = new List<int>(); houses[h] = list; }
            list.Add(row);
        }
        if (houses.Count == 0) return;
        Dictionary<int, int> rowOfRec = new Dictionary<int, int>();
        for (int row = 0; row < tb.Count; row++) rowOfRec[tb.RecNos[row]] = row;
        Seeker s = new Seeker();
        try {
            foreach (List<int> list in houses.Values) {
                int r0 = list[0];
                s.Seek(t, KeyFor(t, new string[] { Field(tb, r0, ulF), Field(tb, r0, domF) }, false));
            }
        } finally {
            s.Close();
        }
        List<int> missing = new List<int>();
        foreach (int rec in s.Found.Keys) if (!rowOfRec.ContainsKey(rec)) missing.Add(rec);
        bool grown = false;
        if (missing.Count > 0) {
            int deleted;
            DbfTable more;
            using (FileStream fs = OpenTable()) more = Fetch(fs, missing, 3000, out deleted);
            for (int row = 0; row < more.Count; row++) {
                int start;
                byte[] buf = more.RecordBuffer(row, out start);
                tb.AddRecord(buf, start, more.RecNos[row]);
                rowOfRec[more.RecNos[row]] = tb.Count - 1;
                grown = true;
            }
        }
        Dictionary<int, bool> shown = new Dictionary<int, bool>();
        foreach (int row in res.Rows) shown[row] = true;
        int added = 0;
        foreach (int rec in s.Found.Keys) {
            int h;
            if (!rowOfRec.TryGetValue(rec, out h) || shown.ContainsKey(h)) continue;
            List<int> list;
            if (!houses.TryGetValue(HouseKey(tb, h), out list)) continue;   // начало номера дома совпало, а дом другой ("180")
            string addr = AddrKey(tb, h), fio = fioF < 0 ? "" : FioKey(Field(tb, h, fioF));
            string hNab = ExtraInfo.Nab(Field(tb, h, nab));
            foreach (int r in list) {
                if (ExtraInfo.Nab(Field(tb, r, nab)) == hNab) continue;
                string why = null;
                if (addr.Length > 0 && addr == AddrKey(tb, r)) why = "тот же адрес - ещё один счётчик";
                else if (fio.Length > 0 && fioF >= 0 && fio == FioKey(Field(tb, r, fioF))) why = "то же ФИО в этом доме";
                if (why == null) continue;
                string main;
                if (Extra != null) {
                    int tr = Extra.Tariff(Field(tb, h, nab), out main);
                    if (tr > 0) {
                        res.Tariff[h] = tr;
                        mainOf[h] = main == null ? hNab : main;
                        string rm;
                        if (mainOf.TryGetValue(r, out rm) && rm == mainOf[h]) { shown[h] = true; break; }   // это тариф того же счётчика
                    }
                }
                res.Rows.Add(h);
                res.Added[h] = true;
                res.Related[h] = why;
                shown[h] = true;
                added++;
                break;
            }
        }
        if (grown) {
            AbonentQuery q = new AbonentQuery(tb, Dir);
            q.Partial = res.Query.Partial;
            q.Address = res.Query.Address;
            q.LastExplain = res.Query.LastExplain;
            q.Miss = res.Query.Miss;
            res.Query = q;
        }
        res.How += "; домов проверено на второй счётчик/то же ФИО: " + houses.Count + ", добавлено абонентов: " + added;
    }

    // Группы (рамка): тарифы одного счётчика, абоненты одного адреса и добавленные к найденному.
    // В группе больше одного счётчика - не тарифные строки оранжевым. Группа стоит на месте первой её строки.
    void Arrange(SearchResult res, Dictionary<int, string> mainOf) {
        DbfTable tb = res.Table;
        Dictionary<int, int> parent = new Dictionary<int, int>();
        foreach (int row in res.Rows) parent[row] = row;
        Dictionary<string, int> firstOf = new Dictionary<string, int>();
        foreach (int row in res.Rows) {
            string m;
            if (mainOf.TryGetValue(row, out m)) Join(parent, firstOf, "T" + m, row);
            string a = AddrKey(tb, row);
            if (a.Length > 0) Join(parent, firstOf, "A" + a, row);
        }
        // один человек (то же ФИО) в одном доме - тоже вместе
        int fioF = tb.FieldIndex("FIO");
        if (fioF >= 0)
            foreach (int row in res.Rows) {
                string h = HouseKey(tb, row), f = FioKey(Field(tb, row, fioF));
                if (h.Length > 0 && f.Length > 0) Join(parent, firstOf, "F" + h + "|" + f, row);
            }
        Dictionary<int, List<int>> members = new Dictionary<int, List<int>>();
        foreach (int row in res.Rows) {
            int root = Find(parent, row);
            List<int> l;
            if (!members.TryGetValue(root, out l)) { l = new List<int>(); members[root] = l; }
            l.Add(row);
        }
        Dictionary<int, int> original = new Dictionary<int, int>();
        for (int i = 0; i < res.Rows.Count; i++) original[res.Rows[i]] = i;
        Dictionary<int, int> clusterPos = new Dictionary<int, int>(), unitPos = new Dictionary<int, int>();
        Dictionary<string, int> unitFirst = new Dictionary<string, int>();
        foreach (KeyValuePair<int, List<int>> kv in members) {
            List<int> l = kv.Value;
            int pos = int.MaxValue;
            Dictionary<string, bool> units = new Dictionary<string, bool>();
            foreach (int row in l) {
                pos = Math.Min(pos, original[row]);
                string m;
                string unit = mainOf.TryGetValue(row, out m) ? "T" + m : "R" + row;
                units[unit] = true;
                int up;
                if (!unitFirst.TryGetValue(unit, out up) || original[row] < up) unitFirst[unit] = original[row];
            }
            foreach (int row in l) {
                clusterPos[row] = pos;
                if (l.Count > 1) res.Group[row] = "G" + kv.Key;
                if (units.Count > 1 && !res.Tariff.ContainsKey(row)) res.Orange[row] = true;
            }
        }
        foreach (int row in res.Rows) {
            string m;
            unitPos[row] = unitFirst[mainOf.TryGetValue(row, out m) ? "T" + m : "R" + row];
        }
        List<int> order = new List<int>(res.Rows);
        order.Sort(delegate(int x, int y) {
            int c = clusterPos[x].CompareTo(clusterPos[y]);
            if (c == 0) c = unitPos[x].CompareTo(unitPos[y]);
            if (c == 0) { int tx, ty; res.Tariff.TryGetValue(x, out tx); res.Tariff.TryGetValue(y, out ty); c = tx.CompareTo(ty); }
            return c != 0 ? c : original[x].CompareTo(original[y]);
        });
        res.Rows = order;
    }

    static int Find(Dictionary<int, int> parent, int x) {
        while (parent[x] != x) { parent[x] = parent[parent[x]]; x = parent[x]; }
        return x;
    }

    static void Join(Dictionary<int, int> parent, Dictionary<string, int> firstOf, string key, int row) {
        int first;
        if (!firstOf.TryGetValue(key, out first)) { firstOf[key] = row; return; }
        int a = Find(parent, first), b = Find(parent, row);
        if (a != b) parent[b] = a;
    }

    // Нашёлся лицевой одного из тарифов - добавляем остальные тарифы того же счётчика (их не искали - помечаются)
    void AddSiblings(SearchResult res, int nab, Dictionary<int, string> mainOf) {
        Dictionary<string, int> rowOf = new Dictionary<string, int>();
        for (int row = 0; row < res.Table.Count; row++) {
            string k = ExtraInfo.Nab(DbfTable.Format(res.Table.GetValue(row, nab, null), true));
            if (!rowOf.ContainsKey(k)) rowOf[k] = row;
        }
        Dictionary<int, bool> shown = new Dictionary<int, bool>();
        foreach (int row in res.Rows) shown[row] = true;
        List<string> missing = new List<string>();
        Dictionary<string, string> wantMain = new Dictionary<string, string>();
        Dictionary<string, bool> doneGroup = new Dictionary<string, bool>();
        foreach (string main in mainOf.Values) {
            if (doneGroup.ContainsKey(main)) continue;
            doneGroup[main] = true;
            foreach (KeyValuePair<int, string> sib in Extra.Group(main)) {
                string k = ExtraInfo.Nab(sib.Value);
                wantMain[k] = main;
                if (!rowOf.ContainsKey(k)) missing.Add(sib.Value);
            }
        }
        bool grown = false;
        if (missing.Count > 0) {
            try {
                DbfTable more;
                Dictionary<string, int> found = LookupNabs(missing, out more);
                foreach (KeyValuePair<string, int> kv in found) {
                    int start;
                    byte[] buf = more.RecordBuffer(kv.Value, out start);
                    res.Table.AddRecord(buf, start, more.RecNos[kv.Value]);
                    rowOf[kv.Key] = res.Table.Count - 1;
                    grown = true;
                }
            } catch (Exception ex) {
                res.How += "; другие тарифы не прочитались: " + ex.Message;
            }
        }
        foreach (KeyValuePair<string, string> kv in wantMain) {
            int row;
            if (!rowOf.TryGetValue(kv.Key, out row) || shown.ContainsKey(row)) continue;
            string main;
            int t = Extra.Tariff(kv.Key, out main);
            if (t <= 0) continue;
            res.Rows.Add(row);
            res.Added[row] = true;
            res.Tariff[row] = t;
            mainOf[row] = kv.Value;
            shown[row] = true;
        }
        if (grown) {
            // в таблице результата добавились записи - столбцы (улица по справочнику) считаем заново
            AbonentQuery q = new AbonentQuery(res.Table, Dir);
            q.Partial = res.Query.Partial;
            q.Address = res.Query.Address;
            q.LastExplain = res.Query.LastExplain;
            q.Miss = res.Query.Miss;
            res.Query = q;
        }
        if (res.Added.Count > 0) res.How += "; добавлено лицевых других тарифов тех же счётчиков: " + res.Added.Count;
    }

    // Записи по номерам абонентов (по индексу NAB): номер (без нулей впереди) -> строка в table. null - нет индекса
    public Dictionary<string, int> LookupNabs(List<string> nabs, out DbfTable table) {
        Dictionary<string, int> result = new Dictionary<string, int>();
        table = null;
        IndexTag t = TagFor("NAB");
        if (t == null) return result;
        int f = Header.FieldIndex("NAB");
        int width = Header.Fields[f].Length;
        Seeker s = new Seeker();
        try {
            foreach (string raw in nabs) {
                string v = raw.Trim();
                if (v.Length == 0) continue;
                List<string> variants = new List<string>();
                AddVariant(variants, v);
                AddVariant(variants, ExtraInfo.Nab(v));
                string core = ExtraInfo.Nab(v);
                for (int w = core.Length + 1; w <= width && Header.Fields[f].Type == 'C'; w++) AddVariant(variants, core.PadLeft(w, '0'));
                foreach (string x in variants) s.Seek(t, KeyFor(t, new string[] { x }, false));
            }
        } finally {
            s.Close();
        }
        int deleted;
        using (FileStream fs = OpenTable()) table = Fetch(fs, new List<int>(s.Found.Keys), 2000, out deleted);
        Dictionary<string, bool> wanted = new Dictionary<string, bool>();
        foreach (string raw in nabs) wanted[ExtraInfo.Nab(raw)] = true;
        for (int row = 0; row < table.Count; row++) {
            string k = ExtraInfo.Nab(DbfTable.Format(table.GetValue(row, f, null), true));
            if (wanted.ContainsKey(k) && !result.ContainsKey(k)) result[k] = row;
        }
        return result;
    }

    static string Digits(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s) if (c >= '0' && c <= '9') sb.Append(c);
        return sb.ToString();
    }

    static string LongestDigits(string s) {
        string best = "", cur = "";
        foreach (char c in s + " ") {
            if (c >= '0' && c <= '9') { cur += c; continue; }
            if (cur.Length > best.Length) best = cur;
            cur = "";
        }
        return best;
    }

    // Как в этой таблице обычно записано поле: длина числа, нули впереди, прижато вправо, заглавные буквы
    class Style {
        public int Digits;
        public bool Zeros, Right, Upper;
    }

    Style StyleOf(string field) {
        Style st = new Style();
        int f = Sample.FieldIndex(field);
        if (f < 0 || Sample.Count == 0) return st;
        Dictionary<int, int> lengths = new Dictionary<int, int>();
        int filled = 0, zeros = 0, right = 0, upper = 0, letters = 0;
        for (int r = 0; r < Sample.Count; r++) {
            int start;
            byte[] buf = Sample.RecordBuffer(r, out start);
            string raw = Sample.Enc.GetString(buf, start + Sample.Fields[f].Offset, Sample.Fields[f].Length);
            string v = raw.Trim();
            if (v.Length == 0) continue;
            filled++;
            if (raw.StartsWith(" ")) right++;
            if (v.StartsWith("0")) zeros++;
            int n;
            lengths.TryGetValue(v.Length, out n);
            lengths[v.Length] = n + 1;
            foreach (char c in v) if (char.IsLetter(c)) { letters++; if (char.IsUpper(c)) upper++; }
        }
        int best = 0;
        foreach (KeyValuePair<int, int> kv in lengths) if (kv.Value > best) { best = kv.Value; st.Digits = kv.Key; }
        st.Zeros = filled > 0 && zeros * 2 > filled;
        st.Right = filled > 0 && right * 2 > filled;
        st.Upper = letters > 0 && upper * 10 > letters * 7;
        return st;
    }

    static void AddVariant(List<string> list, string v) {
        if (v != null && v.Length > 0 && !list.Contains(v)) list.Add(v);
    }

    // текст как записан в поле: прижат вправо, если так принято в таблице
    string Aligned(string field, string v, Style st) {
        int f = Header.FieldIndex(field);
        if (f < 0 || !st.Right || Header.Fields[f].Type != 'C') return v;
        return v.PadLeft(Header.Fields[f].Length);
    }

    // Лицевой: точный номер в нескольких видах (с нулями впереди и без);
    // при частичном совпадении ещё номера, начинающиеся так, и номер внутри длинного Л/С
    bool SeekAccount(string text, Seeker s, SearchResult res, bool partial) {
        string digits = Digits(text);
        if (digits.Length == 0) { res.Note = "Впишите номер."; return false; }
        string core = digits.TrimStart('0');
        if (core.Length == 0) core = "0";
        bool any = false;
        foreach (string field in AccountFields) {
            IndexTag t = TagFor(field);
            if (t == null) continue;
            any = true;
            int f = Header.FieldIndex(field);
            int width = Header.Fields[f].Length;
            List<string> exact = new List<string>();
            AddVariant(exact, digits);
            AddVariant(exact, core);
            if (Header.Fields[f].Type == 'C') for (int w = core.Length + 1; w <= width; w++) AddVariant(exact, core.PadLeft(w, '0'));
            foreach (string v in exact) {
                s.Seek(t, KeyFor(t, new string[] { v }, false));
                if (Header.Fields[f].Type == 'C') s.Seek(t, KeyFor(t, new string[] { v.PadLeft(width) }, false));
            }
            if (s.Found.Count > 0) { res.How = "точный номер (поле " + field + ")"; return true; }
        }
        if (!any) return false;
        if (!partial) { res.How = "точный номер"; return true; }
        // начало номера (только если номер не короче 4 цифр - иначе слишком много)
        if (digits.Length >= 4) {
            foreach (string field in AccountFields) {
                IndexTag t = TagFor(field);
                if (t == null || Header.Fields[Header.FieldIndex(field)].Type != 'C') continue;
                s.Seek(t, KeyFor(t, new string[] { digits }, true));
                if (core != digits) s.Seek(t, KeyFor(t, new string[] { core }, true));
                // в таблице номера с нулями впереди (044060216) - "4406021" ищем и как "04406021"
                Style st = StyleOf(field);
                if (st.Zeros) for (string z = "0" + core; z.Length < st.Digits; z = "0" + z) if (z != digits) s.Seek(t, KeyFor(t, new string[] { z }, true));
                if (s.Found.Count > 0) { res.How = "начало номера (поле " + field + ")"; return true; }
            }
        }
        // длинный Л/С, внутри которого номер абонента
        IndexTag nab = TagFor("NAB");
        if (nab != null && core.Length >= 8) {
            Style st = StyleOf("NAB");
            int len = st.Digits > 0 ? st.Digits : 9;
            for (int from = 0; from + len <= digits.Length && from < 12; from++) {
                string piece = digits.Substring(from, len);
                if (piece.TrimStart('0').Length < 5) continue;
                s.Seek(nab, KeyFor(nab, new string[] { piece }, false));
                s.Seek(nab, KeyFor(nab, new string[] { piece.TrimStart('0') }, false));
            }
            if (s.Found.Count > 0) res.How = "номер абонента внутри длинного номера";
        }
        return true;
    }

    // Счётчик: точный номер (нули впереди, пробелы, "№" не важны); при частичном совпадении - ещё начало номера.
    // Если не нашлось - то же по снятому счётчику (NSC)
    bool SeekMeter(string text, Seeker s, SearchResult res, bool partial) {
        string typed = text.Replace("№", " ").Trim();
        if (typed.Length == 0) { res.Note = "Впишите номер."; return false; }
        bool any = false;
        foreach (string field in new string[] { "NCH", "NSC" }) {
            IndexTag t = TagFor(field);
            if (t == null) continue;
            any = true;
            Style st = StyleOf(field);
            int width = Header.Fields[Header.FieldIndex(field)].Length;
            List<string> variants = new List<string>();
            AddVariant(variants, typed);
            AddVariant(variants, typed.ToUpperInvariant());
            foreach (string key in AbonentQuery.MeterKeys(typed)) {
                AddVariant(variants, key);
                AddVariant(variants, key.ToUpperInvariant());
                if (key.Length > 0 && char.IsDigit(key[0])) for (int w = key.Length + 1; w <= Math.Min(width, 12); w++) AddVariant(variants, key.PadLeft(w, '0'));
            }
            foreach (string v in variants) s.Seek(t, KeyFor(t, new string[] { Aligned(field, v, st) }, false));
            if (partial && s.Found.Count == 0 && !st.Right) foreach (string v in variants) if (v.Length >= 3) s.Seek(t, KeyFor(t, new string[] { v }, true));
            if (s.Found.Count > 0) { res.How = field == "NSC" ? "по снятому счётчику (NSC)" : "поле " + field; return true; }
        }
        return any;
    }

    static byte[] WholeWord(byte[] key, IndexTag t, bool whole) {
        if (key == null || !whole || t.Collated || key.Length >= t.KeyLen) return key;
        byte[] k = new byte[key.Length + 1];
        Buffer.BlockCopy(key, 0, k, 0, key.Length);
        k[key.Length] = (byte)' ';
        return k;
    }

    static string Title(string s) {
        StringBuilder sb = new StringBuilder();
        bool start = true;
        foreach (char c in s) {
            sb.Append(start ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c));
            start = c == ' ' || c == '-' || c == '.';
        }
        return sb.ToString();
    }

    // ФИО: точно - фамилия (и имя, отчество, если вписаны) целыми словами: "иванов" не находит "Иванова";
    // частично - ФИО, начинающиеся так. Если целиком не нашлось - по одной фамилии (дальше сверяется в записях)
    bool SeekName(string text, Seeker s, SearchResult res, bool partial) {
        string[] words = DbfTable.SplitWords(text.Replace('.', ' '));
        if (words.Length == 0) { res.Note = "Впишите фамилию."; return false; }
        IndexTag t = TagFor("FIO");
        if (t == null) return false;
        Style st = StyleOf("FIO");
        for (int n = words.Length; n >= 1; n = n > 1 ? 1 : 0) {
            string q = string.Join(" ", words, 0, n);
            List<string> variants = new List<string>();
            if (st.Upper) AddVariant(variants, q.ToUpperInvariant());
            AddVariant(variants, Title(q));
            AddVariant(variants, q.ToUpperInvariant());
            AddVariant(variants, q);
            // целое слово: после него в ФИО пробел (кроме инициала - "иванов и")
            bool whole = !partial && words[n - 1].Length > 1;
            foreach (string v in variants) s.Seek(t, WholeWord(KeyFor(t, new string[] { v }, true), t, whole));
            if (s.Found.Count > 0) { res.How = n == words.Length ? "начало ФИО" : "по фамилии"; return true; }
        }
        return true;
    }

    // ФИО с адресом: "Иванов Боровляны", "Иванов Иван, Боровляны Заводская 18". Индекса по ФИО в базе нет,
    // поэтому по индексу адреса берём абонентов этого места и среди них ищем ФИО (вся таблица не читается).
    // Адрес - после запятой или последние слова, которые есть в справочнике улиц и населённых пунктов.
    // "Иванов Боровляны" -> ФИО "Иванов", адрес "Боровляны": адрес - после запятой или последние слова,
    // которые есть в справочнике улиц и населённых пунктов
    bool SplitNameAddress(string text, out string name, out string addr) {
        name = null;
        addr = null;
        if (Dir == null) return false;
        int comma = text.IndexOf(',');
        if (comma > 0) {
            name = text.Substring(0, comma).Trim();
            addr = text.Substring(comma + 1).Trim();
        } else {
            string[] w = text.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            // ФИО - первые k слов, остальное - адрес (сначала пробуем адрес подлиннее)
            for (int k = 1; k < w.Length && addr == null; k++) {
                if (char.IsDigit(w[k][0])) continue;   // число - часть ФИО или дом, адрес с него не начинается
                string a = string.Join(" ", w, k, w.Length - k);
                AddressMatch m = Dir.Match(a);
                if (m.Problem == null && m.Codes.Count > 0) { name = string.Join(" ", w, 0, k); addr = a; }
            }
        }
        return addr != null && name != null && name.Length > 0 && addr.Length > 0;
    }

    bool SeekNameByAddress(string text, Seeker s, SearchResult res, bool partial) {
        string name, addr;
        if (!SplitNameAddress(text, out name, out addr)) return false;
        if (TagFor("FIO") != null) {
            // есть индекс по ФИО: ищем ФИО по нему, адрес сверяется в найденных
            AddressMatch m = Dir.Match(addr);
            res.Address = m;
            if (m.Problem != null) { res.Note = m.Problem; return false; }
            if (!SeekName(name, s, res, partial)) return false;
            res.How = "ФИО по индексу, адрес сверяется";
        } else {
            if (!SeekAddressExact(addr, s, res)) return false;
            res.How = "ФИО среди абонентов по адресу";
        }
        res.NameText = name;
        res.AddressText = addr;
        return true;
    }

    // Точный адрес: населённый пункт -> улица в нём -> дом -> корпус -> квартира (см. StreetDirectory.Match).
    // По индексу берём абонентов найденных улиц - сразу с домом, если индекс по улице и дому.
    bool SeekAddressExact(string text, Seeker s, SearchResult res) {
        AddressMatch m = Dir.Match(text);
        res.Address = m;
        if (m.Problem != null) { res.Note = m.Problem; return false; }
        IndexTag t = TagFor("UL", true);
        if (t == null) return false;
        int got = 0;
        if (m.House != null && t.Terms.Count > 1 && t.TermFields[1] == "DOM") {
            Style st = StyleOf("DOM");
            List<string> houses = new List<string>();
            foreach (string h in new string[] { m.House, m.House.ToUpperInvariant() }) {
                AddVariant(houses, h);
                AddVariant(houses, Aligned("DOM", h, st));
            }
            foreach (string code in m.Codes) foreach (string h in houses) got += s.Seek(t, KeyFor(t, new string[] { code, h }, false));
        }
        // дом записан в базе иначе (или дом не указан) - берём улицы целиком, дом сверяется в записях
        if (got == 0) foreach (string code in m.Codes) s.Seek(t, KeyFor(t, new string[] { code }, false));
        return true;
    }

    // Частичное совпадение адреса: улицы, где совпало больше всего слов (как раньше)
    bool SeekAddress(string text, Seeker s, SearchResult res) {
        AddressParts a = AddressParts.Parse(text);
        if (a.Words.Count == 0) { res.Note = a.House != null ? "Впишите улицу или населённый пункт, не только номер дома." : "Впишите адрес."; return false; }
        IndexTag t = TagFor("UL", true);
        if (t == null) return false;
        if (Dir == null) {
            // в UL названия улиц: ищем по началу названия
            foreach (string w in a.Words) {
                if (w.Length < 3) continue;
                s.Seek(t, KeyFor(t, new string[] { w.ToUpperInvariant() }, true));
                s.Seek(t, KeyFor(t, new string[] { Title(w) }, true));
                s.Seek(t, KeyFor(t, new string[] { w }, true));
            }
            res.How = "по началу названия улицы";
            return true;
        }
        int best;
        // улицы ищем во всех сельсоветах: если адрес в невыбранном, покажем, что он там есть (скрытым)
        List<string> codes = Dir.BestStreets(a.Words, null, out best);
        res.How = "улиц по справочнику: " + codes.Count + " (совпало слов " + best + " из " + a.Words.Count + ")";
        if (codes.Count == 0) { res.Note = "Такой улицы или населённого пункта нет в справочнике улиц."; return false; }
        if (codes.Count > 300) { res.Note = "Подходит слишком много улиц (" + codes.Count + ") - уточните адрес."; return false; }
        bool withHouse = a.House != null && t.Terms.Count > 1 && t.TermFields[1] == "DOM";
        foreach (string code in codes) {
            int got = 0;
            if (withHouse) {
                got = s.Seek(t, KeyFor(t, new string[] { code, a.House.ToUpperInvariant() }, false));
                if (got == 0 && a.House.ToUpperInvariant() != a.House) got = s.Seek(t, KeyFor(t, new string[] { code, a.House }, false));
            }
            if (got == 0) s.Seek(t, KeyFor(t, new string[] { code }, false));
        }
        return true;
    }
}

// ---- точный адрес: населённый пункт -> улица в нём -> дом -> корпус -> квартира ----

// Слово адреса после разбора
public class AddrToken {
    public string Text;
    public char Kind;      // w - слово или число названия, h - "дом", k - "корпус", f - "квартира",
                           // t - тип населённого пункта (аг., д.), s - тип улицы (ул., пер.), r - лишнее (область, район, "с/с")
    public bool Ss;        // название сельсовета (стоит перед "с/с")
    public bool Bound;     // число сразу после "дом"/"корп"/"кв" - не часть названия
    public bool Used;      // уже узнано как название
    public bool Digit { get { return Text.Length > 0 && char.IsDigit(Text[0]); } }
}

// Слова адреса и названия из справочников приводятся к одному виду: маленькие буквы, ё -> е,
// без знаков препинания, сокращения через дефис ("пр-т", "с/с") - одним словом
public static class AddressText {
    static readonly string[][] Glued = {
        new string[] { "пр-т", "прт" }, new string[] { "б-р", "бр" }, new string[] { "р-н", "рн" }, new string[] { "с/с", "сс" },
        new string[] { "кв-л", "квл" }, new string[] { "кв-ра", "квра" }, new string[] { "пр-д", "прд" }, new string[] { "г.п.", "гп" },
        new string[] { "р.п.", "рп" }, new string[] { "а.г.", "аг" } };
    public static readonly string[] TownTypes = {
        "г", "город", "аг", "агрогородок", "п", "пос", "поселок", "гп", "рп", "кп", "д", "дер", "деревня", "с", "село",
        "х", "хут", "хутор", "ст", "станция" };
    public static readonly string[] StreetTypes = {
        "ул", "улица", "пер", "переулок", "пр", "прт", "просп", "проспект", "пл", "площадь", "бр", "бульв", "бульвар",
        "ш", "шоссе", "туп", "тупик", "проезд", "прд", "мкр", "мкрн", "микрорайон", "наб", "набережная", "тракт",
        "линия", "квл", "квартал", "аллея", "спуск" };
    static readonly string[] Regions = { "беларусь", "республика", "рб" };

    public static List<string> Words(string s) {
        string t = " " + s.ToLowerInvariant().Replace('ё', 'е') + " ";
        foreach (string[] g in Glued) t = t.Replace(g[0], " " + g[1] + " ");
        StringBuilder sb = new StringBuilder();
        foreach (char c in t) sb.Append(char.IsLetterOrDigit(c) || c == '/' ? c : ' ');
        List<string> list = new List<string>();
        foreach (string w in sb.ToString().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) list.Add(w.Trim('/'));
        list.RemoveAll(delegate(string w) { return w.Length == 0; });
        return list;
    }

    // слова названия из справочника - без типов ("аг. Боровляны" -> боровляны, "Заводская ул." -> заводская)
    public static string[] NameWords(string s) {
        List<string> list = new List<string>();
        foreach (string w in Words(s))
            if (Array.IndexOf(TownTypes, w) < 0 && Array.IndexOf(StreetTypes, w) < 0) list.Add(w);
        return list.ToArray();
    }

    public static List<AddrToken> Tokens(string query) {
        List<string> words = new List<string>();
        foreach (string w in Words(query)) {
            // слитно: "д18", "кв5", "к2"
            string marker = null;
            if (w.Length > 2 && w.StartsWith("кв") && char.IsDigit(w[2])) marker = "кв";
            else if (w.Length > 1 && (w[0] == 'д' || w[0] == 'к') && char.IsDigit(w[1])) marker = w.Substring(0, 1);
            if (marker != null) { words.Add(marker); words.Add(w.Substring(marker.Length)); }
            else words.Add(w);
        }
        List<AddrToken> toks = new List<AddrToken>();
        for (int i = 0; i < words.Count; i++) {
            string w = words[i];
            bool nextDigit = i + 1 < words.Count && char.IsDigit(words[i + 1][0]);
            AddrToken t = new AddrToken();
            t.Text = w;
            AddrToken prev = toks.Count > 0 ? toks[toks.Count - 1] : null;
            if (w == "дом" || (w == "д" && nextDigit)) t.Kind = 'h';
            else if ((w == "к" || w == "корп" || w == "корпус") && nextDigit) t.Kind = 'k';
            else if (w == "кв" || w == "квартира" || w == "квра") t.Kind = 'f';
            else if (w == "к") t.Kind = 'r';
            else if (Array.IndexOf(TownTypes, w) >= 0) t.Kind = 't';
            else if (Array.IndexOf(StreetTypes, w) >= 0) t.Kind = 's';
            else if (w == "сс" || w == "сельсовет") { t.Kind = 'r'; if (prev != null && prev.Kind == 'w' && !prev.Digit) prev.Ss = true; }
            else if (w == "рн" || w == "район" || w == "обл" || w == "область") { t.Kind = 'r'; if (prev != null && prev.Kind == 'w' && !prev.Digit) prev.Kind = 'r'; }
            else if (Array.IndexOf(Regions, w) >= 0) t.Kind = 'r';
            else t.Kind = 'w';
            if (t.Kind == 'w' && t.Digit && prev != null && (prev.Kind == 'h' || prev.Kind == 'k' || prev.Kind == 'f')) t.Bound = true;
            toks.Add(t);
        }
        return toks;
    }
}

// Адрес, узнанный по справочнику. Каждая указанная часть должна совпасть точно, не указанная - не ограничивает.
public class AddressMatch {
    public List<string> Codes = new List<string>();   // коды улиц, где искать
    public string House, Korp, Flat;
    public string Problem;          // почему искать нечего (null - всё понятно)
    public string Said = "";        // что поняли: "аг. Боровляны -> Заводская ул. -> дом 18" (для строки состояния)
    public string Explain = "";     // то же без названий - для отчёта

    // дом, корпус, квартира для сравнения: без пробелов и дефисов, маленькими буквами, без нулей впереди
    public static string Num(string s) {
        if (s == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (char c in s.ToLowerInvariant()) if (c != ' ' && c != '-' && c != '.' && c != ',') sb.Append(c);
        string r = sb.ToString();
        int z = 0;
        while (z < r.Length - 1 && r[z] == '0' && char.IsDigit(r[z + 1])) z++;
        return r.Substring(z);
    }
}

// Адрес, разобранный на части: слова (населённый пункт, улица), дом, корпус, квартира (для частичного совпадения)
public class AddressParts {
    public List<string> Words = new List<string>();
    public string House, Korp, Flat;
    public int Skipped;

    static readonly string[] Skip = {
        "ул", "улица", "пер", "переулок", "пр", "пр-т", "просп", "проспект", "пл", "площадь", "б-р", "бульвар",
        "ш", "шоссе", "туп", "тупик", "проезд", "мкр", "микрорайон", "наб", "набережная", "тракт",
        "г", "город", "аг", "агрогородок", "п", "пос", "поселок", "посёлок", "гп", "рп", "дер", "деревня",
        "с", "село", "х", "хутор", "с/с", "сс", "сельсовет", "р-н", "район", "обл", "область" };

    // "Острошицкий Городок Заводская ул. 18", "аг. Острошицкий Городок, ул. Заводская, д. 18, кв. 5", "заводская 18 5"
    public static AddressParts Parse(string query) {
        AddressParts a = new AddressParts();
        List<string> numbers = new List<string>();
        List<string> early = new List<string>();   // числа до названия: номер строки из списка или сам дом
        string pending = null;                     // "дом", "корп" или "кв" перед числом
        foreach (string raw in DbfTable.SplitWords(query.Replace('.', ' '))) {
            string w = raw;
            // слитно: "кв5", "д18", "к2"
            string marker = null;
            if (w.Length > 2 && w.StartsWith("кв") && IsDigitAt(w, 2)) { marker = "кв"; w = w.Substring(2); }
            else if (w.Length > 1 && (w[0] == 'д' || w[0] == 'к') && IsDigitAt(w, 1)) { marker = w[0] == 'д' ? "дом" : "корп"; w = w.Substring(1); }
            if (marker != null) pending = marker;

            if (w == "д" || w == "дом") { pending = "дом"; continue; }
            if (w == "к" || w == "корп" || w == "корпус") { pending = "корп"; continue; }
            if (w == "кв" || w == "квартира") { pending = "кв"; continue; }
            if (Array.IndexOf(Skip, w) >= 0) { pending = null; a.Skipped++; continue; }

            if (IsDigitAt(w, 0)) {
                if (IsDigits(w) && w.Length >= 5) { pending = null; a.Skipped++; continue; }   // лицевой или код - не адрес
                if (pending == "корп" && a.Korp == null) a.Korp = w;
                else if (pending == "кв" && a.Flat == null) a.Flat = w;
                else if (pending == "дом" && a.House == null) a.House = w;
                else if (a.Words.Count > 0) numbers.Add(w);
                else early.Add(w);
                pending = null;
                continue;
            }
            pending = null;
            if (HasDigit(w)) { a.Skipped++; continue; }
            a.Words.Add(w);
        }
        // числа до названия улицы - обычно номер строки из списка, их пропускаем;
        // но если слов нет совсем ("18"), это и есть дом
        if (a.Words.Count == 0) numbers.InsertRange(0, early);
        else a.Skipped += early.Count;
        // числа без пояснений: первое - дом, второе - квартира
        foreach (string n in numbers) {
            if (a.House == null) a.House = n;
            else if (a.Flat == null) a.Flat = n;
            else if (a.Korp == null) a.Korp = n;
        }
        return a;
    }

    static bool IsDigitAt(string s, int i) {
        return i < s.Length && s[i] >= '0' && s[i] <= '9';
    }

    static bool IsDigits(string s) {
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return s.Length > 0;
    }

    static bool HasDigit(string s) {
        foreach (char c in s) if (c >= '0' && c <= '9') return true;
        return false;
    }
}

public static class FieldNames {
    static Dictionary<string, string> names;

    public static string Label(string field) {
        if (names == null) {
            names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string[] pairs = {
                "NAB", "№ абонента", "FIO", "Ф.И.О.", "UL", "Улица", "DOM", "Дом", "KOR", "Корпус", "KV", "Квартира",
                "TEL", "Телефон", "KNIGA", "№ книги", "LGI", "Код льготы", "SKID", "% скидки", "KTAR", "Код тарифа",
                "GIL", "Жильцов", "KLG", "Льготных жильцов", "TPO", "Тип помещения", "ROZ", "Розеток",
                "TOS", "Точек освещения", "NCH", "№ счётчика", "TIPCH", "Тип счётчика", "RAS", "Разрядность",
                "GODPR", "Год поверки", "KVP", "Квартал поверки", "DZA", "Дата замены счётчика",
                "PZS", "Показание при замене", "NSC", "№ снятого счётчика", "TSC", "Тип снятого счётчика",
                "NPOK", "Показание при установке", "DPR", "Дата проверки", "POK", "Показание", "PPR", "Пр. проверки",
                "OPOK", "Оплаченное показание", "PRA", "Признак архива", "DOT", "Дата отключения", "PRP", "Пр. пл.",
                "DUC", "Дата учёта без пр.", "DOK", "Дата окончания", "N_UCH", "№ участка", "N_F", "№ фидера",
                "LS_ERIP", "Л/С ЕРИП" };
            for (int i = 0; i + 1 < pairs.Length; i += 2) names[pairs[i]] = pairs[i + 1];
        }
        string label;
        return names.TryGetValue(field, out label) ? label : null;
    }

    public static string Title(string field) {
        string label = Label(field);
        return label == null ? field : label + " (" + field + ")";
    }
}

// Карточка записи: все поля целиком, с русскими подписями, где они известны
// Как перейти к соседнему абоненту из списка найденного (PgUp/PgDn в карточке): false - дальше некуда
public delegate bool CardMove(int delta, out SearchResult res, out int row);

// Справочник "код -> название" для карточки (тарифы, льготы, место установки счётчика):
// маленькая таблица базы, читается целиком один раз
public class CodeBook {
    readonly Dictionary<string, string> names = new Dictionary<string, string>();
    public string Note = "";
    static readonly Dictionary<string, CodeBook> cache = new Dictionary<string, CodeBook>(StringComparer.OrdinalIgnoreCase);

    public static CodeBook Get(string folder, string table, string[] keys) {
        string id = folder + "|" + table;
        CodeBook b;
        lock (cache) if (cache.TryGetValue(id, out b)) return b;
        b = new CodeBook();
        try { b.Load(folder, table, keys); }
        catch (Exception ex) { b.Note = "не прочитался: " + ex.Message; }
        Report.Line("Справочник для карточки " + table + ": " + b.Note);
        lock (cache) cache[id] = b;
        return b;
    }

    void Load(string folder, string table, string[] keys) {
        string path = null;
        foreach (string f in Directory.GetFiles(folder))
            if (string.Compare(Path.GetFileName(f), table, StringComparison.OrdinalIgnoreCase) == 0) path = f;
        if (path == null) { Note = "таблицы нет"; return; }
        DbfTable h = DbfTable.LoadHeader(path);
        if (h.Expected > 20000) { Note = "слишком большая (" + h.Expected + " записей)"; return; }
        int key = -1;
        foreach (string k in keys) if (key < 0) key = h.FieldIndex(k);
        for (int i = 0; i < h.Fields.Count && key < 0; i++) if (h.Fields[i].Type == 'N') key = i;
        // название - текстовое поле, лучше длинное и с "говорящим" именем (NAIM..., N...)
        int name = -1, best = 0;
        for (int i = 0; i < h.Fields.Count; i++) {
            DbfField f = h.Fields[i];
            if (i == key || f.Type != 'C') continue;
            string n = f.Name.ToUpperInvariant();
            int score = f.Length >= 10 ? 2 : 1;
            if (n.Contains("NAIM") || n.Contains("NAME") || n.StartsWith("N")) score += 2;
            if (score > best) { best = score; name = i; }
        }
        if (key < 0 || name < 0) { Note = "не понятно, где код и где название"; return; }
        DbfTable t = DbfTable.Load(path, null, null, 20000);
        for (int row = 0; row < t.Count; row++) {
            string k = Code(DbfTable.Format(t.GetValue(row, key, null), true));
            if (k.Length > 0 && !names.ContainsKey(k)) names[k] = DbfTable.Format(t.GetValue(row, name, null), true).Trim();
        }
        Note = "код - поле " + h.Fields[key].Name + ", название - поле " + h.Fields[name].Name + ", кодов " + names.Count;
    }

    static string Code(string s) {
        s = s.Trim();
        decimal d;
        if (decimal.TryParse(s.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out d))
            return d.ToString("0.##########", CultureInfo.InvariantCulture);
        return s.ToUpperInvariant();
    }

    public string Name(string code) {
        string s;
        return names.TryGetValue(Code(code), out s) ? s : "";
    }
}

// Карточка абонента как в программе базы: те же надписи и поля на тех же местах (экран в 80 знаков).
// Ctrl+D - доп. сведения, F2 - все поля списком, PgUp/PgDn - соседний абонент из списка, Esc - закрыть.
public static class Places {
    public const string Network = @"\\fs-dos\communal\BASAFOX";

    // Батник базы подключает \\fs-dos\communal как диск O:, но только пока база запущена,
    // поэтому пробуем и диск, и прямой сетевой путь
    public static string FindFolder(string saved) {
        string[] candidates = { saved, @"O:\BASAFOX", Network };
        foreach (string c in candidates) {
            try { if (c != null && Directory.Exists(c)) return c; }
            catch (Exception) { }
        }
        return saved != null ? saved : Network;
    }

    public static string[] DbfFiles(string folder) {
        List<string> list = new List<string>();
        foreach (string f in Directory.GetFiles(folder))
            if (string.Compare(Path.GetExtension(f), ".dbf", StringComparison.OrdinalIgnoreCase) == 0) list.Add(f);
        string[] files = list.ToArray();
        Array.Sort(files, StringComparer.OrdinalIgnoreCase);
        return files;
    }

    public static string ReadSetting(string file, int line) {
        try {
            if (!File.Exists(file)) return null;
            string[] lines = File.ReadAllLines(file, Encoding.UTF8);
            return line < lines.Length && lines[line].Trim().Length > 0 ? lines[line].Trim() : null;
        } catch (Exception) {
            return null;
        }
    }

    public static void SaveSettings(string file, string[] lines) {
        try { File.WriteAllLines(file, lines, Encoding.UTF8); }
        catch (Exception) { }   // папка может быть только для чтения - тогда просто не запоминаем
    }
}

public class TableInfo {
    public string Name;
    public int Records;
    public bool Abonents;

    public override string ToString() {
        return Name + "   (" + Records + " записей)";
    }

    // Таблицы папки, похожие на таблицу абонентов - первыми; архивы (ARC...) - после основных
    public static List<TableInfo> List(string folder, LoadProgress progress) {
        List<TableInfo> list = new List<TableInfo>();
        string[] files = Places.DbfFiles(folder);
        for (int i = 0; i < files.Length; i++) {
            if (progress != null) progress(i, files.Length);
            TableInfo info = new TableInfo();
            info.Name = Path.GetFileName(files[i]);
            try {
                DbfTable h = DbfTable.LoadHeader(files[i]);
                info.Records = h.Expected;
                info.Abonents = h.FieldIndex("NAB") >= 0 && h.FieldIndex("FIO") >= 0
                    && (h.FieldIndex("DOM") >= 0 || h.FieldIndex("UL") >= 0);
            } catch (Exception) {
                continue;
            }
            list.Add(info);
        }
        list.Sort(delegate(TableInfo a, TableInfo b) {
            if (a.Abonents != b.Abonents) return a.Abonents ? -1 : 1;
            bool aa = a.Name.ToUpperInvariant().StartsWith("ARC"), ba = b.Name.ToUpperInvariant().StartsWith("ARC");
            if (aa != ba) return aa ? 1 : -1;
            if (a.Records != b.Records) return b.Records.CompareTo(a.Records);
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        });
        return list;
    }
}

// Отчёт для настройки: что из какой таблицы взялось и как прошёл поиск.
// Данных абонентов в нём нет: только имена таблиц и полей, вид значений
// (буквы заменены на А/а, цифры на 9) и количества.
public static class Report {
    static string path;
    static int searches;
    static readonly string[] Keep = {
        "ул", "улица", "пер", "пр", "пр-т", "пл", "б-р", "ш", "д", "дом", "кв", "к", "корп", "г", "аг", "п", "пос",
        "гп", "рп", "дер", "с/с", "р-н", "обл", "х", "мкр" };

    public static string FilePath { get { return path; } }

    public static void Start(string file) {
        path = file;
        searches = 0;
        try { File.WriteAllText(file, "", Encoding.UTF8); }
        catch (Exception) { path = null; }
        Line("=== Поиск абонента: запуск " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " ===");
        Line("В отчёте нет данных абонентов: только имена таблиц и полей, вид значений");
        Line("(буквы заменены на А/а, цифры на 9) и количества.");
        Line("Windows " + Environment.OSVersion.Version + ", .NET " + Environment.Version);
    }

    static readonly object sync = new object();

    public static void Line(string text) {
        if (path == null) return;
        lock (sync) {
            try { File.AppendAllText(path, text + "\r\n", Encoding.UTF8); }
            catch (Exception) { }
        }
    }

    public static void Search(string text) {
        if (++searches > 1000) return;
        Line(DateTime.Now.ToString("HH:mm:ss") + "  " + text);
    }

    // вид значения: "Заводская ул. 18а" -> "Ааааааааа ул. 99а"
    public static string Shape(string value, bool keepWords) {
        StringBuilder sb = new StringBuilder();
        foreach (string token in value.Split(' ')) {
            if (token.Length == 0) continue;
            if (sb.Length > 0) sb.Append(' ');
            string core = token.Trim('.', ',', ';', ':').ToLowerInvariant();
            if (keepWords && Array.IndexOf(Keep, core) >= 0) { sb.Append(token.ToLowerInvariant()); continue; }
            foreach (char c in token) sb.Append(char.IsDigit(c) ? '9' : char.IsLetter(c) ? (char.IsUpper(c) ? 'А' : 'а') : c);
            if (sb.Length > 60) { sb.Append("..."); break; }
        }
        return sb.ToString();
    }

    // список полей таблицы: тип, длина, сколько заполнено и как обычно выглядит значение
    public static void Fields(DbfTable t, string indent) {
        int sample = Math.Min(t.Count, 5000);
        for (int j = 0; j < t.Fields.Count; j++) {
            DbfField f = t.Fields[j];
            string label = FieldNames.Label(f.Name);
            string head = indent + (f.Name + "           ").Substring(0, 11) + (f.Type.ToString() + f.Length + "    ").Substring(0, 5)
                + (label == null ? "" : "(" + label + ") ");
            if (f.Type == 'M') { Line(head + "примечание (длинный текст)"); continue; }
            bool isName = string.Compare(f.Name, "FIO", StringComparison.OrdinalIgnoreCase) == 0;
            Dictionary<string, int> shapes = new Dictionary<string, int>();
            int filled = 0;
            for (int r = 0; r < sample; r++) {
                string v = DbfTable.Format(t.GetValue(r, j, null), true).Trim();
                if (v.Length == 0 || v == "0") continue;
                filled++;
                string sh = Shape(v, !isName);
                int n;
                shapes.TryGetValue(sh, out n);
                shapes[sh] = n + 1;
            }
            string top = "";
            int best = 0;
            foreach (KeyValuePair<string, int> kv in shapes) if (kv.Value > best) { best = kv.Value; top = kv.Key; }
            Line(head + "заполнено " + (sample == 0 ? 0 : filled * 100 / sample) + "%"
                + (top.Length == 0 ? "" : ", обычно: " + top + (shapes.Count > 1 ? "  (разных видов: " + shapes.Count + ")" : "")));
        }
    }
}

// Справочник улиц. У абонента в поле UL хранится код улицы. Название улицы берём из справочника
// улиц (у вас STREET.DBF), а населённый пункт и сельсовет - из справочников, на которые он ссылается
// своими кодами (IDNP -> таблица населённых пунктов, IDSS -> таблица сельсоветов). Их ищем сами:
// таблица, где есть поле с тем же именем, каждый код встречается один раз и при нём есть текст.
public class StreetDirectory {
    public string TableName;
    public string KeyField;
    public Dictionary<string, string> Links = new Dictionary<string, string>();   // поле-ссылка -> таблица (IDNP -> SPRNP.DBF)
    readonly Dictionary<string, string> names = new Dictionary<string, string>();
    readonly Dictionary<string, string> lowerNames = new Dictionary<string, string>();
    readonly Dictionary<string, string> ssOf = new Dictionary<string, string>();  // код улицы -> сельсовет
    readonly Dictionary<string, string> townOf = new Dictionary<string, string>();    // код улицы -> населённый пункт ("аг. Боровляны")
    readonly Dictionary<string, string> streetOf = new Dictionary<string, string>();  // код улицы -> своё название ("Заводская ул.")

    public int Count { get { return names.Count; } }

    public string Name(string code) {
        string s;
        return names.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public string LowerName(string code) {
        string s;
        return lowerNames.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    // для карточки: населённый пункт ("аг. Боровляны"), улица ("Заводская ул."), сельсовет
    public string Town(string code) {
        string s;
        return townOf.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public string Street(string code) {
        string s;
        return streetOf.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public string Selsovet(string code) {
        string s;
        return ssOf.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public double Coverage(List<string> codes) {
        if (codes.Count == 0) return 0;
        int hit = 0;
        foreach (string c in codes) if (names.ContainsKey(c)) hit++;
        return (double)hit / codes.Count;
    }

    // сельсоветы, к которым относятся улицы справочника
    public List<string> Selsovety {
        get {
            List<string> list = new List<string>();
            foreach (string v in ssOf.Values) if (v.Length > 0 && !list.Contains(v)) list.Add(v);
            list.Sort(StringComparer.CurrentCultureIgnoreCase);
            return list;
        }
    }

    // улица в одном из выбранных сельсоветов (selected = null - выбраны все; сельсовет неизвестен - показываем)
    public bool InSelsovety(string code, Dictionary<string, bool> selected) {
        if (selected == null) return true;
        string ss;
        if (!ssOf.TryGetValue(NormalizeCode(code), out ss) || ss.Length == 0) return true;
        return selected.ContainsKey(ss);
    }

    // Улицы, в названии которых (с населённым пунктом и сельсоветом) больше всего слов адреса
    public List<string> BestStreets(List<string> words, Dictionary<string, bool> selected, out int best) {
        best = 0;
        List<string> codes = new List<string>();
        List<string> norm = new List<string>();
        foreach (string w in words) norm.Add(w.ToLowerInvariant().Replace('ё', 'е'));
        foreach (KeyValuePair<string, string> kv in lowerNames) {
            int score = 0;
            foreach (string w in norm) if (kv.Value.IndexOf(w, StringComparison.Ordinal) >= 0) score++;
            if (score == 0 || score < best || !InSelsovety(kv.Key, selected)) continue;
            if (score > best) { best = score; codes.Clear(); }
            codes.Add(kv.Key);
        }
        return codes;
    }

    static string CleanSs(string ss) {
        string s = ss.Trim();
        foreach (string tail in new string[] { "с/с", "сельсовет", "с/с.", "сс" })
            if (s.EndsWith(" " + tail, StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - tail.Length - 1).Trim();
        return s;
    }

    public static string NormalizeCode(string s) {
        s = s.Trim();
        string t = s.TrimStart('0');
        return t.Length == 0 && s.Length > 0 ? "0" : t;
    }

    static bool IsDigits(string s) {
        if (s.Length == 0) return false;
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return true;
    }

    static bool HasLetter(string s) {
        foreach (char c in s) if (char.IsLetter(c)) return true;
        return false;
    }

    static string Value(DbfTable t, int row, int field) {
        return DbfTable.Format(t.GetValue(row, field, null), true).Trim();
    }

    // Коды улиц у абонентов, если в UL почти везде числа; иначе null (там названия)
    public static List<string> Codes(DbfTable abon, int ul) {
        if (ul < 0) return null;
        Dictionary<string, bool> codes = new Dictionary<string, bool>();
        int filled = 0, numeric = 0;
        for (int r = 0; r < abon.Count; r++) {
            string v = Value(abon, r, ul);
            if (v.Length == 0) continue;
            filled++;
            if (!IsDigits(v.Replace(",00", "").Replace(".00", ""))) continue;
            numeric++;
            codes[NormalizeCode(v)] = true;
        }
        Report.Line("Поле улицы (UL) в первых " + abon.Count + " записях: заполнено " + filled + ", из них чисел " + numeric + ", разных чисел " + codes.Count
            + (filled > 0 && numeric >= filled * 0.9 ? " -> в UL коды улиц, нужен справочник" : " -> в UL названия, справочник не нужен"));
        if (filled == 0 || numeric < filled * 0.9) return null;
        return new List<string>(codes.Keys);
    }

    // строка таблицы-ссылки: её текст и (на один уровень глубже) тексты её собственных ссылок
    class LinkRow {
        public string Text;
        public Dictionary<string, string> Nested = new Dictionary<string, string>();
    }

    // Где искать таблицы-ссылки: сначала запомненная с прошлого раза, потом все таблицы папки
    class TableSource {
        readonly string folder;
        readonly Dictionary<string, string> hints;
        List<TableInfo> all;
        public readonly Dictionary<string, string> Found = new Dictionary<string, string>();

        public TableSource(string folder, Dictionary<string, string> hints) {
            this.folder = folder;
            this.hints = hints;
        }

        public IEnumerable<TableInfo> Candidates(string field) {
            string name = null;
            if (hints != null && hints.TryGetValue(field.ToUpperInvariant(), out name) && File.Exists(Path.Combine(folder, name))) {
                TableInfo info = null;
                try {
                    info = new TableInfo();
                    info.Name = name;
                    info.Records = DbfTable.LoadHeader(Path.Combine(folder, name)).Expected;
                } catch (Exception) {
                    info = null;
                }
                if (info != null) yield return info;
            }
            if (all == null) all = TableInfo.List(folder, null);
            foreach (TableInfo info in all)
                if (name == null || string.Compare(info.Name, name, StringComparison.OrdinalIgnoreCase) != 0) yield return info;
        }
    }

    // Найти таблицу, где поле field - уникальный код с текстом при нём (например, IDNP в таблице населённых пунктов)
    static Dictionary<string, LinkRow> FindLink(string folder, TableSource tables, string field, List<string> needed,
                                                string exclude, bool nested, out string found) {
        found = null;
        foreach (TableInfo info in tables.Candidates(field)) {
            if (info.Records < 1 || info.Records > 50000) continue;
            if (string.Compare(info.Name, exclude, StringComparison.OrdinalIgnoreCase) == 0) continue;
            string path = Path.Combine(folder, info.Name);
            DbfTable t;
            int k;
            try {
                DbfTable h = DbfTable.LoadHeader(path);
                k = h.FieldIndex(field);
                if (k < 0) continue;
                bool text = false;
                foreach (DbfField f in h.Fields) if (f.Type == 'C' && f.Length >= 3) text = true;
                if (!text) continue;
                t = DbfTable.Load(path, false);
            } catch (Exception) {
                continue;
            }
            Dictionary<string, LinkRow> map = new Dictionary<string, LinkRow>();
            int filled = 0, withText = 0;
            for (int r = 0; r < t.Count; r++) {
                string code = NormalizeCode(Value(t, r, k));
                if (code.Length == 0 || code == "0") continue;
                filled++;
                if (map.ContainsKey(code)) continue;
                LinkRow row = new LinkRow();
                // сокращения ("аг.", "д.", "г.") - вперёд: "аг. Боровляны", а не "Боровляны аг."
                List<string> words = new List<string>();
                int front = 0;
                for (int j = 0; j < t.Fields.Count; j++) {
                    if (j == k || t.Fields[j].Type != 'C') continue;
                    string v = Value(t, r, j);
                    if (v.Length == 0 || !HasLetter(v)) continue;
                    if (v.Length <= 5 && v.EndsWith(".")) words.Insert(front++, v);
                    else words.Add(v);
                }
                row.Text = string.Join(" ", words.ToArray());
                if (row.Text.Length > 0) withText++;
                map[code] = row;
            }
            if (filled == 0 || map.Count < filled * 0.95 || withText < map.Count * 0.5) continue;   // код должен быть уникальным, с текстом
            int hit = 0;
            foreach (string c in needed) if (map.ContainsKey(c)) hit++;
            if (needed.Count > 0 && hit < needed.Count * 0.9) continue;                            // и покрывать почти все коды
            found = info.Name;
            tables.Found[field.ToUpperInvariant()] = info.Name;
            Report.Line("    поле " + field + " -> " + info.Name + " (записей " + t.Count + ", кодов нашлось " + hit + " из " + needed.Count + "). Поля:");
            Report.Fields(t, "        ");
            if (nested) {
                // ссылки самой этой таблицы (например, у населённого пункта - его сельсовет)
                for (int j = 0; j < t.Fields.Count; j++) {
                    if (j == k || t.Fields[j].Type != 'N') continue;
                    List<string> codes = new List<string>();
                    Dictionary<string, bool> seen = new Dictionary<string, bool>();
                    for (int r = 0; r < t.Count; r++) {
                        string c = NormalizeCode(Value(t, r, j));
                        if (c.Length > 0 && c != "0" && !seen.ContainsKey(c)) { seen[c] = true; codes.Add(c); }
                    }
                    if (codes.Count == 0) continue;
                    string inner;
                    Dictionary<string, LinkRow> sub = FindLink(folder, tables, t.Fields[j].Name, codes, info.Name, false, out inner);
                    if (sub == null) continue;
                    for (int r = 0; r < t.Count; r++) {
                        string code = NormalizeCode(Value(t, r, k));
                        LinkRow row, target;
                        if (!map.TryGetValue(code, out row)) continue;
                        if (sub.TryGetValue(NormalizeCode(Value(t, r, j)), out target) && target.Text.Length > 0)
                            row.Nested[t.Fields[j].Name.ToUpperInvariant()] = target.Text;
                    }
                }
            }
            return map;
        }
        return null;
    }

    public static StreetDirectory Load(string folder, string tableName, string keyField) {
        return Load(folder, tableName, keyField, null);
    }

    // links - запомненные с прошлого раза таблицы-ссылки (чтобы не перебирать все таблицы папки)
    public static StreetDirectory Load(string folder, string tableName, string keyField, Dictionary<string, string> links) {
        string path = Path.Combine(folder, tableName);
        DbfTable d = DbfTable.Load(path, false);
        int key = d.FieldIndex(keyField);
        if (key < 0) return null;
        StreetDirectory dir = new StreetDirectory();
        dir.TableName = Path.GetFileName(path);
        dir.KeyField = d.Fields[key].Name;
        Report.Line("Справочник улиц " + dir.TableName + ": код - поле " + dir.KeyField + ", записей " + d.Count + ". Поля:");
        Report.Fields(d, "    ");

        // числовые поля справочника - ссылки на другие справочники (населённый пункт, сельсовет)
        TableSource tables = new TableSource(folder, links);
        Dictionary<int, Dictionary<string, LinkRow>> found = new Dictionary<int, Dictionary<string, LinkRow>>();
        for (int j = 0; j < d.Fields.Count; j++) {
            if (j == key || d.Fields[j].Type != 'N') continue;
            List<string> codes = new List<string>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>();
            for (int r = 0; r < d.Count; r++) {
                string c = NormalizeCode(Value(d, r, j));
                if (c.Length > 0 && c != "0" && !seen.ContainsKey(c)) { seen[c] = true; codes.Add(c); }
            }
            if (codes.Count == 0) continue;
            string table;
            Dictionary<string, LinkRow> map = FindLink(folder, tables, d.Fields[j].Name, codes, dir.TableName, true, out table);
            if (map != null) found[j] = map;
            else Report.Line("    поле " + d.Fields[j].Name + ": справочник к нему не нашёлся");
        }
        dir.Links = tables.Found;

        for (int r = 0; r < d.Count; r++) {
            string code = NormalizeCode(Value(d, r, key));
            if (code.Length == 0 || dir.names.ContainsKey(code)) continue;
            // своё: название улицы
            StringBuilder own = new StringBuilder();
            for (int j = 0; j < d.Fields.Count; j++) {
                if (j == key || d.Fields[j].Type != 'C') continue;
                string v = Value(d, r, j);
                if (v.Length == 0 || !HasLetter(v)) continue;
                if (own.Length > 0) own.Append(", ");
                own.Append(v);
            }
            // по ссылкам: населённый пункт, сельсовет и прочее
            Dictionary<string, string> parts = new Dictionary<string, string>();
            List<string> order = new List<string>();
            foreach (KeyValuePair<int, Dictionary<string, LinkRow>> link in found) {
                LinkRow row;
                if (!link.Value.TryGetValue(NormalizeCode(Value(d, r, link.Key)), out row) || row.Text.Length == 0) continue;
                string name = d.Fields[link.Key].Name.ToUpperInvariant();
                parts[name] = row.Text;
                order.Add(name);
                foreach (KeyValuePair<string, string> n in row.Nested)
                    if (!parts.ContainsKey(n.Key)) { parts[n.Key] = n.Value; order.Add(n.Key); }
            }
            // если своя ссылка пустая, а у населённого пункта она есть - берём оттуда (сельсовет улицы через её деревню)
            string town = "", ss = "", other = "";
            foreach (string name in order) {
                string text = parts[name];
                if (name.Contains("NP")) { if (town.Length == 0) town = text; }
                else if (name.Contains("SS")) { if (ss.Length == 0) ss = text; }
                else other += (other.Length > 0 ? ", " : "") + text;
            }
            StringBuilder display = new StringBuilder();
            if (town.Length > 0) display.Append(town);
            if (own.Length > 0) display.Append((display.Length > 0 ? ", " : "") + own);
            if (ss.Length > 0) {
                string l = ss.ToLowerInvariant();
                display.Append((display.Length > 0 ? ", " : "") + ss + (l.Contains("с/с") || l.Contains("сельсовет") ? "" : " с/с"));
            }
            if (other.Length > 0) display.Append((display.Length > 0 ? ", " : "") + other);
            dir.names[code] = display.ToString();
            dir.lowerNames[code] = display.ToString().ToLowerInvariant().Replace('ё', 'е');
            dir.ssOf[code] = CleanSs(ss);
            dir.townOf[code] = town;
            dir.streetOf[code] = own.ToString();
        }
        return dir;
    }

    // ---- точный адрес ----

    class NameEntry {
        public string[] Words;
        public string Key;
        public List<string> Codes = new List<string>();
    }

    List<NameEntry> towns;                                        // населённые пункты: слова названия -> их улицы
    Dictionary<string, string> streetKey;                    // код улицы -> слова названия улицы
    Dictionary<string, NameEntry> streets;                        // название улицы -> улицы с таким названием (во всех нас. пунктах)

    void PrepareMatch() {
        lock (names) {
            if (towns != null) return;
            Dictionary<string, NameEntry> t = new Dictionary<string, NameEntry>();
            streets = new Dictionary<string, NameEntry>();
            streetKey = new Dictionary<string, string>();
            foreach (string code in names.Keys) {
                string town;
                if (townOf.TryGetValue(code, out town) && town.Length > 0) Add(t, AddressText.NameWords(town), code);
                string own;
                string[] sw = streetOf.TryGetValue(code, out own) ? AddressText.NameWords(own) : new string[0];
                streetKey[code] = string.Join(" ", sw);
                Add(streets, sw, code);
            }
            towns = new List<NameEntry>(t.Values);
        }
    }

    static void Add(Dictionary<string, NameEntry> map, string[] words, string code) {
        if (words.Length == 0) return;
        string key = string.Join(" ", words);
        NameEntry n;
        if (!map.TryGetValue(key, out n)) { n = new NameEntry(); n.Words = words; n.Key = key; map[key] = n; }
        n.Codes.Add(code);
    }

    // слова, которые могут быть названием: подряд, ещё не узнанные
    static List<int> NameSlots(List<AddrToken> toks) {
        List<int> list = new List<int>();
        for (int i = 0; i < toks.Count; i++) if (toks[i].Kind == 'w' && !toks[i].Ss && !toks[i].Bound) list.Add(i);
        return list;
    }

    // где в адресе стоит название целиком (все его слова подряд); -1 - нет
    static int Find(List<AddrToken> toks, List<int> slots, string[] words, bool[] used) {
        for (int p = 0; p + words.Length <= slots.Count; p++) {
            bool ok = true;
            for (int k = 0; k < words.Length && ok; k++) {
                int i = slots[p + k];
                if (used[p + k] || toks[i].Text != words[k]) ok = false;
                if (k > 0 && slots[p + k] - slots[p + k - 1] > 2) ok = false;   // между словами - не больше одного служебного ("аг.", "ул.")
            }
            if (ok) return p;
        }
        return -1;
    }

    class Reading {
        public NameEntry Town, Street;
        public int TownAt = -1, StreetAt = -1;
        public double Score;
    }

    // сколько слов узнано; рядом стоящее "ул." или "аг." подсказывает, улица это или населённый пункт
    static double Score(Reading r, List<AddrToken> toks, List<int> slots) {
        double score = 0;
        if (r.Town != null) {
            score += r.Town.Words.Length;
            if (Near(toks, slots[r.TownAt], slots[r.TownAt + r.Town.Words.Length - 1], 's')) score -= 0.5;
        }
        if (r.Street != null) {
            score += r.Street.Words.Length;
            if (Near(toks, slots[r.StreetAt], slots[r.StreetAt + r.Street.Words.Length - 1], 't')) score -= 0.5;
        }
        return score;
    }

    static bool Near(List<AddrToken> toks, int first, int last, char kind) {
        return (first > 0 && toks[first - 1].Kind == kind) || (last + 1 < toks.Count && toks[last + 1].Kind == kind);
    }

    // самое длинное название улицы из списка, которое есть в адресе
    static void BestStreet(List<AddrToken> toks, List<int> slots, bool[] used, IEnumerable<NameEntry> list, Reading r) {
        foreach (NameEntry n in list) {
            if (r.Street != null && n.Words.Length <= r.Street.Words.Length) continue;
            int p = Find(toks, slots, n.Words, used);
            if (p >= 0) { r.Street = n; r.StreetAt = p; }
        }
    }

    // Адрес как его вписали -> улицы (коды), дом, корпус, квартира.
    // Сначала населённый пункт, потом улица именно в нём, потом дом, корпус, квартира.
    public AddressMatch Match(string query) {
        PrepareMatch();
        AddressMatch m = new AddressMatch();
        List<AddrToken> toks = AddressText.Tokens(query);
        List<int> slots = NameSlots(toks);

        // все прочтения: каждый найденный в адресе населённый пункт + улица в нём; или только улица
        List<Reading> readings = new List<Reading>();
        Reading none = new Reading();
        BestStreet(toks, slots, new bool[slots.Count], streets.Values, none);
        readings.Add(none);
        foreach (NameEntry town in towns) {
            int p = Find(toks, slots, town.Words, new bool[slots.Count]);
            if (p < 0) continue;
            Reading r = new Reading();
            r.Town = town;
            r.TownAt = p;
            bool[] used = new bool[slots.Count];
            for (int k = 0; k < town.Words.Length; k++) used[p + k] = true;
            List<NameEntry> inTown = new List<NameEntry>();
            Dictionary<string, NameEntry> byKey = new Dictionary<string, NameEntry>();
            foreach (string code in town.Codes) {
                string key = streetKey[code];
                if (key.Length == 0) continue;
                NameEntry n;
                if (!byKey.TryGetValue(key, out n)) { n = new NameEntry(); n.Key = key; n.Words = key.Split(' '); byKey[key] = n; inTown.Add(n); }
                n.Codes.Add(code);
            }
            BestStreet(toks, slots, used, inTown, r);
            readings.Add(r);
        }
        // лучшее прочтение: узнано больше слов; при равенстве - с населённым пунктом, потом - что раньше в строке
        foreach (Reading r in readings) r.Score = Score(r, toks, slots);
        Reading best = none;
        foreach (Reading r in readings) {
            if (r == none) continue;
            if (r.Score > best.Score || (r.Score == best.Score && (best.Town == null || r.TownAt < best.TownAt))) best = r;
        }
        if (best.Town != null) for (int k = 0; k < best.Town.Words.Length; k++) toks[slots[best.TownAt + k]].Used = true;
        if (best.Street != null) for (int k = 0; k < best.Street.Words.Length; k++) toks[slots[best.StreetAt + k]].Used = true;

        // сельсовет: перед "с/с" или просто его название среди оставшихся слов
        List<string> ssWanted = new List<string>();
        List<string> leftover = new List<string>();
        int firstName = int.MaxValue;
        if (best.Town != null) firstName = Math.Min(firstName, slots[best.TownAt]);
        if (best.Street != null) firstName = Math.Min(firstName, slots[best.StreetAt]);
        for (int i = 0; i < toks.Count; i++) {
            AddrToken t = toks[i];
            if (t.Kind != 'w' || t.Used || t.Digit || t.Bound) continue;
            string ss = SsByWord(t.Text);
            if (ss != null) { if (!ssWanted.Contains(ss)) ssWanted.Add(ss); t.Used = true; continue; }
            if (t.Ss) { leftover.Add(t.Text); continue; }
            if (t.Text.Length == 1 && i > 0 && toks[i - 1].Digit && !toks[i - 1].Used) continue;   // буква дома: "18 а"
            leftover.Add(t.Text);
        }

        // дом, корпус, квартира
        string[] slot = new string[3];   // 0 - дом, 1 - корпус, 2 - квартира
        List<int> bare = new List<int>();
        List<string> values = new List<string>();
        int pending = -1, lastValue = -1;
        bool hasName = best.Town != null || best.Street != null;
        for (int i = 0; i < toks.Count; i++) {
            AddrToken t = toks[i];
            if (t.Kind == 'h') { pending = 0; continue; }
            if (t.Kind == 'k') { pending = 1; continue; }
            if (t.Kind == 'f') { pending = 2; continue; }
            if (t.Kind != 'w' || t.Used) continue;
            if (!t.Digit) {
                // буква дома, написанная отдельно: "18 а"
                if (t.Text.Length == 1 && lastValue >= 0 && i > 0 && toks[i - 1].Digit) values[lastValue] += t.Text;
                continue;
            }
            bool longNumber = t.Text.Length >= 5 && IsDigits(t.Text);
            if (longNumber || (!t.Bound && pending < 0 && hasName && i < firstName)) { pending = -1; lastValue = -1; continue; }   // лицевой, индекс, номер строки списка
            values.Add(t.Text);
            lastValue = values.Count - 1;
            if (pending >= 0 && slot[pending] == null) slot[pending] = "#" + lastValue;
            else bare.Add(lastValue);
            pending = -1;
        }
        // числа без пояснений: первое - дом, второе - квартира, третье - корпус
        int[] order = { 0, 2, 1 };
        foreach (int b in bare) foreach (int o in order) if (slot[o] == null) { slot[o] = "#" + b; break; }
        m.House = slot[0] == null ? null : values[int.Parse(slot[0].Substring(1))];
        m.Korp = slot[1] == null ? null : values[int.Parse(slot[1].Substring(1))];
        m.Flat = slot[2] == null ? null : values[int.Parse(slot[2].Substring(1))];

        // улицы, где искать
        string townText = best.Town == null ? null : TownLabel(best.Town.Codes);
        if (best.Town != null && best.Street != null) m.Codes.AddRange(best.Street.Codes);
        else if (best.Town != null) {
            if (leftover.Count > 0) {
                m.Problem = "В " + townText + " нет улицы «" + string.Join(" ", leftover.ToArray()) + "»" + Elsewhere(leftover) + ".";
                m.Explain = "нас. пункт узнан, улица не найдена в нём";
                return m;
            }
            m.Codes.AddRange(best.Town.Codes);
        } else if (best.Street != null) m.Codes.AddRange(best.Street.Codes);
        else {
            m.Problem = leftover.Count > 0 ? "Не нашёл в справочниках улиц и населённых пунктов: «" + string.Join(" ", leftover.ToArray()) + "»."
                : m.House != null ? "Впишите улицу или населённый пункт, не только номер дома." : "Впишите адрес.";
            m.Explain = leftover.Count > 0 ? "не узнано ни нас. пункта, ни улицы" : "нет названий";
            return m;
        }
        if (leftover.Count > 0) {
            m.Codes.Clear();
            m.Problem = "Не понял в адресе: «" + string.Join(" ", leftover.ToArray()) + "» - такого нет в справочниках" + (best.Town != null ? " для " + townText : "") + ".";
            m.Explain = "лишние неузнанные слова: " + leftover.Count;
            return m;
        }
        if (ssWanted.Count > 0) {
            List<string> keep = new List<string>();
            foreach (string c in m.Codes) { string s; if (ssOf.TryGetValue(c, out s) && ssWanted.Contains(s)) keep.Add(c); }
            if (keep.Count == 0) {
                m.Codes.Clear();
                m.Problem = "В " + string.Join(", ", ssWanted.ToArray()) + " с/с такого адреса нет.";
                return m;
            }
            m.Codes = keep;
        }

        StringBuilder said = new StringBuilder();
        if (townText != null) said.Append(townText);
        if (best.Street != null) said.Append((said.Length > 0 ? " -> " : "") + StreetLabel(m.Codes, best.Town == null));
        else said.Append(" -> все улицы");
        if (m.House != null) said.Append(" -> дом " + m.House);
        if (m.Korp != null) said.Append(", корп. " + m.Korp);
        if (m.Flat != null) said.Append(", кв. " + m.Flat);
        m.Said = said.ToString();
        m.Explain = "нас. пункт: " + (best.Town == null ? "не указан" : "узнан (вариантов " + TownCount(best.Town.Codes) + ")")
            + ", улица: " + (best.Street == null ? (best.Town != null ? "не указана - весь нас. пункт" : "нет") : "узнана")
            + (ssWanted.Count > 0 ? ", сельсовет указан" : "") + ", улиц для поиска: " + m.Codes.Count
            + ", дом: " + (m.House != null ? "да" : "нет") + ", корпус: " + (m.Korp != null ? "да" : "нет") + ", квартира: " + (m.Flat != null ? "да" : "нет");
        return m;
    }

    string SsByWord(string w) {
        if (w.Length < 5) return null;
        string stem = Stem(w);
        foreach (string s in ssOf.Values) if (s.Length > 0 && Stem(s).Replace('ё', 'е') == stem) return s;
        return null;
    }

    // сколько разных населённых пунктов (одноимённые в разных сельсоветах - разные)
    int TownCount(List<string> codes) {
        Dictionary<string, bool> seen = new Dictionary<string, bool>();
        foreach (string c in codes) {
            string t, s;
            townOf.TryGetValue(c, out t);
            ssOf.TryGetValue(c, out s);
            seen[t + "|" + s] = true;
        }
        return seen.Count;
    }

    string TownLabel(List<string> codes) {
        List<string> list = new List<string>();
        foreach (string c in codes) {
            string t, s;
            townOf.TryGetValue(c, out t);
            ssOf.TryGetValue(c, out s);
            string label = t + (s != null && s.Length > 0 ? " (" + s + " с/с)" : "");
            if (!list.Contains(label)) list.Add(label);
        }
        return list.Count == 1 ? list[0].Substring(0, list[0].IndexOf(" (") > 0 ? list[0].IndexOf(" (") : list[0].Length) : string.Join(" или ", list.ToArray());
    }

    string StreetLabel(List<string> codes, bool withTown) {
        List<string> list = new List<string>();
        foreach (string c in codes) {
            string st, t;
            streetOf.TryGetValue(c, out st);
            townOf.TryGetValue(c, out t);
            string label = withTown && t != null && t.Length > 0 ? st + " (" + t + ")" : st;
            if (!list.Contains(label)) list.Add(label);
        }
        if (list.Count > 4) return list[0] + " и ещё " + (list.Count - 1);
        return string.Join(", ", list.ToArray());
    }

    // где ещё есть улица с таким названием - подсказка
    string Elsewhere(List<string> words) {
        NameEntry n;
        if (!streets.TryGetValue(string.Join(" ", words.ToArray()), out n)) return "";
        List<string> list = new List<string>();
        foreach (string c in n.Codes) { string t; if (townOf.TryGetValue(c, out t) && t.Length > 0 && !list.Contains(t)) list.Add(t); }
        if (list.Count == 0) return "";
        return " (такая улица есть в: " + string.Join(", ", list.GetRange(0, Math.Min(5, list.Count)).ToArray()) + (list.Count > 5 ? "..." : "") + ")";
    }

    // Коды улиц, в адресе которых есть одно из названий (сельсоветы). "Боровлянский" ищется как "боровлянск".
    public Dictionary<string, bool> CodesMatching(string[] wanted, out string report) {
        List<string> stems = new List<string>();
        foreach (string n in wanted) stems.Add(Stem(n));
        int[] perName = new int[stems.Count];
        Dictionary<string, bool> codes = new Dictionary<string, bool>();
        foreach (KeyValuePair<string, string> kv in lowerNames) {
            for (int i = 0; i < stems.Count; i++) {
                if (stems[i].Length == 0 || kv.Value.IndexOf(stems[i], StringComparison.Ordinal) < 0) continue;
                perName[i]++;
                codes[kv.Key] = true;
                break;
            }
        }
        StringBuilder r = new StringBuilder("Сельсоветы из списка - улиц в справочнике: ");
        for (int i = 0; i < stems.Count; i++) r.Append((i > 0 ? ", " : "") + wanted[i].Trim() + " - " + perName[i]);
        r.Append(". Всего улиц: " + codes.Count + " из " + names.Count);
        report = r.ToString();
        return codes;
    }

    public static string Stem(string name) {
        string w = name.Trim().ToLowerInvariant();
        if (w.Length > 5 && (w.EndsWith("ий") || w.EndsWith("ой") || w.EndsWith("ая") || w.EndsWith("ое"))) w = w.Substring(0, w.Length - 2);
        return w;
    }

    public bool HasTowns {
        get { foreach (string t in townOf.Values) if (t.Length > 0) return true; return false; }
    }

    class Candidate {
        public string Table, Key;
        public double Coverage;
        public int Records;
        public bool Ss, Np;
    }

    // Ищем таблицу, где есть поле с теми же кодами (каждый код - один раз) и текстом при них.
    // Берём, только если там нашлись почти все коды улиц абонентов. Из подходящих лучше та, у которой
    // есть ссылки на сельсовет и населённый пункт (STREET.DBF -> SPRSS, SPRNP), а не просто список улиц.
    public static StreetDirectory Detect(string folder, string abonTable, List<string> codes, LoadProgress progress) {
        List<TableInfo> tables = TableInfo.List(folder, null);
        List<Candidate> found = new List<Candidate>();
        // сначала - только поля кода улицы с "UL" в имени (KUL, IDUL): так не читаются все таблицы подряд
        for (int pass = 0; pass < 2 && found.Count == 0; pass++) {
            for (int i = 0; i < tables.Count; i++) {
                if (progress != null) progress(i + pass * tables.Count, tables.Count * 2);
                TableInfo info = tables[i];
                if (info.Abonents || info.Records < 2 || info.Records > 50000) continue;
                if (string.Compare(info.Name, abonTable, StringComparison.OrdinalIgnoreCase) == 0) continue;
                if (info.Name.ToUpperInvariant().StartsWith("ARC")) continue;
                string path = Path.Combine(folder, info.Name);
                DbfTable d;
                try {
                    DbfTable h = DbfTable.LoadHeader(path);
                    bool text = false, ulField = false;
                    foreach (DbfField f in h.Fields) {
                        if (f.Type == 'C' && f.Length >= 5) text = true;
                        if ((f.Type == 'N' || f.Type == 'C') && f.Length <= 10 && f.Name.ToUpperInvariant().Contains("UL")) ulField = true;
                    }
                    if (!text || (pass == 0 && !ulField)) continue;
                    d = DbfTable.Load(path, false);
                } catch (Exception) {
                    continue;
                }
                if (d.Count < 2) continue;
                for (int k = 0; k < d.Fields.Count; k++) {
                    DbfField f = d.Fields[k];
                    if (!((f.Type == 'N' && f.Length <= 10) || (f.Type == 'C' && f.Length <= 10))) continue;
                    if (pass == 0 && !f.Name.ToUpperInvariant().Contains("UL")) continue;
                    Dictionary<string, bool> values = new Dictionary<string, bool>();
                    int filled = 0;
                    for (int r = 0; r < d.Count; r++) {
                        string v = NormalizeCode(Value(d, r, k));
                        if (v.Length == 0) continue;
                        filled++;
                        values[v] = true;
                    }
                    if (filled < 2 || values.Count < filled * 0.95) continue;   // код справочника не повторяется
                    int hit = 0;
                    foreach (string c in codes) if (values.ContainsKey(c)) hit++;
                    double coverage = (double)hit / codes.Count;
                    Candidate cand = new Candidate();
                    cand.Table = info.Name;
                    cand.Key = f.Name;
                    cand.Coverage = coverage;
                    cand.Records = d.Count;
                    foreach (DbfField g in d.Fields) {
                        string n = g.Name.ToUpperInvariant();
                        if (g.Type != 'N' || g == f) continue;
                        if (n.Contains("SS")) cand.Ss = true;
                        if (n.Contains("NP")) cand.Np = true;
                    }
                    if (hit > 0)
                        Report.Line("    кандидат " + info.Name + "." + f.Name + ": записей " + d.Count + ", кодов улиц нашлось " + hit + " из " + codes.Count
                            + " (" + (int)(coverage * 100) + "%)" + (coverage >= 0.9 ? (cand.Ss ? ", есть ссылка на сельсовет" : "") + (cand.Np ? ", есть ссылка на населённый пункт" : "") : " - мало, не подходит"));
                    if (coverage >= 0.9) found.Add(cand);
                }
            }
        }
        if (found.Count == 0) {
            Report.Line("    справочник не найден: нет таблицы, где нашлись бы 90% кодов улиц");
            return null;
        }
        found.Sort(delegate(Candidate x, Candidate y) {
            if (x.Ss != y.Ss) return x.Ss ? -1 : 1;
            if (x.Np != y.Np) return x.Np ? -1 : 1;
            if (x.Coverage != y.Coverage) return y.Coverage.CompareTo(x.Coverage);
            return y.Records.CompareTo(x.Records);
        });
        // берём лучший; если в нём не оказалось ни населённых пунктов, ни сельсоветов - следующий
        StreetDirectory first = null;
        for (int i = 0; i < found.Count && i < 3; i++) {
            Report.Line("    пробую " + found[i].Table + "." + found[i].Key);
            StreetDirectory dir = Load(folder, found[i].Table, found[i].Key);
            if (dir == null) continue;
            if (first == null) first = dir;
            if (dir.HasTowns || dir.Selsovety.Count > 0) { Report.Line("    выбран " + found[i].Table + "." + found[i].Key); return dir; }
        }
        if (first != null) Report.Line("    выбран " + first.TableName + "." + first.KeyField + " (населённых пунктов и сельсоветов в нём нет)");
        return first;
    }
}

public enum SearchMode { Account, Address, Meter, Name }

// Поиск абонента: по лицевому счёту, адресу, номеру счётчика или ФИО - что выбрано
public class AbonentQuery {
    readonly DbfTable t;
    readonly int[] accountFields, addressFields, meterFields, nameFields;
    readonly int oldMeter;   // снятый счётчик (NSC) - ищем по нему, если по текущему не нашлось
    readonly int dom, kor, kv, ul;
    public readonly int[] Columns;   // что показывать в списке найденного
    public string LastExplain = "";  // как прошёл последний поиск - для отчёта (без самих слов)
    readonly string[] street;        // адрес по справочнику улиц для каждой записи (если справочник есть)
    readonly string[] streetLower;
    readonly string streetSource;

    public AbonentQuery(DbfTable t, StreetDirectory dir) {
        this.t = t;
        accountFields = Existing(new string[] { "NAB", "LS", "LIC", "LSCH", "LS_ERIP" });
        meterFields = Existing(new string[] { "NCH" });
        oldMeter = t.FieldIndex("NSC");
        nameFields = Existing(new string[] { "FIO" });
        dom = t.FieldIndex("DOM");
        kor = t.FieldIndex("KOR");
        kv = t.FieldIndex("KV");
        ul = t.FieldIndex("UL");

        // слова адреса (населённый пункт, улица) ищем во всех текстовых полях абонента,
        // кроме ФИО и номеров - где бы ни хранился населённый пункт
        string[] notAddress = { "FIO", "NAB", "LS", "LIC", "LSCH", "NCH", "NSC", "TIPCH", "TSC", "TEL", "DOM", "KOR", "KV", "N_F", "LS_ERIP" };
        List<int> address = new List<int>();
        if (dir != null) {
            // улицы - по справочнику; из полей абонента берём только явно адресные
            address.AddRange(Existing(new string[] { "UL", "NP", "NASP", "GOROD", "ULICA", "ADRES", "ADR" }));
        } else {
            for (int i = 0; i < t.Fields.Count; i++) {
                bool skip = false;
                foreach (string n in notAddress) if (string.Compare(t.Fields[i].Name, n, StringComparison.OrdinalIgnoreCase) == 0) skip = true;
                if (!skip && (t.Fields[i].Type == 'C' || i == ul)) address.Add(i);
            }
        }
        addressFields = address.ToArray();

        if (dir != null && ul >= 0) {
            streetSource = dir.TableName;
            street = new string[t.Count];
            streetLower = new string[t.Count];
            for (int r = 0; r < t.Count; r++) {
                string code = DbfTable.Format(t.GetValue(r, ul, null), true);
                street[r] = dir.Name(code);
                streetLower[r] = dir.LowerName(code);
            }
        }
        Columns = Existing(new string[] { "NAB", "FIO", "UL", "DOM", "KOR", "KV", "NCH", "TIPCH", "TEL" });
        if (Columns.Length == 0) {
            Columns = new int[Math.Min(8, t.Fields.Count)];
            for (int i = 0; i < Columns.Length; i++) Columns[i] = i;
        }
    }

    public bool HasStreets { get { return street != null; } }

    // Записи, среди которых ищем: null - вся таблица, иначе только эти (выбранные сельсоветы)
    public int[] Scope;

    int RowCount { get { return Scope == null ? t.Count : Scope.Length; } }

    int RowAt(int i) { return Scope == null ? i : Scope[i]; }

    // Абоненты, у которых в адресе (справочник улиц или текстовые поля) есть одно из названий.
    // "Боровлянский" ищется как "боровлянск" - находится и "Боровлянский с/с", и "Боровлянского".
    public int[] FindByNames(string[] names, out string report) {
        List<string> stems = new List<string>();
        foreach (string n in names) {
            string w = n.Trim().ToLowerInvariant();
            if (w.Length > 5 && (w.EndsWith("ий") || w.EndsWith("ой") || w.EndsWith("ая") || w.EndsWith("ое"))) w = w.Substring(0, w.Length - 2);
            if (w.Length > 0) stems.Add(w);
        }
        SearchWord[] words = new SearchWord[stems.Count];
        for (int i = 0; i < words.Length; i++) words[i] = Part(stems[i]);
        int[] perName = new int[stems.Count];
        Dictionary<string, int> sources = new Dictionary<string, int>();
        List<int> rows = new List<int>();
        for (int row = 0; row < t.Count; row++) {
            for (int i = 0; i < words.Length; i++) {
                string source = null;
                if (streetLower != null && streetLower[row] != null && streetLower[row].IndexOf(stems[i], StringComparison.Ordinal) >= 0) source = "справочник улиц";
                else foreach (int f in addressFields) if (t.FieldContains(row, f, words[i])) { source = "поле " + t.Fields[f].Name; break; }
                if (source == null) continue;
                perName[i]++;
                int c;
                sources.TryGetValue(source, out c);
                sources[source] = c + 1;
                rows.Add(row);
                break;
            }
        }
        StringBuilder r = new StringBuilder();
        r.Append("Сельсоветы из списка: ");
        for (int i = 0; i < stems.Count; i++) r.Append((i > 0 ? ", " : "") + names[i].Trim() + " - " + perName[i]);
        r.Append(". Всего абонентов в них: " + rows.Count + " из " + t.Count + ". Где нашлись названия: ");
        if (sources.Count == 0) r.Append("нигде");
        bool first = true;
        foreach (KeyValuePair<string, int> kv in sources) { r.Append((first ? "" : ", ") + kv.Key + " - " + kv.Value); first = false; }
        report = r.ToString();
        return rows.ToArray();
    }

    public string FieldList(SearchMode mode) {
        int[] fields = mode == SearchMode.Account ? accountFields : mode == SearchMode.Address ? addressFields
            : mode == SearchMode.Meter ? meterFields : nameFields;
        List<string> names = new List<string>();
        foreach (int f in fields) names.Add(t.Fields[f].Name);
        string list = names.Count == 0 ? "вся запись (нужных полей нет)" : string.Join(", ", names.ToArray());
        if (mode == SearchMode.Address) {
            list += "; дом: " + (dom < 0 ? "нет поля" : t.Fields[dom].Name) + ", корпус: " + (kor < 0 ? "нет поля" : t.Fields[kor].Name)
                + ", квартира: " + (kv < 0 ? "нет поля" : t.Fields[kv].Name) + (street != null ? "; плюс справочник улиц" : "");
        }
        return list;
    }

    // адрес по справочнику для записи или null
    public string StreetOf(int row) {
        return street == null ? null : street[row];
    }

    // текст для списка: улица по справочнику вместо кода
    public string CellText(int row, int field) {
        string value = DbfTable.Format(t.GetValue(row, field, null), false);
        if (!IsStreetColumn(field)) return value;
        string s = street[row];
        return s == null ? "код " + value : s;
    }

    public string ColumnTitle(int field) {
        if (IsStreetColumn(field)) return "Населённый пункт, улица";
        string label = FieldNames.Label(t.Fields[field].Name);
        return label == null ? t.Fields[field].Name : label;
    }

    // дополнительная строка карточки: адрес по справочнику
    public string[] CardExtra(int row) {
        string s = StreetOf(row);
        return s == null ? null : new string[] { "Адрес по справочнику " + streetSource, s };
    }

    public bool IsStreetColumn(int field) {
        return field == ul && street != null;
    }

    int[] Existing(string[] names) {
        List<int> list = new List<int>();
        foreach (string n in names) {
            int i = t.FieldIndex(n);
            if (i >= 0) list.Add(i);
        }
        return list.ToArray();
    }

    public bool Partial;            // частичное совпадение; иначе - только точное
    public AddressMatch Address;    // точный адрес, узнанный по справочнику (иначе адрес ищется по словам)
    public string Miss = "";        // на чём адрес перестал совпадать ("дома 999 на этой улице нет")

    public List<int> Find(string query, SearchMode mode) {
        List<int> found;
        switch (mode) {
            case SearchMode.Account: found = Partial ? FindAccount(query) : FindAccountExact(query); break;
            case SearchMode.Address: found = !Partial && Address != null ? FindAddressExact() : FindAddress(query); break;
            case SearchMode.Meter: found = Partial ? FindMeter(query) : FindMeterExact(query); break;
            default: found = Partial ? FindWords(query, nameFields) : FindNameExact(query); break;
        }
        if (found.Count <= 20000) SortByAddress(found);
        return found;
    }

    // ---- точное совпадение ----

    string Text(int row, int field) {
        return field < 0 ? "" : DbfTable.Format(t.GetValue(row, field, null), true).Trim();
    }

    // лицевой: номер равен вписанному целиком (нули впереди не важны)
    List<int> FindAccountExact(string query) {
        List<int> found = new List<int>();
        StringBuilder digits = new StringBuilder();
        foreach (char c in query) if (c >= '0' && c <= '9') digits.Append(c);
        LastExplain = "цифр: " + digits.Length + ", только точное совпадение";
        if (digits.Length == 0) return found;
        string core = digits.ToString().TrimStart('0');
        if (core.Length == 0) core = "0";
        for (int i = 0; i < RowCount; i++) {
            int row = RowAt(i);
            foreach (int f in accountFields) {
                string v = Text(row, f);
                if (v.Length == 0) continue;
                string vc = v.TrimStart('0');
                if ((vc.Length == 0 ? "0" : vc) == core) { found.Add(row); break; }
            }
        }
        return found;
    }

    // номер счётчика для сравнения: только буквы и цифры, маленькие, без нулей впереди ("№ 0050-001" -> 50001)
    public static string MeterKey(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s.ToLowerInvariant()) if (char.IsLetterOrDigit(c)) sb.Append(c);
        string r = sb.ToString().TrimStart('0');
        return r.Length == 0 && sb.Length > 0 ? "0" : r;
    }

    // номер из запроса: весь запрос и каждое его слово с цифрами ("СЕ-208BY 10011761" -> и 10011761)
    public static List<string> MeterKeys(string query) {
        List<string> keys = new List<string>();
        string q = query.Replace("№", " ");
        string all = MeterKey(q);
        if (all.Length > 0) keys.Add(all);
        foreach (string w in q.Split(new char[] { ' ', '\t', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
            int digits = 0;
            foreach (char c in w) if (char.IsDigit(c)) digits++;
            string k = MeterKey(w);
            if (digits >= 3 && k.Length >= 4 && !keys.Contains(k)) keys.Add(k);
        }
        return keys;
    }

    // счётчик: номер совпадает целиком; не нашёлся среди текущих - ищем среди снятых (NSC)
    List<int> FindMeterExact(string query) {
        List<int> found = new List<int>();
        List<string> keys = MeterKeys(query);
        LastExplain = "вариантов номера: " + keys.Count + ", только точное совпадение";
        if (keys.Count == 0) return found;
        foreach (int f in meterFields)
            for (int i = 0; i < RowCount; i++) { int row = RowAt(i); if (!found.Contains(row) && keys.Contains(MeterKey(Text(row, f)))) found.Add(row); }
        if (found.Count == 0 && oldMeter >= 0) {
            for (int i = 0; i < RowCount; i++) { int row = RowAt(i); if (keys.Contains(MeterKey(Text(row, oldMeter)))) found.Add(row); }
            if (found.Count > 0) LastExplain += ", найден по снятому счётчику (NSC)";
        }
        return found;
    }

    public static List<string> PersonWords(string s) {
        List<string> list = new List<string>();
        foreach (string w in s.ToLowerInvariant().Replace('ё', 'е').Split(new char[] { ' ', '.', ',', '\t' }, StringSplitOptions.RemoveEmptyEntries)) list.Add(w);
        return list;
    }

    // в одном поле ФИО бывает несколько человек: "ИВАНОВ И.И./ПЕТРОВА П.П."
    static readonly char[] PersonSeparators = { '/', '\\', ';', '+' };

    // частичное совпадение: каждое слово - часть ФИО
    public static bool ContainsAll(string fio, List<string> words) {
        string f = fio.ToLowerInvariant().Replace('ё', 'е');
        foreach (string w in words) if (f.IndexOf(w, StringComparison.Ordinal) < 0) return false;
        return true;
    }

    // ФИО подходит, если у кого-то из записанных людей слова идут так же: фамилия, имя, отчество - целиком;
    // инициал подходит к имени ("иванов и" - "Иванов Иван", "иванов иван" - "Иванов И.И.").
    // Двойная фамилия: "Свидлер Лишик Л.А." находится и по "лишик", и по "свидлер-лишик" - по "лишик".
    public static bool NameMatches(string fio, List<string> q) {
        if (q.Count == 0) return false;
        foreach (string person in fio.Split(PersonSeparators, StringSplitOptions.RemoveEmptyEntries)) {
            List<string> w = PersonWords(person);
            // с какого слова может начинаться фамилия: с первого; со второго - если слов много (двойная фамилия)
            for (int from = 0; from <= (w.Count >= 4 ? 1 : 0); from++) {
                if (w.Count - from < q.Count) break;
                bool ok = true;
                for (int k = 0; k < q.Count && ok; k++) {
                    string a = q[k], b = w[from + k];
                    ok = a == b || (Initial(a) && b.StartsWith(a)) || (Initial(b) && k > 0 && a.StartsWith(b))
                        || (k == 0 && Array.IndexOf(b.Split('-'), a) >= 0);
                }
                if (ok) return true;
            }
        }
        return false;
    }

    static bool Initial(string w) {
        return w.Length == 1 && char.IsLetter(w[0]);
    }

    // ФИО: вписанные слова - это фамилия, имя, отчество по порядку, каждое целиком.
    // Одна буква - инициал ("иванов и"); если в базе записан инициал ("Иванов И.И.") - он подходит к имени.
    List<int> FindNameExact(string query) {
        List<int> found = new List<int>();
        List<string> q = PersonWords(query);
        LastExplain = "слов: " + q.Count + ", только точное совпадение (с фамилии, словами целиком)";
        if (q.Count == 0 || nameFields.Length == 0) return found;
        for (int i = 0; i < RowCount; i++) {
            int row = RowAt(i);
            foreach (int f in nameFields)
                if (NameMatches(Text(row, f), q)) { found.Add(row); break; }
        }
        return found;
    }

    // адрес: улица из найденных по справочнику, дом, корпус и квартира - точно (если вписаны)
    List<int> FindAddressExact() {
        List<int> found = new List<int>();
        AddressMatch m = Address;
        Dictionary<string, bool> codes = new Dictionary<string, bool>();
        foreach (string c in m.Codes) codes[StreetDirectory.NormalizeCode(c)] = true;
        string house = m.House == null ? null : AddressMatch.Num(m.House);
        string korp = m.Korp == null ? null : AddressMatch.Num(m.Korp);
        string flat = m.Flat == null ? null : AddressMatch.Num(m.Flat);
        int onStreet = 0, inHouse = 0, inKorp = 0;
        for (int i = 0; i < RowCount; i++) {
            int row = RowAt(i);
            if (ul < 0 || !codes.ContainsKey(StreetDirectory.NormalizeCode(Text(row, ul)))) continue;
            onStreet++;
            if (house != null && (dom < 0 || AddressMatch.Num(Text(row, dom)) != house)) continue;
            inHouse++;
            if (korp != null && (kor < 0 || AddressMatch.Num(Text(row, kor)) != korp)) continue;
            inKorp++;
            if (flat != null && (kv < 0 || AddressMatch.Num(Text(row, kv)) != flat)) continue;
            found.Add(row);
        }
        // где оборвалась цепочка: улица -> дом -> корпус -> квартира
        if (found.Count == 0) {
            if (onStreet == 0) Miss = "Абонентов на этой улице нет.";
            else if (inHouse == 0) Miss = "Дома " + m.House + " здесь нет.";
            else if (inKorp == 0) Miss = "В доме " + m.House + " нет корпуса " + m.Korp + ".";
            else Miss = "В доме " + m.House + (m.Korp != null ? " корп. " + m.Korp : "") + " нет квартиры " + m.Flat + ".";
        }
        LastExplain = m.Explain + "; записей на этих улицах прочитано: " + onStreet + ", по дому: " + inHouse + ", по корпусу: " + inKorp + ", итог: " + found.Count;
        return found;
    }

    // ---- частичное совпадение (галочка "Частичное совпадение") ----

    // лицевой: берём из строки только цифры, нули впереди не важны
    List<int> FindAccount(string query) {
        List<int> found = new List<int>();
        StringBuilder digits = new StringBuilder();
        foreach (char c in query) if (c >= '0' && c <= '9') digits.Append(c);
        LastExplain = "цифр: " + digits.Length;
        if (digits.Length == 0) return found;
        string q = digits.ToString();
        SearchWord w = Part(q);
        string core = q.TrimStart('0');
        foreach (int f in accountFields) {
            for (int i = 0; i < RowCount; i++) {
                int row = RowAt(i);
                if (t.FieldEquals(row, f, w)) { found.Add(row); continue; }
                string v = DbfTable.Format(t.GetValue(row, f, null), true).Trim().TrimStart('0');
                if (core.Length > 0 && v == core) found.Add(row);
            }
        }
        if (found.Count > 0) { LastExplain += ", точное совпадение"; return found; }
        for (int i = 0; i < RowCount; i++) {
            int row = RowAt(i);
            if (InAny(row, accountFields, w)) found.Add(row);
        }
        if (found.Count > 0) { LastExplain += ", часть номера"; return found; }
        // длинный номер (Л/С) может содержать номер абонента внутри себя
        string qt = q.TrimStart('0');
        if (accountFields.Length > 0 && qt.Length >= 8) {
            int f = accountFields[0];
            for (int i = 0; i < RowCount; i++) {
                int row = RowAt(i);
                string nab = DbfTable.Format(t.GetValue(row, f, null), true).Trim().TrimStart('0');
                if (nab.Length >= 6 && nab.Length < qt.Length && qt.IndexOf(nab, StringComparison.Ordinal) >= 0) found.Add(row);
            }
            if (found.Count > 0) LastExplain += ", номер абонента внутри длинного номера";
        }
        return found;
    }

    // номер счётчика или ФИО: все слова должны быть в этом поле
    List<int> FindWords(string query, int[] fields) {
        List<int> found = new List<int>();
        List<SearchWord> words = new List<SearchWord>();
        foreach (string w in DbfTable.SplitWords(query)) words.Add(Part(w));
        LastExplain = "слов: " + words.Count;
        if (words.Count == 0) return found;

        for (int si = 0; si < RowCount; si++) {
            int row = RowAt(si);
            bool ok = true;
            foreach (SearchWord w in words) if (!InAny(row, fields, w)) { ok = false; break; }
            if (ok) found.Add(row);
        }
        return found;
    }

    // Номер счётчика: точно; потом как часть; потом по одним цифрам (без "№", пробелов, дефисов),
    // в том числе когда в списке номер длиннее или короче записанного; потом - по снятому счётчику
    List<int> FindMeter(string query) {
        List<int> found = new List<int>();
        string[] words = DbfTable.SplitWords(query);
        StringBuilder sb = new StringBuilder();
        foreach (char c in query) if (c >= '0' && c <= '9') sb.Append(c);
        string digits = sb.ToString().TrimStart('0');
        LastExplain = "слов: " + words.Length + ", цифр: " + sb.Length;
        if (words.Length == 0) return found;
        int[] where = meterFields.Length > 0 ? meterFields : new int[] { -1 };
        string how = MeterIn(where, words, digits, found);
        if (found.Count == 0 && oldMeter >= 0) {
            how = MeterIn(new int[] { oldMeter }, words, digits, found);
            if (found.Count > 0) how += ", найден по снятому счётчику (NSC)";
        }
        if (found.Count > 0) LastExplain += ", " + how;
        return found;
    }

    string MeterIn(int[] fields, string[] words, string digits, List<int> found) {
        if (words.Length == 1 && fields[0] >= 0) {
            SearchWord w = Part(words[0]);
            foreach (int f in fields)
                for (int i = 0; i < RowCount; i++) { int row = RowAt(i); if (t.FieldEquals(row, f, w)) found.Add(row); }
            if (found.Count > 0) return "точное совпадение";
        }
        List<SearchWord> parts = new List<SearchWord>();
        foreach (string w in words) parts.Add(Part(w));
        for (int i = 0; i < RowCount; i++) {
            int row = RowAt(i);
            bool ok = true;
            foreach (SearchWord w in parts) if (!InAny(row, fields, w)) { ok = false; break; }
            if (ok) found.Add(row);
        }
        if (found.Count > 0) return "часть номера";
        if (digits.Length < 5 || fields[0] < 0) return "";
        foreach (int f in fields) {
            for (int i = 0; i < RowCount; i++) {
                int row = RowAt(i);
                string v = DbfTable.Format(t.GetValue(row, f, null), true);
                StringBuilder d = new StringBuilder();
                foreach (char c in v) if (c >= '0' && c <= '9') d.Append(c);
                string nd = d.ToString().TrimStart('0');
                if (nd.Length < 5) continue;
                if (nd.IndexOf(digits, StringComparison.Ordinal) >= 0 || (nd.Length >= 6 && digits.IndexOf(nd, StringComparison.Ordinal) >= 0)) found.Add(row);
            }
        }
        return found.Count > 0 ? "по цифрам номера" : "";
    }

    // слово ищется как часть поля (и короткие числа тоже - это часть номера, а не дом)
    SearchWord Part(string word) {
        SearchWord w = t.PrepareWord(word);
        w.WholeNumber = false;
        return w;
    }

    // Адрес целиком, как его скопировали: "Острошицкий Городок Заводская ул. 18",
    // "аг. Острошицкий Городок, ул. Заводская, д. 18, кв. 5", "заводская 18 5".
    // Дом, корпус и квартира должны совпасть точно; из слов (населённый пункт, улица)
    // показываем записи, где их совпало больше всего - лишние слова не мешают.
    List<int> FindAddress(string query) {
        List<int> found = new List<int>();
        AddressParts a = AddressParts.Parse(query);
        List<SearchWord> text = new List<SearchWord>();
        List<string> textLower = new List<string>();
        foreach (string w in a.Words) { text.Add(Part(w)); textLower.Add(w.Replace('ё', 'е')); }
        SearchWord house = a.House == null ? null : t.PrepareWord(a.House);
        SearchWord korp = a.Korp == null ? null : t.PrepareWord(a.Korp);
        SearchWord flat = a.Flat == null ? null : t.PrepareWord(a.Flat);
        StringBuilder explain = new StringBuilder();
        explain.Append("слов адреса: " + text.Count + ", дом: " + (house != null ? "да" : "нет") + ", корпус: " + (korp != null ? "да" : "нет")
            + ", квартира: " + (flat != null ? "да" : "нет") + ", пропущено служебных слов и номеров: " + a.Skipped);
        if (text.Count == 0 && house == null && korp == null && flat == null) { LastExplain = explain.ToString(); return found; }

        // где нашлось каждое слово: в каком поле абонента или в справочнике
        int[,] fieldHits = new int[text.Count, addressFields.Length];
        int[] dirHits = new int[text.Count];
        int candidates = 0, best = 0;
        for (int si = 0; si < RowCount; si++) {
            int row = RowAt(si);
            if (house != null && !Is(row, dom, house)) continue;
            if (korp != null && !Is(row, kor, korp)) continue;
            if (flat != null && !Is(row, kv, flat)) continue;
            candidates++;
            int score = 0;
            for (int i = 0; i < text.Count; i++) {
                bool hit = false;
                for (int k = 0; k < addressFields.Length && !hit; k++)
                    if (t.FieldContains(row, addressFields[k], text[i])) { fieldHits[i, k]++; hit = true; }
                if (!hit && streetLower != null && streetLower[row] != null && streetLower[row].IndexOf(textLower[i], StringComparison.Ordinal) >= 0) {
                    dirHits[i]++;
                    hit = true;
                }
                if (hit) score++;
            }
            if (score < best) continue;
            if (score > best) { best = score; found.Clear(); }
            found.Add(row);
        }
        if (text.Count > 0 && best == 0) found.Clear();   // ни одно слово адреса не нашлось - улицы нет

        explain.Append("; подходят по дому/корпусу/квартире: " + candidates);
        for (int i = 0; i < text.Count; i++) {
            List<string> where = new List<string>();
            for (int k = 0; k < addressFields.Length; k++)
                if (fieldHits[i, k] > 0) where.Add("поле " + t.Fields[addressFields[k]].Name + " - " + fieldHits[i, k]);
            if (dirHits[i] > 0) where.Add("справочник - " + dirHits[i]);
            explain.Append("; слово " + (i + 1) + " (" + textLower[i].Length + " букв): " + (where.Count == 0 ? "не найдено" : string.Join(", ", where.ToArray())));
        }
        if (text.Count > 0) explain.Append("; лучшее совпадение: " + best + " из " + text.Count + " слов");
        LastExplain = explain.ToString();
        return found;
    }

    bool InAny(int row, int[] fields, SearchWord w) {
        if (fields.Length == 0) return t.FieldContains(row, -1, w);
        foreach (int f in fields) if (t.FieldContains(row, f, w)) return true;
        return false;
    }

    bool Is(int row, int field, SearchWord w) {
        return field < 0 ? t.FieldContains(row, -1, w) : t.FieldEquals(row, field, w);
    }

    class SortItem {
        public int Row;
        public string Street, House, Flat;
    }

    // по улице, дому и квартире, с числами по порядку: 2, 10, 18а
    void SortByAddress(List<int> rows) {
        if (ul < 0 && dom < 0) return;
        List<SortItem> items = new List<SortItem>(rows.Count);
        foreach (int r in rows) {
            SortItem it = new SortItem();
            it.Row = r;
            it.Street = street != null && street[r] != null ? street[r] : ul < 0 ? "" : DbfTable.Format(t.GetValue(r, ul, null), true);
            it.House = dom < 0 ? "" : DbfTable.Format(t.GetValue(r, dom, null), true);
            it.Flat = kv < 0 ? "" : DbfTable.Format(t.GetValue(r, kv, null), true);
            items.Add(it);
        }
        items.Sort(delegate(SortItem a, SortItem b) {
            int c = string.Compare(a.Street, b.Street, StringComparison.CurrentCultureIgnoreCase);
            if (c == 0) c = Natural(a.House, b.House);
            if (c == 0) c = Natural(a.Flat, b.Flat);
            return c != 0 ? c : a.Row.CompareTo(b.Row);
        });
        rows.Clear();
        foreach (SortItem it in items) rows.Add(it.Row);
    }

    static int Natural(string a, string b) {
        int na = LeadingNumber(a), nb = LeadingNumber(b);
        if (na != nb) return na.CompareTo(nb);
        return string.Compare(a, b, StringComparison.CurrentCultureIgnoreCase);
    }

    // пусто - раньше всех (частный дом без квартиры), без цифр - позже всех
    static int LeadingNumber(string s) {
        int n = 0, i = 0;
        s = s.Trim();
        if (s.Length == 0) return -1;
        while (i < s.Length && i < 9 && s[i] >= '0' && s[i] <= '9') n = n * 10 + (s[i++] - '0');
        return i == 0 ? int.MaxValue : n;
    }

    static bool IsDigitAt(string s, int i) {
        return i < s.Length && s[i] >= '0' && s[i] <= '9';
    }

    static bool IsDigits(string s) {
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return s.Length > 0;
    }

    static bool HasDigit(string s) {
        foreach (char c in s) if (c >= '0' && c <= '9') return true;
        return false;
    }
}

// Главное окно: одна строка поиска. Каждый поиск идёт по индексам прямо в базе - данные всегда свежие.

// Проверка по базе абонентов (только чтение, тем же движком, что «Поиск абонента»):
// есть ли уже такой лицевой (копия лицевого ломает базу) и есть ли абоненты по этому адресу.
public class BaseLink {
    public LiveBase Base;
    public string Folder, Table, Error;
    public bool Ready { get { return Base != null; } }
    readonly object sync = new object();

    // открыть базу; запомненное (папка, таблица, справочник улиц) - в settings, найденное заново - туда же
    public void Open(Dictionary<string, string> settings) {
        try {
            string saved = Get(settings, "БазаПапка");
            Folder = Places.FindFolder(saved.Length > 0 ? saved : null);
            if (!Directory.Exists(Folder)) throw new Exception("нет доступа к папке базы " + Folder);
            Table = Get(settings, "БазаТаблица");
            if (Table.Length == 0 || !File.Exists(Path.Combine(Folder, Table))) {
                Table = File.Exists(Path.Combine(Folder, "ABONENT.DBF")) ? "ABONENT.DBF" : "";
                if (Table.Length == 0) {
                    List<TableInfo> tables = TableInfo.List(Folder, null);
                    if (tables.Count > 0 && tables[0].Abonents) Table = tables[0].Name;
                }
                if (Table.Length == 0) throw new Exception("в папке " + Folder + " не найдена таблица абонентов");
            }
            LiveBase b = LiveBase.Open(Path.Combine(Folder, Table));
            b.Dir = Streets(b, settings);
            settings["БазаПапка"] = Folder;
            settings["БазаТаблица"] = Table;
            if (b.Dir != null) {
                settings["БазаУлицы"] = b.Dir.TableName ?? "";
                settings["БазаУлицыКлюч"] = b.Dir.KeyField ?? "";
                List<string> links = new List<string>();
                foreach (KeyValuePair<string, string> kv in b.Dir.Links) links.Add(kv.Key + "=" + kv.Value);
                settings["БазаСвязи"] = string.Join(";", links.ToArray());
            }
            Base = b;
        } catch (Exception ex) {
            Error = ex.Message;
        }
    }

    static string Get(Dictionary<string, string> d, string k) {
        string v;
        return d.TryGetValue(k, out v) && v != null ? v.Trim() : "";
    }

    // справочник улиц: запомненный или найденный по кодам улиц (как в «Поиске»)
    StreetDirectory Streets(LiveBase b, Dictionary<string, string> settings) {
        if (b.Header.FieldIndex("UL") < 0) return null;
        string table = Get(settings, "БазаУлицы"), key = Get(settings, "БазаУлицыКлюч");
        Dictionary<string, string> links = new Dictionary<string, string>();
        foreach (string pair in Get(settings, "БазаСвязи").Split(';')) {
            int eq = pair.IndexOf('=');
            if (eq > 0) links[pair.Substring(0, eq).Trim().ToUpperInvariant()] = pair.Substring(eq + 1).Trim();
        }
        if (table.Length > 0 && key.Length > 0 && File.Exists(Path.Combine(Folder, table))) {
            try {
                StreetDirectory d = StreetDirectory.Load(Folder, table, key, links);
                if (d != null && d.Count > 0) return d;
            } catch { }
        }
        DbfTable sample = b.SpreadSample(300);
        List<string> codes = StreetDirectory.Codes(sample, sample.FieldIndex("UL"));
        if (codes == null) return null;
        return StreetDirectory.Detect(Folder, Table, codes, null);
    }

    // лицевой в программе 5002123456789: 5002 - приписка района, в базе - 123456789
    public static string Nab(string ls) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in ls) if (char.IsDigit(c)) sb.Append(c);
        string d = sb.ToString();
        return d.Length == 13 ? d.Substring(4) : d;
    }

    static string Clean(string s) { return s.Trim().TrimStart('0'); }

    string F(DbfTable t, int row, string field) {
        int i = t.FieldIndex(field);
        return i < 0 ? "" : DbfTable.Format(t.GetValue(row, i, null), true).Trim();
    }

    // «123456789 - Иванов Иван Иванович, аг. Боровляны, Лесная ул., д.7 кв.2»
    string Describe(DbfTable t, int row) {
        StringBuilder sb = new StringBuilder(F(t, row, "NAB"));
        string fio = F(t, row, "FIO");
        if (fio.Length > 0) sb.Append(" - ").Append(fio);
        string ul = F(t, row, "UL"), place = Base.Dir != null && ul.Length > 0 ? Base.Dir.Name(ul) : null;
        List<string> a = new List<string>();
        if (place != null && place.Length > 0) a.Add(place);
        string dom = F(t, row, "DOM"), kor = F(t, row, "KOR"), kv = F(t, row, "KV");
        string house = (dom.Length > 0 ? "д." + dom : "") + (kor.Length > 0 ? " корп." + kor : "") + (kv.Length > 0 ? " кв." + kv : "");
        if (house.Trim().Length > 0) a.Add(house.Trim());
        if (a.Count > 0) sb.Append(", ").Append(string.Join(", ", a.ToArray()));
        return sb.ToString();
    }

    // записи с этим лицевым (точное совпадение номера)
    public List<string> Accounts(string nab) {
        List<string> found = new List<string>();
        if (Clean(nab).Length == 0) return found;
        lock (sync) {
            SearchResult r = Base.Search(nab, SearchMode.Account, null, false);
            if (r.Error != null) throw new Exception(r.Error);
            if (r.Table == null) return found;
            foreach (int row in r.Rows)
                if (Clean(F(r.Table, row, "NAB")) == Clean(nab)) found.Add(Describe(r.Table, row));
        }
        return found;
    }

    // абоненты по адресу; note - если адрес не понят (нет такого населённого пункта или улицы в справочнике)
    public List<string> Address(string query, out string note) {
        List<string> found = new List<string>();
        note = null;
        lock (sync) {
            SearchResult r = Base.Search(query, SearchMode.Address, null, false);
            if (r.Error != null) throw new Exception(r.Error);
            if (r.Table != null)
                foreach (int row in r.Rows) if (!r.Added.ContainsKey(row)) found.Add(Describe(r.Table, row));
            if (found.Count == 0 && r.Note.Length > 0) note = r.Note;
        }
        return found;
    }
}

// Оформление договора электроснабжения: форма -> шаблон Word (shablon.docx) -> готовый договор
public class DogovorForm : Form {
    readonly string here, templateFile, outDir, dataDir;
    readonly Lists lists;
    readonly Memory memory;
    readonly Dictionary<string, Control> inp = new Dictionary<string, Control>();
    readonly Dictionary<string, Card> cardOf = new Dictionary<string, Card>();
    readonly List<Card> cards = new List<Card>();
    readonly List<NavItem> navs = new List<NavItem>();
    readonly Panel content = new Panel();
    readonly Label head = new Label(), headSub = new Label(), status = new Label();
    readonly ToolTip tip = new ToolTip();
    readonly Timer timer = new Timer();
    BoxPanel meter2;
    readonly List<Control> cardOfLater = new List<Control>();
    // выбор счётчика: что уже разобрано по спискам, что подставлено само
    readonly string[] lastModel = new string[3], lastCase = new string[3], autoTip = new string[3], autoModem = new string[3];
    const string Own = "свое";
    bool loading, showMissing, dirty, unsavedAuto;
    string openedFile;                      // файл данных договора, который сейчас в форме (открыт или сохранён)
    readonly string autoFile;               // автосохранение - один файл, перезаписывается
    readonly Timer autoTimer = new Timer();
    // проверка по базе абонентов: копия лицевого запрещена, такой же адрес - предупреждение
    BaseLink baseLink;
    readonly Label lsNote = Ui.Note(""), adresNote = Ui.Note(""), baseLabel = new Label();
    readonly Timer checkTimer = new Timer();
    bool lsPending, adrPending;
    int lsSeq, adrSeq;
    string lsChecked, lsError, lsSaved;              // для какого ЛС есть ответ базы; ошибка; другой сохранённый договор с этим ЛС
    List<string> lsInBase = new List<string>();      // записи базы с этим лицевым
    string autoInic = "", autoAdres = "";

    public DogovorForm(string here) {
        this.here = here;
        templateFile = Path.Combine(here, "shablon.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = new Memory(Path.Combine(here, "nastroiki.txt"));
        outDir = memory.Get("ПапкаДоговоров", "");
        if (outDir.Length == 0) outDir = Path.Combine(here, "Договоры");
        dataDir = Path.Combine(outDir, "данные");
        autoFile = Path.Combine(here, "avtosohranenie.txt");

        Text = "Оформление договора электроснабжения";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1280, wa.Width), Math.Min(900, wa.Height));
        MinimumSize = new Size(900, 560);
        KeyPreview = true;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

        BuildCards();

        content.Dock = DockStyle.Fill;
        content.AutoScroll = true;
        content.BackColor = St.Back;
        foreach (Card c in cards) content.Controls.Add(c.Panel);
        content.Resize += delegate { LayoutCards(); };
        Controls.Add(content);
        Controls.Add(BuildHeader());
        Controls.Add(BuildBottom());
        Controls.Add(BuildNav());

        timer.Interval = 150;
        timer.Tick += delegate { SyncNav(); };
        timer.Start();

        FillLists();
        NewContract();
        // списки - ещё раз, когда ширина полей известна (длинные варианты переносятся по ширине)
        Shown += delegate { LayoutCards(); FillLists(); FocusField(inp["ls"]); RestoreAuto(); StartBase(); };
        checkTimer.Interval = 700;
        checkTimer.Tick += delegate { checkTimer.Stop(); RunChecks(); };
        Activated += delegate {
            if (!lists.Changed()) return;
            try { lists.Load(); FillLists(); } catch { }
        };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.S) { e.SuppressKeyPress = true; Create(true, true); }
            else if (e.Control && e.KeyCode == Keys.O) { e.SuppressKeyPress = true; OpenSaved(); }
            else if (e.Control && e.KeyCode == Keys.N) { e.SuppressKeyPress = true; AskNew(); }
            else if (e.KeyCode == Keys.Enter && !e.Control && !e.Alt) {
                // Enter - к следующему полю, как в программе базы
                Control a = ActiveControl;
                if (a == null || (a is TextBox && ((TextBox)a).Multiline) || a is Button) return;
                if (a is ComboBox && ((ComboBox)a).DroppedDown) return;
                e.SuppressKeyPress = true;
                SelectNextControl(a, true, true, true, true);
            }
        };
        FormClosing += delegate(object s, FormClosingEventArgs e) {
            // выключение Windows - спрашивать некогда: всё уже в автосохранении
            if (e.CloseReason == CloseReason.WindowsShutDown || e.CloseReason == CloseReason.TaskManagerClosing) { AutoSave(); return; }
            if (!AskSave("Закрыть программу")) e.Cancel = true;
        };
        // автосохранение: через полминуты после изменений (отключат свет - пропадёт не больше полминуты работы)
        autoTimer.Interval = 30000;
        autoTimer.Tick += delegate { if (unsavedAuto) AutoSave(); };
        autoTimer.Start();
    }

    // ---------- поля ----------

    TextBox T(string key, int lines, string cue) {
        TextBox t = Ui.Text(lines);
        Reg(key, t);
        if (cue != null) t.HandleCreated += delegate { Ui.Cue(t, cue); };
        return t;
    }

    Combo C(string key, bool editable, bool wrap, string cue) {
        Combo c = new Combo(editable, wrap);
        Reg(key, c);
        if (cue != null) c.HandleCreated += delegate { Ui.Cue(c, cue); };
        return c;
    }

    Seg S(string key, params string[] options) {
        Seg s = new Seg(options);
        Reg(key, s);
        return s;
    }

    DateBox D(string key) {
        DateBox d = new DateBox();
        Reg(key, d);
        return d;
    }

    CheckBox X(string key, string text) {
        CheckBox c = new CheckBox();
        c.Text = text;
        c.Font = St.Base;
        c.ForeColor = St.Text;
        c.BackColor = Color.Transparent;
        c.Cursor = Cursors.Hand;
        Reg(key, c);
        return c;
    }

    void Reg(string key, Control c) {
        inp[key] = c;
        c.Name = key;
        EventHandler h = delegate { Changed(key); };
        if (c is Seg) ((Seg)c).Changed += h;
        else if (c is ComboBox) { ((ComboBox)c).TextChanged += h; ((ComboBox)c).SelectedIndexChanged += h; }
        else if (c is CheckBox) ((CheckBox)c).CheckedChanged += h;
        else if (c is DateBox) ((DateBox)c).ValueChanged += h;
        else if (c is Heaters) { ((Heaters)c).Changed += h; ((Heaters)c).Resized += delegate { LayoutCards(); }; }
        else c.TextChanged += h;
    }

    public string V(string key) {
        Control c = inp[key];
        if (c is Combo && c.Text.Trim() == Own) return "";
        if (c is Seg) return ((Seg)c).Value;
        if (c is CheckBox) return ((CheckBox)c).Checked ? "1" : "";
        if (c is DateBox) return ((DateBox)c).Text2;
        if (c is Heaters) return ((Heaters)c).Value;
        return c.Text.Trim();
    }

    public void SetV(string key, string v) {
        Control c;
        if (!inp.TryGetValue(key, out c)) return;
        v = v ?? "";
        if (c is Seg) ((Seg)c).Value = v;
        else if (c is CheckBox) ((CheckBox)c).Checked = v == "1";
        else if (c is DateBox) {
            DateTime d;
            if (DateTime.TryParseExact(v, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) ((DateBox)c).Value = d;
        } else if (c is Combo) ((Combo)c).Value = v;
        else if (c is Heaters) ((Heaters)c).Value = v;
        else c.Text = v;
    }

    // ---------- разделы ----------

    Card NewCard(string name, string sub, int cols) {
        Card c = new Card(name, sub, cols);
        cards.Add(c);
        return c;
    }

    void Row(Card card, params Cell[] cells) {
        card.Grid.Row(cells);
        foreach (Cell c in cells) if (c.Ctl != null && c.Ctl.Name.Length > 0) cardOf[c.Ctl.Name] = card;
    }

    static Cell F(string caption, Control c, int span) { return Cell.F(caption, c, span); }

    void BuildCards() {
        Card c = NewCard("Договор", "дата, номер и кто оформляет", 4);
        Button remember = Ui.Button("Запомнить подписанта", false);
        remember.Font = St.Small;
        remember.Size = new Size(190, 28);
        remember.Click += delegate { RememberSigner(); };
        remember.TabStop = false;
        tip.SetToolTip(remember, "Записать должность, доверенность и подпись в справочник,\nчтобы в следующий раз выбрать подписанта из списка");
        c.AddExtra(remember);
        Row(c, F("Дата договора", D("data"), 1), F("Лицевой счёт №", T("ls", 1, "номер лицевого счёта"), 1), F("Договор оформил(а)", C("oformil", true, false, "фамилия И.О."), 2));
        Row(c, Cell.Empty(1), F(null, lsNote, 3));
        Row(c, F("Исполнитель в лице (подписывает)", C("podp", true, false, "выберите из списка"), 1), F("Подпись подписанта", T("podp_inic", 1, "И.О.Фамилия"), 1),
            F("Доверенность", T("podp_dov", 1, "№ ... от ..."), 2));
        Row(c, F("Должность подписанта (в лице кого)", T("podp_dolzh", 3, null), 3), F("Ф.И.О. подписанта (кого)", T("podp_fio", 3, null), 1));
        Row(c, F("Сообщать о неисправностях (п. 3.6): адрес", T("soob_adres", 1, null), 2), F("телефон", T("soob_tel", 1, null), 2));

        c = NewCard("Потребитель", "гражданин(ка), с которым заключается договор", 4);
        Row(c, F("Фамилия, собственное имя, отчество", T("fio", 1, "Иванов Иван Иванович"), 2), F("Подпись (инициалы)", T("potr_inic", 1, "заполнится сама"), 1),
            F("Идентификационный номер", T("lichny", 1, "14 знаков"), 1));
        Button same = Ui.Button("= адрес объекта", false);
        same.Font = St.Small;
        same.Click += delegate { SetV("adres_prozh", V("adres")); };
        same.TabStop = false;
        tip.SetToolTip(same, "Потребитель проживает по адресу объекта");
        Row(c, F("Адрес проживания (регистрации)", T("adres_prozh", 1, "г.Минск, ул. ..., д. ..., кв. ..."), 2), F(" ", same, 1), F("Телефон", T("telefon", 1, "необязательно"), 1));

        c = NewCard("Объект", "что подключаем и где", 8);
        Row(c, F("Объект электроснабжения", C("obekt", true, false, null), 4), F("Разрешённая мощность, кВт", T("moshn", 1, "например 10"), 2),
            F("из неё на отопление и ГВС, кВт", T("moshn_otopl", 1, "0,0"), 2));
        Row(c, F("Тип н.п.", C("np_tip", true, false, null), 1), F("Населённый пункт / садоводческое товарищество", C("np", true, false, "Боровляны"), 3),
            F("Сельсовет", C("selsovet", true, false, null), 2), F("Кадастровый номер", T("kadastr", 1, "цифры или № по свидетельству"), 2));
        Row(c, F("Улица", C("ulica_tip", true, false, null), 1), F(" ", C("ulica", true, false, "Дружная"), 3), F("Дом / участок", T("dom", 1, null), 2),
            F("Корпус", T("korp", 1, null), 1), F("Квартира", T("kv", 1, null), 1));
        Row(c, F("Адрес объекта - так будет в договоре (собирается сам из полей выше, можно поправить)", T("adres", 1, null), 8));
        Row(c, F(null, adresNote, 8));
        Row(c, F("Допуск: акт осмотра №", T("dopusk_nom", 1, "например 407-Ф"), 2), F("от", D("dopusk_data"), 2),
            F("кем выдан", C("dopusk_kem", true, false, null), 4));
        Row(c, F("Прочие условия (пункт 28.1)", C("prochie", true, true, null), 8));

        c = NewCard("Счётчики", "средства расчётного учёта и тарифы", 2);
        Seg count = S("schetchikov", "1 счётчик", "2 счётчика");
        count.Size = new Size(240, Ui.InputHeightCached);
        c.AddExtra(count);
        Grid g1 = Meter(1), g2 = Meter(2);
        meter2 = (BoxPanel)g2.Host;
        c.Grid.Row(Cell.G(g1, 1), Cell.G(g2, 1));
        foreach (Grid g in new Grid[] { g1, g2 })
            foreach (Cell x in g.Cells()) if (x.Ctl != null && x.Ctl.Name.Length > 0) cardOf[x.Ctl.Name] = c;
        foreach (Control x in cardOfLater) cardOf[x.Name] = c;
        cardOf["schetchikov"] = c;

        c = NewCard("Пломбы и оборудование", "", 4);
        Row(c, F("Пломба на вводном аппарате", T("plomba_vvod", 1, "пусто - прочерк"), 1), F("на клеммной крышке", T("plomba_klemm", 1, "пусто - прочерк"), 1),
            F("другие пломбы", T("plomba_drugoe", 1, "например ДСТП 1234567"), 2));
        Row(c, F("Электроплита стационарная", S("plita", "да", "нет"), 1), F("Центр. газоснабжение", S("gaz", "да", "нет"), 1),
            F("Центр. теплоснабжение", S("teplo", "да", "нет"), 1), F("Центр. горячее водоснабжение", S("gvs", "да", "нет"), 1));

        c = NewCard("Дополнительно", "льгота, особые условия, СМС", 4);
        Row(c, F("СМС-оповещение", S("sms", "согласен(а)", "отказ"), 1), F("Номер для СМС", T("sms_tel", 1, "+375 ..."), 1), F("Оформлена льгота", T("lgota", 1, "пусто - прочерк"), 2));
        Row(c, F("Особые условия", T("osobye", 2, null), 4));
        Row(c, F(" ", X("utratil", "Считать утратившим силу приложение к договору"), 2), F("№", T("utratil_nom", 1, null), 1), F("от", D("utratil_ot"), 1));
        Row(c, F("Дополнительная информация", T("dop_info", 2, null), 4));

        c = NewCard("Порядок и документы", "для себя - в договор не попадает", 2);
        Label steps = Ui.Note(Poryadok());
        steps.ForeColor = St.Text;
        Grid docs = new Grid(new Panel(), 1);
        docs.Host.BackColor = Color.Transparent;
        docs.PadX = 0;
        docs.PadBottom = 0;
        docs.RowGap = 2;
        string[][] list = new string[][] {
            new string[] { "doc_tu", "Технические условия" },
            new string[] { "doc_proekt", "Проект электроснабжения" },
            new string[] { "doc_razgr", "Акт разграничения балансовой принадлежности" },
            new string[] { "doc_osmotr", "Акт осмотра ГУ «Госэнергогазнадзор»" },
            new string[] { "doc_pasport", "Паспортные данные владельца" },
            new string[] { "doc_pravo", "Правоудостоверяющие документы на участок / строение" } };
        foreach (string[] d in list) docs.Row(Cell.F(null, X(d[0], d[1]), 1));
        c.Grid.Row(Cell.F(null, steps, 1), Cell.G(docs, 1));
    }

    Grid Meter(int i) {
        string p = "pu" + i + "_";
        BoxPanel box = new BoxPanel();
        box.Title = "Счётчик " + i;
        Grid g = new Grid(box, 2);
        g.PadX = 14;
        g.PadTop = 38;
        g.PadBottom = 16;
        g.RowGap = 10;
        // тип счётчика - по частям из справочника; в конце каждого списка «свое»
        Grid parts = new Grid(new Panel(), 4);
        parts.Host.BackColor = Color.Transparent;
        parts.PadX = 0;
        parts.PadBottom = 0;
        parts.Row(F("Модель", C(p + "model", true, false, null), 1), F("Тип корпуса", C(p + "korpus", true, false, null), 1),
            F("Модификация", C(p + "modif", true, false, null), 2));
        g.Row(Cell.G(parts, 2));
        foreach (Cell x in parts.Cells()) cardOfLater.Add(x.Ctl);
        g.Row(F("Тип счётчика - так будет в договоре (собирается сам, можно поправить)", T(p + "tip", 1, null), 2));
        g.Row(F("Заводской номер", T(p + "zav", 1, null), 1), F("Дата поверки (год-квартал)", T(p + "poverka", 1, "например 2025-2"), 1));
        g.Row(F("Цель использования электроэнергии", C(p + "cel", true, false, null), 2));
        g.Row(F("Фазность", S(p + "faz", "1 ф", "3 ф"), 1), F("Пломба госповерителя", C(p + "plomba_gp", true, false, null), 1));
        g.Row(F("Место установки", C(p + "mesto", true, false, null), 2));
        g.Row(F("Принадлежность", C(p + "prinadl", true, false, null), 1), F("Тип модема", C(p + "modem", true, false, null), 1));
        g.Row(F("Включение в АСКУЭ", S(p + "askue", "да", "нет"), 1), F("Показания, кВт·ч", T(p + "pokaz", 1, null), 1));
        g.Row(F("Вводной аппарат: тип", C(p + "vvod_tip", true, false, null), 1), F("ток, А", C(p + "vvod_i", true, false, null), 1));
        g.Row(F("Аппарат защиты: тип", C(p + "zasch_tip", true, false, null), 1), F("ток, А", C(p + "zasch_i", true, false, null), 1));
        g.Row(F("Модуль отображения (сплит-счётчик), зав. №", T(p + "modul", 1, "нет модуля - прочерк"), 2));
        g.Row(F("Тариф", S(p + "tarif", Tarifs), 2));
        g.Row(F("Тарифная группа", C(p + "gruppa", true, true, null), 2));
        Heaters heat = new Heaters();
        Reg(p + "nagrev", heat);
        g.Row(F("Электронагревательные приборы: вид, шт., кВт одного (+ ещё прибор)", heat, 2));
        return g;
    }

    // одноставочный или дифференцированный по 2 или 3 временным периодам
    static readonly string[] Tarifs = new string[] { "одноставочный", "дифф., 2 периода", "дифф., 3 периода" };

    static string Periods(string tarif) {
        return tarif == Tarifs[1] ? "2" : tarif == Tarifs[2] ? "3" : "";
    }

    string Poryadok() {
        string f = Path.Combine(here, "poryadok.txt");
        try { if (File.Exists(f)) return File.ReadAllText(f, Encoding.UTF8).Trim(); } catch { }
        return "1. Сельский совет - технические условия.\n" +
            "2. Проект электроснабжения (любая сертифицированная организация).\n" +
            "3. Монтаж по проекту.\n" +
            "4. Филиал электрических сетей - акт разграничения балансовой принадлежности и эксплуатационной ответственности.\n" +
            "5. ГУ «Госэнергогазнадзор» (Старовиленская, 100А) - акт осмотра электроустановок для ввода в эксплуатацию.\n" +
            "6. Документы - в филиал «Энергосбыт» Острошицкий ПСУ для заключения договора.";
    }

    // ---------- рамка формы: меню слева, заголовок, кнопки внизу ----------

    Control BuildNav() {
        Panel nav = new Panel();
        nav.Dock = DockStyle.Left;
        nav.Width = 236;
        nav.BackColor = St.Nav;
        Label title = new Label();
        title.Text = "Оформление\nдоговора";
        title.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
        title.ForeColor = Color.White;
        title.SetBounds(22, 22, 210, 64);
        Label sub = new Label();
        sub.Text = "электроснабжения";
        sub.Font = St.Small;
        sub.ForeColor = St.NavText;
        sub.SetBounds(24, 86, 200, 20);
        nav.Controls.Add(title);
        nav.Controls.Add(sub);
        int y = 128;
        foreach (Card c in cards) {
            NavItem n = new NavItem(c.Name, false);
            n.SetBounds(0, y, nav.Width, n.Height);
            Card card = c;
            n.Click += delegate { ScrollTo(card); };
            navs.Add(n);
            nav.Controls.Add(n);
            y += n.Height;
        }
        NavItem folder = new NavItem("Папка с договорами", true);
        folder.Click += delegate { OpenFolder(); };
        NavItem dirs = new NavItem("Справочники (списки)", true);
        dirs.Click += delegate { OpenLists(); };
        foreach (NavItem n in new NavItem[] { folder, dirs }) {
            n.Anchor = AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right;
            nav.Controls.Add(n);
        }
        baseLabel.Font = St.Small;
        baseLabel.ForeColor = St.NavText;
        baseLabel.BackColor = Color.Transparent;
        baseLabel.Text = "База: подключаю...";
        nav.Controls.Add(baseLabel);
        nav.Resize += delegate {
            dirs.SetBounds(0, nav.ClientSize.Height - 20 - dirs.Height, nav.Width, dirs.Height);
            folder.SetBounds(0, dirs.Top - folder.Height, nav.Width, folder.Height);
            baseLabel.SetBounds(22, folder.Top - 58, nav.Width - 34, 52);
        };
        return nav;
    }

    Control BuildHeader() {
        Panel p = new Panel();
        p.Dock = DockStyle.Top;
        p.Height = 74;
        p.BackColor = St.Back;
        head.Font = St.Big;
        head.ForeColor = St.Text;
        head.AutoSize = true;
        head.Location = new Point(28, 14);
        headSub.Font = St.Small;
        headSub.ForeColor = St.Hint;
        headSub.AutoSize = true;
        headSub.Location = new Point(30, 46);
        p.Controls.Add(head);
        p.Controls.Add(headSub);
        return p;
    }

    Control BuildBottom() {
        Panel p = new Panel();
        p.Dock = DockStyle.Bottom;
        p.Height = 66;
        p.BackColor = Color.White;
        p.Paint += delegate(object s, PaintEventArgs e) {
            using (Pen pen = new Pen(St.Border)) e.Graphics.DrawLine(pen, 0, 0, p.Width, 0);
        };
        Button nw = Ui.Button("Создать новый договор", false);
        nw.SetBounds(24, 14, 200, 38);
        nw.Click += delegate { AskNew(); };
        tip.SetToolTip(nw, "Чистая форма для нового договора; несохранённый текущий - спросит, сохранить ли (Ctrl+N)");
        Button open = Ui.Button("Открыть сохранённый…", false);
        open.SetBounds(234, 14, 200, 38);
        open.Click += delegate { OpenSaved(); };
        tip.SetToolTip(open, "Открыть ранее оформленный договор, чтобы исправить или сделать похожий (Ctrl+O)");
        Button create = Ui.Button("Сохранить договор", true);
        create.Size = new Size(210, 40);
        create.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        create.Click += delegate { Create(true, true); };
        tip.SetToolTip(create, "Заполнить шаблон, сохранить договор в папку «Договоры» и открыть в Word (Ctrl+S)");
        status.AutoSize = false;
        status.Font = St.Base;
        status.TextAlign = ContentAlignment.MiddleLeft;
        status.Cursor = Cursors.Hand;
        status.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
        status.Click += delegate { GoToMissing(); };
        p.Controls.Add(nw);
        p.Controls.Add(open);
        p.Controls.Add(create);
        p.Controls.Add(status);
        p.Resize += delegate {
            create.Location = new Point(p.ClientSize.Width - 24 - create.Width, 13);
            status.SetBounds(open.Right + 24, 14, Math.Max(10, create.Left - open.Right - 48), 38);
        };
        return p;
    }

    void LayoutCards() {
        content.SuspendLayout();
        int full = content.Width - SystemInformation.VerticalScrollBarWidth;
        int w = Math.Max(600, Math.Min(full - 48, 1240));
        int x = Math.Max(24, (full - w) / 2);
        int y = 4 + content.AutoScrollPosition.Y;
        foreach (Card c in cards) {
            int h = c.Layout(w);
            c.Panel.Location = new Point(x + content.AutoScrollPosition.X, y);
            y += h + 18;
        }
        content.AutoScrollMinSize = new Size(0, y - content.AutoScrollPosition.Y + 6);
        content.ResumeLayout();
    }

    void ScrollTo(Card c) {
        int top = c.Panel.Top - content.AutoScrollPosition.Y - 4;
        content.AutoScrollPosition = new Point(0, Math.Max(0, top));
        SyncNav();
    }

    void SyncNav() {
        int active = 0;
        // раздел, который сейчас вверху; в самом низу - последний
        int view = content.ClientSize.Height;
        bool bottom = -content.AutoScrollPosition.Y + view >= content.DisplayRectangle.Height - 2;
        for (int i = 0; i < cards.Count; i++) if (cards[i].Panel.Top <= 60) active = i;
        if (bottom && content.AutoScrollPosition.Y != 0) active = cards.Count - 1;
        for (int i = 0; i < navs.Count; i++) if (navs[i].Active != (i == active)) { navs[i].Active = i == active; navs[i].Invalidate(); }
    }

    // ---------- справочники ----------

    void Items(string key, List<string> items) {
        Combo c = inp[key] as Combo;
        if (c == null) return;
        bool was = loading;
        loading = true;
        try { c.SetItems(items); } finally { loading = was; }
    }

    List<string> With(List<string> a, List<string> b) {
        List<string> r = new List<string>(a);
        foreach (string s in b) {
            bool found = false;
            foreach (string x in r) if (string.Compare(x, s, StringComparison.OrdinalIgnoreCase) == 0) found = true;
            if (!found) r.Add(s);
        }
        return r;
    }

    void FillLists() {
        Items("oformil", lists.Get("Оформил"));
        List<string> signers = new List<string>();
        foreach (string[] s in lists.Signers()) signers.Add(s[0]);
        Items("podp", signers);
        Items("obekt", lists.Get("Объект электроснабжения"));
        Items("np_tip", lists.Get("Тип населённого пункта"));
        Items("ulica_tip", lists.Get("Тип улицы"));
        Items("np", memory.Hist("нп"));
        Items("ulica", memory.Hist("улица"));
        Items("selsovet", With(lists.Get("Сельсовет"), memory.Hist("сельсовет")));
        Items("dopusk_kem", With(lists.Get("Кем выдан допуск"), memory.Hist("допуск")));
        Items("prochie", lists.Get("Прочие условия"));
        for (int i = 1; i <= 2; i++) {
            string p = "pu" + i + "_";
            MeterLists(i, false);
            Items(p + "cel", lists.Get("Цель использования"));
            Items(p + "plomba_gp", lists.Get("Пломба госповерителя"));
            Items(p + "mesto", With(lists.Get("Место установки"), memory.Hist("место")));
            Items(p + "prinadl", lists.Get("Принадлежность"));
            Items(p + "modem", lists.Get("Тип модема"));
            Items(p + "vvod_tip", WithOwn(lists.Get("Вводной аппарат")));
            Items(p + "zasch_tip", WithOwn(lists.Get("Аппарат защиты")));
            Items(p + "vvod_i", WithOwn(lists.Get("Ток аппарата, А")));
            Items(p + "zasch_i", WithOwn(lists.Get("Ток аппарата, А")));
            Items(p + "gruppa", lists.Get("Тарифная группа"));
            ((Heaters)inp[p + "nagrev"]).SetKinds(lists.Get("Электроприборы"));
        }
    }

    static List<string> WithOwn(List<string> a) {
        List<string> r = new List<string>(a);
        r.Add(Own);
        return r;
    }

    // списки с «свое» в конце: выбрали «свое» - поле очищается для своего значения
    static bool OwnList(string key) {
        return key.EndsWith("_vvod_tip") || key.EndsWith("_zasch_tip") || key.EndsWith("_vvod_i") || key.EndsWith("_zasch_i");
    }

    static bool NoInput(string tip) { return tip.Trim().ToLower() == "отсутствует"; }

    void OpenLists() {
        try { Process.Start("notepad.exe", "\"" + lists.File + "\""); }
        catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть " + lists.File + ":\n" + ex.Message, Text); }
    }

    void OpenFolder() {
        try {
            Directory.CreateDirectory(outDir);
            Process.Start("explorer.exe", "\"" + outDir + "\"");
        } catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть папку " + outDir + ":\n" + ex.Message, Text); }
    }

    void RememberSigner() {
        string[] v = new string[] { V("podp"), Flat(V("podp_dolzh")), Flat(V("podp_fio")), V("podp_dov"), V("podp_inic") };
        if (v[0].Length == 0) v[0] = Short(v[2]);
        if (v[0].Length == 0 || v[1].Length == 0 || v[2].Length == 0) {
            MessageBox.Show(this, "Заполните должность и Ф.И.О. подписанта - тогда его можно запомнить.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try {
            lists.SaveSigner(v);
            FillLists();
            SetV("podp", v[0]);
            MessageBox.Show(this, "Подписант «" + v[0] + "» записан в справочник.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        } catch (Exception ex) { MessageBox.Show(this, "Не удалось записать справочник:\n" + ex.Message, Text); }
    }

    static string Flat(string s) { return s.Replace("\r", " ").Replace("\n", " ").Replace("  ", " ").Trim(); }

    // "Зенченко Дмитрия Ивановича" -> "Зенченко Д.И."
    static string Short(string fio) {
        string[] w = fio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (w.Length == 0) return "";
        StringBuilder sb = new StringBuilder(w[0]);
        if (w.Length > 1) sb.Append(' ');
        for (int i = 1; i < w.Length && i < 3; i++) sb.Append(char.ToUpper(w[i][0])).Append('.');
        return sb.ToString();
    }

    // ---------- реакция на ввод ----------

    void Changed(string key) {
        if (loading) return;
        dirty = true;
        unsavedAuto = true;
        loading = true;
        try {
            if (key == "fio") {
                string auto = Initials(V("fio"));
                if (V("potr_inic") == autoInic || V("potr_inic").Length == 0) SetV("potr_inic", auto);
                autoInic = auto;
            }
            if (key == "np_tip" || key == "np" || key == "ulica_tip" || key == "ulica" || key == "dom" || key == "korp" || key == "kv" || key == "selsovet") {
                string auto = Address();
                if (V("adres") == autoAdres || V("adres").Length == 0) SetV("adres", auto);
                autoAdres = auto;
            }
            if (key == "podp") {
                foreach (string[] s in lists.Signers())
                    if (s[0] == V("podp")) { SetV("podp_dolzh", s[1]); SetV("podp_fio", s[2]); SetV("podp_dov", s[3]); SetV("podp_inic", s[4]); }
            }
            if (key == "obekt") {
                // стройплощадка - своё условие 28.1
                List<string> pr = lists.Get("Прочие условия");
                string build = null;
                foreach (string s in pr) if (s.ToLower().Contains("после завершения строительства")) build = s;
                bool site = V("obekt").ToLower().Contains("стройплощадка");
                if (site && build != null) SetV("prochie", build);
                else if (!site && pr.Count > 0 && (V("prochie").Length == 0 || V("prochie") == build)) SetV("prochie", pr[0]);
            }
            for (int i = 1; i <= 2; i++) {
                string p = "pu" + i + "_";
                if (key == p + "model") MeterPart(i, 0);
                else if (key == p + "korpus") MeterPart(i, 1);
                else if (key == p + "modif") MeterPart(i, 2);
            }
            if (OwnList(key) && inp[key].Text.Trim() == Own) {
                Control c = inp[key];
                BeginInvoke((MethodInvoker)delegate { loading = true; try { c.Text = ""; } finally { loading = false; } c.Focus(); });
            }
            // вводного аппарата нет - нет и тока
            for (int i = 1; i <= 2; i++)
                if (key == "pu" + i + "_vvod_tip" && NoInput(V(key))) SetV("pu" + i + "_vvod_i", "");
            if (key == "lichny") {
                TextBox t = (TextBox)inp["lichny"];
                int pos = t.SelectionStart;
                string up = t.Text.ToUpper().Replace(" ", "");
                if (up != t.Text) { t.Text = up; t.SelectionStart = Math.Min(pos, up.Length); }
            }
        } finally { loading = false; }
        if (key == "ls") Recheck(true, false);
        if (key == "adres" || key == "np_tip" || key == "np" || key == "ulica_tip" || key == "ulica" || key == "dom" || key == "korp" || key == "kv") Recheck(false, true);
        if (key.EndsWith("tarif") || key.EndsWith("vvod_tip") || key == "schetchikov" || key == "sms" || key == "utratil" || key == "plita") Enable();
        RefreshState();
    }

    void Enable() {
        bool two = V("schetchikov") != "1 счётчик";
        meter2.Enabled = two;
        meter2.BackColor = two ? St.Box : St.Disabled;
        meter2.Invalidate();
        inp["sms_tel"].Enabled = V("sms") == "согласен(а)";
        inp["utratil_nom"].Enabled = inp["utratil_ot"].Enabled = V("utratil") == "1";
        for (int i = 1; i <= 2; i++) inp["pu" + i + "_vvod_i"].Enabled = !NoInput(V("pu" + i + "_vvod_tip"));
    }

    // ---------- проверка по базе абонентов ----------

    void Post(MethodInvoker m) {
        try { if (!IsDisposed && IsHandleCreated) BeginInvoke(m); } catch { }
    }

    // база открывается в фоне: заголовок, индексы и справочник улиц (как в «Поиске»); самих абонентов не читаем
    void StartBase() {
        Dictionary<string, string> settings = new Dictionary<string, string>(memory.Values);
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            BaseLink b = new BaseLink();
            b.Open(settings);
            Post(delegate {
                if (b.Ready) {
                    baseLink = b;
                    foreach (string k in new string[] { "БазаПапка", "БазаТаблица", "БазаУлицы", "БазаУлицыКлюч", "БазаСвязи" })
                        if (settings.ContainsKey(k)) memory.Values[k] = settings[k];
                    memory.Save();
                    baseLabel.Text = "База: " + b.Table + "\n" + b.Base.RecordCount.ToString("#,0") + " записей - лицевые и адреса проверяются";
                    baseLabel.ForeColor = Color.FromArgb(150, 210, 170);
                } else {
                    baseLabel.Text = "База недоступна: " + b.Error + "\nлицевые не проверяются";
                    baseLabel.ForeColor = Color.FromArgb(240, 150, 150);
                }
                Recheck(true, true);
            });
        });
        th.IsBackground = true;
        th.Start();
    }

    void Recheck(bool ls, bool adres) {
        if (ls) lsPending = true;
        if (adres) adrPending = true;
        checkTimer.Stop();
        checkTimer.Start();
    }

    void RunChecks() {
        if (lsPending) { lsPending = false; CheckLs(); }
        if (adrPending) { adrPending = false; CheckAdres(); }
    }

    static string Digits(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s) if (char.IsDigit(c)) sb.Append(c);
        return sb.ToString();
    }

    // другой сохранённый договор с этим лицевым (тот, что сейчас открыт, не в счёт)
    string SavedWithLs(string ls) {
        try {
            if (ls.Length == 0 || !Directory.Exists(dataDir)) return null;
            foreach (string f in Directory.GetFiles(dataDir, "*.txt")) {
                if (openedFile != null && string.Compare(Path.GetFullPath(f), Path.GetFullPath(openedFile), StringComparison.OrdinalIgnoreCase) == 0) continue;
                foreach (string line in File.ReadAllLines(f, Encoding.UTF8))
                    if (line.StartsWith("ls=")) { if (Digits(line.Substring(3)) == ls) return Path.GetFileNameWithoutExtension(f); break; }
            }
        } catch { }
        return null;
    }

    void Note(Label l, string text, Color color, bool bold) {
        if (l.Text == text && l.ForeColor == color) return;
        l.Text = text;
        l.ForeColor = color;
        l.Font = bold ? new Font(St.Small, FontStyle.Bold) : St.Small;
        LayoutCards();
    }

    static readonly Color Warn = Color.FromArgb(190, 105, 0);

    void CheckLs() {
        string ls = Digits(V("ls"));
        lsSaved = SavedWithLs(ls);
        string nab = BaseLink.Nab(ls);
        if (ls.Length == 0 || lsSaved != null || baseLink == null || ls.Length < 9) { ShowLs(); RefreshState(); return; }
        if (lsChecked == ls) { ShowLs(); return; }
        int seq = ++lsSeq;
        Note(lsNote, "Проверяю лицевой " + nab + " по базе...", St.Hint, false);
        BaseLink b = baseLink;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            List<string> found = null;
            string err = null;
            try { found = b.Accounts(nab); } catch (Exception ex) { err = ex.Message; }
            Post(delegate {
                if (seq != lsSeq) return;
                lsChecked = err == null ? ls : null;
                lsError = err;
                lsInBase = found ?? new List<string>();
                ShowLs();
                RefreshState();
            });
        });
        th.IsBackground = true;
        th.Start();
    }

    // текст запрета: копия лицевого (в базе или в другом сохранённом договоре); null - можно
    string LsBlockText() {
        string ls = Digits(V("ls"));
        if (ls.Length == 0) return null;
        if (lsSaved != null) return "Лицевой счёт " + ls + " уже есть в сохранённом договоре «" + lsSaved + "».";
        if (lsChecked == ls && lsInBase.Count > 0)
            return "Лицевой " + BaseLink.Nab(ls) + " уже есть в базе: " + lsInBase[0] + (lsInBase.Count > 1 ? " (и ещё " + (lsInBase.Count - 1) + ")" : "") + ".";
        return null;
    }

    void ShowLs() {
        string ls = Digits(V("ls"));
        string block = LsBlockText();
        if (ls.Length == 0) Note(lsNote, "", St.Hint, false);
        else if (block != null) Note(lsNote, block + " Копию лицевого сохранить нельзя - она ломает базу.", St.Bad, true);
        else if (ls.Length != 13 && ls.Length != 9) Note(lsNote, "Лицевой - 13 цифр: приписка района (5002...) и номер в базе. Сейчас цифр: " + ls.Length, St.Hint, false);
        else if (baseLink == null) Note(lsNote, "База не подключена - лицевой по базе не проверен.", St.Hint, false);
        else if (lsError != null && lsChecked == null) Note(lsNote, "Проверить по базе не удалось: " + lsError, Warn, false);
        else if (lsChecked == ls) Note(lsNote, "Лицевого " + BaseLink.Nab(ls) + " в базе нет - можно оформлять.", St.Good, false);
    }

    // перед сохранением: ответ базы по этому лицевому обязателен (если база есть)
    string LsBlock(bool ask, out bool cancel) {
        cancel = false;
        string ls = Digits(V("ls"));
        if (ls.Length == 0) return null;
        lsSaved = SavedWithLs(ls);
        if (baseLink != null && lsChecked != ls && ls.Length >= 9) {
            Cursor = Cursors.WaitCursor;
            try {
                lsInBase = baseLink.Accounts(BaseLink.Nab(ls));
                lsChecked = ls;
                lsError = null;
            } catch (Exception ex) {
                lsError = ex.Message;
            } finally { Cursor = Cursors.Default; }
            ShowLs();
        }
        string block = LsBlockText();
        if (block != null) return block;
        if (ask && lsChecked != ls) {
            string why = baseLink == null ? "база недоступна" : "проверка не удалась" + (lsError != null ? " (" + lsError + ")" : "");
            if (MessageBox.Show(this, "Лицевой " + ls + " не проверен по базе: " + why + ".\n\nКопия лицевого ломает базу. Сохранить договор без проверки?",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) cancel = true;
        }
        return null;
    }

    // адрес для поиска в базе: из полей (как его понимает «Поиск»), а если строку адреса правили руками - она сама
    string AddressQuery() {
        if (V("adres") != autoAdres && V("adres").Length > 0) return V("adres");
        if (V("np").Length == 0 && V("ulica").Length == 0) return "";
        List<string> parts = new List<string>();
        if (V("np").Length > 0) parts.Add((V("np_tip") + " " + V("np")).Trim());
        if (V("ulica").Length > 0) parts.Add((V("ulica_tip") + " " + V("ulica")).Trim());
        if (V("dom").Length > 0) parts.Add("д. " + V("dom"));
        if (V("korp").Length > 0) parts.Add("корп. " + V("korp"));
        if (V("kv").Length > 0) parts.Add("кв. " + V("kv"));
        return string.Join(", ", parts.ToArray());
    }

    void CheckAdres() {
        string q = AddressQuery();
        if (q.Length == 0 || baseLink == null || V("dom").Length == 0 && V("adres") == autoAdres) { Note(adresNote, "", St.Hint, false); return; }
        int seq = ++adrSeq;
        Note(adresNote, "Проверяю адрес по базе...", St.Hint, false);
        BaseLink b = baseLink;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            List<string> found = null;
            string err = null, note = null;
            try { found = b.Address(q, out note); } catch (Exception ex) { err = ex.Message; }
            Post(delegate {
                if (seq != adrSeq) return;
                if (err != null) Note(adresNote, "Проверить адрес по базе не удалось: " + err, Warn, false);
                else if (found.Count > 0) {
                    List<string> show = found.GetRange(0, Math.Min(3, found.Count));
                    Note(adresNote, "По этому адресу в базе уже есть: " + string.Join("; ", show.ToArray())
                        + (found.Count > 3 ? " и ещё " + (found.Count - 3) : "") + ". Сохранить можно - просто проверьте.", Warn, true);
                } else if (note != null) Note(adresNote, "В справочнике базы адрес не найден: " + note, St.Hint, false);
                else Note(adresNote, "Абонентов по этому адресу в базе нет.", St.Good, false);
            });
        });
        th.IsBackground = true;
        th.Start();
    }

    // ---------- выбор счётчика: модель -> корпус -> модификация ----------

    List<string> Distinct(List<string> a) {
        List<string> r = new List<string>();
        foreach (string s in a) {
            bool found = false;
            foreach (string x in r) if (MeterType.Same(x, s)) found = true;
            if (!found && s.Length > 0) r.Add(s);
        }
        return r;
    }

    // списки корпусов и модификаций - только для выбранной модели и корпуса; clear - очистить выбор ниже
    void MeterLists(int i, bool clear) {
        string p = "pu" + i + "_";
        List<string> models = new List<string>(), cases = new List<string>(), mods = new List<string>();
        foreach (MeterType m in lists.Meters()) {
            models.Add(m.Model);
            if (!MeterType.Same(m.Model, V(p + "model"))) continue;
            cases.Add(m.Case);
            if (MeterType.Same(m.Case, V(p + "korpus"))) mods.Add(m.Mod);
        }
        models = Distinct(models); cases = Distinct(cases); mods = Distinct(mods);
        models.Add(Own); cases.Add(Own); mods.Add(Own);
        Items(p + "model", models);
        Items(p + "korpus", cases);
        Items(p + "modif", mods);
        if (clear) { SetV(p + "korpus", ""); SetV(p + "modif", ""); }
    }

    // полный тип: как в справочнике, если все части оттуда; иначе - части через пробел
    string MeterTip(int i) {
        string p = "pu" + i + "_", model = V(p + "model"), cs = V(p + "korpus"), mod = V(p + "modif");
        foreach (MeterType m in lists.Meters())
            if (MeterType.Same(m.Model, model) && MeterType.Same(m.Case, cs) && MeterType.Same(m.Mod, mod)) return m.Full;
        List<string> parts = new List<string>();
        foreach (string x in new string[] { model, cs, mod }) if (x.Length > 0) parts.Add(x);
        return string.Join(" ", parts.ToArray());
    }

    void MeterPart(int i, int level) {
        string p = "pu" + i + "_";
        string key = p + (level == 0 ? "model" : level == 1 ? "korpus" : "modif");
        Combo c = (Combo)inp[key];
        bool own = c.Text.Trim() == Own;
        // «свое» - пустое поле для своего значения; в списках ниже остаётся только «свое»
        if (own) BeginInvoke((MethodInvoker)delegate { loading = true; try { c.Text = ""; } finally { loading = false; } c.Focus(); });
        if (level == 0 && (own || !MeterType.Same(V(p + "model"), lastModel[i] ?? ""))) {
            lastModel[i] = V(p + "model");
            lastCase[i] = "";
            MeterLists(i, true);
            string phase = MeterType.Phase(V(p + "model"));
            if (phase.Length > 0) SetV(p + "faz", phase);
            // у модели один корпус - он и выбран
            Combo k = (Combo)inp[p + "korpus"];
            if (k.Items.Count == 2) { SetV(p + "korpus", k.Items[0].ToString()); level = 1; }
        }
        if (level == 1 && (own || !MeterType.Same(V(p + "korpus"), lastCase[i] ?? ""))) {
            lastCase[i] = V(p + "korpus");
            MeterLists(i, false);
            SetV(p + "modif", "");
            Combo m = (Combo)inp[p + "modif"];
            if (m.Items.Count == 2) SetV(p + "modif", m.Items[0].ToString());
        }
        // тип для договора и модем по модификации (если их не правили руками)
        string tip = MeterTip(i);
        if (V(p + "tip") == (autoTip[i] ?? "") || V(p + "tip").Length == 0) SetV(p + "tip", tip);
        autoTip[i] = tip;
        string modem = MeterType.Modem(V(p + "modif"));
        if (V(p + "modem") == (autoModem[i] ?? "") || V(p + "modem").Length == 0) SetV(p + "modem", modem);
        autoModem[i] = modem;
    }

    // ---------- проверка ----------

    class Problem {
        public string Key, Text;
        public bool Required, Blocker;      // Blocker - сохранить нельзя совсем (копия лицевого)
        public Problem(string key, string text, bool required) { Key = key; Text = text; Required = required; }
    }

    List<Problem> Check() {
        List<Problem> r = new List<Problem>();
        string[][] req = new string[][] {
            new string[] { "ls", "лицевой счёт" }, new string[] { "oformil", "кто оформил договор" },
            new string[] { "podp_dolzh", "должность подписанта" }, new string[] { "podp_fio", "Ф.И.О. подписанта" },
            new string[] { "podp_dov", "доверенность подписанта" }, new string[] { "podp_inic", "подпись подписанта" },
            new string[] { "soob_adres", "адрес для сообщений (п. 3.6)" }, new string[] { "soob_tel", "телефон для сообщений (п. 3.6)" },
            new string[] { "fio", "Ф.И.О. потребителя" }, new string[] { "potr_inic", "подпись потребителя" },
            new string[] { "lichny", "идентификационный номер" }, new string[] { "adres_prozh", "адрес проживания" },
            new string[] { "obekt", "объект электроснабжения" }, new string[] { "moshn", "разрешённая мощность" },
            new string[] { "np", "населённый пункт" }, new string[] { "adres", "адрес объекта" },
            new string[] { "prochie", "прочие условия (28.1)" },
            new string[] { "plita", "электроплита да/нет" }, new string[] { "gaz", "газоснабжение да/нет" },
            new string[] { "teplo", "теплоснабжение да/нет" }, new string[] { "gvs", "горячее водоснабжение да/нет" },
            new string[] { "sms", "СМС-оповещение: согласие или отказ" } };
        foreach (string[] q in req) if (V(q[0]).Length == 0) r.Add(new Problem(q[0], q[1], true));
        string block = LsBlockText();
        if (block != null) { Problem b = new Problem("ls", block, true); b.Blocker = true; r.Add(b); }
        if (V("sms") == "согласен(а)" && V("sms_tel").Length == 0) r.Add(new Problem("sms_tel", "номер для СМС", true));
        if (V("utratil") == "1" && V("utratil_nom").Length == 0) r.Add(new Problem("utratil_nom", "номер приложения, утратившего силу", true));
        int n = V("schetchikov") == "1 счётчик" ? 1 : 2;
        for (int i = 1; i <= n; i++) {
            string p = "pu" + i + "_", w = " (счётчик " + i + ")";
            string[][] mq = new string[][] {
                new string[] { "tip", "тип" }, new string[] { "zav", "заводской номер" }, new string[] { "cel", "цель использования" },
                new string[] { "faz", "фазность" }, new string[] { "askue", "АСКУЭ да/нет" }, new string[] { "tarif", "тариф" },
                new string[] { "gruppa", "тарифная группа" } };
            foreach (string[] q in mq) if (V(p + q[0]).Length == 0) r.Add(new Problem(p + q[0], q[1] + w, true));
        }
        // не обязательно, но стоит проверить
        string id = V("lichny");
        if (id.Length > 0 && id.Length != 14) r.Add(new Problem("lichny", "в идентификационном номере " + id.Length + " знаков, обычно 14", false));
        string ls = V("ls");
        foreach (char ch in ls) if (!char.IsDigit(ch)) { r.Add(new Problem("ls", "в лицевом счёте не только цифры", false)); break; }
        if (V("dopusk_nom").Length == 0) r.Add(new Problem("dopusk_nom", "не указан акт допуска в эксплуатацию", false));
        string g1 = V("pu1_gruppa");
        if (V("plita") == "да" && g1.Contains("НЕ ОБОРУДОВАННЫЙ")) r.Add(new Problem("pu1_gruppa", "плита есть, а тарифная группа - «НЕ оборудованный плитой»", false));
        if (V("plita") == "нет" && g1.Contains("ОБОРУДОВАННЫЙ") && !g1.Contains("НЕ ОБОРУДОВАННЫЙ")) r.Add(new Problem("pu1_gruppa", "плиты нет, а тарифная группа - «оборудованный плитой»", false));
        double all, heat;
        if (Num(V("moshn"), out all) && Num(V("moshn_otopl"), out heat) && heat > all) r.Add(new Problem("moshn_otopl", "мощность на отопление больше всей разрешённой", false));
        return r;
    }

    void RefreshState() {
        List<Problem> ps = Check();
        Dictionary<Card, int> per = new Dictionary<Card, int>();
        foreach (Card c in cards) per[c] = 0;
        Dictionary<string, bool> miss = new Dictionary<string, bool>();
        int required = 0;
        foreach (Problem p in ps) {
            if (!p.Required) continue;
            required++;
            miss[p.Key] = true;
            Card c;
            if (cardOf.TryGetValue(p.Key, out c)) per[c]++;
        }
        for (int i = 0; i < cards.Count && i < navs.Count; i++) {
            int cnt = cards[i].Name == "Порядок и документы" ? -1 : per[cards[i]];
            if (navs[i].Count != cnt) { navs[i].Count = cnt; navs[i].Invalidate(); }
        }
        foreach (KeyValuePair<string, Control> kv in inp) {
            bool m = showMissing && miss.ContainsKey(kv.Key);
            if (kv.Value is Seg) { if (((Seg)kv.Value).Missing != m) ((Seg)kv.Value).Missing = m; }
            else if (kv.Value is TextBox || kv.Value is ComboBox) {
                Color want = m ? St.Missing : (kv.Value.Enabled ? Color.White : St.Disabled);
                if (kv.Value.BackColor != want) kv.Value.BackColor = want;
            }
        }
        if (required > 0) {
            status.ForeColor = St.Bad;
            status.Text = LsBlockText() != null ? "Копия лицевого счёта - сохранить нельзя" :
                "Не заполнено полей: " + required + " - показать";
        } else {
            status.ForeColor = St.Good;
            status.Text = "Всё обязательное заполнено";
        }
    }

    static string Plural(int n, string one, string few, string many) {
        int a = n % 100, b = n % 10;
        if (a >= 11 && a <= 14) return many;
        return b == 1 ? one : b >= 2 && b <= 4 ? few : many;
    }

    void GoToMissing() {
        foreach (Problem p in Check()) if (p.Required) { FocusField(inp[p.Key]); return; }
    }

    void FocusField(Control c) {
        Card card;
        if (cardOf.TryGetValue(c.Name, out card)) {
            int top = card.Panel.Top - content.AutoScrollPosition.Y;
            // поле - на экран, с запасом сверху
            Point pt = content.PointToClient(c.Parent.PointToScreen(c.Location));
            int abs = pt.Y - content.AutoScrollPosition.Y;
            if (pt.Y < 60 || pt.Y > content.ClientSize.Height - 80) content.AutoScrollPosition = new Point(0, Math.Max(0, Math.Min(top - 4, abs - 120)));
        }
        if (c is Seg) ((Seg)c).Button(0).Focus();
        else if (c is DateBox) ((DateBox)c).Focus();
        else c.Focus();
    }

    // ---------- значения для шаблона ----------

    static bool Num(string s, out double v) {
        return double.TryParse(s.Replace(',', '.').Replace(" ", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out v);
    }

    static string Kw(string s, string empty) {
        double v;
        if (s.Length == 0) return empty;
        if (!Num(s, out v)) return s;
        return v.ToString("0.0##", CultureInfo.InvariantCulture).Replace('.', ',');
    }

    // "Колошич Елена Петровна" -> "Е.П.Колошич"; "Гаврош В.А." -> "В.А.Гаврош"
    public static string Initials(string fio) {
        string[] w = fio.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (w.Length == 0) return "";
        StringBuilder sb = new StringBuilder();
        int n = 0;
        for (int i = 1; i < w.Length && n < 2; i++)
            foreach (string part in w[i].Split('.'))
                if (part.Length > 0 && n < 2 && char.IsLetter(part[0])) { sb.Append(char.ToUpper(part[0])).Append('.'); n++; }
        return sb.ToString() + w[0];
    }

    string Address() {
        string tip = V("np_tip"), np = V("np");
        List<string> parts = new List<string>();
        bool st = tip.ToUpper() == "СТ" || tip.ToLower() == "с/т";
        if (np.Length > 0) parts.Add(st ? tip + " «" + np.Trim('«', '»', '"') + "»" : (tip + " " + np).Trim());
        if (V("ulica").Length > 0) parts.Add((V("ulica_tip") + " " + V("ulica")).Trim());
        string house = "";
        if (V("dom").Length > 0) house = (st ? "уч." : "д.") + V("dom");
        if (V("korp").Length > 0) house += (house.Length > 0 ? " " : "") + "корп." + V("korp");
        if (V("kv").Length > 0) house += (house.Length > 0 ? " " : "") + "кв." + V("kv");
        if (house.Length > 0) parts.Add(house);
        string s = string.Join(", ", parts.ToArray());
        if (V("selsovet").Length > 0) s += (s.Length > 0 ? " " : "") + "(" + V("selsovet") + " с/с)";
        return s;
    }

    static string Line(int n, char c) { return new string(c, n); }

    static string LowerFirst(string s) {
        if (s.Length < 2 || !char.IsUpper(s[0]) || char.IsUpper(s[1])) return s;
        return char.ToLower(s[0]) + s.Substring(1);
    }

    public Dictionary<string, string> Compose() {
        Dictionary<string, string> v = new Dictionary<string, string>();
        v["ДАТА"] = V("data");
        v["ЛС"] = V("ls");
        string tip = V("np_tip");
        v["НП"] = tip.ToUpper() == "СТ" || tip.ToLower() == "с/т" ? tip + " «" + V("np").Trim('«', '»', '"') + "»" : V("np");
        v["АДРЕС_ОБЪЕКТА"] = V("adres");
        v["ПОДП_В_ЛИЦЕ"] = (Flat(V("podp_dolzh")) + " " + Flat(V("podp_fio"))).Trim();
        v["ПОДП_ДОВЕРЕННОСТЬ"] = V("podp_dov");
        v["ПОДП_ИНИЦ"] = V("podp_inic");
        v["ОФОРМИЛ_ИНИЦ"] = Initials(V("oformil"));
        v["АДРЕС_СООБЩАТЬ"] = V("soob_adres");
        v["ТЕЛ_СООБЩАТЬ"] = V("soob_tel");
        v["ФИО"] = V("fio");
        v["ПОТР_ИНИЦ"] = V("potr_inic");
        v["ЛИЧНЫЙ_НОМЕР"] = V("lichny");
        v["АДРЕС_ПРОЖИВАНИЯ"] = V("adres_prozh");
        v["ТЕЛЕФОН"] = V("telefon").Length > 0 ? V("telefon") : Line(31, '_');
        v["ОБЪЕКТ"] = LowerFirst(V("obekt"));
        string kad = V("kadastr");
        bool digits = kad.Length > 0;
        foreach (char ch in kad) if (!char.IsDigit(ch)) digits = false;
        v["КАДАСТР"] = digits ? "к.н." + kad : kad;
        v["МОЩНОСТЬ"] = Kw(V("moshn"), "");
        v["МОЩНОСТЬ_ОТОПЛ"] = Kw(V("moshn_otopl"), "0,0");
        v["ДОПУСК"] = V("dopusk_nom").Length == 0 ? "" :
            (V("dopusk_nom") + " от " + V("dopusk_data") + " г. " + V("dopusk_kem")).Trim();
        string pr = V("prochie").TrimEnd('.', ' ');
        if (pr.Length == 0) v["ПРОЧИЕ_УСЛОВИЯ"] = "";
        else if (pr.ToLower() == "отсутствуют") v["ПРОЧИЕ_УСЛОВИЯ"] = " отсутствуют.";
        else v["ПРОЧИЕ_УСЛОВИЯ"] = "«" + char.ToUpper(pr[0]) + pr.Substring(1) + "».";

        bool two = V("schetchikov") != "1 счётчик";
        for (int i = 1; i <= 2; i++) {
            string p = "pu" + i + "_", f = "ПУ" + i + "_";
            if (i == 2 && !two) {
                string dash = Line(42, '-'), shortDash = Line(16, '-');
                // как в эталонном договоре с одним счётчиком
                v[f + "НОМ"] = "_";
                v[f + "НОМ_Т"] = "______" + shortDash;
                v[f + "ЦЕЛЬ"] = shortDash;
                foreach (string k in new string[] { "ТИП", "ЗАВ", "ПОВЕРКА", "ПЛОМБА_ГП", "ФАЗ", "МЕСТО", "ПРИНАДЛ", "МОДЕМ", "ПОКАЗ", "МОДУЛЬ" }) v[f + k] = dash;
                foreach (string k in new string[] { "ВВОД_ТИП", "ВВОД_I", "ЗАЩ_ТИП", "ЗАЩ_I" }) v[f + k] = shortDash;
                v["□" + f + "АСКУЭ_ДА"] = v["□" + f + "АСКУЭ_НЕТ"] = "0";
                v["□" + f + "ОДНОСТ"] = v["□" + f + "ДИФФ"] = "0";
                v[f + "ПЕРИОДЫ"] = "__" + shortDash + "_";
                v[f + "ТАРИФ"] = v[f + "ТАРИФ_Д"] = v[f + "ЗПТ"] = v[f + "ЗПТ_Д"] = "";
                v[f + "НАГРЕВ"] = v[f + "НАГРЕВ_Д"] = "";
                continue;
            }
            v[f + "НОМ"] = i.ToString();
            if (i == 2) v[f + "НОМ_Т"] = "_____2____";
            v[f + "ЦЕЛЬ"] = V(p + "cel");
            v[f + "ТИП"] = V(p + "tip");
            v[f + "ЗАВ"] = V(p + "zav");
            v[f + "ПОВЕРКА"] = V(p + "poverka");
            v[f + "ПЛОМБА_ГП"] = V(p + "plomba_gp");
            v[f + "ФАЗ"] = V(p + "faz");
            v[f + "МЕСТО"] = V(p + "mesto");
            v[f + "ПРИНАДЛ"] = V(p + "prinadl");
            v[f + "МОДЕМ"] = V(p + "modem");
            v[f + "ПОКАЗ"] = V(p + "pokaz");
            v[f + "ВВОД_ТИП"] = V(p + "vvod_tip");
            v[f + "ВВОД_I"] = NoInput(V(p + "vvod_tip")) ? Line(16, '-') : V(p + "vvod_i");
            v[f + "ЗАЩ_ТИП"] = V(p + "zasch_tip");
            v[f + "ЗАЩ_I"] = V(p + "zasch_i");
            v[f + "МОДУЛЬ"] = V(p + "modul").Length > 0 ? V(p + "modul") : Line(42, '-');
            v["□" + f + "АСКУЭ_ДА"] = V(p + "askue") == "да" ? "1" : "0";
            v["□" + f + "АСКУЭ_НЕТ"] = V(p + "askue") == "нет" ? "1" : "0";
            string per = Periods(V(p + "tarif")), group = V(p + "gruppa");
            bool diff = per.Length > 0;
            v["□" + f + "ОДНОСТ"] = V(p + "tarif") == Tarifs[0] ? "1" : "0";
            v["□" + f + "ДИФФ"] = diff ? "1" : "0";
            v[f + "ПЕРИОДЫ"] = diff ? per : Line(11, '_');
            // тарифная группа (жирным, после запятой) и электронагреватели - под выбранным тарифом
            string heat = V(p + "nagrev"), comma = group.Length > 0 ? ", " : "";
            v[f + "ЗПТ"] = diff ? "" : comma;
            v[f + "ТАРИФ"] = diff ? "" : group;
            v[f + "ЗПТ_Д"] = diff ? comma : "";
            v[f + "ТАРИФ_Д"] = diff ? group : "";
            v[f + "НАГРЕВ"] = diff ? "" : heat;
            v[f + "НАГРЕВ_Д"] = diff ? heat : "";
        }

        v["ПЛОМБА_ВВОД"] = V("plomba_vvod").Length > 0 ? V("plomba_vvod") : Line(16, '-');
        v["ПЛОМБА_КЛЕММ"] = V("plomba_klemm").Length > 0 ? V("plomba_klemm") : Line(16, '-');
        v["ПЛОМБА_ДРУГОЕ"] = V("plomba_drugoe").Length > 0 ? V("plomba_drugoe") : Line(16, '-');
        foreach (string[] q in new string[][] { new string[] { "plita", "ПЛИТА" }, new string[] { "gaz", "ГАЗ" }, new string[] { "teplo", "ТЕПЛО" }, new string[] { "gvs", "ГВС" } }) {
            v["□" + q[1] + "_ДА"] = V(q[0]) == "да" ? "1" : "0";
            v["□" + q[1] + "_НЕТ"] = V(q[0]) == "нет" ? "1" : "0";
        }
        v["ДОП_ИНФО"] = V("dop_info").Length > 0 ? V("dop_info") : Line(94, '_');
        v["ЛЬГОТА"] = V("lgota").Length > 0 ? V("lgota") : Line(78, '_');
        string os = V("osobye");
        if (V("utratil") == "1") {
            string u = "Считать утратившим силу приложение к договору № " + V("utratil_nom") + " от " + V("utratil_ot") + ".";
            os = os.Length > 0 ? os + "\n" + u : u;
        }
        v["ОСОБЫЕ"] = os.Length > 0 ? os : Line(93, '_');
        bool sms = V("sms") == "согласен(а)";
        v["□СМС_ДА"] = sms ? "1" : "0";
        v["□СМС_НЕТ"] = V("sms") == "отказ" ? "1" : "0";
        v["СМС_ТЕЛ"] = sms && V("sms_tel").Length > 0 ? V("sms_tel") : Line(13, '_');
        return v;
    }

    // ---------- создать, открыть, новый ----------

    string FileName() {
        string fio = V("fio");
        string[] w = fio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        string who = w.Length > 0 ? Short(fio) : "";
        string name = (V("ls") + " " + who).Trim();
        if (name.Length == 0) name = "Договор " + DateTime.Now.ToString("yyyy-MM-dd HH-mm");
        foreach (char ch in Path.GetInvalidFileNameChars()) name = name.Replace(ch, '_');
        return name.TrimEnd('.');
    }

    string Title() { return ("Договор " + V("ls") + " " + Short(V("fio"))).Trim(); }

    // для проверок без окон: create без вопросов
    public string Create(bool openWord) { return Create(openWord, openWord); }

    // ask - с вопросами (проверка заполнения, замена файла); openWord - открыть готовый договор в Word
    public string Create(bool openWord, bool ask) {
        // копия лицевого счёта ломает базу - такой договор не сохраняется совсем
        bool cancel;
        string block = LsBlock(ask, out cancel);
        if (cancel) return null;
        if (block != null) {
            if (ask) MessageBox.Show(this, block + "\n\nДоговор не сохранён.", "Копия лицевого счёта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            FocusField(inp["ls"]);
            return null;
        }
        List<Problem> ps = Check();
        StringBuilder req = new StringBuilder(), warn = new StringBuilder();
        foreach (Problem p in ps) if (!p.Blocker) (p.Required ? req : warn).Append("  • ").Append(p.Text).Append('\n');
        if (req.Length > 0 || warn.Length > 0) {
            showMissing = true;
            RefreshState();
            string msg = (req.Length > 0 ? "Не заполнено:\n" + req : "") + (warn.Length > 0 ? (req.Length > 0 ? "\n" : "") + "Проверьте:\n" + warn : "") +
                "\nВсё равно сохранить договор? (Незаполненное можно дописать в Word.)";
            if (ask && MessageBox.Show(this, msg, "Проверка перед сохранением", MessageBoxButtons.YesNo,
                    req.Length > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) {
                foreach (Problem p in ps) { FocusField(inp[p.Key]); break; }
                return null;
            }
        }
        if (!File.Exists(templateFile)) {
            MessageBox.Show(this, "Нет файла шаблона:\n" + templateFile + "\n\nОн должен лежать рядом с программой.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }
        string name = FileName();
        string docx = Path.Combine(outDir, name + ".docx");
        try {
            Directory.CreateDirectory(outDir);
            Directory.CreateDirectory(dataDir);
            string data = Path.Combine(dataDir, name + ".txt");
            bool same = openedFile != null && string.Compare(Path.GetFullPath(openedFile), Path.GetFullPath(data), StringComparison.OrdinalIgnoreCase) == 0;
            if (ask && !same && File.Exists(docx) && MessageBox.Show(this, "Договор «" + name + "» уже есть в папке.\n\nЗаменить его?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
            string tmp = docx + ".tmp";
            List<string> missing = DocxTemplate.Fill(templateFile, tmp, Compose());
            try {
                if (File.Exists(docx)) File.Delete(docx);
            } catch (IOException) {
                File.Delete(tmp);
                MessageBox.Show(this, "Договор «" + name + "» сейчас открыт в Word.\nЗакройте его в Word и нажмите «Сохранить договор» ещё раз.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            File.Move(tmp, docx);
            SaveData(data);
            Remember();
            dirty = false;
            openedFile = data;
            DropAuto();
            head.Text = Title();
            headSub.Text = "сохранён " + DateTime.Now.ToString("dd.MM.yyyy HH:mm") + " - " + docx;
            headSub.ForeColor = St.Hint;
            if (missing.Count > 0)
                MessageBox.Show(this, "В шаблоне есть поля, которые программа не знает (остались пустыми):\n" + string.Join(", ", missing.ToArray()), Text);
            if (openWord) {
                try { Process.Start(docx); }
                catch (Exception ex) { MessageBox.Show(this, "Договор сохранён:\n" + docx + "\n\nОткрыть в Word не получилось: " + ex.Message, Text); }
            }
            return docx;
        } catch (Exception ex) {
            MessageBox.Show(this, "Не удалось сохранить договор:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }
    }

    void Remember() {
        memory.Remember("нп", V("np"));
        memory.Remember("улица", V("ulica"));
        memory.Remember("сельсовет", V("selsovet"));
        memory.Remember("допуск", V("dopusk_kem"));
        for (int i = 1; i <= 2; i++) {
            if (i == 2 && V("schetchikov") == "1 счётчик") break;
            memory.Remember("место", V("pu" + i + "_mesto"));
        }
        foreach (string k in Sticky) memory.Values[k] = V(k);
        memory.Save();
        FillLists();
    }

    // эти поля переходят в следующий договор
    static readonly string[] Sticky = new string[] { "oformil", "podp", "podp_dolzh", "podp_fio", "podp_dov", "podp_inic", "soob_adres", "soob_tel", "dopusk_kem" };

    void SaveData(string file) {
        StringBuilder sb = new StringBuilder();
        sb.Append("# Данные договора для программы «Оформление договора»\r\n");
        sb.Append("сохранён=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
        foreach (KeyValuePair<string, Control> kv in inp) sb.Append(kv.Key).Append('=').Append(Memory.Escape(V(kv.Key))).Append("\r\n");
        File.WriteAllText(file, sb.ToString(), new UTF8Encoding(true));
    }

    public static Dictionary<string, string> ReadData(string file) {
        Dictionary<string, string> d = new Dictionary<string, string>();
        foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
            int eq = line.IndexOf('=');
            if (eq <= 0 || line.StartsWith("#")) continue;
            d[line.Substring(0, eq)] = Memory.Unescape(line.Substring(eq + 1));
        }
        return d;
    }

    // несохранённый договор: сохранить (да), бросить (нет) или вернуться к нему (отмена). true - можно идти дальше
    bool AskSave(string action) {
        if (!dirty) { DropAuto(); return true; }
        DialogResult r = MessageBox.Show(this, "Сохранить договор" + (V("ls").Length > 0 || V("fio").Length > 0 ? " «" + Title().Replace("Договор ", "") + "»" : "") + "?\n\n"
            + "Да - сохранить и " + action.ToLower() + "\nНет - не сохранять\nОтмена - вернуться к договору", Text,
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) return Create(false, true) != null;
        DropAuto();
        return true;
    }

    void AskNew() {
        if (!AskSave("Создать новый")) return;
        NewContract();
    }

    // ---------- автосохранение ----------

    void AutoSave() {
        if (!dirty) return;
        try {
            StringBuilder sb = new StringBuilder();
            sb.Append("# Автосохранение несохранённого договора (программа «Оформление договора»)\r\n");
            sb.Append("автосохранение=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
            if (openedFile != null) sb.Append("открыт=").Append(Memory.Escape(openedFile)).Append("\r\n");
            foreach (KeyValuePair<string, Control> kv in inp) sb.Append(kv.Key).Append('=').Append(Memory.Escape(V(kv.Key))).Append("\r\n");
            // сначала во временный файл, потом подмена: свет погаснет посреди записи - старое автосохранение цело
            string tmp = autoFile + ".tmp";
            File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(true));
            if (File.Exists(autoFile)) File.Replace(tmp, autoFile, null);
            else File.Move(tmp, autoFile);
            unsavedAuto = false;
        } catch { }
    }

    void DropAuto() {
        unsavedAuto = false;
        try { if (File.Exists(autoFile)) File.Delete(autoFile); } catch { }
    }

    // после отключения света - продолжить с последнего автосохранения
    void RestoreAuto() {
        if (!File.Exists(autoFile)) return;
        try {
            Dictionary<string, string> d = ReadData(autoFile);
            LoadFile(autoFile);
            string when, from;
            openedFile = d.TryGetValue("открыт", out from) && from.Length > 0 ? from : null;
            dirty = true;
            headSub.Text = "Восстановлен несохранённый договор (автосохранение " + (d.TryGetValue("автосохранение", out when) ? when : "") +
                "). Проверьте и нажмите «Сохранить договор»";
            headSub.ForeColor = St.Bad;
        } catch { }
    }

    public void TestAutoSave() { AutoSave(); }

    public void NewContract() {
        loading = true;
        try {
            foreach (KeyValuePair<string, Control> kv in inp) {
                if (kv.Value is DateBox) ((DateBox)kv.Value).Value = DateTime.Today;
                else SetV(kv.Key, "");
            }
            foreach (string k in Sticky) SetV(k, memory.Get(k, ""));
            if (V("soob_adres").Length == 0) SetV("soob_adres", "Руссиянова 24");
            if (V("soob_tel").Length == 0) SetV("soob_tel", "275-81-38");
            if (V("dopusk_kem").Length == 0) { List<string> k = lists.Get("Кем выдан допуск"); if (k.Count > 0) SetV("dopusk_kem", k[0]); }
            if (V("podp").Length == 0) { List<string[]> s = lists.Signers(); if (s.Count > 0) { SetV("podp", s[0][0]); SetV("podp_dolzh", s[0][1]); SetV("podp_fio", s[0][2]); SetV("podp_dov", s[0][3]); SetV("podp_inic", s[0][4]); } }
            SetV("np_tip", "д.");
            SetV("ulica_tip", "ул.");
            SetV("schetchikov", "1 счётчик");
            SetV("moshn_otopl", "0,0");
            List<string> o = lists.Get("Объект электроснабжения");
            if (o.Count > 0) SetV("obekt", o[0]);
            List<string> pr = lists.Get("Прочие условия");
            if (pr.Count > 0) SetV("prochie", pr[0]);
            SetV("gaz", "нет");
            SetV("teplo", "нет");
            SetV("gvs", "нет");
            for (int i = 1; i <= 2; i++) {
                string p = "pu" + i + "_";
                List<string> cel = lists.Get("Цель использования");
                if (cel.Count > 0) SetV(p + "cel", cel[i == 2 && cel.Count > 1 ? 1 : 0]);
                SetV(p + "tarif", Tarifs[0]);
                List<string> pl = lists.Get("Пломба госповерителя");
                if (pl.Count > 0) SetV(p + "plomba_gp", pl[0]);
                List<string> own = lists.Get("Принадлежность");
                if (own.Count > 0) SetV(p + "prinadl", own[0]);
            }
            autoInic = "";
            autoAdres = "";
            for (int i = 1; i <= 2; i++) {
                lastModel[i] = lastCase[i] = autoTip[i] = autoModem[i] = "";
                MeterLists(i, false);
            }
        } finally { loading = false; }
        showMissing = false;
        dirty = false;
        openedFile = null;
        head.Text = "Новый договор";
        headSub.Text = "Заполните поля - программа подставит их в шаблон Word и сохранит договор в папку «Договоры»";
        headSub.ForeColor = St.Hint;
        Enable();
        RefreshState();
        Recheck(true, true);
        content.AutoScrollPosition = new Point(0, 0);
    }

    public void LoadFile(string file) {
        Dictionary<string, string> d = ReadData(file);
        NewContract();
        loading = true;
        try {
            foreach (KeyValuePair<string, string> kv in d) if (inp.ContainsKey(kv.Key)) SetV(kv.Key, kv.Value);
            autoInic = Initials(V("fio"));
            autoAdres = Address();
            for (int i = 1; i <= 2; i++) {
                string p = "pu" + i + "_";
                lastModel[i] = V(p + "model");
                lastCase[i] = V(p + "korpus");
                MeterLists(i, false);
                autoTip[i] = MeterTip(i);
                autoModem[i] = MeterType.Modem(V(p + "modif"));
            }
        } finally { loading = false; }
        dirty = false;
        openedFile = file;
        head.Text = Title();
        string saved;
        headSub.Text = "открыт сохранённый" + (d.TryGetValue("сохранён", out saved) ? " (" + saved + ")" : "") +
            ". Исправьте что нужно и нажмите «Сохранить договор»";
        headSub.ForeColor = St.Hint;
        Enable();
        RefreshState();
        Recheck(true, true);
    }

    void OpenSaved() {
        if (!AskSave("Открыть другой")) return;
        using (SavedDialog dlg = new SavedDialog(dataDir)) {
            if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Chosen != null) {
                try { LoadFile(dlg.Chosen); }
                catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть:\n" + ex.Message, Text); }
            }
        }
    }

    // для проверки: прокрутить к полю
    public void Reveal(string key) {
        Control c = inp[key];
        Point pt = content.PointToClient(c.Parent.PointToScreen(c.Location));
        content.AutoScrollPosition = new Point(0, Math.Max(0, pt.Y - content.AutoScrollPosition.Y - 120));
    }

    public void ScrollToCard(int card) { ScrollTo(cards[card]); }
}

// Список сохранённых договоров с поиском
public class SavedDialog : Form {
    public string Chosen;
    readonly ListView list = new ListView();
    readonly TextBox find = Ui.Text(1);
    readonly List<string[]> rows = new List<string[]>();

    public SavedDialog(string dir) {
        Text = "Сохранённые договоры";
        Font = St.Base;
        BackColor = Color.White;
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(980, 600);
        MinimizeBox = false;
        ShowIcon = false;
        ShowInTaskbar = false;
        Label cap = Ui.Caption("Поиск: лицевой счёт, фамилия или адрес");
        cap.SetBounds(16, 12, 400, 20);
        find.SetBounds(16, 34, 940, 26);
        find.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
        find.TextChanged += delegate { Filter(); };
        find.KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Down && list.Items.Count > 0) { list.Focus(); list.Items[0].Selected = true; e.Handled = true; }
            if (e.KeyCode == Keys.Enter) { Pick(); e.SuppressKeyPress = true; }
        };
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.BorderStyle = BorderStyle.FixedSingle;
        list.SetBounds(16, 72, 940, 430);
        list.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
        list.Columns.Add("Дата", 95);
        list.Columns.Add("Лицевой счёт", 140);
        list.Columns.Add("Потребитель", 250);
        list.Columns.Add("Адрес объекта", 300);
        list.Columns.Add("Сохранён", 130);
        list.DoubleClick += delegate { Pick(); };
        list.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { Pick(); e.SuppressKeyPress = true; } };
        Button ok = Ui.Button("Открыть", true);
        ok.SetBounds(756, 512, 200, 38);
        ok.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        ok.Click += delegate { Pick(); };
        Button cancel = Ui.Button("Отмена", false);
        cancel.SetBounds(636, 512, 110, 38);
        cancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        cancel.DialogResult = DialogResult.Cancel;
        CancelButton = cancel;
        Controls.Add(cap);
        Controls.Add(find);
        Controls.Add(list);
        Controls.Add(ok);
        Controls.Add(cancel);
        try {
            if (Directory.Exists(dir)) {
                List<KeyValuePair<DateTime, string[]>> all = new List<KeyValuePair<DateTime, string[]>>();
                foreach (string f in Directory.GetFiles(dir, "*.txt")) {
                    try {
                        Dictionary<string, string> d = DogovorForm.ReadData(f);
                        string[] r = new string[] { Get(d, "data"), Get(d, "ls"), Get(d, "fio"), Get(d, "adres"), Get(d, "сохранён"), f };
                        all.Add(new KeyValuePair<DateTime, string[]>(File.GetLastWriteTime(f), r));
                    } catch { }
                }
                all.Sort(delegate(KeyValuePair<DateTime, string[]> a, KeyValuePair<DateTime, string[]> b) { return b.Key.CompareTo(a.Key); });
                foreach (KeyValuePair<DateTime, string[]> kv in all) rows.Add(kv.Value);
            }
        } catch { }
        Filter();
        Shown += delegate { find.Focus(); };
    }

    static string Get(Dictionary<string, string> d, string k) {
        string v;
        return d.TryGetValue(k, out v) ? v : "";
    }

    void Filter() {
        string q = find.Text.Trim().ToLower();
        list.BeginUpdate();
        list.Items.Clear();
        foreach (string[] r in rows) {
            if (q.Length > 0 && (r[1] + " " + r[2] + " " + r[3]).ToLower().IndexOf(q) < 0) continue;
            ListViewItem it = new ListViewItem(new string[] { r[0], r[1], r[2], r[3], r[4] });
            it.Tag = r[5];
            list.Items.Add(it);
        }
        if (list.Items.Count == 0) {
            ListViewItem it = new ListViewItem(new string[] { "", "", rows.Count == 0 ? "Сохранённых договоров пока нет" : "Ничего не найдено", "", "" });
            it.ForeColor = St.Hint;
            list.Items.Add(it);
        }
        list.EndUpdate();
    }

    void Pick() {
        ListViewItem it = list.SelectedItems.Count > 0 ? list.SelectedItems[0] : (list.Items.Count == 1 ? list.Items[0] : null);
        if (it == null || it.Tag == null) return;
        Chosen = (string)it.Tag;
        DialogResult = DialogResult.OK;
        Close();
    }
}

public static class DogovorApp {
    public static void Run(string here, string launch) {
        try { Application.EnableVisualStyles(); } catch { }
        using (DogovorForm f = new DogovorForm(here)) {
            if (launch != null) f.Tag = launch;
            Application.Run(f);
        }
    }
}
'@
# Программа собирается из этого скрипта ОДИН раз после обновления и сохраняется готовой (файл dogovor_*.dll
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
$name = "dogovor_$stamp.dll"
$cacheDirs = @()
if ($env:LOCALAPPDATA) { $cacheDirs += (Join-Path $env:LOCALAPPDATA 'oformlenie-dogovora') }
if ($env:PUBLIC) { $cacheDirs += (Join-Path $env:PUBLIC 'oformlenie-dogovora') }
$cacheDirs += (Join-Path $here 'kesh-programmy')

function Test-Ready { return ($null -ne ('DogovorApp' -as [type])) }

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
        $splash.Text = 'Оформление договора'
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
            Get-ChildItem -Path $d -Filter 'dogovor_*.dll' -ErrorAction SilentlyContinue |
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

if (-not (Test-Ready)) {
    [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n" + ($errors -join "`n`n") + "`n`nСфотографируйте это окно и пришлите.", 'Оформление договора - ошибка')
    exit
}

try {
    [DogovorApp]::Run($here, $launch)
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("$_", 'Оформление договора - ошибка')
}
