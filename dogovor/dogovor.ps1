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

// Шаблон договора: поля «ИМЯ» заменяются текстом, «□ИМЯ» - флажком ☑ или ☐ (символы Wingdings, как в бланке).
// Поле может быть разбито Word-ом на несколько кусков текста - собираем его по всему абзацу.
public static class DocxTemplate {
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly Regex Field = new Regex("«(□?)([A-ZА-ЯЁ0-9_]+)»");

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
            foreach (XmlNode p in doc.SelectNodes("//w:p", ns))
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
            bool box = match.Groups[1].Value.Length > 0;
            string value;
            if (!values.TryGetValue((box ? "□" : "") + name, out value)) { value = ""; if (!missing.Contains(name)) missing.Add(name); }
            Put(doc, first, firstOffset, match.Length, value, box);
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
                    string n = m.Groups[1].Value + m.Groups[2].Value;
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
"[Тип счётчика]\r\n" +
"СЕ 318 R32.146.JR.UVFL\r\n" +
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
"[Аппарат]\r\n" +
"# вводной аппарат и аппарат защиты\r\n" +
"ВА\r\nВыключатель автоматический\r\nВыключатель нагрузки\r\nРубильник\r\nОтсутствует\r\n";
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
            Size s = TextRenderer.MeasureText(c.Text, c.Font, new Size(Math.Max(width, 10), int.MaxValue), TextFormatFlags.WordBreak);
            return s.Height + 4;
        }
        if (c is CheckBox) return Math.Max(24, c.Font.Height + 8);
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

public class DateBox : DateTimePicker {
    public DateBox() {
        Font = St.Base;
        Format = DateTimePickerFormat.Custom;
        CustomFormat = "dd.MM.yyyy";
    }
    protected override void OnMouseWheel(MouseEventArgs e) { Ui.Wheel(this, e); }
    public string Text2 { get { return Value.ToString("dd.MM.yyyy"); } }
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
    bool loading, showMissing, dirty;
    string autoInic = "", autoAdres = "";

    public DogovorForm(string here) {
        this.here = here;
        templateFile = Path.Combine(here, "shablon.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = new Memory(Path.Combine(here, "nastroiki.txt"));
        outDir = memory.Get("ПапкаДоговоров", "");
        if (outDir.Length == 0) outDir = Path.Combine(here, "Договоры");
        dataDir = Path.Combine(outDir, "данные");

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
        Shown += delegate { LayoutCards(); FillLists(); FocusField(inp["ls"]); };
        Activated += delegate {
            if (!lists.Changed()) return;
            try { lists.Load(); FillLists(); } catch { }
        };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.S) { e.SuppressKeyPress = true; Create(true); }
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
            if (dirty && MessageBox.Show(this, "Договор не создан - введённые данные пропадут.\n\nЗакрыть программу?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) e.Cancel = true;
        };
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
        else if (c is DateTimePicker) ((DateTimePicker)c).ValueChanged += h;
        else c.TextChanged += h;
    }

    public string V(string key) {
        Control c = inp[key];
        if (c is Seg) return ((Seg)c).Value;
        if (c is CheckBox) return ((CheckBox)c).Checked ? "1" : "";
        if (c is DateBox) return ((DateBox)c).Text2;
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
        cardOf["schetchikov"] = c;

        c = NewCard("Пломбы и оборудование", "", 4);
        Row(c, F("Пломба на вводном аппарате", T("plomba_vvod", 1, "пусто - прочерк"), 1), F("на клеммной крышке", T("plomba_klemm", 1, "пусто - прочерк"), 1),
            F("другие пломбы", T("plomba_drugoe", 1, "например ДСТП 1234567"), 2));
        Row(c, F("Электроплита стационарная", S("plita", "да", "нет"), 1), F("Центр. газоснабжение", S("gaz", "да", "нет"), 1),
            F("Центр. теплоснабжение", S("teplo", "да", "нет"), 1), F("Центр. горячее водоснабжение", S("gvs", "да", "нет"), 1));
        Row(c, F("Примечание к плите (под словом «да»)", T("plita_prim", 1, "например Эл. плита - 2,5 кВт"), 2), Cell.Empty(2));

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
        g.Row(F("Тип", C(p + "tip", true, false, "например СЕ 318"), 2));
        g.Row(F("Заводской номер", T(p + "zav", 1, null), 1), F("Дата поверки (год-квартал)", T(p + "poverka", 1, "например 2025-2"), 1));
        g.Row(F("Цель использования электроэнергии", C(p + "cel", true, false, null), 2));
        g.Row(F("Фазность", S(p + "faz", "1 ф", "3 ф"), 1), F("Пломба госповерителя", C(p + "plomba_gp", true, false, null), 1));
        g.Row(F("Место установки", C(p + "mesto", true, false, null), 2));
        g.Row(F("Принадлежность", C(p + "prinadl", true, false, null), 1), F("Тип модема", C(p + "modem", true, false, null), 1));
        g.Row(F("Включение в АСКУЭ", S(p + "askue", "да", "нет"), 1), F("Показания, кВт·ч", T(p + "pokaz", 1, null), 1));
        g.Row(F("Вводной аппарат: тип", C(p + "vvod_tip", true, false, null), 1), F("ток, А", T(p + "vvod_i", 1, null), 1));
        g.Row(F("Аппарат защиты: тип", C(p + "zasch_tip", true, false, null), 1), F("ток, А", T(p + "zasch_i", 1, null), 1));
        g.Row(F("Модуль отображения (сплит-счётчик), зав. №", T(p + "modul", 1, "нет модуля - прочерк"), 2));
        g.Row(F("Тариф", S(p + "tarif", Tarifs), 2));
        g.Row(F("Тарифная группа", C(p + "gruppa", true, true, null), 2));
        g.Row(F("Электронагревательные приборы (в договоре - под тарифом)", T(p + "nagrev", 2, null), 2));
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
        nav.Resize += delegate {
            dirs.SetBounds(0, nav.ClientSize.Height - 20 - dirs.Height, nav.Width, dirs.Height);
            folder.SetBounds(0, dirs.Top - folder.Height, nav.Width, folder.Height);
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
        Button nw = Ui.Button("Новый", false);
        nw.SetBounds(24, 14, 110, 38);
        nw.Click += delegate { AskNew(); };
        tip.SetToolTip(nw, "Очистить форму для нового договора (Ctrl+N)");
        Button open = Ui.Button("Открыть сохранённый…", false);
        open.SetBounds(144, 14, 200, 38);
        open.Click += delegate { OpenSaved(); };
        tip.SetToolTip(open, "Открыть ранее оформленный договор, чтобы исправить или сделать похожий (Ctrl+O)");
        Button create = Ui.Button("Создать договор", true);
        create.Size = new Size(210, 40);
        create.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        create.Click += delegate { Create(true); };
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
            Items(p + "tip", With(memory.Hist("тип счётчика"), lists.Get("Тип счётчика")));
            Items(p + "cel", lists.Get("Цель использования"));
            Items(p + "plomba_gp", lists.Get("Пломба госповерителя"));
            Items(p + "mesto", With(lists.Get("Место установки"), memory.Hist("место")));
            Items(p + "prinadl", lists.Get("Принадлежность"));
            Items(p + "modem", lists.Get("Тип модема"));
            Items(p + "vvod_tip", With(lists.Get("Аппарат"), memory.Hist("аппарат")));
            Items(p + "zasch_tip", With(lists.Get("Аппарат"), memory.Hist("аппарат")));
            Items(p + "gruppa", lists.Get("Тарифная группа"));
        }
    }

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
            if (key == "lichny") {
                TextBox t = (TextBox)inp["lichny"];
                int pos = t.SelectionStart;
                string up = t.Text.ToUpper().Replace(" ", "");
                if (up != t.Text) { t.Text = up; t.SelectionStart = Math.Min(pos, up.Length); }
            }
        } finally { loading = false; }
        if (key.EndsWith("tarif") || key == "schetchikov" || key == "sms" || key == "utratil" || key == "plita") Enable();
        RefreshState();
    }

    void Enable() {
        bool two = V("schetchikov") != "1 счётчик";
        meter2.Enabled = two;
        meter2.BackColor = two ? St.Box : St.Disabled;
        meter2.Invalidate();
        inp["sms_tel"].Enabled = V("sms") == "согласен(а)";
        inp["utratil_nom"].Enabled = inp["utratil_ot"].Enabled = V("utratil") == "1";
        inp["plita_prim"].Enabled = V("plita") == "да";
    }

    // ---------- проверка ----------

    class Problem {
        public string Key, Text;
        public bool Required;
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
            status.Text = "Не заполнено: " + required + " " + Plural(required, "поле", "поля", "полей") + " - нажмите, чтобы перейти к первому";
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
                v[f + "НОМ"] = "";
                v[f + "ЦЕЛЬ"] = shortDash;
                foreach (string k in new string[] { "ТИП", "ЗАВ", "ПОВЕРКА", "ПЛОМБА_ГП", "ФАЗ", "МЕСТО", "ПРИНАДЛ", "МОДЕМ", "ПОКАЗ", "МОДУЛЬ" }) v[f + k] = dash;
                foreach (string k in new string[] { "ВВОД_ТИП", "ВВОД_I", "ЗАЩ_ТИП", "ЗАЩ_I" }) v[f + k] = shortDash;
                v["□" + f + "АСКУЭ_ДА"] = v["□" + f + "АСКУЭ_НЕТ"] = "0";
                v["□" + f + "ОДНОСТ"] = v["□" + f + "ДИФФ"] = "0";
                v[f + "ПЕРИОДЫ"] = shortDash;
                v[f + "ТАРИФ"] = v[f + "ТАРИФ_Д"] = "";
                v[f + "НАГРЕВ"] = "";
                continue;
            }
            v[f + "НОМ"] = i.ToString();
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
            v[f + "ВВОД_I"] = V(p + "vvod_i");
            v[f + "ЗАЩ_ТИП"] = V(p + "zasch_tip");
            v[f + "ЗАЩ_I"] = V(p + "zasch_i");
            v[f + "МОДУЛЬ"] = V(p + "modul").Length > 0 ? V(p + "modul") : Line(42, '-');
            v["□" + f + "АСКУЭ_ДА"] = V(p + "askue") == "да" ? "1" : "0";
            v["□" + f + "АСКУЭ_НЕТ"] = V(p + "askue") == "нет" ? "1" : "0";
            string per = Periods(V(p + "tarif")), group = V(p + "gruppa").Length > 0 ? ", " + V(p + "gruppa") : "";
            bool diff = per.Length > 0;
            v["□" + f + "ОДНОСТ"] = V(p + "tarif") == Tarifs[0] ? "1" : "0";
            v["□" + f + "ДИФФ"] = diff ? "1" : "0";
            v[f + "ПЕРИОДЫ"] = diff ? per : Line(11, '_');
            // тарифная группа и электронагреватели - под выбранным тарифом
            string heat = V(p + "nagrev");
            v[f + "ТАРИФ"] = diff ? "" : group;
            v[f + "ТАРИФ_Д"] = !diff ? "" : heat.Length > 0 ? group + "\n" + heat : group;
            v[f + "НАГРЕВ"] = diff ? "" : i == 2 && heat.Length > 0 ? "\n" + heat : heat;
        }

        v["ПЛОМБА_ВВОД"] = V("plomba_vvod").Length > 0 ? V("plomba_vvod") : Line(16, '-');
        v["ПЛОМБА_КЛЕММ"] = V("plomba_klemm").Length > 0 ? V("plomba_klemm") : Line(16, '-');
        v["ПЛОМБА_ДРУГОЕ"] = V("plomba_drugoe").Length > 0 ? V("plomba_drugoe") : Line(16, '-');
        foreach (string[] q in new string[][] { new string[] { "plita", "ПЛИТА" }, new string[] { "gaz", "ГАЗ" }, new string[] { "teplo", "ТЕПЛО" }, new string[] { "gvs", "ГВС" } }) {
            v["□" + q[1] + "_ДА"] = V(q[0]) == "да" ? "1" : "0";
            v["□" + q[1] + "_НЕТ"] = V(q[0]) == "нет" ? "1" : "0";
        }
        v["ПЛИТА_ПРИМ"] = V("plita") == "да" && V("plita_prim").Length > 0 ? "\n" + V("plita_prim") : "";
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

    public string Create(bool openWord) {
        List<Problem> ps = Check();
        StringBuilder req = new StringBuilder(), warn = new StringBuilder();
        foreach (Problem p in ps) (p.Required ? req : warn).Append("  • ").Append(p.Text).Append('\n');
        if (req.Length > 0 || warn.Length > 0) {
            showMissing = true;
            RefreshState();
            string msg = (req.Length > 0 ? "Не заполнено:\n" + req : "") + (warn.Length > 0 ? (req.Length > 0 ? "\n" : "") + "Проверьте:\n" + warn : "") +
                "\nВсё равно создать договор? (Незаполненное можно дописать в Word.)";
            if (openWord && MessageBox.Show(this, msg, "Проверка перед созданием", MessageBoxButtons.YesNo,
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
            if (openWord && File.Exists(docx) && MessageBox.Show(this, "Договор «" + name + "» уже есть в папке.\n\nЗаменить его?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
            string tmp = docx + ".tmp";
            List<string> missing = DocxTemplate.Fill(templateFile, tmp, Compose());
            try {
                if (File.Exists(docx)) File.Delete(docx);
            } catch (IOException) {
                File.Delete(tmp);
                MessageBox.Show(this, "Договор «" + name + "» сейчас открыт в Word.\nЗакройте его в Word и нажмите «Создать договор» ещё раз.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            File.Move(tmp, docx);
            SaveData(Path.Combine(dataDir, name + ".txt"));
            Remember();
            dirty = false;
            head.Text = Title();
            headSub.Text = "создан " + DateTime.Now.ToString("dd.MM.yyyy HH:mm") + " - " + docx;
            if (missing.Count > 0)
                MessageBox.Show(this, "В шаблоне есть поля, которые программа не знает (остались пустыми):\n" + string.Join(", ", missing.ToArray()), Text);
            if (openWord) {
                try { Process.Start(docx); }
                catch (Exception ex) { MessageBox.Show(this, "Договор сохранён:\n" + docx + "\n\nОткрыть в Word не получилось: " + ex.Message, Text); }
            }
            return docx;
        } catch (Exception ex) {
            MessageBox.Show(this, "Не удалось создать договор:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            memory.Remember("тип счётчика", V("pu" + i + "_tip"));
            memory.Remember("место", V("pu" + i + "_mesto"));
            memory.Remember("аппарат", V("pu" + i + "_vvod_tip"));
            memory.Remember("аппарат", V("pu" + i + "_zasch_tip"));
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

    void AskNew() {
        if (dirty && MessageBox.Show(this, "Договор не создан - введённые данные пропадут.\n\nНачать новый?", Text,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        NewContract();
    }

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
        } finally { loading = false; }
        showMissing = false;
        dirty = false;
        head.Text = "Новый договор";
        headSub.Text = "Заполните поля - программа подставит их в шаблон Word и сохранит договор в папку «Договоры»";
        Enable();
        RefreshState();
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
        } finally { loading = false; }
        dirty = false;
        head.Text = Title();
        string saved;
        headSub.Text = "открыт сохранённый" + (d.TryGetValue("сохранён", out saved) ? " (" + saved + ")" : "") +
            ". Исправьте что нужно и нажмите «Создать договор»";
        Enable();
        RefreshState();
    }

    void OpenSaved() {
        if (dirty && MessageBox.Show(this, "Договор не создан - введённые данные пропадут.\n\nОткрыть другой?", Text,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        using (SavedDialog dlg = new SavedDialog(dataDir)) {
            if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Chosen != null) {
                try { LoadFile(dlg.Chosen); }
                catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть:\n" + ex.Message, Text); }
            }
        }
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
