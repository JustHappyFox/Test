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
using System.Drawing.Imaging;
using System.Drawing.Text;
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

    // values: имя поля -> текст ("\n" - перенос строки, "\t" - табуляция); для флажков "1" - отмечен. Возвращает поля, которых не было в values
    public static List<string> Fill(string template, string output, Dictionary<string, string> values) {
        return Fill(template, output, values, null);
    }

    // sizes: имя поля -> размер шрифта значения в полупунктах (длинный текст в узкой ячейке - мельче, чтобы таблица не росла)
    public static List<string> Fill(string template, string output, Dictionary<string, string> values, Dictionary<string, int> sizes) {
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
                if (FillParagraph(doc, p, values, sizes, missing)) changed = true;
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

    static bool FillParagraph(XmlDocument doc, XmlNode p, Dictionary<string, string> values, Dictionary<string, int> sizes, List<string> missing) {
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
            XmlNode run = first.ParentNode;
            Put(doc, first, firstOffset, match.Length, value, box);
            int size;
            if (!box && sizes != null && sizes.TryGetValue(name, out size)) Resize(doc, run, p, size);
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
            string[] cells = lines[i].Split('\t');
            for (int j = 0; j < cells.Length; j++) {
                if (j > 0) parts.Add(doc.CreateElement("w", "tab", W));
                if (cells[j].Length > 0) parts.Add(MakeText(doc, cells[j]));
            }
        }
        if (after.Length > 0) parts.Add(MakeText(doc, after));
        foreach (XmlNode n in parts) run.InsertBefore(n, t);
        run.RemoveChild(t);
    }

    // размер шрифта куска (полупункты); точный межстрочный интервал абзаца - под новый размер
    static void Resize(XmlDocument doc, XmlNode run, XmlNode p, int size) {
        XmlNode rPr = null;
        foreach (XmlNode c in run.ChildNodes) if (c.LocalName == "rPr") rPr = c;
        if (rPr == null) { rPr = doc.CreateElement("w", "rPr", W); run.PrependChild(rPr); }
        foreach (string n in new string[] { "sz", "szCs" }) {
            XmlElement e = null;
            foreach (XmlNode c in rPr.ChildNodes) if (c.LocalName == n) e = (XmlElement)c;
            if (e == null) {
                e = doc.CreateElement("w", n, W);
                // порядок в rPr: ... sz, szCs, highlight, u ...
                XmlNode before = null;
                foreach (XmlNode c in rPr.ChildNodes)
                    if (before == null && (c.LocalName == "highlight" || c.LocalName == "u" || c.LocalName == "effect" || c.LocalName == "bdr" ||
                        c.LocalName == "shd" || c.LocalName == "fitText" || c.LocalName == "vertAlign" || c.LocalName == "rtl" ||
                        c.LocalName == "cs" || c.LocalName == "em" || c.LocalName == "lang" || c.LocalName == "eastAsianLayout" ||
                        (n == "sz" && c.LocalName == "szCs"))) before = c;
                rPr.InsertBefore(e, before);
            }
            e.SetAttribute("val", W, size.ToString());
        }
        foreach (XmlNode c in p.ChildNodes) {
            if (c.LocalName != "pPr") continue;
            foreach (XmlNode sp in c.ChildNodes) {
                XmlElement e = sp as XmlElement;
                if (e == null || e.LocalName != "spacing" || e.GetAttribute("lineRule", W) != "exact") continue;
                e.SetAttribute("line", W, Math.Min(int.Parse(e.GetAttribute("line", W)), size * 12).ToString());
            }
        }
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

// Ширина текста в шрифте бланка (Times New Roman) - чтобы решить, влезет ли значение на своё место.
// Ширины букв - из таблицы (единицы 1/2048 кегля, как в самом шрифте; сняты с метрически совместимого
// Liberation Serif): так одинаково на любом компьютере, без измерения на экране
public static class Fit {
    const string Chars = " !\u0022#%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[]_abcdefghijklmnopqrstuvwxyz\u00AB\u00BB\u00A0\u2013\u2014\u2116\u0401\u0410\u0411\u0412\u0413\u0414\u0415\u0416\u0417\u0418\u0419\u041A\u041B\u041C\u041D\u041E\u041F\u0420\u0421\u0422\u0423\u0424\u0425\u0426\u0427\u0428\u0429\u042A\u042B\u042C\u042D\u042E\u042F\u0430\u0431\u0432\u0433\u0434\u0435\u0436\u0437\u0438\u0439\u043A\u043B\u043C\u043D\u043E\u043F\u0440\u0441\u0442\u0443\u0444\u0445\u0446\u0447\u0448\u0449\u044A\u044B\u044C\u044D\u044E\u044F\u0451\u0406\u0456\u040E\u045E";
    static readonly int[] Regular = {
        512, 682, 836, 1024, 1706, 1593, 369, 682, 682, 1024, 1155, 512, 682, 512, 569, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024,
        1024, 569, 569, 1155, 1155, 1155, 909, 1886, 1479, 1366, 1366, 1479, 1251, 1139, 1479, 1479, 682, 797, 1479, 1251, 1821, 1479, 1479, 1139,
        1479, 1366, 1139, 1251, 1479, 1479, 1933, 1479, 1479, 1251, 682, 682, 1024, 909, 1024, 909, 1024, 909, 682, 1024, 1024, 569, 569, 1024,
        569, 1593, 1024, 1024, 1024, 1024, 682, 797, 569, 1024, 1024, 1479, 1024, 1024, 909, 1024, 1024, 512, 1024, 2048, 1954, 1251, 1479, 1176,
        1366, 1184, 1397, 1251, 1835, 1026, 1479, 1479, 1366, 1389, 1821, 1479, 1479, 1479, 1139, 1366, 1251, 1450, 1618, 1479, 1479, 1331, 2066, 2066,
        1446, 1786, 1176, 1352, 2105, 1366, 909, 1042, 967, 840, 1042, 909, 1415, 809, 1096, 1096, 995, 1022, 1296, 1096, 1024, 1096, 1024, 909,
        895, 1024, 1327, 1024, 1096, 1030, 1577, 1577, 1059, 1376, 934, 879, 1530, 942, 909, 682, 569, 1450, 1024 };
    static readonly int[] Bold = {
        512, 682, 1137, 1024, 2048, 1706, 569, 682, 682, 1024, 1167, 512, 682, 512, 569, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024,
        1024, 682, 682, 1167, 1167, 1167, 1024, 1905, 1479, 1366, 1479, 1479, 1366, 1251, 1593, 1593, 797, 1024, 1593, 1366, 1933, 1479, 1593, 1251,
        1593, 1479, 1139, 1366, 1479, 1479, 2048, 1479, 1479, 1366, 682, 682, 1024, 1024, 1139, 909, 1139, 909, 682, 1024, 1139, 569, 682, 1139,
        569, 1706, 1139, 1024, 1139, 1139, 909, 797, 682, 1139, 1024, 1479, 1024, 1024, 909, 1024, 1024, 512, 1024, 2048, 2058, 1366, 1479, 1354,
        1366, 1303, 1409, 1366, 2025, 1081, 1593, 1593, 1485, 1526, 1933, 1593, 1593, 1593, 1251, 1479, 1366, 1503, 1759, 1479, 1593, 1503, 2249, 2249,
        1567, 2011, 1354, 1389, 2306, 1479, 1024, 1024, 1106, 930, 1036, 909, 1485, 823, 1180, 1180, 1180, 1149, 1395, 1180, 1024, 1180, 1139, 909,
        1006, 1024, 1417, 1024, 1180, 1155, 1729, 1729, 1194, 1599, 1083, 885, 1565, 1108, 909, 797, 569, 1503, 1024 };

    // ширина строки, пт
    public static float Width(string s, float pt, bool bold) {
        int[] w = bold ? Bold : Regular;
        long sum = 0;
        foreach (char c in s) {
            int i = Chars.IndexOf(c);
            sum += i >= 0 ? w[i] : 1100;     // незнакомый знак - как средняя заглавная
        }
        return sum * pt / 2048f;
    }

    // сколько строк займёт текст в колонке шириной width пт (перенос по словам)
    public static int Lines(string s, float pt, float width) {
        int lines = 1;
        float x = 0, space = Width(" ", pt, false);
        foreach (string w in s.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
            float ww = Width(w, pt, false);
            if (x > 0 && x + space + ww > width) { lines++; x = 0; }
            x += (x > 0 ? space : 0) + ww;
            while (x > width) { lines++; x -= width; }   // слово длиннее строки - Word режет его
        }
        return lines;
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
        try { AddMeters(); } catch { }
    }

    // Новые типы счётчиков из встроенного списка дописываются в раздел [Счётчики] старого справочника - один раз
    // на каждое обновление списка (метка в файле), чтобы удалённые пользователем строки не возвращались каждый запуск
    public const int MetersVersion = 2;     // то же число - в строке-метке встроенного справочника (Default)
    const string MetersMark = "# список счётчиков обновлён программой, версия ";

    void AddMeters() {
        int section = -1, last = -1, version = 0;
        List<string> have = new List<string>();
        for (int i = 0; i < lines.Count; i++) {
            string s = lines[i].Trim();
            if (s.StartsWith("[") && s.EndsWith("]")) {
                if (section >= 0) break;
                if (Key(s.Substring(1, s.Length - 2)) == Key("Счётчики")) { section = i; last = i; }
                continue;
            }
            if (section < 0) continue;
            if (s.StartsWith(MetersMark)) {
                string d = s.Substring(MetersMark.Length).Trim();
                int n = 0;
                while (n < d.Length && char.IsDigit(d[n])) n++;
                if (n > 0) version = Math.Max(version, int.Parse(d.Substring(0, n)));
            }
            if (s.Length > 0) last = i;
            if (s.Length > 0 && !s.StartsWith("#")) have.Add(Squeeze(s));
        }
        if (section < 0 || version >= MetersVersion) return;
        List<string> add = new List<string>();
        bool take = false;
        foreach (string raw in Default.Replace("\r\n", "\n").Split('\n')) {
            string s = raw.Trim();
            if (s.StartsWith("[") && s.EndsWith("]")) { take = Key(s.Substring(1, s.Length - 2)) == Key("Счётчики"); continue; }
            if (!take || s.Length == 0 || s.StartsWith("#")) continue;
            if (!have.Contains(Squeeze(s))) { add.Add(s); have.Add(Squeeze(s)); }
        }
        add.Insert(0, MetersMark + MetersVersion + (add.Count > 0 ? " (дописано " + add.Count + ")" : ""));
        lines.InsertRange(last + 1, add);
        System.IO.File.WriteAllLines(File, lines.ToArray(), new UTF8Encoding(true));
        Load();
    }

    // для сравнения типов: без пробелов и точек, латиницей, заглавными ("СЕ 318BY R32.043" = "CE318BYR32043")
    public static string Squeeze(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s.ToUpperInvariant()) {
            if (c == ' ' || c == '.' || c == '\t') continue;
            int k = "АВЕКМНОРСТХ".IndexOf(c);
            sb.Append(k >= 0 ? "ABEKMHOPCTX"[k] : c);
        }
        return sb.ToString();
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
"# Расшифровка (по руководствам по эксплуатации Энергомеры): корпус S - в щиток, R - на DIN-рейку, C - сплит на опору\r\n" +
"# (с выносным дисплеем СЕ901); код вроде 146 или 145.1: класс точности, 4 - 230 В, ток (3 - 5(10) А, 4 - 5(40) А,\r\n" +
"# 5 - 5(60) А, 6 - 5(100) А, 9 - 5(80) А), у СЕ 208 после точки - число измерительных элементов;\r\n" +
"# интерфейсы: J - оптопорт, A - RS-485, P - PLC-модем, R - радиомодем, G - GSM; дальше - доп. функции\r\n" +
"# (Q - реле нагрузки, U - параметры сети, Y - два направления, V - эл. пломбы, F, M - датчики поля, L - подсветка).\r\n" +
"# список счётчиков обновлён программой, версия 2\r\n" +
"# Однофазные СЕ 208\r\n" +
"CE208BY S53 145.1 JA.UKVFLZ\r\nCE208BY S53 145.1 JR.UKVFLZ\r\nCE208BY S53 145.1 JP.UKVFLZ\r\nCE208BY S53 145.1 JPR.UKVFLZ\r\n" +
"CE208BY S53 145.1 JR.QUKVFLZ\r\nCE208BY S53 145.1 JPR.QUKVFLZ\r\nCE208BY S53 146.1 JR.UKVFLZ\r\nCE208BY S53 146.1 JPR.UKVFLZ\r\n" +
"CE208BY S53 146.1 JR.QUKVFLZ\r\nCE208BY S53 146.1 JPR.QUKVFLZ\r\nCE208BY S51 145.1 JR.UKVFLZ\r\nCE208BY S51 145.1 JR.QUKVFLZ\r\n" +
"CE208BY S51 145.2 JR.UKVFLZ\r\nCE208BY S51 145.2 JR.QUKVFLZ\r\nCE208BY S7 145.1 JPR.UKVFLZ\r\nCE208BY S7 146.1 JPR.UKVFLZ\r\n" +
"CE208BY S7 149.1 JPR.QUKVFLZ\r\nCE208BY S7 145.1 JR.UKVFLZ\r\nCE208BY S7 145.2 JR.UKVFLZ\r\nCE208BY S7 145.2 JPR.UKVFLZ\r\n" +
"CE208BY S8 146.2 JG.QUKVFLZ\r\nCE208BY R7 145.2 JR.QUKVFLZ\r\nCE208BY R7 745.2 JR.QUYKVFLZ\r\nCE208BY C1.145.2.JR.QUKVFZ\r\n" +
"CE208BY C1.146.2.JR.QUKVFZ\r\nCE208BY C1.145.2.JPR.QUKVFZ\r\nCE208BY C1.146.2.JPR.QUKVFZ\r\nCE208BY C2 145.2.JPR.QUKVFZ\r\n" +
"CE208BY C2 146.2.JPR.QUKVFZ\r\nCE208BY S52/2 145.1 R.LFZ\r\n" +
"# Трёхфазные СЕ 318\r\n" +
"CE318BY R32 043 JA.UVFL\r\nCE318BY R32 043 JR.UVFL\r\nCE318BY R32 146 JA.UVFL\r\nCE318BY R32 146 JR.UVFL\r\n" +
"CE318BY R32 146 JP.UVFL\r\nCE318BY R32 146 JPR.UVFL\r\nCE318BY R32 746 JA.UYVFL\r\nCE318BY R32 643 JR.UYVFL\r\n" +
"CE318BY R32 746 JR.UYVFL\r\nCE318BY S31 043 JPR.UVFL\r\nCE318BY S31 146 JA.UVFL\r\nCE318BY S31 146 JR.UVFL\r\n" +
"CE318BY S31 146 JPR.UVFL\r\nCE318BY S31 146 JR.QUVFL\r\nCE318BY S31 146 JPR.QUVFL\r\nCE318BY S31 043 JA.UVFL\r\n" +
"CE318BY S31 149 JR.QUVFL\r\nCE318BY S31 149 JPR.QUVFL\r\nCE318BY S31 643 JA.UYVFL\r\nCE318BY S31 643 JPR.UYVFL\r\n" +
"CE318BY S31 746 JPR.UYVFL\r\nCE318BY S34 146 JPR.UVFL\r\nCE318BY S34 146 JPR.QUVFL\r\nCE318BY S35 043 JPR.UVFL\r\n" +
"CE318BY S35 146 JPR.UVFL\r\nCE318BY S35 149.JPR.QUVFL\r\nCE318BY S35 746 JPR.UYVFL\r\nCE318BY S35 643 JAPR.SUYVMFL\r\n" +
"CE318BY S37 146 JPR.QUVFL\r\nCE318BY S37 043 JAPRG.UVFL\r\nCE318BY S37 146 JAPRG.QUVFL\r\nCE318BY S37 643 JAPR.UYVFL\r\n" +
"CE318BY S37 643 JAPRG.UYVFL\r\nCE318BY S37 746 JAPR.QUYVFL\r\nCE318BY S37 746 JAPRG.QUYVFL\r\nCE318BY S39 643 JAPRG.UYVFL\r\n" +
"CE318BY S39 746 JAPRG.UYVFL\r\nCE318BY S39 749 JAPRG.QUYVFL\r\nCE318BY C3 146 JR.QUVF\r\nCE318BY C3 146 JPR.QUVF\r\n" +
"CE318BY C3 746 JR.QUYVF\r\n" +
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

    // интерфейсы в модификации - часть перед последней («146 JPR.QUVFL» -> JPR, «145.1 R.LFZ» -> R):
    // P - PLC-модем, R - радиомодем (JPR - оба)
    public static string Modem(string mod) {
        string[] t = mod.Split(new char[] { ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);
        if (t.Length < 2) return "";
        string code = t[t.Length - 2].ToUpperInvariant().Replace('Р', 'P');
        foreach (char c in code) if (!char.IsLetterOrDigit(c)) return "";
        if (code.Length == 0 || char.IsDigit(code[0])) return "";
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
    public static readonly Color Caption = Color.FromArgb(58, 66, 78);
    public static readonly Color Text = Color.FromArgb(31, 36, 43);
    public static readonly Color Hint = Color.FromArgb(128, 137, 150);
    public static readonly Color Missing = Color.FromArgb(255, 222, 222);
    public static readonly Color MissingBorder = Color.FromArgb(220, 90, 90);
    public static readonly Color Bad = Color.FromArgb(200, 55, 55);
    public static readonly Color Good = Color.FromArgb(46, 140, 87);
    public static readonly Color Box = Color.FromArgb(248, 250, 252);
    public static readonly Color Disabled = Color.FromArgb(236, 239, 243);
    // рамка полей ввода - чёрная, чтобы поле не сливалось с фоном; у недоступного - серая
    public static readonly Color Field = Color.Black;
    public static readonly Color FieldOff = Color.FromArgb(170, 176, 184);
    // поле, в котором курсор
    public static readonly Color Focus = Color.FromArgb(255, 247, 205);

    // свой цвет у каждого раздела: фон, полоса заголовка, метка (полоска слева, рамка, значок в меню)
    public static Tint TintOf(string card) {
        switch (card) {
            case "Договор": return new Tint(236, 243, 252, 211, 227, 248, 52, 110, 190);
            case "Фото документов": return new Tint(243, 240, 251, 225, 218, 246, 118, 90, 190);
            case "Потребитель": return new Tint(235, 247, 238, 208, 237, 215, 46, 138, 78);
            case "Объект": return new Tint(253, 245, 229, 248, 229, 188, 196, 126, 24);
            case "Счётчики": return new Tint(232, 246, 247, 201, 234, 237, 22, 126, 138);
            case "Пломбы и оборудование": return new Tint(252, 238, 241, 245, 212, 219, 184, 68, 94);
            case "Дополнительно": return new Tint(244, 246, 230, 228, 235, 196, 116, 138, 26);
            default: return new Tint(241, 242, 244, 223, 226, 231, 100, 110, 125);
        }
    }

    public static readonly Font Base = new Font("Segoe UI", 10f);
    public static readonly Font Small = new Font("Segoe UI", 9f);
    public static readonly Font Bold = new Font("Segoe UI", 10f, FontStyle.Bold);
    public static readonly Font Title = new Font("Segoe UI", 12.5f, FontStyle.Bold);
    public static readonly Font Big = new Font("Segoe UI", 16f, FontStyle.Bold);
    public static readonly Font NavFont = new Font("Segoe UI", 10.5f);
}

public class Tint {
    public readonly Color Body, Head, Mark;
    public Tint(int r1, int g1, int b1, int r2, int g2, int b2, int r3, int g3, int b3) {
        Body = Color.FromArgb(r1, g1, b1);
        Head = Color.FromArgb(r2, g2, b2);
        Mark = Color.FromArgb(r3, g3, b3);
    }
    // тёмный оттенок метки - для заголовков
    public Color Dark { get { return Color.FromArgb(Mark.R * 6 / 10, Mark.G * 6 / 10, Mark.B * 6 / 10); } }
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
        if (c is PhotoStrip) return ((PhotoStrip)c).HeightFor(width);
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

    // своя чёрная рамка поверх плоской серой: список не сливается с фоном (как текстовые поля)
    protected override void WndProc(ref Message m) {
        base.WndProc(ref m);
        if (m.Msg == 0x000F && IsHandleCreated) {
            try {
                using (Graphics g = Graphics.FromHwnd(Handle))
                using (Pen p = new Pen(!Enabled ? St.FieldOff : BackColor == St.Missing ? St.MissingBorder : St.Field))
                    g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            } catch { }
        }
    }

    protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Invalidate(); }
    protected override void OnBackColorChanged(EventArgs e) { base.OnBackColorChanged(e); Invalidate(); }

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
        BackColor = St.Field;
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
        set { missing = value; BackColor = value ? St.MissingBorder : Enabled ? St.Field : St.FieldOff; }
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

    protected override void OnEnabledChanged(EventArgs e) {
        base.OnEnabledChanged(e);
        if (!missing) BackColor = Enabled ? St.Field : St.FieldOff;
    }
}

// Карточка раздела: светлый фон цвета раздела, сверху полоса с названием, слева цветная полоска
public class CardPanel : Panel {
    public Tint Tint = St.TintOf("");
    public int HeadHeight = 38;
    public CardPanel() {
        BackColor = St.Card;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
    }
    public void SetTint(Tint t) { Tint = t; BackColor = t.Body; Invalidate(); }
    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Tint.Head)) g.FillRectangle(b, 0, 0, Width, HeadHeight);
        using (SolidBrush b = new SolidBrush(Tint.Mark)) {
            g.FillRectangle(b, 0, 0, 6, Height);
            g.FillRectangle(b, 0, HeadHeight - 2, Width, 2);
        }
        using (Pen p = new Pen(Tint.Mark)) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
    }
}

// Рамка внутри карточки (счётчик 1, счётчик 2): белая, с рамкой цвета раздела
public class BoxPanel : Panel {
    public string Title = "";
    public Color Mark = St.Border;
    public BoxPanel() {
        BackColor = Color.White;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
    }
    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        Color c = Enabled ? Mark : St.FieldOff;
        using (Pen p = new Pen(c, 2f)) e.Graphics.DrawRectangle(p, 1, 1, Width - 2, Height - 2);
        TextRenderer.DrawText(e.Graphics, Title, St.Bold, new Point(12, 7), Enabled ? St.Text : St.Hint);
    }
}

// Ячейка сетки: подпись над полем; поле может занимать несколько колонок
public class Cell {
    public string Caption;
    public Label Cap;
    public Control Ctl;
    public Grid Sub;
    public int Span = 1;
    public int W;              // ширина поля в знаках (по длине значения); 0 - на всю оставшуюся ширину

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
    public int PadX = 18, PadTop = 0, PadBottom = 12, Gap = 10, RowGap = 6;
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

    static double charW;
    // ширина одного знака основного шрифта (цифры), точки
    public static double CharW {
        get {
            if (charW == 0) charW = TextRenderer.MeasureText("0000000000", St.Base).Width / 10.0;
            return charW;
        }
    }

    // строка, где у полей своя ширина: поля друг за другом слева направо, не влезло - на следующую линию;
    // поля без ширины делят остаток линии
    int Flow(Cell[] row, int width, int y) {
        int capH = St.Small.Height + 2, avail = width - 2 * PadX;
        int[] w = new int[row.Length];
        bool[] fill = new bool[row.Length];
        for (int i = 0; i < row.Length; i++) {
            Cell c = row[i];
            int cap = c.Caption != null && c.Caption.Trim().Length > 0 ? TextRenderer.MeasureText(c.Caption, St.Small).Width + 4 : 0;
            fill[i] = c.W <= 0;
            w[i] = fill[i] ? Math.Max(cap, (int)(15 * CharW)) : Math.Max(cap, (int)(c.W * CharW + 10));
            w[i] = Math.Min(w[i], avail);
        }
        int start = 0;
        while (start < row.Length) {
            // линия: сколько полей влезает
            int end = start, used = 0;
            while (end < row.Length && (end == start || used + Gap + w[end] <= avail)) { used += (end > start ? Gap : 0) + w[end]; end++; }
            int fills = 0;
            for (int i = start; i < end; i++) if (fill[i]) fills++;
            int extra = avail - used;
            bool anyCap = false;
            for (int i = start; i < end; i++) if (row[i].Caption != null) anyCap = true;
            int x = PadX, lineH = 0;
            for (int i = start; i < end; i++) {
                Cell c = row[i];
                int cw = w[i] + (fill[i] && fills > 0 ? extra / fills : 0);
                int cy = y;
                if (c.Cap != null) c.Cap.SetBounds(x, cy, cw, capH);
                if (anyCap && c.Sub == null) cy += capH + 1;
                int h = 0;
                // подпись длиннее поля - поле всё равно своей ширины (короткое значение - короткое поле)
                int fw = fill[i] ? cw : Math.Min(cw, (int)(c.W * CharW + 10));
                if (c.Sub != null) { h = c.Sub.Layout(cw); c.Sub.Host.SetBounds(x, cy, cw, h); }
                else if (c.Ctl != null) { h = Ui.HeightOf(c.Ctl, fw); c.Ctl.SetBounds(x, cy, fw, h); }
                lineH = Math.Max(lineH, cy + h - y);
                x += cw + Gap;
            }
            y += lineH + RowGap;
            start = end;
        }
        return y;
    }

    public int Layout(int width) {
        int capH = St.Small.Height + 2;
        int y = PadTop;
        float colW = (width - 2 * PadX - Gap * (Cols - 1)) / (float)Cols;
        foreach (Cell[] row in rows) {
            bool flow = false;
            foreach (Cell c in row) if (c.W > 0) flow = true;
            if (flow) { y = Flow(row, width, y); continue; }
            int col = 0, rowH = 0;
            bool anyCap = false;
            foreach (Cell c in row) if (c.Caption != null) anyCap = true;
            foreach (Cell c in row) {
                int cx = PadX + (int)Math.Round(col * (colW + Gap));
                int cw = (int)Math.Round(c.Span * colW + (c.Span - 1) * Gap);
                int cy = y;
                // поля одной строки - на одном уровне, даже если у какого-то нет подписи
                if (c.Cap != null) c.Cap.SetBounds(cx, cy, cw, capH);
                if (anyCap && c.Sub == null) cy += capH + 1;
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
        Sub.ForeColor = St.Caption;
        Sub.AutoSize = true;
        Sub.BackColor = Color.Transparent;
        Panel.Controls.Add(Title);
        Panel.Controls.Add(Sub);
        Tint t = St.TintOf(name);
        Panel.SetTint(t);
        Title.ForeColor = t.Dark;
        Grid = new Grid(Panel, cols);
        Grid.PadTop = Panel.HeadHeight + 8;
    }

    public void AddExtra(Control c) {
        Extra.Add(c);
        Panel.Controls.Add(c);
    }

    public int Layout(int width) {
        Panel.Width = width;
        int head = Panel.HeadHeight - 2;
        Title.Location = new Point(18, (head - Title.Height) / 2);
        Sub.Location = new Point(Title.Right + 10, Title.Top + Title.Height - Sub.Height - 2);
        int x = width - 10;
        for (int i = Extra.Count - 1; i >= 0; i--) {
            Control c = Extra[i];
            x -= c.Width;
            c.Location = new Point(x, Math.Max(2, (head - c.Height) / 2));
            x -= 8;
        }
        // узко - пояснение рядом с заголовком прячем, чтобы кнопки на него не налезали
        Sub.Visible = Extra.Count == 0 || x >= Sub.Right;
        int h = Grid.Layout(width);
        Panel.Height = h;
        return h;
    }
}

// Пункт меню слева: название раздела и сколько в нём не заполнено
public class NavItem : Control {
    public bool Active, Small;
    public int Count = -1;
    public Color Mark = Color.Empty;    // цвет раздела - значок слева
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
        int tx = 22;
        if (Mark != Color.Empty) {
            using (SolidBrush b = new SolidBrush(Mark)) g.FillRectangle(b, 14, (Height - 12) / 2, 12, 12);
            using (Pen p = new Pen(Color.FromArgb(150, 255, 255, 255))) g.DrawRectangle(p, 14, (Height - 12) / 2, 11, 11);
            tx = 34;
        }
        Rectangle tr = new Rectangle(tx, 0, Width - tx - 38, Height);
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
        box.Enter += delegate { box.BackColor = St.Focus; };
        box.Leave += delegate { box.BackColor = Enabled ? Color.White : St.Disabled; };
        pick.FlatStyle = FlatStyle.Flat;
        pick.FlatAppearance.BorderColor = St.Field;
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
        r.Power.FlatAppearance.BorderColor = St.Field;
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

// Машиночитаемая зона паспорта (две строки по 44 знака с «<<<» внизу страницы с фото): своё распознавание,
// без сторонних программ. Знаки - шрифт OCR-B, образцы зашиты в программу (MrzData). Проверка - контрольные цифры.
public class MrzResult {
    public string Line1 = "", Line2 = "";
    public string Personal = "", Number = "", Surname = "", Names = "";
    public bool Valid;          // сошлись все контрольные цифры
    public bool PersonalOk;     // сошлась контрольная цифра личного номера
    public int Checks;          // сколько контрольных цифр сошлось (из 5)
    public string Info = "";    // как искали - без самих данных
}

// Серое изображение (0 - чёрный, 255 - белый)
public class GrayImage {
    public int W, H;
    public byte[] P;

    public GrayImage(int w, int h) { W = w; H = h; P = new byte[w * h]; }

    public static GrayImage FromBitmap(Bitmap b) {
        GrayImage g = new GrayImage(b.Width, b.Height);
        BitmapData d = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        try {
            byte[] row = new byte[Math.Abs(d.Stride)];
            for (int y = 0; y < b.Height; y++) {
                Marshal.Copy(new IntPtr(d.Scan0.ToInt64() + (long)y * d.Stride), row, 0, row.Length);
                int o = y * g.W;
                for (int x = 0; x < b.Width; x++) {
                    int i = x * 3;
                    g.P[o + x] = (byte)((row[i] * 114 + row[i + 1] * 587 + row[i + 2] * 299) / 1000);
                }
            }
        } finally { b.UnlockBits(d); }
        return g;
    }

    // поворот на 90° по часовой
    public GrayImage Rotate90() {
        GrayImage r = new GrayImage(H, W);
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++) r.P[x * r.W + (H - 1 - y)] = P[y * W + x];
        return r;
    }

    public float Sample(double x, double y) {
        if (x < 0) x = 0; if (y < 0) y = 0;
        if (x > W - 1.001) x = W - 1.001; if (y > H - 1.001) y = H - 1.001;
        int x0 = (int)x, y0 = (int)y;
        double fx = x - x0, fy = y - y0;
        int i = y0 * W + x0;
        return (float)((P[i] * (1 - fx) + P[i + 1] * fx) * (1 - fy) + (P[i + W] * (1 - fx) + P[i + W + 1] * fx) * fy);
    }
}

public static class Mrz {
    public const string Chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ<";
    public const int GW = 16, GH = 20;
    public const int TV = 9;        // видов образца у каждого знака: толщина штриха x размытие

    public class Comp { public int X0, Y0, X1, Y1, N; public double Cx { get { return (X0 + X1) / 2.0; } } public double Cy { get { return (Y0 + Y1) / 2.0; } } public int Wd { get { return X1 - X0 + 1; } } public int Ht { get { return Y1 - Y0 + 1; } } }

    // ---------- загрузка: поворот по EXIF, уменьшение ----------

    // как повёрнут снимок (метка EXIF от телефона): 1 - как есть, 3 - вверх ногами, 6 и 8 - на боку
    public static int Orientation(Image img) {
        try {
            foreach (int id in img.PropertyIdList)
                if (id == 0x0112) {
                    byte[] v = img.GetPropertyItem(id).Value;
                    if (v.Length < 2) return 1;
                    int o = v[0] | v[1] << 8;
                    if (o < 1 || o > 8) o = v[1] | v[0] << 8;      // байты в обратном порядке (EXIF «MM»)
                    return o >= 1 && o <= 8 ? o : 1;
                }
        } catch { }
        return 1;
    }

    // работа с картинками (GDI+) - по очереди: так надёжнее и старому компьютеру легче
    public static readonly object Gdi = new object();

    public static Bitmap Load(string path, int maxSide) {
        byte[] bytes = File.ReadAllBytes(path);
        lock (Gdi)
        using (MemoryStream ms = new MemoryStream(bytes))
        using (Image img = Image.FromStream(ms)) {
            double k = Math.Min(1.0, (double)maxSide / Math.Max(img.Width, img.Height));
            return Draw(img, Math.Max(1, (int)(img.Width * k)), Math.Max(1, (int)(img.Height * k)), Orientation(img));
        }
    }

    // снимок -> новая картинка w x h (размер до поворота), повёрнутая по метке EXIF: 3 - на 180°, 6 - на 90° по часовой, 8 - против
    public static Bitmap Draw(Image src, int w, int h, int orient) {
        bool side = orient == 6 || orient == 8;
        Bitmap b = new Bitmap(side ? h : w, side ? w : h, PixelFormat.Format24bppRgb);
        using (Graphics g = Graphics.FromImage(b)) {
            g.InterpolationMode = InterpolationMode.HighQualityBilinear;
            if (orient == 6) { g.TranslateTransform(b.Width, 0); g.RotateTransform(90); }
            else if (orient == 8) { g.TranslateTransform(0, b.Height); g.RotateTransform(270); }
            else if (orient == 3) { g.TranslateTransform(b.Width, b.Height); g.RotateTransform(180); }
            g.DrawImage(src, 0, 0, w, h);
        }
        return b;
    }

    // снимок целиком (для чтения знаков) и уменьшенный (для поиска строк)
    public static void LoadPair(string path, out GrayImage full, out GrayImage small) {
        lock (Gdi)
        using (Bitmap b = Load(path, 4200)) {
            full = GrayImage.FromBitmap(b);
            double k = Math.Min(1.0, 2400.0 / Math.Max(b.Width, b.Height));
            using (Bitmap sm = new Bitmap(Math.Max(1, (int)(b.Width * k)), Math.Max(1, (int)(b.Height * k)), PixelFormat.Format24bppRgb)) {
                using (Graphics gr = Graphics.FromImage(sm)) { gr.InterpolationMode = InterpolationMode.HighQualityBilinear; gr.DrawImage(b, 0, 0, sm.Width, sm.Height); }
                small = GrayImage.FromBitmap(sm);
            }
        }
    }

    public static MrzResult Read(string path) {
        // строки ищем на уменьшенном снимке, а читаем по полному - мелкие знаки так чётче
        GrayImage g, full;
        LoadPair(path, out full, out g);
        MrzResult best = null;
        StringBuilder info = new StringBuilder();
        for (int rot = 0; rot < 4; rot++) {
            if (rot > 0) { g = g.Rotate90(); full = full.Rotate90(); }
            MrzResult r = Find(g, full, info);
            if (r != null && (best == null || r.Checks > best.Checks)) best = r;
            if (best != null && best.Valid) break;
        }
        if (best == null) { best = new MrzResult(); best.Info = "строки с <<< не найдены. " + info; }
        else best.Info = info.ToString() + best.Info;
        return best;
    }

    // ---------- поиск двух строк ----------

    static MrzResult Find(GrayImage g, GrayImage full, StringBuilder info) {
        bool[] dark = Binarize(g);
        List<Comp> comps = Components(dark, g.W, g.H);
        List<Comp> cand = new List<Comp>();
        foreach (Comp c in comps)
            if (c.Ht >= 7 && c.Ht <= g.H / 10 && c.Wd <= 3 * c.Ht && c.N >= 8) cand.Add(c);
        cand.Sort(delegate(Comp a, Comp b) { return a.X0.CompareTo(b.X0); });
        List<List<Comp>> lines = Chains(cand);
        // пара строк: друг под другом, похожие, по ~44 знака
        List<Comp> top = null, bottom = null;
        double bestScore = double.MinValue;
        for (int i = 0; i < lines.Count; i++)
            for (int j = 0; j < lines.Count; j++) {
                if (i == j) continue;
                List<Comp> a = lines[i], b = lines[j];
                double ha = CapHeight(a), hb = CapHeight(b);
                if (ha / hb > 1.3 || hb / ha > 1.3) continue;
                double hc = (ha + hb) / 2, dy = MeanCy(b) - MeanCy(a);
                if (dy < 1.1 * hc || dy > 3.4 * hc) continue;
                int ax0 = a[0].X0, ax1 = a[a.Count - 1].X1, bx0 = b[0].X0, bx1 = b[b.Count - 1].X1;
                int ov = Math.Min(ax1, bx1) - Math.Max(ax0, bx0);
                if (ov < 0.7 * Math.Min(ax1 - ax0, bx1 - bx0)) continue;
                double score = -Math.Abs(a.Count - 44) - Math.Abs(b.Count - 44) - Math.Abs((ax1 - ax0) - (bx1 - bx0)) / hc;
                if (score > bestScore) { bestScore = score; top = a; bottom = b; }
            }
        if (top == null) return null;
        info.Append("строки найдены (" + top.Count + " и " + bottom.Count + " знаков, высота знака " + CapHeight(top).ToString("0") + " точек); ");
        double f = (double)full.W / g.W;
        List<double[]> s1 = ReadLine(full, Scale(top, f)), s2 = ReadLine(full, Scale(bottom, f));
        if (s1 == null || s2 == null) return null;
        return Decode(s1, s2);
    }

    public static bool[] Binarize(GrayImage g) {
        int w = g.W, h = g.H, r = Math.Max(10, Math.Min(w, h) / 50);
        int[] sum = new int[(w + 1) * (h + 1)];
        for (int y = 0; y < h; y++) {
            int row = 0;
            for (int x = 0; x < w; x++) {
                row += g.P[y * w + x];
                sum[(y + 1) * (w + 1) + x + 1] = sum[y * (w + 1) + x + 1] + row;
            }
        }
        bool[] dark = new bool[w * h];
        for (int y = 0; y < h; y++) {
            int ya = Math.Max(0, y - r), yb = Math.Min(h, y + r + 1);
            for (int x = 0; x < w; x++) {
                int xa = Math.Max(0, x - r), xb = Math.Min(w, x + r + 1);
                int s = sum[yb * (w + 1) + xb] - sum[ya * (w + 1) + xb] - sum[yb * (w + 1) + xa] + sum[ya * (w + 1) + xa];
                int mean = s / ((xb - xa) * (yb - ya));
                int v = g.P[y * w + x];
                dark[y * w + x] = mean - v > Math.Max(14, mean * 18 / 100);
            }
        }
        return dark;
    }

    static int Root(int[] p, int a) { while (p[a] != a) { p[a] = p[p[a]]; a = p[a]; } return a; }

    // связные области (8 соседей)
    public static List<Comp> Components(bool[] dark, int w, int h) {
        int[] lab = new int[w * h];
        List<int> parent = new List<int>();
        parent.Add(0);
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++) {
                int i = y * w + x;
                if (!dark[i]) continue;
                int best = 0;
                int[] nb = { x > 0 ? lab[i - 1] : 0, y > 0 ? lab[i - w] : 0, y > 0 && x > 0 ? lab[i - w - 1] : 0, y > 0 && x < w - 1 ? lab[i - w + 1] : 0 };
                foreach (int n in nb) if (n > 0 && (best == 0 || n < best)) best = n;
                if (best == 0) { best = parent.Count; parent.Add(best); }
                lab[i] = best;
                foreach (int n in nb) if (n > 0 && n != best) {
                    // объединение двух областей
                    int ra = n, rb = best;
                    while (parent[ra] != ra) ra = parent[ra];
                    while (parent[rb] != rb) rb = parent[rb];
                    if (ra != rb) parent[Math.Max(ra, rb)] = Math.Min(ra, rb);
                }
            }
        int[] par = parent.ToArray();
        Dictionary<int, Comp> map = new Dictionary<int, Comp>();
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++) {
                int l = lab[y * w + x];
                if (l == 0) continue;
                int r = Root(par, l);
                Comp c;
                if (!map.TryGetValue(r, out c)) { c = new Comp(); c.X0 = c.X1 = x; c.Y0 = c.Y1 = y; map[r] = c; }
                if (x < c.X0) c.X0 = x; if (x > c.X1) c.X1 = x;
                if (y < c.Y0) c.Y0 = y; if (y > c.Y1) c.Y1 = y;
                c.N++;
            }
        return new List<Comp>(map.Values);
    }

    // цепочки знаков слева направо: соседний справа, той же высоты строки, недалеко
    static List<List<Comp>> Chains(List<Comp> cand) {
        List<List<Comp>> lines = new List<List<Comp>>();
        bool[] used = new bool[cand.Count];
        for (int s = 0; s < cand.Count; s++) {
            if (used[s]) continue;
            List<Comp> chain = new List<Comp>();
            chain.Add(cand[s]);
            int cur = s;
            double href = cand[s].Ht;
            while (true) {
                Comp c = cand[cur];
                int next = -1;
                for (int j = cur + 1; j < cand.Count; j++) {
                    Comp d = cand[j];
                    if (d.X0 > c.X1 + 2.0 * href) break;
                    if (used[j] || d.X0 < c.X0 + 0.4 * c.Wd) continue;
                    if (Math.Abs(d.Cy - c.Cy) > 0.45 * href) continue;
                    if (d.Ht < 0.3 * href || d.Ht > 1.45 * href) continue;
                    if (next < 0 || d.X0 < cand[next].X0) next = j;
                }
                if (next < 0) break;
                chain.Add(cand[next]);
                cur = next;
                href = Math.Max(href, Math.Min(cand[next].Ht, href * 1.45));
            }
            if (chain.Count >= 28) {
                foreach (Comp c in chain) used[cand.IndexOf(c)] = true;
                lines.Add(chain);
            }
        }
        return lines;
    }

    public static List<Comp> Scale(List<Comp> line, double f) {
        List<Comp> r = new List<Comp>();
        foreach (Comp c in line) {
            Comp d = new Comp();
            d.X0 = (int)(c.X0 * f); d.Y0 = (int)(c.Y0 * f); d.X1 = (int)((c.X1 + 1) * f) - 1; d.Y1 = (int)((c.Y1 + 1) * f) - 1; d.N = c.N;
            r.Add(d);
        }
        return r;
    }

    static double CapHeight(List<Comp> line) {
        List<int> hs = new List<int>();
        foreach (Comp c in line) hs.Add(c.Ht);
        hs.Sort();
        return hs[(int)(hs.Count * 0.8)];
    }

    static double MeanCy(List<Comp> line) {
        double s = 0;
        foreach (Comp c in line) s += c.Cy;
        return s / line.Count;
    }

    // ---------- строка: выпрямить, разрезать на 44 знака, сравнить с образцами ----------

    // для каждого из 44 знаков - оценки сходства со всеми знаками OCR-B
    static List<double[]> ReadLine(GrayImage g, List<Comp> line) {
        double cap = CapHeight(line);
        // наклон строки - по центрам высоких знаков
        double sx = 0, sy = 0, sxx = 0, sxy = 0; int n = 0;
        foreach (Comp c in line) if (c.Ht >= 0.75 * cap) { sx += c.Cx; sy += c.Cy; sxx += c.Cx * c.Cx; sxy += c.Cx * c.Cy; n++; }
        if (n < 10) return null;
        double b = (n * sxy - sx * sy) / (n * sxx - sx * sx), a = (sy - b * sx) / n;
        double norm = Math.Sqrt(1 + b * b), dx = 1 / norm, dy = b / norm;
        double x0 = line[0].X0 - 0.8 * cap, x1 = line[line.Count - 1].X1 + 0.8 * cap;
        int L = (int)((x1 - x0) * norm), SH = (int)(2.6 * cap);
        double px = x0, py = a + b * x0;
        float[] strip = new float[L * SH];
        for (int v = 0; v < SH; v++)
            for (int u = 0; u < L; u++) {
                double off = v - SH / 2.0;
                strip[v * L + u] = g.Sample(px + u * dx - off * dy, py + u * dy + off * dx);
            }
        // уровни фона и краски в полосе
        float[] sorted = (float[])strip.Clone();
        Array.Sort(sorted);
        float bg = sorted[(int)(sorted.Length * 0.85)], fg = sorted[(int)(sorted.Length * 0.02)];
        if (bg - fg < 20) return null;
        float[] ink = new float[strip.Length];
        bool[] bin = new bool[strip.Length];
        for (int i = 0; i < strip.Length; i++) {
            float k = (bg - strip[i]) / (bg - fg);
            ink[i] = k < 0 ? 0 : k > 1 ? 1 : k;
            bin[i] = ink[i] > 0.45f;
        }
        List<Comp> cs = Components(bin, L, SH);
        List<Comp> glyphs = new List<Comp>();
        foreach (Comp c in cs)
            if (c.Ht >= 0.25 * cap && c.Ht <= 1.6 * cap && Math.Abs(c.Cy - SH / 2.0) < 0.7 * cap && c.N >= 6) glyphs.Add(c);
        if (glyphs.Count < 20) return null;
        glyphs.Sort(delegate(Comp p, Comp q) { return p.Cx.CompareTo(q.Cx); });
        List<int> hs = new List<int>();
        foreach (Comp c in glyphs) hs.Add(c.Ht);
        hs.Sort();
        double capS = hs[(int)(hs.Count * 0.8)];
        double mid = 0; int m = 0;
        foreach (Comp c in glyphs) if (c.Ht >= 0.75 * capS) { mid += c.Cy; m++; }
        mid = m > 0 ? mid / m : SH / 2.0;
        double c0 = glyphs[0].Cx, c43 = glyphs[glyphs.Count - 1].Cx, pitch = (c43 - c0) / 43;
        if (pitch < 0.6 * capS || pitch > 1.7 * capS) return null;
        List<double[]> res = new List<double[]>();
        for (int i = 0; i < 44; i++) {
            double ci = c0 + i * pitch, lo = ci - pitch / 2, hi = ci + pitch / 2;
            int bx0 = int.MaxValue, by0 = int.MaxValue, bx1 = -1, by1 = -1;
            foreach (Comp c in glyphs) {
                if (c.X1 < lo || c.X0 > hi) continue;
                if (c.Cx < lo || c.Cx >= hi) { if (c.Wd < 1.3 * pitch) continue; }
                bx0 = Math.Min(bx0, Math.Max(c.X0, (int)Math.Ceiling(lo)));
                bx1 = Math.Max(bx1, Math.Min(c.X1, (int)Math.Floor(hi)));
                by0 = Math.Min(by0, c.Y0);
                by1 = Math.Max(by1, c.Y1);
            }
            double[] sc = new double[Chars.Length];
            if (bx1 < 0) { for (int k = 0; k < sc.Length; k++) sc[k] = -2; res.Add(sc); continue; }
            float[] grid = Grid(ink, L, bx0, by0, bx1, by1);
            double relH = (by1 - by0 + 1) / capS, relY = ((by0 + by1) / 2.0 - mid) / capS, asp = (bx1 - bx0 + 1.0) / (by1 - by0 + 1);
            for (int k = 0; k < Chars.Length; k++) sc[k] = Score(grid, relH, relY, asp, k);
            res.Add(sc);
        }
        return res;
    }

    // знак -> сетка GW x GH: высота знака - во всю сетку, ширина - по пропорции, по центру
    public static float[] Grid(float[] ink, int stride, int x0, int y0, int x1, int y1) { return Grid(ink, stride, x0, y0, x1, y1, GW, GH); }

    public static float[] Grid(float[] ink, int stride, int x0, int y0, int x1, int y1, int GW, int GH) {
        float[] grid = new float[GW * GH];
        double h = y1 - y0 + 1, w = x1 - x0 + 1;
        double k = Math.Min(GH / h, GW / w);       // точек сетки на точку знака
        double gw = w * k, gh = h * k, ox = (GW - gw) / 2, oy = (GH - gh) / 2;
        for (int gy = 0; gy < GH; gy++)
            for (int gx = 0; gx < GW; gx++) {
                // клетка сетки -> прямоугольник в знаке; среднее по нему (по 3x3 точкам)
                double s = 0; int cnt = 0;
                for (int sy = 0; sy < 3; sy++)
                    for (int sxi = 0; sxi < 3; sxi++) {
                        double fx = (gx + (sxi + 0.5) / 3 - ox) / k, fy = (gy + (sy + 0.5) / 3 - oy) / k;
                        cnt++;
                        if (fx < 0 || fy < 0 || fx >= w || fy >= h) continue;
                        s += ink[(y0 + (int)fy) * stride + x0 + (int)fx];
                    }
                grid[gy * GW + gx] = (float)(s / cnt);
            }
        return grid;
    }

    static float[][] templ;          // [знак * TV + вид]
    static double[] tRelH, tRelY, tAsp;

    static void LoadTemplates() {
        if (templ != null) return;
        int n = Chars.Length, size = GW * GH;
        float[][] t = new float[n * TV][];
        tRelH = new double[n]; tRelY = new double[n]; tAsp = new double[n];
        for (int k = 0; k < n * TV; k++) {
            t[k] = new float[size];
            for (int i = 0; i < size; i++) t[k][i] = Convert.ToInt32(MrzData.Grids[k * size + i].ToString(), 16) / 15f;
        }
        for (int k = 0; k < n; k++) {
            string[] p = MrzData.Shapes[k].Split(';');
            tRelH[k] = double.Parse(p[0], System.Globalization.CultureInfo.InvariantCulture);
            tRelY[k] = double.Parse(p[1], System.Globalization.CultureInfo.InvariantCulture);
            tAsp[k] = double.Parse(p[2], System.Globalization.CultureInfo.InvariantCulture);
        }
        templ = t;
    }

    // сходство знака с образцом: лучшее из видов образца (корреляция) минус штраф за размер и положение
    static double Score(float[] grid, double relH, double relY, double asp, int k) {
        LoadTemplates();
        double ma = 0;
        for (int i = 0; i < grid.Length; i++) ma += grid[i];
        ma /= grid.Length;
        double saa = 0;
        for (int i = 0; i < grid.Length; i++) saa += (grid[i] - ma) * (grid[i] - ma);
        double best = 0;
        for (int v = 0; v < TV; v++) {
            float[] t = templ[k * TV + v];
            double mb = 0;
            for (int i = 0; i < grid.Length; i++) mb += t[i];
            mb /= grid.Length;
            double sab = 0, sbb = 0;
            for (int i = 0; i < grid.Length; i++) { double db = t[i] - mb; sab += (grid[i] - ma) * db; sbb += db * db; }
            double corr = saa > 0 && sbb > 0 ? sab / Math.Sqrt(saa * sbb) : 0;
            if (v == 0 || corr > best) best = corr;
        }
        return best - 1.0 * Math.Abs(relH - tRelH[k]) - 0.6 * Math.Abs(relY - tRelY[k]) - 0.35 * Math.Abs(Math.Log(asp / tAsp[k]));
    }

    // ---------- строки -> данные: что где может стоять и контрольные цифры ----------

    const string Digits = "0123456789", Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    public static int Val(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'A' && c <= 'Z') return c - 'A' + 10;
        return 0;
    }

    public static char Check(string s) {
        int[] w = { 7, 3, 1 };
        int sum = 0;
        for (int i = 0; i < s.Length; i++) sum += Val(s[i]) * w[i % 3];
        return (char)('0' + sum % 10);
    }

    // допустимые знаки в позиции второй строки (паспорт РБ: номер ББЦЦЦЦЦЦЦ, личный номер ЦЦЦЦЦЦЦБЦЦЦББЦ)
    static string Allowed2(int i) {
        if (i <= 1) return Letters;
        if (i <= 9) return Digits;
        if (i <= 12) return Letters;
        if (i <= 19) return Digits;
        if (i == 20) return "MF<";
        if (i <= 27) return Digits;
        if (i <= 34) return Digits;
        if (i == 35) return "ABCEHKM";
        if (i <= 38) return Digits;
        if (i <= 40) return Letters;
        return Digits;
    }

    static string Allowed1(int i) {
        if (i == 0) return "P";
        if (i == 1) return Letters + "<";
        if (i <= 4) return Letters;
        return Letters + "<";
    }

    class Pick { public char C; public double S; public char C2; public double S2; }

    static Pick Best(double[] sc, string allowed) {
        Pick p = new Pick();
        p.S = p.S2 = double.MinValue;
        foreach (char c in allowed) {
            double s = sc[Chars.IndexOf(c)];
            if (s > p.S) { p.C2 = p.C; p.S2 = p.S; p.C = c; p.S = s; }
            else if (s > p.S2) { p.C2 = c; p.S2 = s; }
        }
        return p;
    }

    static MrzResult Decode(List<double[]> s1, List<double[]> s2) {
        StringBuilder a = new StringBuilder(), b = new StringBuilder();
        for (int i = 0; i < 44; i++) a.Append(Best(s1[i], Allowed1(i)).C);
        Pick[] p2 = new Pick[44];
        for (int i = 0; i < 44; i++) { p2[i] = Best(s2[i], Allowed2(i)); b.Append(p2[i].C); }
        char[] l2 = b.ToString().ToCharArray();
        bool blr = new string(l2, 10, 3) == "BLR";
        if (blr) {
            // в личном номере РБ - дата рождения (ДДММГГ), та же, что в поле даты (ГГММДД): читаем обе вместе
            int[][] pairs = { new int[] { 13, 33 }, new int[] { 14, 34 }, new int[] { 15, 31 }, new int[] { 16, 32 }, new int[] { 17, 29 }, new int[] { 18, 30 } };
            foreach (int[] pr in pairs) {
                char bestC = l2[pr[0]];
                double bestS = double.MinValue;
                foreach (char c in Digits) {
                    double sc = s2[pr[0]][Chars.IndexOf(c)] + s2[pr[1]][Chars.IndexOf(c)];
                    if (sc > bestS) { bestS = sc; bestC = c; }
                }
                l2[pr[0]] = l2[pr[1]] = bestC;
            }
            // первая цифра личного номера - пол и век: 1, 3, 5 - мужской, 2, 4, 6 - женский
            char sx = l2[20], d0 = l2[28];
            if (sx == 'M' && "135".IndexOf(d0) < 0 || sx == 'F' && "246".IndexOf(d0) < 0) {
                string allowedSex = sx == 'M' ? "135" : "246";
                double bs = double.MinValue;
                foreach (char c in allowedSex) if (s2[28][Chars.IndexOf(c)] > bs) { bs = s2[28][Chars.IndexOf(c)]; l2[28] = c; }
            }
        }
        // поля с контрольной цифрой: [начало, конец) и место контрольной цифры
        int[][] fields = { new int[] { 0, 9, 9 }, new int[] { 13, 19, 19 }, new int[] { 21, 27, 27 }, new int[] { 28, 42, 42 } };
        foreach (int[] f in fields) Repair(l2, p2, f[0], f[1], f[2]);
        string line2 = new string(l2);
        MrzResult r = new MrzResult();
        r.Line1 = a.ToString();
        r.Line2 = line2;
        int ok = 0;
        foreach (int[] f in fields) if (Check(line2.Substring(f[0], f[1] - f[0])) == line2[f[2]]) ok++;
        string comp = line2.Substring(0, 10) + line2.Substring(13, 7) + line2.Substring(21, 22);
        if (Check(comp) == line2[43]) ok++;
        r.Checks = ok;
        r.Valid = ok == 5;
        r.PersonalOk = Check(line2.Substring(28, 14)) == line2[42];
        if (blr) {
            string dob = line2.Substring(13, 6), pn = line2.Substring(28, 14);
            bool same = pn.Substring(1, 2) == dob.Substring(4, 2) && pn.Substring(3, 2) == dob.Substring(2, 2) && pn.Substring(5, 2) == dob.Substring(0, 2);
            bool sex = line2[20] == 'M' ? "135".IndexOf(pn[0]) >= 0 : "246".IndexOf(pn[0]) >= 0;
            if (!same || !sex) { r.PersonalOk = false; r.Valid = false; }
        }
        r.Number = line2.Substring(0, 9).Replace("<", "");
        r.Personal = line2.Substring(28, 14);
        string names = r.Line1.Substring(5).TrimEnd('<');
        int sep = names.IndexOf("<<");
        r.Surname = (sep < 0 ? names : names.Substring(0, sep)).Replace('<', ' ').Trim();
        r.Names = sep < 0 ? "" : names.Substring(sep + 2).Replace('<', ' ').Trim();
        r.Info = "контрольных цифр сошлось " + ok + " из 5";
        return r;
    }

    // контрольная цифра не сошлась - заменяем ОДИН самый сомнительный знак на второй вариант,
    // и только если он и правда почти равноценен (иначе подгонка под контрольную цифру даст чужой номер)
    static void Repair(char[] l, Pick[] p, int from, int to, int chk) {
        if (Check(new string(l, from, to - from)) == l[chk]) return;
        int found = -1, count = 0;
        for (int i = from; i <= to; i++) {
            int pos = i == to ? chk : i;
            if (p[pos].C2 == '\0' || p[pos].S - p[pos].S2 > 0.08) continue;
            char[] t = (char[])l.Clone();
            t[pos] = p[pos].C2;
            if (Check(new string(t, from, to - from)) == t[chk]) { found = pos; count++; }
        }
        if (count == 1) l[found] = p[found].C2;
    }
}

// Образцы знаков OCR-B (шрифт OCR-B, общественное достояние) для распознавания строк паспорта
public static class MrzData {
    // знак за знаком (0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ<), у каждого 9 вида (толщина x размытие), по 320 клеток, яркость краски 0..F
    public const string Grids =
        "000037DFEC7200000005FFFFFFFE3000001EFFC98BFFD000009FF600007FF70000EFB000000CFD0000FF9000000AFF0002FF50000007FF0004FF4000" +
        "0005FF2005FF30000004FF3005FF30000004FF4005FF20000004FF4005FF30000004FF4005FF30000005FF3003FF50000007FF1000FF9000000AFF00" +
        "00FFB000000CFD00009FF600008FF700001EFFB88BFFD0000004EFFFFFFF4000000027DFFE930000001257899975310001259BBBBBB963100159BB98" +
        "89BB9520026AB854458AB730038B95211259B940049B84100137BA4004AB73100026AA5005AB72000026AA5005AB62000025AB5005AB62000025AB60" +
        "05AB62000025AB6005AB62000025AB6005AB62000026AB5004AB73100026BA50049B84100137BA40038BA6211259B940027AB854458BB7300148BB98" +
        "89BB952000259BBBBBBA6310001257899986310000046777777640000004777777774000000576666667400000046655556640000004654344564000" +
        "000454323346400000045322234540000004532122454000000443211235400000044321123540000004432112354000000453212235400000045322" +
        "234540000004543223454000000454333356400000046544455640000004665556675000000477666677400000047777777740000004677777764000" +
        "00004AEFFD9400000009FFFFFFFF7000006FFFFFFFFFF40000BFFFC99CFFFA0001FFFB1001DFFE0005FFF500005FFF3008FFF300004FFF6009FFF000" +
        "001FFF700AFFD000000FFF800AFFD000000EFF900AFFC000000EFF900AFFD000000EFF900AFFE000001FFF8009FFF300004FFF7006FFF500006FFF40" +
        "02FFFC1002DFFE1000BFFFC99CFFFB00005FFFFFFFFFF4000008FFFFFFFF800000004AEFFFB50000001368AAAA9641000137ADDDDDDB7410016ADEDD" +
        "CDEDB620028CEDA88ACED93003ADDA6436ADEA4004BEC831137CEB5005BEB620015BEC5005CEB510015ADC5005CEA510014ADC6005CDA5100149DD60" +
        "05CDA5100149DD6005CEB510014ADD6005CEB620015ADC6005BEC720026BEC5004BED942148CEB50039DEB8557BEEA40027CEDCAABDEC8300159CEED" +
        "DEEDA61000259BCDDDCA631000124688987531000134677887754310024689999987642003689AAAAAA976300479AA9999A987300489A98778999740" +
        "048998666789984005899754457998400589864334689850058986433468985005898642246898500589864224689850058986433468985005898643" +
        "34689850058997544579985004899866678998400489A987789A98400479AA9999AA973003689AAAAAA9863002467899999864200134677887764310" +
        "000037CEEB6200000002FFFFFFFE1000001DF953349FC000005FA000000BF500009F10000003F70000EF00000000FC0001FF00000000FE0003FD0000" +
        "0000FF2005FD00000000EF3005FD00000000EF4005FC00000000EF4005FC00000000EF4005FD00000000EF3002FE00000000FF1000EF00000000FD00" +
        "00AF20000003F800005FB000000CF500001CF94334AFC0000001EFFFFFFF1000000026CEFD7300000012568887642100002589999998520001479866" +
        "668974100169853223588510027963100136971002895200002598200288410000248820029841000014892002984100001489200398410000148920" +
        "039841000014892003984100001489200298410000148920029841000025982002895200002598200279631001369720016985322368961001479866" +
        "678974100025899999985200001256888875210000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000004DE7000000000005FFFE00000000009FFFFE0000000019FFECFE00000000AFFC49FE00000000FFD209FE00000000260009FE000000000000" +
        "09FE00000000000009FE00000000000009FE00000000000009FE00000000000009FE00000000000009FE00000000000009FE00000000000009FE0000" +
        "0000000009FE00000000000009FE00000000000009FE00000000000009FE00000000000003E70000000001247996000000001258BDC900000000369B" +
        "DDD90000000059BCCCC9000000008BA99AC900000000798669B900000000454358B900000000221148B900000000000148B900000000000148B90000" +
        "0000000148B900000000000148B900000000000148B900000000000148B900000000000148B900000000000148B900000000000148B9000000000001" +
        "48B900000000000148A90000000000013687000015566677777766522566777787777662266777888887776226778888888887722777888999988772" +
        "277888999998887227888999999988722788899999998872278889999999887227788899999888722778888999988872267788888888887226777888" +
        "888887722666777888888772256667777887777225566677777777721455566777777772144556667777776214445566677777621344455667777762" +
        "000000019EFB10000000002CFFFF4000000003DFFFFF400000004EFFFFFF40000003FFFFFFFF40000005FFFEAFFF40000003FFF34FFF400000001210" +
        "4FFF4000000000004FFF4000000000004FFF4000000000004FFF4000000000004FFF4000000000004FFF4000000000004FFF4000000000004FFF4000" +
        "000000004FFF4000000000004FFF4000000000004FFF4000000000004FFF4000000000001BEB1000000001257997200000001359CDDA2000000036AD" +
        "EFEB300000016ADEFFEB300000029DEEEEEB30000002ADDBCDEB300000028A989CEB3000000145447BDB3000000011126BDB3000000000026BDB3000" +
        "000000026BDB3000000000026BDB3000000000026BDB3000000000026BDB3000000000026BDB3000000000026BDB3000000000026BDB300000000002" +
        "6BDB3000000000025ACA20000000000137872000000135677876200000015789AA9720000002689BBBA82000000279BCCBA8200000028ABCCCB92000" +
        "000289ABBBA920000002789ABBA82000000256799AA82000000145689998200000012456899820000000234679982000000012367998200000001235" +
        "789820000000123578982000000012357898200000001235789820000000123578982000000012357888200000001235688720000000113467762000" +
        "000000002AE6000000000003DFFE00000000005DFFFE0000000004FF95FE000000009FF504FE00000000AD4004FE00000000010004FE000000000000" +
        "04FE00000000000004FE00000000000004FE00000000000004FE00000000000004FE00000000000004FE00000000000004FE00000000000004FE0000" +
        "0000000004FE00000000000004FE00000000000004FE00000000000004FE00000000000001B70000000001247874000000001357ABA5000000002589" +
        "BCB50000000048999AA500000000587679A5000000004654589500000000232247950000000011114795000000000001479500000000000147950000" +
        "000000014795000000000001479500000000000147950000000000014795000000000001479500000000000147950000000000014795000000000001" +
        "479500000000000147950000000000013684000066777777777777776677777777777777777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777766777777777777776677777777777777" +
        "00025ADFEC720000005EFFFFFFFE300000CFFEB99DFFE200007C8100018FF90000000000000EFA0000000000000CFB0000000000000EFA0000000000" +
        "005FF9000000000009FFE10000000002BFFF40000000006EFFD4000000000BFFFA1000000001DFFE300000000008FFC200000000002FFD1000000000" +
        "009FF3000000000000AFF0000000000000AFF8888888840000AFFFFFFFFFFE00005EFFFFFFFFF90000247899987521000048ABBBBBB952000059AA98" +
        "8ABB940000377544469BB50000233211136AB600000110000259B60000000000126AB60000000001258BA5000000001358BB8300000012469BB95200" +
        "0001258ABB852100001258ABA752100000158BB9632100000037BB8521000000004AB95210000000005BB84211111000005BC96444443200006CCB99" +
        "99997400005ACCBBBBBB9500003788888888730000156667777662000015566667776200001445556677720000133444566772000012223345677200" +
        "000112234567720000011123456772000001123456776200000123456677620000023456677652000013456676654100001456676654310000156776" +
        "654321000026777654322000002777765432200000278876543321000028887665443100002888876655410000278887766651000026777776665100" +
        "00037BEFFD940000008FFFFFFFFF800002FFFFFFFFFFF70002FFFFD9BFFFFB00009DB20001DFFF0000000000009FFF200000000000AFFF1000000000" +
        "08FFFD0000000001AFFFF9000000005DFFFFC200000018FFFFFD30000001AFFFFFB200000009FFFFD6000000005FFFFC1000000000AFFFC000000000" +
        "00BFFE100000000000CFFD888888850000CFFFFFFFFFFF3000CFFFFFFFFFFF40008FFFFFFFFFFC10003579AAA9853100007ACDDDDDCA6300009DEEDD" +
        "DEEDA500009CCA989BDEC7000067754347BED9000023221125ADD9000001000126AED9000000001259CEC800000001369CEEB6000001247ADEEC8400" +
        "001259CDEDB8420000259CEEDA7420000049CEDB96310000006BEEB742100000008DEC9521111000009DEC854333320000ADEDA999988500009DEEDD" +
        "DDDDB800008CDDDDDDDDB80000578999999885000045677877654300005789999987640000678999AAA97500005788899AAA86000045666789AA9700" +
        "0034445678AA97000022334568AA97000012234579AA97000012345789A98600002345789A99750000245789A99764000045799A9876430000579AA9" +
        "875432000068AAA9764321000079ABA876433200008ABBA976554300008ABBA9887654000079ABAA9987650000689AA9998865000056788888776400" +
        "00025ADFEA610000005EFFFFFFFD000000EF954345CFA00000340000002EF400000000000008FA00000000000006FA00000000000009F80000000000" +
        "005EE2000000000002DF7000000000007EF900000000003CFE700000000006FFC200000000008EF6100000000005FE4000000000001DF50000000000" +
        "006FD0000000000000AF80000000000000CF83333333200000CFFFFFFFFFF800005EFFFFFFFFE6000024678887642100003799999997520000477766" +
        "678973000024432224698400001111000148950000000000013795000000000001489500000000001369840000000012368963000000012478974100" +
        "000013589874200000014689853200000014688642100000002688642100000000389741000000000059852100000000005A853222221100006AA866" +
        "66654200005ABA9999987300004799988888730000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "01ADDDDDDDDDC40004FFFFFFFFFFFA00007AAAAAAAFFFA000000000008FFC300000000008FFE300000000007FFF100000000006FFD300000000005FF" +
        "F700000000000EFFFFF80000000005ACFFFFB1000000000017CFF80000000000000EFE00000000000006FF30000000000004FF50000000000006FF40" +
        "00000000000DFE000065000003AFF90003FFEA89CFFFD10002EFFFFFFFFB20000037CEFEC85000000078999999997400008BBBBBCCDCA60000688889" +
        "ACDDA500003444469BDB8400001111359BB8520000001258BB9521000000258BCA731000000137BDDA731000000138BDDB96410000012689ABBA7300" +
        "0000124568ABA60000000111247BB800000000000249B900000000000148B900000000000149B90000221111247AB8000055544457ABA60000899888" +
        "9ABB8400009BBBBBBBA74200006899998763210000006777776500000000778888760000000078899987000000007889998600000000678998760000" +
        "000067899875000000006789986500000000678987640000000067898865000000006788886500000000567888760000000045667776000000002345" +
        "677700000000223456770000000012345677000000002234567700000000333456770000000044556777000000005566777600000000667777650000" +
        "03CDDDDDDDDDD70009FFFFFFFFFFFD0008FFFFFFFFFFFD00019AAAAADFFFFC0000000003FFFFE3000000003FFFFE3000000004FFFFF3000000002DFF" +
        "FFC2000000005FFFFFFE700000002EFFFFFFF80000000268AFFFFE100000000003DFFF7000000000004FFF9000000000000FFFA000000000006FFF90" +
        "0023000004EFFF5008FFFD9AEFFFFD000AFFFFFFFFFFF50007FFFFFFFFFE40000049CEFFEA620000037999999999862005ACDDDDDDDDC83004ACCCDD" +
        "EFFFD9300368889BDEFEC83001233469DEEDA6200001147BEEDA6310000126ADEEC83100000149CEFEC8410000015ADEFEDB7410000148BCDEEDB720" +
        "000124689BDEDA400000112347ADEB5000000000137CEC6000000000026BEC6001111111148CEC600255543458BDEB50049BA999ACDED93005BDDDDD" +
        "DEEDA62004ACDDDDDDC9631002689AAA987521000047788888765300005899AAAA98630000589ABBCCB9740000478ABCCCB97300003679BCCCB96300" +
        "002579ACCCA86300002469ACCBA75200002469ACCBA85200002468ABCBA863000013579ABBA97400001345789AA984000012345689AA850000112345" +
        "78AA950000122335689A95000023344579AA95000035566789AA8500004778899AA9740000589999A998630000589999987652000047788776543100" +
        "019DDDDDDDDDC20001EFFFFFFFFFF8000025555555BFF3000000000003FF7000000000003EF5000000000001EF9000000000003EF8000000000002DF" +
        "F4000000000008FFFFB20000000001569DFE700000000000017DF300000000000007FC00000000000001FF20000000000000DF40000000000000EF30" +
        "000000000005FD0000000000001CF50002BA843359DFA00002EFFFFFFFF900000007BDFEB7100000006888888998620000789999ABCB830000456666" +
        "8ACB73000022223469A95200000001257985210000000247986310000000147AA74100000001269CA74210000001369BB97421000000246889975200" +
        "000012334689730000000001136895000000000001479500000000000137960000000000013795000011100012589500003332223579840000676666" +
        "789852000079999999863100005788887642100000067777777760000007777777777000000777777777700000077777777770000007777777777000" +
        "000777777777700000077777777770000007777777776000000777777777600000077777777660000006777777666000000677777766600000067777" +
        "766660000006777766665000000777776666500000077777666650000007777766655000000777776665500000077777666550000007777766665000" +
        "00000008E50000000000001FFA000000000000AFF2000000000001FFB0000000000008FF5000000000002FFB0000000000006FF5000000000001DFD1" +
        "000000000007FF5001510000000CFF000BF70000006FF7000DFA000000DFD1000DFA000002FFA3333EFB310005FFFFFFFFFFFF1004FFFFFFFFFFFF20" +
        "005555555EFC5500000000000DFA0000000000000DFA0000000000000DFA00000000000005D4000000000137863100000000026AB74100000000148B" +
        "A63100000000269B95200000000138BA73100000000259B95200000000147BB74100000000259BA5211110000137AB83234421000159BA6225774100" +
        "027BB84237A96200029BB74248BA731003ACB8667ACC951003BDCBA9ACDDB82003ACCCBBCDEECA3002688889ACDDB820013444456ACB951000111112" +
        "48BA73000000000137A96200000000002577410000000000000000000000000000000000122222222211111155677766554433225677776655544433" +
        "6677777665555443677877766666655477888777777776657888888788888776788888888889887678889889999999877788889999AAA98767788889" +
        "99AAA9986677788899AAA988556667788999998744455667788998773334445667788776111111122222222200000000000000000000000000000000" +
        "0000002DFB000000000000AFFF100000000002EFFE000000000009FFF900000000001EFFE300000000008FFFA00000000000DFFF400000000007FFFA" +
        "00000000000CFFF41AC70000006FFFB05FFE000000BFFF505FFF000005FFFC005FFF00000AFFFEDDEFFFDB200AFFFFFFFFFFFF900AFFFFFFFFFFFF80" +
        "029AAAAACFFFA920000000005FFF0000000000005FFF0000000000005FFF0000000000002BD8000000000147974100000000137BCA510000000025AD" +
        "DA510000000137CED941000000015ADEB620000000137CED9410000000149DEB731000000036BEDA633210000149DEC866763100027BEDA66AB95100" +
        "039CEC857BDB620004BEEC868CEC841005CEEDAABDEDA73006CEFEDDEEFEDB5005BDEDDEEFFFEB500389AAAACEEEB940013444469CEC852000111113" +
        "7CDB6310000000026ACA5200000000014897410000123567765421000123467887642100012357899764210001346899975421000245799997532100" +
        "034689A98753211003579AA98754321004689AA9876533200579AAA9877643200689AAA998876530079ABBAAA9987640079BBBBAAAAA975008ABBBBB" +
        "BBBBA86008ABBBBBCCCCB970079ABBBBCCCCBA8006899AABBCCCBA7005678899ABBBA970034566789AAA986002334456799987500112233567777540" +
        "00000006D20000000000000FF20000000000005FB0000000000000CF60000000000007FB0000000000000AF40000000000003FE0000000000000CF60" +
        "000000000004FD00000000000008F60003A40000002FE20008F9000000AF800008FA000001FE100008FA000005FE88888CFD850004FFFFFFFFFFFF10" +
        "009AAAAAADFDA7000000000008FA00000000000008FA00000000000008FA00000000000004E500000000013686310000000012588631000000001479" +
        "742000000000258963100000000147985210000000025897310000000013798521000000002489741111100001369853112321000147974213554200" +
        "02688521247863100379752125897310049A8533479A852005AB976679BCA730059BBA99ABCDC950037899999BCDB9400245555568BB973001122222" +
        "469A84200000000025896310000000001478521000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "777666677777777677777777777777767777777788888877777777778888888766777777888888876666777788888887566667777888888755566667" +
        "778888774455556677777776444455566777776600000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "0004CDDDDDDDC200000CFFFFFFFFF600000DFFAAAAAA9000000EFE0000000000000FFD0000000000000FFC0000000000000FFE9851000000000FFFFF" +
        "FE600000000EFFFFFFFD1000000144456BFFC0000000000000CFF20000000000002FF70000000000000EFA0000000000000EF90000000000005FF600" +
        "0000000002FFF000000000028EFF700000078ADFFFFA0000007FFFFFFE400000003EEDB85100000000168999999861000029BCCBBBBA8200002ACB98" +
        "88875100002AB96444432000002AB84211111000002BB95321100000002BCB8764310000002ACCBAA86420000028AAABBBA74100001456678ABA7100" +
        "0001222247AB920000000000148BA20000000000027BB30000000000137BB20000000000249BA2000001112358BB8200001344579AB961000026889A" +
        "BBA730000029BBBBA8631000002899876421000000266777766651000026788776655100002788877665410000278887655441000027888765443100" +
        "002788876544310000278887665431000027788776554100002677777766510000156677777662000014555667776200001334455677720000022233" +
        "456772000001122345677200000111223567720000011122456772000001223345677200001233445677720000134455677762000014556677766200" +
        "0008DDDDDDDDD400000FFFFFFFFFFC00000FFFFFFFFFFB00000FFFCAAAAAA200001FFF5000000000002FFF5000000000003FFFFEC7100000003FFFFF" +
        "FFE90000003FFFFFFFFFA0000007999CFFFFF300000000003CFFFA000000000002EFFE000000000000AFFF000000000000DFFF000000000007FFFD00" +
        "00000002AFFFF60000189BEFFFFFE00000BFFFFFFFFE300000DFFFFFFFA20000006FFEDA6300000000268999999973000049CDDDDDDDA500005BEEDD" +
        "CCCCA500005BEDB988886300005BEC9543332100005BEC9543211000005CEDCA87531000005BEEEDCB973100005ACDDDEEDB73000036889ABDEDA400" +
        "001334457ADEC6000000111137BEC6000000000025ADD7000000000026BED7000000011259CEC600001234569CEEB500003789ABDEEC8300005BCDDE" +
        "EDC95200006CDDDDCA8521000048AA98753210000035678887765300004789999987630000489AAA998763000059ABAA987652000059ABA987654200" +
        "0059BBA9876532000059ABAA986542000058AAAA98764200004799AAA9976300003678899AA974000024566789A984000012344679AA850000112235" +
        "689A950000112235689A95000012334679AA85000023456789A98400003567899AA9730000367899A997530000478999987642000036788776543100" +
        "0004CDDDDDDDA000000CFFFFFFFFF000000DF85555552000000EF30000000000000FF20000000000000FF10000000000000FF54200000000000FFFFF" +
        "DA300000000AEDDEFFF500000000000037FE500000000000004FD00000000000000BF10000000000000AF40000000000000AF40000000000000CF100" +
        "00000000007FC0000000000008FE20000002357ADFE40000003FFFFFFC200000002EEDA61000000000158888888751000028AAA9999851000028A976" +
        "665531000029974222221000002995210000000000299632110000000029B975432100000028AA998753100000168888998630000013444567985100" +
        "000111122579720000000000126982000000000002589200000000000258920000000000136982000000001125897100000122346898510000146678" +
        "997530000027899987531000002688765321000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000002C91000000000001DFF300000000001CFF700000000000CFF9000000000008FFB000000000005FFC100000000002EFF300000000000AFFDA" +
        "AA400000003FFFFFFFFA300000BFFFFFFFFFC00000FFFA0000AFFA0002FFB000000BFF0004FF60000007FF2005FF30000004FF4004FF30000005FF30" +
        "01FF80000009FF1000EFE500006EFD00005FFFC99DFFF3000006FFFFFFFE5000000037CEEC620000000000247974100000000147AB8410000000137A" +
        "B96310000001369BA7310000001269BA7410000000149BC9531000000137BDC964321000025ADEDBA9863100037BEEDCBBBA7310049CDCA99ABB9620" +
        "05ACB854458AB83005BC85211258BA4006BB63100136BA5006BA52000026AB5005BA62100136BA5005AB84211258BA50049BA864568BB830026ABB99" +
        "9ABBA6200137ABBBBBB96310001367999875310002345677765432100345677876543210035678888755322004578899876543200568999987654320" +
        "06789AA99876543006899AA9987765400789999988777650078999887777776007888876666777600788776555667770077776544456677007776543" +
        "344567700777654333456770077765433445677007776654445667700677766555667770067777666667776005677777777776600456677777776650" +
        "000000008DC3000000000009FFF800000000007FFFF70000000005FFFFC1000000003EFFFD1000000001CFFFD20000000009FFFF40000000002FFFFF" +
        "FFC4000000AFFFFFFFFFA00000EFFFFFFFFFF90005FFFFD77DFFFE1009FFF90000AFFF600AFFF300004FFF800AFFD000000FFF900AFFF200003FFF80" +
        "07FFFA1001BFFF6002EFFFC9ACFFFE10009FFFFFFFFFF800001AFFFFFFFF9000000049DFFD940000000000248995200000000248CDB730000000148C" +
        "EEB62000000137BEEC94100000126ADED952000000259DEEB73100000148CEFEB8532100026AEFFEDBA75210038CEFFEEEDC952004ADFEDCCDEEC930" +
        "05BEECA889CEEB5006CED963358CEC6006CEB731126BED6006CEB620015ADD7006CEC731136BED6005BEDA64469DEC6004ADEDB99BCEDB50037BDEED" +
        "DDEEC8300147ACDDDDDB85100013689AA9864200001234677765321001124578987532100124679AA9753210023579AAA975321003468ABBA9753210" +
        "0468ABCCB97643200579BCCCBA875420068ABCCCBA987530069BCCCBBBA98650079BBBAAAAAA9860079AAA98899AA870089AA9777799A970089A9865" +
        "5689A970089A98656689A970079A98766799A970079AA988889AA8700689AA9999AA986005789AAAAAA9875004578999998875400245677877765420" +
        "000000002A70000000000000AFA0000000000008FC2000000000006FD1000000000005FE2000000000003DF5000000000000CF90000000000006FF77" +
        "74000000002EFFFFFFD90000005FFFDAADFF800000BFC200002DF50001EF20000003FA0004FE00000000FF1005FC00000000EF3004FD00000000EF30" +
        "01FF00000000FE00008F60000008F700003EF853359FD3000003FFFFFFFE2000000036BEEB6200000000012578631000000012479752000000001469" +
        "853100000001368863100000001368974100000000258A96421100000147BCA865421000026ACCBA98753200038BCBA999986410049BB87667898620" +
        "04AA74222347983005A95210012589400598410000148950059741000014895005984100001489500498521111269940038975333358983002689876" +
        "678985200146899999986310001356788764310000000000000000001222222222222222556666777777776655666777777777665666777777777776" +
        "666777778877777666777788888777776677788888887777677788888888777777778888888877777778888888887777777888888887777777788888" +
        "887777777778888877777766777777777777766677777777777766667777777777666666777777776666665522222222222222220000000000000000" +
        "02CDDDDDDDDDDB2005FFFFFFFFFFFF50019AAAAAAAACFF5000000000000BFE2000000000008FF9000000000003FFF000000000003EFF400000000000" +
        "AFF9000000000009FFA000000000003EFD100000000000BFE4000000000000FFA0000000000006FF5000000000000AFE1000000000000EFA00000000" +
        "00002FF70000000000003FF50000000000004FF50000000000005FF50000000000001CD2000000000279999999999720029BBBBBBBCCCA2002788888" +
        "89ACDA3001344444568BC92000111111258BB7200000000136ABA5100000001269BA7310000000259BB8410000000147BB95200000001269B9621000" +
        "0000148BA7310000000026AB85100000000137BB63100000000149BA5200000000025AB94100000000026AB83100000000036BA72100000000037BA7" +
        "2000000000037A9620000000000258741000000000000000000000006777777777777666777777777777776677777888888887767777788888888877" +
        "777778888888887767777888899888876677788889988887666777888898888756667788888888775566777888888877556667788888877645566777" +
        "888877764556677778877766455667777777766545566777777766554556677777766655556667777766655455666777766655440000000000000000" +
        "04DDDDDDDDDDDC300AFFFFFFFFFFFF9009FFFFFFFFFFFF9001AAAAAAAAEFFF600000000005FFFF10000000002EFFFB0000000001CFFFF10000000009" +
        "FFFF50000000006FFFF90000000000DFFFB00000000006FFFD20000000000AFFF700000000000DFFD000000000003FFFA000000000006FFF60000000" +
        "00008FFF3000000000009FFF000000000000AFFE000000000000AFFD0000000000004EF700000000038999999999973005BDDDDDDDDDDB5005ACCCCC" +
        "DDEFEC50037888889BDEEC500123333469DEEA40000001136ADEC830000000259DEDA51000000148CEDB73100000126BDEB841000000149DEC841000" +
        "000026BEDA620000000148DEC731000000015ADDA510000000026BEC8310000000027CEB7200000000038CEB6200000000038CDA5100000000038CDA" +
        "5100000000037BC8410000000002588631000000026777788887651002788999AAA987200278999AABBA98200267889ABBCBA82002566789ABCBA820" +
        "01345679ABBB972001234578ABBA862000124579AAA9751000124679AA97641000235789A98643100024689A98753200013568999754210001357999" +
        "864311000146899975321000014689986431100001578997632100000157899753210000015789865311000001578886421000000146777542100000" +
        "01CDDDDDDDDDDA1002FFFFFFFFFFFF30004555555555FE10000000000005FA0000000000003FF3000000000000BF9000000000000AFD000000000000" +
        "6FC2000000000002EF4000000000000DF70000000000004FC0000000000000CF60000000000003FD10000000000008F90000000000000AF400000000" +
        "00000BF20000000000000DF00000000000000EF00000000000000FF00000000000000AA00000000002788888888886200289999999ABB92002566666" +
        "678AB920012333333469A820000000011369961000000001247975100000000247985300000000136886310000000125887410000000024897420000" +
        "000013698521000000002489731000000001369852000000000147974100000000015896310000000002599520000000001269852000000000137984" +
        "100000000013798410000000001268631000000000000000000000000000000000000000777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777700000000000000000000000000000000" +
        "000026CEEB6200000003EFFFFFFD2000001EFFD99EFFD000007FF700008FF500009FE000001FF700007FF300005FF600001EFF4005FFD0000006FFFB" +
        "BFFF500000006EFFFFE5000000006FFFFFF50000000AFFF88FFF9000007FFD3003EFF60000FFD100002DFE0003FF60000008FF2005FF20000004FF40" +
        "04FF60000008FF2000FFC400004DFE00009FFFB98BFFF7000008FFFFFFFF8000000048DEFE940000001357899874210000269BBBBBA852000159BB98" +
        "9ABB8410016AB954469BA610017BB731238BB710017BB842248BA6100159BA7668AB95100037ACBAABCA630000259CDDDDC9520000259CDDDDC95200" +
        "0148BCA99ACA7410017AB96556ABA610029BA631136AB92002AB83100148BA2003BB73000137BA2002AB84111248BA20029BA754457AB920017ABA98" +
        "89BBA7100147ABBBBBBA7400002367999986310000356777776653000036777777776300004778777787740000478877778874000047887777887400" +
        "004788888888740000478899998874000047899999987400004789AAAA987400004789AAAA9874000047899999987400004788888888740000477766" +
        "667774000047765555677400005766544567750000576654456775000057765555677500004776666677740000477777777774000036677777776300" +
        "000039DFFD8300000007FFFFFFFF6000005FFFFFFFFFF30000AFFFCAADFFFA0000AFFD1002EFFA0000AFFE2003FFFA0000AFFFE45EFFF900003FFFFF" +
        "FFFFE2000006FFFFFFFF50000006FFFFFFFF5000006FFFFFFFFFF40001EFFFE44EFFFC0007FFFC1002DFFF500AFFF300004FFF900AFFE000002FFF90" +
        "09FFF700009FFF8005FFFFB99CFFFF3000BFFFFFFFFFFA00002BFFFFFFFFB10000015AEFFFB610000013689AA98531000137BCDDDDCA6300026BDEDD" +
        "DEEDA510038CEDB99BDEC710039DEC8558CEC810038DEC8559CEC710027CEDB99CEEB610015ADEEDDEED94100148CEFFFFEC73000149DEFEEFEC8410" +
        "027BEEDCCDEDB61004ADEC978ADED92005BED95336ADEB3006CEB731138CEC3006CEB731138CEC3005CEDA6557ADEB3004ADEDBAABDED920037BDEED" +
        "DEEDB7100147ACCDDDCA731000135788987531000134567787754310014678999997641001579AAAAAA986100268ABBBBBBA97200279ABBAABBA9720" +
        "0279ABBBBBBB97200279ABBBBBBB97200269ACCCCCCB97200269ACCDDDCB97200269BCCDDCCB97200279ABCBBCBB97200279ABAAAABAA820028AAA98" +
        "89AAA820028AA987789AA930029AA987789AA930028AA988889AA9200289AA9999AA982002689AAAAAA9872001568999999875100135677887765410" +
        "000016BEEA5100000001EFFFFFFD0000001CFA5445AFB000005FB000000DF400005F9000000AF500004FD000001EF300000CF90000AFB0000000EFE8" +
        "8FFD000000001AFFFF91000000001CFFFFB100000003FFB22BFF2000002EF700008FD200009F50000007F80002FE00000000FF1005FC00000000EF30" +
        "04FD00000000FF2000CF30000005FA00004FE753347FF3000006FFFFFFFF5000000047CEFD8400000012467887642100002479999997520000379876" +
        "67897400005896422369960000599521014896000058963113599500003798644589740000247A9999A8520000136ACCCCA7310000136ABBBBA74100" +
        "003589877898630000579753347985000079842112479700008962100025890000985200001589000089621000258900007985322347980000589866" +
        "567986000036899999986300001356888875310000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "777777777777777677777777777777777777888888887777777888888888777777788888888887777778888888888777777888888888777777778888" +
        "888877777777777777777777677777777777777600000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000016CEEC6200000002DFFFFFFD3000001DFFD99DFFE20000BFF800007FFB0000FFA000000AFF0004FF50000006FF2005FF30000004FF4004FF5000" +
        "0006FF3001FFA0000009FF1000CFF900005EFF00001DFFFFFFFFF9000003BFFFFFFFE200000005AAAEFF9000000000003FFE100000000001CFF50000" +
        "0000000AFF9000000000009FFD000000000008FFD100000000004FFD2000000000001AB200000000001357899875310001369ABBBBB963100259BBA9" +
        "9ABB9520038BB965569BB83004AB85211258BA4005AB73100137BA5006BA62000026AB5005BA62100026BB5005AB84211248BA50049BA754457ACA40" +
        "026ABBA99ABCC8300147ABBBCCDDB7200013689ABDED9510000123469CDB6310000011359CB9410000000137AB9621000000137ABA6310000001369B" +
        "A7410000000148BA74100000000147974200000005567777777665400567777777777650067776666677766006776655566777600777655445567760" +
        "077665434456777007765433345677700776544334567770077765444556777007776655566787700677766667788870066777777888887005667788" +
        "8999987004567788999997600345678999A9876002345678999986500223567889987540012345788887653001234567887654300123456777654320" +
        "000038DFFD8300000006FFFFFFFF7000005FFFFFFFFFF60001CFFFDAADFFFC0005FFFC2002CFFF5009FFF400004FFF800AFFE000000FFF900AFFF200" +
        "002FFF9008FFF900006FFF7002FFFFD77CFFFF30009FFFFFFFFFFD00000BFFFFFFFFF90000005CFFFFFFE10000000004FFFF80000000001DFFFC1000" +
        "000001CFFFE3000000001CFFFF60000000009FFFF90000000000AFFFA000000000004DD7000000000013579AA98632000136ACDDDDCB7410026ADEED" +
        "DEEDB830039DEDBA9BDEDA4005BEDB7546ADEC5006CEC841136BED6006CEB620015ADD7006CEB731125ADD6005CED963348CED6004ADECA889BDEC60" +
        "038CEEDCCDEEEB500158BDEEEFFFD930002479BCEFFEB7200011357ADFED94100000135ADEDA62000000148CEEB73100000138BEEC84100000025ADE" +
        "C952000000026BDC95210000000158A8521000000235677777765320035688999988653005689AAAAAA986500679AAA99AAA9760078AA998899AA860" +
        "079A99777799A970089A98666689A970089A98655689A980079A9976679AA970079AA98889AAA9700689AAAAAABBA97005789AABBCCBA860045789AB" +
        "CCCBA850034578ABCCCB97400234679BCCBA85300123579ABBA864200123579ABA9753200123579AA976421001245789876431100123567876432100" +
        "000015AEEA6100000000DFFFFFFD1000001CFB5445AFC200006FB000000AF60000DF00000000FD0003FE00000000FF2005FC00000000EF3003FE0000" +
        "0000FF2000CF20000000FD00006FC2000008F9000009FFDAACFFF50000009DFFFFFFD1000000004779FF40000000000009FC0000000000004FD30000" +
        "00000002EF5000000000001DF7000000000002CFA000000000000BFA0000000000000891000000000012467887642100013589999998631002589877" +
        "778985200379854334589730048963111136984005984100001589500597410000148950059741000014895004985210012589500489742222479940" +
        "02689866678AA830014789999ABCA72000235789ABCC9510001124568ACA73100000112469A842000000001479952100000001368873100000001358" +
        "974200000001257984210000000136875210000000000000000000004444444444555554666666777777777766667777777777776677777777777777" +
        "777777777777777777777778888777777777778888877777777778888888777777778888888777777777888888877776777788888877777677777888" +
        "877777667777778877777766777777777777766677777777777766656777777777766665667777777666665544444544444444330000000000000000" +
        "000001CFFD300000000005FFFF600000000007FFFFA0000000000CFFEFE0000000001FFEAFF4000000005FF96FF800000000AFF51FFD00000001EFE1" +
        "0CFF20000004FFB007FF50000005FF7003FF60000008FF2000DFB000000CFE8888DFE000002FFFFFFFFFF400006FFFFFFFFFF90000BFF755555FFD00" +
        "01FFE000000BFF3005FFA0000007FF7009FF60000002FFA00AFF20000000DFA005E9000000008F700000247997420000000137ACCA631000000158CE" +
        "EC84100000026ACDDC95200000137ACCCCA6200000148BBAABB7310000149BA89BB841000126AB9669B952000137BB8559BA62100148BB7448BB7310" +
        "0259CB8558BC8410026ADCA88ACC9520037BDDCAACDDA630048CDCCBBCDDB740059CCA9889ACB85006ABA754457AB96008BB84211248BA7008BA6310" +
        "0137AB8008A9520001259A8007873100001478700234567887654320023467888876542003457899998754300345789AA98765300356789AA9876530" +
        "0456789AA998654004568899998866400466889999887640046788999988765005678899998876500577899999887750067889999998876006789999" +
        "999988600688999999998870078999999999987007899999999998700788998888998870078888877888887007787776667777700777765555667770" +
        "000005EFFF80000000000DFFFFF0000000000FFFFFF1000000002FFFFFF5000000007FFFFFFA00000000CFFFFFFE10000002FFFFFFFF40000007FFFD" +
        "AFFF9000000AFFF95FFFA000000BFFF41FFFD000001FFFE33CFFF300005FFFFFFFFFF80000AFFFFFFFFFFD0000FFFFFFFFFFFF0001FFFEAAAADFFF20" +
        "04FFFB00008FFF700AFFF600003FFFC00DFFF200000FFFF00DFFF000000DFFF006FF80000005EF9000012589A8621000000148CDDC94100000025ADF" +
        "FEB6200000026BEFFEC7310000137CEFFFD8410000149DEFFFD952000026AEEEEEEB62000037BEDCCDEC83100138CECAACED94100149DEC99BEDA510" +
        "026BEECBACEEB620027BEFEDCDEEC830038CEFEEEEFED94004ADEEDCCDEEDA5005BEECA889BEEB6007CED964459DEC7008DEC731136BED9008DDB620" +
        "025ADD9008CC94100148BC80058863100125886001234677776432100124578998764310023569ABBA975320023579BCCBA8642002468ACCCCA86430" +
        "03579BCDDCB9753003579BCDDCB9753004679BCDDCBA86400468ABCCCCBA86400579ACCCCCCA97500579BCCCCCCB9850068ABCCCCCCBA860068ABCCC" +
        "CCCBA970079ABCCCCCCBA970079ABBBBBBBBA970089ABAAA9AABA980089AA998889AA980089998766789998007888655556888800777654334567770" +
        "0000009FF8000000000003FFFE100000000007FEFF50000000000BF89F90000000001FF23FD0000000005FC00EF3000000009FA00AF600000000EF80" +
        "09FC00000003FF4006FF10000005FF1002FF40000007FA0000CF5000000BFA3333BF8000000EFFFFFFFFC000005FFFFFFFFFF300009F8000000AF600" +
        "00DF40000006FB0003FE00000001FF1007FB00000000CF400AFA00000000AF8005D3000000005D4000002478864100000001369BB9620000000148BC" +
        "CA731000000259BBBA841000001269A99A95200000136997799621000014798668973100002489745797410000258963469852000136985335886310" +
        "01479853358974100259A965569A8410026ABA9889BB9520037ACBAAAABCA620038ABA9999ABA730048A97655679A830049974222247984005985200" +
        "012589500597410000148950058631000013685002456677776654200245677777765420024567788776653003556778877665300356677887766530" +
        "035667777776653003566777777665300356677777766530035666777766653003566677776665300366677777766630046677777777664004677777" +
        "777776400467777777777640047777777777774004778888788777400477888778877750057777777777775004777777777777500477777667777740" +
        "0CFFFFFFEDA400000FFFFFFFFFFFB1000FFEDDDDEFFFFB000FFA0000027FFF400FFA00000005FFA00FFA00000001FFC00FFA00000003FFB00FFA0000" +
        "138FFF500FFFFFFFFFFFFC000FFFFFFFFFFFF2000FFEDDDDDFFFFE100FFA0000016EFF900FFA00000003FFE00FFA00000000AFF00FFA00000000AFF0" +
        "0FFA00000003EFE00FFA0000028EFF900FFEDDDDEFFFFE100FFFFFFFFFFFC2000CFFFFFFEC950000047888888765311007BCCBBBBBBA742009CCBAAA" +
        "ABBBA73009CB8666679BB95009BA63222359BA7009B952111147BB7009BA63212258BB7009CB8655568ACA6009CDBA99AACDC95009DDCCBBCCDDC940" +
        "09CCB9999ABCC95009CB8655568ACA7009BA63222247BB8009B952000125AB9009B952101136AB9009BA63222358BB8009CB8666678ABA6009CCBAAA" +
        "ABBBA74007BCCBBBBBBA7520047888888765321005677777776655400677877777776650067887777777766007788766667777600778776666778770" +
        "078887666677887007888766667888700788887777889870078998877889987007899888888998700789988777899870078887766778887007887765" +
        "566788700777765555677870077776544556777007777655556777700678776666677770067887776777776006778777777766500567777776666540" +
        "1CFFFFFFFEB710005FFFFFFFFFFFE4005FFFFFFFFFFFFE205FFFEDDEFFFFFF605FFF700005CFFFA05FFF7000004FFFC05FFF7000006FFFC05FFFA555" +
        "6BFFFFA05FFFFFFFFFFFFF605FFFFFFFFFFFFF105FFFFFFFFFFFFF705FFFB8889AFFFFD05FFF7000005FFFF35FFF7000000CFFF45FFF7000001EFFF4" +
        "5FFF700016CFFFF15FFFEDDEFFFFFFB05FFFFFFFFFFFFF405FFFFFFFFFFFF6001CFFFFFFFDB61000157899998876421028BDDDDDDCCA852029DEEDDD" +
        "DEEDC8412AEEDBBBBCDEDB612ADDB76568ADEC822ADDA532358CED822ADDA533358CED822ADEC97789BEEC822AEEECCCCDEEEC712AEFEEEEEEFFEB61" +
        "2AEEECCCCDEEEC722ADEC98789BDED922ADDA633358CEDA22ADD9411126ADEA32ADD9522247BDEA22ADDB76567ADED922AEEDBBBBCDEEC7229DEEDDD" +
        "DEEDC95128BDDDDDDCCA8620157899998876421015677888877654312689999999987641279AAAAAAAA99761289AAAA99AAAA87228AAAA9999ABA972" +
        "28AAA99889ABBA8228AAAA9899ABBA8228ABBA999ABBBA8228ABBBAAABBCBA8228ABBBAAABBCBA8228ABBBAAAABBBA8228ABAA9999ABBA8228AAA987" +
        "889ABA92289A9877789AAA92289AA987789AAA82289AA988899AA982289AAA9999AAA972279AAAA99AA9986226899999999876511567788887765431" +
        "08FFFFFFED8200000FFFFFFFFFFF80000FF5333345AFFB000FF200000005FF000FF200000000CF400FF200000000BF600FF200000000DF400FF20000" +
        "001AFF000FFBAAAAABFFF8000FFFFFFFFFFFE0000FF9888889EFFB000FF200000007FF200FF2000000009F800FF2000000005FC00FF2000000004FD0" +
        "0FF2000000008FA00FF200000004EF400FF5333345BFFE000FFFFFFFFFFF900008FFFFFFEC830000058999888876431007ABAA999998752008BA8766" +
        "6789984008A86433335799600897421111358970089741000024797008A852111236897008A975444568A96008BB988888ABB94008CCBA9AAABCB940" +
        "08BB9877789AA95008A965444467996008A842111235798008974100001369800897410000136980089742111124798008A864333357897008BA8766" +
        "6789985007ABAA9999997520058999888876431000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000022222222222222227777666666667777777776666667777777777766666777777777776666677777777776666666777777776666" +
        "566667772222222222222222000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "00000049CC82000000001CFFFFFF60000000BFFFFFFFF3000003FFE3006FFA000009FF40000CF800000EFE0000000000003FFA0000000000005FF700" +
        "00000000007FF60000000000008FF50000000000008FF50000000000007FF70000000000005FF80000000000003FFA0000000000000EFE0000000000" +
        "0009FF30000CF8000003FFD1005FFA000000CFFFEEFFF50000002DFFFFFFA0000000006CFEB500000001245788753100000248ABBBBA710000037ABB" +
        "AABBA20000169BA8667AA2000017AB84334772000029B96211133100002AB85100011000002AB84100000000002AB73100000000002AB73100000000" +
        "002AB73100000000002AB73100000000002AB84100000000002AB851000110000029B962111331000017AB842247720000169BA7557AA30000047ABB" +
        "AABBA200000248ABBBBA8200000124689986410000000567777000000000067777700000000007777660000000000777655000000000076654400000" +
        "000007654330000000000664322000000000065432100000000006542110000000000654211000000000065421100000000006542110000000000654" +
        "321000000000066532200000000007654330000000000776554000000000077766500000000007777760000000000677777000000000056777700000" +
        "0000047BDDB5100000004FFFFFFFB0000003FFFFFFFFF9000008FFFFDEFFFA00001FFFFB02EFFA00006FFFB0003FF800009FFF800000000000AFFF40" +
        "0000000000AFFF200000000000AFFF100000000000AFFF100000000000AFFF200000000000AFFF5000000000009FFF8000000000006FFFB0004FF800" +
        "001FFFF901EFFA000009FFFFBDFFFA000004FFFFFFFFFA0000005FFFFFFFD1000000049EFFD72000000135799986420000136ACDDDCB83000026ADEE" +
        "EEEEB5000038CEEDBCDEC600004ADEC978ACB600005BEDA634688400006CEC8311244200006CEB7200111100006CDA6200000000006CDA5200000000" +
        "006CDA5200000000006CDA6200000000006CEB7200111100006CEC8311244200005BEDA533588400004ADEC977ACB6000038CEECBBDEC6000026ADEE" +
        "EEEEB60000136ACDDDDB840000013689AA975200001345678877620000146789999872000015789AAAA98200002689AAAA99820000279AA998877200" +
        "002899987666520000289987654441000028998643222100002899753211100000289875321110000028997532111000002899754221100000289986" +
        "43322100002899976544410000279A987766620000279AA998887200002689AAAA9982000015789AAAA9820000145789999872000013456777775200" +
        "00000028CC710000000008FFFFFD600000006FFB56BFF0000002EE60000BF4000009F80000018000000DF50000000000000FF40000000000000FF200" +
        "00000000002FF10000000000002FF00000000000002FF00000000000001FF10000000000000FF20000000000000FF40000000000000DF50000000000" +
        "0009F800000290000002FE40000BF40000007FF8339FF10000000AFFFFFF80000000004AEEA200000001246788864100000247999999720000136998" +
        "667882000014798533467200001698531113410000279741000110000028963100000000002896210000000000299521000000000029952100000000" +
        "002995210000000000299621000000000028963100000000002897310000100000279742000121000016886311244100001479864357720000035898" +
        "778982000001368999997200000013578876410000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "008EEC830000000000FFFFFFB400000000FFFFFFFF60000000FFC16EFFF5000000FFC0018FFF300000FFC0000AFFA00000FFC00002EFF20000FFC000" +
        "006FF90000FFC000001FFD0000FFC000000DFF0000FFC000000CFF0000FFC000000DFF0000FFC000001FFE0000FFC000007FF90000FFC00002EFF400" +
        "00FFC0002BFFA00000FFC14AFFFF300000FFFFFFFFF5000000FFFFFFFB300000008FFEC7200000000168987642100000028BCCBA86421000029CCCBB" +
        "B974200002ACB99ABB97310002AB96568AB96300029B953358BB8410029B852136ABA610029B8520147AB820029B8510036AB820029B85100259B920" +
        "029B85100259B920029B85100259B920029B8510026AB920029B8521148BB720029B952236ABA61002AB964469BB841002ACB989ABB95200029CCBBB" +
        "BB963100028BCCBBA8631000016899876421000000067777665430000007788877654000000788888776500000078888777760000007887777776000" +
        "000777766667700000077765556670000007765545567000000666544456600000066654444560000006665434456000000666544445600000077654" +
        "445660000007766555567000000777666666700000077777667770000007888777776000000788888777600000078888777650000006777776654000" +
        "01CFFEB61000000005FFFFFFFA10000005FFFFFFFFC2000005FFFFFFFFFD100005FFF99FFFFFA00005FFF703EFFFF30005FFF7002FFFFB0005FFF700" +
        "08FFFF1005FFF70001EFFF4005FFF70000AFFF5005FFF700007FFF5005FFF700009FFF5005FFF70001EFFF5005FFF70008FFFF2005FFF7005FFFFB00" +
        "05FFF86BFFFFF30005FFFFFFFFFFA00005FFFFFFFFF9100005FFFFFFFF80000001CFFFDA60000000026999875321000002ACDDDCA852100003BEEEEE" +
        "DB95310003BEEEDEEEC9520003BEDCABDEEB841003BDD977ADEDA61003BDC8546BDEC82003BDC84248CED92003BDC83126ADDA3003BDC831159DDB30" +
        "03BDC831149DEB3003BDC831159DDB3003BDC84126ADDA3003BDC84348CED92003BDD9668BDEC82003BEDB9ABDEDA61003BEEDCDEEDB841003BEEEEE" +
        "EDB8520002ACDDDDCA7421000269999875311000015778876543210002789AA9886532100289ABBAA9865310028ABBBBAA986510028ABBBAAAA97610" +
        "028AAAA99AA987200289A99899AA97200289987778999820028998766799982002899866678998200289986667899820028998766789982002899877" +
        "789998200289A988899A9820028AAA9999AA9720028ABBAAAAA98620028ABBBBAA9875100289ABBAA9875410027899A9987543100156777766542200" +
        "008EDB610000000000FFFFFE8100000000FF96BFFF20000000FF7003CFD2000000FF70001BFC000000FF700001FF700000FF7000006FF20000FF7000" +
        "000EF80000FF7000000AFC0000FF70000007FF0000FF70000006FF0000FF70000007FF0000FF70000009FD0000FF7000000DF90000FF7000007FF200" +
        "00FF700001FF900000FF70005DFD100000FF958DFFD2000000FFFFFFE8000000008FFDA61000000000578876421000000079AA9875310000008AA988" +
        "98641000008A965689863100008A74335898520000897411368964000089731013688600008973100247970000897310013798000089731001369800" +
        "00897310013698000089731001369800008973100137980000897310024797000089741124698500008A742246897400008A865578985200008AA988" +
        "998631000079BAA987531000005788875421000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "00CFFFFFFFFFFB0003FFFFFFFFFFFF2004FFEDDDDDDDD90004FF90000000000004FF90000000000004FF90000000000004FF90000000000004FF9000" +
        "0000000004FFFFFFFFFE700004FFFFFFFFFFF00004FFEDDDDDDD600004FF90000000000004FF90000000000004FF90000000000004FF900000000000" +
        "04FF90000000000004FF90000000000004FFEDDDDDDDD90003FFFFFFFFFFFF2000BFFFFFFFFFFB000047888888887600007ACCBBBBBBA800009CCBAA" +
        "AAAA9700009CB86655555400009BA63222222100009B952111100000009BA63211111000009CB86555443100009CDBA999985300009CDCCBBBBA7300" +
        "009CCB9999875300009CB86555543100009BA63211111000009B951000000000009B952100000000009BA63222222100009CB86655555400009CCBAA" +
        "AAAA9700007ACCBBBBBBA800004788888888750000002667777200000000277877720000000027888772000000002788776200000000278876520000" +
        "000028887652000000002888776200000000288887620000000028998872000000002899987200000000289988720000000028888762000000002788" +
        "765200000000277765510000000027776541000000002777665100000000278776620000000027887772000000002778777200000000266777720000" +
        "02DFFFFFFFFFFD2009FFFFFFFFFFFF8009FFFFFFFFFFFF8009FFFEDDDDDDDB1009FFF4000000000009FFF4000000000009FFF4000000000009FFF855" +
        "5555200009FFFFFFFFFFE30009FFFFFFFFFFF50009FFFFFFFFFFF30009FFFA888888400009FFF4000000000009FFF4000000000009FFF40000000000" +
        "09FFF4000000000009FFFEDDDDDDDB1009FFFFFFFFFFFF8009FFFFFFFFFFFF8002DFFFFFFFFFFD200168999999998620029CDDDDDDDDC92002AEEEDD" +
        "DDDDCA2003BEEDBBAAAAA82003BEDA755555531003BEC9422222210003BEC9532222110003BEDB877776420003BEEDCCBBBA841003BEFEEEDDDC9510" +
        "03BEEDCCBBBB841003BEDB987776520003BEC9533332210003BDC8311111000003BEC8421111110003BEDA755555531003BEEDBBAAAAA82002AEEEDD" +
        "DDDDCA20029CDDDDDDDDC92001689999999986200035778888777400004789999998850000489AAA999885000059AAAA988774000059AAA987766300" +
        "0059AAA9876653000059AAA9876653000059ABA9987663000059BBBAA98763000059BBBBA98874000059BBBA998763000059ABA9877653000059AA98" +
        "765542000059AA98755442000059AA98765542000059AAA8876653000059AAA99887740000489AAA9998850000478999999885000035778888777400" +
        "00AFFFFFFFFFFD0000FFFFFFFFFFFF0000FF63333333310000FF40000000000000FF40000000000000FF40000000000000FF40000000000000FF4000" +
        "0000000000FFBAAAAAA9300000FFFFFFFFFFD00000FFA8888888200000FF40000000000000FF40000000000000FF40000000000000FF400000000000" +
        "00FF40000000000000FF40000000000000FF63333333310000FFFFFFFFFFFF0000AFFFFFFFFFFD0000378998888884000049BAA999999500005AA876" +
        "66665300005A964333332100005A8421111110000059841000000000005A853211111000005A975444442100005BB98888775200006BCBAA99986200" +
        "005AB98777764200005A975444432100005A85211111100000598410000000000059741000000000005A842111111000005A964333332100005AA876" +
        "666653000049BAA999999500003789988888840006677777776666600677777777766660077777777777666007777777777776600777777777777760" +
        "077777777777777007777777777777700777777777777770077777777777777007777777777777700777777777777770077777777777777007777777" +
        "777777700777777777777770077777777777777007777777777777700777777777777770077777777777766006777777777766600677777777766660" +
        "000BFFFFFFFFB100004FFFFFFFFFF400004FFEDDDDDD9000004FF90000000000004FF90000000000004FF90000000000004FF90000000000004FF900" +
        "00000000004FFFFFFFF90000004FFFFFFFFD0000004FFEDDDDD70000004FF90000000000004FF90000000000004FF90000000000004FF90000000000" +
        "004FF90000000000004FF90000000000004FF90000000000004FF80000000000000AC2000000000000057888888860000008BCBBBBBA80000009CCBA" +
        "AAA970000009CB97666650000009BA63222220000009B952100000000009B952111100000009CB8544431000000ACCA988753000000ADDCBBBA84000" +
        "000ADDBAAA9740000009CB97665420000009CA74332210000009B952100000000009B951000000000009B951000000000009B951000000000009B951" +
        "000000000008A84100000000000786310000000000005677777700000000677787770000000067888877000000007788877600000000778887660000" +
        "000077887766000000007788776500000000788877660000000078888766000000007888887600000000789988770000000078999887000000007899" +
        "988700000000789998870000000078888877000000007888877600000000788877650000000077777654000000007777654400000000677665430000" +
        "003DFFFFFFFFE400009FFFFFFFFFFA00009FFFFFFFFFF900009FFFEDDDDDC200009FFF4000000000009FFF4000000000009FFF4000000000009FFFBA" +
        "AAA80000009FFFFFFFFF5000009FFFFFFFFF5000009FFFFFFFFF3000009FFF6333310000009FFF4000000000009FFF4000000000009FFF4000000000" +
        "009FFF4000000000009FFF4000000000009FFF4000000000009FFF4000000000002DEB00000000000036899999986300005ACDDDDDDCA500006CEEED" +
        "DDDCA500006CEECBAAAA8400006CEC9655554200006CEC8422221100006CEC8432221000006CEDA776642100006CEEDBBBA84100006CEFEEDDCA5200" +
        "006CEEDCCCB95100006CEDB988753100006CEC8533321000006CEB7311100000006CDB6200000000006CDB6200000000006CDB6200000000006CDB62" +
        "00000000005BCA6200000000004897410000000000156778887762000026789999887200002789AAA998820000279AAAA988720000289AAA99876200" +
        "00289AAA9876520000289AAA9876510000289ABA9876510000289ABBA98762000028ABBBA98762000028ABBBBA8762000028ABBBA987620000289ABA" +
        "9876510000289AA987654100002899987543210000289987643210000027898754211000002788865321000000278876431100000026776543110000" +
        "000BFFFFFFFFB000000FFFFFFFFFE000000FF63333331000000FF40000000000000FF40000000000000FF40000000000000FF40000000000000FF400" +
        "00000000000FFBAAAAA50000000FFFFFFFFB0000000FFA8888710000000FF40000000000000FF40000000000000FF40000000000000FF40000000000" +
        "000FF40000000000000FF40000000000000FF40000000000000FF30000000000000AA0000000000000057899888870000007ABAA999980000008BA98" +
        "777760000009A975444430000008A8521111100000089741000000000008A852111100000009A964333210000009BA97766530000009BCBA99874000" +
        "0009BCA9998740000009BA87665430000009A964332210000008A8421110000000089741000000000008974100000000000897410000000000089741" +
        "000000000008974100000000000786310000000004777777777766400477777777776640047777777777764004777777777777400477777777777740" +
        "047777777777774004777777777777400577777777777750057777777777775005777777777777500577777777777750057777777777775005777777" +
        "777777500577777777777750047777777777774004777777777777400477777777777740047777777777764004777777777766400467777777766640" +
        "0000007CEDA4000000007FFFFFFFD2000005FFFEACFFFB00002FFF80002CFF10006FFA000002A60000BFE1000000000001FF80000000000004FF5000" +
        "0000000005FF50000000000005FF5002CDDDDA0005FF5009FFFFFF5005FF5003AAACFF5003FF70000005FF5001EFA0000005FF5000BFF1000005FF50" +
        "006FFB000005FF50002FFF910006FF500005FFFEBBDFFF5000007FFFFFFFFC100000006BEEC830000001357898764210001479BBCBBA85200136ABBA" +
        "AABBA7300159BB8655789730027BB95322356420038BA62100112210049B84100000100005AB73111111110005AB73124454431005AB732478998630" +
        "05AB73359BBCCA4005AB732478ABCB5005AB84223469CB50049B95211247BA50038BA6310137BA50027BB9532247BB500159BB965579CB500136ABBA" +
        "99ABCA40001479BBBBBBA730000135789987641000466777776654000056777776665400006777666665540000677665555443000077665444443300" +
        "007665433333320000765543333333000076554444444300007655444555540000765555566665000076655556777600007665555677770000766555" +
        "567787000077655556778700007766555667870000777665567787000067776666778700006777777778770000567777777776000045667777776500" +
        "0000049DEEB720000000BFFFFFFFE400000CFFFFFFFFFF40005FFFFFCEFFFF6000CFFFD4018FFF4003FFFF200003640007FFF800000000000AFFF500" +
        "000000000AFFF404555555100AFFF23FFFFFFF800AFFF35FFFFFFFA00AFFF52EFFFFFFA009FFF502556FFFA007FFF900002FFFA002FFFF40002FFFA0" +
        "00BFFFD5003FFFA0005FFFFFCCFFFFA0000CFFFFFFFFFFA00000CFFFFFFFFE400000059DEEDA520000124689A997531000248BCDDDDC96200148CEEE" +
        "EEEEC940027BEEDBABCDDA40039DEDA7668AA84004AEEB732245652005BEC8411112221006CEB7322233321006CEB7456777652006CEB768ABBBB940" +
        "06CEB779CDEEEC6006CEB768ABCEEC6006CEC855679CEC6005BED953237BEC6004ADEB74236BEC60039DEDA7568CEC60027BEEDBAACDEC600148CEEE" +
        "DDEEEB5000248BCDDDDCB84000124689AA9875200134567888775420024678999998753003679AAAAA99763003789AA9998876300489AA9887776530" +
        "048AA98766665420059A997665555420059A987666665420059A987777776530059A998788988730059A998889AA9840059A998889AAA850059AA988" +
        "89AAA950059AA98888AAA950048AAA9889AAA9500479AAA999AAA95003689AAAAAAAA840035799AAAAAA974002457899999986300134567788776520" +
        "0000005BED92000000004EFFFFFFA0000003FFD856AFFA00000BFA100003CA00004FD1000000000000AF70000000000001FF20000000000003FE0000" +
        "0000000005FC00000000000005FC00006888850005FC0003FFFFFF5005FD00008AAAEF5003FF10000000CF5000EF40000000CF5000AF90000000CF50" +
        "004FE1000000CF50000BFA100000CF500003FFE8557BFF5000004EFFFFFFFC100000005ADEB720000001356788753200001368999998741000368997" +
        "778885100158975333466410017985311112320002897310000111000299620000000000029851001111100003984112234432000398411357776510" +
        "0398411479AAA82003985213578AB920029952112357A920029962101126992002897310002599200179853111269920015898643347A92000368997" +
        "7789A8200024689999AA9610000135678876531000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "009E40000002DB1002FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFB000" +
        "0008FF5002FFEDDDDDDEFF5002FFFFFFFFFFFF5002FFFFFFFFFFFF5002FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFB0000008FF50" +
        "02FFB0000008FF5002FFB0000008FF5002FFB0000008FF5002FFA0000008FF50008C30000001CA00048752000035884005A963000137AA4005BA7310" +
        "0147AB5005BA73100147BB5005BA73100147BB5005BA73100148BB5005BB85222258BB5005BCA755557ACB5005CDCA9999ACDC5005CDDCBBBBCDDC50" +
        "05CDCA9999BCDC5005CCA765568ACB5005BB85322359BB5005BA74111148BB5005BA73100147BB5005BA73100147BB5005BA73100147BB5005BA7310" +
        "0147AB5005A9630001379A40048752000025783000000000000000000000000000000000000000000000000022222211111222226777654444567777" +
        "777776555566777777887766667788877888877777788887789988877888998778999888888999887899988888899988788988877788998778888777" +
        "777888877788766666678887777766555556777767776544445667772222221111122222000000000000000000000000000000000000000000000000" +
        "00CFE100000CFD1002FFF600003FFF5002FFF600003FFF5002FFF600003FFF5002FFF600003FFF5002FFF600003FFF5002FFF600003FFF5002FFF733" +
        "335FFF5002FFFFFFFFFFFF5002FFFFFFFFFFFF5002FFFFFFFFFFFF5002FFFCAAAABFFF5002FFF600003FFF5002FFF600003FFF5002FFF600003FFF50" +
        "02FFF600003FFF5002FFF600003FFF5002FFF600003FFF5002FFF500003FFF5000BDC100000BDC10037873100147873005BCA520026ACA4005CDB620" +
        "137BDB5005CDB620137BDB5005CDB620137BDB5005CDB620137BDB5005CDB742248CDB5005CEDA7667ADEC5005CEECBAABDEEC5005CEEEDDDEEFEC50" +
        "05CEEDDCCDEEEC5005CEDB9889BDEC5005CEC854469CEB5005CDB731138CDB5005CDB620137BDB5005CDB620137BDB5005CDB620137BDB5005CDB620" +
        "137BDB5005BCA520026ACA40038974100147973045544322223455446787654334577876788876444567888778987654456899877999876556789987" +
        "799998766789A99789AAA988889AAA9789ABAA9999AABA9789ABBBAAAABBBA9789BBBBBAABBBBA9889BBBBAAAABBBA9889ABAA9999AABA9789AAA988" +
        "889AAA9779AA98777789AA97799987655678998778998655557899877898765445678987788875444467888767876543345678766777643334567776" +
        "008C10000000990000CF50000002FF0000CF60000003FF0000CF60000003FF0000CF60000003FF0000CF60000003FF0000CF60000003FF0000CF6000" +
        "0003FF0000CFB8888889FF0000CFFFFFFFFFFF0000CFCAAAAAABFF0000CF60000003FF0000CF60000003FF0000CF60000003FF0000CF60000003FF00" +
        "00CF60000003FF0000CF60000003FF0000CF60000003FF0000CF50000002FF00007A0000000088000786410000147860089741000025897008984100" +
        "002589800898410000258980089842000025898008A8521001258A8008A8632222369A8008AA86555568AA8008BBA988889ABB8009BCBAA99AABCB80" +
        "08BBA998899BCB8008AA86555668AA8008A9632222469A8008A8521001258A8008984200002589800898410000258980089841000025898008984100" +
        "002589800897410000258970078631000014786000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000022222222222222227777666666667777777776666667777777777766667777777777776666777777777776666677777777776666" +
        "666677772222222222222222000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "0006EFFFFFFF9000000DFFFFFFFFF1000004CDEFFEDD70000000005FF80000000000005FF80000000000005FF80000000000005FF80000000000005F" +
        "F80000000000005FF80000000000005FF80000000000005FF80000000000005FF80000000000005FF80000000000005FF80000000000005FF8000000" +
        "0000005FF80000000000005FF8000000001BDDEFFEDDC300007FFFFFFFFFFA00002DFFFFFFFFE400000368888887300000049BCCCCB9400000038ACD" +
        "DCA84000000247ACCA7520000001248BB84210000000137AB83100000000027AB83100000000027AB83100000000027AB83100000000027AB8310000" +
        "0000027AB83100000000027AB83100000000027AB83100000000027AB83100000000137AB83100000001248BB8521000000257ACCA75300000059ACD" +
        "DCB950000005ABCCCCBA6000000478888887400000000477774000000000048888400000000005899850000000000488884000000000047887400000" +
        "000004777740000000000367763000000000036776300000000003666630000000000366663000000000036666300000000003666630000000000367" +
        "763000000000036776300000000004777740000000000478874000000000058888400000000005899850000000000588885000000000047777400000" +
        "000AFFFFFFFFC100003FFFFFFFFFF500003FFFFFFFFFF5000008DDEFFFDDA100000000AFFD000000000000AFFD000000000000AFFD000000000000AF" +
        "FD000000000000AFFD000000000000AFFD000000000000AFFD000000000000AFFD000000000000AFFD000000000000AFFD000000000000AFFD000000" +
        "000000AFFD000000003CDDEFFFDDD400008FFFFFFFFFFA00008FFFFFFFFFFA00004EFFFFFFFFF50000158899998851000028CDDDDDDB82000029CDEE" +
        "EEDC82000027ACDEEDBA6100001358BEEB853100000126ADDA521000000015ADD9410000000015ADD9410000000015ADD9410000000015ADD9410000" +
        "000015ADD9410000000015ADD9410000000015ADD9410000000025ADD9410000000136ADDA521000001468BEEB8541000028ACDEEDCA8200003ADEEE" +
        "EEDDA300003ACDDDDDDCA300002789999998620000006678776500000000789999870000000079ABBA970000000079ABBA970000000078ABBA860000" +
        "0000679AA975000000005789986500000000468998640000000046899764000000004689975400000000468997540000000046899764000000004689" +
        "9864000000005789987500000000689AA9760000000079ABBA870000000089ABBA980000000089ABBA9800000000789AA98700000000677887760000" +
        "0006EFFFFFFE5000000BFFFFFFFF90000000237FF63200000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005F" +
        "F30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF3000000" +
        "0000005FF30000000000005FF30000000001337FF6331000005FFFFFFFFFF400002DFFFFFFFFC1000002689999862000000279BCBB972000000157AB" +
        "B96510000001248AA8420000000013799621000000001269962000000000026996200000000002699620000000000269962000000000026996200000" +
        "0000026996200000000002699620000000000269962000000000026996200000000012699620000000001379962100000001348AA8431000000157AB" +
        "BA75100000028ABCBB982000000278999987200000000047740000000000004775000000000000477400000000000046640000000000004664000000" +
        "000000366400000000000035530000000000003553000000000000355300000000000035530000000000003553000000000000355300000000000035" +
        "530000000000003553000000000000466400000000000046640000000000004774000000000000477500000000000057750000000000004775000000" +
        "00000000005E70000000000000DFF0000000000000DFF0000000000000DFF0000000000000DFF0000000000000DFF0000000000000DFF00000000000" +
        "00DFF0000000000000DFF0000000000000DFF0000000000000DFF0000000000000DFF0000000000000DFF0000007FC0000DFF000000CFF0000DFF000" +
        "000CFF0000EFF0000008FF7002FFC0000001FFFFFFFF300000006FFFFFFA00000000019DDA200000000000001378700000000000249A800000000000" +
        "259B900000000000259B900000000000259B900000000000259B900000000000259B900000000000259B900000000000259B900000000000259B9000" +
        "00000000259B900000011100259B900000044310259B900000077521259B90000009A74225AB9000000AB95347AB90000009BB9779BA70000007ABBB" +
        "BBB9600000047ABBBB963000000246888753100000000000000000000000000000000000000000000000000044444555566677775555555566667777" +
        "555555566667777766666666677777776666666677777777666677777777777777777777777777777777777777777777777778888887777677778888" +
        "887777767777888887777766777777777777766667777777777766656667777777766655000000000000000000000000000000000000000000000000" +
        "0000000003CFB1000000000008FFF5000000000008FFF5000000000008FFF5000000000008FFF5000000000008FFF5000000000008FFF50000000000" +
        "08FFF5000000000008FFF5000000000008FFF5000000000008FFF5000000000008FFF5000008ED4008FFF500001FFFA008FFF500003FFFA008FFF500" +
        "002FFFE31CFFF500000DFFFFDFFFF2000007FFFFFFFFA0000000BFFFFFFD2000000006BEEC71000000000000248861000000000136BC920000000001" +
        "37CDB2000000000138CDB2000000000138CDB2000000000138CDB2000000000138CDB2000000000138CDB2000000000138CDB2000000000138CDB200" +
        "0001110138CDB2000002331138CDB2000016874238CDB2000029BA6348DDB200002BDC8559DDB200002BEEB89BEDA200002ADEDCCDED92000018CEEE" +
        "EEEB710000158BDDDDB84100000247899864200000000112467774000000012356787400000001235788850000000123578985000000012357898500" +
        "000001235789850000000123578985000000112357898500001112235789850000122234678985000023334567898500002445557899850000366667" +
        "889985000047788889998500004889999AA9850000589AAAAAA9840000489AABBAA9740000479AAAAA98740000367899998753000025677887654200" +
        "00000000002C300000000000007FA00000000000008FA00000000000008FA00000000000008FA00000000000008FA00000000000008FA00000000000" +
        "008FA00000000000008FA00000000000008FA00000000000008FA00000000000008FA00000000000008FA00000036100008FA000000AF500008FA000" +
        "000AF600008FA0000007FB0000CF50000001DFC65BFD000000006DFFFFD500000000018DD71000000000000012587000000000001369800000000000" +
        "136980000000000013699000000000001369900000000000136990000000000013699000000000001369900000000000136990000000000013699000" +
        "000000001369900000011000136990000002211013699000000453101369800000067421136980000008863124798000000798544589700000068988" +
        "889850000004699999863000000246888864200000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "06E800000004DA000AFE0000007FFF200AFE000003EFF9000AFE00007EFF80000AFE0003FFFA00000AFE005EFFA000000AFE04FFFB0000000AFE4EFF" +
        "900000000AFFFFFB000000000AFFFFC1000000000AFFFFF7000000000AFEAFFF300000000AFE08FFE60000000AFE00BFFF4000000AFE0008FFD50000" +
        "0AFE0000BFFF30000AFE00001AFFD3000AFE000000AFFE400AFE0000001DFFA007FA00000000BF70058741001247884006A95201247AB94006BA6212" +
        "57ABA73006BA63247ABA852006BA7446ABB8531006BB868ABA85210006CBA9ABA852100007CCCBCB8531000007CEEDCA5310000007DEEEC942100000" +
        "07CEEDCA5210000007CDCCCB8521000006BBA9ABA852100006BA868ABA85210006BA7457ABA7421006BA63247ABA852006BA621257ABA83006BA6211" +
        "257ABA5006A952001247AA60047741000124785006777777766667700677888777777770067888888887777007789999988887700788999999888760" +
        "07899AAAA99887600789AAAAAA9887600789AABBAA9887600789ABBBAA9876500799ABBBBA9876500789ABBBAA9876500789AABBAA9887600789AAAA" +
        "AA98876007899AAAA9988760078899999998876007889999988887700778888888887770067888888777777006777777777777700677777666666660" +
        "07EE7000003DFD100EFFE00003EFFF300EFFE0003DFFFF200EFFE001EFFFFD000EFFE03EFFFFB0000EFFE3DFFFFA20000EFFEEFFFFD000000EFFFFFF" +
        "FC0000000EFFFFFFB20000000EFFFFFE100000000EFFFFFF700000000EFFFFFFF70000000EFFEFFFFF6000000EFFE6FFFFF500000EFFE08FFFFF5000" +
        "0EFFE007FFFFF4000EFFE0008FFFFF400EFFE00008FFFFD00EFFE000009FFFE008FF80000009FF80058852101358873008CB8311369CCA5009DC9423" +
        "69CEDA5009DD95469CEEC84009DDA779DEEC952009DDBABCEEC8521009DEDCDEEC95210009DFEEEEC952100009DFFFED9621000009DFFFEC84100000" +
        "09DFFFED9520000009DFFFEEB741000009DEEDEEDB73100009DECBCDEDB7410009DDB89BDEDA741009DDA657BDEDA73009DD94348BDEDA6009DC9412" +
        "48BDEC8008CB8310248BDC800698621012479960256665545567776226787766678888722789988888999872279AA99999A9986228AABBAAAAA98751" +
        "28ABBBBBBA99764128ACCCCCBA97643128BCDDDCBA86532028BCDDDCB976431028BCDDDCB975421028BCDDDCB975421028BCDDDCBA86432028ACCDCC" +
        "BA87542128ABCCCBBA98753128ABBBBAAAA98751289AAAA9AAAA986227999988899A9972278988777899997226777655567888722566654444677772" +
        "02C200000001B80009F80000001DFA0009F9000001CFD20009F900001CFD200009F90001CFD2000009F9001BFD20000009F901BFE200000009F90EFE" +
        "0000000009FDCFE30000000009FFFF500000000009FFFFC10000000009F92FFB0000000009F906FF7000000009F9002EFB00000009F90003FFA00000" +
        "09F900006EF7000009F9000003EF900009F90000004FF70009F900000007FE4004E3000000005E400476410001357740059852001357884005985211" +
        "3579863005985212579864100598533479864210059964579874210005A976788742100005AA99997421000006BCCBA85210000006BDDCB831000000" +
        "06BCCBA84210000005ABA9996421000005AA878986321000059965579863210005985335898631000598522368986310059852114689852005985210" +
        "246898400598520012468850047741000124675004667777777666400467777777777640046777777777774004777888888777400477888888887740" +
        "047788888888775004778889988887500478889999888750057889999998885005788999999888500578899999988850057889999998885004788899" +
        "998887500477888998888750047788888888775004778888888877400477788888877740046777777777774004677777777776400466777777766640" +
        "03EB10000000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF5000" +
        "0000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF50000000000007FF500000000000" +
        "07FF50000000000007FF50000000000007FFEDDDDDDDDC3007FFFFFFFFFFFF8003EFFFFFFFFFFE40048852100000000005AA63100000000005AA7310" +
        "0000000005BB73100000000005BB73100000000005BB73100000000005BB73100000000005BB73100000000005BB73100000000005BB731000000000" +
        "05BB73100000000005BB73100000000005BB73100000000005BB73100000000005BB73100000000005BB85322222221005BCA8766666653005BCCBAA" +
        "AAAA9840049BCBBBBBBBB950036788888888874000000000000000000000000000000000777776665544332277777666554433327777776655544333" +
        "777777666554443377777777665554447777777766655544777777777666555577788877776666557778888777776666777888887777766677778888" +
        "877777667777888887777777677777887777777766777777777777776666777777777777566667777777777700000000000000000000000000000000" +
        "05FE8000000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100" +
        "000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF100000000000CFFF10000000000" +
        "0CFFF100000000000CFFFDDDDDDDDC500CFFFFFFFFFFFFD00CFFFFFFFFFFFFE005FFFFFFFFFFFE60037873100000000005BC95100000000005CDB520" +
        "0000000005CDB6200000000005CDB6200000000005CDB6200000000005CDB6200000000005CDB6200000000005CDB6200000000005CDB62000000000" +
        "05CDB6200000000005CDB6200000000005CDB6200000000005CDB6200000000005CDB7321111110005CEC9655555542005CEECBAAAAAA94005CEEEDD" +
        "DDDDDB5004ACDDDDDDDDCB50037899999999873004777653210000000478765321000000058887542100000005898754210000000589876421000000" +
        "058987642100000005898764210000000589876421000000058987642100000005898764210000000589876421100000058987642211110005899865" +
        "3221111005899876433322100589A987655544200589AA98777666300589AAA9988877400489AAAA9998885003689999999887500356788888777740" +
        "02C400000000000007FA00000000000007FB00000000000007FB00000000000007FB00000000000007FB00000000000007FB00000000000007FB0000" +
        "0000000007FB00000000000007FB00000000000007FB00000000000007FB00000000000007FB00000000000007FB00000000000007FB000000000000" +
        "07FB00000000000007FB00000000000007FB33333333310007FFFFFFFFFFFF5002CFFFFFFFFFFD300277420000000000028852000000000002995210" +
        "000000000299521000000000029952100000000002995210000000000299521000000000029952100000000002995210000000000299521000000000" +
        "029952100000000002995210000000000299521000000000029952100000000002995210000000000299632111111100029A854444444310029BA877" +
        "77777620028AAA9999999820016899988888872000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "08FFF600008FFE600FFFFF1003FFFFE00FFFFF6007FFFFE00FFFFFB00DFFFFE00FFCEFF00FFCEFE00FFC9FF23FF7EFE00FFC5FF79FF5EFE00FFC3FFC" +
        "DFF2EFE00FFC0DFFFFB0EFE00FFC08FFFF60EFE00FFC02FFFF10EFE00FFC00CFFB00EFE00FFC007FF500EFE00FFC001FE000EFE00FFC00011000EFE0" +
        "0FFC00000000EFE00FFC00000000EFE00FFC00000000EFE00FFC00000000EFE00AF6000000008F80168986311358986128CDC952258CDC9229DEEB74" +
        "36ADEDA32ADEEC8557BEEDA32ADEDCA779CDEDA329CDDCB99ACCDCA329CCBCCBBCCBCCA329BBABCDDDBAABA329BA9ACEEDB9ABA329BA89CEECA89BA3" +
        "29B977ADDB878BA329B9669CDA768BA329B9658BB9658BA329B9535786458BA329B9523453248BA329B9521221148BA329B9510000148BA329B95100" +
        "00148BA329A8410000147A9227863100001368720000000000000000000000000000000044554443344555446888877667788876789A99877899A987" +
        "89AAAA9999AAAA9789ABBAAAAAABBA9889ABBBBAABBBBA9889ABBBBBBBBBBA9779AABBBBBBBBAA87789AABBBBBBAA987789AAABBBBAAA98778999AAA" +
        "AAA998877788999999998876678888888888887667777777777777766777666666677776677665544556677600000000000000000000000000000000" +
        "0CFFFE3005EFFFA00FFFFFB00DFFFFF00FFFFFF23FFFFFF00FFFFFF55FFFFFF00FFFFFF8AFFFFFF00FFFFFFDEFFFFFF00FFFFFFFFFFFFFF00FFFFFFF" +
        "FFFFFFF00FFFCFFFFFFDFFF00FFF8FFFFFE9FFF00FFF7AFFFF89FFF00FFF76FFFF59FFF00FFF74FFFF39FFF00FFF70DFFB09FFF00FFF70277109FFF0" +
        "0FFF70000009FFF00FFF70000009FFF00FFF70000009FFF00FFF70000009FFF00DFF30000004FFC0169AA8533589A96228CDDC8558BDDD922ADFFDA7" +
        "7ADFFEA32AEFFEC99BEFFEB32AEFFFDCCDEFFEB32AEFFFEEEEFFFEB32AEFFFFFFFFFFEB32AEFFFFFFFFFFEB32AEEEEFFFFEEEEB32ADEDDEFFEDDEEB3" +
        "2ADDCCDEFDCCDEB32ADDBACEEDBBDEB32ADDA9BDDB9ACDB32ADDA789A879CDB32ADC95566558CDB32ADC94222238CDB32ADC94100138CDB32ADC9310" +
        "0138CDA329CB83100027BC9216985200002589723455554444555543679999888899997679ABBAA99AABBA867ABCCCBBBBCCCB978ACDDDCCCCDDDCA7" +
        "8ACDEDDDDDEEDCA88ACDEEEEEEEEDCA88ACDEEEEEEEEDCA88ACDEEEEEEEEDCA88ACDDEEEEEEDDBA88ABCDDDDDDDDCBA78ABCCCCCDCCCCB9779BBBBBB" +
        "BBBBBA9779ABBAAAAAABBA9779AA99888899AA9779999877778999977899876556789987788875444467888768876533345678764554332112344544" +
        "08FFD300004EFE600FFFFB0000CFFFE00FFFFF1002FFFFE00FFDFF5007FFEFE00FF8BF900AFAAFE00FF76FB00DF59FE00FF72FF12FE19FE00FF70BF8" +
        "AF909FE00FF706FFFF409FE00FF702FFFE109FE00FF700AFF9009FE00FF7006FF5009FE00FF7005FF4009FE00FF7000980009FE00FF7000000009FE0" +
        "0FF7000000009FE00FF7000000009FE00FF7000000009FE00FF7000000009FE00AE2000000003E80168875211247886128BCA7411369BB9229CDC953" +
        "258BDDA229CDCA74369CDCA329BBBA8547ABBBA329AAAA8668AAAA9229A9999889999A9229A879AAAA977992299768ACCB867992299757ACDB756992" +
        "2996569CCA6569922996458BB95459922996336997435992299632467533599229963224432259922996311121125992299631000012599229962100" +
        "000259922895210000025882277520000001477200000000000000000000000000000000000000000000000000000000000000006777765555667776" +
        "788887666677888778888877777888877899888778889987789998888889998778899998889998877889999999999887678899999999887767788899" +
        "998888766778888998888776667788888888777666777777777777660000000000000000000000000000000000000000000000000000000000000000" +
        "02DFA0000001BD1009FFF6000005FF7009FFFD000006FF7009FFFF700006FF7009FFFFE00006FF7009FFDFF30006FF7009FF9FFB0006FF7009FF5DFF" +
        "4006FF7009FF45FFD006FF7009FF40DFF106FF7009FF403FFA06FF7009FF400FFF26FF7009FF4007FFB6FF7009FF4001EFF9FF7009FF40006FFDFF70" +
        "09FF40000FFFFF7009FF40000AFFFF7009FF400002FFFF7009FF4000009FFF7004EC1000001DFE20037874200013663005BCA7310025995006CEC951" +
        "0126AA5006CEEB731126AB5006CEEC942126AB5006CEEDB63126AB5006CDCCC85236AB5006BCBBBA7337AB5006BB99AB9547AB5006BA879BA657AB50" +
        "06BA757BB878BB5006BA6459BA89BB5006BA6337ABAACB5006BA62259CCCCB5006BA62147BDDDC5006BA62125ADEEC5006BA621138BEEC5005BA6201" +
        "26ADEC5005A95200148BCB50036631000247973047777653333445534788876443445553589998654445666458AAA9765555666459AAA98765566764" +
        "58AAA98766667764589AA99876677774589AA99877777774489999988888877447899988888888744788888888898874477888888999987447777778" +
        "899A998447777677899AA9854677666779AAA98546766556789AAA8546665555689AAA85466654455789A98535554434567898743554433345678774" +
        "07EFF8000008FE500EFFFE10000FFFC00EFFFFA0001FFFC00EFFFFF2001FFFC00EFFFFF7001FFFC00EFFFFFE201FFFC00EFFFFFF801FFFC00EFFFFFF" +
        "D01FFFC00EFFEEFFF71FFFC00EFFE8FFFE1FFFC00EFFE3FFFF6FFFC00EFFE0AFFFCFFFC00EFFE02FFFFFFFC00EFFE00AFFFFFFC00EFFE004FFFFFFC0" +
        "0EFFE000AFFFFFC00EFFE0004FFFFFC00EFFE0000CFFFFC00EFFE00004FFFFC008FF800000BFFF70048996310136873005BDCA621159CB5006DFEC84" +
        "125ADC6006DFFEA6226ADC6006DFFEC8426ADC6006DFFFDA636ADC6006DFFFEC856ADC6006DFFFFDA77ADC6006DEEEEEC99BDC6006DEDCDEDBACEC60" +
        "06DDCABDEDCDEC6006DDB89CEEEEEC6006DDA77ADFFFFC6006DDA658CEFFFC6006DDA536AEFFFC6006DDA5248CEFFC6006DDA5226AEFFC6006CD9511" +
        "48DEEC6005BB841126ACDA5003775310136897303678765434456653479998755556776448ABBA876667887459BCCB987778887459BCCCA988889984" +
        "59BCDCBA9889998459BCDDCBA99AA98459BCDDCBBAAAAA8459BCDDCCBBBBBA8459BCCCCCCCCCBA8559ABCCCCCCCCCA9558ABBBBBCCDDCB9558AAAAAB" +
        "BCDDCB95489A999ABCDDCB9548999889ABCDCB95489988789BCCCB954888876789BCBA954788765678AABA8446776555578999743565543345677763" +
        "02DE400000007C1009FFE1000000AF7009FFF8000000BF7009FFFE100000BF7009FFFF800000BF7009FC8FC00000BF7009FA1EF40000BF7009FA05FD" +
        "0000BF7009FA00DF6000BF7009FA006FD000BF7009FA001FF500BF7009FA0009FA00BF7009FA0002FF20BF7009FA00007FB0BF7009FA00001EF5BF70" +
        "09FA00000AFEFF7009FA000003FFFF7009FA000000BFFF7009FA0000004FFF7004E700000008FE20048864100013563005BB96200013684006CDB831" +
        "0014795006CDCA521014895006CDCA631014895005BCBA852114895005AA99974124895005A978985224895005986689642489500598557985358950" +
        "05974368974589500597424798668A500597423689779A50059742247999AA500597411258ABBB500597410147BCDC500597410136ADDC5005974100" +
        "249CDC5005863100126ABB50036521000147984047776643323334434788765433334453588877644334455358888765444445535888876554445553" +
        "488887765545555347888776555555644778777665555664477777666666666446777666666666644666666666667764466666666677777446665556" +
        "667787743565555566788874355555556678888535554445567888853555444456788885355443344577888535443333456788853444333334567774" +
        "000004AEEA4000000000AFFFFFFA00000009FFFEEFFF8000003FFE3003EFF10000AFF400005FF80003FFD000000EFF1007FF70000009FF500BFF3000" +
        "0004FF900DFF00000002FFB00EFE00000001FFC00EFE00000001FFC00DFF10000002FFB00AFF30000005FF8006FF8000000AFF4002FFE000000FFE00" +
        "009FF400006FF700002FFE3004FFF1000008FFFEEFFF70000000AFFFFFF90000000004BFEA400000000236899763200001257ABBBBA742000137ABBA" +
        "ABBA7310036ABA7557ABA520048BB742257BB740069B94200259B95007AB83100138BA6008BA62000127AB8009B962000026AB8009B951000026AB80" +
        "09B951000026AB8008BA62000026AB8008BA72100127AB7007AB83100148BA60069B9520025AB950047BB742258BB740036ABA7558ABA5200137ABBA" +
        "ABBA731000257ABBBBA74100000236899763200000006677777620000000677777772000000077777777200000007666666720000000665555662000" +
        "000065444456200000006543334520000000543223452000000054322235200000005432223520000000543222352000000054322235200000005432" +
        "234520000000654333452000000065444456200000006655556620000000766666672000000077777777200000006777777720000000667777762000" +
        "000037DFFD7300000002DFFFFFFD2000000FFFFFFFFFE000008FFFFCCFFFF60002FFFFB11BFFFE1005FFFE0000FFFF5009FFF700008FFF700CFFF100" +
        "003FFFA00EFFC000000EFFC00EFFA000000BFFC00EFFA000000CFFC00DFFC000000EFFB00BFFF200004FFF9008FFF800009FFF6005FFFF0000FFFF50" +
        "01FFFFB11CFFFE00007FFFFCCFFFF600000EFFFFFFFFE0000002DFFFFFFD1000000037DFFD7200000013579AA975210001269BDDDDB952100259CEEE" +
        "EEEC9520037CEECBBDEEB73005ADEC9779DED95007BEDA5336ADEB6008CEC831138CEC7009DDA620027BED8009DD9510015ADD900ADD94100159DD90" +
        "0ADD94100159DD9009DD9510026ADD9009DEB620027BEC8008CEC831148CEC7007BEDA6336ADEB60059DEC977ADED940037CEECBBDEEB7300259CEEE" +
        "EEEC942001269BDDDDB952100013579AA975210002346778887653200346789AA998653004679AAAAAA9864005789AAAAAAA97500689AA99999A9860" +
        "078999877899987007999866568999700799875445789980089986543468998008998643346899800899864334689980089986543468998007999764" +
        "45789980079998766689997006899987789998700689AA9999AA986005789AAAAAAA9750045799AAAAA986400346789AA99865300234577887764320" +
        "0000029EE920000000007EFFFFE700000004FF9449FF3000002EF600007FD10000AF80000009F80001FF50000005FE0005FE10000002FF4006FC0000" +
        "0000EF5008FA00000000BF6009F900000000AF7009F900000000AF7007FA00000000CF5005FC00000000EF5004FF10000003FF4001EF50000005FD00" +
        "009F9000000AF700001DF600007FC1000003FF9339FE200000007EFFFFE600000000029EE91000000001357787532000001368999986410001268987" +
        "789863100148975434798520026985211247973003797310013698400488620000258940059851000014795005984100001479500597410000137950" +
        "059741000013795005984100001479500498520000147950048862000025894003797310013698300269852112489630014897543579852001268987" +
        "789863100013689999864100000135788753200000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "03EFFFFFEC83000007FFFFFFFFFE910007FFEDDDDFFFFB0007FF5000016DFF4007FF50000005FF7007FF50000005FF9007FF50000005FF8007FF5000" +
        "0008FF6007FF855555AFFF3007FFFFFFFFFFF70007FFFFFFFFFE500007FFA8887652000007FF50000000000007FF50000000000007FF500000000000" +
        "07FF50000000000007FF50000000000007FF50000000000007FF50000000000003EC100000000000046888888765311007ABBBBBBBA9642008BCCBAA" +
        "ABBBA74009CC9876789BB96009BB74322469BB7009BA63100137BB8009BA62100136AB9009BA63212358BB8009BB8654568ABA7009CCB98899BBB850" +
        "09CDDCBBBBBA852009CDCBAAA987421009CC98666543210009BA74222111000009BA63100000000009BA62100000000009BA62000000000008BA6200" +
        "0000000008A9520000000000068742000000000022222222222222216677777777776655677887777777766567888777667777767788776666677776" +
        "778776655566777777777655556677777777765555567777778776655566777777887766666777777888877766777776788988777777776678999887" +
        "777776657899988777766654789988776665554378888776655443327788776554433322777766543322221167776543322111112222211110000000" +
        "06FFFFFFFDA610000CFFFFFFFFFFE4000CFFFFFFFFFFFD100CFFFDDDEFFFFF700CFFF10000BFFFC00CFFF100004FFFE00CFFF100002FFFE00CFFF100" +
        "03CFFFD00CFFFAABCFFFFF900CFFFFFFFFFFFD200CFFFFFFFFFFD6000CFFFFFFFFD910000CFFF433310000000CFFF100000000000CFFF10000000000" +
        "0CFFF100000000000CFFF100000000000CFFF100000000000CFFF1000000000005FF800000000000047899998876321007ACDDDDDCBA742008CEEEDD" +
        "DEEDB84009DEDCBBBCDEDA6009CEB86568BDEC8009CDA632248CED9009CDA621137BED9009CEB74446ACEC9009DECA989ACEEB7009DEEDCCDDEED960" +
        "09DEEEEDDDDC963009DEECBBBAA8631009DEC9766654210009CDA6322211000009CDA5210000000009CDA5100000000009CDA5100000000009CD9510" +
        "0000000008BC841000000000058863100000000035678888777654323688999999887652489AAAAAAAA99863589AAAA99AAA997458AAAA9999AAA984" +
        "59AAA988899AAA9559AA9987889AAA9559AAA988889AAA9559AAA99889AAAA8559ABBAA999AAA98459ABBBAAAAAA987459ABBBAAA999875359ABBAA9" +
        "9887654259ABAA987766543259AA98765443322158998765433221105899865422111100589876432111000047887542110000004777653210000000" +
        "02CFFFFFEB71000007FFFFFFFFFE700007FC3333349FF70007FB00000005FE1007FB00000000CF6007FB000000009F8007FB00000000AF7007FB0000" +
        "0001EF4007FB0000127CFA0007FFFFFFFFFFE10007FFFFFFFFF9200007FD88887510000007FB00000000000007FB00000000000007FB000000000000" +
        "07FB00000000000007FB00000000000007FB00000000000007FA00000000000002D4000000000000047899888775421005ABAA999998752006AB9877" +
        "7889973006A964333457995006A85211112589600698420000147960069842000014796006A852111135896006A964333457995006BB977777899840" +
        "06BCBAA99998752006BBA9888765421006AA86554432110006A963222111000006A85210000000000698420000000000069841000000000006984100" +
        "000000000597410000000000047631000000000077777777777666667777777777777666777777777777766677777777777777667777777777777776" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777677777777777777667777777777777666777777777777666677777777777766667777777777666666" +
        "000006ADC61000000003EFFFFFF30000000EFFFFFFFE100000AFF9000AFFC00002FFE00000CFF00008FF6000002FF6000BFF2000000FFA000EFF0000" +
        "000DFD000FFF0000000BFF000FFF0001000AFF100FFF003FD00BFF000DFF108FFA0DFD000BFF301FFF6FFB0007FF8004FFFFF60002FFE0007FFFF100" +
        "008FFB201FFFE000001EFFFDFFFFF5000001BFFFFFEFFF60000007CFE818FFD00000000000007E60001246788753110001258ABBBB9742000158ABBA" +
        "ABBA6310037ABA8679BB9520049BA742358BA73005AB85211369B93005BB73100148BA4006BA62100137BA5006BA62111136AA5006BA63233236AB50" +
        "06BA63456547AB5006BB74589879BB5006BB7459BBABCA5005BB8547ACCDCA4004AB96458CDEC930038BB8668BEEC9300269BB99ACDEC94001369BBB" +
        "CCCDCB500023689AAA9ABB60000124566556895003566777776654200366777777766520046777777777653004777666667766300477655555666630" +
        "047665444566664004765443445666400476544344567640057665444566774005776555566777400577665667788740057776667888874005777777" +
        "899998500577788899AA9850047888899AAA985004778899AAAAA85004678899AAAA9850036678899AAA985003566778899987400244566777887640" +
        "000048CDD94000000007FFFFFFF90000006FFFFFFFFF800002EFFFFDFFFFF10005FFFF504FFFF60009FFF80004FFFC000CFFE10000EFFF100EFFB000" +
        "00BFFF400FFFA000007FFF500FFFA07A926FFF500FFFB0FFFC7FFF500EFFD0FFFFEFFF400CFFF2DFFFFFFF2009FFF83DFFFFFD0005FFFF36FFFFF800" +
        "01EFFFFFFFFFFA00004FFFFFFFFFFF600008FFFFFFFFFFD0000029EFFC6FFFE0000000000006FF80001357999864210001369CDDDCB84200036ADEEE" +
        "EEEB8410049CEDCBCDEDA62006BDEB978BDEC83008CEC85347BED95009DEB631259DDA6009DDA521147CDB6009DDA533347CDB700ADDA767778CDB70" +
        "09DDB89BBBBDEB7009DEB9ADEDDEEB6008DECAACEFFFEB6008CEDAABEFFFDA5006BEECBBDFFFD950049CEEDDEFFFDA60026ADEEEEEFFEC800137ACDD" +
        "DDDEED9000135799999BCC9000012344445688600235677887654320035689999987643004689AAAAA9875400579AAAAAAA987500689AA9999AA9750" +
        "0799A988889A986007999877789998700899987667899970089A9877789AA970089AA988899AA970089AAA999ABBB970089ABAAABBCCBA70079ABBBB" +
        "CCDCBA70079ABCCCDDDDCA70068ABCCDDDDDCA700689BBCDDDDDCA800578ABBCCDDCBA80045789ABBBBBB970024567899AAA98700123456677777760" +
        "0000049CB50000000001BFFFFFB30000000CFE756DFB0000008FB10002CF600002FF4000002FE00005FE0000000AF50006FB00000006F90008FA0000" +
        "0005FC000AFA00000005FE000AFA00000005FF0009FA00194005FF0008FB004FE305FD0005FD000BFC06FA0005FF1001CFBCF50001EF50001DFFD000" +
        "008FB00009FF90000019FCA89DFFD20000009FFFFFDCFD10000005AAA603EF600000000000003D400013467887642100012578999987520001479876" +
        "778974100269864334688520037974211247973004886310013688400598520000258840059841000014795005A841000013795005A8421111136950" +
        "05A842234324795005A852356545795005985335887799500498532489AAA9400489642369BCB8400379853358BDC8300258876569BDC83001368888" +
        "9ABCB9400013578999999950000134667655685000000000000000002222222222222222555666677777777656666677777777766666677777887777" +
        "666777778888877766677788888888776777778888888877677778888888887777778888888888877777888889888887777888888888888777788888" +
        "888888877777888888888877777788888888887777777888888887776777777788877776667777777777777622222222222222220000000000000000" +
        "01CFFFFFD940000005FFFFFFFFFC400005FFEDDDFFFFF30005FF600016BFFC0005FF6000000DFF0005FF6000000DFE0005FF6000001FFC0005FF6000" +
        "007FF80005FF955557FFE20005FFFFFFFFFF500005FFFFFFFFF8000005FFB88FFF60000005FF6009FFC0000005FF6001DFF6000005FF60004FFC0000" +
        "05FF60000BFF600005FF600004FFE30005FF600000AFFB0005FF6000002DFF2001DE30000005FA000157888877542100028ACBBBBAA86310029CCBAA" +
        "ABBB961002ACB97778ABB82002AB9633347ABA2002AB95211259BA3002AB85200259BA2002AB9632247AB92002ACA754579BA72002ACCA989ABB9610" +
        "02ADDCCCDDCA631002ACDCBCDDB8420002ACB98ACCA7310002AB96469BB8410002AB95347AB9520002AB852259BB741002AB852136ABA610029B8410" +
        "148BB820029A84101369A920027863100136882025677777776665512677777777776652267888877777776227788877777777722778877766777772" +
        "277877766667777227777766666777722777776666677772278887766667777227888777777777722788888888888772288999988888877228899999" +
        "999887622889999999998762288999999999876228899999999887622788888888888762278888888888776227777777777777622777766666777762" +
        "02EFFFFFEB74000005FFFFFFFFFF800005FFFFFFFFFFF80005FFFEDEFFFFFD0005FFF50017EFFF0005FFF50000AFFF0005FFF50000DFFE0005FFF500" +
        "19FFFB0005FFFCAADFFFFA0005FFFFFFFFFFF30005FFFFFFFFFF600005FFFFFFFFF4000005FFF75FFFF9000005FFF50DFFFF100005FFF503FFFFA000" +
        "05FFF5008FFFF50005FFF5002EFFFC0005FFF50006FFFF2005FFF50000CFFF3001EFC100004EFC000158999988653100028CDDDDDCB9741002ADEEED" +
        "EEEDA71002AEEDBBCDEED92002ADDA7679CEEB3002ADC95336ADEB3002ADC842259DEB3002ADDA6458BDDA2002AEEC999BDED92002AEEEDCDEEEB710" +
        "02AEFEEEFFEC841002AEEDDEEFDB620002AEDCABDEDA520002ADD978BEEB730002ADC9569DED951002ADC8447BEEB71002ADC84248CEDA2002ADC831" +
        "26ADDB30029CB731148BCA300168752002579720045678887776543006789999998875400689AAAAAAA98750079AAAAAAAAA9860079AAAA99AAAA870" +
        "079AAA9999AAA970079AA99899AAA970089AAA9999AAA970089ABAAAAAAAA870089BBBBBBBBA9860089BCCCCCCBA9750089BCCCCCCBA9750089BCCCC" +
        "CCBA8650089ABBBBBBBA8640079AAAAABBBA8650079AAA99AAAA97500789998889AA9760078988777899986006787766678887600567665455677760" +
        "009FFFFFD940000000FFFFFFFFFA300000FF833348EFF30000FF6000001BFB0000FF60000008FF0000FF60000008FD0000FF6000000AFB0000FF6000" +
        "000BF60000FF6000039FC10000FFFFFFFFFE200000FFFFFFFFD3000000FFB88AFF20000000FF6000EF90000000FF60007FF3000000FF60001DF70000" +
        "00FF600002FF300000FF6000008FC00000FF6000001EF50000FF60000008FE00009C10000001CA000168998887653100028ABAA999975310029BA877" +
        "88998610029A8543446898200299731112379920029962100126992002996210013698200299731112479820029A854334688610029BA87778997510" +
        "02ACCAAABBA85200029BB99ABBA63100029A9768AA952100029A74358995210002997323589741000299621146985200029962112579741002996210" +
        "135886100289621001478720027752000125772000000000000000000000000000000000000000000000000000000000000000007777776666666666" +
        "777777776666666677777777777666667778877777777777778888887777777777888888887777777788888888877777778888888887777777788888" +
        "887777777777777777777777777777777777777777777777777777760000000000000000000000000000000000000000000000000000000000000000" +
        "0000049CDA5000000003EFFFFFFE4000001CFFFBAEFFC100007FFA2001BFF60000BFF300003FF70000BFF30000046000007FF80000000000001EFFB1" +
        "000000000006FFFEA300000000004EFFFF9000000000005EFFFE5000000000017CFFC10000000000009FF70000030000003FFB0000AF9000001FFD00" +
        "00DFF200004FFA00009FF91002CFF500002DFFEBBFFFB0000005FFFFFFFD2000000006BDD9300000001246888764210000269BBBBBA852000049BBAA" +
        "ABBA7300005BB97557AB9400006BA74224688300006BA73111355200005BB85311122100004ABB86321100000037ABBA86421000001479BBB9752100" +
        "00013579BBB95200000012358ABB83000011111246ABA40000243100148BB50000475310137AB500005A9632248BB500005BB96567ABA4000049BBAA" +
        "ABBA730000269BBBBBA84100001357888764200000266777776652000026777777765200002777777666520000278776665551000027877655544100" +
        "002777765544410000277776554431000026777665543100002667776655410000156677776551000014556677766200001344566777620000133445" +
        "667772000013344556777200001444455677720000155555567772000025666666777200002666667777720000267777777762000026677777766200" +
        "000037BDDC8300000006FFFFFFFF7000005FFFFFFFFFF40000DFFFFDDFFFFB0000FFFF7106FFFD0000FFFE0000CFFB0000FFFF800003200000BFFFFC" +
        "70000000003EFFFFFD5100000005FFFFFFFB100000001AFFFFFFD10000000049FFFFFA00000000003BFFFF0000AED20000EFFF0000FFFA0000DFFF00" +
        "00FFFF5017FFFF0000EFFFFCDFFFFB00007FFFFFFFFFF3000008FFFFFFFE5000000048CDDB62000000135799997531000036ACDDDDCA7300006ADEEE" +
        "EEEDA600008CEECBBCEEC800009DEC9668BDC80000ADEC74348A9600009DEC9633455300008CEEC975432100005ADEEDB97421000036ACEEEDB85200" +
        "001358ACDEEC950000123468ADEEB700003443346ADED9000068752237BEDA00009CB84236BEDA0000ADDB8669CED900009DEDCABCEEC800007BDEEE" +
        "EEEDA6000047ACDDDDB9630000136899997531000035678888765300005689999998650000689AAAAA9975000079AAAAAA9976000079ABAA99887500" +
        "0089BBA9988765000079ABA9887654000078AAAA9876540000679AAA99875400005689AAAA98640000457899AAA97500004567899AAA860000456678" +
        "9AAA97000056677789AA98000067788889AA9800006899999AAA980000689AAAAAAA970000689AAAAAA9860000568999999865000045678887765300" +
        "0000038CC94000000001BFFFFFFC1000000BFFA658EFB000006FC200002DF50000AFA0000008F40000AFA00000001000005FD10000000000000BFC50" +
        "000000000000AFFC50000000000019FFFB30000000000007EFF900000000000019FF800000000000003EF50000000000000AFA00006A2000000AFA00" +
        "00AF9000000AFA00008FC100003DF400001DFE855AFF90000002DFFFFFFA0000000005ADC82000000013578888753100002589999998520000389876" +
        "678973000049964334677300005985211134520000598520001221000049964210011000003799753210000000257998653210000013578997642000" +
        "000123578997520000001124579973000000001125798400001211001258950000243100014795000047531112489500004886433469940000489876" +
        "678973000026899999985200001357898865210000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "08FFFFFFFFFFFF700EFFFFFFFFFFFFD007DDDDEFFEDDDD600000005FF50000000000005FF50000000000005FF50000000000005FF50000000000005F" +
        "F50000000000005FF50000000000005FF50000000000005FF50000000000005FF50000000000005FF50000000000005FF50000000000005FF5000000" +
        "0000005FF50000000000005FF50000000000005FF50000000000005FF50000000000003EE20000000478888888888740059BBBCCCCBBB950058AABCD" +
        "DCBA9840035678ADCA8765300122358BB85322100000137BB73100000000137BB73100000000137BB73100000000137BB73100000000137BB7310000" +
        "0000137BB73100000000137BB73100000000137BB73100000000137BB73100000000137BB73100000000137BB73100000000137BB73100000000137B" +
        "B73100000000137AA6210000000012588520000000000000000000000000000000000000444555555555444477777888888777767778888888888777" +
        "778888999888887777888999999888777788899999988877778889999998887767788899998887766677888998887766667788888888776556677888" +
        "888776655566778888776655456677777777665445566777777665444456677777665544233444455444433200000000000000000000000000000000" +
        "2CFFFFFFFFFFFFB15FFFFFFFFFFFFFF45FFFFFFFFFFFFFF30ADDDDFFFFDDDD80000000FFFD000000000000FFFD000000000000FFFD000000000000FF" +
        "FD000000000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD000000" +
        "000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD0000000000008FF6000000057899999999885008BCDDDDDDDDDC8008CDDEEE" +
        "EEEDDC90069ABCDEEDCBA970034558BDEC865430011125ADDA6321100000149DDA5100000000149DDA5100000000149DDA5100000000149DDA510000" +
        "0000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149C" +
        "DA5100000000138BC941000000000268863100000066777777776600007889999998870000789AABBBA9870000789ABBBBA9870000678ABBBBA97600" +
        "005679AABA9765000034689AA98654000024578999754300002346899875320000124689986432000012468998643100001246899864310000124689" +
        "986431000012468998643100001246899864310000124688986431000012467898643100001246788864310000124578876421000012356776542100" +
        "08FFFFFFFFFFFE600CFFFFFFFFFFFFA00033337FF63332000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005F" +
        "F30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF30000000000005FF3000000" +
        "0000005FF30000000000005FF30000000000005FF30000000000004FF30000000000001CB0000000027889999999872002899ABBBBAA9820026778AB" +
        "CA8876200133357AA8543310001113699631110000001258963100000000025896310000000002589631000000000258963100000000025896310000" +
        "000002589631000000000258963100000000025896310000000002589631000000000258963100000000025896310000000002589631000000000258" +
        "963100000000025886210000000002477521000000000000000000000000000000000000222222222222222267777777777777767777777777777777" +
        "777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777777" +
        "777777777777777777777777777777777777777667777777777777666667777777777666222222222222222200000000000000000000000000000000" +
        "03DA00000001BD2009FF40000005FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF4000" +
        "0006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7009FF40000006FF7008FF50000007FF60" +
        "05FF9000000AFF3001DFE610037FFD00008FFFFEEFFFF700000AFFFFFFFF9000000029DFFC920000068741000024786008A9520000259A8009B95200" +
        "0025AB9009BA52000026AB9009BA52000026AB9009BA52000026AB9009BA52000026AB9009BA52000026AB9009BA52000026AB9009BA52000026AB90" +
        "09BA52000026AB9009BA52000026AB9009BA52000026AB9009BA62100136AB8008BA63100137BB8007BB85322369BA70069BB976679BB950047ABBA9" +
        "9ABBA7400247ABBBBBBA7420012467899876321000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000077655444444566777766554444556677777666555566677777776666666777777777776666777776677777777777777666777777" +
        "777777665666777777776665000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "07EE70000008FE500EFFE000000FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000" +
        "001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFE000001FFFC00EFFF000001FFFC00EFFF500006FFFC0" +
        "0BFFFC1003DFFFA005FFFFFEFFFFFF3000BFFFFFFFFFFA00002DFFFFFFFFC10000016BEFFEA61000058853100136885008BC84100148CB8009CD9510" +
        "0159DC9009DD9510015ADC9009DD9510015ADC9009DD9510015ADC9009DD9510015ADC9009DD9510015ADC9009DD9510015ADC9009DD9510015ADC90" +
        "09DD9510015ADC9009DD9510015ADC9009DD9510025ADC9009DDA620126AEC9009CEB742248CEC8008CEDA7557BDEC8006ADEDBABCDEDA60048CEEED" +
        "DEEEB8400258BDDDDDCB85200124789AA987421047776432235677744787653233568874588875433457888558987543345789855898754334678985" +
        "589875433467898558987543346789855898754334678985589875433467898558987643346789855898764334678985589976444568998558998755" +
        "567899855899986667899985489A99877899A984479AAA9999AAA9744789AAAAAAAA986336799AAAAAA9876324678899998875422345677777765431" +
        "03D6000000007C2009FA00000000AF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA0000" +
        "0000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7009FA00000000BF7008FA00000000BF60" +
        "05FC00000001EF3000DF7000000AFB00004FFA5336CFF3000006EFFFFFFE5000000018CEEC7100000786310000136870089731000014798008974100" +
        "001489800897410000148980089741000014898008974100001489800897410000148980089741000014898008974100001489800897410000148980" +
        "089741000014898008974100001489800897410000148980089841000014898008984100002589700799631111369960058985433468985003689887" +
        "889986300246899999986320012357888875321000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "08E5000000006E600FFE00000001FFD00BFF30000005FFA007FF80000009FF5002FFA000000AFF1000DFC000000DFB0000AFF100003FFA00009FF600" +
        "008FF700004FFB0000DFF200000EFE0002FFD000000AFF5007FF80000006FF900BFF40000001FFE00EFD00000000CFF01FF9000000007FF35FF60000" +
        "00005FF89FF4000000003FFCDFE1000000000DFFFFB00000000009FFFF700000000002EFFD100000078631000013787009A8420000259A8009BA6210" +
        "0136AB8008BA73100137BB8007AB84100258BA60069B95200269B950058BA630137AB840037AB841148BA730036AB941259BA6200259B95226AB9520" +
        "0148BA7337BB73100137BB8458BA63100126AB9669BA520000259BA88AB9410000148BBAABB8310000137BCCCCA7310000026ACDDC952000000259CE" +
        "EC841000000147BDDA731000000125799742000007777666666677700777777666777770077777777777777007777777777777700777788888877770" +
        "0778888888888770077888888888877007788899998887700778899999988870077899999999887007789999999988700778999AAA998770067899AA" +
        "AA9987700678999AAA9987700678899AA999876006678999999887600567889999887660056778888887765005567788887765500455667777766540" +
        "0BFE20000003EFA00FFF9000000BFFF00FFFE000001FFFF00FFFF300005FFFE00BFFF80000AFFF9006FFFA0000BFFF5002FFFD0000EFFF1000CFFF30" +
        "05FFFA00007FFF800AFFF600005FFFC00EFFF500002FFFF34FFFF100000DFFF55FFFB0000008FFF79FFF60000003FFFCEFFF10000000FFFFFFFF0000" +
        "0000EFFFFFFB000000008FFFFFF6000000004FFFFFF2000000000EFFFFC00000000007FFFF600000068852000025886009CB84100148BC8009DDA510" +
        "026ADD9009DEB620027BEC8008CEC831138CEC7006BED94114ADEB6005ADDA5226BDD950049DEB6337CED840037CEC8448DEC730026BED966ADDB520" +
        "015ADEB88BED94100149DECAACEC83100137CEDCCDEB72000026BEEDEEDA62000015ADFFFED9410000148CFFFEC8310000037BEFFEB6200000026AEF" +
        "FDA52000000148CDDB741000000125899742000027776532235677722888754334678872289986544568998228998755567899822899976556899982" +
        "2799987667899872278999877899987226899988889998622689AA9999AA97621579AAAAAAAA97511579AABBBBAA96511468ABBCCBBA86411468ABCC" +
        "CCB9854114589BCDDCB9753113579BCDDCB9743113479ACDDCA8642102468ABCCB986320023579ABBA97532001246799997642100123567887643210" +
        "08C1000000002D600DF8000000009FB009FA00000000AF7005FC00000000EF3001EF20000004FD0000BF60000008F900007FA000000CF500001FF100" +
        "002FE000000DF400006FB0000009F80000AF70000005FE0001FF50000004FF3005FF20000001FF7008FD00000000AFA00AF8000000006FB00DF40000" +
        "00002FF12FE0000000000CF68FA00000000008FEEF600000000004FFFF200000000000BFF90000000575310000125750059741000013686005985200" +
        "001479600598631000258950048973100136994003897420014798400379852102489830026996311258973002589741136996200148984214798520" +
        "01379863258974100136897436896310002589754798521000147986689842000013699779974100001269A999963100000258ABBA852100000137AC" +
        "CA7410000001259BB9631000000024688742100047777777777776624777777777777762477777777777777247777777777777725777777777777772" +
        "577777777777777257777777777777725777788888777772577778888877777257777888888777725777788888877772577778888887777257777888" +
        "887777725777788888777772577777888777777247777777777777724777777777777762467777777777766246677777777776624666777777776662" +
        "08E4000000006E600FFC00000000EFE00FFD00000000FFD00EFE00000001FFC00EFE00000002FFC00DFF00000003FFB00CFF10000004FFA00BFF2005" +
        "5005FF9009FF407FF406FF7008FF50CFFA07FF6007FF72FFFF19FF5005FF87FFFF5AFF3003FFACFFFFBAFF1001FFCFFDEFFBFE0000EFFFF89FFFFC00" +
        "00CFFFF34FFFFA0000AFFFF00FFFFA0000AFFFD00EFFF800008FFF8009FFF500002DFD2001AFC100078621000013687009A8310000148A900AB84100" +
        "00149BA00AB9410000249BA00AB9410000259BA00AB9421111259B900AB952222235AB9009B964566546AB9009BA657AA757AB9009BA779CC978BB80" +
        "08BB99BEEB99BB8008BBABCEECBBCB7008BCCCDEEDCCCB7007BDDDDDDDDEDA6006ADEEDBBDEEDA6006ADEEB99CEED950059DEDA77ADEC840048CEC95" +
        "59CDB740037ACA7337ABA630024787422478742000000000000000002222211111122222777665444456677777776655556677777787777667777877" +
        "78888888888888877889999999999887789AAAAAAAAAA987899AABBBBBBBA98789ABBBBCCBBBBA9889ABBCCCCCCBBA9889ABCCCCCCCBBA9889ABBCCC" +
        "CCCBBA9779ABBBBBBBBBBA8778AABBBBBBBBA987789AAAAA9AAA99866789998888998876567888777778776522222222222222210000000000000000" +
        "0CFE20000003EEA00FFF70000009FFF00FFF8000000AFFF00FFF9000000BFFF00FFFA000000CFFF00FFFB000000DFFF00FFFC000000EFFF00FFFD08C" +
        "C60FFFF00FFFE4FFFF4FFFD00EFFF6FFFF8FFFC00DFFFCFFFFDFFFB00BFFFEFFFFFFFF9009FFFFFFFFFFFF7007FFFFFFFFFFFF5005FFFFFFFFFFFF30" +
        "02FFFFFEFFFFFF0000EFFFF9AFFFFC0000BFFFF56FFFF900008FFFF34FFFF600004FFFA007EFE300068852100025786008BB83100127BC9009DD9410" +
        "0138CDA009DD94100138CDA009DD94100139DDA009DDA5211249DDA009DDA644445ADDA009DDB878877BDD9008CECABCCBACED9008CEDCDEEDDDED90" +
        "07CEEEEFFEEEEC8007CEFFFFFFFFEC8006BEFFFFFFFFEC7006BEFFFFFFFFEB7005AEFFFFFFFFEB60059DFFEDDEFFEA50049DFFDCBDEFD940038CEECA" +
        "9BEEC840026ACCA769BCA730014688643578742067765432234567767887653223467887789875433456898779987654456789987999876556789998" +
        "79A9987767899A9879AAA998889AAA9879ABBBAAAABBBA9879ACCCCCCCCCCB9879BCDDDDDDDDCB9879BCDEEEEEEDDB9879BDDEEEEEEEDB9779BDEEEF" +
        "FEEEDB9779BDEEEEEEEEDB9768ACDEEEEEEDDB9768ACDDDDDDDDCA86579BCCDDDDCCBA86578ABBBBBBBBA87546799999999997643567888777887654" +
        "08C1000000002C600FF7000000009FD00FF800000000AFD00EF900000000AFC00DF900000000AFB00CFA00000000AFB00BFA00000000AFA00AFA0001" +
        "0000AF8009FA003EE200BF7008FA006FF500CF6006FB00BFF900EF4005FC01FFFE00FF3003FE05FFFF32FF1001FF3AF8AF95FD0000EFBEF24FEDFC00" +
        "00BFFFD00DFFF900008FFFA00AFFF600006FFF8009FFF300002FFF3003FFF100000AFA00007E90002874100000024872298520000002598229952000" +
        "001259922995200000126992299520000013699229962111111369822896322332236982289633455434798228964469864479822897569BB8657982" +
        "289867ACCA768972279878BDDB88997227999ABCCBA9A97117ABBBBBBBBBB96116ACCCA99ACCC96116ADDC9779CDD951159DDB8558BDD841148CCA64" +
        "46ACC741137BB852258AA630025886311368742000000000000000000000000000000000000000000000000000000000000000006777777777777776" +
        "67788888888888777788899999998877788999999999988778999AAAAAA999877899AAAAAAAA99877899AAAAAAAA99877899AAAAAAAA998778999999" +
        "999999877889999889999887678888888888887667777777777777760000000000000000000000000000000000000000000000000000000000000000" +
        "00BD30000004EA0003FFB000000CFF1000CFF300005FFB00006FFD1002EFF500000CFF7009FFB0000004EFE01FFE200000008FF79FF7000000003FFE" +
        "FFE30000000008FFFF700000000001EFFD000000000003FFFF20000000000BFFFFA0000000004FFDEFF300000000AFF46FF900000005FFD00EFF4000" +
        "001EFF5007FFD000007FFB0001DFF60000DFF200003FFC0005FFA000000BFF3002ED20000004ED100278631001368720029A85211269A820028BA841" +
        "248BB820016AB95336ABA6100149BB8558BB94100026ABA99AB9620000158BCBBCB84100001369CDDC963100000248CEEB741000000137BEEB631000" +
        "000147CEEB731000000259CEDC85200000147ACCCCA7410000269BBAABB952000138BB9779BA7300015ABA6447ABA510027BB842259BB720029BA731" +
        "137AB920029A74100258A920027752100135772000777777777777000077788888877700007788888888770000788899998877000078899999987700" +
        "0078899AA9987600006789AAAA987600006789AAAA987600006789AAAA987600006789AAAA987600006789AAAA987600006789AAAA987600006789AA" +
        "AA987600006789AAA99876000078899AA998770000788999999877000078889999887700007788888888770000777777777777000077777777777700" +
        "03DFB100002CFC2008FFF90000AFFF6008FFFD1003EFFF6002FFFF900BFFFE10006FFFF45FFFF600002EFFFBCFFFD1000006FFFFFFFF50000000EFFF" +
        "FFFD000000005FFFFFF3000000000AFFFF90000000000DFFFFC0000000008FFFFFF600000000FFFFFFFE00000008FFFFFFFF7000002FFFF9AFFFE200" +
        "008FFFF34FFFF70003FFFF7009FFFF200AFFFD1001EFFF800AFFF800009FFF9006FFB100002DFE40037874200257873005ACB731148BCA4005BDDA63" +
        "36ADDA5004ADEC8559CED940037CEDB99BEEB7200259DEDCCDED95100137CEEEEEEB731000259DEFFEC9410000137BEFFEB7210000026AEFFDA52000" +
        "00026AEFFDA6200000138CEFFEB731000025ADEFFEDA52000148CEEDDEEC7310026ADEDBBDEDA520038CEDA88BDEC73004ADEC8458CEDA4005BEDA52" +
        "36ADEB5005BCB731137BCB5004897410024798400477766556677740058888777788884005899998899998500489AAAAAAAA98400479ABBBBBBA9740" +
        "0478ABCCCCBA87300368ABCCCCB9863003579BCDDCB9752002579BCDDCB9652002569BCDDCB8642002569BCDDCB8642002579BCDDCB9752003679BCD" +
        "DCB975300368ABCCCCBA86300479ABCCCCBA87400489ABBBBBBA97400589AAAAAAAA9850058999888899985005888877778888500477765556677740" +
        "00A900000000A80001EF40000006FD00006FD000001EF500000DF600008FB0000005FE3004FF40000000DF900AFC000000005FE12EF30000000009FD" +
        "EF700000000001EFFD1000000000008FF7000000000000AFF9000000000004FFFF20000000000CFABFA0000000008FC00DF600000001EF8009FE1000" +
        "0006FD1002EF6000001EF500006FD000008FB000000DF60002EE30000004FE1002D900000000AD100078631001357700007985211247880000689631" +
        "13689700004798523479850000368974568873000024799778985200001258AAAA963100000147ACBA852000000126ACCA7310000000259CCA631000" +
        "0001269CCA731000000137ACCA842000001258AAAA863100001369988997420000258985579863000047986335897400005897422468860000798531" +
        "125897000088741001368800007752100124770000267777777762000027777777777200002777777777720000277788887772000027788888877200" +
        "002778888887720000278888888772000027888888877200002788888888720000278888888872000027888888887200002788888888720000278888" +
        "888772000027888888877200002778888887720000277888888772000027778888777200002777788777720000277777777772000027777777776200" +
        "00CD20000003DB0004FFA000000BFF2000DFE200003FFC00007FFB0000CFF600001EFF5006FFD0000006FFD00EFF40000000BFF45FF9000000004FFC" +
        "DFF4000000000CFFFFA00000000004FFFF3000000000008FF70000000000005FF30000000000005FF30000000000005FF30000000000005FF3000000" +
        "0000005FF30000000000005FF30000000000005FF30000000000005FF30000000000001FE00000000288631001258820029A85210147AA30028BA742" +
        "136AB920017AB963258BA8200158BA8547AB96100137BB9779BB741000249BBBABB9520000137ACCCCB83100000159CDDC962100000137ADDB841000" +
        "0000258CC95200000000147BB84200000000137AB84100000000137AB84100000000137AB84100000000137AB84100000000137AB84100000000137A" +
        "B841000000001369A7310000000012588631000047777777777777725777777777777772577888888888877257788888888887724788899999888772" +
        "477889999998877247788999999887724778899999988772467889999998876246788999999877623667889999887662366788999988765235677888" +
        "888776513566788888776551355677888877655134566788877665412456677777766541245567777766554124556677776654412445667777665441" +
        "04DFA100001BFD3009FFF800009FFF8009FFFD0001DFFF7003FFFF7008FFFF20008FFFF33FFFF600002FFFF89FFFE2000009FFFFFFFF70000000FFFF" +
        "FFFF000000008FFFFFF7000000001EFFFFD00000000005FFFF500000000001FFFE100000000000FFFD000000000000FFFD000000000000FFFD000000" +
        "000000FFFD000000000000FFFD000000000000FFFD000000000000FFFD0000000000009FF7000000027874200147882002ACB731137BCB2002BDDA63" +
        "259DEB3002ADEC8547BEDB20017CEDB87ADEC820015ADEDBBCEDA6100037BEEDDEEC840000259DEFFEDA620000136BEFFEC74100000249DEFDA52000" +
        "000127BEEC831000000015ADEB6200000000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149DDA5100000000149C" +
        "DA5100000000148BC941000000000268863100000477765555677740058887766788885005899988889998500589AA999AAA98500489AABBBBAA9840" +
        "0479ABBCCBBA97400368ABCCCCBA863003589BCCCCB9753002479ABCCBA8742002468ABCBB986420013579ABBA975310013568AAA98642100124689A" +
        "A97642100124579998754210012357899875321001235789986531100113568988653100011356788764310000134678876431000013456776542100" +
        "009A000000007B0000BF70000004FE00008FE000000DFA00001EF400001FF3000005FD1000BF80000000BF8004FE100000003FE10CF6000000000CFC" +
        "AFE00000000004FFFF800000000000BFFE1000000000001FF50000000000000FF30000000000000FF30000000000000FF30000000000000FF3000000" +
        "0000000FF30000000000000FF30000000000000FF30000000000000FF300000000000009C00000000087521001257700008864100247980000798521" +
        "135897000068974224798500004798644689740000358976689852000024799999974100001258ABBA852100000147ACCA7310000001259BB9521000" +
        "0000248AA74100000000137A973100000000137996310000000013699631000000001369963100000000136996310000000013699631000000001369" +
        "963100000000136986310000000012577521000007777777777777700777777777777770077777777777777007777777777777700777777777777770" +
        "077777777777777007777777777777700777777777777770077777777777777007777777777777700777777777777770077777777777777007777777" +
        "777777700777777777777770077777777777777006777777777777700667777777777770066777777777776006667777777776600666777777777660" +
        "01CFFFFFFFFFA00005FFFFFFFFFFF30000ADDDDDDEFFC000000000000BFF6000000000004FFE100000000000CFF6000000000005FFD000000000001E" +
        "FF3000000000005FFB000000000000BFF5000000000006FFC000000000001DFF5000000000007FFC000000000002FFF2000000000006FFB000000000" +
        "001CFF5000000000008FFB000000000001EFFEDDDDDDD90005FFFFFFFFFFFF4002CFFFFFFFFFFB1000378888888752000059BBBBCCCB83000048AAAA" +
        "BCDC9300002455669BDB7300001122248BCA5200000001248BB8310000000136AB95200000000148BB7310000000136AB95200000000148BA7310000" +
        "000136AB95200000000148BB7310000000136AB95200000000148BA7310000000026AB95210000000038CB8532222100005ADC9765555200006BDDBA" +
        "AAAA8400005BCCCBBBBBA500003788888888740000035667777640000003567788874000000356778998500000024567888850000002345678874000" +
        "000123567887400000012356777630000001235677653000000124567665200000013466765420000002456766431000000245676543100000035677" +
        "653210000003677765321000000477876543100000047887764420000005888876553000000588988765300000047888776640000004677777663000" +
        "03EFFFFFFFFFD20005FFFFFFFFFFF90005FFFFFFFFFFF80002CDDDDDEFFFF20000000002EFFFA0000000000AFFFF40000000001FFFFA00000000008F" +
        "FFF20000000003EFFFA00000000009FFFF30000000002FFFF80000000000AFFFF10000000004FFFFA0000000000BFFFE20000000003FFFF700000000" +
        "00CFFFF00000000004FFFFEDDDDDDB2005FFFFFFFFFFFF5005FFFFFFFFFFFF5003EFFFFFFFFFFD300168999999986310029CDDDDDDDC9510029CDDDE" +
        "EEEEB6100279AABCDFFEA51001345579CEEC841000112259CEDB62000000125ADEC841000000138CEDA62000000025ADEC841000000148CEDA620000" +
        "00026ADEC831000000138CEDA52000000025ADEC831000000148CEDA52100000016AEEC952211100028CEEC97555431002ADFFDCBAAA972003BEFEEE" +
        "DDDDC920029CDDDDDDDDC9200168999999998610001567778876520000267899AA9872000026799ABBB982000026789ABBBA82000014678ABBBA8200" +
        "00134689ABB9720000123579AAA8620000023579AA97510000024689A986410000134689997531000014579A986420000014689A97532000001679AA" +
        "8753200000279AAA8753210000289BBA976431000028ABBB987651000029ABBBA98762000028ABBBA9987200002789AA998872000026778877776200" +
        "00CFFFFFFFFFA00000FFFFFFFFFFF2000013333334FF70000000000004FE1000000000000CF70000000000005FD0000000000001EF40000000000008" +
        "FE0000000000001EF80000000000008FE1000000000000FF50000000000006FC0000000000001EF40000000000007FC0000000000002FF2000000000" +
        "000AFD0000000000003FF6000000000000AFF5333333310000FFFFFFFFFFFF0000CFFFFFFFFFFB00002688889999620000279999ABBB820000146666" +
        "79BB82000002233358AA6100000011124799510000000012589731000000001379852000000001258964100000000136985200000000125896310000" +
        "0000247974200000000136896310000000014798420000000002689731000000001489851000000000159974111110000027AA75333321000029BB97" +
        "666652000029BBAA99998200002799998888820000001556677200000000145667720000000014556672000000001345666200000000134456620000" +
        "000013445552000000001344555200000000134455510000000013445541000000001445554100000000145554410000000014555431000000001555" +
        "443100000000255544310000000025554431000000002665543100000000266654310000000027665541000000002776655100000000277766520000" +
        "0000000000001AD5000000000005EFFD0000000001AFFFFA000000005EFFFF700000002BFFFFC200000006EFFFF700000002BFFFFC200000006EFFFF" +
        "700000002BFFFFC200000000DFFFF70000000000DFFFE700000000002BFFFFB200000000006EFFFE700000000002BFFFFB200000000006EFFFE70000" +
        "0000002BFFFFB200000000006EFFFE700000000002BFFFF9000000000006EFFD0000000000002AD60000000012357985000000112468ABA500000124" +
        "579BBA84000113578ABBA8620012468ABBB97531012468ABBB97531013579BBBA8642100268ABBA87531100049BCCA85421000005ADDC96321000000" +
        "5ACDC9642100000048ACCB96531100002579BBB98642100012468ABBB97531100123579BBBA8642100123579ABBA86410000124679BBB97300000123" +
        "468ABB950000000123579AA5000000000124688404445566667777700455566677777770055566777777777005667777777777700667777888877760" +
        "067778888887776007778888888777600778888888877760078888988887776007888899888776600788889988877660078888988887766007788888" +
        "888777600777888888877760067778888887776006677788888777600566777777777770055566777777777004556666777777700445556666777770" +
        "000000000001AED500000000005EFFFD000000003CFFFFFE00000007FFFFFFFA000001CFFFFFFF7000004DFFFFFFE500004DFFFFFFE5000006FFFFFF" +
        "FD100000AFFFFFFF91000000FFFFFFD400000000FFFFFFD300000000AFFFFFFF8100000006FFFFFFFD100000004DFFFFFFE5000000004DFFFFFFE500" +
        "000001CFFFFFFF7000000007FFFFFFF9000000003CFFFFFE00000000006EFFFD000000000002AED60000000012357886000000012468BCB800000123" +
        "69BCEEC900001347ACDEEDB80012368ADEEECA85012479BDEEDC97532468BDEEDCA85311479BDEEDB97421007ACEEEDB863210008CEFFECA63110000" +
        "8CEFFECA531100007ACEEEDB85321000479CDEEDB96421002468BDEEDCA85311012479BDEEDC97530012468ADEEECA8500001347ACDEEDB800000123" +
        "69BDEEC9000000012469BCB800000000123578860011223456677774011233456788988512234567889999851234567899A99984234667899AA99874" +
        "3567899AAA9987633689AAAAA9876542479ABBBAA876543258ABBCBA9865432159ABCCBA9765432159ABCCBA9765432158ABBCBA98654321479ABBBA" +
        "A87654313689AAAAA98765423567899AAA998753234677899AA998741234567899999984122345678899998501123345678898850011223456677774" +
        "00000000000007C2000000000002CFF700000000007FFF90000000003CFFD40000000007FFF91000000003CFFD40000000007FFF91000000003CFFD4" +
        "0000000008FFF910000000008FFF5000000000008FFF50000000000008FFF81000000000003CFFD30000000000008FFF81000000000003CFFD300000" +
        "00000008FFF81000000000003CFFD30000000000007FFF80000000000003CFF700000000000007C20000000012346785000000012357898500000123" +
        "4678987400011245789886420012346799876421012357899864321013568998754211002578998753210000479AA8642110000059BB974210000000" +
        "59BB97421000000047AAA864211000002578998753210000135689987542110001245789986432100012346799875421000112457898864200000123" +
        "467898740000001123578985000000001235678500000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" +
        "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000";
    // высота знака / высота заглавной; смещение середины; ширина / высота
    public static readonly string[] Shapes = {
        "1.087;-0.043;0.640",
        "1.087;-0.043;0.400",
        "1.072;-0.051;0.608",
        "1.072;-0.036;0.635",
        "1.087;-0.043;0.640",
        "1.072;-0.036;0.568",
        "1.087;-0.043;0.640",
        "1.072;-0.036;0.649",
        "1.087;-0.043;0.640",
        "1.087;-0.043;0.640",
        "1.000;0.000;0.681",
        "0.971;0.000;0.716",
        "1.014;-0.007;0.557",
        "0.971;0.000;0.597",
        "0.971;0.000;0.642",
        "0.986;0.007;0.529",
        "1.029;0.000;0.634",
        "1.000;0.000;0.623",
        "0.971;0.000;0.552",
        "1.014;0.007;0.500",
        "1.000;0.000;0.681",
        "0.986;-0.007;0.676",
        "1.000;0.000;0.696",
        "1.000;0.000;0.667",
        "1.014;-0.007;0.686",
        "0.986;0.007;0.676",
        "1.014;-0.007;0.686",
        "0.986;0.007;0.618",
        "1.029;0.000;0.592",
        "0.986;0.007;0.706",
        "1.000;0.000;0.667",
        "1.000;0.000;0.696",
        "1.000;0.000;0.696",
        "1.000;0.000;0.638",
        "1.000;0.000;0.638",
        "0.971;0.000;0.627",
        "0.884;-0.043;0.787",
    };
}

// Имена и отчества (по-русски, заглавными) - для проверки распознанного с паспорта.
// Имени нет в списке - не беда: тогда программа подставит распознанное, но попросит проверить.
public static class NamesData {
    public const string Female =
        "АВГУСТА АГАТА АГНИЯ АДА АДЕЛИНА АДЕЛЬ АКСИНЬЯ АЛЕВТИНА АЛЕКСАНДРА АЛЕНА АЛЁНА АЛЕСЯ АЛИНА АЛИСА АЛЛА АЛЬБИНА " +
        "АМЕЛИЯ АНАСТАСИЯ АНГЕЛИНА АНЖЕЛА АНЖЕЛИКА АННА АНТОНИНА АРИАДНА АРИНА АСЯ БЕАТА БОГДАНА БРОНИСЛАВА ВАЛЕНТИНА ВАЛЕРИЯ " +
        "ВАНДА ВАРВАРА ВАСИЛИСА ВЕРА ВЕРОНИКА ВИКТОРИЯ ВИОЛЕТТА ВИТАЛИНА ВЛАДА ВЛАДИСЛАВА ГАЛИНА ГЕНОВЕФА ГЛАФИРА ДАНА " +
        "ДАНУТА ДАРИНА ДАРЬЯ ДАРИЯ ДИАНА ДИНА ДОМИНИКА ЕВА ЕВГЕНИЯ ЕВДОКИЯ ЕКАТЕРИНА ЕЛЕНА ЕЛИЗАВЕТА ЕФРОСИНЬЯ ЖАННА ЗИНАИДА " +
        "ЗЛАТА ЗОЯ ИВАННА ИЗАБЕЛЛА ИНГА ИНЕССА ИННА ИРИНА ИРМА ЯДВИГА ЯНА ЯНИНА ЯРОСЛАВА ЮЗЕФА ЮЛИАНА ЮЛИЯ ЭВЕЛИНА ЭЛЕОНОРА " +
        "ЭЛИНА ЭЛЛА ЭЛЬВИРА ЭМИЛИЯ ЭММА ТАИСИЯ ТАМАРА ТАТЬЯНА ТАТИАНА ТЕРЕЗА УЛЬЯНА ФАИНА ФЕОДОРА СВЕТЛАНА СНЕЖАНА СОФИЯ " +
        "СОФЬЯ СТАНИСЛАВА СТЕФАНИЯ СТЕФАНИДА РАИСА РЕГИНА РИММА РОЗА РУСЛАНА ОКСАНА ОЛЕСЯ ОЛЬГА ПАВЛИНА ПОЛИНА ПРАСКОВЬЯ " +
        "НАДЕЖДА НАТАЛЬЯ НАТАЛИЯ НЕЛЛИ НИКА НИНА НОННА МАЙЯ МАРГАРИТА МАРИАННА МАРИНА МАРИЯ МАРЬЯ МАРЬЯНА МАРТА МИЛАНА " +
        "МИЛЕНА МИРОСЛАВА ЛАРИСА ЛЕОКАДИЯ ЛЕСЯ ЛИДИЯ ЛИЛИЯ ЛИЯ ЛЮБОВЬ ЛЮДМИЛА КАРИНА КАРОЛИНА КИРА КЛАВДИЯ КРИСТИНА КСЕНИЯ " +
        "ЧЕСЛАВА ЭЛЕНА ЕВГЕНИЯ АЛЬБЕРТИНА ВЕНЕРА ЯНИНА АУРИКА ВИОЛЕТА ВЕРОНИКА ВИТА ДАЯНА МИЯ ЕСЕНИЯ";

    public const string Male =
        "АЛЕКСАНДР АЛЕКСЕЙ АЛЕС АЛЬБЕРТ АНАТОЛИЙ АНДРЕЙ АНТОН АРКАДИЙ АРСЕНИЙ АРТЕМ АРТЁМ АРТУР АФАНАСИЙ БОГДАН БОРИС " +
        "БРОНИСЛАВ ВАДИМ ВАЛЕНТИН ВАЛЕРИЙ ВАСИЛИЙ ВЕНИАМИН ВИКТОР ВИТАЛИЙ ВИТОЛЬД ВЛАДИМИР ВЛАДИСЛАВ ВЯЧЕСЛАВ ГЕННАДИЙ " +
        "ГЕОРГИЙ ГЕРМАН ГЛЕБ ГРИГОРИЙ ДАВИД ДАНИИЛ ДАНИЛ ДАНИЛА ДЕНИС ДМИТРИЙ ЕВГЕНИЙ ЕГОР ЕМЕЛЬЯН ЕФИМ ЗАХАР ЗЕНОН ЗИНОВИЙ " +
        "ИВАН ИГНАТ ИГОРЬ ИЛЬЯ ИННОКЕНТИЙ ИОСИФ КАЗИМИР КИРИЛЛ КОНСТАНТИН КУЗЬМА ЛАВРЕНТИЙ ЛЕВ ЛЕОНИД ЛЕОНТИЙ ЛУКА МАКАР " +
        "МАКСИМ МАРК МАТВЕЙ МИРОН МИРОСЛАВ МИХАИЛ НАЗАР НИКИТА НИКОЛАЙ ОЛЕГ ОСИП ПАВЕЛ ПЕТР ПЁТР ПЛАТОН ПРОХОР РОДИОН РОМАН " +
        "РОСТИСЛАВ РУДОЛЬФ РУСЛАН САВВА СВЯТОСЛАВ СЕМЕН СЕМЁН СЕРГЕЙ СТАНИСЛАВ СТЕПАН ТАРАС ТИМОФЕЙ ТИМУР ТИХОН ФЕДОР ФЁДОР " +
        "ФЕЛИКС ФИЛИПП ФОМА ЭДУАРД ЭДУАРД ЭМИЛЬ ЭРИК ЮЛИАН ЮЗЕФ ЮРИЙ ЯКОВ ЯН ЯРОСЛАВ ЧЕСЛАВ ВЛАДИЛЕН ГЕНРИХ ДАНИЭЛЬ " +
        "МИХАЙЛО ЕЛИСЕЙ АРСЕН ВИКЕНТИЙ ГАВРИИЛ ДАВЫД ИЗРАИЛЬ КЛИМ ЛЕОНАРД НЕСТОР ОСКАР РАФАЭЛЬ СЕВАСТЬЯН ТРОФИМ ФРАНЦ";

    // отчества от необычных имён (остальные - по правилам)
    static readonly string[] Special = {
        "ИЛЬЯ=ИЛЬИЧ,ИЛЬИНИЧНА", "НИКИТА=НИКИТИЧ,НИКИТИЧНА", "КУЗЬМА=КУЗЬМИЧ,КУЗЬМИНИЧНА", "ФОМА=ФОМИЧ,ФОМИНИЧНА",
        "ЛУКА=ЛУКИЧ,ЛУКИНИЧНА", "САВВА=САВВИЧ,САВВИЧНА", "ДАНИЛА=ДАНИЛОВИЧ,ДАНИЛОВНА", "ПАВЕЛ=ПАВЛОВИЧ,ПАВЛОВНА",
        "ЛЕВ=ЛЬВОВИЧ,ЛЬВОВНА", "ПЁТР=ПЕТРОВИЧ,ПЕТРОВНА", "ЯКОВ=ЯКОВЛЕВИЧ,ЯКОВЛЕВНА", "МИХАИЛ=МИХАЙЛОВИЧ,МИХАЙЛОВНА",
        "АРТЁМ=АРТЕМОВИЧ,АРТЕМОВНА", "СЕМЁН=СЕМЕНОВИЧ,СЕМЕНОВНА", "ФЁДОР=ФЕДОРОВИЧ,ФЕДОРОВНА", "МИХАЙЛО=МИХАЙЛОВИЧ,МИХАЙЛОВНА" };

    static List<string> female, male, patronymic;
    static readonly object sync = new object();

    public static List<string> FemaleNames { get { Build(); return female; } }
    public static List<string> MaleNames { get { Build(); return male; } }
    public static List<string> Patronymics { get { Build(); return patronymic; } }

    static void Build() {
        lock (sync) {
            if (patronymic != null) return;
            female = Words(Female);
            male = Words(Male);
            List<string> p = new List<string>();
            Dictionary<string, string[]> sp = new Dictionary<string, string[]>();
            foreach (string s in Special) {
                int eq = s.IndexOf('=');
                sp[s.Substring(0, eq)] = s.Substring(eq + 1).Split(',');
            }
            foreach (string m in male) {
                string[] f;
                if (sp.TryGetValue(m, out f)) { Add(p, f[0]); Add(p, f[1]); continue; }
                foreach (string stem in Stems(m)) { Add(p, stem + "ИЧ"); Add(p, stem + "НА"); }
            }
            patronymic = p;
        }
    }

    // основа отчества с «-ВИЧ»: СЕРГЕЙ -> СЕРГЕЕВ, ИВАН -> ИВАНОВ, ДМИТРИЙ -> ДМИТРИЕВ (и ДМИТРЬЕВ), ИГОРЬ -> ИГОРЕВ
    static List<string> Stems(string m) {
        List<string> r = new List<string>();
        if (m.EndsWith("ИЙ")) {
            string b = m.Substring(0, m.Length - 2);
            r.Add(b + "ЬЕВ");
            r.Add(b + "ИЕВ");
        } else if (m.EndsWith("ЕЙ") || m.EndsWith("АЙ")) r.Add(m.Substring(0, m.Length - 1) + "ЕВ");
        else if (m.EndsWith("Ь")) r.Add(m.Substring(0, m.Length - 1) + "ЕВ");
        else if (m.EndsWith("А") || m.EndsWith("Я") || m.EndsWith("О")) r.Add(m.Substring(0, m.Length - 1) + "ОВ");
        else r.Add(m + "ОВ");
        return r;
    }

    static void Add(List<string> l, string s) { if (!l.Contains(s)) l.Add(s); }

    static List<string> Words(string s) {
        List<string> l = new List<string>();
        foreach (string w in s.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) Add(l, w);
        return l;
    }
}

// Предпоследняя страница паспорта РБ: Ф.И.О. (по-белорусски / по-русски) и идентификационный номер.
// Своё распознавание, без сторонних программ: буквы - жирный рубленый шрифт (как Arial), образцы рисуются
// шрифтами этого компьютера; номер и даты - OCR-B (образцы те же, что для строк «<<<»).

// Образцы знаков: у каждого знака несколько видов (шрифт x размытие)
public class GlyphSet {
    public readonly string Chars;
    readonly List<float[]> grids = new List<float[]>();     // нормированы: среднее 0, длина 1
    readonly List<int> owner = new List<int>();
    readonly List<double> relH = new List<double>(), relY = new List<double>(), asp = new List<double>();
    public int Fonts;
    public int GW = Mrz.GW, GH = Mrz.GH;      // сетка сравнения

    GlyphSet(string chars) { Chars = chars; }

    public float[] Grid(float[] ink, int stride, int x0, int y0, int x1, int y1) { return Mrz.Grid(ink, stride, x0, y0, x1, y1, GW, GH); }

    static float[] Norm(float[] g) {
        double m = 0;
        foreach (float v in g) m += v;
        m /= g.Length;
        double ss = 0;
        foreach (float v in g) ss += (v - m) * (v - m);
        double k = ss > 1e-9 ? 1 / Math.Sqrt(ss) : 0;
        float[] r = new float[g.Length];
        for (int i = 0; i < g.Length; i++) r[i] = (float)((g[i] - m) * k);
        return r;
    }

    void Add(int ch, float[] grid, double h, double y, double a) {
        grids.Add(Norm(grid));
        owner.Add(ch);
        relH.Add(h);
        relY.Add(y);
        asp.Add(a);
    }

    // оценка каждого знака (лучший из его видов): сходство рисунка минус штраф за размер, положение, пропорции
    public double[] Score(float[] grid, double h, double y, double a) {
        float[] g = Norm(grid);
        double[] best = new double[Chars.Length];
        for (int i = 0; i < best.Length; i++) best[i] = -9;
        for (int t = 0; t < grids.Count; t++) {
            float[] tg = grids[t];
            double c = 0;
            for (int i = 0; i < g.Length; i++) c += g[i] * tg[i];
            double s = c - 1.0 * Math.Abs(h - relH[t]) - 0.6 * Math.Abs(y - relY[t]) - 0.3 * Math.Abs(Math.Log(a / asp[t]));
            if (s > best[owner[t]]) best[owner[t]] = s;
        }
        return best;
    }

    // образцы рисуются жирными шрифтами этого компьютера (каких нет - пропускаются)
    public static GlyphSet FromFonts(string chars, string[] families) { return FromFonts(chars, families, new FontStyle[] { FontStyle.Bold }); }

    public static GlyphSet FromFonts(string chars, string[] families, FontStyle[] styles) { return FromFonts(chars, families, styles, Mrz.GW, Mrz.GH); }

    public static GlyphSet FromFonts(string chars, string[] families, FontStyle[] styles, int gw, int gh) {
        return FromFonts(chars, families, styles, gw, gh, new double[] { 0, 0.045, 0.09 });
    }

    // blurs - размытие образцов (доля высоты заглавной)
    public static GlyphSet FromFonts(string chars, string[] families, FontStyle[] styles, int gw, int gh, double[] blurs) {
        GlyphSet s = new GlyphSet(chars);
        s.GW = gw;
        s.GH = gh;
        const int W = 200, H = 200;
        List<float[]> seen = new List<float[]>();
        lock (Mrz.Gdi) {
            foreach (string fam in families)
            foreach (FontStyle style in styles) {
                Font f = null;
                try { f = new Font(fam, 64f, style, GraphicsUnit.Pixel); } catch { f = null; }
                if (f == null) continue;
                using (f) {
                    int x0, y0, x1, y1;
                    if (!Box(Render(f, "Н", W, H), W, H, out x0, out y0, out x1, out y1)) continue;
                    double cap = y1 - y0 + 1, capMid = (y0 + y1) / 2.0;
                    // тот же шрифт под другим именем (Windows подставляет замену) - второй раз не берём
                    float[] probe = Probe(f, W, H);
                    bool dup = false;
                    foreach (float[] p in seen) {
                        double c = 0;
                        for (int i = 0; i < p.Length; i++) c += p[i] * probe[i];
                        if (c > 0.995) dup = true;
                    }
                    if (dup) continue;
                    seen.Add(probe);
                    s.Fonts++;
                    for (int k = 0; k < chars.Length; k++) {
                        float[] ink = Render(f, chars[k].ToString(), W, H);
                        foreach (double sigma in blurs) {
                            float[] b = sigma > 0 ? Blur(ink, W, H, sigma * cap) : ink;
                            if (!Box(b, W, H, out x0, out y0, out x1, out y1)) continue;
                            s.Add(k, s.Grid(b, W, x0, y0, x1, y1), (y1 - y0 + 1) / cap, ((y0 + y1) / 2.0 - capMid) / cap, (x1 - x0 + 1.0) / (y1 - y0 + 1));
                        }
                    }
                }
            }
        }
        return s;
    }

    static float[] Probe(Font f, int W, int H) {
        List<float> all = new List<float>();
        foreach (string t in new string[] { "Ж", "Я", "Ю" }) {
            float[] ink = Render(f, t, W, H);
            int x0, y0, x1, y1;
            if (Box(ink, W, H, out x0, out y0, out x1, out y1)) all.AddRange(Mrz.Grid(ink, W, x0, y0, x1, y1));
            all.Add((float)(x1 - x0 + 1) / (y1 - y0 + 1));
        }
        return Norm(all.ToArray());
    }

    static float[] Render(Font f, string text, int W, int H) {
        using (Bitmap b = new Bitmap(W, H, PixelFormat.Format24bppRgb)) {
            using (Graphics g = Graphics.FromImage(b)) {
                g.Clear(Color.White);
                g.TextRenderingHint = TextRenderingHint.AntiAlias;
                g.DrawString(text, f, Brushes.Black, 40, 30);
            }
            GrayImage gi = GrayImage.FromBitmap(b);
            float[] ink = new float[W * H];
            for (int i = 0; i < ink.Length; i++) ink[i] = 1 - gi.P[i] / 255f;
            return ink;
        }
    }

    public static bool Box(float[] ink, int w, int h, out int x0, out int y0, out int x1, out int y1) {
        x0 = w; y0 = h; x1 = -1; y1 = -1;
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++)
                if (ink[y * w + x] > 0.45f) {
                    if (x < x0) x0 = x;
                    if (x > x1) x1 = x;
                    if (y < y0) y0 = y;
                    if (y > y1) y1 = y;
                }
        return x1 >= 0;
    }

    // размытие по Гауссу (как у снимка телефоном) - только вокруг самого знака
    static float[] Blur(float[] src, int w, int h, double sigma) {
        int r = Math.Max(1, (int)Math.Ceiling(sigma * 3));
        int bx0 = w, by0 = h, bx1 = -1, by1 = -1;
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++)
                if (src[y * w + x] > 0.01f) { if (x < bx0) bx0 = x; if (x > bx1) bx1 = x; if (y < by0) by0 = y; if (y > by1) by1 = y; }
        float[] dst = new float[w * h];
        if (bx1 < 0) return dst;
        bx0 = Math.Max(0, bx0 - r); by0 = Math.Max(0, by0 - r); bx1 = Math.Min(w - 1, bx1 + r); by1 = Math.Min(h - 1, by1 + r);
        double[] k = new double[2 * r + 1];
        double sum = 0;
        for (int i = -r; i <= r; i++) { k[i + r] = Math.Exp(-i * i / (2 * sigma * sigma)); sum += k[i + r]; }
        for (int i = 0; i < k.Length; i++) k[i] /= sum;
        float[] tmp = new float[w * h];
        for (int y = by0; y <= by1; y++)
            for (int x = bx0; x <= bx1; x++) {
                double s = 0;
                for (int i = -r; i <= r; i++) { int xx = x + i; if (xx >= 0 && xx < w) s += src[y * w + xx] * k[i + r]; }
                tmp[y * w + x] = (float)s;
            }
        for (int y = by0; y <= by1; y++)
            for (int x = bx0; x <= bx1; x++) {
                double s = 0;
                for (int i = -r; i <= r; i++) { int yy = y + i; if (yy >= 0 && yy < h) s += tmp[yy * w + x] * k[i + r]; }
                dst[y * w + x] = (float)s;
            }
        return dst;
    }

    // образцы OCR-B (из тех же данных, что для строк «<<<»)
    public static GlyphSet Ocrb(string chars) {
        GlyphSet s = new GlyphSet(chars);
        int size = Mrz.GW * Mrz.GH;
        for (int n = 0; n < chars.Length; n++) {
            int k = Mrz.Chars.IndexOf(chars[n]);
            string[] p = MrzData.Shapes[k].Split(';');
            double rh = double.Parse(p[0], System.Globalization.CultureInfo.InvariantCulture);
            double ry = double.Parse(p[1], System.Globalization.CultureInfo.InvariantCulture);
            double a = double.Parse(p[2], System.Globalization.CultureInfo.InvariantCulture);
            for (int v = 0; v < Mrz.TV; v++) {
                float[] g = new float[size];
                for (int i = 0; i < size; i++) g[i] = Convert.ToInt32(MrzData.Grids[(k * Mrz.TV + v) * size + i].ToString(), 16) / 15f;
                s.Add(n, g, rh, ry, a);
            }
        }
        s.Fonts = 1;
        return s;
    }
}

// Распознанный знак строки
public class Glyph {
    public char C;
    public double S, S2;          // оценка лучшего и второго варианта
    public char C2;
    public double[] Sc;           // оценки всех знаков набора
    public bool Space;            // перед знаком пробел
    public int X0, X1;
    internal object Src;          // строка, откуда знак (для перечитывания по образцу)
    internal int CutA, CutB;
}

public class PageResult {
    public string Surname = "", Name = "", Patronymic = "";          // по-русски, заглавными
    public string SurnameBy = "", NameBy = "", PatronymicBy = "";    // по-белорусски
    public bool SurnameOk, NameOk, PatronymicOk;                      // уверены
    public bool NameKnown;                                            // имя есть в списке имён
    public string Id = "";                                            // идентификационный номер
    public bool IdOk;
    public string Birth = "";                                         // дата рождения ДД.ММ.ГГГГ
    public bool Found;                                                // похоже на страницу с данными
    public string Info = "";
    public List<string> Lines = new List<string>();                   // для проверки программы

    public bool FioOk { get { return SurnameOk && NameOk && PatronymicOk; } }

    public string Fio {
        get { return (Title(Surname) + " " + Title(Name) + " " + Title(Patronymic)).Trim(); }
    }

    // «АБАКУНЧИК-ПЕТРОВА» -> «Абакунчик-Петрова»
    public static string Title(string s) {
        StringBuilder sb = new StringBuilder();
        bool start = true;
        foreach (char c in s) {
            sb.Append(start ? char.ToUpper(c) : char.ToLower(c));
            start = c == '-' || c == ' ';
        }
        return sb.ToString();
    }
}

public static class PassportPage {
    // рубленые шрифты, которые есть в Windows 7; в паспорте - похожий на Arial
    public static string[] Families = { "Arial", "Tahoma", "Verdana", "Segoe UI", "Calibri", "Trebuchet MS", "Franklin Gothic Medium", "Microsoft Sans Serif" };
    public static string DebugDir;      // для проверки программы: картинки строк с разрезами
    public const string SansChars = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯІЎ0123456789/-.,";
    const string OcrbChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    static GlyphSet sans, ocrb;
    static readonly object sync = new object();

    public static void Init() {
        lock (sync) {
            if (sans != null) return;
            ocrb = GlyphSet.Ocrb(OcrbChars);
            sans = GlyphSet.FromFonts(SansChars, Families);
        }
    }

    public static int FontCount { get { Init(); return sans.Fonts; } }

    internal class Line {
        public List<Mrz.Comp> Comps;
        public bool Mixed;                  // обычный текст (строчные и заглавные), а не одни заглавные
        public double Cap, Y, MedH, H80;
        public int X0;
        public List<Glyph> G;
        public string Text = "";
    }

    // ---------- строки текста на странице ----------

    internal static List<Line> FindLines(GrayImage g) { return FindLines(g, false); }

    // mixed - обычный текст: строчные буквы, широкие пробелы (текст по ширине)
    internal static List<Line> FindLines(GrayImage g, bool mixed) {
        bool[] dark = Mrz.Binarize(g);
        List<Mrz.Comp> cand = new List<Mrz.Comp>();
        foreach (Mrz.Comp c in Mrz.Components(dark, g.W, g.H))
            // слипшиеся буквы (размытый жирный шрифт) - широкая, но плотная область; линии узора - редкие
            if (c.Ht >= 8 && c.Ht <= g.H / 12 && c.N >= 12 && (c.Wd <= 4 * c.Ht || c.Wd <= 40 * c.Ht && c.N >= 0.2 * c.Wd * c.Ht)) cand.Add(c);
        cand.Sort(delegate(Mrz.Comp a, Mrz.Comp b) { return a.X0.CompareTo(b.X0); });
        List<Line> lines = new List<Line>();
        bool[] used = new bool[cand.Count];
        for (int s = 0; s < cand.Count; s++) {
            if (used[s]) continue;
            List<int> chain = new List<int>();
            chain.Add(s);
            int cur = s;
            double href = cand[s].Ht, yref = cand[s].Cy, slope = 0, xref = cand[s].Cx;
            while (true) {
                Mrz.Comp c = cand[cur];
                int next = -1;
                for (int j = cur + 1; j < cand.Count; j++) {
                    Mrz.Comp d = cand[j];
                    if (d.X0 > c.X1 + (mixed ? 4.0 : 1.4) * href) break;
                    if (used[j] || d.X0 < c.X0 + 0.3 * c.Wd) continue;
                    if (Math.Abs(d.Cy - (yref + slope * (d.Cx - xref))) > 0.45 * href) continue;
                    if (d.Ht < 0.3 * href || d.Ht > 1.7 * href) continue;
                    if (next < 0 || d.X0 < cand[next].X0) next = j;
                }
                if (next < 0) break;
                chain.Add(next);
                cur = next;
                Mrz.Comp n = cand[next];
                // строка может идти с наклоном: следующий знак ищем по продолжению линии
                if (n.Ht >= (mixed ? 0.5 : 0.7) * href && n.Ht <= 1.3 * href) {
                    double dxs = n.Cx - xref;
                    if (dxs > 0.5 * href) slope = 0.6 * slope + 0.4 * Math.Max(-0.4, Math.Min(0.4, (n.Cy - yref) / dxs));
                    yref = n.Cy;
                    xref = n.Cx;
                }
                href = Math.Max(href, Math.Min(n.Ht, href * 1.4));
            }
            if (chain.Count < (mixed ? 3 : 4)) continue;
            Line l = new Line();
            l.Mixed = mixed;
            l.Comps = new List<Mrz.Comp>();
            foreach (int i in chain) { used[i] = true; l.Comps.Add(cand[i]); }
            List<int> hs = new List<int>();
            foreach (Mrz.Comp c in l.Comps) hs.Add(c.Ht);
            hs.Sort();
            l.Cap = hs[(int)(hs.Count * 0.8)];
            // в обычном тексте строка из одних заглавных (заголовок) читается как заглавные
            l.MedH = hs[(int)(hs.Count * 0.4)];
            l.H80 = hs[(int)(hs.Count * 0.8)];
            if (l.Mixed) {
                // высота заглавной - по строчным (их больше всего): у шрифта с засечками заглавная в 1,48 раза выше
                l.Cap = 1.48 * hs[(int)(hs.Count * 0.4)];
                List<double> cys = new List<double>();
                foreach (Mrz.Comp c in l.Comps) cys.Add(c.Cy);
                cys.Sort();
                l.Y = cys[cys.Count / 2];
                l.X0 = l.Comps[0].X0;
                lines.Add(l);
                continue;
            }
            // строка заглавных: почти все знаки высотой с заглавную (у подписей полей - строчные)
            int tall = 0;
            double ys = 0;
            foreach (Mrz.Comp c in l.Comps) if (c.Ht >= 0.75 * l.Cap && c.Ht <= 1.3 * l.Cap) { tall++; ys += c.Cy; }
            if (tall < 0.7 * l.Comps.Count) continue;
            l.Y = ys / tall;
            l.X0 = l.Comps[0].X0;
            lines.Add(l);
        }
        if (mixed && lines.Count > 0) {
            // строки из одних заглавных (заголовки): их буквы заметно выше строчных основного текста
            List<double> med = new List<double>();
            foreach (Line l in lines) for (int k = 0; k < l.Comps.Count; k++) med.Add(l.MedH);
            med.Sort();
            double body = med[med.Count / 2];
            foreach (Line l in lines)
                if (l.MedH >= 1.3 * body) { l.Mixed = false; l.Cap = l.H80; }
        }
        lines.Sort(delegate(Line a, Line b) { return a.Y.CompareTo(b.Y); });
        return lines;
    }

    // ---------- строка -> знаки ----------

    static int[] Label(bool[] bin, int w, int h, out int count) {
        int[] lab = new int[w * h];
        count = 0;
        int[] stack = new int[w * h];
        for (int i = 0; i < bin.Length; i++) {
            if (!bin[i] || lab[i] != 0) continue;
            count++;
            int sp = 0;
            stack[sp++] = i;
            lab[i] = count;
            while (sp > 0) {
                int p = stack[--sp], x = p % w, y = p / w;
                for (int dy = -1; dy <= 1; dy++)
                    for (int dx = -1; dx <= 1; dx++) {
                        int xx = x + dx, yy = y + dy;
                        if (xx < 0 || yy < 0 || xx >= w || yy >= h) continue;
                        int q = yy * w + xx;
                        if (bin[q] && lab[q] == 0) { lab[q] = count; stack[sp++] = q; }
                    }
            }
        }
        return lab;
    }


    internal static Strip ReadLine(GrayImage g, Line line, GlyphSet set) {
        double cap = line.Cap;
        bool mixed = line.Mixed;
        // наклон строки - по центрам знаков высотой с заглавную (в обычном тексте - по низу строчных)
        double sx = 0, sy = 0, sxx = 0, sxy = 0;
        int n = 0;
        foreach (Mrz.Comp c in line.Comps) {
            double y;
            if (!mixed) { if (c.Ht < 0.75 * cap || c.Ht > 1.3 * cap) continue; y = c.Cy; }
            else { if (c.Ht < 0.5 * cap || c.Ht > 0.9 * cap) continue; y = c.Y1 + 0.5 - cap / 2; }
            sx += c.Cx; sy += y; sxx += c.Cx * c.Cx; sxy += c.Cx * y; n++;
        }
        if (n < 3) return null;
        double den = n * sxx - sx * sx;
        double b = Math.Abs(den) < 1e-9 ? 0 : (n * sxy - sx * sy) / den, a = (sy - b * sx) / n;
        double norm = Math.Sqrt(1 + b * b), dx = 1 / norm, dy = b / norm;
        double x0 = line.Comps[0].X0 - 0.8 * cap, x1 = line.Comps[line.Comps.Count - 1].X1 + 0.8 * cap;
        int L = (int)((x1 - x0) * norm), SH = (int)((mixed ? 2.6 : 2.4) * cap);
        if (L < 10 || SH < 10) return null;
        double px = x0, py = a + b * x0;
        float[] strip = new float[L * SH];
        for (int v = 0; v < SH; v++)
            for (int u = 0; u < L; u++) {
                double off = v - SH / 2.0;
                strip[v * L + u] = g.Sample(px + u * dx - off * dy, py + u * dy + off * dx);
            }
        float[] sorted = (float[])strip.Clone();
        Array.Sort(sorted);
        float bg = sorted[(int)(sorted.Length * 0.8)], fg = sorted[(int)(sorted.Length * 0.02)];
        if (bg - fg < 25) return null;
        float[] ink = new float[strip.Length];
        bool[] bin = new bool[strip.Length];
        // только полоса строки (с запасом на кратки и хвосты букв): соседние строки не слипаются с нашими буквами
        int bandTop = Math.Max(0, (int)(SH / 2.0 - (mixed ? 0.75 : 0.8) * cap)), bandBot = Math.Min(SH - 1, (int)(SH / 2.0 + (mixed ? 0.95 : 0.8) * cap));
        for (int i = 0; i < strip.Length; i++) {
            float k = (bg - strip[i]) / (bg - fg);
            ink[i] = k < 0 ? 0 : k > 1 ? 1 : k;
            int row = i / L;
            bin[i] = ink[i] > 0.45f && row >= bandTop && row <= bandBot;
        }
        int count;
        int[] lab = Label(bin, L, SH, out count);
        int[] cx0 = new int[count + 1], cx1 = new int[count + 1], cy0 = new int[count + 1], cy1 = new int[count + 1], cn = new int[count + 1];
        double[] cink = new double[count + 1], csx = new double[count + 1];
        for (int i = 1; i <= count; i++) { cx0[i] = L; cy0[i] = SH; cx1[i] = -1; cy1[i] = -1; }
        for (int y = 0; y < SH; y++)
            for (int x = 0; x < L; x++) {
                int l = lab[y * L + x];
                if (l == 0) continue;
                if (x < cx0[l]) cx0[l] = x;
                if (x > cx1[l]) cx1[l] = x;
                if (y < cy0[l]) cy0[l] = y;
                if (y > cy1[l]) cy1[l] = y;
                cn[l]++;
                cink[l] += ink[y * L + x];
                csx[l] += x;
            }
        // знаки строки: задевают середину полосы; чужие строки сверху и снизу - мимо
        double mid = SH / 2.0;
        bool[] keep = new bool[count + 1];
        List<double> hts = new List<double>();
        for (int i = 1; i <= count; i++) {
            if (cn[i] < 4) continue;
            if (cy1[i] < mid - 0.45 * cap || cy0[i] > mid + (mixed ? 0.6 : 0.45) * cap) continue;
            // обрезок буквы соседней строки у края полосы; край страницы, тень
            if (cy0[i] == bandTop && (cy0[i] + cy1[i]) / 2.0 < mid - 0.3 * cap) continue;
            if (cy0[i] == bandTop && cy1[i] == bandBot) continue;
            if (cx0[i] == 0 || cx1[i] == L - 1) continue;
            if (cy1[i] == bandBot && (cy0[i] + cy1[i]) / 2.0 > mid + 0.3 * cap) continue;
            if (cy1[i] - cy0[i] + 1 > 1.9 * cap) continue;
            keep[i] = true;
            double h = cy1[i] - cy0[i] + 1;
            if (!mixed ? h > 0.7 * cap && h < 1.3 * cap : h > 0.5 * cap && h < 0.85 * cap) hts.Add(h);
        }
        // в обычном тексте мелкие бледные пятнышки - обрывки защитного узора, а не точки и кратки
        if (mixed)
            for (int i = 1; i <= count; i++)
                if (keep[i] && cy1[i] - cy0[i] + 1 <= 0.45 * cap && cink[i] / Math.Max(1, cn[i]) < 0.6) keep[i] = false;
        // мелочь над и под строкой (кратка над Й, Ў, точки Ё) - тоже своя, если рядом с буквой
        for (int i = 1; i <= count; i++) {
            if (keep[i] || cn[i] < 3) continue;
            double h = cy1[i] - cy0[i] + 1, w = cx1[i] - cx0[i] + 1;
            if (h > 0.45 * cap || w > 0.9 * cap) continue;
            if (cy1[i] < mid - 1.0 * cap || cy0[i] > mid + 1.0 * cap) continue;
            if (cy0[i] == bandTop || cy1[i] == bandBot) continue;
            if (mixed && cink[i] / Math.Max(1, cn[i]) < 0.6) continue;
            keep[i] = true;
        }
        if (hts.Count < 2) return null;
        hts.Sort();
        double capS = hts[hts.Count / 2];
        double ym = 0; int ymn = 0;
        if (!mixed) {
            for (int i = 1; i <= count; i++) {
                double h = cy1[i] - cy0[i] + 1;
                if (keep[i] && h > 0.85 * capS && h < 1.15 * capS) { ym += (cy0[i] + cy1[i]) / 2.0; ymn++; }
            }
            if (ymn > 0) mid = ym / ymn;
        } else {
            // по строчным: середина заглавной - выше линии строки на половину заглавной
            double xh = capS;
            capS = 1.48 * xh;
            List<int> bot = new List<int>();
            for (int i = 1; i <= count; i++) {
                double h = cy1[i] - cy0[i] + 1;
                if (keep[i] && h > 0.85 * xh && h < 1.15 * xh) bot.Add(cy1[i]);
            }
            bot.Sort();
            if (bot.Count > 0) mid = bot[bot.Count / 2] + 0.5 - capS / 2;
        }
        // только свои знаки: области строки
        bool[] own = (bool[])keep.Clone();
        // обычный текст: точки и запятые - отдельно (иначе «г.» под перекладиной «г» читается как «п»)
        List<Glyph> punct = new List<Glyph>();
        if (mixed) {
            double baseY = mid + capS / 2;
            List<int> low = new List<int>(), high = new List<int>();
            for (int i = 1; i <= count; i++) {
                if (!own[i]) continue;
                double h = cy1[i] - cy0[i] + 1, w = cx1[i] - cx0[i] + 1;
                if (h > 0.3 * capS || w > 0.35 * capS) continue;
                if (cy1[i] >= baseY - 0.12 * capS && cy0[i] >= baseY - 0.3 * capS) low.Add(i);
                else if (cy1[i] < baseY - 0.25 * capS && cy0[i] > mid - 0.25 * capS) high.Add(i);
            }
            foreach (int i in low) {
                // запятая - вытянута вниз под строку, точка - круглая
                double ph = cy1[i] - cy0[i] + 1, pw = cx1[i] - cx0[i] + 1;
                char c = cy1[i] > baseY + 0.12 * capS || ph > 1.4 * pw ? ',' : '.';
                // двоеточие / точка с запятой: над точкой - ещё точка
                foreach (int j in high)
                    if (own[j] && Math.Min(cx1[i], cx1[j]) - Math.Max(cx0[i], cx0[j]) >= 0) { c = c == ',' ? ';' : ':'; own[j] = false; }
                own[i] = false;
                Glyph pg = new Glyph();
                pg.C = c; pg.S = 1; pg.C2 = c; pg.S2 = 0;
                pg.Sc = new double[set.Chars.Length];
                for (int k = 0; k < pg.Sc.Length; k++) pg.Sc[k] = set.Chars[k] == c ? 1 : -1;
                pg.X0 = cx0[i];
                pg.X1 = cx1[i];
                punct.Add(pg);
            }
        }
        float[] inkOwn = new float[ink.Length];
        for (int i = 0; i < ink.Length; i++) inkOwn[i] = lab[i] != 0 ? (own[lab[i]] ? ink[i] : 0) : Math.Min(ink[i], 0.44f);
        // разрезы: по столбцам с краской - промежутки между буквами и провалы внутри слипшихся букв
        double[] col = new double[L];
        int r0 = Math.Max(0, (int)(mid - 0.55 * capS)), r1 = Math.Min(SH - 1, (int)(mid + 0.55 * capS));
        for (int x = 0; x < L; x++) {
            double sum = 0;
            for (int y = 0; y < SH; y++) {
                float v = inkOwn[y * L + x];
                if (v > 0.45f) sum += y >= r0 && y <= r1 ? v : 0.01;
            }
            col[x] = sum;
        }
        List<int> cuts = new List<int>();          // x разрезов по порядку
        List<bool> gapAfter = new List<bool>();    // от этого разреза до следующего - пусто (промежуток)
        List<int> gapW = new List<int>();
        List<bool> inside = new List<bool>();      // разрез посреди краски (а не в промежутке)
        int xs = 0;
        while (xs < L && col[xs] <= 0) xs++;
        while (xs < L) {
            int xe = xs;
            while (xe < L && col[xe] > 0) xe++;             // краска: [xs, xe)
            cuts.Add(xs); gapAfter.Add(false); gapW.Add(0); inside.Add(false);
            int w = xe - xs;
            if (w > 0.8 * capS) {
                double mx = 0;
                for (int x = xs; x < xe; x++) mx = Math.Max(mx, col[x]);
                int minGap = Math.Max(2, (int)(0.18 * capS));
                for (int x = xs + minGap; x < xe - minGap; x++) {
                    if (col[x] > 0.35 * mx) continue;
                    bool lm = true;
                    for (int k = Math.Max(xs, x - minGap / 2); k <= Math.Min(xe - 1, x + minGap / 2); k++)
                        if (col[k] < col[x] || (col[k] == col[x] && k < x)) lm = false;
                    if (lm) { cuts.Add(x); gapAfter.Add(false); gapW.Add(0); inside.Add(true); }
                }
            }
            int nx = xe;
            while (nx < L && col[nx] <= 0) nx++;
            cuts.Add(xe);                                     // дальше - промежуток до следующей краски
            inside.Add(false);
            gapAfter.Add(true);
            gapW.Add(nx < L ? nx - xe : 0);
            xs = nx;
        }
        if (cuts.Count < 2) return null;
        Strip st = new Strip();
        st.Ink = inkOwn; st.L = L; st.SH = SH; st.Mid = mid; st.Cap = capS;
        st.Cuts = cuts; st.Gap = gapAfter; st.GapW = gapW; st.Inside = inside; st.Set = set; st.Mixed = mixed;
        if (mixed) {
            st.Lab = lab; st.Own = own; st.CX0 = cx0; st.CX1 = cx1; st.Raw = ink;
            st.Cen = new double[count + 1];
            for (int i = 1; i <= count; i++) st.Cen[i] = cn[i] > 0 ? csx[i] / cn[i] : 0;
        }
        st.G = Decode(st, null);
        st.Punct = punct;
        if (punct.Count > 0) st.G = WithPunct(st.G, punct, capS);
        return st;
    }

    // строка, готовая к делению на знаки: полоса, разрезы, оценки уже проверенных кусков
    internal class Strip {
        public float[] Ink;
        public int L, SH;
        public double Mid, Cap;
        public List<int> Cuts, GapW;
        public List<bool> Gap, Inside;
        public bool Mixed;
        public List<Glyph> Punct;
        public int[] Lab, CX0, CX1;           // обычный текст: буква берётся целиком своей областью краски
        public bool[] Own;
        public double[] Cen;
        public float[] Raw;
        public GlyphSet Set;
        public Dictionary<long, Glyph> Cache = new Dictionary<long, Glyph>();
        public List<Glyph> G;
    }

    // точки и запятые - на свои места среди букв; пробелы - заново по промежуткам
    static List<Glyph> WithPunct(List<Glyph> gs, List<Glyph> punct, double cap) {
        List<Glyph> all = new List<Glyph>(gs);
        foreach (Glyph p in punct) {
            Glyph c = new Glyph();
            c.C = p.C; c.S = p.S; c.C2 = p.C2; c.S2 = p.S2; c.Sc = p.Sc; c.X0 = p.X0; c.X1 = p.X1;
            all.Add(c);
        }
        all.Sort(delegate(Glyph a, Glyph b) { return (a.X0 + a.X1).CompareTo(b.X0 + b.X1); });
        for (int i = 0; i < all.Count; i++) all[i].Space = i > 0 && all[i].X0 - all[i - 1].X1 > 0.33 * cap;
        return all;
    }

    // какие знаки могут стоять на этом месте строки (null - любые)
    internal delegate string Alphabet(int x0, int x1);

    const string ByLetters = "АБВГДЕЁЖЗІЙКЛМНОПРСТУЎФХЦЧШЫЬЭЮЯ-";      // в белорусском нет И, Щ, Ъ
    const string RuLetters = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ-";     // в русском нет І, Ў

    static Glyph Piece(Strip st, int i, int j) {
        long key = (long)i * 100000 + j;
        Glyph g;
        if (!st.Cache.TryGetValue(key, out g)) {
            g = st.Mixed ? ClassifyWhole(st, st.Cuts[i], st.Cuts[j]) : Classify(st.Ink, st.L, st.SH, st.Cuts[i], st.Cuts[j], st.Mid, st.Cap, st.Set);
            st.Cache[key] = g;
        }
        return g;
    }

    // знак с выбором только из разрешённых (копия - оценки общие)
    // деление по разрезу посреди краски и склейка через промежуток бывают реже - небольшой штраф
    public static double CutPenalty = 0.12, GapPenalty = 0.1;

    static double Penalty(Strip st, int ci, int cj) {
        if (!st.Mixed) return 0;          // в паспорте (жирные заглавные) буквы часто слипаются - без штрафа
        double p = st.Inside[ci] ? CutPenalty : 0;
        // узкий промежуток - внутри буквы бывает (ы, Ы), широкий - между буквами
        for (int k = ci + 1; k < cj; k++) if (st.Gap[k]) p += st.GapW[k] > 0.1 * st.Cap ? GapPenalty : 0.02;
        return p;
    }

    internal static Glyph Restrict(Glyph g, string allowed, GlyphSet set) {
        Glyph r = new Glyph();
        r.Sc = g.Sc;
        r.X0 = g.X0;
        r.X1 = g.X1;
        r.S = r.S2 = double.MinValue;
        for (int k = 0; k < set.Chars.Length; k++) {
            if (allowed != null && allowed.IndexOf(set.Chars[k]) < 0) continue;
            double v = g.Sc[k];
            if (v > r.S) { r.S2 = r.S; r.C2 = r.C; r.S = v; r.C = set.Chars[k]; }
            else if (v > r.S2) { r.S2 = v; r.C2 = set.Chars[k]; }
        }
        return r.S == double.MinValue ? null : r;
    }

    // знак - от разреза до разреза внутри куска или через узкий промежуток (Ы, разорванные буквы);
    // лучшее деление строки - где сумма (ширина x сходство) больше
    internal static List<Glyph> Decode(Strip st, Alphabet abc) {
        List<int> cuts = st.Cuts;
        double capS = st.Cap;
        int m = cuts.Count;
        double[] best2 = new double[m];
        int[] from = new int[m];
        Glyph[] got = new Glyph[m];
        for (int i = 0; i < m; i++) { best2[i] = double.MinValue; from[i] = -1; }
        best2[0] = 0;
        for (int j = 1; j < m; j++) {
            // промежуток или соринка - без знака (в обычном тексте пропуск краски - со штрафом)
            double drop = st.Mixed && !st.Gap[j - 1] ? -0.3 * (cuts[j] - cuts[j - 1]) / capS : 0;
            if (best2[j - 1] > double.MinValue && best2[j - 1] + drop > best2[j]) { best2[j] = best2[j - 1] + drop; from[j] = j - 1; got[j] = null; }
            if (st.Gap[j - 1]) continue;
            for (int i = j - 1; i >= 0; i--) {
                if (cuts[j] - cuts[i] > 1.7 * capS) break;
                if (st.Gap[i]) {
                    if (st.GapW[i] > 0.22 * capS) break;
                    continue;
                }
                if (best2[i] == double.MinValue) continue;
                Glyph raw = Piece(st, i, j);
                if (raw == null) continue;
                Glyph gl = Restrict(raw, abc == null ? null : abc(cuts[i], cuts[j]), st.Set);
                if (gl == null) continue;
                double v = best2[i] + (cuts[j] - cuts[i]) / capS * gl.S - Penalty(st, i, j);
                if (v > best2[j]) { best2[j] = v; from[j] = i; got[j] = gl; }
            }
        }
        List<Glyph> res = new List<Glyph>();
        int at = m - 1;
        while (at > 0 && from[at] >= 0) {
            if (got[at] != null) { got[at].Src = st; got[at].CutA = from[at]; got[at].CutB = at; res.Add(got[at]); }
            at = from[at];
        }
        res.Reverse();
        for (int i = 1; i < res.Count; i++) if (res[i].X0 - res[i - 1].X1 > 0.33 * capS) res[i].Space = true;
        // « / » между белорусской и русской частью: отдельный знак между пробелами, похожий на косую черту
        int sl = st.Set.Chars.IndexOf('/');
        if (sl >= 0 && abc == null)
            for (int i = 1; i + 1 < res.Count; i++)
                if (res[i].Space && res[i + 1].Space && res[i].C != '/' && res[i].Sc[sl] > res[i].S - 0.2) {
                    res[i].C2 = res[i].C; res[i].S2 = res[i].S; res[i].C = '/'; res[i].S = res[i].Sc[sl];
                }
        if (DebugDir != null) {
            Dump(st.Ink, st.L, st.SH, cuts, res, st.Set);
            StringBuilder sb = new StringBuilder("L" + dumpNo.ToString("000") + (abc != null ? " (2)" : "") + " cap=" + capS.ToString("0.0") + "\n");
            foreach (Glyph gl in res) sb.Append("   " + gl.X0 + "-" + gl.X1 + " " + gl.C + " " + gl.S.ToString("0.00") + " / " + gl.C2 + " " + gl.S2.ToString("0.00") + "\n");
            System.IO.File.AppendAllText(System.IO.Path.Combine(DebugDir, "log.txt"), sb.ToString());
        }
        return res;
    }

    static int dumpNo;
    static void Dump(float[] ink, int L, int SH, List<int> cuts, List<Glyph> res, GlyphSet set) {
        try {
            using (Bitmap b = new Bitmap(L, SH + 4, PixelFormat.Format24bppRgb)) {
                for (int y = 0; y < SH; y++)
                    for (int x = 0; x < L; x++) { int v = 255 - (int)(255 * ink[y * L + x]); b.SetPixel(x, y, Color.FromArgb(v, v, v)); }
                foreach (int c in cuts) if (c < L) for (int y = 0; y < 4; y++) b.SetPixel(c, SH + y, Color.Red);
                foreach (Glyph g in res) for (int x = g.X0; x <= g.X1 && x < L; x++) b.SetPixel(x, SH + 2, Color.Blue);
                dumpNo++;
                b.Save(System.IO.Path.Combine(DebugDir, "L" + dumpNo.ToString("000") + (set == sans ? "s" : "o") + ".png"), ImageFormat.Png);
            }
        } catch { }
    }

    // обычный текст: буквы с засечками заходят друг на друга - область краски относится к куску, где её середина,
    // и берётся целиком; режется по столбцам только большая область (несколько слипшихся букв)
    static Glyph ClassifyWhole(Strip st, int xa, int xb) {
        int L = st.L, SH = st.SH;
        int wa = Math.Max(0, (int)(xa - 0.6 * st.Cap)), wb = Math.Min(L, (int)(xb + 0.6 * st.Cap));
        int ww = wb - wa;
        if (ww <= 0) return null;
        float[] m = new float[ww * SH];
        bool any = false;
        for (int y = 0; y < SH; y++)
            for (int x = wa; x < wb; x++) {
                int i = y * L + x, l = st.Lab[i];
                bool take;
                if (l == 0) take = x >= xa && x < xb;
                else if (!st.Own[l]) take = false;
                else if (st.CX1[l] - st.CX0[l] + 1 > 1.2 * st.Cap) take = x >= xa && x < xb;
                else take = st.Cen[l] >= xa && st.Cen[l] < xb;
                if (!take) continue;
                m[y * ww + x - wa] = l == 0 ? Math.Min(st.Raw[i], 0.44f) : st.Raw[i];
                if (l != 0) any = true;
            }
        if (!any) return null;
        Glyph g = Classify(m, ww, SH, 0, ww, st.Mid, st.Cap, st.Set);
        if (g != null) { g.X0 += wa; g.X1 += wa; }
        return g;
    }

    static Glyph Classify(float[] ink, int L, int SH, int xa, int xb, double mid, double cap, GlyphSet set) {
        int bx0 = L, by0 = SH, bx1 = -1, by1 = -1;
        double mass = 0;
        for (int y = 0; y < SH; y++)
            for (int x = xa; x < xb; x++) {
                float v = ink[y * L + x];
                if (v <= 0.45f) continue;
                mass += v;
                if (x < bx0) bx0 = x;
                if (x > bx1) bx1 = x;
                if (y < by0) by0 = y;
                if (y > by1) by1 = y;
            }
        if (bx1 < 0 || mass < 0.02 * cap * cap) return null;
        float[] grid = set.Grid(ink, L, bx0, by0, bx1, by1);
        double h = (by1 - by0 + 1) / cap, yy = ((by0 + by1) / 2.0 - mid) / cap, asp = (bx1 - bx0 + 1.0) / (by1 - by0 + 1);
        double[] sc = set.Score(grid, h, yy, asp);
        Glyph g = new Glyph();
        g.Sc = sc;
        g.S = g.S2 = double.MinValue;
        for (int k = 0; k < sc.Length; k++) {
            if (sc[k] > g.S) { g.S2 = g.S; g.C2 = g.C; g.S = sc[k]; g.C = set.Chars[k]; }
            else if (sc[k] > g.S2) { g.S2 = sc[k]; g.C2 = set.Chars[k]; }
        }
        g.X0 = bx0;
        g.X1 = bx1;
        return g;
    }

    internal static string Text(List<Glyph> gs) {
        StringBuilder sb = new StringBuilder();
        foreach (Glyph g in gs) { if (g.Space) sb.Append(' '); sb.Append(g.C); }
        return sb.ToString();
    }

    // ---------- страница ----------

    public static PageResult Read(string path) {
        Init();
        GrayImage full, small;
        Mrz.LoadPair(path, out full, out small);
        // строки ищем на уменьшенном снимке, читаем по полному; текст поперёк снимка - повернуть,
        // вверх ногами - решат оценки букв
        List<Line> l0 = FindLines(small);
        GrayImage sr = small.Rotate90();
        List<Line> l1 = FindLines(sr);
        if (Weight(l1) > Weight(l0)) { small = sr; l0 = l1; full = full.Rotate90(); }
        double f = (double)full.W / small.W;
        Parsed best = Parse(full, Scale(l0, f));
        if (!best.FioOk) {
            GrayImage ff = Rotate180(full), sf = Rotate180(small);
            Parsed other = Parse(ff, Scale(FindLines(sf), f));
            if (other.Score > best.Score) best = other;
        }
        return best.Result;
    }

    internal static List<Line> Scale(List<Line> lines, double f) {
        foreach (Line l in lines) {
            l.Comps = Mrz.Scale(l.Comps, f);
            l.Cap *= f;
            l.Y *= f;
            l.X0 = (int)(l.X0 * f);
        }
        return lines;
    }

    internal static int Weight(List<Line> lines) {
        int w = 0;
        foreach (Line l in lines) w += l.Comps.Count;
        return w;
    }

    internal static GrayImage Rotate180(GrayImage g) {
        GrayImage r = new GrayImage(g.W, g.H);
        int n = g.P.Length;
        for (int i = 0; i < n; i++) r.P[n - 1 - i] = g.P[i];
        return r;
    }

    class Parsed {
        public PageResult Result = new PageResult();
        public double Score;
        public List<Field> Fields;
        public bool FioOk { get { return Result.FioOk; } }
    }

    // строка «ПА-БЕЛАРУСКУ / ПО-РУССКИ»: слова до и после отдельно стоящей косой черты
    class Field {
        public List<Glyph> By = new List<Glyph>(), Ru = new List<Glyph>();
        public Glyph Slash;
        public Strip St;
        public int RuFrom, RuTo, ByFrom, ByTo;
        public string ByText = "", RuText = "";
    }

    static bool IsLetter(char c) { return c >= 'А' && c <= 'Я' || c == 'Ё' || c == 'І' || c == 'Ў' || c == '-'; }

    // знак внутри слова, прочитанный как цифра или черта, - лучшая из букв
    static void AsLetter(Glyph g) {
        if (IsLetter(g.C)) return;
        double best = double.MinValue;
        char bc = g.C;
        for (int k = 0; k < SansChars.Length; k++)
            if (IsLetter(SansChars[k]) && g.Sc[k] > best) { best = g.Sc[k]; bc = SansChars[k]; }
        g.C2 = g.C; g.S2 = g.S;
        g.C = bc; g.S = best;
    }

    static Field Split(List<Glyph> gs) {
        List<List<Glyph>> words = new List<List<Glyph>>();
        foreach (Glyph g in gs) {
            if (g.Space || words.Count == 0) words.Add(new List<Glyph>());
            words[words.Count - 1].Add(g);
        }
        int sep = -1;
        for (int i = 1; i < words.Count - 1; i++)
            if (words[i].Count == 1 && words[i][0].C == '/') {
                if (sep < 0 || Math.Abs(i - words.Count / 2.0) < Math.Abs(sep - words.Count / 2.0)) sep = i;
            }
        if (sep < 0) return null;
        // строка значения: почти одни буквы, и читаются они хорошо (заголовок и подписи - другим шрифтом)
        int n = 0, let = 0;
        double mean = 0;
        foreach (Glyph g in gs) { if (g.C == '/') continue; n++; mean += g.S; if (IsLetter(g.C)) let++; }
        if (n == 0 || let < 0.7 * n || mean / n < 0.6) return null;
        Field f = new Field();
        f.Slash = words[sep][0];
        for (int i = 0; i < words.Count; i++) {
            if (i == sep) continue;
            List<Glyph> to = i < sep ? f.By : f.Ru;
            foreach (Glyph g in words[i]) {
                AsLetter(g);
                if (to.Count > 0 && g == words[i][0]) g.Space = true;
                else g.Space = false;
                to.Add(g);
            }
        }
        f.ByText = Text(f.By);
        f.RuText = Text(f.Ru);
        return f;
    }

    static Parsed Parse(GrayImage g, List<Line> lines) {
        Parsed p = new Parsed();
        PageResult r = p.Result;
        List<Field> named = new List<Field>();
        double total = 0;
        int cnt = 0;
        foreach (Line l in lines) {
            Strip st = ReadLine(g, l, sans);
            if (st == null || st.G.Count == 0) continue;
            l.G = st.G;
            l.Text = Text(l.G);
            foreach (Glyph gl in l.G) { total += gl.S; cnt++; }
            r.Lines.Add(l.Text);
            ReadNumber(g, l, r);
            if (l.Text.IndexOf(',') >= 0) continue;                  // место рождения, орган
            Field f = Split(l.G);
            if (f == null || f.By.Count < 2 || f.Ru.Count < 2) continue;
            // ещё раз, зная, где какой язык: слева - только белорусские буквы, справа - только русские
            int sx0 = f.Slash.X0, sx1 = f.Slash.X1;
            Field f2 = Split(Decode(st, delegate(int a, int b) {
                if (b <= sx0 + 1) return ByLetters;
                if (a >= sx1) return RuLetters;
                if (a >= sx0 - 2 && b <= sx1 + 3) return "/";
                return "";
            }));
            if (f2 != null && f2.By.Count >= 2 && f2.Ru.Count >= 2) f = f2;
            f.St = st;
            f.RuTo = st.Cuts.Count - 1;
            f.RuFrom = f.RuTo;
            for (int k = 0; k < st.Cuts.Count; k++) if (st.Cuts[k] > f.Slash.X1) { f.RuFrom = k; break; }
            f.ByFrom = 0;
            f.ByTo = 0;
            for (int k = 0; k < st.Cuts.Count; k++) if (st.Cuts[k] <= f.Slash.X0) f.ByTo = k;
            if (Similar(f.RuText, "РЕСПУБЛИКА БЕЛАРУСЬ") || Similar(f.ByText, "РЭСПУБЛІКА БЕЛАРУСЬ")) continue;
            named.Add(f);
        }
        p.Score = cnt > 0 ? total / cnt : 0;
        // в номере - дата рождения (ДДММГГ): с датой на странице не сходится - номер под сомнением
        if (r.IdOk && r.Birth.Length == 10 && r.Id.Substring(1, 6) != r.Birth.Substring(0, 2) + r.Birth.Substring(3, 2) + r.Birth.Substring(8, 2)) {
            r.IdOk = false;
            r.Info += "номер не сходится с датой рождения; ";
        }
        if (named.Count >= 2) {
            r.Found = true;
            // какая строка - отчество: та, что лучше всех похожа на отчество из списка (обычно третья)
            int pi = -1;
            string pword = null;
            for (int k = 1; k < named.Count && k <= 3; k++) {
                double loss, margin;
                Field fk = named[k];
                string w = Match(fk.St, fk.RuFrom, fk.RuTo, fk.Ru.Count, RuLetters, NamesData.Patronymics, out loss, out margin);
                r.Lines.Add("  отчество? " + w + " loss=" + loss.ToString("0.000") + " margin=" + margin.ToString("0.000"));
                if (w != null && loss <= LossMax && margin >= MarginMin) { pi = k; pword = w; break; }
            }
            if (pi < 0 && named.Count >= 3) pi = 2;
            Field pat = pi >= 0 ? named[pi] : null;
            Field nm = pi >= 1 ? named[pi - 1] : named[1];
            Field sur = pi >= 2 ? named[pi - 2] : (pi < 0 ? named[0] : null);
            bool female = false, known = pat != null;
            if (pat != null) {
                r.PatronymicBy = pat.ByText;
                r.Patronymic = pat.RuText;
                if (pword != null) { r.Patronymic = pword; r.PatronymicOk = true; }
                female = r.Patronymic.EndsWith("НА");
            }
            r.NameBy = nm.ByText;
            r.Name = nm.RuText;
            {
                double loss, margin;
                List<string> list = !known ? Both() : female ? NamesData.FemaleNames : NamesData.MaleNames;
                string w = Match(nm.St, nm.RuFrom, nm.RuTo, nm.Ru.Count, RuLetters, list, out loss, out margin);
                r.Lines.Add("  имя? " + w + " loss=" + loss.ToString("0.000") + " margin=" + margin.ToString("0.000"));
                // имени нет в списке - остаётся прочитанное (но без уверенности)
                if (w != null && loss <= LossMax) {
                    r.Name = w;
                    r.NameKnown = true;
                    r.NameOk = margin >= MarginMin;
                }
            }
            if (sur != null) {
                r.SurnameBy = sur.ByText;
                r.Surname = sur.RuText;
                r.SurnameOk = Sure(sur.Ru) && Sure(sur.By) && Skeleton(sur.ByText) == Skeleton(sur.RuText);
                if (Skeleton(sur.ByText) != Skeleton(sur.RuText)) {
                    // написания расходятся по согласным: одна из частей прочитана с ошибкой. Пробуем прочесть
                    // каждую часть по «скелету» другой - какая при этом теряет меньше, та и ошиблась
                    double inkR, inkB;
                    double fr = Free(sur.St, sur.RuFrom, sur.RuTo, RuLetters, out inkR), fb = Free(sur.St, sur.ByFrom, sur.ByTo, ByLetters, out inkB);
                    string ru2, by2;
                    double vr = Pattern(sur.St, sur.RuFrom, sur.RuTo, Units(Skeleton(sur.ByText), false), out ru2);
                    double vb = Pattern(sur.St, sur.ByFrom, sur.ByTo, Units(Skeleton(sur.RuText), true), out by2);
                    double lr = ru2 == null ? 9 : (fr - vr) / Math.Max(inkR, 0.5), lb = by2 == null ? 9 : (fb - vb) / Math.Max(inkB, 0.5);
                    r.Lines.Add("  фамилия по скелету: " + ru2 + " " + lr.ToString("0.000") + " | " + by2 + " " + lb.ToString("0.000"));
                    if (lr < lb && lr <= 0.06) r.Surname = ru2;
                }
                r.Lines.Add("  фамилия: " + Weakest(sur.Ru) + " | " + Weakest(sur.By) + " | " + Skeleton(sur.ByText) + " " + Skeleton(sur.RuText));
            }
            p.Fields = named;
        }
        return p;
    }

    public static double LossMax = 0.06, MarginMin = 0.03;

    static List<string> Both() {
        List<string> l = new List<string>(NamesData.FemaleNames);
        l.AddRange(NamesData.MaleNames);
        return l;
    }

    static string Weakest(List<Glyph> gs) {
        double ms = 9, mm = 9;
        foreach (Glyph g in gs) {
            double second = double.MinValue;
            for (int k = 0; k < SansChars.Length; k++)
                if (SansChars[k] != g.C && IsLetter(SansChars[k]) && g.Sc[k] > second) second = g.Sc[k];
            ms = Math.Min(ms, g.S);
            mm = Math.Min(mm, g.S - second);
        }
        return "minS=" + ms.ToString("0.00") + " minM=" + mm.ToString("0.00");
    }

    // все буквы прочитаны уверенно
    static bool Sure(List<Glyph> gs) {
        foreach (Glyph g in gs) {
            if (g.S < 0.55) return false;
            double second = double.MinValue;
            for (int k = 0; k < SansChars.Length; k++)
                if (SansChars[k] != g.C && IsLetter(SansChars[k]) && g.Sc[k] > second) second = g.Sc[k];
            if (g.S - second < 0.06) return false;
        }
        return true;
    }

    // «скелет» слова: белорусское и русское написание фамилии должны совпасть по согласным
    // (гласные и мягкий знак в двух языках пишутся по-разному: ЛЯШЧЭНКА / ЛЕЩЕНКО, ІВАНОЎ / ИВАНОВ)
    public static string Skeleton(string s) {
        StringBuilder sb = new StringBuilder();
        string t = s.Replace("ДЗ", "Д").Replace("ШЧ", "Щ");
        foreach (char c in t) {
            if ("АОЕЁЭЯИЫІЙ".IndexOf(c) >= 0) { if (sb.Length == 0 || sb[sb.Length - 1] != 'V') sb.Append('V'); }
            else if (c == 'Ь' || c == 'Ъ' || c == '\'' || c == ' ') continue;
            else if (c == 'Ў') sb.Append('В');
            else if (c == 'Ц') sb.Append('Т');
            else sb.Append(c);
        }
        return sb.ToString();
    }

    static void Part(Line l, out string by, out string ru) {
        int slash = l.Text.IndexOf('/');
        by = l.Text.Substring(0, slash).Trim();
        ru = l.Text.Substring(slash + 1).Trim();
    }

    static double Letters(string s) {
        int n = 0, a = 0;
        foreach (char c in s) { if (c == ' ') continue; n++; if (char.IsLetter(c) || c == '-') a++; }
        return n == 0 ? 0 : (double)a / n;
    }

    static bool Similar(string a, string b) { return Distance(a, b) <= Math.Max(3, b.Length / 4); }

    public static int Distance(string a, string b) {
        int[,] d = new int[a.Length + 1, b.Length + 1];
        for (int i = 0; i <= a.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= b.Length; j++) d[0, j] = j;
        for (int i = 1; i <= a.Length; i++)
            for (int j = 1; j <= b.Length; j++)
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1));
        return d[a.Length, b.Length];
    }

    // ---------- прочтение по образцу ----------

    // лучшее прочтение участка строки (разрезы from..to) ровно по образцу: sets[t] - какие знаки могут стоять на месте t.
    // Возвращает сумму (ширина x сходство); пропуск краски (соринка) - штраф.
    internal static double Constrained(Strip st, int from, int to, string[] sets, out string text) {
        int m = to - from + 1, n = sets.Length;
        double[,] f = new double[m, n + 1];
        int[,] bi = new int[m, n + 1];
        char[,] bc = new char[m, n + 1];
        for (int k = 0; k < m; k++) for (int t = 0; t <= n; t++) { f[k, t] = double.MinValue; bi[k, t] = -1; }
        f[0, 0] = 0;
        List<int> cuts = st.Cuts;
        for (int k = 1; k < m; k++) {
            int cj = from + k;
            bool gap = st.Gap[cj - 1];
            double skip = gap ? 0 : -0.3 * (cuts[cj] - cuts[cj - 1]) / st.Cap;
            for (int t = 0; t <= n; t++)
                if (f[k - 1, t] > double.MinValue && f[k - 1, t] + skip > f[k, t]) { f[k, t] = f[k - 1, t] + skip; bi[k, t] = k - 1; bc[k, t] = '\0'; }
            if (gap) continue;
            for (int i = k - 1; i >= 0; i--) {
                int ci = from + i;
                if (cuts[cj] - cuts[ci] > 1.7 * st.Cap) break;
                if (st.Gap[ci]) {
                    if (st.GapW[ci] > 0.22 * st.Cap) break;
                    continue;
                }
                Glyph raw = Piece(st, ci, cj);
                if (raw == null) continue;
                double w = (cuts[cj] - cuts[ci]) / st.Cap;
                for (int t = 1; t <= n; t++) {
                    if (f[i, t - 1] == double.MinValue) continue;
                    double bs = double.MinValue;
                    char c = '\0';
                    foreach (char ch in sets[t - 1]) {
                        int q = st.Set.Chars.IndexOf(ch);
                        if (q >= 0 && raw.Sc[q] > bs) { bs = raw.Sc[q]; c = ch; }
                    }
                    if (c == '\0') continue;
                    double v = f[i, t - 1] + w * bs - Penalty(st, ci, cj);
                    if (v > f[k, t]) { f[k, t] = v; bi[k, t] = i; bc[k, t] = c; }
                }
            }
        }
        text = null;
        if (f[m - 1, n] == double.MinValue) return double.MinValue;
        char[] res = new char[n];
        int kk = m - 1, tt = n;
        while (kk > 0) {
            char c = bc[kk, tt];
            int pk = bi[kk, tt];
            if (c != '\0') { res[tt - 1] = c; tt--; }
            kk = pk;
        }
        text = new string(res);
        return f[m - 1, n];
    }

    // лучшее прочтение участка без образца (любое число знаков из алфавита) и общая ширина краски
    internal static double Free(Strip st, int from, int to, string alphabet, out double inkWidth) {
        int m = to - from + 1;
        double[] f = new double[m];
        for (int k = 0; k < m; k++) f[k] = double.MinValue;
        f[0] = 0;
        inkWidth = 0;
        List<int> cuts = st.Cuts;
        for (int k = 1; k < m; k++) {
            int cj = from + k;
            bool gap = st.Gap[cj - 1];
            if (!gap) inkWidth += (cuts[cj] - cuts[cj - 1]) / st.Cap;
            double skip = gap ? 0 : -0.3 * (cuts[cj] - cuts[cj - 1]) / st.Cap;
            if (f[k - 1] > double.MinValue) f[k] = f[k - 1] + skip;
            if (gap) continue;
            for (int i = k - 1; i >= 0; i--) {
                int ci = from + i;
                if (cuts[cj] - cuts[ci] > 1.7 * st.Cap) break;
                if (st.Gap[ci]) {
                    if (st.GapW[ci] > 0.22 * st.Cap) break;
                    continue;
                }
                if (f[i] == double.MinValue) continue;
                Glyph raw = Piece(st, ci, cj);
                if (raw == null) continue;
                Glyph gl = Restrict(raw, alphabet, st.Set);
                if (gl == null) continue;
                double v = f[i] + (cuts[cj] - cuts[ci]) / st.Cap * gl.S - Penalty(st, ci, cj);
                if (v > f[k]) f[k] = v;
            }
        }
        return f[m - 1];
    }

    // слово из списка, лучше всего подходящее к участку строки: loss - насколько (на ширину буквы) хуже свободного
    // прочтения, margin - насколько лучше следующего слова из списка
    static string Match(Strip st, int from, int to, int glyphs, string alphabet, List<string> words, out double loss, out double margin) {
        double ink;
        double free = Free(st, from, to, alphabet, out ink);
        string best = null;
        double b1 = double.MinValue, b2 = double.MinValue;
        foreach (string w in words) {
            if (Math.Abs(w.Length - glyphs) > 2) continue;
            string[] sets = new string[w.Length];
            for (int i = 0; i < w.Length; i++) sets[i] = w[i].ToString();
            string text;
            double v = Constrained(st, from, to, sets, out text);
            if (v > b1) { b2 = b1; b1 = v; best = w; }
            else if (v > b2) b2 = v;
        }
        ink = Math.Max(ink, 0.5);
        loss = best == null ? 9 : (free - b1) / ink;
        margin = b2 == double.MinValue ? 9 : (b1 - b2) / ink;
        return best;
    }

    // ---------- сверка с латиницей строк «<<<» (последняя страница того же паспорта) ----------

    // как белорусские буквы пишутся латиницей в паспорте (бывает по-разному - все варианты)
    static string[] Lat(char c) {
        switch (c) {
            case 'А': return new string[] { "A" };
            case 'Б': return new string[] { "B" };
            case 'В': return new string[] { "V" };
            case 'Г': return new string[] { "H", "G" };
            case 'Д': return new string[] { "D" };
            case 'Е': return new string[] { "E", "IE", "YE" };
            case 'Ё': return new string[] { "IO", "YO", "E", "O" };
            case 'Ж': return new string[] { "ZH" };
            case 'З': return new string[] { "Z" };
            case 'І': return new string[] { "I", "Y" };
            case 'И': return new string[] { "I" };
            case 'Й': return new string[] { "I", "Y", "J", "" };
            case 'К': return new string[] { "K" };
            case 'Л': return new string[] { "L" };
            case 'М': return new string[] { "M" };
            case 'Н': return new string[] { "N" };
            case 'О': return new string[] { "O" };
            case 'П': return new string[] { "P" };
            case 'Р': return new string[] { "R" };
            case 'С': return new string[] { "S" };
            case 'Т': return new string[] { "T" };
            case 'У': return new string[] { "U" };
            case 'Ў': return new string[] { "U", "W" };
            case 'Ф': return new string[] { "F" };
            case 'Х': return new string[] { "KH", "H", "X" };
            case 'Ц': return new string[] { "TS", "C" };
            case 'Ч': return new string[] { "CH" };
            case 'Ш': return new string[] { "SH" };
            case 'Ы': return new string[] { "Y", "I" };
            case 'Ь': return new string[] { "" };
            case 'Ъ': return new string[] { "" };
            case 'Э': return new string[] { "E" };
            case 'Ю': return new string[] { "IU", "YU", "U" };
            case 'Я': return new string[] { "IA", "YA", "A" };
            case '-': return new string[] { " ", "" };
            case ' ': return new string[] { " ", "" };
        }
        return new string[] { "" };
    }

    // может ли кириллица дать эту латиницу (пробелы в латинице - на месте «<»)
    public static bool LatinMatch(string cyr, string lat) {
        lat = lat.Replace('<', ' ').Trim();
        int n = cyr.Length, m = lat.Length;
        if (n == 0 || m == 0) return false;
        bool[,] ok = new bool[n + 1, m + 1];
        ok[0, 0] = true;
        for (int i = 0; i < n; i++)
            for (int j = 0; j <= m; j++) {
                if (!ok[i, j]) continue;
                foreach (string t in Lat(cyr[i]))
                    if (j + t.Length <= m && string.CompareOrdinal(lat, j, t, 0, t.Length) == 0) ok[i + 1, j + t.Length] = true;
            }
        return ok[n, m];
    }

    // строки «<<<» того же паспорта подтверждают прочитанное (фамилия и имя там - белорусские, латиницей)
    public static void Confirm(PageResult r, MrzResult m) {
        if (r == null || m == null || !r.Found || m.Line2.Length != 44 || m.Line2.Substring(10, 3) != "BLR") return;
        // тот же паспорт: номер на странице и в строках совпал (если оба прочитаны)
        if (r.IdOk && m.PersonalOk && r.Id != m.Personal) return;
        if (!r.SurnameOk && r.SurnameBy.Length > 0 && LatinMatch(r.SurnameBy, m.Surname) && Skeleton(r.SurnameBy) == Skeleton(r.Surname)) {
            r.SurnameOk = true;
            r.Info += "фамилия сошлась с латиницей паспорта; ";
        }
        string first = m.Names.Split(' ')[0];
        if (!r.NameOk && r.NameBy.Length > 0 && first.Length > 0 && LatinMatch(r.NameBy, first) && r.NameKnown) {
            r.NameOk = true;
            r.Info += "имя сошлось с латиницей паспорта; ";
        }
    }

    // ---------- чтение по «скелету» (согласные белорусского и русского написания совпадают) ----------

    const string RuVowels = "АОЕЁЭЯИЫЙ", ByVowels = "АОЕЁЭЯЫІЙ";

    // скелет -> образец: на месте гласных - 1..3 гласные, после согласной может стоять мягкий знак
    static List<List<string[]>> Units(string skel, bool by) {
        List<List<string[]>> units = new List<List<string[]>>();
        string v = by ? ByVowels : RuVowels;
        foreach (char c in skel) {
            List<string[]> alts = new List<string[]>();
            if (c == 'V') {
                alts.Add(new string[] { v });
                alts.Add(new string[] { v, v });
                alts.Add(new string[] { v, v, v });
            } else {
                List<string[]> core = new List<string[]>();
                if (c == 'В') core.Add(new string[] { by ? "ВЎ" : "В" });
                else if (c == 'Т') core.Add(new string[] { "ТЦ" });
                else if (c == 'Щ') core.Add(by ? new string[] { "Ш", "Ч" } : new string[] { "Щ" });
                else if (c == 'Д') { core.Add(new string[] { "Д" }); core.Add(new string[] { "Д", "З" }); }
                else core.Add(new string[] { c.ToString() });
                foreach (string[] a in core) {
                    alts.Add(a);
                    string[] soft = new string[a.Length + 1];
                    a.CopyTo(soft, 0);
                    soft[a.Length] = "Ь";
                    alts.Add(soft);
                }
            }
            units.Add(alts);
        }
        return units;
    }

    // лучшее прочтение участка строки по образцу из блоков (у блока - варианты, вариант - 1..3 знака)
    static double Pattern(Strip st, int from, int to, List<List<string[]>> units, out string text) {
        int m = to - from + 1, n = units.Count;
        double[,] f = new double[m, n + 1];
        int[,] pk = new int[m, n + 1], pu = new int[m, n + 1];
        string[,] pc = new string[m, n + 1];
        for (int k = 0; k < m; k++) for (int u = 0; u <= n; u++) { f[k, u] = double.MinValue; pk[k, u] = -1; }
        f[0, 0] = 0;
        List<int> cuts = st.Cuts;
        for (int k = 0; k < m; k++) {
            if (k > 0)
                for (int u = 0; u <= n; u++) {
                    // пропуск промежутка (или соринки - со штрафом)
                    if (f[k - 1, u] == double.MinValue) continue;
                    int cj = from + k;
                    double skip = st.Gap[cj - 1] ? 0 : -0.3 * (cuts[cj] - cuts[cj - 1]) / st.Cap;
                    if (f[k - 1, u] + skip > f[k, u]) { f[k, u] = f[k - 1, u] + skip; pk[k, u] = k - 1; pu[k, u] = u; pc[k, u] = ""; }
                }
            for (int u = 0; u < n; u++) {
                if (f[k, u] == double.MinValue) continue;
                foreach (string[] alt in units[u]) Chains(st, from, k, m, alt, 0, f[k, u], "", u, k, f, pk, pu, pc);
            }
        }
        text = null;
        if (f[m - 1, n] == double.MinValue) return double.MinValue;
        StringBuilder sb = new StringBuilder();
        int kk = m - 1, uu = n;
        while (kk > 0 || uu > 0) {
            if (pk[kk, uu] < 0) break;
            sb.Insert(0, pc[kk, uu]);
            int nk = pk[kk, uu], nu = pu[kk, uu];
            kk = nk;
            uu = nu;
        }
        text = sb.ToString();
        return f[m - 1, n];
    }

    // все цепочки знаков варианта alt от разреза k; концы - в f[.., u+1]
    static void Chains(Strip st, int from, int k, int m, string[] alt, int pos, double acc, string chars, int u, int k0,
            double[,] f, int[,] pk, int[,] pu, string[,] pc) {
        if (pos == alt.Length) {
            if (acc > f[k, u + 1]) { f[k, u + 1] = acc; pk[k, u + 1] = k0; pu[k, u + 1] = u; pc[k, u + 1] = chars; }
            return;
        }
        List<int> cuts = st.Cuts;
        int ci = from + k;
        if (ci >= cuts.Count - 1 || st.Gap[ci]) return;
        for (int j = k + 1; j < m; j++) {
            int cj = from + j;
            if (cuts[cj] - cuts[ci] > 1.7 * st.Cap) break;
            if (st.Gap[cj - 1]) {
                if (st.GapW[cj - 1] > 0.22 * st.Cap) break;
                continue;
            }
            Glyph raw = Piece(st, ci, cj);
            if (raw == null) continue;
            double bs = double.MinValue;
            char bc = '\0';
            foreach (char c in alt[pos]) {
                int q = st.Set.Chars.IndexOf(c);
                if (q >= 0 && raw.Sc[q] > bs) { bs = raw.Sc[q]; bc = c; }
            }
            if (bc == '\0') continue;
            Chains(st, from, j, m, alt, pos + 1, acc + (cuts[cj] - cuts[ci]) / st.Cap * bs - Penalty(st, ci, cj), chars + bc, u, k0, f, pk, pu, pc);
        }
    }

    // идентификационный номер (OCR-B, 14 знаков, последний - контрольная цифра) и дата рождения «ДД ММ ГГГГ»
    static void ReadNumber(GrayImage g, Line l, PageResult r) {
        if (l.Comps.Count < 7 || l.Comps.Count > 20) return;
        int digits = 0;
        foreach (char c in l.Text) if (char.IsDigit(c) || c == 'З' || c == 'О') digits++;
        if (digits < 4) return;
        Strip st = ReadLine(g, l, ocrb);
        if (st == null) return;
        int m = st.Cuts.Count - 1;
        double ink;
        double free = Free(st, 0, m, null, out ink);
        ink = Math.Max(ink, 0.5);
        if (!r.IdOk && st.G.Count >= 12 && st.G.Count <= 17) {
            string[] sets = new string[14];
            for (int i = 0; i < 14; i++) sets[i] = Allowed(i);
            string id;
            double v = Constrained(st, 0, m, sets, out id);
            // у граждан РБ в конце - «PB»: если так тоже читается почти без потерь - так и есть
            if (id != null && id.Substring(11, 2) != "PB") {
                sets[11] = "P"; sets[12] = "B";
                string pb;
                double v2 = Constrained(st, 0, m, sets, out pb);
                if (pb != null && (v - v2) / ink < 0.05 && Mrz.Check(pb.Substring(0, 13)) == pb[13]) { id = pb; v = v2; }
            }
            if (id != null && (free - v) / ink <= 0.08 && Mrz.Check(id.Substring(0, 13)) == id[13] && "123456".IndexOf(id[0]) >= 0) {
                r.Id = id;
                r.IdOk = true;
            }
        }
        if (r.Birth.Length == 0 && st.G.Count >= 7 && st.G.Count <= 10) {
            string[] sets = new string[8];
            for (int i = 0; i < 8; i++) sets[i] = "0123456789";
            string d;
            double v = Constrained(st, 0, m, sets, out d);
            DateTime dt;
            if (d != null && (free - v) / ink <= 0.08 &&
                DateTime.TryParseExact(d, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt))
                r.Birth = dt.ToString("dd.MM.yyyy");
        }
    }

    static string Allowed(int i) {
        if (i <= 6) return "0123456789";
        if (i == 7) return "ABCEHKM";
        if (i <= 10) return "0123456789";
        if (i <= 12) return "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        return "0123456789";
    }

    static char Best(Glyph g, string allowed) {
        char best = allowed[0];
        double bs = double.MinValue;
        foreach (char c in allowed) {
            int k = OcrbChars.IndexOf(c);
            if (g.Sc[k] > bs) { bs = g.Sc[k]; best = c; }
        }
        return best;
    }
}

// Свидетельство (удостоверение) о государственной регистрации: адрес объекта, кадастровый номер участка,
// правообладатель. Своё распознавание обычного текста (шрифт с засечками, как Times New Roman).
// Адрес читается прямо с листа (не по справочнику улиц): бывают новые дома и улицы.

public class CertResult {
    public bool Found;                                // это свидетельство и в нём нашёлся адрес
    public string Address = "";                       // как в свидетельстве
    public string NpTip = "", Np = "", Selsovet = "", UlicaTip = "", Ulica = "", Dom = "", Korp = "", Kv = "";
    public bool Ok;                                   // всё прочитано уверенно
    // уверенно ли прочитаны названия (программа может подтвердить их своими списками: сельсоветы, н.п., улицы)
    public bool NpSure = true, SelsovetSure = true, UlicaSure = true;
    public bool Rest;                                 // номера и разбор адреса - уверенно
    public string Kadastr = "";
    public bool KadastrOk;
    public string Owner = "";                         // правообладатель (Ф.И.О.)
    public string Info = "";
    public List<string> Lines = new List<string>();

    // коротко, как в договоре: «д. Боровляны, ул. Лесная, д.7 (Боровлянский с/с)»
    public string Short {
        get {
            List<string> p = new List<string>();
            if (Np.Length > 0) p.Add(NpTip == "СТ" ? "СТ «" + Np + "»" : (NpTip + " " + Np).Trim());
            if (Ulica.Length > 0) p.Add((UlicaTip + " " + Ulica).Trim());
            string h = (Dom.Length > 0 ? (NpTip == "СТ" ? "уч." : "д.") + Dom : "") + (Korp.Length > 0 ? " корп." + Korp : "") + (Kv.Length > 0 ? " кв." + Kv : "");
            if (h.Trim().Length > 0) p.Add(h.Trim());
            string s = string.Join(", ", p.ToArray());
            if (Selsovet.Length > 0) s += " (" + Selsovet + " с/с)";
            return s;
        }
    }
}

public static class CertPage {
    public static string[] Families = { "Times New Roman" };
    // ı - палочка без точки: правая часть «ы» (ь + палочка), отдельно не встречается
    public const string Chars = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюяı0123456789.,-/:;()\"№";
    const string Letters = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя";
    const string Digits = "0123456789";
    static GlyphSet serif;
    public static double[] Blurs = { 0, 0.045, 0.09, 0.14 };
    public static int GridW = 24, GridH = 30;      // сетка сравнения мельче, чем для паспорта: строчные буквы похожи
    static readonly object sync = new object();

    public static void Init() {
        lock (sync) {
            if (serif == null) serif = GlyphSet.FromFonts(Chars, Families, new FontStyle[] { FontStyle.Regular, FontStyle.Bold }, GridW, GridH, Blurs);
        }
    }

    public static CertResult Read(string path) {
        Init();
        GrayImage full, small;
        Mrz.LoadPair(path, out full, out small);
        List<PassportPage.Line> l0 = PassportPage.FindLines(small, true);
        GrayImage sr = small.Rotate90();
        List<PassportPage.Line> l1 = PassportPage.FindLines(sr, true);
        if (PassportPage.Weight(l1) > PassportPage.Weight(l0)) { small = sr; l0 = l1; full = full.Rotate90(); }
        double f = (double)full.W / small.W;
        CertResult r = ReadPage(full, PassportPage.Scale(l0, f));
        if (!r.Found) {
            GrayImage ff = PassportPage.Rotate180(full), sf = PassportPage.Rotate180(small);
            CertResult o = ReadPage(ff, PassportPage.Scale(PassportPage.FindLines(sf, true), f));
            if (o.Found || o.Lines.Count > r.Lines.Count) r = o;
        }
        return r;
    }

    // ---------- текст страницы ----------

    static CertResult ReadPage(GrayImage g, List<PassportPage.Line> lines) {
        CertResult r = new CertResult();
        // строки по ширине бывают разорваны большими пробелами - куски одной строки собираем по высоте
        lines.Sort(delegate(PassportPage.Line a, PassportPage.Line b) {
            if (Math.Abs(a.Y - b.Y) < 0.4 * Math.Min(a.Cap, b.Cap)) return a.X0.CompareTo(b.X0);
            return a.Y.CompareTo(b.Y);
        });
        List<Glyph> all = new List<Glyph>();      // знаки по порядку; null - пробел или конец строки
        StringBuilder text = new StringBuilder();
        double prevY = double.MinValue, prevCap = 0;
        int read = 0;
        foreach (PassportPage.Line l in lines) {
            PassportPage.Strip st = PassportPage.ReadLine(g, l, serif);
            if (st == null || st.G.Count == 0) continue;
            bool newRow = Math.Abs(l.Y - prevY) >= 0.4 * Math.Min(l.Cap, prevCap > 0 ? prevCap : l.Cap);
            if (text.Length > 0) { text.Append(newRow ? '\n' : ' '); all.Add(null); }
            foreach (Glyph gl in st.G) {
                if (gl.Space) { text.Append(' '); all.Add(null); }
                text.Append(gl.C);
                all.Add(gl);
            }
            if (newRow) r.Lines.Add(PassportPage.Text(st.G));
            else if (r.Lines.Count > 0) r.Lines[r.Lines.Count - 1] += "  " + PassportPage.Text(st.G);
            prevY = l.Y;
            prevCap = l.Cap;
            read++;
            // первые строки - не свидетельство (или вверх ногами) - дальше не читаем
            if (read == 10 && !LooksLikeCert(text.ToString())) return r;
        }
        string t = text.ToString().Replace('\n', ' ');
        if (!LooksLikeCert(t)) return r;
        int s = Find(t, "по адресу", 0, 2);
        if (s < 0) { r.Info = "не найдено «по адресу»"; return r; }
        while (s < t.Length && (t[s] == ':' || t[s] == ' ' || t[s] == ';' || t[s] == '.')) s++;
        int e = t.Length;
        foreach (string stop in new string[] { "площадь", "назначение", "наименование", "произведена" }) {
            int k = FindStart(t, stop, s, 2);
            if (k >= 0 && k < e) e = k;
        }
        e = Math.Min(e, s + 200);
        while (e > s && (t[e - 1] == ' ' || t[e - 1] == ',')) e--;
        if (e <= s) return r;
        Parse(r, all, s, e);
        // кадастровый номер (у земельного участка)
        int kd = Find(t, "кадастровым номером", 0, 3);
        if (kd >= 0) {
            List<Glyph> num = Token(all, kd);
            string digits = "";
            bool ok = num.Count >= 10;
            foreach (Glyph gl in num) {
                char c = BestOf(gl, Digits);
                digits += c;
                if (Margin(gl, Digits) < 0.08) ok = false;
            }
            if (digits.Length >= 10) { r.Kadastr = digits; r.KadastrOk = ok; }
        }
        // правообладатель: после «гражданин(ка) Республики Беларусь» - три слова
        int ow = Find(t, "правообладатель", 0, 3);
        if (ow >= 0) {
            int rb = Find(t, "Беларусь", ow, 2);
            if (rb >= 0 && rb - ow < 80) {
                List<string> words = new List<string>();
                int at = rb;
                for (int w = 0; w < 6 && words.Count < 3; w++) {
                    List<Glyph> tok = Token(all, at);
                    if (tok.Count == 0) break;
                    at = Skip(all, at);
                    if (tok.Count < 2) continue;              // соринка, а не слово
                    StringBuilder sb = new StringBuilder();
                    foreach (Glyph gl in tok) sb.Append(BestOf(gl, Letters + "-"));
                    words.Add(PageResult.Title(sb.ToString()));
                }
                if (words.Count == 3) r.Owner = string.Join(" ", words.ToArray());
            }
        }
        return r;
    }

    static bool LooksLikeCert(string t) {
        foreach (string w in new string[] { "РЕСПУБЛИКА", "РЕГИСТР", "СВИДЕТЕЛЬСТВО", "РЕГИСТРАЦИИ", "заявлению", "недвижимого" })
            if (Find(t, w, 0, 2) >= 0) return true;
        return false;
    }

    static char Up(char c) { c = char.ToUpperInvariant(c); return c == 'Ё' ? 'Е' : c; }

    // нечёткий поиск (до maxErr ошибок, без учёта регистра): индекс сразу после найденного или -1
    public static int Find(string text, string pat, int from, int maxErr) {
        int m = pat.Length;
        int[] prev = new int[m + 1], cur = new int[m + 1];
        for (int j = 0; j <= m; j++) prev[j] = j;
        int bestEnd = -1, bestErr = maxErr + 1;
        for (int i = from; i < text.Length; i++) {
            cur[0] = 0;
            for (int j = 1; j <= m; j++)
                cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + (Up(text[i]) == Up(pat[j - 1]) ? 0 : 1));
            if (cur[m] < bestErr) { bestErr = cur[m]; bestEnd = i + 1; }
            else if (bestEnd >= 0 && i + 1 - bestEnd > m) break;   // первое хорошее совпадение уже позади
            int[] tmp = prev; prev = cur; cur = tmp;
        }
        return bestErr <= maxErr ? bestEnd : -1;
    }

    // то же, но начало найденного
    static int FindStart(string text, string pat, int from, int maxErr) {
        int e = Find(text, pat, from, maxErr);
        if (e < 0) return -1;
        int best = -1, be = maxErr + 1;
        for (int st = Math.Max(from, e - pat.Length - maxErr); st <= e - pat.Length + maxErr && st < e; st++) {
            int d = PassportPage.Distance(text.Substring(st, e - st).ToUpperInvariant().Replace('Ё', 'Е'), pat.ToUpperInvariant().Replace('Ё', 'Е'));
            if (d < be) { be = d; best = st; }
        }
        return best;
    }

    // знаки следующего слова начиная с позиции at (в тексте пробелы и концы строк - null)
    static List<Glyph> Token(List<Glyph> all, int at) {
        // позиции текста и списка знаков совпадают
        List<Glyph> r = new List<Glyph>();
        int i = at;
        while (i < all.Count && (all[i] == null || all[i].C == ':' || all[i].C == '-' || all[i].C == ',')) i++;
        for (; i < all.Count && all[i] != null && all[i].C != ','; i++) r.Add(all[i]);
        return r;
    }

    static int Skip(List<Glyph> all, int at) {
        int i = at;
        while (i < all.Count && (all[i] == null || all[i].C == ',')) i++;
        while (i < all.Count && all[i] != null && all[i].C != ',') i++;
        return i;
    }

    static char BestOf(Glyph g, string allowed) {
        char best = g.C;
        double bs = double.MinValue;
        foreach (char c in allowed) {
            int k = Chars.IndexOf(c);
            if (k >= 0 && g.Sc[k] > bs) { bs = g.Sc[k]; best = c; }
        }
        return best;
    }

    // насколько лучший из разрешённых знаков лучше следующего (заглавная и строчная одной буквы - не соперники)
    static double Margin(Glyph g, string allowed) {
        char best = BestOf(g, allowed);
        double b1 = g.Sc[Chars.IndexOf(best)], b2 = double.MinValue;
        foreach (char c in allowed) {
            if (char.ToUpperInvariant(c) == char.ToUpperInvariant(best)) continue;
            int k = Chars.IndexOf(c);
            if (k >= 0 && g.Sc[k] > b2) b2 = g.Sc[k];
        }
        return b2 == double.MinValue ? 9 : b1 - b2;
    }

    // ---------- адрес -> поля ----------

    class Word {
        public string Text = "";
        public List<Glyph> G = new List<Glyph>();
    }

    static void Parse(CertResult r, List<Glyph> all, int s, int e) {
        // слова адреса (запятая - граница частей)
        List<List<Word>> parts = new List<List<Word>>();
        List<Word> part = new List<Word>();
        Word w = null;
        StringBuilder raw = new StringBuilder();
        for (int i = s; i < e; i++) {
            Glyph gl = all[i];
            raw.Append(gl == null ? ' ' : gl.C);
            if (gl == null) { if (w != null) { part.Add(w); w = null; } continue; }
            bool sep = gl.C == ',' || gl.C == ';';
            if (gl.C == '.' && w != null) {
                // точка вместо запятой: после слова, которое не сокращение (или вторая точка подряд)
                string prev = Clean(w.G).ToLowerInvariant().TrimEnd('.');
                bool abbr = Array.IndexOf(DotAbbr, prev) >= 0;
                bool before = i + 1 < e && all[i + 1] == null;
                if (before && !abbr) sep = true;
                if (w.G.Count > 0 && w.G[w.G.Count - 1].C == '.') sep = true;
            }
            if (sep) {
                if (w != null) { part.Add(w); w = null; }
                if (part.Count > 0) { parts.Add(part); part = new List<Word>(); }
                continue;
            }
            if (w == null) w = new Word();
            w.G.Add(gl);
        }
        if (w != null) part.Add(w);
        if (part.Count > 0) parts.Add(part);
        r.Address = Regex.Replace(raw.ToString(), " +", " ").Trim();
        // слова: цифры - цифрами, буквы - буквами; сокращения («г.», «аг.», «ул.», «кв») - перечитать по списку
        foreach (List<Word> p in parts)
            for (int i = 0; i < p.Count; i++) {
                p[i].Text = Clean(p[i].G);
                string ab = Abbreviation(p[i]);
                if (ab != null) p[i].Text = ab;
            }
        bool ok = true;
        // «пр т.» - дефис не прочитался: это «пр-т»
        foreach (List<Word> p in parts)
            if (p.Count >= 2 && p[0].Text.ToLowerInvariant().TrimEnd('.') == "пр" && p[1].Text.ToLowerInvariant().TrimEnd('.') == "т") {
                p[0].Text = "пр-т";
                p[0].G.AddRange(p[1].G);
                p.RemoveAt(1);
            }
        foreach (List<Word> p in parts) {
            string first = p[0].Text, low = first.ToLowerInvariant().TrimEnd('.');
            string joined = Join(p, 0);
            if (Has(p, "обл") || Has(p, "р-н") || Has(p, "район") || Has(p, "область")) continue;
            if (p.Count >= 2 && (Like(p[p.Count - 1].Text, "с/с") || Like(p[p.Count - 1].Text, "сельсовет"))) {
                r.Selsovet = Title(Join(p, 0, p.Count - 1));
                r.SelsovetSure = Sure(p, 0, p.Count - 1, false);
                continue;
            }
            // садоводческое товарищество
            if (Like(first, "СТ") || Like(first, "с/т") || low.StartsWith("садовод")) {
                int k = 1;
                while (k < p.Count && p[k].Text.ToLowerInvariant().StartsWith("товарищ")) k++;
                r.NpTip = "СТ";
                r.Np = Join(p, k).Trim('"', '«', '»', ' ', '.', ',');
                r.NpSure = Sure(p, k, p.Count, false);
                continue;
            }
            // населённый пункт: «г. Минск», «д. Боровляны», «аг. Ждановичи», «г.п. Мачулищи»
            string np = NpType(first);
            if (np != null && p.Count >= 2 && !char.IsDigit(p[1].Text[0])) {
                r.NpTip = np;
                int k = 1;
                // улица в той же части без запятой: «г. Минск ул. Ленина 5» - редко; берём до типа улицы
                int end = p.Count;
                for (int q = 2; q < p.Count; q++) if (StreetType(p[q].Text) != null) { end = q; break; }
                r.Np = Title(Join(p, k, end));
                r.NpSure = Sure(p, k, end, false);
                if (end < p.Count) Street(r, p.GetRange(end, p.Count - end), ref ok);
                continue;
            }
            if (StreetType(first) != null) { Street(r, p, ref ok); continue; }
            // квартира, корпус, дом, участок
            if (low == "кв" || low == "квартира") { r.Kv = Num(p, 1, ref ok); continue; }
            if (low == "корп" || low == "корпус" || low == "к") { r.Korp = Num(p, 1, ref ok); continue; }
            if (low == "участок" || low == "уч" || low == "д" || low == "дом") { r.Dom = Num(p, 1, ref ok); continue; }
            if (char.IsDigit(first[0])) { r.Dom = Num(p, 0, ref ok); continue; }
            r.Info += "непонятная часть адреса: " + joined + "; ";
            ok = false;
        }
        r.Found = r.Np.Length > 0 || r.Ulica.Length > 0;
        if (r.Dom.Length == 0 && r.Kv.Length == 0) ok = false;
        r.Rest = ok;
        r.Ok = r.Found && ok && r.NpSure && r.SelsovetSure && r.UlicaSure;
    }

    // «ул. Лесная, 7» / «пр-т. Партизанский 88» - тип, название, может быть сразу и номер дома
    static void Street(CertResult r, List<Word> p, ref bool ok) {
        r.UlicaTip = StreetType(p[0].Text);
        int end = p.Count;
        while (end > 1 && char.IsDigit(p[end - 1].Text[0])) end--;
        r.Ulica = Title(Join(p, 1, end));
        r.UlicaSure = Sure(p, 1, end, false);
        if (end < p.Count) r.Dom = Num(p, end, ref ok);
    }

    static string Num(List<Word> p, int from, ref bool ok) {
        if (from >= p.Count) { ok = false; return ""; }
        string s = "";
        for (int i = from; i < p.Count; i++) s += p[i].Text;
        ok &= Sure(p, from, p.Count, true);
        return s;
    }

    // сокращения с точкой: точка после них - не конец части адреса
    static readonly string[] DotAbbr = { "г", "д", "аг", "п", "гп", "г.п", "пос", "ул", "пер", "пр-т", "пр", "просп", "туп", "мкр", "корп", "к", "уч", "обл", "кв" };

    static readonly string[] Abbr = { "г.", "д.", "аг.", "п.", "г.п.", "гп.", "пос.", "ул.", "пер.", "пр-т.", "пр-т", "пр.", "просп.", "туп.",
        "мкр.", "б-р", "кв", "кв.", "корп.", "к.", "уч.", "участок", "СТ", "с/т", "с/с", "обл.", "р-н", "д" };

    // похоже на сокращение - перечитать участок строки ровно по каждому варианту и взять лучший
    static string Abbreviation(Word w) {
        if (w.G.Count == 0 || w.G.Count > 8 || w.Text.Length > 8) return null;
        bool near = false;
        foreach (string a in Abbr) if (PassportPage.Distance(w.Text.ToLowerInvariant(), a.ToLowerInvariant()) <= 2) near = true;
        if (!near) return null;
        Glyph g0 = w.G[0], g1 = w.G[w.G.Count - 1];
        PassportPage.Strip st = g0.Src as PassportPage.Strip;
        if (st == null || g1.Src != g0.Src) return null;
        int from = g0.CutA, to = g1.CutB;
        // захватить и точку сразу после слова (если отдельной точкой не прочиталась)
        while (to + 1 < st.Cuts.Count && !st.Gap[to] && st.Cuts[to + 1] - st.Cuts[to] < 0.3 * st.Cap) to++;
        double ink;
        double free = PassportPage.Free(st, from, to, null, out ink);
        ink = Math.Max(ink, 0.3);
        string best = null;
        double b1 = double.MinValue, b2 = double.MinValue;
        foreach (string a in Abbr) {
            string[] sets = new string[a.Length];
            for (int i = 0; i < a.Length; i++) sets[i] = a[i].ToString();
            string text;
            double v = PassportPage.Constrained(st, from, to, sets, out text);
            if (v > b1) { b2 = b1; b1 = v; best = a; }
            else if (v > b2) b2 = v;
        }
        if (best == null || (free - b1) / ink > 0.15) return null;
        return best;
    }

    // слово: если в нём цифры - читать как номер (цифры, в конце буква: 7а, 40Б, 7/1), иначе - буквами
    static string Clean(List<Glyph> gs) {
        int dig = 0, let = 0;
        foreach (Glyph g in gs) { if (char.IsDigit(g.C)) dig++; else if (char.IsLetter(g.C)) let++; }
        StringBuilder sb = new StringBuilder();
        if (dig > 0 && dig >= let) {
            for (int i = 0; i < gs.Count; i++) {
                char c = gs[i].C;
                if (c == '/' || c == '-') { sb.Append(c); continue; }
                if (".,:;\"".IndexOf(c) >= 0) continue;
                bool tail = i == gs.Count - 1 && i > 0 && char.IsLetter(c) && gs[i].S - gs[i].Sc[Chars.IndexOf(BestOf(gs[i], Digits))] > 0.15;
                sb.Append(tail ? char.ToLowerInvariant(c) : BestOf(gs[i], Digits));
            }
        } else {
            bool afterHyphen = false;
            for (int i = 0; i < gs.Count; i++) {
                Glyph g = gs[i];
                if (g.C == 'ı') {
                    // «ь» + палочка = «ы»
                    if (sb.Length > 0 && (sb[sb.Length - 1] == 'ь' || sb[sb.Length - 1] == 'Ь')) { sb[sb.Length - 1] = sb[sb.Length - 1] == 'ь' ? 'ы' : 'Ы'; continue; }
                    sb.Append(BestOf(g, Letters));
                    continue;
                }
                if (char.IsDigit(g.C) && afterHyphen) { sb.Append(g.C); continue; }   // «Строитель-2»
                afterHyphen = g.C == '-' || afterHyphen && char.IsDigit(g.C);
                char ch = char.IsLetter(g.C) || ".-/\"".IndexOf(g.C) >= 0 ? g.C : BestOf(g, Letters);
                // русское письмо: й - только после гласной; ь, ъ - не после гласной и не в начале слова
                char before = sb.Length > 0 ? char.ToLowerInvariant(sb[sb.Length - 1]) : ' ';
                bool vowel = "аеёиоуыэюя".IndexOf(before) >= 0;
                char lo = char.ToLowerInvariant(ch);
                if (lo == 'й' && !vowel || (lo == 'ь' || lo == 'ъ') && (vowel || !char.IsLetter(before))) {
                    string other = Letters.Replace("й", "").Replace("Й", "").Replace("ь", "").Replace("Ь", "").Replace("ъ", "").Replace("Ъ", "");
                    ch = BestOf(g, other);
                    g.S = -1;                                   // и такое слово - под сомнением
                }
                sb.Append(ch);
            }
        }
        return sb.ToString();
    }

    // прочитано уверенно: у каждого знака заметный отрыв от следующего варианта
    static bool Sure(List<Word> p, int from, int to, bool number) {
        for (int i = from; i < to && i < p.Count; i++)
            foreach (Glyph g in p[i].G) {
                if (".-/\"«»".IndexOf(g.C) >= 0) continue;
                bool d = number && char.IsDigit(BestOf(g, Digits + Letters)) || number && char.IsDigit(g.C);
                if (g.S < 0.5) return false;
                if (Margin(g, d ? Digits : Letters) < (d ? 0.08 : 0.05)) return false;
            }
        return true;
    }

    static string Join(List<Word> p, int from) { return Join(p, from, p.Count); }

    static string Join(List<Word> p, int from, int to) {
        List<string> l = new List<string>();
        for (int i = from; i < to && i < p.Count; i++) l.Add(p[i].Text);
        return string.Join(" ", l.ToArray());
    }

    static bool Has(List<Word> p, string w) {
        foreach (Word x in p) if (Like(x.Text, w)) return true;
        return false;
    }

    // одно и то же слово с точностью до одной ошибки распознавания (и точки в конце)
    static bool Like(string a, string b) {
        string x = a.ToLowerInvariant().TrimEnd('.'), y = b.ToLowerInvariant().TrimEnd('.');
        if (x == y) return true;
        return y.Length >= 3 && PassportPage.Distance(x, y) <= 1;
    }

    static string NpType(string w) {
        string x = w.ToLowerInvariant().Replace(" ", "");
        if (x == "г." || x == "г") return "г.";
        if (x == "д." || x == "д") return "д.";
        if (x == "аг." || x == "аг") return "аг.";
        if (x == "п." || x == "пос." || x == "п") return "п.";
        if (x == "г.п." || x == "гп." || x == "г.п" || x == "гп") return "г.п.";
        return null;
    }

    static string StreetType(string w) {
        string x = w.ToLowerInvariant().TrimEnd('.');
        if (x == "ул" || x == "улица") return "ул.";
        if (x == "пер" || x == "переулок") return "пер.";
        if (x == "пр-т" || x == "просп" || x == "проспект" || x == "пр-кт") return "пр-т";
        if (x == "пр" || x == "проезд") return "пр.";
        if (x == "туп" || x == "тупик") return "туп.";
        if (x == "мкр" || x == "микрорайон") return "мкр.";
        if (x == "б-р" || x == "бульвар") return "б-р";
        return null;
    }

    // «боровляны» -> «Боровляны», «новая боровая» -> «Новая Боровая», «Строитель-2» как есть
    static string Title(string s) {
        StringBuilder sb = new StringBuilder();
        bool start = true;
        foreach (char c in s) {
            sb.Append(start ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c));
            start = c == ' ' || c == '-' || c == '"';
        }
        return sb.ToString();
    }
}

// Счётчик на фото (Энергомера CE208BY / CE318BY): заводской номер - по штрихкоду наклейки (Code 128, своя контрольная
// сумма), тип - по надписи модификации со сверкой со справочником, показания - с табло (семисегментные цифры).
public class MeterResult {
    public bool Found;                    // на снимке счётчик: табло, или штрихкод вместе с надписью типа
    public string Serial = "";            // заводской номер
    public int SerialRows;                // сколько строк развёртки штрихкода дали этот номер
    public string Type = "";              // полное обозначение из справочника
    public bool TypeOk;
    public string TypeText = "";          // как прочитана надпись
    public string Reading = "";           // показания как на табло, с запятой: «05801,969» («?» - цифра не разобрана)
    public bool ReadingOk;                // цифры чёткие, экран 1.8.0 (сумма)
    public string Code = "";              // код экрана: «1.8.0» - сумма
    public string Year = "";              // год на наклейке со штрихкодом (выпуск, первая поверка)
    public string Info = "";
}

public static class MeterPhoto {
    // ---------- штрихкод Code 128 ----------

    static readonly string[] P128 = ("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 " +
        "113222 123122 123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 " +
        "131123 131321 112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 " +
        "213131 311123 311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 " +
        "122411 142112 142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 " +
        "412121 111143 111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232").Split(' ');
    static Dictionary<int, List<int>> eTable;

    static int EKey(int e1, int e2, int e3, int e4) { return ((e1 * 10 + e2) * 10 + e3) * 10 + e4; }

    static void InitTable() {
        if (eTable != null) return;
        Dictionary<int, List<int>> t = new Dictionary<int, List<int>>();
        for (int v = 0; v < P128.Length; v++) {
            int[] w = new int[6];
            for (int i = 0; i < 6; i++) w[i] = P128[v][i] - '0';
            int k = EKey(w[0] + w[1], w[1] + w[2], w[2] + w[3], w[3] + w[4]);
            List<int> l;
            if (!t.TryGetValue(k, out l)) { l = new List<int>(); t[k] = l; }
            l.Add(v);
        }
        eTable = t;
    }

    // знак по ширинам 6 элементов (полоса, пробел, ...): по расстояниям между краями - так не мешает растекание краски
    static int Symbol(double[] r, int i) {
        double T = 0;
        for (int k = 0; k < 6; k++) T += r[i + k];
        double m = T / 11;
        int[] e = new int[4];
        for (int k = 0; k < 4; k++) {
            double x = (r[i + k] + r[i + k + 1]) / m;
            e[k] = (int)Math.Round(x);
            if (e[k] < 2 || e[k] > 7) return -1;
        }
        List<int> c;
        if (!eTable.TryGetValue(EKey(e[0], e[1], e[2], e[3]), out c)) return -1;
        if (c.Count == 1) return c[0];
        // несколько знаков с теми же расстояниями - по сумме ширин полос
        double bars = (r[i] + r[i + 2] + r[i + 4]) / m;
        int best = -1;
        double bd = 9;
        foreach (int v in c) {
            int b = P128[v][0] - '0' + P128[v][2] - '0' + P128[v][4] - '0';
            if (Math.Abs(b - bars) < bd) { bd = Math.Abs(b - bars); best = v; }
        }
        return best;
    }

    static bool Stop(double[] r, int i) {
        if (i + 7 > r.Length) return false;
        double T = 0;
        for (int k = 0; k < 7; k++) T += r[i + k];
        double m = T / 13;
        int[] want = { 2, 3, 3, 1, 1, 1, 2 };
        for (int k = 0; k < 6; k++) {
            double x = (r[i + k] + r[i + k + 1]) / m;
            if (Math.Abs(x - (want[k] + want[k + 1])) > 0.6) return false;
        }
        return true;
    }

    // разбор последовательности ширин (r[0] - полоса); null - не Code 128
    static string Decode128(double[] r, int start) { int end; return Decode128(r, start, out end); }

    // end - индекс последнего участка знака «стоп»
    static string Decode128(double[] r, int start, out int end) {
        end = -1;
        int s = Symbol(r, start);
        if (s < 103 || s > 105) return null;
        // тихая зона перед штрихкодом
        double mod = 0;
        for (int k = 0; k < 6; k++) mod += r[start + k];
        mod /= 11;
        if (start > 0 && r[start - 1] < 3 * mod) return null;
        List<int> vals = new List<int>();
        vals.Add(s);
        int i = start + 6;
        while (true) {
            if (Stop(r, i)) { end = i + 6; break; }
            if (i + 6 > r.Length) return null;
            int v = Symbol(r, i);
            if (v < 0 || v > 102) return null;
            vals.Add(v);
            i += 6;
            if (vals.Count > 40) return null;
        }
        if (vals.Count < 3) return null;
        int sum = vals[0];
        for (int k = 1; k < vals.Count - 1; k++) sum += k * vals[k];
        if (sum % 103 != vals[vals.Count - 1]) return null;
        // текст: набор C - пары цифр, B - знаки ASCII
        StringBuilder sb = new StringBuilder();
        int set = s == 105 ? 'C' : s == 104 ? 'B' : 'A';
        for (int k = 1; k < vals.Count - 1; k++) {
            int v = vals[k];
            if (set == 'C') {
                if (v < 100) sb.Append(v.ToString("00"));
                else if (v == 100) set = 'B';
                else if (v == 101) set = 'A';
            } else {
                if (v == 99) set = 'C';
                else if (v == 100 && set == 'A') set = 'B';
                else if (v == 101 && set == 'B') set = 'A';
                else if (v < 96) sb.Append((char)(set == 'B' || v < 64 ? v + 32 : v - 64));
            }
        }
        return sb.ToString();
    }

    // развёртка: строки снимка через шаг; в каждой - края полос с точностью до доли точки
    // where: x0, y0, x1, y1 штрихкода с самым частым текстом, и 1 - если он перевёрнут
    static Dictionary<string, int> Scan(GrayImage g, out int[] where) {
        InitTable();
        Dictionary<string, int> votes = new Dictionary<string, int>();
        where = null;
        Dictionary<string, int[]> box = new Dictionary<string, int[]>();
        int w = g.W, R = Math.Max(6, w / 160), step = Math.Max(2, g.H / 600);
        double[] mean = new double[w];
        float[] v = new float[w];
        int bestN = 0;
        string bestText = null;
        for (int y = 0; y < g.H; y += step) {
            int o = y * w;
            for (int x = 0; x < w; x++) v[x] = g.P[o + x];
            // порог - середина между самым светлым и самым тёмным рядом (среднее по окну сдвигает края у тихой зоны)
            for (int x0 = 0; x0 < w; x0 += 4) {
                int xa = Math.Max(0, x0 - R), xb = Math.Min(w - 1, x0 + 3 + R);
                float lo = 255, hi = 0;
                for (int x = xa; x <= xb; x++) { if (v[x] < lo) lo = v[x]; if (v[x] > hi) hi = v[x]; }
                double t = hi - lo < 30 ? -1 : (lo + hi) / 2.0;
                for (int x = x0; x < x0 + 4 && x < w; x++) mean[x] = t;
            }
            // края: переходы через порог
            List<double> edges = new List<double>();
            List<bool> darkAfter = new List<bool>();
            bool prev = v[0] < mean[0];
            for (int x = 1; x < w; x++) {
                bool d = v[x] < mean[x];
                if (d == prev) continue;
                double t0 = v[x - 1] - mean[x - 1], t1 = v[x] - mean[x];
                double f = t0 == t1 ? 0.5 : t0 / (t0 - t1);
                edges.Add(x - 1 + f);
                darkAfter.Add(d);
                prev = d;
            }
            if (edges.Count < 30) continue;
            double[] runs = new double[edges.Count - 1];
            for (int k = 0; k + 1 < edges.Count; k++) runs[k] = edges[k + 1] - edges[k];
            // в обе стороны (штрихкод может быть вверх ногами); участок k тёмный, если после края k - тёмное
            int n = runs.Length;
            for (int dir = 0; dir < 2; dir++) {
                double[] r = runs;
                if (dir == 1) {
                    r = new double[n];
                    for (int k = 0; k < n; k++) r[k] = runs[n - 1 - k];
                }
                for (int k = 0; k + 13 < n; k++) {
                    if (darkAfter[dir == 0 ? k : n - 1 - k] == false) continue;
                    int end;
                    string txt = Decode128(r, k, out end);
                    if (txt == null || txt.Length == 0) continue;
                    int c;
                    votes.TryGetValue(txt, out c);
                    votes[txt] = c + 1;
                    if (c + 1 > bestN) { bestN = c + 1; bestText = txt; }
                    double xa = dir == 0 ? edges[k] : edges[n - 1 - end], xb = dir == 0 ? edges[end + 1] : edges[n - k];
                    int[] bx;
                    if (!box.TryGetValue(txt, out bx)) { bx = new int[] { (int)xa, y, (int)xb, y, dir }; box[txt] = bx; }
                    bx[0] = Math.Min(bx[0], (int)xa); bx[2] = Math.Max(bx[2], (int)xb); bx[1] = Math.Min(bx[1], y); bx[3] = Math.Max(bx[3], y);
                }
            }
        }
        if (bestText != null) where = box[bestText];
        return votes;
    }

    // заводской номер со штрихкода: строка с наибольшим числом подтверждений (не меньше двух)
    public static string Barcode(GrayImage full, out int rows) { string year; return Barcode(full, out rows, out year); }

    // year - год, напечатанный поперёк справа от штрихкода (на наклейке Энергомеры)
    public static string Barcode(GrayImage full, out int rows, out string year) {
        rows = 0;
        year = "";
        string best = "";
        GrayImage img = full;
        for (int turn = 0; turn < 2; turn++) {
            int[] where;
            Dictionary<string, int> v = Scan(img, out where);
            foreach (KeyValuePair<string, int> kv in v)
                if (kv.Value > rows && kv.Key.Length >= 6) { rows = kv.Value; best = kv.Key; }
            if (rows >= 2) {
                if (where != null) try { year = Year(img, where); } catch (Exception) { year = ""; }
                break;
            }
            img = full.Rotate90();       // снимок на боку - штрихкод вертикальный
        }
        if (rows < 2) { rows = 0; return ""; }
        return best;
    }

    // год справа от штрихкода (у перевёрнутого - слева), написан снизу вверх: 4 цифры 20xx
    static string Year(GrayImage g, int[] b) {
        Init();
        int w = b[2] - b[0];
        bool flip = b[4] == 1;
        // надпись: по высоте - около трети ширины штрихкода, середина - чуть ниже середины полос
        double cy = (b[1] + b[3]) / 2.0 + (flip ? -0.08 : 0.08) * w;
        int x0 = flip ? b[0] - (int)(0.32 * w) : b[2] + (int)(0.02 * w), x1 = flip ? b[0] - (int)(0.02 * w) : b[2] + (int)(0.32 * w);
        int y0 = (int)(cy - 0.32 * w), y1 = (int)(cy + 0.32 * w);
        if (x1 - x0 < 10 || y1 - y0 < 10) return "";
        GrayImage crop = Crop(g, x0, y0, x1, y1);
        // поворот: снизу вверх - по часовой; у перевёрнутого снимка - против
        GrayImage r = crop.Rotate90();
        if (flip) r = r.Rotate90().Rotate90();
        // крупнее и с белыми полями - поиск строк рассчитан на текст не мельче 20 точек
        r = Grow(r, Math.Max(1, (int)Math.Ceiling(80.0 / Math.Max(1, r.H))));
        foreach (GrayImage img in new GrayImage[] { r, r.Rotate90().Rotate90() }) {
            foreach (PassportPage.Line l in PassportPage.FindLines(img, false)) {
                PassportPage.Strip st = PassportPage.ReadLine(img, l, latin);
                if (st == null || st.G == null) continue;
                StringBuilder sb = new StringBuilder();
                foreach (Glyph gl in st.G) sb.Append(gl.C);
                string t = Key(sb.ToString());
                for (int i = 0; i + 4 <= t.Length; i++) {
                    string y4 = t.Substring(i, 4);
                    int yv;
                    if (int.TryParse(y4, out yv) && yv >= 2005 && yv <= DateTime.Now.Year) return y4;
                }
            }
        }
        return "";
    }

    // ---------- табло: семисегментные цифры ----------

    // сегменты: a - верх, b - правый верхний, c - правый нижний, d - низ, e - левый нижний, f - левый верхний, g - середина
    static readonly string[] Seg7 = { "abcdef", "bc", "abdeg", "abcdg", "bcfg", "acdfg", "acdefg", "abc", "abcdefg", "abcdfg" };

    class Blob { public int X0, Y0, X1, Y1; public int Ht { get { return Y1 - Y0 + 1; } } public int Wd { get { return X1 - X0 + 1; } } }

    // тёмное относительно среднего по окну (r - полуширина окна)
    static bool[] Dark(GrayImage g, int r) {
        int w = g.W, h = g.H;
        long[] sum = new long[(w + 1) * (h + 1)];
        for (int y = 0; y < h; y++) {
            long row = 0;
            for (int x = 0; x < w; x++) {
                row += g.P[y * w + x];
                sum[(y + 1) * (w + 1) + x + 1] = sum[y * (w + 1) + x + 1] + row;
            }
        }
        bool[] d = new bool[w * h];
        for (int y = 0; y < h; y++) {
            int ya = Math.Max(0, y - r), yb = Math.Min(h, y + r + 1);
            for (int x = 0; x < w; x++) {
                int xa = Math.Max(0, x - r), xb = Math.Min(w, x + r + 1);
                long s = sum[yb * (w + 1) + xb] - sum[ya * (w + 1) + xb] - sum[yb * (w + 1) + xa] + sum[ya * (w + 1) + xa];
                int mean = (int)(s / ((xb - xa) * (yb - ya)));
                int v = g.P[y * w + x];
                d[y * w + x] = mean - v > Math.Max(12, mean * 15 / 100);
            }
        }
        return d;
    }

    // утолщение тёмного на r точек (сегменты одной цифры сливаются в одно пятно)
    static bool[] Dilate(bool[] m, int w, int h, int r) {
        if (r <= 0) return m;
        bool[] t = new bool[w * h], o = new bool[w * h];
        for (int y = 0; y < h; y++) {
            int cnt = 0, b = y * w;
            for (int x = 0; x < Math.Min(w, r); x++) if (m[b + x]) cnt++;
            for (int x = 0; x < w; x++) {
                if (x + r < w && m[b + x + r]) cnt++;
                if (x - r - 1 >= 0 && m[b + x - r - 1]) cnt--;
                t[b + x] = cnt > 0;
            }
        }
        for (int x = 0; x < w; x++) {
            int cnt = 0;
            for (int y = 0; y < Math.Min(h, r); y++) if (t[y * w + x]) cnt++;
            for (int y = 0; y < h; y++) {
                if (y + r < h && t[(y + r) * w + x]) cnt++;
                if (y - r - 1 >= 0 && t[(y - r - 1) * w + x]) cnt--;
                o[y * w + x] = cnt > 0;
            }
        }
        return o;
    }

    // значение, вокруг которого больше всего других (в пределах tol); из этой группы - крайнее (min или max)
    static double Cluster(List<int> v, double tol, bool min) {
        int best = 0;
        double res = v[0];
        foreach (int a in v) {
            int cnt = 0;
            double ext = a;
            foreach (int b in v) if (Math.Abs(b - a) <= tol) { cnt++; ext = min ? Math.Min(ext, b) : Math.Max(ext, b); }
            if (cnt > best || cnt == best && (min ? ext < res : ext > res)) { best = cnt; res = ext; }
        }
        return res;
    }

    // увеличить в k раз и обвести полем цвета фона: сверху и снизу - широким (поиск строк ждёт строку
    // не выше 1/12 высоты картинки, как на странице)
    static GrayImage Grow(GrayImage g, int k) {
        int pad = 20 * k, padY = 6 * g.H * k;
        GrayImage o = new GrayImage(g.W * k + 2 * pad, g.H * k + 2 * padY);
        for (int i = 0; i < o.P.Length; i++) o.P[i] = 255;
        // поле - цвета фона наклейки (светлее всего)
        byte bg = 0;
        foreach (byte v in g.P) if (v > bg) bg = v;
        for (int i = 0; i < o.P.Length; i++) o.P[i] = bg;
        for (int y = 0; y < g.H * k; y++)
            for (int x = 0; x < g.W * k; x++) o.P[(y + padY) * o.W + x + pad] = (byte)g.Sample((double)x / k, (double)y / k);
        return o;
    }

    static void Save(GrayImage g, string file) {
        using (Bitmap bm = new Bitmap(g.W, g.H)) {
            for (int y = 0; y < g.H; y++) for (int x = 0; x < g.W; x++) { int v = g.P[y * g.W + x]; bm.SetPixel(x, y, Color.FromArgb(v, v, v)); }
            bm.Save(file);
        }
    }

    static GrayImage Crop(GrayImage g, int x0, int y0, int x1, int y1) {
        x0 = Math.Max(0, x0); y0 = Math.Max(0, y0); x1 = Math.Min(g.W - 1, x1); y1 = Math.Min(g.H - 1, y1);
        GrayImage c = new GrayImage(Math.Max(1, x1 - x0 + 1), Math.Max(1, y1 - y0 + 1));
        for (int y = 0; y < c.H; y++) Array.Copy(g.P, (y0 + y) * g.W + x0, c.P, y * c.W, c.W);
        return c;
    }

    // ряды пятен одной высоты слева направо (кандидаты в строку цифр)
    static List<List<Blob>> Rows(List<Blob> bl, int minCount) {
        bl.Sort(delegate(Blob a, Blob b) { return a.X0.CompareTo(b.X0); });
        List<List<Blob>> rows = new List<List<Blob>>();
        bool[] used = new bool[bl.Count];
        for (int s = 0; s < bl.Count; s++) {
            if (used[s]) continue;
            List<Blob> row = new List<Blob>();
            row.Add(bl[s]);
            int cur = s;
            while (true) {
                int next = -1;
                Blob c = bl[cur];
                for (int j = cur + 1; j < bl.Count; j++) {
                    Blob d = bl[j];
                    double h = Math.Max(c.Ht, d.Ht);
                    if (d.X0 > c.X1 + 0.9 * h) break;
                    if (used[j] || d.X0 < c.X1 - 0.15 * h) continue;
                    // с соседом слева: верх на одном уровне, низ - почти (у «7» и «1» низ бывает выше)
                    if (Math.Abs(d.Y0 - c.Y0) > 0.12 * h || Math.Abs(d.Y1 - c.Y1) > 0.2 * h || Math.Min(c.Ht, d.Ht) < 0.75 * h) continue;
                    if (next < 0 || d.X0 < bl[next].X0) next = j;
                }
                if (next < 0) break;
                row.Add(bl[next]);
                cur = next;
            }
            if (row.Count >= minCount) {
                foreach (Blob b in row) used[bl.IndexOf(b)] = true;
                rows.Add(row);
            }
        }
        return rows;
    }

    public static bool Debug;
    public static string DebugDump;
    public static double DebugH0, DebugH1;

    class Digits {
        public string Text = "";        // цифры, «?» - не разобрана; точка - где нашлась
        public double Score;            // сумма уверенностей
        public bool Sure = true;
        public int Count, Bad;
        public double Top, Base, Left;  // в координатах полного снимка
        public double H;
        public double Tilt;             // поворот строки (dy/dx), с которым она прочитана
        public double TiltFix;          // уточнение поворота по самим цифрам
        public double X0, X1, Y0, Y1;   // где строка на полном снимке (для отсева повторов)
    }

    // прочитать ряд по полному снимку. Пятна ряда - на уменьшенном снимке, выпрямленном сдвигом x' = x + sh*y: по ним
    // только полоса строки; цифры в полосе делятся по пустым столбцам (у семисегментных цифр между ними всегда просвет)
    static Digits ReadRow(GrayImage full, List<Blob> row, double f, double sh) { return ReadRow(full, row, f, sh, double.NaN); }

    static Digits ReadRow(GrayImage full, List<Blob> row, double f, double sh, double tiltSet) {
        int n = row.Count;
        // ряд может быть чуть повёрнут: прямая по низу пятен
        double tilt = 0, xref = 0;
        if (!double.IsNaN(tiltSet)) tilt = tiltSet;
        else if (n >= 2) {
            double sx = 0, sy = 0, sxx = 0, sxy = 0;
            foreach (Blob b in row) { double x = (b.X0 + b.X1) / 2.0, y = b.Y1; sx += x; sy += y; sxx += x * x; sxy += x * y; }
            double den = n * sxx - sx * sx;
            if (den > 0) tilt = (n * sxy - sx * sy) / den;
            if (Math.Abs(tilt) > 0.3) tilt = 0;
        }
        xref = (row[0].X0 + row[n - 1].X1) / 2.0 * f;
        List<double> tops = new List<double>(), bots = new List<double>();
        foreach (Blob b in row) {
            double off = tilt * ((b.X0 + b.X1) / 2.0 * f - xref);
            tops.Add(b.Y0 * f - off); bots.Add((b.Y1 + 1) * f - off);
        }
        tops.Sort(); bots.Sort();
        double top = tops[n / 2], bot = bots[bots.Count - 1 - n / 3];
        double h = bot - top;
        if (h < 10) return null;
        // полоса шире ряда: недостающие цифры (не собравшиеся в пятно) найдутся по столбцам
        int core0 = (int)(row[0].X0 * f), core1 = (int)((row[n - 1].X1 + 1) * f);
        int bx0 = (int)(core0 - 6 * h), bx1 = (int)(core1 + 6 * h);
        int by0 = (int)(top - 0.2 * h), by1 = (int)(bot + 0.2 * h);
        GrayImage band = new GrayImage(Math.Max(1, bx1 - bx0 + 1), Math.Max(1, by1 - by0 + 1));
        for (int y = 0; y < band.H; y++) {
            for (int x = 0; x < band.W; x++) {
                double xs = bx0 + x, yf = by0 + y + tilt * (xs - xref);
                double xf = xs - sh * yf;
                band.P[y * band.W + x] = xf < 0 || xf > full.W - 1 || yf < 0 || yf > full.H - 1 ? (byte)255 : (byte)full.Sample(xf, yf);
            }
        }
        bool[] dk = Dark(band, Math.Max(4, (int)(0.6 * h)));
        int W = band.W, H = band.H;
        double tp = top - by0, bs = bot - by0;
        if (DebugDump != null && h > DebugH0 && h < DebugH1) {
            try {
                using (Bitmap bm = new Bitmap(W, H * 2)) {
                    for (int y = 0; y < H; y++) for (int x = 0; x < W; x++) {
                        int v = band.P[y * W + x];
                        bm.SetPixel(x, y, Color.FromArgb(v, v, v));
                        bm.SetPixel(x, H + y, dk[y * W + x] ? Color.Black : Color.White);
                    }
                    bm.Save(DebugDump + "_" + sh.ToString("0.00") + "_" + bx0 + "_" + by0 + ".png");
                }
            } catch { }
        }
        int c0x = core0 - bx0, c1x = core1 - bx0;
        // остаток наклона: при верном - меньше всего занятых столбцов
        int ya = (int)(tp + 0.04 * h), yb = (int)(bs - 0.24 * h);
        double slant = 0;
        int bestCols = int.MaxValue;
        for (double sl = -0.08; sl <= 0.0801; sl += 0.02) {
            bool[] col = new bool[W + 2 * H];
            for (int y = ya; y <= yb; y++) {
                if (y < 0 || y >= H) continue;
                int d = (int)Math.Round(sl * (bs - y));
                for (int x = c0x; x < c1x && x < W; x++) if (x >= 0 && dk[y * W + x]) { int xx = x - d + H; if (xx >= 0 && xx < col.Length) col[xx] = true; }
            }
            int cnt = 0;
            foreach (bool b in col) if (b) cnt++;
            if (cnt < bestCols) { bestCols = cnt; slant = sl; }
        }
        // проекция на столбцы (без нижней полосы, где точка)
        int[] proj = new int[W];
        for (int y = ya; y <= yb; y++) {
            if (y < 0 || y >= H) continue;
            for (int x = 0; x < W; x++) {
                int xi = (int)Math.Round(x + slant * (bs - y));
                if (xi >= 0 && xi < W && dk[y * W + xi]) proj[x]++;
            }
        }
        int minInk = Math.Max(2, (int)(0.05 * h));
        List<int[]> cells = new List<int[]>();
        for (int x = 0; x < W; ) {
            if (proj[x] < minInk) { x++; continue; }
            int s0 = x;
            while (x < W && proj[x] >= minInk) x++;
            if (cells.Count > 0 && s0 - cells[cells.Count - 1][1] < 0.04 * h) cells[cells.Count - 1][1] = x;
            else cells.Add(new int[] { s0, x });
        }
        // ячейки - цифры: по высоте почти во всю строку; цепочка вокруг ядра ряда без больших разрывов
        List<int[]> dig = new List<int[]>();
        List<int> digTop = new List<int>(), digBot = new List<int>();
        foreach (int[] c in cells) {
            int y0 = H, y1 = -1;
            for (int y = 0; y < H; y++)
                for (int x = c[0]; x < c[1]; x++) {
                    int xi = (int)Math.Round(x + slant * (bs - y));
                    if (xi >= 0 && xi < W && dk[y * W + xi]) { if (y < y0) y0 = y; if (y > y1) y1 = y; break; }
                }
            if (y1 - y0 < 0.55 * h || y0 > tp + 0.25 * h || y1 < bs - 0.3 * h) continue;
            if (c[1] - c[0] > 1.3 * h) continue;
            // цифра стоит отдельно: тёмное не тянется сплошь за верх или низ строки (рамка табло, линии - тянутся)
            bool[] rowInk = new bool[H];
            for (int y = 0; y < H; y++)
                for (int x = c[0]; x < c[1]; x++) {
                    int xi = (int)Math.Round(x + slant * (bs - y));
                    if (xi >= 0 && xi < W && dk[y * W + xi]) { rowInk[y] = true; break; }
                }
            int yu = (int)(tp + 0.25 * h), yd = (int)(bs - 0.25 * h);
            if (yu >= 0 && yu < H && rowInk[yu]) { while (yu > 0 && (rowInk[yu - 1] || yu > 1 && rowInk[yu - 2])) yu--; } else yu = H;
            if (yd >= 0 && yd < H && rowInk[yd]) { while (yd < H - 1 && (rowInk[yd + 1] || yd < H - 2 && rowInk[yd + 2])) yd++; } else yd = -1;
            if (yu <= tp - 0.12 * h || yd >= bs + 0.12 * h) continue;
            dig.Add(c);
            // верх и низ самой цифры - сплошной тёмный участок от середины (без значков над и под строкой)
            digTop.Add(yu < H ? yu : y0); digBot.Add(yd >= 0 ? yd : y1);
        }
        if (dig.Count == 0) return null;
        // ядро: ячейки, пересекающиеся с рядом пятен; дальше - соседи с разрывом не больше высоты строки
        int k0 = -1, k1 = -1;
        for (int k = 0; k < dig.Count; k++) if (dig[k][1] > c0x && dig[k][0] < c1x) { if (k0 < 0) k0 = k; k1 = k; }
        if (k0 < 0) return null;
        while (k0 > 0 && dig[k0][0] - dig[k0 - 1][1] < 0.8 * h) k0--;
        while (k1 < dig.Count - 1 && dig[k1 + 1][0] - dig[k1][1] < 0.8 * h) k1++;
        dig = dig.GetRange(k0, k1 - k0 + 1);
        n = dig.Count;
        // верх и низ строки - по самим цифрам: самая частая высота верха (у «1», «4», «7» верх ниже) и низа
        tp = Cluster(digTop.GetRange(k0, n), 0.08 * h, true);
        bs = Cluster(digBot.GetRange(k0, n), 0.08 * h, false) + 1;
        h = bs - tp;
        if (h < 8) return null;
        List<double> widths = new List<double>();
        foreach (int[] c in dig) widths.Add(c[1] - c[0]);
        widths.Sort();
        double maxW = widths[widths.Count - 1];
        List<double> wide = new List<double>();
        foreach (double x in widths) if (x > 0.6 * maxW) wide.Add(x);
        double Wt = wide[wide.Count / 2];
        // толщина штриха: тёмные отрезки поперёк вертикальных сегментов
        List<int> runs = new List<int>();
        foreach (int[] c in dig)
            foreach (double v in new double[] { 0.3, 0.7 }) {
                int y = (int)(tp + v * h), run = 0;
                if (y < 0 || y >= H) continue;
                for (int x = c[0] - 2; x <= c[1] + 2; x++) {
                    int xi = (int)Math.Round(x + slant * (bs - y));
                    bool d = xi >= 0 && xi < W && dk[y * W + xi];
                    if (d) run++;
                    else { if (run > 0) runs.Add(run); run = 0; }
                }
                if (run > 0) runs.Add(run);
            }
        runs.Sort();
        double t = runs.Count > 0 ? runs[runs.Count / 2] : 0.14 * h;
        t = Math.Max(0.07 * h, Math.Min(0.25 * h, t));
        // узкая ячейка - «1» (два сегмента во всю высоту); мало тёмного - соринка или блик, не цифра
        for (int k = dig.Count - 1; k >= 0; k--) {
            int[] c = dig[k];
            if (c[1] - c[0] >= 0.6 * Wt) continue;
            int ink = 0;
            for (int x = c[0]; x < c[1]; x++) ink += proj[x];
            if (ink < 0.45 * (yb - ya) * t) { dig.RemoveAt(k); digTop.RemoveAt(k0 + k); digBot.RemoveAt(k0 + k); }
        }
        n = dig.Count;
        if (n == 0) return null;
        if (Debug) Console.WriteLine("row n=" + n + " sh=" + sh + " h=" + h.ToString("0") + " slant=" + slant.ToString("0.00") + " Wt=" + Wt.ToString("0") + " t=" + t.ToString("0.0") + " band=" + bx0 + "," + by0 + " " + W + "x" + H);
        Digits res = new Digits();
        res.Top = by0 + tp; res.Base = by0 + bs; res.H = h; res.Left = bx0 + dig[0][0] - sh * res.Base;
        res.Tilt = tilt;
        res.X0 = bx0 + dig[0][0] - sh * res.Base; res.X1 = bx0 + dig[n - 1][1] - sh * res.Top; res.Y0 = res.Top; res.Y1 = res.Base;
        // уточнение поворота: прямая по низу цифр во всю высоту
        if (n >= 3) {
            double sx = 0, sy = 0, sxx = 0, sxy = 0;
            int m = 0;
            for (int k = 0; k < n; k++) {
                double yb2 = digBot[k0 + k];
                if (Math.Abs(yb2 - bs) > 0.2 * h) continue;
                double x = (dig[k][0] + dig[k][1]) / 2.0;
                sx += x; sy += yb2; sxx += x * x; sxy += x * yb2; m++;
            }
            double den = m * sxx - sx * sx;
            if (m >= 3 && den > 0) res.TiltFix = (m * sxy - sx * sy) / den;
        }
        StringBuilder sb = new StringBuilder();
        double[] cl = new double[n], cr = new double[n];
        bool[] sure = new bool[n];
        string segs = "abcdefg";
        for (int k = 0; k < n; k++) {
            double c0 = dig[k][0], c1 = dig[k][1];
            if (c1 - c0 < 0.6 * Wt) c0 = c1 - Wt;           // «1» - только правые сегменты
            cl[k] = c0; cr[k] = c1;
            double[] fr = new double[7];
            for (int q = 0; q < 7; q++) {
                double u0, u1, v0, v1;   // u - по ширине (точки), v - по высоте (точки от верха цифры)
                char sg = segs[q];
                if (sg == 'a' || sg == 'd' || sg == 'g') {
                    u0 = c0 + 0.3 * (c1 - c0); u1 = c1 - 0.3 * (c1 - c0);
                    double vc = sg == 'a' ? t / 2 : sg == 'd' ? h - t / 2 : h / 2;
                    v0 = vc - 0.35 * t; v1 = vc + 0.35 * t;
                } else {
                    bool left = sg == 'e' || sg == 'f', upper = sg == 'b' || sg == 'f';
                    double uc = left ? c0 + t / 2 : c1 - t / 2;
                    u0 = uc - 0.35 * t; u1 = uc + 0.35 * t;
                    v0 = upper ? 0.2 * h : 0.62 * h; v1 = upper ? 0.38 * h : 0.8 * h;
                }
                int on = 0, all = 0;
                for (double v = v0; v <= v1; v += Math.Max(1, (v1 - v0) / 6))
                    for (double u = u0; u <= u1; u += Math.Max(1, (u1 - u0) / 6)) {
                        int y = (int)Math.Round(tp + v), x = (int)Math.Round(u + slant * (h - v));
                        all++;
                        if (x >= 0 && x < W && y >= 0 && y < H && dk[y * W + x]) on++;
                    }
                fr[q] = all == 0 ? 0 : (double)on / all;
            }
            // цифра, которой лучше всего соответствуют сегменты
            int best = -1;
            double b1 = -1, b2 = -1;
            for (int dgt = 0; dgt < 10; dgt++) {
                double sc = 0;
                for (int q = 0; q < 7; q++) sc += Seg7[dgt].IndexOf(segs[q]) >= 0 ? Math.Min(1, fr[q] / 0.6) : Math.Min(1, (1 - fr[q]) / 0.6);
                sc /= 7;
                if (sc > b1) { b2 = b1; b1 = sc; best = dgt; } else if (sc > b2) b2 = sc;
            }
            bool ok = b1 >= 0.85 && b1 - b2 >= 0.1;
            // всё тёмное в ячейке должно лежать на сегментах этой цифры (иначе это не семисегментная цифра: буква, штрихкод, рамка)
            if (best >= 0) {
                int dark = 0, extra = 0;
                for (int y = (int)tp; y <= (int)bs; y++) {
                    if (y < 0 || y >= H) continue;
                    double v = y - tp;
                    for (int x = (int)Math.Floor(Math.Min(c0, dig[k][0])); x <= (int)Math.Ceiling(c1); x++) {
                        int xi = (int)Math.Round(x + slant * (bs - y));
                        if (xi < 0 || xi >= W || !dk[y * W + xi]) continue;
                        dark++;
                        bool inSeg = false;
                        foreach (char sg in Seg7[best]) {
                            bool z;
                            if (sg == 'a') z = v <= 1.4 * t;
                            else if (sg == 'd') z = v >= h - 1.4 * t;
                            else if (sg == 'g') z = Math.Abs(v - h / 2) <= 0.9 * t;
                            else {
                                bool left = sg == 'e' || sg == 'f', upper = sg == 'b' || sg == 'f';
                                z = (left ? x <= c0 + 1.4 * t : x >= c1 - 1.4 * t) && (upper ? v <= h / 2 + 0.4 * t : v >= h / 2 - 0.4 * t);
                            }
                            if (z) { inSeg = true; break; }
                        }
                        if (!inSeg) extra++;
                    }
                }
                if (dark == 0 || extra > 0.12 * dark) b1 = Math.Min(b1, 0.6);
                if (best == 1 && dig[k][1] - dig[k][0] > 2.2 * t) b1 = Math.Min(b1, 0.6);   // «1» - узкая
                // каждый сегмент цифры виден хотя бы наполовину (иначе это обрывок линии или надписи)
                for (int q = 0; q < 7; q++) if (Seg7[best].IndexOf(segs[q]) >= 0 && fr[q] < 0.45) b1 = Math.Min(b1, 0.65);
            }
            if (Debug) { StringBuilder db = new StringBuilder(); foreach (double x in fr) db.Append(x.ToString("0.00") + " "); Console.WriteLine("  cell " + c0.ToString("0") + ".." + c1.ToString("0") + " fr " + db + " -> " + best + " " + b1.ToString("0.00") + "/" + b2.ToString("0.00")); }
            if (b1 < 0.7) { sb.Append('?'); res.Bad++; }
            else { sb.Append((char)('0' + best)); sure[k] = ok; }
            if (b1 >= 0.7) res.Score += b1;
            res.Count++;
        }
        // десятичная точка: тёмное пятнышко у нижнего края в промежутке между цифрами, над ним пусто
        int dotAfter = -1;
        double dotInk = 0;
        for (int k = 0; k + 1 < n; k++) {
            double g0 = cr[k] - 0.2 * t, g1 = cl[k + 1] + 0.2 * t;
            int on = 0, all = 0, above = 0, allA = 0;
            for (double v = h - 1.2 * t; v <= h + 0.1 * t; v += 1)
                for (double u = g0; u <= g1; u += 1) {
                    int y = (int)Math.Round(tp + v), x = (int)Math.Round(u + slant * (h - v));
                    all++;
                    if (x >= 0 && x < W && y >= 0 && y < H && dk[y * W + x]) on++;
                }
            for (double v = h - 2.8 * t; v <= h - 1.8 * t; v += 1)
                for (double u = cr[k] + 0.2 * t; u <= cl[k + 1] - 0.2 * t; u += 1) {
                    int y = (int)Math.Round(tp + v), x = (int)Math.Round(u + slant * (h - v));
                    allA++;
                    if (x >= 0 && x < W && y >= 0 && y < H && dk[y * W + x]) above++;
                }
            double ink = all == 0 ? 0 : (double)on / all, inkA = allA == 0 ? 0 : (double)above / allA;
            // в промежутке без точки внизу ничего (сегменты d соседей в промежуток не заходят)
            double extra = ink - 0.0;
            if (extra > 0.12 && inkA < 0.1 && ink > dotInk) { dotInk = ink; dotAfter = k; }
        }
        string txt = sb.ToString();
        if (dotAfter >= 0) txt = txt.Substring(0, dotAfter + 1) + "." + txt.Substring(dotAfter + 1);
        // неразобранные ячейки по краям - не цифры табло (рамка, надписи рядом)
        int cut0 = 0, cut1 = 0;
        while (cut0 < txt.Length && txt[cut0] == '?') cut0++;
        while (cut1 < txt.Length - cut0 && txt[txt.Length - 1 - cut1] == '?') cut1++;
        res.Bad -= cut0 + cut1;
        res.Count -= cut0 + cut1;
        res.Text = txt.Substring(cut0, txt.Length - cut0 - cut1).Trim('.');
        // уверенно: все оставшиеся цифры разобраны чётко
        res.Sure = res.Bad == 0 && res.Count > 0;
        string digitsOnly = sb.ToString();
        for (int k = 0; k < n; k++) {
            if (k < cut0 || k >= n - cut1) continue;
            if (digitsOnly[k] != '?' && !sure[k]) res.Sure = false;
        }
        return res;
    }

    // сдвиг строк маски: x' = x + sh*y (наклонные цифры становятся прямыми)
    static bool[] Shear(bool[] m, int w, int h, double sh, out int w2) {
        w2 = w + (int)Math.Ceiling(sh * h) + 1;
        bool[] o = new bool[w2 * h];
        for (int y = 0; y < h; y++) {
            int d = (int)Math.Round(sh * y), a = y * w, b = y * w2 + d;
            for (int x = 0; x < w; x++) if (m[a + x]) o[b + x] = true;
        }
        return o;
    }

    // сегменты одной цифры - в одно пятно (по высоте зазор больше: у «7», «1», «0» верх и низ разделены местом среднего сегмента)
    static List<Blob> Group(bool[] m, int w, int h) {
        List<Mrz.Comp> cs = new List<Mrz.Comp>();
        foreach (Mrz.Comp c in Mrz.Components(m, w, h))
            if (c.Ht >= 4 && c.Ht < h / 4 && c.Wd < w / 4 && c.N >= 8) cs.Add(c);
        cs.Sort(delegate(Mrz.Comp a, Mrz.Comp b) { return a.X0.CompareTo(b.X0); });
        int nc = cs.Count;
        int[] par = new int[nc];
        double[] th = new double[nc];
        for (int i = 0; i < nc; i++) { par[i] = i; th[i] = (double)cs[i].N / Math.Max(cs[i].Wd, cs[i].Ht); }
        for (int i = 0; i < nc; i++) {
            Mrz.Comp a = cs[i];
            bool dotA = a.Wd < 2.2 * th[i] && a.Ht < 2.2 * th[i];
            for (int j = i + 1; j < nc; j++) {
                Mrz.Comp b = cs[j];
                if (b.X0 > a.X1 + 2 * th[i] + 2) break;
                bool dotB = b.Wd < 2.2 * th[j] && b.Ht < 2.2 * th[j];
                if (dotA || dotB) continue;                      // точка между цифрами склеила бы соседей
                double t = Math.Min(th[i], th[j]);
                if (Math.Max(th[i], th[j]) > 2.5 * t) continue;
                double gx = Math.Max(b.X0 - a.X1, a.X0 - b.X1), gy = Math.Max(b.Y0 - a.Y1, a.Y0 - b.Y1);
                if (gx <= 0.5 * t + 1 && gy <= 1.8 * t + 1) {
                    int ra = i, rb = j;
                    while (par[ra] != ra) ra = par[ra];
                    while (par[rb] != rb) rb = par[rb];
                    if (ra != rb) par[Math.Max(ra, rb)] = Math.Min(ra, rb);
                }
            }
        }
        Dictionary<int, Blob> groups = new Dictionary<int, Blob>();
        for (int i = 0; i < nc; i++) {
            int ri = i;
            while (par[ri] != ri) ri = par[ri];
            Blob b;
            if (!groups.TryGetValue(ri, out b)) { b = new Blob(); b.X0 = cs[i].X0; b.Y0 = cs[i].Y0; b.X1 = cs[i].X1; b.Y1 = cs[i].Y1; groups[ri] = b; }
            b.X0 = Math.Min(b.X0, cs[i].X0); b.Y0 = Math.Min(b.Y0, cs[i].Y0); b.X1 = Math.Max(b.X1, cs[i].X1); b.Y1 = Math.Max(b.Y1, cs[i].Y1);
        }
        List<Blob> bl = new List<Blob>();
        foreach (Blob b in groups.Values) {
            if (b.Ht < 12 || b.Ht > h / 4) continue;
            double ar = (double)b.Ht / b.Wd;
            if (ar < 1.0 || ar > 12) continue;
            bl.Add(b);
        }
        return bl;
    }

    static double Height(List<Blob> row) {
        double h = 0;
        foreach (Blob b in row) h = Math.Max(h, b.Ht);
        return h;
    }

    // показания с табло: строка из 5-10 крупных цифр; над ней слева - код экрана («1.8.0» - сумма)
    static Digits Lcd(GrayImage full, GrayImage small, out Digits code) {
        code = null;
        double f = (double)full.W / small.W;
        bool[] dk = Dark(small, Math.Max(10, Math.Min(small.W, small.H) / 40));
        Digits best = null;
        double bestSh = 0;
        foreach (double sh in new double[] { 0, 0.08, 0.16, 0.24, 0.32 }) {
            List<Digits> done = new List<Digits>();        // повторы - только при этом же сдвиге (при другом читается иначе)
            int w2;
            bool[] m = Shear(dk, small.W, small.H, sh, out w2);
            List<Blob> bl = Group(m, w2, small.H);
            // предложения: ряды пятен (хотя бы два) и отдельные пятна, похожие на цифру; крупные - первыми
            List<List<Blob>> props = Rows(bl, 2);
            foreach (Blob b in bl) {
                double ar = (double)b.Ht / b.Wd;
                if (ar < 1.2 || ar > 3.5 || b.Ht < 14) continue;
                List<Blob> one = new List<Blob>();
                one.Add(b);
                props.Add(one);
            }
            props.Sort(delegate(List<Blob> a, List<Blob> b) { return Height(b).CompareTo(Height(a)); });
            if (Debug) Console.WriteLine("sh=" + sh + " blobs " + bl.Count + " proposals " + props.Count);
            foreach (List<Blob> row in props) {
                if (row.Count > 11) continue;
                // уже прочитанное место (при этом или другом сдвиге) - не повторять
                double cx = ((row[0].X0 + row[row.Count - 1].X1) / 2.0 - sh * (row[0].Y0 + row[0].Y1) / 2.0) * f;
                double cy = (row[0].Y0 + row[0].Y1) / 2.0 * f, hh = Height(row) * f;
                bool seen = false;
                foreach (Digits d0 in done)
                    if (cx >= d0.X0 - 0.3 * d0.H && cx <= d0.X1 + 0.3 * d0.H && cy >= d0.Y0 - 0.3 * d0.H && cy <= d0.Y1 + 0.3 * d0.H && hh < 1.4 * d0.H && hh > 0.5 * d0.H) seen = true;
                if (seen) continue;
                Digits d = ReadRow(full, row, f, sh);
                if (d == null || d.Count == 0) continue;
                // строка повёрнута сильнее, чем казалось по пятнам, - прочитать ещё раз с поправкой
                if (Math.Abs(d.TiltFix) > 0.015) {
                    Digits d2 = ReadRow(full, row, f, sh, d.Tilt + d.TiltFix);
                    if (d2 != null && d2.Score > d.Score) d = d2;
                }
                done.Add(d);
                if (Debug) Console.WriteLine("ROW " + d.Text + " score " + d.Score.ToString("0.0") + " n " + d.Count + " bad " + d.Bad + " h " + d.H.ToString("0"));
                // качество: доля уверенных цифр, размер; печатные строки разбираются плохо
                if (d.Count < 4) continue;
                double q = d.Score / d.Count;
                if (q < 0.75 || d.Bad > 1) continue;
                if (best == null || d.Score * d.H > best.Score * best.H) { best = d; bestSh = sh; }
            }
        }
        if (best == null) return null;
        // код экрана («1.8.0»): строка мелких цифр над основным рядом, у левого края. Мелкие цифры с точками
        // слипаются, поэтому строку ищем по тёмному в полосе над рядом, а цифры в ней делит ReadRow
        double left = (best.Left + bestSh * best.Base) / f, topS = best.Top / f, hs = best.H / f;   // в выпрямленных координатах
        int xa = (int)(left - 0.5 * hs), xb = (int)(left + 2.5 * hs);
        int yA = (int)Math.Max(0, topS - 1.6 * hs), yB = (int)(topS - 0.05 * hs);
        if (yB > yA) {
            int[] ink = new int[yB - yA + 1];
            for (int y = yA; y <= yB; y++)
                for (int x = xa; x <= xb; x++) {
                    int xi = (int)Math.Round(x - bestSh * y);
                    if (xi >= 0 && xi < small.W && dk[y * small.W + xi]) ink[y - yA]++;
                }
            int need = Math.Max(2, (xb - xa) / 40);
            int yb = yB;
            while (yb >= yA && ink[yb - yA] < need) yb--;
            int yt = yb, gap = 0;
            while (yt > yA && gap <= 0.06 * hs) { yt--; if (ink[yt - yA] < need) gap++; else gap = 0; }
            yt += gap;
            double hc = yb - yt + 1;
            if (yb > yA && hc >= 0.25 * hs && hc <= 0.85 * hs) {
                Blob cb = new Blob();
                cb.X0 = xa; cb.X1 = xb; cb.Y0 = yt; cb.Y1 = yb;
                List<Blob> crow = new List<Blob>();
                crow.Add(cb);
                code = ReadRow(full, crow, f, bestSh);
            }
            if (Debug) Console.WriteLine("code band " + yt + ".." + yb + " hs " + hs.ToString("0") + " -> " + (code == null ? "-" : code.Text));
        }
        return best;
    }

    // для проверки: текст табло и кода экрана («?» в конце - неуверенно)
    public static string LcdText(GrayImage full, GrayImage small, out string code) {
        Digits c;
        Digits d = Lcd(full, small, out c);
        code = c == null ? "" : c.Text + (c.Sure ? "" : "?");
        return d == null ? "" : d.Text + (d.Sure ? "" : "?");
    }

    // ---------- тип: надпись модификации («R32.043.JA.UVFL») ----------

    public static string[] Families = { "Arial", "Arial Narrow", "Tahoma", "Verdana", "Segoe UI", "Calibri", "Trebuchet MS", "Microsoft Sans Serif" };
    const string LatinChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ.";
    static GlyphSet latin;
    static readonly object sync = new object();

    static void Init() {
        lock (sync) {
            if (latin == null) latin = GlyphSet.FromFonts(LatinChars, Families, new FontStyle[] { FontStyle.Regular, FontStyle.Bold });
        }
    }

    // для сравнения: только буквы и цифры, похожие знаки - одинаково (O, U и 0, I и 1, B и 8, S и 5, Z и 2, Y и V)
    static string Key(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c0 in s.ToUpperInvariant()) {
            char c = c0;
            int k = "АВЕКМНОРСТХ".IndexOf(c);
            if (k >= 0) c = "ABEKMHOPCTX"[k];
            if (!char.IsLetterOrDigit(c)) continue;
            if (c == 'O' || c == 'Q' || c == 'D' || c == 'U') c = '0';
            else if (c == 'I') c = '1';
            else if (c == 'B') c = '8';
            else if (c == 'S') c = '5';
            else if (c == 'Z') c = '2';
            else if (c == 'Y') c = 'V';
            sb.Append(c);
        }
        return sb.ToString();
    }

    // наименьшее число правок, чтобы key встретился где-то внутри text
    static int Within(string text, string key) {
        int m = key.Length;
        int[] prev = new int[m + 1], cur = new int[m + 1];
        for (int j = 0; j <= m; j++) prev[j] = j;
        int best = prev[m];
        foreach (char ch in text) {
            cur[0] = 0;
            for (int j = 1; j <= m; j++)
                cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + (key[j - 1] == ch ? 0 : 1));
            if (cur[m] < best) best = cur[m];
            int[] t = prev; prev = cur; cur = t;
        }
        return best;
    }

    // тип по надписи: строки снимка читаются латиницей и цифрами и сравниваются с модификациями справочника
    public static string ReadType(GrayImage full, GrayImage small, List<string> types, out bool ok, out string seen) {
        Init();
        ok = false;
        seen = "";
        List<string> keys = new List<string>();
        foreach (string t in types) {
            int sp = t.IndexOf(' ');
            keys.Add(Key(sp < 0 ? t : t.Substring(sp + 1)));
        }
        double f = (double)full.W / small.W;
        int best = -1, d1 = 99, d2 = 99;
        foreach (PassportPage.Line l in PassportPage.Scale(PassportPage.FindLines(small, false), f)) {
            if (l.Comps.Count < 8 || l.Comps.Count > 60) continue;
            PassportPage.Strip st = PassportPage.ReadLine(full, l, latin);
            if (st == null || st.G == null) continue;
            StringBuilder sb = new StringBuilder();
            foreach (Glyph g in st.G) sb.Append(g.C);
            string text = Key(sb.ToString());
            if (text.Length < 6) continue;
            for (int i = 0; i < keys.Count; i++) {
                if (keys[i].Length == 0) continue;
                int d = Within(text, keys[i]);
                if (d < d1) {
                    if (best >= 0 && keys[best] != keys[i]) d2 = d1;
                    d1 = d; best = i; seen = sb.ToString();
                } else if (d < d2 && (best < 0 || keys[i] != keys[best])) d2 = d;
            }
        }
        if (best < 0 || d1 > Math.Max(2, keys[best].Length / 5)) return "";
        // уверенно: совпало точно (или с одной ошибкой у длинного обозначения), а следующий вариант заметно дальше
        ok = d1 == 0 && d2 >= 1 || d1 == 1 && d2 >= 3 && keys[best].Length >= 10;
        return types[best];
    }

    // ---------- всё вместе ----------

    public static MeterResult Read(string path, List<string> types) {
        MeterResult r = new MeterResult();
        GrayImage full, small;
        Mrz.LoadPair(path, out full, out small);
        int rows;
        string year;
        string serial = Barcode(full, out rows, out year);
        // заводской номер - только цифры (6-15)
        bool digits = serial.Length >= 6 && serial.Length <= 15;
        foreach (char c in serial) if (!char.IsDigit(c)) digits = false;
        if (digits) { r.Serial = serial; r.SerialRows = rows; r.Year = year; }
        Digits code;
        Digits d = Lcd(full, small, out code);
        if (d != null) {
            // показания: не меньше 6 цифр, после точки 2-3 знака
            string t = d.Text;
            int dot = t.IndexOf('.'), nd = 0;
            foreach (char c in t) if (char.IsDigit(c) || c == '?') nd++;
            int dec = dot < 0 ? 0 : t.Length - dot - 1;
            if (nd >= 6 && nd <= 10 && dot > 0 && dec >= 2 && dec <= 3) {
                r.Reading = t.Replace('.', ',');
                string cd = code == null ? "" : code.Text.Replace(".", "");
                if (cd.Length == 3) r.Code = cd[0] + "." + cd[1] + "." + cd[2];
                r.ReadingOk = d.Sure && t.IndexOf('?') < 0 && r.Code == "1.8.0" && code.Sure;
            }
        }
        if (r.Serial.Length > 0 || r.Reading.Length > 0) {
            bool ok;
            string seen;
            r.Type = ReadType(full, small, types, out ok, out seen);
            r.TypeOk = ok;
            r.TypeText = seen;
        }
        // счётчик на снимке: табло прочитано целиком (печатные цифры документов так не читаются), или номер со штрихкода и тип
        bool lcd = r.Reading.Length > 0 && r.Reading.IndexOf('?') < 0 && d.H >= 0.015 * Math.Max(full.W, full.H);
        r.Found = lcd || r.Serial.Length > 0 && r.Type.Length > 0;
        return r;
    }
}

// Фото документов абонента (паспорт, свидетельство): у каждого договора своя папка с фото,
// миниатюры в форме, окно просмотра рядом с формой, перенос фото с телефона по кабелю.

public static class Photo {
    public static bool IsPhoto(string f) {
        string e = Path.GetExtension(f).ToLowerInvariant();
        return e == ".jpg" || e == ".jpeg" || e == ".png" || e == ".bmp";
    }

    // фото целиком (повёрнуто как снимали), не больше max точек по длинной стороне; файл не занимается
    public static Bitmap Open(string path, int max) { return Mrz.Load(path, max); }

    // миниатюра: встроенная в снимок (быстро), если её нет - из самого снимка
    public static Bitmap Thumb(string path, int max) {
        byte[] bytes = File.ReadAllBytes(path);
        lock (Mrz.Gdi)
        using (MemoryStream ms = new MemoryStream(bytes))
        using (Image img = Image.FromStream(ms, false, false)) {
            int orient = Mrz.Orientation(img);
            Image small = null;
            MemoryStream ts = null;
            try {
                foreach (int id in img.PropertyIdList)
                    if (id == 0x501B) {
                        ts = new MemoryStream(img.GetPropertyItem(id).Value);
                        small = Image.FromStream(ts);
                    }
            } catch { small = null; }
            // встроенная миниатюра другой формы или слишком мелкая - не годится
            if (small != null && (Math.Abs((double)small.Width / small.Height - (double)img.Width / img.Height) > 0.05 || Math.Max(small.Width, small.Height) < max / 2)) {
                small.Dispose();
                small = null;
            }
            try {
                Image src = small ?? img;
                double k = Math.Min(1.0, (double)max / Math.Max(src.Width, src.Height));
                return Mrz.Draw(src, Math.Max(1, (int)(src.Width * k)), Math.Max(1, (int)(src.Height * k)), orient);
            } finally {
                if (small != null) small.Dispose();
                if (ts != null) ts.Dispose();
            }
        }
    }

    // свободное имя файла в папке: IMG_1.jpg, IMG_1 (2).jpg, ...
    public static string Free(string dir, string name) {
        string to = Path.Combine(dir, name);
        string b = Path.GetFileNameWithoutExtension(name), e = Path.GetExtension(name);
        for (int i = 2; File.Exists(to); i++) to = Path.Combine(dir, b + " (" + i + ")" + e);
        return to;
    }
}

public delegate void PhotoHandler(string path);

// Миниатюры фото договора; щелчок - открыть, правая кнопка - меню
public class PhotoStrip : Control {
    public const int TW = 150, TH = 112, CapH = 22, Gap = 12;
    readonly List<string> files = new List<string>();
    readonly Dictionary<string, Bitmap> thumbs = new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, string> marks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, Color> markColors = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
    readonly Queue<string> queue = new Queue<string>();
    System.Threading.Thread loader;
    int hover = -1;
    ContextMenuStrip menu;
    string menuFile;
    public string Hint = "";
    public event PhotoHandler OpenPhoto, RemovePhoto;
    public event EventHandler Resized;

    public PhotoStrip() {
        BackColor = St.Card;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }

    public List<string> Files { get { return new List<string>(files); } }

    public void SetFiles(List<string> list) {
        int before = Rows(Width);
        files.Clear();
        files.AddRange(list);
        // миниатюры исчезнувших фото - выбросить
        List<string> gone = new List<string>();
        foreach (string k in thumbs.Keys) if (!files.Contains(k)) gone.Add(k);
        foreach (string k in gone) { thumbs[k].Dispose(); thumbs.Remove(k); }
        lock (queue) {
            queue.Clear();
            foreach (string f in files) if (!thumbs.ContainsKey(f)) queue.Enqueue(f);
            if (queue.Count > 0) {
                if (loader == null) {
                    loader = new System.Threading.Thread(Load);
                    loader.IsBackground = true;
                    loader.Priority = System.Threading.ThreadPriority.BelowNormal;
                    loader.Start();
                } else System.Threading.Monitor.Pulse(queue);
            }
        }
        hover = -1;
        Invalidate();
        if (Rows(Width) != before && Resized != null) Resized(this, EventArgs.Empty);
    }

    // пометка под миниатюрой (по имени файла): «личный номер прочитан» и т.п.
    public void Mark(string path, string text, Color color) {
        string n = Path.GetFileName(path);
        if (text == null) { marks.Remove(n); markColors.Remove(n); }
        else { marks[n] = text; markColors[n] = color; }
        Invalidate();
    }

    public void ClearMarks() { marks.Clear(); markColors.Clear(); Invalidate(); }

    void Load() {
        while (true) {
            string f;
            lock (queue) {
                while (queue.Count == 0) System.Threading.Monitor.Wait(queue);
                f = queue.Dequeue();
            }
            Bitmap b = null;
            try { b = Photo.Thumb(f, 2 * TW); } catch { }
            string file = f;
            Bitmap bmp = b;
            try {
                if (IsHandleCreated && !IsDisposed)
                    BeginInvoke((MethodInvoker)delegate {
                        if (bmp == null) return;
                        if (!files.Contains(file) || thumbs.ContainsKey(file)) { bmp.Dispose(); return; }
                        thumbs[file] = bmp;
                        Invalidate();
                    });
                else if (b != null) b.Dispose();
            } catch { if (b != null) b.Dispose(); }
        }
    }

    int Cols(int width) { return Math.Max(1, (width + Gap) / (TW + Gap)); }
    int Rows(int width) { return files.Count == 0 ? 0 : (files.Count + Cols(width) - 1) / Cols(width); }

    public int HeightFor(int width) {
        if (files.Count == 0) return 46;
        return Rows(width) * (TH + CapH + Gap) - Gap + 2;
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        Invalidate();
    }

    Rectangle Tile(int i) {
        int c = Cols(Width);
        return new Rectangle((i % c) * (TW + Gap), (i / c) * (TH + CapH + Gap), TW, TH);
    }

    int At(Point p) {
        for (int i = 0; i < files.Count; i++) {
            Rectangle r = Tile(i);
            r.Height += CapH;
            if (r.Contains(p)) return i;
        }
        return -1;
    }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        if (files.Count == 0) {
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (Pen p = new Pen(St.InputBorder)) { p.DashStyle = DashStyle.Dash; g.DrawRectangle(p, r); }
            TextRenderer.DrawText(g, Hint, St.Small, Rectangle.Inflate(r, -12, 0), St.Hint, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.WordBreak);
            return;
        }
        for (int i = 0; i < files.Count; i++) {
            Rectangle r = Tile(i);
            if (!e.ClipRectangle.IntersectsWith(new Rectangle(r.X, r.Y, r.Width + 1, r.Height + CapH + 1))) continue;
            using (SolidBrush b = new SolidBrush(St.Box)) g.FillRectangle(b, r);
            Bitmap t;
            if (thumbs.TryGetValue(files[i], out t)) {
                double k = Math.Min((double)(r.Width - 4) / t.Width, (double)(r.Height - 4) / t.Height);
                int w = (int)(t.Width * k), h = (int)(t.Height * k);
                g.InterpolationMode = InterpolationMode.HighQualityBilinear;
                g.DrawImage(t, r.X + (r.Width - w) / 2, r.Y + (r.Height - h) / 2, w, h);
            } else TextRenderer.DrawText(g, "загружаю...", St.Small, r, St.Hint, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            using (Pen p = new Pen(i == hover ? St.Accent : St.Border, i == hover ? 2f : 1f)) g.DrawRectangle(p, r.X, r.Y, r.Width - 1, r.Height - 1);
            string name = Path.GetFileName(files[i]), mark;
            Rectangle cap = new Rectangle(r.X, r.Bottom + 2, r.Width, CapH - 2);
            if (marks.TryGetValue(name, out mark))
                TextRenderer.DrawText(g, mark, St.Small, cap, markColors[name], TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
            else TextRenderer.DrawText(g, name, St.Small, cap, St.Hint, TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
        }
    }

    protected override void OnMouseMove(MouseEventArgs e) {
        base.OnMouseMove(e);
        int h = At(e.Location);
        Cursor = h >= 0 ? Cursors.Hand : Cursors.Default;
        if (h != hover) { hover = h; Invalidate(); }
    }

    protected override void OnMouseLeave(EventArgs e) {
        base.OnMouseLeave(e);
        if (hover >= 0) { hover = -1; Invalidate(); }
    }

    protected override void OnMouseWheel(MouseEventArgs e) { Ui.Wheel(this, e); }

    protected override void OnMouseUp(MouseEventArgs e) {
        base.OnMouseUp(e);
        int i = At(e.Location);
        if (i < 0) return;
        string f = files[i];
        if (e.Button == MouseButtons.Left) { if (OpenPhoto != null) OpenPhoto(f); }
        else if (e.Button == MouseButtons.Right) {
            if (menu == null) {
                menu = new ContextMenuStrip();
                menu.Font = St.Base;
                menu.Items.Add("Открыть", null, delegate { if (OpenPhoto != null) OpenPhoto(menuFile); });
                menu.Items.Add("Удалить фото...", null, delegate { if (RemovePhoto != null) RemovePhoto(menuFile); });
            }
            menuFile = f;
            menu.Show(this, e.Location);
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            foreach (Bitmap b in thumbs.Values) b.Dispose();
            thumbs.Clear();
            if (menu != null) menu.Dispose();
        }
        base.Dispose(disposing);
    }
}

// Окно просмотра фото: рядом с формой (можно печатать в форме, глядя на фото), крупнее/мельче колёсиком,
// двигать мышью, повернуть, листать
public class PhotoViewer : Form {
    readonly List<string> files = new List<string>();
    int index;
    Bitmap img;
    double scale = 1, ox, oy;
    bool fit = true, drag;
    Point last;
    readonly Panel bar = new Panel();
    readonly Label name = new Label(), hint = new Label();
    readonly Button prev, next;
    readonly View view = new View();
    readonly ToolTip tip = new ToolTip();

    class View : Control {
        public View() {
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.Selectable, true);
            BackColor = Color.FromArgb(46, 50, 56);
        }
    }

    public PhotoViewer() {
        Text = "Фото документов";
        Font = St.Base;
        BackColor = Color.White;
        ShowInTaskbar = false;
        MinimizeBox = false;
        KeyPreview = true;
        StartPosition = FormStartPosition.Manual;
        MinimumSize = new Size(420, 320);

        bar.Dock = DockStyle.Top;
        bar.Height = 50;
        bar.BackColor = Color.White;
        prev = Bar("‹", 40);
        next = Bar("›", 40);
        prev.Font = next.Font = new Font("Segoe UI", 14f, FontStyle.Bold);
        prev.Click += delegate { Go(index - 1); };
        next.Click += delegate { Go(index + 1); };
        Button turn = Bar("Повернуть", 100);
        turn.Click += delegate { Turn(); };
        Button fitB = Bar("Целиком", 86);
        fitB.Click += delegate { Fit(); };
        tip.SetToolTip(turn, "Повернуть на четверть оборота (R)");
        tip.SetToolTip(fitB, "Всё фото по размеру окна (0)");
        // правая кнопка: открыть в программе Windows (распечатать и т.п.)
        ContextMenuStrip menu = new ContextMenuStrip();
        menu.Font = St.Base;
        menu.Items.Add("Повернуть", null, delegate { Turn(); });
        menu.Items.Add("Целиком", null, delegate { Fit(); });
        menu.Items.Add("Открыть в программе Windows", null, delegate {
            if (index >= 0 && index < files.Count) try { System.Diagnostics.Process.Start(files[index]); } catch (Exception ex) { MessageBox.Show(this, ex.Message, Text); }
        });
        view.ContextMenuStrip = menu;
        name.AutoSize = false;
        name.AutoEllipsis = true;
        name.Font = St.Small;
        name.ForeColor = St.Text;
        name.TextAlign = ContentAlignment.MiddleLeft;
        hint.AutoSize = false;
        hint.Font = St.Small;
        hint.ForeColor = St.Hint;
        hint.Text = "колёсико - крупнее / мельче, мышью - двигать, правая кнопка - меню";
        hint.Dock = DockStyle.Bottom;
        hint.Height = 22;
        hint.TextAlign = ContentAlignment.MiddleLeft;
        hint.Padding = new Padding(8, 0, 0, 0);
        bar.Controls.Add(name);
        bar.Resize += delegate {
            int x = 10;
            foreach (Button b in new Button[] { prev, next }) { b.SetBounds(x, 8, b.Width, 34); x += b.Width + 6; }
            int r = bar.ClientSize.Width - 10;
            foreach (Button b in new Button[] { fitB, turn }) { r -= b.Width; b.SetBounds(r, 8, b.Width, 34); r -= 6; }
            name.SetBounds(x + 6, 8, Math.Max(10, r - x - 12), 34);
        };
        bar.Paint += delegate(object s, PaintEventArgs e) { using (Pen p = new Pen(St.Border)) e.Graphics.DrawLine(p, 0, bar.Height - 1, bar.Width, bar.Height - 1); };

        view.Dock = DockStyle.Fill;
        view.Paint += delegate(object s, PaintEventArgs e) { Draw(e.Graphics); };
        view.Resize += delegate { if (fit) Fit(); else view.Invalidate(); };
        view.MouseWheel += delegate(object s, MouseEventArgs e) { Zoom(e.Delta > 0 ? 1.25 : 0.8, e.Location); };
        view.MouseDown += delegate(object s, MouseEventArgs e) { view.Focus(); if (e.Button == MouseButtons.Left) { drag = true; last = e.Location; view.Cursor = Cursors.SizeAll; } };
        view.MouseMove += delegate(object s, MouseEventArgs e) {
            if (!drag) return;
            ox += e.X - last.X; oy += e.Y - last.Y;
            last = e.Location;
            fit = false;
            view.Invalidate();
        };
        view.MouseUp += delegate { drag = false; view.Cursor = Cursors.Default; };
        view.DoubleClick += delegate {
            if (img == null) return;
            if (fit) { Point p = view.PointToClient(Cursor.Position); Zoom(2.5, p); }
            else Fit();
        };
        Controls.Add(view);
        Controls.Add(hint);
        Controls.Add(bar);
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) Close();
            else if (e.KeyCode == Keys.Left || e.KeyCode == Keys.PageUp) Go(index - 1);
            else if (e.KeyCode == Keys.Right || e.KeyCode == Keys.PageDown) Go(index + 1);
            else if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus) Zoom(1.25, Center());
            else if (e.KeyCode == Keys.Subtract || e.KeyCode == Keys.OemMinus) Zoom(0.8, Center());
            else if (e.KeyCode == Keys.D0 || e.KeyCode == Keys.NumPad0) Fit();
            else if (e.KeyCode == Keys.R) Turn();
            else return;
            e.Handled = true;
        };
        FormClosed += delegate { if (img != null) { img.Dispose(); img = null; } };
        Shown += delegate { view.Focus(); };
    }

    Button Bar(string text, int width) {
        Button b = Ui.Button(text, false);
        b.Width = width;
        b.TabStop = false;
        bar.Controls.Add(b);
        return b;
    }

    public string Current { get { return index < files.Count ? files[index] : null; } }

    public void ShowFiles(List<string> list, string current) {
        files.Clear();
        files.AddRange(list);
        int i = current == null ? 0 : files.IndexOf(current);
        if (i < 0) i = Math.Min(index, files.Count - 1);
        index = -1;
        Go(Math.Max(0, i));
    }

    void Go(int i) {
        if (files.Count == 0) { if (img != null) { img.Dispose(); img = null; } name.Text = "фото нет"; view.Invalidate(); return; }
        if (i < 0 || i >= files.Count || i == index) return;
        index = i;
        Cursor old = Cursor.Current;
        Cursor.Current = Cursors.WaitCursor;
        try {
            if (img != null) { img.Dispose(); img = null; }
            img = Photo.Open(files[i], 3200);
        } catch (Exception ex) {
            img = null;
            name.Text = Path.GetFileName(files[i]) + " - не открывается: " + ex.Message;
        } finally { Cursor.Current = old; }
        if (img != null) name.Text = "Фото " + (i + 1) + " из " + files.Count + "\n" + Path.GetFileName(files[i]);
        prev.Enabled = i > 0;
        next.Enabled = i < files.Count - 1;
        Fit();
    }

    Point Center() { return new Point(view.ClientSize.Width / 2, view.ClientSize.Height / 2); }

    void Fit() {
        fit = true;
        if (img != null) {
            scale = Math.Min((double)view.ClientSize.Width / img.Width, (double)view.ClientSize.Height / img.Height) * 0.98;
            if (scale <= 0) scale = 1;
            ox = (view.ClientSize.Width - img.Width * scale) / 2;
            oy = (view.ClientSize.Height - img.Height * scale) / 2;
        }
        view.Invalidate();
    }

    void Zoom(double k, Point at) {
        if (img == null) return;
        double fitScale = Math.Min((double)view.ClientSize.Width / img.Width, (double)view.ClientSize.Height / img.Height) * 0.98;
        double s = Math.Max(fitScale * 0.5, Math.Min(6.0, scale * k));
        ox = at.X - (at.X - ox) * s / scale;
        oy = at.Y - (at.Y - oy) * s / scale;
        scale = s;
        fit = false;
        view.Invalidate();
    }

    void Turn() {
        if (img == null) return;
        Bitmap t = Mrz.Draw(img, img.Width, img.Height, 6);
        img.Dispose();
        img = t;
        Fit();
    }

    void Draw(Graphics g) {
        if (img == null) return;
        // рисуем только видимую часть снимка
        RectangleF dest = new RectangleF((float)ox, (float)oy, (float)(img.Width * scale), (float)(img.Height * scale));
        RectangleF vis = RectangleF.Intersect(dest, new RectangleF(0, 0, view.ClientSize.Width, view.ClientSize.Height));
        if (vis.Width <= 0 || vis.Height <= 0) return;
        RectangleF src = new RectangleF((float)((vis.X - ox) / scale), (float)((vis.Y - oy) / scale), (float)(vis.Width / scale), (float)(vis.Height / scale));
        g.InterpolationMode = scale < 1 ? InterpolationMode.HighQualityBilinear : InterpolationMode.Bilinear;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        g.DrawImage(img, vis, src, GraphicsUnit.Pixel);
    }
}

// ---------- фото с телефона по кабелю ----------
// Телефон (Android, режим «Передача файлов») виден в «Компьютер» как устройство, а не как диск: обычными
// файловыми функциями его не прочитать. Берём его через оболочку Windows (Shell.Application, как Проводник):
// ищем на телефоне папку DCIM, в ней новые снимки и копируем их в папку договора.

public delegate object ShellMaker();
public delegate void TextHandler(string text);

public class PhoneResult {
    public string Error;                                   // ничего не вышло (нет телефона, закрыт и т.п.)
    public string Device = "";
    public readonly List<string> Copied = new List<string>();
    public readonly List<string> Failed = new List<string>();
    public bool NoneNew;                                   // новых фото нет
    public DateTime Newest = DateTime.MinValue;            // время самого нового забранного фото (по телефону)
    public int Total;                                      // фото на телефоне всего
}

public static class Phone {
    // оболочку можно подменить (проверка программы без телефона)
    public static ShellMaker MakeShell = delegate { return Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application")); };
    public static int CopySeconds = 120;   // сколько ждать одно фото
    public const int MaxNew = 15;          // за раз - не больше стольких новых

    class Pic { public object Item; public string Name; public DateTime Date; public long Size; }

    static object Call(object o, string m, params object[] a) { return o.GetType().InvokeMember(m, System.Reflection.BindingFlags.InvokeMethod, null, o, a); }
    static object Prop(object o, string p) { return o.GetType().InvokeMember(p, System.Reflection.BindingFlags.GetProperty, null, o, null); }

    static List<object> Items(object folder) {
        List<object> r = new List<object>();
        if (folder == null) return r;
        object items = Call(folder, "Items");
        if (items == null) return r;
        int n = Convert.ToInt32(Prop(items, "Count"));
        for (int i = 0; i < n; i++) {
            object it = Call(items, "Item", i);
            if (it != null) r.Add(it);
        }
        return r;
    }

    static string Str(object o, string p) { try { object v = Prop(o, p); return v == null ? "" : v.ToString(); } catch { return ""; } }
    static bool Flag(object o, string p) { try { return Convert.ToBoolean(Prop(o, p)); } catch { return false; } }
    static object Sub(object item) { try { return Prop(item, "GetFolder"); } catch { return null; } }

    static string Ext(object item, string key) {
        try { object v = Call(item, "ExtendedProperty", key); return v == null ? "" : v.ToString(); } catch { return ""; }
    }

    // имя файла с расширением (Проводник может прятать расширения)
    static string FileName(object item) {
        string n = Ext(item, "System.FileName");
        if (n.Length == 0) n = Str(item, "Name");
        if (Path.GetExtension(n).Length == 0) {
            string t = Str(item, "Type").ToUpperInvariant();
            if (t.Contains("JPG") || t.Contains("JPEG")) n += ".jpg";
        }
        return n;
    }

    static DateTime Date(object item, string name) {
        try {
            object d = Prop(item, "ModifyDate");
            if (d is DateTime && ((DateTime)d).Year > 2000) return (DateTime)d;
        } catch { }
        return NameDate(name);
    }

    // время снимка по имени файла: IMG_20261016_143012.jpg, 20261016_143012.jpg
    static DateTime NameDate(string name) {
        Match m = Regex.Match(name, @"(20\d{6})[_-]?(\d{6})");
        DateTime t;
        if (m.Success && DateTime.TryParseExact(m.Groups[1].Value + m.Groups[2].Value, "yyyyMMddHHmmss", null, System.Globalization.DateTimeStyles.None, out t)) return t;
        return DateTime.MinValue;
    }

    static object Named(List<object> items, string name) {
        foreach (object it in items) if (string.Compare(Str(it, "Name"), name, StringComparison.OrdinalIgnoreCase) == 0 && Flag(it, "IsFolder")) return it;
        return null;
    }

    // забрать с телефона: новые фото (снятые позже since) или, если last > 0, последние last фото
    public static PhoneResult Take(string dir, DateTime since, int last, TextHandler progress) {
        PhoneResult r = new PhoneResult();
        try {
            object shell = MakeShell();
            object computer = Call(shell, "NameSpace", 17);       // «Компьютер»
            List<object> devices = new List<object>();
            foreach (object it in Items(computer))
                if (!Flag(it, "IsFileSystem") && Flag(it, "IsFolder")) devices.Add(it);
            if (devices.Count == 0) {
                r.Error = "Телефон не найден. Подключите его кабелем, разблокируйте и на телефоне выберите «Передача файлов».";
                return r;
            }
            List<Pic> pics = new List<Pic>();
            int heic = 0;
            bool dcim = false;
            List<string> names = new List<string>();
            foreach (object dev in devices) {
                string devName = Str(dev, "Name");
                names.Add("«" + devName + "»");
                if (progress != null) progress("Телефон «" + devName + "»: ищу фото...");
                // DCIM - на самом устройстве или в его памяти («Phone», «Внутренний общий накопитель», карта)
                List<object> found = new List<object>();
                List<object> top = Items(Sub(dev));
                object d0 = Named(top, "DCIM");
                if (d0 != null) found.Add(d0);
                foreach (object st in top) {
                    if (!Flag(st, "IsFolder") || st == d0) continue;
                    object d1 = Named(Items(Sub(st)), "DCIM");
                    if (d1 != null) found.Add(d1);
                }
                if (found.Count == 0) continue;
                dcim = true;
                r.Device = devName;
                foreach (object d in found) {
                    List<object> folders = new List<object>();
                    folders.Add(d);
                    foreach (object sub in Items(Sub(d)))
                        if (Flag(sub, "IsFolder") && !Str(sub, "Name").StartsWith(".")) folders.Add(sub);
                    foreach (object f in folders) {
                        foreach (object it in Items(Sub(f))) {
                            // фото за прошлые дни (видно по имени) пропускаем сразу: на телефоне их могут быть тысячи
                            DateTime byName = NameDate(Str(it, "Name"));
                            if (last == 0 && byName != DateTime.MinValue && byName.Date < since.Date) continue;
                            if (Flag(it, "IsFolder")) continue;
                            string n = FileName(it);
                            string e = Path.GetExtension(n).ToLowerInvariant();
                            if (e == ".heic" || e == ".heif") { heic++; continue; }
                            if (e != ".jpg" && e != ".jpeg") continue;
                            Pic p = new Pic();
                            p.Item = it;
                            p.Name = n;
                            p.Date = Date(it, n);
                            try { p.Size = Convert.ToInt64(Prop(it, "Size")); } catch { p.Size = 0; }
                            pics.Add(p);
                        }
                    }
                }
            }
            if (!dcim) {
                r.Error = "Телефон " + string.Join(", ", names.ToArray()) + " виден, но фото на нём не открываются: разблокируйте телефон " +
                    "и на вопросе «Разрешить доступ к данным?» нажмите «Разрешить» (или выберите «Передача файлов»).";
                return r;
            }
            r.Total = pics.Count;
            if (pics.Count == 0) {
                r.Error = heic > 0 ? "На телефоне фото в формате HEIC - Windows 7 их не открывает. В настройках камеры телефона выключите «Фото в формате HEIF» и переснимите."
                    : "На телефоне «" + r.Device + "» в папке DCIM нет фото.";
                return r;
            }
            // по времени снимка (без времени - по имени), новые - в конце
            pics.Sort(delegate(Pic a, Pic b) {
                int c = a.Date.CompareTo(b.Date);
                return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
            });
            List<Pic> take = new List<Pic>();
            if (last > 0) for (int i = Math.Max(0, pics.Count - last); i < pics.Count; i++) take.Add(pics[i]);
            else {
                foreach (Pic p in pics) if (p.Date > since) take.Add(p);
                if (take.Count > MaxNew) take.RemoveRange(0, take.Count - MaxNew);
            }
            if (take.Count == 0) { r.NoneNew = true; return r; }

            Directory.CreateDirectory(dir);
            object target = Call(shell, "NameSpace", dir);
            if (target == null) { r.Error = "Не открывается папка для фото: " + dir; return r; }
            for (int i = 0; i < take.Count; i++) {
                Pic p = take[i];
                if (p.Date > r.Newest) r.Newest = p.Date;
                // уже забрано раньше в эту папку
                if (File.Exists(Path.Combine(dir, p.Name))) continue;
                if (progress != null) progress("Копирую с телефона фото " + (i + 1) + " из " + take.Count + "...");
                Dictionary<string, bool> before = Snapshot(dir);
                Call(target, "CopyHere", p.Item, 4 | 16 | 512 | 1024);   // без вопросов и окон
                string got = WaitNew(dir, before, p.Size);
                if (got != null) r.Copied.Add(got);
                else r.Failed.Add(p.Name);
            }
        } catch (Exception ex) {
            Exception e = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
            r.Error = "Не получилось забрать фото с телефона: " + e.Message;
        }
        return r;
    }

    static Dictionary<string, bool> Snapshot(string dir) {
        Dictionary<string, bool> d = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (string f in Directory.GetFiles(dir)) d[f] = true;
        return d;
    }

    // копирование с телефона идёт в фоне: ждём новый файл, пока он не допишется (размер не растёт и файл свободен)
    static string WaitNew(string dir, Dictionary<string, bool> before, long size) {
        DateTime end = DateTime.Now.AddSeconds(CopySeconds);
        string cand = null;
        long len = -1;
        int same = 0;
        while (DateTime.Now < end) {
            System.Threading.Thread.Sleep(300);
            if (cand == null)
                foreach (string f in Directory.GetFiles(dir))
                    if (!before.ContainsKey(f) && Photo.IsPhoto(f)) { cand = f; break; }
            if (cand == null) continue;
            long l;
            try { l = new FileInfo(cand).Length; } catch { continue; }
            if (l == len && l > 0 && (size <= 0 || l >= size)) same++;
            else same = 0;
            len = l;
            if (same < 2) continue;
            try { using (FileStream fs = new FileStream(cand, FileMode.Open, FileAccess.Read, FileShare.None)) { } }
            catch (IOException) { same = 0; continue; }
            return cand;
        }
        return null;
    }
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
    // ищем только по индексам базы (CDX), перебора таблицы нет; нет индекса - проверить нельзя (а не «нет в базе»)
    public bool AccountIndex { get { return Base != null && Base.TagFor("NAB") != null; } }
    public bool AddressIndex { get { return Base != null && Base.TagFor("UL", true) != null; } }
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
            if (r.NoIndex) throw new Exception("в базе нет индекса по лицевому счёту");
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
            if (r.NoIndex) throw new Exception("в базе нет индекса по улице и дому");
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
    // фото документов: у каждого договора своя папка (Договоры\фото\<договор>)
    const string PhotoCard = "Фото документов";
    readonly string photoRoot;
    string photoDir;                                 // папка фото этого договора; null - фото ещё нет
    readonly PhotoStrip strip = new PhotoStrip();
    readonly Label photoNote = Ui.Note(""), lichnyNote = Ui.Note(""), fioNote = Ui.Note(""), certNote = Ui.Note("");
    CertResult lastCert;                            // прочитанное со свидетельства о госрегистрации
    readonly Label[] meterNotes = { Ui.Note(""), Ui.Note("") };
    readonly MeterResult[] photoMeter = new MeterResult[2];    // прочитанное с фото счётчиков 1 и 2
    readonly string[] meterWhere = { "", "" };
    int lastMeterIdx;
    string certWhere = "";
    bool certSure;
    Button phoneButton, photoAdd, photoFolder;
    PhotoViewer viewer;
    bool phoneBusy;
    int photoGen;                                    // меняется с договором: ответы по фото прошлого договора не нужны
    // прочитанное с фото паспорта этого договора
    MrzResult lastMrz;
    PageResult lastPage;
    string mrzWhere = "", pageWhere = "", importNote = "";
    string photoPersonal = "", photoFio = "";
    bool photoFioOk, mrzWeak;
    static readonly object mrzLock = new object();

    public DogovorForm(string here) {
        this.here = here;
        templateFile = Path.Combine(here, "shablon.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = new Memory(Path.Combine(here, "nastroiki.txt"));
        outDir = memory.Get("ПапкаДоговоров", "");
        if (outDir.Length == 0) outDir = Path.Combine(here, "Договоры");
        dataDir = Path.Combine(outDir, "данные");
        photoRoot = Path.Combine(outDir, "фото");
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
        // поле, в котором курсор, - светло-жёлтое: видно, где печатаешь
        if (c is TextBox || c is ComboBox) {
            c.Enter += delegate { c.BackColor = St.Focus; };
            c.Leave += delegate { c.BackColor = c.Enabled ? Color.White : St.Disabled; checkTimer.Stop(); checkTimer.Start(); };
        }
        // с заглавной, когда уходим из поля: «боровляны» -> «Боровляны», «иванов пётр» -> «Иванов Пётр»
        if (key == "np" || key == "ulica" || key == "fio" || key == "adres_prozh") c.Leave += delegate {
            string s = V(key), t = key == "fio" ? UpperWords(s) : UpperFirst(s);
            if (t != s) SetV(key, t);
        };
    }

    static string UpperFirst(string s) { return s.Length == 0 ? s : char.ToUpper(s[0]) + s.Substring(1); }

    static string UpperWords(string s) {
        char[] a = s.ToCharArray();
        for (int i = 0; i < a.Length; i++)
            if (char.IsLetter(a[i]) && (i == 0 || a[i - 1] == ' ' || a[i - 1] == '-')) a[i] = char.ToUpper(a[i]);
        return new string(a);
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
        foreach (Cell c in cells) if (c.Ctl != null && c.Ctl.Name.Length > 0) { cardOf[c.Ctl.Name] = card; Field(c.Ctl.Name, c.Caption, card.Name); }
    }

    // ---------- обязательные поля (выбирает пользователь) ----------

    // поля, которые можно сделать обязательными: ключ (у счётчиков - «pu_tip», для обоих), подпись, раздел - в порядке формы
    readonly List<string[]> fieldList = new List<string[]>();
    Dictionary<string, bool> requiredSet;

    // подписи понятнее, чем над полем в форме («ток, А», «№»)
    static readonly string[] FieldNames = {
        "np_tip=Тип населённого пункта", "ulica_tip=Тип улицы", "ulica=Улица (название)", "korp=Корпус", "kv=Квартира",
        "soob_tel=Сообщать о неисправностях: телефон", "telefon=Телефон потребителя", "dopusk_kem=Допуск: кем выдан",
        "plomba_klemm=Пломба на клеммной крышке", "plomba_drugoe=Другие пломбы", "utratil_nom=№ приложения, утратившего силу",
        "pu_vvod_i=Вводной аппарат: ток, А", "pu_zasch_i=Аппарат защиты: ток, А", "pu_model=Модель счётчика",
        "pu_korpus=Тип корпуса счётчика", "pu_modif=Модификация счётчика", "adres=Адрес объекта (строка в договоре)",
        "pu_tip=Тип счётчика (строка в договоре)", "moshn_otopl=Мощность на отопление и ГВС, кВт",
        "pu_nagrev=Электронагревательные приборы", "pu_modul=Модуль отображения (сплит-счётчик)" };

    // обязательные сначала - как было до настройки
    static readonly string[] DefaultRequired = {
        "ls", "oformil", "podp_dolzh", "podp_fio", "podp_dov", "podp_inic", "soob_adres", "soob_tel", "fio", "potr_inic", "lichny",
        "adres_prozh", "obekt", "moshn", "np", "adres", "prochie", "plita", "gaz", "teplo", "gvs", "sms", "sms_tel", "utratil_nom",
        "pu_tip", "pu_zav", "pu_cel", "pu_faz", "pu_askue", "pu_tarif", "pu_gruppa" };

    void Field(string key, string caption, string card) {
        Control c;
        if (!inp.TryGetValue(key, out c) || c is CheckBox || c is DateBox) return;
        if (key.StartsWith("pu2_")) return;
        if (key.StartsWith("pu1_")) key = "pu_" + key.Substring(4);
        foreach (string[] f in fieldList) if (f[0] == key) return;
        string name = (caption ?? "").Trim();
        foreach (string n in FieldNames) if (n.StartsWith(key + "=")) name = n.Substring(key.Length + 1);
        if (name.Length == 0) name = key;
        fieldList.Add(new string[] { key, name, card });
    }

    // для проверки без окон
    public void ShowRequired() {
        if (requiredSet == null) LoadRequired();
        RequiredDialog d = new RequiredDialog(fieldList, requiredSet, DefaultRequired);
        d.StartPosition = FormStartPosition.Manual;
        d.Location = new Point(120, 40);
        d.Show(this);
    }
    public void SetRequired(string csv) { memory.Values["ОбязательныеПоля"] = csv; memory.Save(); LoadRequired(); RefreshState(); }
    public string StatusText { get { return status.Text; } }

    void LoadRequired() {
        requiredSet = new Dictionary<string, bool>();
        string v = memory.Get("ОбязательныеПоля", null);
        string[] keys = v == null ? DefaultRequired : v.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string k in keys) requiredSet[k.Trim()] = true;
    }

    void SetupRequired() {
        RequiredDialog d = new RequiredDialog(fieldList, requiredSet, DefaultRequired);
        if (d.ShowDialog(this) != DialogResult.OK) return;
        requiredSet = d.Result;
        List<string> keys = new List<string>();
        foreach (string[] f in fieldList) if (requiredSet.ContainsKey(f[0])) keys.Add(f[0]);
        memory.Values["ОбязательныеПоля"] = string.Join(",", keys.ToArray());
        memory.Save();
        RefreshState();
    }

    static Cell F(string caption, Control c, int span) {
        Cell x = Cell.F(caption, c, span);
        x.W = WidthOf(c.Name);
        return x;
    }

    // ширина поля в знаках - по длине значения (у списков - со стрелкой); нет в списке - на всю оставшуюся ширину
    static readonly string[] Widths = {
        "data=14", "ls=16", "oformil=22", "podp=20", "podp_inic=16", "podp_dov=26", "podp_fio=30", "soob_adres=30", "soob_tel=16",
        "fio=42", "potr_inic=18", "lichny=17", "telefon=18",
        "obekt=34", "moshn=8", "moshn_otopl=8", "np_tip=7", "np=28", "selsovet=18", "kadastr=22", "ulica_tip=7", "ulica=28",
        "dom=8", "korp=6", "kv=6", "dopusk_nom=12", "dopusk_data=14",
        "plomba_vvod=16", "plomba_klemm=16", "plomba_drugoe=24", "plita=12", "gaz=12", "teplo=12", "gvs=12",
        "sms=22", "sms_tel=18", "utratil_nom=10", "utratil_ot=14",
        "pu_model=12", "pu_korpus=9", "pu_modif=24", "pu_zav=16", "pu_poverka=11", "pu_pokaz=14", "pu_faz=12", "pu_askue=12",
        "pu_plomba_gp=14", "pu_prinadl=22", "pu_modem=14", "pu_vvod_tip=16", "pu_vvod_i=8", "pu_zasch_tip=16", "pu_zasch_i=8",
        "pu_modul=24" };

    static int WidthOf(string key) {
        if (key == null || key.Length == 0) return 0;
        if (key.StartsWith("pu1_") || key.StartsWith("pu2_")) key = "pu_" + key.Substring(4);
        foreach (string s in Widths) if (s.StartsWith(key + "=")) return int.Parse(s.Substring(key.Length + 1));
        return 0;
    }

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
        Row(c, F(null, lsNote, 4));
        Row(c, F("Исполнитель в лице (подписывает)", C("podp", true, false, "выберите из списка"), 1), F("Подпись подписанта", T("podp_inic", 1, "И.О.Фамилия"), 1),
            F("Доверенность", T("podp_dov", 1, "№ ... от ..."), 2));
        Row(c, F("Должность подписанта (в лице кого)", T("podp_dolzh", 3, null), 3), F("Ф.И.О. подписанта (кого)", T("podp_fio", 3, null), 1));
        Row(c, F("Сообщать о неисправностях (п. 3.6): адрес", T("soob_adres", 1, null), 2), F("телефон", T("soob_tel", 1, null), 2));

        BuildPhotos();

        c = NewCard("Потребитель", "гражданин(ка), с которым заключается договор", 4);
        Row(c, F("Фамилия, собственное имя, отчество", T("fio", 1, "Иванов Иван Иванович"), 2), F("Подпись (инициалы)", T("potr_inic", 1, "заполнится сама"), 1),
            F("Идентификационный номер", T("lichny", 1, "14 знаков"), 1));
        Row(c, F(null, fioNote, 4));
        Row(c, F(null, lichnyNote, 4));
        Button same = Ui.Button("= адрес объекта", false);
        same.Font = St.Small;
        same.Click += delegate { SetV("adres_prozh", V("adres")); };
        same.TabStop = false;
        tip.SetToolTip(same, "Потребитель проживает по адресу объекта");
        Cell sameCell = F(" ", same, 1);
        sameCell.W = 16;
        Row(c, F("Адрес проживания (регистрации)", T("adres_prozh", 1, "г.Минск, ул. ..., д. ..., кв. ..."), 2), sameCell, F("Телефон", T("telefon", 1, "необязательно"), 1));

        c = NewCard("Объект", "что подключаем и где", 8);
        Row(c, F("Объект электроснабжения", C("obekt", true, false, null), 4), F("Мощность, кВт", T("moshn", 1, "например 10"), 2),
            F("из неё отопление и ГВС, кВт", T("moshn_otopl", 1, "0,0"), 2));
        Row(c, F("Тип н.п.", C("np_tip", true, false, null), 1), F("Населённый пункт / садоводческое товарищество", C("np", true, false, "Боровляны"), 3),
            F("Сельсовет", C("selsovet", true, false, null), 2), F("Кадастровый номер", T("kadastr", 1, "цифры или № по свидетельству"), 2));
        Row(c, F("Улица", C("ulica_tip", true, false, null), 1), F(" ", C("ulica", true, false, "Дружная"), 3), F("Дом / участок", T("dom", 1, null), 2),
            F("Корпус", T("korp", 1, null), 1), F("Квартира", T("kv", 1, null), 1));
        Row(c, F("Адрес объекта - так будет в договоре (собирается сам из полей выше, можно поправить)", T("adres", 1, null), 8));
        Row(c, F(null, adresNote, 8));
        Row(c, F(null, certNote, 8));
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
        Cell utr = F(" ", X("utratil", "Считать утратившим силу приложение к договору"), 2);
        utr.W = 46;
        Row(c, utr, F("№", T("utratil_nom", 1, null), 1), F("от", D("utratil_ot"), 1));
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
        box.Mark = St.TintOf("Счётчики").Mark;
        Grid g = new Grid(box, 2);
        g.PadX = 12;
        g.PadTop = 30;
        g.PadBottom = 12;
        g.RowGap = 6;
        // тип счётчика - по частям из справочника; в конце каждого списка «свое»
        Grid parts = new Grid(new Panel(), 4);
        parts.Host.BackColor = Color.Transparent;
        parts.PadX = 0;
        parts.PadBottom = 0;
        parts.Row(F("Модель", C(p + "model", true, false, null), 1), F("Тип корпуса", C(p + "korpus", true, false, null), 1),
            F("Модификация", C(p + "modif", true, false, null), 2));
        g.Row(Cell.G(parts, 2));
        foreach (Cell x in parts.Cells()) cardOfLater.Add(x.Ctl);
        g.Row(F("Тип счётчика в договоре (собирается сам, можно поправить)", T(p + "tip", 1, null), 2));
        Label mn = meterNotes[i - 1];
        int mi = i;
        mn.Click += delegate { if (photoMeter[mi - 1] != null) { PutMeter(mi, true); MeterNote(mi); } };
        g.Row(F(null, mn, 2));
        g.Row(F("Заводской номер", T(p + "zav", 1, null), 1), F("Поверка (год-квартал)", T(p + "poverka", 1, "2025-2"), 1),
            F("Показания, кВт·ч", T(p + "pokaz", 1, null), 1));
        g.Row(F("Цель использования электроэнергии", C(p + "cel", true, false, null), 2));
        g.Row(F("Фазность", S(p + "faz", "1 ф", "3 ф"), 1), F("АСКУЭ", S(p + "askue", "да", "нет"), 1), F("Пломба госповерителя", C(p + "plomba_gp", true, false, null), 1));
        g.Row(F("Место установки", C(p + "mesto", true, false, null), 2));
        g.Row(F("Принадлежность", C(p + "prinadl", true, false, null), 1), F("Тип модема", C(p + "modem", true, false, null), 1));
        g.Row(F("Вводной аппарат", C(p + "vvod_tip", true, false, null), 1), F("ток, А", C(p + "vvod_i", true, false, null), 1),
            F("Аппарат защиты", C(p + "zasch_tip", true, false, null), 1), F("ток, А", C(p + "zasch_i", true, false, null), 1));
        g.Row(F("Модуль отображения (сплит-счётчик), зав. №", T(p + "modul", 1, "нет модуля - прочерк"), 2));
        g.Row(F("Тариф", S(p + "tarif", Tarifs), 2));
        g.Row(F("Тарифная группа", C(p + "gruppa", true, true, null), 2));
        Heaters heat = new Heaters();
        Reg(p + "nagrev", heat);
        g.Row(F("Электронагревательные приборы: вид, шт., кВт одного (+ ещё прибор)", heat, 2));
        if (i == 1) {
            foreach (Cell x in parts.Cells()) if (x.Ctl != null && x.Ctl.Name.Length > 0) Field(x.Ctl.Name, x.Caption, "Счётчики");
            foreach (Cell x in g.Cells()) if (x.Ctl != null && x.Ctl.Name.Length > 0) Field(x.Ctl.Name, x.Caption, "Счётчики");
        }
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
        nav.Width = 252;
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
            n.Mark = St.TintOf(c.Name).Mark;
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
        NavItem reqs = new NavItem("Обязательные поля", true);
        reqs.Click += delegate { SetupRequired(); };
        foreach (NavItem n in new NavItem[] { reqs, folder, dirs }) {
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
            reqs.SetBounds(0, folder.Top - reqs.Height, nav.Width, reqs.Height);
            baseLabel.SetBounds(22, reqs.Top - 58, nav.Width - 34, 52);
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
            // узкое окно (рядом открыто фото) - короткие надписи
            bool narrow = p.ClientSize.Width < 900;
            nw.Text = narrow ? "Новый договор" : "Создать новый договор";
            nw.Width = narrow ? 140 : 200;
            open.Text = narrow ? "Открыть…" : "Открыть сохранённый…";
            open.SetBounds(nw.Right + 10, 14, narrow ? 110 : 200, 38);
            create.Width = narrow ? 190 : 210;
            create.Location = new Point(p.ClientSize.Width - 24 - create.Width, 13);
            status.Font = narrow ? St.Small : St.Base;
            status.SetBounds(open.Right + (narrow ? 14 : 24), 14, Math.Max(10, create.Left - open.Right - (narrow ? 28 : 48)), 38);
        };
        return p;
    }

    void LayoutCards() {
        content.SuspendLayout();
        int full = content.Width - SystemInformation.VerticalScrollBarWidth;
        // не шире 1120: поля ближе друг к другу, глазам не бегать по всему экрану
        int w = Math.Max(600, Math.Min(full - 32, 1120));
        int x = Math.Max(16, (full - w) / 2);
        int y = 4 + content.AutoScrollPosition.Y;
        // узкая форма (рядом открыто фото) - кнопки фото короче
        bool narrow = w < 900;
        photoFolder.Text = narrow ? "Папка" : "Папка фото";
        photoFolder.Width = narrow ? 76 : 120;
        photoAdd.Text = narrow ? "Из папки…" : "Добавить из папки…";
        photoAdd.Width = narrow ? 104 : 170;
        phoneButton.Text = narrow ? "Забрать с телефона" : "Забрать фото с телефона";
        phoneButton.Width = narrow ? 180 : 220;
        foreach (Card c in cards) {
            int h = c.Layout(w);
            c.Panel.Location = new Point(x + content.AutoScrollPosition.X, y);
            y += h + 10;
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
                FioNote();
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
                LichnyNote();
            }
        } finally { loading = false; }
        if (key == "ls") Recheck(true, false);
        if (lastCert != null && (Array.IndexOf(CertKeys, key) >= 0 || key == "fio" || key == "kadastr")) CertNote();
        for (int i = 1; i <= 2; i++)
            if (photoMeter[i - 1] != null && (key == "pu" + i + "_zav" || key == "pu" + i + "_tip" || key == "pu" + i + "_pokaz" || key == "pu" + i + "_poverka")) MeterNote(i);
        if (key == "adres" || key == "np_tip" || key == "np" || key == "ulica_tip" || key == "ulica" || key == "dom" || key == "korp" || key == "kv") Recheck(false, true);
        if (key.EndsWith("tarif") || key.EndsWith("vvod_tip") || key == "schetchikov" || key == "sms" || key == "utratil" || key == "plita") Enable();
        RefreshState();
    }

    void Enable() {
        bool two = V("schetchikov") != "1 счётчик";
        meter2.Enabled = two;
        meter2.BackColor = two ? Color.White : St.Disabled;
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
                    // проверка - поиском по индексу базы (несколько чтений), не перебором всех записей
                    if (b.AccountIndex && b.AddressIndex) {
                        baseLabel.Text = "База: " + b.Table + "\nлицевой и адрес ищутся по индексу базы";
                        baseLabel.ForeColor = Color.FromArgb(150, 210, 170);
                    } else {
                        string what = !b.AccountIndex && !b.AddressIndex ? "индексов - лицевые и адреса" : !b.AccountIndex ? "индекса по лицевому - лицевые" : "индекса по адресу - адреса";
                        baseLabel.Text = "База: " + b.Table + "\nнет " + what + " не проверяются";
                        baseLabel.ForeColor = Color.FromArgb(240, 150, 150);
                    }
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
        if (requiredSet == null) LoadRequired();
        int n = V("schetchikov") == "1 счётчик" ? 1 : 2;
        // обязательные - по настройке; недоступное поле (СМС при отказе, ток без вводного аппарата) не в счёт
        foreach (string[] f in fieldList) {
            if (!requiredSet.ContainsKey(f[0])) continue;
            List<string[]> ks = new List<string[]>();
            if (f[0].StartsWith("pu_")) { for (int i = 1; i <= n; i++) ks.Add(new string[] { "pu" + i + "_" + f[0].Substring(3), " (счётчик " + i + ")" }); }
            else ks.Add(new string[] { f[0], "" });
            foreach (string[] k in ks) {
                Control c;
                if (!inp.TryGetValue(k[0], out c) || !c.Enabled) continue;
                if (V(k[0]).Length == 0) r.Add(new Problem(k[0], f[1] + k[1], true));
            }
        }
        string block = LsBlockText();
        if (block != null) { Problem b = new Problem("ls", block, true); b.Blocker = true; r.Add(b); }
        if (requiredSet.ContainsKey("sms_tel") && V("sms") == "согласен(а)" && V("sms_tel").Length > 0 && Digits(V("sms_tel")).Length < 9)
            r.Add(new Problem("sms_tel", "номер для СМС не полный (в договоре будет линия)", true));
        for (int i = 1; i <= n; i++)
            if (V("pu" + i + "_poverka").EndsWith("-"))
                r.Add(new Problem("pu" + i + "_poverka", "дата поверки: допишите квартал (счётчик " + i + ")", requiredSet.ContainsKey("pu_poverka")));

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
            int cnt = cards[i].Name == "Порядок и документы" || cards[i].Name == PhotoCard ? -1 : per[cards[i]];
            if (navs[i].Count != cnt) { navs[i].Count = cnt; navs[i].Invalidate(); }
        }
        foreach (KeyValuePair<string, Control> kv in inp) {
            bool m = showMissing && miss.ContainsKey(kv.Key);
            if (kv.Value is Seg) { if (((Seg)kv.Value).Missing != m) ((Seg)kv.Value).Missing = m; }
            else if (kv.Value is TextBox || kv.Value is ComboBox) {
                Color want = kv.Value.ContainsFocus ? St.Focus : m ? St.Missing : (kv.Value.Enabled ? Color.White : St.Disabled);
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
        string tip = V("np_tip"), np = UpperFirst(V("np"));
        List<string> parts = new List<string>();
        bool st = tip.ToUpper() == "СТ" || tip.ToLower() == "с/т";
        if (np.Length > 0) parts.Add(st ? tip + " «" + np.Trim('«', '»', '"') + "»" : (tip + " " + np).Trim());
        if (V("ulica").Length > 0) parts.Add((V("ulica_tip") + " " + UpperFirst(V("ulica"))).Trim());
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

    // размер шрифта значений в узких ячейках: цель использования в шапке таблицы счётчиков.
    // В бланке там «Бытовое использование» - две строки 10 пт; длинная цель («Бытовое потребление, отопление и
    // горячее водоснабжение») в 10 пт - пять строк, таблица растёт и её конец уезжает на следующую страницу
    public static Dictionary<string, int> Sizes(Dictionary<string, string> v) {
        Dictionary<string, int> r = new Dictionary<string, int>();
        foreach (string k in new string[] { "ПУ1_ЦЕЛЬ", "ПУ2_ЦЕЛЬ" }) {
            string s;
            if (!v.TryGetValue(k, out s) || s.Length == 0) continue;
            // ширина текста в ячейке 69 пт (берём 68); высота - не больше 40 пт (было 24, запас на странице - около 24)
            foreach (int hp in new int[] { 20, 18, 16, 14 }) {
                int lines = Fit.Lines(s, hp / 2f, 68f);
                if (lines * (hp == 20 ? 12f : hp * 0.6f) <= 40f || hp == 14) { if (hp < 20) r[k] = hp; break; }
            }
        }
        return r;
    }

    static string LowerFirst(string s) {
        if (s.Length < 2 || !char.IsUpper(s[0]) || char.IsUpper(s[1])) return s;
        return char.ToLower(s[0]) + s.Substring(1);
    }

    public Dictionary<string, string> Compose() {
        Dictionary<string, string> v = new Dictionary<string, string>();
        v["ДАТА"] = V("data");
        v["ЛС"] = V("ls");
        string tip = V("np_tip");
        string np = UpperFirst(V("np"));
        if (tip.ToUpper() == "СТ" || tip.ToLower() == "с/т") np = tip + " «" + np.Trim('«', '»', '"') + "»";
        // за названием - линия до места, где она кончается в бланке (табуляция с подчёркиванием, 122,7 пт).
        // Название длиннее линии - без табуляции (иначе дата и номер съедут), а почти такое же - добиваем
        // подчёркнутыми пробелами, чтобы точно перешло за конец линии
        if (Fit.Width(np, 12f, true) <= 116f) np += "\t";
        else while (Fit.Width(np, 12f, true) < 130f) np += "\u00A0";
        v["НП"] = np;
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
        v["СМС_ТЕЛ"] = sms && Digits(V("sms_tel")).Length >= 9 ? V("sms_tel") : Line(13, '_');
        return v;
    }

    // ---------- фото документов ----------

    void BuildPhotos() {
        Card c = NewCard(PhotoCard, "с телефона или из папки", 4);
        Button folder = photoFolder = Ui.Button("Папка фото", false);
        folder.Size = new Size(120, 30);
        folder.Click += delegate { OpenPhotoFolder(); };
        tip.SetToolTip(folder, "Открыть папку с фото этого договора");
        Button add = photoAdd = Ui.Button("Добавить из папки…", false);
        add.Size = new Size(170, 30);
        add.Click += delegate { AddFromFolder(); };
        tip.SetToolTip(add, "Скопировать в договор фото, которые уже лежат на компьютере");
        phoneButton = Ui.Button("Забрать фото с телефона", true);
        phoneButton.Size = new Size(220, 30);
        phoneButton.Click += delegate { TakeFromPhone(0); };
        tip.SetToolTip(phoneButton, "Телефон подключён кабелем и разблокирован, на телефоне выбрано «Передача файлов».\nНовые фото с телефона попадут в папку этого договора");
        foreach (Button b in new Button[] { folder, add, phoneButton }) { b.TabStop = false; c.AddExtra(b); }
        strip.Hint = "Фото пока нет. Сфотографируйте телефоном паспорт (последнюю страницу - со строками «<<<», предпоследнюю, прописку) " +
            "свидетельство и счётчик, подключите телефон кабелем и нажмите «Забрать фото с телефона». Личный номер, Ф.И.О., адрес объекта, номер и тип счётчика программа прочитает с фото сама.";
        strip.OpenPhoto += delegate(string f) { ShowPhoto(f); };
        strip.RemovePhoto += delegate(string f) { RemovePhoto(f); };
        strip.Resized += delegate { LayoutCards(); };
        lichnyNote.Click += delegate { if (photoPersonal.Length > 0 && V("lichny") != photoPersonal) SetV("lichny", photoPersonal); };
        fioNote.Click += delegate { if (photoFio.Length > 0 && Same(V("fio")) != Same(photoFio)) SetV("fio", photoFio); };
        certNote.Click += delegate { if (lastCert != null && !CertSame()) { PutCert(); CertNote(); } };
        Row(c, F(null, photoNote, 4));
        Row(c, F(null, strip, 4));
    }

    void PhotoNote(string text, Color color) { Note(photoNote, text, color, false); }

    // в данных договора: имя папки в «Договоры\фото» (или полный путь, если папка где-то ещё)
    string PhotoRef() {
        if (photoDir == null) return "";
        try {
            if (string.Compare(Path.GetDirectoryName(Path.GetFullPath(photoDir)), Path.GetFullPath(photoRoot), StringComparison.OrdinalIgnoreCase) == 0) return Path.GetFileName(photoDir);
        } catch { }
        return photoDir;
    }

    // папка для фото: у договора уже есть - она; нет - по имени договора (пока не сохранён - «Договор дата время»)
    string PhotoTarget() {
        if (photoDir != null) return photoDir;
        string n = FileName(), d = Path.Combine(photoRoot, n);
        for (int i = 2; Directory.Exists(d); i++) d = Path.Combine(photoRoot, n + " (" + i + ")");
        return d;
    }

    // при сохранении договора папка фото получает его имя
    void MovePhotos(string name) {
        if (photoDir == null || !Directory.Exists(photoDir)) return;
        string want = Path.Combine(photoRoot, name);
        try {
            if (string.Compare(Path.GetFullPath(want), Path.GetFullPath(photoDir), StringComparison.OrdinalIgnoreCase) == 0 || Directory.Exists(want)) return;
            Directory.Move(photoDir, want);
            photoDir = want;
        } catch { }
        RefreshPhotos();
    }

    void RefreshPhotos() {
        List<string> list = new List<string>();
        try {
            if (photoDir != null && Directory.Exists(photoDir))
                foreach (string f in Directory.GetFiles(photoDir)) if (Photo.IsPhoto(f)) list.Add(f);
        } catch { }
        list.Sort(StringComparer.OrdinalIgnoreCase);
        strip.SetFiles(list);
        if (viewer != null && !viewer.IsDisposed && viewer.Visible) viewer.ShowFiles(list, viewer.Current);
    }

    void OpenPhotoFolder() {
        string d = photoDir != null && Directory.Exists(photoDir) ? photoDir : photoRoot;
        try { Directory.CreateDirectory(d); Process.Start(d); } catch (Exception ex) { MessageBox.Show(this, "Не открывается папка " + d + ":\n" + ex.Message, Text); }
    }

    void ShowPhoto(string f) {
        if (viewer == null || viewer.IsDisposed) {
            viewer = new PhotoViewer();
            viewer.Owner = this;
            SideBySide();
        }
        viewer.ShowFiles(strip.Files, f);
        if (!viewer.Visible) viewer.Show();
        viewer.Activate();
    }

    // фото - справа, форма - слева: печатать, глядя на фото; фото закрыли - форма на прежнее место
    void SideBySide() {
        Rectangle wa = Screen.FromControl(this).WorkingArea;
        int vw = Math.Max(380, Math.Min(wa.Width * 45 / 100, wa.Width - MinimumSize.Width));
        bool max = WindowState == FormWindowState.Maximized;
        Rectangle before = max ? RestoreBounds : Bounds;
        if (max || Bounds.Right > wa.Right - vw) {
            if (max) WindowState = FormWindowState.Normal;
            Bounds = new Rectangle(wa.Left, wa.Top, Math.Max(MinimumSize.Width, wa.Width - vw), wa.Height);
            viewer.FormClosed += delegate {
                if (IsDisposed || WindowState == FormWindowState.Minimized) return;
                Bounds = before;
                if (max) WindowState = FormWindowState.Maximized;
            };
        }
        viewer.Bounds = new Rectangle(wa.Right - vw, wa.Top, vw, wa.Height);
    }

    void RemovePhoto(string f) {
        if (MessageBox.Show(this, "Удалить фото «" + Path.GetFileName(f) + "» из папки договора?\n\nНа телефоне оно останется.", Text,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        try { File.Delete(f); }
        catch (Exception ex) { MessageBox.Show(this, "Не удалось удалить фото:\n" + ex.Message, Text); return; }
        strip.Mark(f, null, Color.Empty);
        RefreshPhotos();
    }

    void AddFromFolder() {
        string[] chosen;
        using (OpenFileDialog d = new OpenFileDialog()) {
            d.Title = "Добавить фото документов";
            d.Filter = "Фото (*.jpg; *.jpeg; *.png; *.bmp)|*.jpg;*.jpeg;*.png;*.bmp|Все файлы|*.*";
            d.Multiselect = true;
            string from = memory.Get("ФотоИзПапки", "");
            if (from.Length > 0 && Directory.Exists(from)) d.InitialDirectory = from;
            if (d.ShowDialog(this) != DialogResult.OK) return;
            chosen = d.FileNames;
        }
        AddPhotos(chosen);
    }

    // скопировать фото в папку договора (оригиналы остаются где были)
    public void AddPhotos(string[] chosen) {
        if (chosen.Length == 0) return;
        memory.Values["ФотоИзПапки"] = Path.GetDirectoryName(chosen[0]);
        memory.Save();
        string dir = PhotoTarget();
        List<string> added = new List<string>(), skipped = new List<string>();
        string error = null;
        try {
            Directory.CreateDirectory(dir);
            foreach (string f in chosen) {
                if (!Photo.IsPhoto(f)) { skipped.Add(Path.GetFileName(f)); continue; }
                if (string.Compare(Path.GetDirectoryName(Path.GetFullPath(f)), Path.GetFullPath(dir), StringComparison.OrdinalIgnoreCase) == 0) continue;
                string to = Photo.Free(dir, Path.GetFileName(f));
                File.Copy(f, to);
                added.Add(to);
            }
        } catch (Exception ex) { error = ex.Message; }
        if (Directory.Exists(dir)) photoDir = dir;
        string note = error != null ? "Не все фото скопировались: " + error : "Добавлено фото: " + added.Count;
        if (skipped.Count > 0) note += ". Не фото (пропущены): " + string.Join(", ", skipped.ToArray());
        Added(added, note, error != null ? St.Bad : St.Good);
    }

    void Added(List<string> added, string note, Color color) {
        if (added.Count > 0) { dirty = true; unsavedAuto = true; }
        importNote = note;
        RefreshPhotos();
        PhotoNote(note, color);
        if (added.Count > 0) ReadPassport(added);
    }

    // с какого времени фото на телефоне считаются новыми: после последнего забранного (в первый раз - за последний час)
    DateTime PhoneSince() {
        DateTime t;
        if (DateTime.TryParseExact(memory.Get("ТелефонФотоДо", ""), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out t)) return t;
        return DateTime.Now.AddHours(-1);
    }

    // last > 0 - забрать последние last фото, даже если они уже забирались раньше
    public void TakeFromPhone(int last) {
        if (phoneBusy) return;
        phoneBusy = true;
        phoneButton.Enabled = false;
        PhotoNote("Ищу телефон...", St.Hint);
        string dir = PhotoTarget();
        DateTime since = PhoneSince();
        int gen = photoGen;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            PhoneResult r = Phone.Take(dir, since, last, delegate(string t) { Post(delegate { if (gen == photoGen) PhotoNote(t, St.Hint); }); });
            Post(delegate { PhoneDone(r, dir, since, gen); });
        });
        // оболочка Windows (как Проводник) работает только в таком потоке
        th.SetApartmentState(System.Threading.ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
    }

    public bool Busy { get { return phoneBusy || mrzRunning > 0; } }
    int mrzRunning;
    public bool NoQuestions;                         // для проверки без окон: на вопросы - «нет»

    // для проверки: открыть фото по номеру, что написано про фото
    public void ShowPhotoAt(int i) { List<string> f = strip.Files; if (i < f.Count) ShowPhoto(f[i]); }
    public string PhotoStatus { get { return photoNote.Text + " | " + lichnyNote.Text + " | " + fioNote.Text + " | " + certNote.Text + " | " + meterNotes[0].Text + " | " + meterNotes[1].Text + " | папка: " + PhotoRef(); } }

    void PhoneDone(PhoneResult r, string dir, DateTime since, int gen) {
        phoneBusy = false;
        phoneButton.Enabled = true;
        if (r.Newest > since) {
            memory.Values["ТелефонФотоДо"] = r.Newest.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            memory.Save();
        }
        if (gen != photoGen) {
            if (r.Copied.Count > 0) PhotoNote("Фото с телефона попали в папку прошлого договора: " + dir, Warn);
            return;
        }
        if (r.Error != null) { PhotoNote(r.Error, St.Bad); return; }
        if (r.NoneNew) {
            PhotoNote("На телефоне «" + r.Device + "» нет новых фото (снятых после " + since.ToString("dd.MM.yyyy HH:mm") + ").", St.Hint);
            if (!NoQuestions && MessageBox.Show(this, "На телефоне нет новых фото - снятых после " + since.ToString("dd.MM.yyyy HH:mm") + ".\n\nЗабрать последние 3 фото с телефона?",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) TakeFromPhone(3);
            return;
        }
        if (Directory.Exists(dir)) photoDir = dir;
        string note = r.Copied.Count > 0 ? "С телефона «" + r.Device + "» забрано фото: " + r.Copied.Count : "Эти фото уже есть в папке договора";
        if (r.Failed.Count > 0) note += ". Не скопировались: " + string.Join(", ", r.Failed.ToArray()) + " - нажмите «Забрать фото с телефона» ещё раз";
        Added(r.Copied, note, r.Failed.Count > 0 ? St.Bad : St.Good);
    }

    // с фото паспорта: личный номер - со строк «<<<» последней страницы, Ф.И.О. (и номер) - с предпоследней
    void ReadPassport(List<string> files) {
        int gen = photoGen;
        List<string> list = new List<string>(files);
        List<string> mtypes = new List<string>();
        foreach (MeterType mt in lists.Meters()) mtypes.Add(mt.Full);
        foreach (string f in list) strip.Mark(f, "читаю...", St.Hint);
        mrzRunning++;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            lock (mrzLock) {
                foreach (string f in list) {
                    MrzResult r = null;
                    PageResult pg = null;
                    CertResult ce = null;
                    MeterResult me = null;
                    try {
                        r = Mrz.Read(f);
                        // строк «<<<» нет - может, это счётчик (штрихкод, табло, надпись типа)
                        if (!r.Valid) me = MeterPhoto.Read(f, mtypes);
                        bool meter = me != null && me.Found;
                        // может, страница паспорта с Ф.И.О.
                        if (!r.Valid && !meter) pg = PassportPage.Read(f);
                        // и не она - может, свидетельство о госрегистрации (адрес)
                        if (!r.Valid && !meter && (pg == null || !pg.Found)) ce = CertPage.Read(f);
                    } catch (Exception ex) { r = new MrzResult(); r.Info = "не открывается: " + ex.Message; pg = null; ce = null; me = null; }
                    string file = f;
                    MrzResult res = r;
                    PageResult page = pg;
                    CertResult cert = ce;
                    MeterResult mres = me;
                    Post(delegate { if (gen == photoGen) PhotoRead(file, res, page, cert, mres); });
                }
            }
            Post(delegate { mrzRunning--; if (gen == photoGen) PhotosRead(); });
        });
        th.IsBackground = true;
        th.Priority = System.Threading.ThreadPriority.BelowNormal;
        th.Start();
    }

    static bool Blr(MrzResult r) { return r.Line2.Length == 44 && r.Line2.Substring(10, 3) == "BLR"; }

    void PhotoRead(string f, MrzResult r, PageResult page, CertResult cert, MeterResult meter) {
        string name = Path.GetFileName(f);
        if (r != null && r.Info.StartsWith("не открывается")) { strip.Mark(f, "фото не открывается", St.Bad); return; }
        if (meter != null && meter.Found) {
            int i = ApplyMeter(meter, name);
            bool sure = (meter.Serial.Length > 0) && (meter.Reading.Length == 0 || meter.ReadingOk) && (meter.Type.Length == 0 || meter.TypeOk);
            strip.Mark(f, "счётчик " + i + (sure ? " прочитан" : " - проверьте"), sure ? St.Good : Warn);
            return;
        }
        if (cert != null && cert.Found) {
            lastCert = cert;
            certWhere = name;
            ApplyCert();
            strip.Mark(f, certSure ? "адрес прочитан" : "адрес - проверьте", certSure ? St.Good : Warn);
            return;
        }
        if (page != null && page.Found) {
            lastPage = page;
            pageWhere = name;
        } else if (r != null && r.Line2.Length > 0) {
            if (!r.Valid) { strip.Mark(f, "строки паспорта нечётко", Warn); mrzWeak = true; }
            else if (!Blr(r)) strip.Mark(f, "паспорт не РБ", St.Hint);
            else { strip.Mark(f, "личный номер прочитан", St.Good); lastMrz = r; mrzWhere = name; }
        } else strip.Mark(f, null, Color.Empty);
        ApplyPassport();
    }

    // прочитанное с фото - в поля (только в пустые; в заполненные - подсказка рядом)
    void ApplyPassport() {
        if (lastPage != null && lastMrz != null) PassportPage.Confirm(lastPage, lastMrz);
        if (lastPage != null) {
            foreach (string f in strip.Files)
                if (Path.GetFileName(f) == pageWhere) strip.Mark(f, lastPage.FioOk ? "ФИО прочитано" : "ФИО - проверьте", lastPage.FioOk ? St.Good : Warn);
            photoFio = lastPage.Surname.Length > 0 && lastPage.Name.Length > 0 ? lastPage.Fio : "";
            photoFioOk = lastPage.FioOk;
            if (photoFio.Length > 0 && photoFioOk && V("fio").Length == 0) SetV("fio", photoFio);
        }
        string pn = lastMrz != null ? lastMrz.Personal : lastPage != null && lastPage.IdOk ? lastPage.Id : "";
        if (pn.Length > 0) {
            photoPersonal = pn;
            if (V("lichny").Length == 0) SetV("lichny", pn);
        }
        LichnyNote();
        FioNote();
    }

    void PhotosRead() {
        List<string> parts = new List<string>();
        bool good = true;
        if (lastMrz != null) parts.Add("личный номер прочитан с фото «" + mrzWhere + "»");
        else if (lastPage != null && lastPage.IdOk) parts.Add("личный номер прочитан с фото «" + pageWhere + "»");
        else if (mrzWeak) {
            parts.Add("строки «<<<» прочитаны неуверенно - личный номер не подставлен: переснимите последнюю страницу ближе, ровно, без бликов");
            good = false;
        }
        if (lastPage != null) {
            if (photoFio.Length == 0) { parts.Add("Ф.И.О. на фото «" + pageWhere + "» прочитать не удалось"); good = false; }
            else if (photoFioOk) parts.Add("Ф.И.О. прочитано с фото «" + pageWhere + "»");
            else { parts.Add("Ф.И.О. с фото «" + pageWhere + "» прочитано неуверенно - проверьте подсказку под полем"); good = false; }
        }
        for (int i = 1; i <= 2; i++) {
            MeterResult m = photoMeter[i - 1];
            if (m == null) continue;
            List<string> got = new List<string>();
            if (m.Serial.Length > 0) got.Add("зав. №");
            if (m.Type.Length > 0) got.Add(m.TypeOk ? "тип" : "тип (проверьте)");
            if (m.Reading.Length > 0) got.Add(m.ReadingOk ? "показания" : "показания (проверьте)");
            parts.Add("счётчик " + i + " (фото «" + meterWhere[i - 1] + "»): " + (got.Count > 0 ? string.Join(", ", got.ToArray()) : "прочитать не удалось"));
            if (!(m.Reading.Length == 0 || m.ReadingOk) || !(m.Type.Length == 0 || m.TypeOk)) good = false;
        }
        if (lastCert != null) {
            if (certSure) parts.Add("адрес объекта прочитан со свидетельства (фото «" + certWhere + "»)");
            else { parts.Add("адрес со свидетельства (фото «" + certWhere + "») прочитан неуверенно - проверьте подсказку под адресом"); good = false; }
        }
        if (parts.Count == 0) { PhotoNote(importNote, photoNote.ForeColor); return; }
        string s = string.Join("; ", parts.ToArray());
        PhotoNote((importNote.Length > 0 ? importNote.TrimEnd('.') + ". " : "") + char.ToUpper(s[0]) + s.Substring(1) + ". Сверьте с документами.", good ? St.Good : Warn);
    }

    static string Same(string s) { return s.Trim().ToUpper().Replace('Ё', 'Е').Replace("  ", " "); }

    void FioNote() {
        string v = V("fio");
        if (photoFio.Length == 0) Note(fioNote, "", St.Hint, false);
        else if (Same(v) == Same(photoFio)) Note(fioNote, "как на фото паспорта - сверьте с паспортом", St.Good, false);
        else if (!photoFioOk) Note(fioNote, "с фото паспорта: " + photoFio + " - проверьте " + Doubtful() + "; щёлкните, чтобы подставить", Warn, true);
        else Note(fioNote, "на фото паспорта: " + photoFio + " (щёлкните, чтобы подставить)", Warn, true);
        fioNote.Cursor = photoFio.Length > 0 && Same(v) != Same(photoFio) ? Cursors.Hand : Cursors.Default;
    }

    // что в Ф.И.О. с фото прочитано неуверенно
    string Doubtful() {
        List<string> d = new List<string>();
        if (lastPage == null) return "";
        if (!lastPage.SurnameOk) d.Add("фамилию");
        if (!lastPage.NameOk) d.Add("имя");
        if (!lastPage.PatronymicOk) d.Add("отчество");
        return string.Join(", ", d.ToArray());
    }

    // ---------- счётчик с фото ----------

    // к какому счётчику фото: тот же заводской номер - к нему; иначе - к первому незаполненному
    int TargetMeter(MeterResult m) {
        int n = V("schetchikov") == "1 счётчик" ? 1 : 2;
        if (m.Serial.Length > 0)
            for (int i = 1; i <= n; i++)
                if (V("pu" + i + "_zav") == m.Serial || photoMeter[i - 1] != null && photoMeter[i - 1].Serial == m.Serial) return i;
        // снимок табло без наклейки - к тому счётчику, что снимали только что
        if (m.Serial.Length == 0 && lastMeterIdx > 0) return lastMeterIdx;
        for (int i = 1; i <= n; i++) if (V("pu" + i + "_zav").Length == 0 && photoMeter[i - 1] == null) return i;
        return 1;
    }

    // прочитанное с фото - в пустые поля счётчика (уверенное), остальное - подсказкой под типом
    int ApplyMeter(MeterResult m, string name) {
        int i = TargetMeter(m);
        MeterResult old = photoMeter[i - 1];
        if (old != null && (m.Serial.Length == 0 || old.Serial == m.Serial)) {
            // второй снимок того же счётчика: недостающее - с нового
            if (m.Serial.Length == 0) { m.Serial = old.Serial; m.SerialRows = old.SerialRows; }
            if (m.Type.Length == 0 || !m.TypeOk && old.TypeOk) { m.Type = old.Type; m.TypeOk = old.TypeOk; m.TypeText = old.TypeText; }
            if (m.Reading.Length == 0 || !m.ReadingOk && old.ReadingOk) { m.Reading = old.Reading; m.ReadingOk = old.ReadingOk; m.Code = old.Code; }
        }
        photoMeter[i - 1] = m;
        meterWhere[i - 1] = name;
        lastMeterIdx = i;
        PutMeter(i, false);
        MeterNote(i);
        return i;
    }

    // all - по щелчку на подсказке: всё прочитанное (и неуверенное), поверх введённого
    void PutMeter(int i, bool all) {
        MeterResult m = photoMeter[i - 1];
        string p = "pu" + i + "_";
        if (m.Serial.Length > 0 && (all || V(p + "zav").Length == 0)) SetV(p + "zav", m.Serial);
        if (m.Type.Length > 0 && (all || m.TypeOk && V(p + "model").Length == 0 && V(p + "tip").Length == 0)) {
            MeterType t = MeterType.Parse(m.Type);
            SetV(p + "model", t.Model);
            SetV(p + "korpus", t.Case);
            SetV(p + "modif", t.Mod);
        }
        if (m.Reading.Length > 0 && m.Reading.IndexOf('?') < 0 && (all || m.ReadingOk && V(p + "pokaz").Length == 0)) SetV(p + "pokaz", m.Reading);
        // год с наклейки - в дату поверки, квартал дописать
        if (m.Year.Length > 0 && (V(p + "poverka").Length == 0 || all && !V(p + "poverka").StartsWith(m.Year))) SetV(p + "poverka", m.Year + "-");
    }

    void MeterNote(int i) {
        Label l = meterNotes[i - 1];
        MeterResult m = photoMeter[i - 1];
        if (m == null) { Note(l, "", St.Hint, false); l.Cursor = Cursors.Default; return; }
        string p = "pu" + i + "_";
        List<string> diff = new List<string>();
        if (m.Serial.Length > 0 && V(p + "zav") != m.Serial) diff.Add("зав. № " + m.Serial);
        if (m.Type.Length > 0 && Lists.Squeeze(V(p + "tip")) != Lists.Squeeze(m.Type)) {
            diff.Add("тип " + m.Type + (m.TypeOk ? "" : " (надпись прочитана неточно - проверьте)"));
        }
        if (m.Reading.Length > 0 && V(p + "pokaz") != m.Reading) {
            string why = m.Reading.IndexOf('?') >= 0 ? " («?» - цифра не разобрана, впишите сами)" :
                !m.ReadingOk ? (m.Code.Length > 0 && m.Code != "1.8.0" ? " - на табло экран " + m.Code + ", а нужен 1.8.0 (сумма)" : " - проверьте по фото") : "";
            diff.Add("показания " + m.Reading + why);
        }
        if (m.Year.Length > 0 && !V(p + "poverka").StartsWith(m.Year)) diff.Add("год на наклейке " + m.Year);
        string quarter = V(p + "poverka").EndsWith("-") ? " Дата поверки: допишите квартал (год " + V(p + "poverka").TrimEnd('-') + " - с наклейки)." : "";
        if (diff.Count == 0) {
            Note(l, "Как на фото «" + meterWhere[i - 1] + "» - сверьте со счётчиком." + quarter, St.Good, false);
            l.Cursor = Cursors.Default;
        } else {
            Note(l, "С фото «" + meterWhere[i - 1] + "»: " + string.Join("; ", diff.ToArray()) + ". Щёлкните, чтобы подставить." + quarter, Warn, true);
            l.Cursor = Cursors.Hand;
        }
    }

    // ---------- адрес со свидетельства ----------

    // название из своих списков (сельсоветы, прежние н.п. и улицы) - прочитано верно, даже если буквы были неуверенные
    bool Known(string key, string v) {
        if (v.Length == 0) return true;
        ComboBox cb = inp[key] as ComboBox;
        if (cb != null) foreach (object o in cb.Items) if (Same(o.ToString()) == Same(v)) return true;
        return false;
    }

    // поля адреса со свидетельства
    static readonly string[] CertKeys = { "np_tip", "np", "selsovet", "ulica_tip", "ulica", "dom", "korp", "kv" };

    string[] CertValues() {
        CertResult c = lastCert;
        return new string[] { c.NpTip, c.Np, c.Selsovet, c.UlicaTip, c.Ulica, c.Dom, c.Korp, c.Kv };
    }

    bool CertSame() {
        string[] v = CertValues();
        for (int i = 0; i < CertKeys.Length; i++) if (Same(V(CertKeys[i])) != Same(v[i])) return false;
        return true;
    }

    void PutCert() {
        string[] v = CertValues();
        for (int i = 0; i < CertKeys.Length; i++) SetV(CertKeys[i], v[i]);
        if (lastCert.Kadastr.Length > 0 && lastCert.KadastrOk && V("kadastr").Length == 0) SetV("kadastr", lastCert.Kadastr);
    }

    // прочитанное со свидетельства - в пустые поля адреса, если уверенно; иначе подсказка под адресом
    void ApplyCert() {
        CertResult c = lastCert;
        // сельсовет - из закрытого списка: прочитанное с 1-2 ошибками («Воровлянский») - ближайший из справочника
        if (c.Selsovet.Length > 0 && !Known("selsovet", c.Selsovet)) {
            string best = null;
            int bd = 99, cnt = 0;
            foreach (string k in With(lists.Get("Сельсовет"), memory.Hist("сельсовет"))) {
                int d = PassportPage.Distance(Same(k), Same(c.Selsovet));
                if (d < bd) { bd = d; best = k; cnt = 1; } else if (d == bd) cnt++;
            }
            if (best != null && cnt == 1 && bd <= 2 && bd * 5 <= best.Length) { c.Selsovet = best; c.SelsovetSure = true; }
        }
        certSure = c.Found && c.Rest && (c.NpSure || Known("np", c.Np) || Same(c.Np) == "МИНСК") &&
            (c.SelsovetSure || Known("selsovet", c.Selsovet)) && (c.UlicaSure || Known("ulica", c.Ulica));
        bool empty = true;
        foreach (string k in CertKeys) if (k != "np_tip" && k != "ulica_tip" && V(k).Length > 0) empty = false;
        if (certSure && empty) PutCert();
        else if (c.Kadastr.Length > 0 && c.KadastrOk && V("kadastr").Length == 0 && CertSame()) SetV("kadastr", c.Kadastr);
        CertNote();
    }

    void CertNote() {
        if (lastCert == null) { Note(certNote, "", St.Hint, false); certNote.Cursor = Cursors.Default; return; }
        CertResult c = lastCert;
        string owner = "";
        // правообладатель - тот, с кем договор? (сравниваем фамилию)
        if (c.Owner.Length > 0 && V("fio").Length > 0) {
            string a = Same(c.Owner).Split(' ')[0], b = Same(V("fio")).Split(' ')[0];
            if (a != b && PassportPage.Distance(a, b) > 1) owner = " Правообладатель по свидетельству: " + c.Owner + " - не тот, с кем договор?";
        }
        string kad = c.Kadastr.Length > 0 && V("kadastr").Length > 0 && Digits(V("kadastr")) != c.Kadastr ? " Кадастровый номер в свидетельстве: " + c.Kadastr + "." : "";
        bool same = CertSame();
        if (same) Note(certNote, "Адрес как в свидетельстве (фото «" + certWhere + "») - сверьте с документом." + kad + owner, owner.Length > 0 ? Warn : St.Good, false);
        else Note(certNote, "Со свидетельства (фото «" + certWhere + "»): " + c.Short + (certSure ? "" : " - прочитано неуверенно, проверьте") +
            "; щёлкните, чтобы подставить." + kad + owner, Warn, true);
        certNote.Cursor = same ? Cursors.Default : Cursors.Hand;
    }

    void LichnyNote() {
        string v = V("lichny");
        if (photoPersonal.Length == 0) Note(lichnyNote, "", St.Hint, false);
        else if (v == photoPersonal) Note(lichnyNote, "как на фото паспорта - сверьте с паспортом", St.Good, false);
        else Note(lichnyNote, "на фото паспорта: " + photoPersonal + " (щёлкните, чтобы подставить)", Warn, true);
        lichnyNote.Cursor = photoPersonal.Length > 0 && v != photoPersonal ? Cursors.Hand : Cursors.Default;
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
            Dictionary<string, string> vals = Compose();
            List<string> missing = DocxTemplate.Fill(templateFile, tmp, vals, Sizes(vals));
            try {
                if (File.Exists(docx)) File.Delete(docx);
            } catch (IOException) {
                File.Delete(tmp);
                MessageBox.Show(this, "Договор «" + name + "» сейчас открыт в Word.\nЗакройте его в Word и нажмите «Сохранить договор» ещё раз.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            File.Move(tmp, docx);
            MovePhotos(name);
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
        sb.Append("фото=").Append(Memory.Escape(PhotoRef())).Append("\r\n");
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
            sb.Append("фото=").Append(Memory.Escape(PhotoRef())).Append("\r\n");
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
        photoGen++;
        photoDir = null;
        lastMrz = null;
        lastPage = null;
        lastCert = null;
        photoMeter[0] = photoMeter[1] = null;
        meterWhere[0] = meterWhere[1] = "";
        lastMeterIdx = 0;
        mrzWhere = pageWhere = importNote = photoPersonal = photoFio = certWhere = "";
        photoFioOk = mrzWeak = certSure = false;
        if (viewer != null && !viewer.IsDisposed) viewer.Close();
        strip.ClearMarks();
        RefreshPhotos();
        Note(photoNote, "", St.Hint, false);
        LichnyNote();
        FioNote();
        CertNote();
        MeterNote(1);
        MeterNote(2);
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
        string photos;
        photoDir = d.TryGetValue("фото", out photos) && photos.Length > 0 ? (Path.IsPathRooted(photos) ? photos : Path.Combine(photoRoot, photos)) : null;
        RefreshPhotos();
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
// Какие поля обязательны: галочки по разделам формы (у счётчиков - для обоих)
public class RequiredDialog : Form {
    public Dictionary<string, bool> Result;
    readonly Dictionary<string, CheckBox> boxes = new Dictionary<string, CheckBox>();

    public RequiredDialog(List<string[]> fields, Dictionary<string, bool> current, string[] defaults) {
        Text = "Обязательные поля";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(820, Math.Min(760, Screen.PrimaryScreen.WorkingArea.Height - 40));
        MinimumSize = new Size(600, 420);
        ShowInTaskbar = false;
        MinimizeBox = false;
        Label top = new Label();
        top.Dock = DockStyle.Top;
        top.Height = 78;
        top.Padding = new Padding(14, 10, 14, 6);
        top.Font = St.Base;
        top.ForeColor = St.Text;
        top.Text = "Отметьте поля, без которых договор не готов. Пустое обязательное поле считается в «Не заполнено полей» внизу " +
            "и в красных цифрах у разделов слева. Сохранить договор можно и с пустыми - программа предупредит.";
        Panel list = new Panel();
        list.Dock = DockStyle.Fill;
        list.AutoScroll = true;
        list.BackColor = St.Back;
        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 56;
        bottom.BackColor = Color.White;
        Button ok = Ui.Button("Сохранить", true), cancel = Ui.Button("Отмена", false), reset = Ui.Button("Как было сначала", false);
        ok.Size = new Size(140, 34); cancel.Size = new Size(110, 34); reset.Size = new Size(170, 34);
        bottom.Controls.Add(ok); bottom.Controls.Add(cancel); bottom.Controls.Add(reset);
        bottom.Resize += delegate {
            ok.Location = new Point(bottom.Width - ok.Width - 14, 11);
            cancel.Location = new Point(ok.Left - cancel.Width - 10, 11);
            reset.Location = new Point(14, 11);
        };
        ok.Click += delegate {
            Result = new Dictionary<string, bool>();
            foreach (KeyValuePair<string, CheckBox> kv in boxes) if (kv.Value.Checked) Result[kv.Key] = true;
            DialogResult = DialogResult.OK;
        };
        cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
        reset.Click += delegate {
            foreach (KeyValuePair<string, CheckBox> kv in boxes) kv.Value.Checked = Array.IndexOf(defaults, kv.Key) >= 0;
        };
        AcceptButton = ok;
        CancelButton = cancel;
        // разделы - полосами своего цвета, под ними галочки в две колонки
        List<Control> heads = new List<Control>();
        string card = null;
        int col = 0;
        foreach (string[] f in fields) {
            if (f[2] != card) {
                card = f[2];
                Label h = new Label();
                h.Text = card + (card == "Счётчики" ? " (для обоих счётчиков)" : "");
                h.Font = St.Bold;
                Tint t = St.TintOf(card);
                h.BackColor = t.Head;
                h.ForeColor = t.Dark;
                h.TextAlign = ContentAlignment.MiddleLeft;
                h.Padding = new Padding(8, 0, 0, 0);
                h.Tag = "head";
                list.Controls.Add(h);
                heads.Add(h);
                col = 0;
            }
            CheckBox cb = new CheckBox();
            cb.Text = f[1];
            cb.Checked = current.ContainsKey(f[0]);
            cb.Font = St.Base;
            cb.ForeColor = St.Text;
            cb.AutoEllipsis = true;
            cb.Tag = col++ % 2;
            boxes[f[0]] = cb;
            list.Controls.Add(cb);
        }
        list.Resize += delegate { Relayout(list); };
        Controls.Add(list);
        Controls.Add(top);
        Controls.Add(bottom);
        Shown += delegate { Relayout(list); };
    }

    static void Relayout(Panel list) {
        int w = list.ClientSize.Width - 30, y = 6 + list.AutoScrollPosition.Y, x0 = 14, half = w / 2;
        bool open = false;        // в строке уже стоит левая галочка
        foreach (Control c in list.Controls) {
            if ("head".Equals(c.Tag)) {
                if (open) { y += 25; open = false; }
                y += 8;
                c.SetBounds(x0, y, w, 26);
                y += 30;
                continue;
            }
            int k = (int)c.Tag;
            c.SetBounds(x0 + 8 + k * half, y, half - 12, 24);
            if (k == 1) { y += 25; open = false; } else open = true;
        }
    }
}

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
