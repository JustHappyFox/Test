# Поиск абонента в базе FoxPro (таблицы .dbf) - ТОЛЬКО ЧТЕНИЕ.
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
    byte[] headerBytes;

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
        return LoadCore(path, otherEncoding, query, progress, false, null, 0);
    }

    // keep - какие записи оставлять; maxRecords > 0 - прочитать только начало файла
    public static DbfTable Load(string path, RecordFilter keep, LoadProgress progress, int maxRecords) {
        return LoadCore(path, false, null, progress, false, keep, maxRecords);
    }

    // Только заголовок: список полей и число записей, без чтения самих записей
    public static DbfTable LoadHeader(string path) {
        return LoadCore(path, false, null, null, true, null, 0);
    }

    static DbfTable LoadCore(string path, bool otherEncoding, string query, LoadProgress progress, bool headerOnly,
                             RecordFilter keep, int maxRecords) {
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
                f.Offset = offset;
                offset += f.Length;
                if (f.Type != '0') t.Fields.Add(f);                  // '0' - служебное поле VFP
            }
            if (offset > recordLen)
                throw new InvalidDataException("Не удалось разобрать структуру таблицы.");
            t.Expected = count;
            t.headerBytes = new byte[headerLen];
            Buffer.BlockCopy(head, 0, t.headerBytes, 0, 32);
            Buffer.BlockCopy(desc, 0, t.headerBytes, 32, desc.Length);
            if (headerOnly) return t;

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
                if (progress != null) progress(index, limit);
            }
        }
        t.Seconds = (DateTime.Now - started).TotalSeconds;
        return t;
    }

    // Сохранить записи, что в памяти, отдельным файлом DBF той же структуры (быстрый запуск в следующий раз)
    public void SaveCopy(string file) {
        string tmp = file + ".tmp";
        using (FileStream fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16)) {
            byte[] head = (byte[])headerBytes.Clone();
            byte[] n = BitConverter.GetBytes(Count);
            Buffer.BlockCopy(n, 0, head, 4, 4);
            fs.Write(head, 0, head.Length);
            for (int b = 0; b < data.Count; b++) {
                int records = Math.Min(BlockRecords, Count - b * BlockRecords);
                fs.Write(data[b], 0, records * recordLength);
            }
            fs.WriteByte(0x1A);
        }
        if (File.Exists(file)) File.Delete(file);
        File.Move(tmp, file);
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
            folder = Places.FindFolder(ReadSetting(0));
            Cursor = Cursors.Default;
            FillTables(lastTable);
            searchBox.Focus();
        };
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
        try { files = Places.DbfFiles(folder); }
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
        Card.Show(this, table, view[viewRow]);
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

// Подписи полей таблицы абонентов - как на карточке "АБОНЕНТ" в старой программе
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
public static class Card {
    public static void Show(IWin32Window owner, DbfTable table, int row) {
        Show(owner, table, row, null);
    }

    // extra - пары "подпись, значение", показываются первыми
    public static void Show(IWin32Window owner, DbfTable table, int row, string[] extra) {
        object[] values;
        using (MemoReader memo = table.OpenMemo()) values = table.GetRow(row, memo);
        Form card = new Form();
        card.Text = "Карточка - " + Path.GetFileName(table.FilePath);
        card.Size = new Size(650, 650);
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
        if (extra != null)
            for (int i = 0; i + 1 < extra.Length; i += 2) g.Rows.Add(extra[i], extra[i + 1]);
        for (int j = 0; j < values.Length; j++)
            g.Rows.Add(FieldNames.Title(table.Fields[j].Name), DbfTable.Format(values[j], true));
        card.Controls.Add(g);
        card.Show(owner);
    }
}

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
    readonly Dictionary<string, string> names = new Dictionary<string, string>();
    readonly Dictionary<string, string> lowerNames = new Dictionary<string, string>();

    public int Count { get { return names.Count; } }

    public string Name(string code) {
        string s;
        return names.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public string LowerName(string code) {
        string s;
        return lowerNames.TryGetValue(NormalizeCode(code), out s) ? s : null;
    }

    public double Coverage(List<string> codes) {
        if (codes.Count == 0) return 0;
        int hit = 0;
        foreach (string c in codes) if (names.ContainsKey(c)) hit++;
        return (double)hit / codes.Count;
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

    // Найти таблицу, где поле field - уникальный код с текстом при нём (например, IDNP в таблице населённых пунктов)
    static Dictionary<string, LinkRow> FindLink(string folder, List<TableInfo> tables, string field, List<string> needed,
                                                string exclude, bool nested, out string found) {
        found = null;
        foreach (TableInfo info in tables) {
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
        List<TableInfo> tables = null;
        Dictionary<int, Dictionary<string, LinkRow>> links = new Dictionary<int, Dictionary<string, LinkRow>>();
        for (int j = 0; j < d.Fields.Count; j++) {
            if (j == key || d.Fields[j].Type != 'N') continue;
            List<string> codes = new List<string>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>();
            for (int r = 0; r < d.Count; r++) {
                string c = NormalizeCode(Value(d, r, j));
                if (c.Length > 0 && c != "0" && !seen.ContainsKey(c)) { seen[c] = true; codes.Add(c); }
            }
            if (codes.Count == 0) continue;
            if (tables == null) tables = TableInfo.List(folder, null);
            string found;
            Dictionary<string, LinkRow> map = FindLink(folder, tables, d.Fields[j].Name, codes, dir.TableName, true, out found);
            if (map != null) links[j] = map;
            else Report.Line("    поле " + d.Fields[j].Name + ": справочник к нему не нашёлся");
        }

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
            foreach (KeyValuePair<int, Dictionary<string, LinkRow>> link in links) {
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
            dir.lowerNames[code] = display.ToString().ToLowerInvariant();
        }
        return dir;
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

    // Ищем таблицу, где есть поле с теми же кодами (каждый код - один раз) и текстом при них.
    // Берём, только если там нашлись почти все коды улиц абонентов.
    public static StreetDirectory Detect(string folder, string abonTable, List<string> codes, LoadProgress progress) {
        List<TableInfo> tables = TableInfo.List(folder, null);
        string bestTable = null, bestKey = null;
        double bestCoverage = 0.9;
        int bestRecords = int.MaxValue;
        for (int i = 0; i < tables.Count; i++) {
            if (progress != null) progress(i, tables.Count);
            TableInfo info = tables[i];
            if (info.Abonents || info.Records < 2 || info.Records > 50000) continue;
            if (string.Compare(info.Name, abonTable, StringComparison.OrdinalIgnoreCase) == 0) continue;
            if (info.Name.ToUpperInvariant().StartsWith("ARC")) continue;
            string path = Path.Combine(folder, info.Name);
            DbfTable d;
            try {
                DbfTable h = DbfTable.LoadHeader(path);
                bool text = false;
                foreach (DbfField f in h.Fields) if (f.Type == 'C' && f.Length >= 5) text = true;
                if (!text) continue;
                d = DbfTable.Load(path, false);
            } catch (Exception) {
                continue;
            }
            if (d.Count < 2) continue;
            for (int k = 0; k < d.Fields.Count; k++) {
                DbfField f = d.Fields[k];
                if (!((f.Type == 'N' && f.Length <= 10) || (f.Type == 'C' && f.Length <= 10))) continue;
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
                // справочник со ссылками на населённый пункт и сельсовет (IDNP, IDSS) лучше просто списка улиц
                bool linked = false;
                foreach (DbfField g in d.Fields) {
                    string n = g.Name.ToUpperInvariant();
                    if (g.Type == 'N' && g != f && (n.Contains("NP") || n.Contains("SS"))) linked = true;
                }
                if (hit > 0)
                    Report.Line("    кандидат " + info.Name + "." + f.Name + ": записей " + d.Count + ", кодов улиц нашлось " + hit + " из " + codes.Count
                        + " (" + (int)(coverage * 100) + "%)" + (coverage >= 0.9 ? (linked ? ", есть ссылки на населённый пункт/сельсовет" : "") : " - мало, не подходит"));
                if (coverage >= 0.9 && linked) coverage += 0.05;
                if (coverage > bestCoverage || (coverage == bestCoverage && bestTable != null && d.Count < bestRecords)) {
                    bestCoverage = coverage;
                    bestRecords = d.Count;
                    bestTable = info.Name;
                    bestKey = f.Name;
                }
            }
        }
        if (bestTable == null) {
            Report.Line("    справочник не найден: нет таблицы, где нашлись бы 90% кодов улиц");
            return null;
        }
        Report.Line("    выбран " + bestTable + "." + bestKey + " - больше всех кодов улиц");
        return Load(folder, bestTable, bestKey);
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

    // слова, которые в адресе ничего не значат: тип улицы и населённого пункта
    static readonly string[] Skip = {
        "ул", "улица", "пер", "переулок", "пр", "пр-т", "просп", "проспект", "пл", "площадь", "б-р", "бульвар",
        "ш", "шоссе", "туп", "тупик", "проезд", "мкр", "микрорайон", "наб", "набережная", "тракт",
        "г", "город", "аг", "агрогородок", "п", "пос", "поселок", "посёлок", "гп", "рп", "дер", "деревня",
        "с", "село", "х", "хутор", "с/с", "сс", "сельсовет", "р-н", "район", "обл", "область" };

    public AbonentQuery(DbfTable t, StreetDirectory dir) {
        this.t = t;
        accountFields = Existing(new string[] { "NAB", "LS", "LIC", "LSCH" });
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

    public List<int> Find(string query, SearchMode mode) {
        List<int> found;
        switch (mode) {
            case SearchMode.Account: found = FindAccount(query); break;
            case SearchMode.Address: found = FindAddress(query); break;
            case SearchMode.Meter: found = FindMeter(query); break;
            default: found = FindWords(query, nameFields); break;
        }
        if (found.Count <= 20000) SortByAddress(found);
        return found;
    }

    // лицевой: берём из строки только цифры, нули впереди не важны
    List<int> FindAccount(string query) {
        List<int> found = new List<int>();
        StringBuilder digits = new StringBuilder();
        foreach (char c in query) if (c >= '0' && c <= '9') digits.Append(c);
        LastExplain = "цифр: " + digits.Length;
        if (digits.Length == 0) return found;
        string q = digits.ToString();
        SearchWord w = Part(q);
        foreach (int f in accountFields) {
            for (int i = 0; i < RowCount; i++) {
                int row = RowAt(i);
                if (t.FieldEquals(row, f, w)) found.Add(row);
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
        List<SearchWord> text = new List<SearchWord>();
        List<string> textLower = new List<string>();
        List<string> numbers = new List<string>();
        List<string> early = new List<string>();   // числа до названия: номер строки из списка или сам дом
        SearchWord house = null, korp = null, flat = null;
        string pending = null;   // "дом", "корп" или "кв" перед числом
        bool seenText = false;
        int skipped = 0;
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
            if (Array.IndexOf(Skip, w) >= 0) { pending = null; skipped++; continue; }

            if (IsDigitAt(w, 0)) {
                if (IsDigits(w) && w.Length >= 5) { pending = null; skipped++; continue; }   // лицевой или код - не адрес
                if (pending == "корп" && korp == null) korp = t.PrepareWord(w);
                else if (pending == "кв" && flat == null) flat = t.PrepareWord(w);
                else if (pending == "дом" && house == null) house = t.PrepareWord(w);
                else if (seenText) numbers.Add(w);
                else early.Add(w);
                pending = null;
                continue;
            }
            pending = null;
            if (HasDigit(w)) { skipped++; continue; }
            text.Add(Part(w));
            textLower.Add(w);
            seenText = true;
        }
        // числа до названия улицы - обычно номер строки из списка, их пропускаем;
        // но если слов нет совсем ("18"), это и есть дом
        if (text.Count == 0) numbers.InsertRange(0, early);
        // числа без пояснений: первое - дом, второе - квартира
        foreach (string n in numbers) {
            if (house == null) house = t.PrepareWord(n);
            else if (flat == null) flat = t.PrepareWord(n);
            else if (korp == null) korp = t.PrepareWord(n);
        }
        StringBuilder explain = new StringBuilder();
        explain.Append("слов адреса: " + text.Count + ", дом: " + (house != null ? "да" : "нет") + ", корпус: " + (korp != null ? "да" : "нет")
            + ", квартира: " + (flat != null ? "да" : "нет") + ", пропущено служебных слов и номеров: " + (skipped + (text.Count > 0 ? early.Count : 0)));
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

// Главное окно: одна строка поиска по базе абонентов
public class AbonentSearch : Form {
    readonly string settingsFile;
    readonly string viewerSettings;
    string folder;
    string tableName;
    DbfTable table;
    AbonentQuery query;
    StreetDirectory directory;
    bool streetCodes;    // в UL у абонентов коды улиц, а не названия
    int[] mine;          // абоненты выбранных сельсоветов, если улицы записаны названиями (отбор после чтения)
    string mineNote = "";
    readonly CheckBox onlyMine = new CheckBox();
    Dictionary<string, bool> allowed;   // коды улиц моих сельсоветов - читаем только их абонентов (null - всю базу)
    DbfField ulField;
    int totalInBase;
    DateTime dataTime;
    string dataNote = "";
    bool refreshing;
    bool suppressMine;
    int generation;      // номер загрузки: фоновое обновление от старой загрузки не подменяет новую

    static readonly string[] DefaultSelsovety = { "Шершунский", "Папернянский", "Юзуфовский", "Острошицкий", "Боровлянский", "Колодищанский" };
    bool loading;
    List<int> found = new List<int>();
    internal SearchMode mode = SearchMode.Account;
    readonly RadioButton[] modeButtons = new RadioButton[4];
    readonly Label hint = new Label();

    static readonly string[] ModeNames = { "&Лицевой счёт", "&Адрес", "№ &счётчика", "&ФИО" };
    static readonly string[] Prompts = { "лицевой счёт", "адрес", "номер счётчика", "фамилию" };
    static readonly string[] Hints = {
        "Например:  044060216   (нули впереди можно не писать)",
        "Вставьте адрес целиком:  Острошицкий Городок Заводская ул. 18   или коротко:  заводская 18 кв 5",
        "Номер счётчика целиком или его часть",
        "Например:  иванов   или   иванов иван" };

    internal readonly TextBox queryBox = new TextBox();
    readonly DataGridView grid = new DataGridView();
    readonly Label status = new Label();
    readonly Panel bottom = new Panel();
    readonly Timer timer = new Timer();
    readonly Timer logTimer = new Timer();   // поиск пишем в отчёт, когда ввод успокоился

    public static void Run(string dir) {
        Application.EnableVisualStyles();
        Application.Run(new AbonentSearch(Path.Combine(dir, "poisk.txt"), Path.Combine(dir, "nastroiki.txt")));
    }

    public AbonentSearch(string settingsFile, string viewerSettings) {
        this.settingsFile = settingsFile;
        this.viewerSettings = viewerSettings;
        Text = "Поиск абонента (только чтение)";
        Size = new Size(1000, 650);
        StartPosition = FormStartPosition.CenterScreen;

        Panel top = new Panel();
        top.Dock = DockStyle.Top;
        top.Height = 100;
        Label caption = new Label();
        caption.Text = "Искать по:";
        caption.AutoSize = true;
        caption.Margin = new Padding(6, 9, 6, 0);
        // первая строка - в ряд, по фактической ширине надписей
        FlowLayoutPanel row = new FlowLayoutPanel();
        row.AutoSize = true;
        row.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        row.WrapContents = false;
        row.Location = new Point(4, 2);
        row.Controls.Add(caption);
        for (int i = 0; i < modeButtons.Length; i++) {
            RadioButton rb = new RadioButton();
            rb.Text = ModeNames[i];
            rb.AutoSize = true;
            rb.Margin = new Padding(3, 7, 8, 0);
            rb.Tag = (SearchMode)i;
            rb.CheckedChanged += new EventHandler(OnModeChanged);
            modeButtons[i] = rb;
            row.Controls.Add(rb);
        }
        queryBox.Font = new Font(Font.FontFamily, 14f);
        queryBox.Location = new Point(12, 38);
        queryBox.Width = top.Width - 24;
        queryBox.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        queryBox.TextChanged += new EventHandler(OnQueryChanged);
        queryBox.KeyDown += new KeyEventHandler(OnQueryKey);
        hint.ForeColor = SystemColors.GrayText;
        hint.AutoSize = true;
        hint.Location = new Point(12, 76);
        Button listButton = new Button();
        listButton.Text = "Поиск по списку...";
        listButton.AutoSize = true;
        listButton.Margin = new Padding(12, 3, 3, 0);
        listButton.Click += delegate { OpenBatch(); };
        row.Controls.Add(listButton);
        onlyMine.Text = "Только мои сельсоветы";
        onlyMine.AutoSize = true;
        onlyMine.Checked = true;
        onlyMine.Margin = new Padding(14, 7, 3, 0);
        onlyMine.CheckedChanged += delegate { OnOnlyMineChanged(); };
        row.Controls.Add(onlyMine);
        Button ssButton = new Button();
        ssButton.Text = "Сельсоветы...";
        ssButton.AutoSize = true;
        ssButton.Margin = new Padding(6, 3, 3, 0);
        ssButton.Click += delegate { EditSelsovety(); };
        row.Controls.Add(ssButton);
        top.Controls.Add(row);
        Load += delegate { Width = Math.Max(Width, row.PreferredSize.Width + 60); };
        top.Controls.Add(queryBox);
        top.Controls.Add(hint);

        grid.Dock = DockStyle.Fill;
        grid.VirtualMode = true;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = false;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.CellValueNeeded += new DataGridViewCellValueEventHandler(OnCellValueNeeded);
        grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0) OpenCard(e.RowIndex); };
        grid.KeyDown += new KeyEventHandler(OnGridKey);

        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 34;
        status.Dock = DockStyle.Fill;
        status.TextAlign = ContentAlignment.MiddleLeft;
        bottom.Controls.Add(status);
        bottom.Controls.Add(BottomButton("Отчёт", delegate { OpenReport(); }));
        bottom.Controls.Add(BottomButton("Все таблицы...", delegate { new DbfViewer(this.viewerSettings).Show(); }));
        bottom.Controls.Add(BottomButton("Где искать...", delegate { ChooseBase(); }));
        bottom.Controls.Add(BottomButton("Обновить", delegate { if (allowed != null && table != null) RefreshInBackground(); else LoadTable(); }));

        timer.Interval = 300;
        timer.Tick += delegate { timer.Stop(); ApplySearch(); };
        logTimer.Interval = 2000;
        logTimer.Tick += delegate { logTimer.Stop(); LogSearch(); };

        Controls.Add(grid);
        Controls.Add(bottom);
        Controls.Add(top);

        Shown += delegate { Start(); };
    }

    Button BottomButton(string text, EventHandler click) {
        Button b = new Button();
        b.Text = text;
        b.Dock = DockStyle.Right;
        b.Width = 120;
        b.Click += click;
        return b;
    }

    // адрес, скопированный из Excel, приходит с табуляциями - показываем их пробелами
    void OnQueryChanged(object sender, EventArgs e) {
        string text = queryBox.Text;
        if (text.IndexOf('\t') >= 0 || text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) {
            int caret = queryBox.SelectionStart;
            queryBox.Text = text.Replace("\r\n", " ").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').TrimEnd();
            queryBox.SelectionStart = Math.Min(caret, queryBox.Text.Length);
            return;   // смена текста вызовет этот обработчик ещё раз
        }
        timer.Stop();
        timer.Start();
    }

    void OnModeChanged(object sender, EventArgs e) {
        RadioButton rb = (RadioButton)sender;
        if (!rb.Checked) return;
        mode = (SearchMode)rb.Tag;
        hint.Text = Hints[(int)mode];
        SaveSettings();
        queryBox.Focus();
        queryBox.SelectAll();
        ApplySearch();
    }

    void SaveSettings() {
        if (folder == null) return;
        Places.SaveSettings(settingsFile, new string[] {
            folder, tableName == null ? "" : tableName, ((int)mode).ToString(),
            directory == null ? "" : directory.TableName, directory == null ? "" : directory.KeyField });
    }

    void OpenReport() {
        LogSearch();
        if (Report.FilePath == null || !File.Exists(Report.FilePath)) {
            MessageBox.Show(this, "Отчёт не получилось записать (папка программы только для чтения?).", Text);
            return;
        }
        try { Process.Start("notepad.exe", "\"" + Report.FilePath + "\""); }
        catch (Exception ex) { MessageBox.Show(this, "Отчёт лежит здесь:\n" + Report.FilePath + "\n\n" + ex.Message, Text); }
    }

    string lastLogged;

    void LogSearch() {
        logTimer.Stop();
        if (query == null || DbfTable.SplitWords(queryBox.Text).Length == 0) return;
        string line = "Поиск по: " + Prompts[(int)mode] + "; " + query.LastExplain + " -> найдено " + found.Count;
        if (line == lastLogged) return;
        lastLogged = line;
        Report.Search(line);
    }

    void Start() {
        Report.Start(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(settingsFile)), "otchet.txt"));
        int saved;
        if (int.TryParse(Places.ReadSetting(settingsFile, 2), out saved) && saved >= 0 && saved < modeButtons.Length) mode = (SearchMode)saved;
        modeButtons[(int)mode].Checked = true;
        hint.Text = Hints[(int)mode];
        queryBox.Focus();
        SetBusy(true, "Ищу папку базы...");
        folder = Places.FindFolder(Places.ReadSetting(settingsFile, 0));
        Report.Line("Папка базы: " + folder);
        tableName = Places.ReadSetting(settingsFile, 1);
        if (tableName != null && File.Exists(Path.Combine(folder, tableName))) {
            Report.Line("Таблица абонентов: " + tableName + " (запомнена с прошлого раза)");
        } else {
            tableName = null;
            try {
                List<TableInfo> tables = TableInfo.List(folder, delegate(int done, int total) {
                    if (IsDisposed) throw new OperationCanceledException();
                    status.Text = "Ищу таблицу абонентов: проверено " + done + " из " + total + " таблиц...";
                    Application.DoEvents();
                });
                if (tables.Count > 0 && tables[0].Abonents) tableName = tables[0].Name;
                Report.Line("Таблиц в папке: " + tables.Count + ". Похожие на таблицу абонентов (есть NAB, FIO и DOM или UL):");
                int shown = 0;
                foreach (TableInfo info in tables) if (info.Abonents && shown++ < 10) Report.Line("    " + info);
                if (shown == 0) Report.Line("    таких нет");
                Report.Line("Таблица абонентов: " + (tableName == null ? "не выбрана" : tableName + " (выбрана сама: не архив, больше всех записей)"));
            } catch (OperationCanceledException) {
                return;
            } catch (Exception ex) {
                SetBusy(false, "Не удалось открыть папку " + folder + ": " + ex.Message + "   Нажмите \"Где искать...\".");
                return;
            }
        }
        SetBusy(false, null);
        if (tableName == null) {
            status.Text = "Не нашёл таблицу абонентов в папке " + folder + ". Нажмите \"Где искать...\" и выберите её.";
            return;
        }
        LoadTable();
    }

    // ---- загрузка: справочники -> коды улиц моих сельсоветов -> только их абоненты ----

    string Dir { get { return Path.GetDirectoryName(Path.GetFullPath(settingsFile)); } }
    string CacheFile { get { return Path.Combine(Dir, "kesh_abonenty.dbf"); } }
    string CacheInfo { get { return Path.Combine(Dir, "kesh_abonenty.txt"); } }

    // что лежит в сохранённой выборке: база, справочник и список сельсоветов
    string CacheKey() {
        return Path.Combine(folder, tableName) + "|" + (directory == null ? "" : directory.TableName + "." + directory.KeyField)
            + "|" + string.Join(";", ReadSelsovety());
    }

    internal void LoadTable() {
        if (loading || tableName == null) return;
        string path = Path.Combine(folder, tableName);
        generation++;
        SetBusy(true, "Готовлю справочники...");
        try {
            DbfTable header = DbfTable.LoadHeader(path);
            totalInBase = header.Expected;
            Report.Line("Таблица абонентов " + tableName + ": записей по заголовку " + header.Expected);
            directory = PrepareDirectory(path, header);
            PrepareFilter(header);
            if (allowed != null && TryLoadCache(path)) {
                AfterLoad();
                if (DateTime.Now - dataTime > TimeSpan.FromHours(1)) RefreshInBackground();
                return;
            }
            SetBusy(true, allowed != null ? "Читаю базу и отбираю абонентов моих сельсоветов..." : "Читаю всю базу абонентов...");
            table = DbfTable.Load(path, MakeFilter(), delegate(int done, int expected) {
                if (IsDisposed) throw new OperationCanceledException();
                status.Text = "Читаю базу " + tableName + ": " + done + " из " + expected + " записей"
                    + (allowed != null ? " (оставляю только мои сельсоветы)" : "") + "...";
                Application.DoEvents();
            }, 0);
            dataTime = DateTime.Now;
            ReportRead("Прочитана база", table);
            if (allowed != null) SaveCache(table, CacheKey(), dataTime);
            AfterLoad();
        } catch (OperationCanceledException) {
            return;
        } catch (Exception ex) {
            table = null;
            query = null;
            Report.Line("ОШИБКА чтения " + path + ": " + ex.Message);
            SetBusy(false, "Не удалось прочитать " + path + ": " + ex.Message);
        }
    }

    void ReportRead(string what, DbfTable t) {
        double mb = t.BytesRead / 1048576.0;
        Report.Line(what + ": " + mb.ToString("0") + " МБ за " + t.Seconds.ToString("0.0") + " с (" + (t.Seconds > 0 ? (mb / t.Seconds).ToString("0.0") : "?")
            + " МБ/с), записей " + t.Total + ", удалённых пропущено " + t.Deleted + ", оставлено " + t.Count
            + (allowed != null ? " (мои сельсоветы)" : " (вся база)"));
    }

    void AfterLoad() {
        query = new AbonentQuery(table, directory);
        if (allowed == null && directory == null && onlyMine.Checked) FindMine();   // улицы названиями - отбор по тексту адреса
        else query.Scope = null;
        Report.Line("Поля таблицы абонентов (по загруженным записям):");
        Report.Fields(table, "    ");
        Report.Line("Где ищется: лицевой - " + query.FieldList(SearchMode.Account) + " | адрес - " + query.FieldList(SearchMode.Address)
            + " | счётчик - " + query.FieldList(SearchMode.Meter) + " | ФИО - " + query.FieldList(SearchMode.Name));
        Report.Line("--- поиски (сами слова не записываются) ---");
        dataNote = "данные на " + dataTime.ToString("dd.MM HH:mm");
        SaveSettings();
        BuildColumns();
        SetBusy(false, null);
        ApplySearch();
    }

    // Справочник улиц: запомненный с прошлого раза или ищем по кодам улиц из начала базы
    StreetDirectory PrepareDirectory(string path, DbfTable header) {
        streetCodes = false;
        int ul = header.FieldIndex("UL");
        ulField = ul < 0 ? null : header.Fields[ul];
        if (ul < 0) return null;
        string savedTable = Places.ReadSetting(settingsFile, 3), savedKey = Places.ReadSetting(settingsFile, 4);
        if (savedTable != null && savedKey != null && File.Exists(Path.Combine(folder, savedTable))) {
            try {
                Report.Line("Справочник улиц с прошлого раза: " + savedTable + "." + savedKey);
                StreetDirectory d = StreetDirectory.Load(folder, savedTable, savedKey);
                if (d != null && d.Count > 0) { streetCodes = true; return d; }
            } catch (Exception ex) {
                Report.Line("    не открылся: " + ex.Message);
            }
        }
        status.Text = "Смотрю, как записаны улицы...";
        status.Refresh();
        DbfTable sample = DbfTable.Load(path, null, null, 30000);
        List<string> codes = StreetDirectory.Codes(sample, sample.FieldIndex("UL"));
        streetCodes = codes != null;
        if (codes == null) return null;
        Report.Line("Ищу справочник улиц среди таблиц папки:");
        try {
            return StreetDirectory.Detect(folder, tableName, codes, delegate(int done, int total) {
                if (IsDisposed) throw new OperationCanceledException();
                status.Text = "В поле улицы - коды. Ищу справочник улиц: проверено " + done + " из " + total + " таблиц...";
                Application.DoEvents();
            });
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            Report.Line("    ошибка: " + ex.Message);
            return null;
        }
    }

    // Коды улиц моих сельсоветов - по справочникам. null - читаем всю базу.
    void PrepareFilter(DbfTable header) {
        allowed = null;
        mineNote = "";
        string[] names = ReadSelsovety();
        onlyMine.Text = "Только мои сельсоветы (" + names.Length + ")";
        if (!onlyMine.Checked) { Report.Line("Отбор по сельсоветам выключен - читаю всю базу"); return; }
        if (names.Length == 0) { Report.Line("Сельсоветы: список пуст - читаю всю базу"); return; }
        if (directory == null || ulField == null) {
            if (streetCodes) mineNote = "   Справочник улиц не нашёлся - сельсоветы не проверить, показана вся база.";
            return;   // улицы названиями - отберём после чтения по тексту адреса
        }
        string report;
        Dictionary<string, bool> codes = directory.CodesMatching(names, out report);
        Report.Line(report);
        if (codes.Count == 0) {
            mineNote = "   Названия сельсоветов в справочниках не нашлись - показана вся база. Нажмите \"Отчёт\" и пришлите его.";
            onlyMine.Text = "Только мои сельсоветы (не нашлись)";
            return;
        }
        allowed = codes;
    }

    RecordFilter MakeFilter() {
        if (allowed == null || ulField == null) return null;
        Dictionary<string, bool> codes = allowed;
        DbfField f = ulField;
        return delegate(byte[] record, int start) {
            return codes.ContainsKey(StreetDirectory.NormalizeCode(Encoding.ASCII.GetString(record, start + f.Offset, f.Length)));
        };
    }

    bool TryLoadCache(string path) {
        if (!File.Exists(CacheFile) || !File.Exists(CacheInfo)) return false;
        string key = Places.ReadSetting(CacheInfo, 0);
        long ticks;
        if (key != CacheKey() || !long.TryParse(Places.ReadSetting(CacheInfo, 1), out ticks)) {
            Report.Line("Сохранённая выборка не подходит (поменялись база, справочник или сельсоветы) - читаю базу");
            return false;
        }
        try {
            status.Text = "Беру сохранённую выборку абонентов...";
            status.Refresh();
            table = DbfTable.Load(CacheFile, false);
            table.FilePath = path;   // примечания - из исходной базы
            dataTime = new DateTime(ticks);
            Report.Line("Абоненты взяты из сохранённой выборки от " + dataTime.ToString("dd.MM.yyyy HH:mm") + ": " + table.Count
                + " записей за " + table.Seconds.ToString("0.0") + " с" + (DateTime.Now - dataTime > TimeSpan.FromHours(1) ? "; ей больше часа - обновляю в фоне" : ""));
            return true;
        } catch (Exception ex) {
            Report.Line("Сохранённая выборка не открылась: " + ex.Message);
            table = null;
            return false;
        }
    }

    void SaveCache(DbfTable t, string key, DateTime when) {
        try {
            t.SaveCopy(CacheFile);
            Places.SaveSettings(CacheInfo, new string[] { key, when.Ticks.ToString() });
        } catch (Exception ex) {
            Report.Line("Сохранить выборку не удалось: " + ex.Message);
        }
    }

    // Свежие данные из базы - в фоне, пока работаете со старыми; потом тихо подменяем
    void RefreshInBackground() {
        if (refreshing || loading || table == null || tableName == null) return;
        refreshing = true;
        string path = Path.Combine(folder, tableName);
        RecordFilter filter = MakeFilter();
        string key = CacheKey();
        bool save = allowed != null;
        int gen = generation;
        SetDataNote("обновляю данные из базы...");
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            DbfTable fresh = null;
            string error = null;
            int shown = 0;
            try {
                fresh = DbfTable.Load(path, filter, delegate(int done, int expected) {
                    if (done - shown < 50000) return;
                    shown = done;
                    string note = "обновляю данные из базы: " + (int)(done * 100L / Math.Max(expected, 1)) + "%";
                    try { BeginInvoke(new MethodInvoker(delegate { SetDataNote(note); })); } catch (Exception) { }
                }, 0);
                if (save) SaveCache(fresh, key, DateTime.Now);
            } catch (Exception ex) {
                error = ex.Message;
            }
            try { BeginInvoke(new MethodInvoker(delegate { FinishRefresh(fresh, error, gen); })); } catch (Exception) { }
        });
        th.IsBackground = true;
        th.Start();
    }

    void FinishRefresh(DbfTable fresh, string error, int gen) {
        refreshing = false;
        if (gen != generation) return;   // пока обновлялось, базу уже перечитали заново
        if (error != null) {
            Report.Line("Обновить данные не удалось: " + error);
            SetDataNote("данные на " + dataTime.ToString("dd.MM HH:mm") + " (обновить не удалось)");
            return;
        }
        fresh.FilePath = Path.Combine(folder, tableName);
        table = fresh;
        dataTime = DateTime.Now;
        ReportRead("Данные обновлены из базы", table);
        query = new AbonentQuery(table, directory);
        dataNote = "данные обновлены в " + dataTime.ToString("HH:mm");
        ApplySearch();
    }

    void SetDataNote(string note) {
        dataNote = note;
        UpdateStatus();
    }

    void SetBusy(bool busy, string text) {
        loading = busy;
        foreach (Control c in bottom.Controls) if (c is Button) c.Enabled = !busy;
        Cursor = busy ? Cursors.AppStarting : Cursors.Default;
        if (text != null) { status.Text = text; status.Refresh(); }
    }

    void BuildColumns() {
        grid.RowCount = 0;
        grid.Columns.Clear();
        foreach (int f in query.Columns) {
            DbfField field = table.Fields[f];
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = query.ColumnTitle(f);
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            string n = field.Name.ToUpperInvariant();
            c.FillWeight = n == "FIO" ? 300 : n == "UL" ? 200 : (n == "NAB" || n == "NCH" || n == "TIPCH") ? 130 : 70;
            grid.Columns.Add(c);
        }
    }

    internal void ApplySearch() {
        if (loading || table == null) return;
        found = query.Find(queryBox.Text, mode);
        logTimer.Stop();
        logTimer.Start();
        grid.RowCount = 0;
        grid.RowCount = found.Count;
        grid.Invalidate();
        UpdateStatus();
    }

    void UpdateStatus() {
        if (table == null || query == null) return;
        string where;
        if (allowed != null) where = "абонентов в моих сельсоветах " + table.Count + " (всего в базе " + totalInBase + ")";
        else if (query.Scope != null) where = "абонентов в моих сельсоветах " + query.Scope.Length + " из " + table.Count;
        else where = tableName + ", " + table.Count + " абонентов";
        if (directory != null) where += ", улицы - из " + directory.TableName;
        where += "; " + dataNote + mineNote;
        string warning = streetCodes && directory == null && mode == SearchMode.Address
            ? "   В поле улицы - коды, а справочник улиц не нашёлся: адрес ищется по дому и квартире, название улицы может не находиться."
            : "";
        if (DbfTable.SplitWords(queryBox.Text).Length == 0)
            status.Text = "Впишите " + Prompts[(int)mode] + ".   База: " + where + warning;
        else if (found.Count == 0)
            status.Text = "Ничего не найдено.   База: " + where + warning;
        else
            status.Text = "Найдено: " + found.Count + ". Двойной щелчок или Enter - карточка.   База: " + where;
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        if (table == null || e.RowIndex >= found.Count || e.ColumnIndex >= query.Columns.Length) return;
        e.Value = query.CellText(found[e.RowIndex], query.Columns[e.ColumnIndex]);
    }

    void OnQueryKey(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) {
            e.Handled = true;
            e.SuppressKeyPress = true;
            timer.Stop();
            ApplySearch();
            LogSearch();
            if (found.Count == 1) OpenCard(0);   // нашёлся один - сразу карточка
            else if (found.Count > 1) grid.Focus();
        } else if (e.KeyCode == Keys.Down && found.Count > 0) {
            e.Handled = true;
            grid.Focus();
        }
    }

    void OnGridKey(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter && grid.CurrentCell != null) {
            e.Handled = true;
            e.SuppressKeyPress = true;
            OpenCard(grid.CurrentCell.RowIndex);
        }
    }

    internal void OpenCard(int index) {
        if (table == null || index < 0 || index >= found.Count) return;
        Card.Show(this, table, found[index], query.CardExtra(found[index]));
    }

    string SelsovetyFile {
        get { return Path.Combine(Path.GetDirectoryName(Path.GetFullPath(settingsFile)), "selsovety.txt"); }
    }

    // список сельсоветов: из файла selsovety.txt (если его нет - создаём с вашими шестью)
    string[] ReadSelsovety() {
        string file = SelsovetyFile;
        if (!File.Exists(file)) SaveSelsovety(DefaultSelsovety);
        List<string> names = new List<string>();
        try {
            foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                string n = line.Trim();
                if (n.Length > 0 && !n.StartsWith("#")) names.Add(n);
            }
        } catch (Exception) {
            names.AddRange(DefaultSelsovety);
        }
        return names.ToArray();
    }

    void SaveSelsovety(string[] names) {
        List<string> lines = new List<string>();
        lines.Add("# Сельсоветы, абоненты которых нужны в поиске - по одному в строке");
        lines.AddRange(names);
        Places.SaveSettings(SelsovetyFile, lines.ToArray());
    }

    void FindMine() {
        string[] names = ReadSelsovety();
        mine = null;
        mineNote = "";
        if (names.Length == 0) {
            Report.Line("Сельсоветы: список пуст - ищу по всей базе");
        } else {
            string report;
            int[] rows = query.FindByNames(names, out report);
            Report.Line(report);
            if (rows.Length > 0) {
                mine = rows;
            } else {
                mineNote = "   Названия сельсоветов в базе не нашлись - показана вся база. Нажмите \"Отчёт\" и пришлите его.";
                Report.Line("    названия сельсоветов не нашлись ни в справочнике улиц, ни в полях абонента - отбор не применён");
            }
        }
        onlyMine.Enabled = mine != null;
        onlyMine.Text = mine == null ? "Только мои сельсоветы (не нашлись)" : "Только мои сельсоветы (" + names.Length + ")";
        ApplyScope();
    }

    void ApplyScope() {
        if (query == null) return;
        query.Scope = onlyMine.Checked ? mine : null;
    }

    void OnOnlyMineChanged() {
        if (suppressMine || table == null) return;
        if (!onlyMine.Checked && allowed != null) {
            string q = "Без отбора программа прочитает всю базу (около " + totalInBase + " записей) - это может занять несколько минут "
                + "и много памяти. Продолжить?";
            if (MessageBox.Show(this, q, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
                suppressMine = true;
                onlyMine.Checked = true;
                suppressMine = false;
                return;
            }
        }
        if (allowed == null && directory == null && mine != null) { ApplyScope(); ApplySearch(); return; }   // отбор по тексту - без перечитывания
        LoadTable();
    }

    void EditSelsovety() {
        Form d = new Form();
        d.Text = "Мои сельсоветы";
        d.Size = new Size(360, 380);
        d.StartPosition = FormStartPosition.CenterParent;
        d.MinimizeBox = false;
        d.MaximizeBox = false;
        Label l = new Label();
        l.Text = "По одному сельсовету в строке:";
        l.AutoSize = true;
        l.Location = new Point(12, 10);
        TextBox box = new TextBox();
        box.Multiline = true;
        box.AcceptsReturn = true;
        box.ScrollBars = ScrollBars.Vertical;
        box.Location = new Point(12, 30);
        box.Size = new Size(320, 260);
        box.Text = string.Join("\r\n", ReadSelsovety());
        Button ok = new Button();
        ok.Text = "OK";
        ok.Location = new Point(170, 300);
        ok.DialogResult = DialogResult.OK;
        Button cancel = new Button();
        cancel.Text = "Отмена";
        cancel.Location = new Point(255, 300);
        cancel.DialogResult = DialogResult.Cancel;
        d.CancelButton = cancel;
        d.Controls.AddRange(new Control[] { l, box, ok, cancel });
        if (d.ShowDialog(this) != DialogResult.OK) return;
        List<string> names = new List<string>();
        foreach (string line in box.Text.Replace("\r\n", "\n").Split('\n')) if (line.Trim().Length > 0) names.Add(line.Trim());
        SaveSelsovety(names.ToArray());
        if (table == null) return;
        suppressMine = true;
        onlyMine.Checked = true;
        suppressMine = false;
        LoadTable();
    }

    void OpenBatch() {
        if (table == null || loading) {
            MessageBox.Show(this, "Подождите, пока загрузится база абонентов.", Text);
            return;
        }
        new BatchSearch(table, query, mode).Show(this);
    }

    // Выбор папки и таблицы абонентов, если сама не нашлась или нашлась не та
    void ChooseBase() {
        Form d = new Form();
        d.Text = "Где искать абонентов";
        d.Size = new Size(560, 420);
        d.StartPosition = FormStartPosition.CenterParent;
        d.MinimizeBox = false;
        d.MaximizeBox = false;

        TextBox folderBox = new TextBox();
        folderBox.Text = folder;
        folderBox.Location = new Point(12, 30);
        folderBox.Width = 330;
        Button browse = new Button();
        browse.Text = "Обзор...";
        browse.Location = new Point(350, 28);
        Button show = new Button();
        show.Text = "Показать";
        show.Location = new Point(430, 28);
        Label l1 = new Label();
        l1.Text = "Папка с таблицами базы:";
        l1.AutoSize = true;
        l1.Location = new Point(12, 10);
        Label l2 = new Label();
        l2.Text = "Таблица абонентов (подходящие - вверху):";
        l2.AutoSize = true;
        l2.Location = new Point(12, 62);
        ListBox list = new ListBox();
        list.Location = new Point(12, 82);
        list.Size = new Size(520, 250);
        Button ok = new Button();
        ok.Text = "OK";
        ok.Location = new Point(370, 345);
        ok.DialogResult = DialogResult.OK;
        Button cancel = new Button();
        cancel.Text = "Отмена";
        cancel.Location = new Point(455, 345);
        cancel.DialogResult = DialogResult.Cancel;
        d.AcceptButton = ok;
        d.CancelButton = cancel;
        d.Controls.AddRange(new Control[] { l1, folderBox, browse, show, l2, list, ok, cancel });

        EventHandler fill = delegate {
            list.Items.Clear();
            d.Cursor = Cursors.WaitCursor;
            try {
                foreach (TableInfo info in TableInfo.List(folderBox.Text.Trim(), null)) {
                    list.Items.Add(info);
                    if (string.Compare(info.Name, tableName, StringComparison.OrdinalIgnoreCase) == 0) list.SelectedItem = info;
                }
                if (list.SelectedIndex < 0 && list.Items.Count > 0) list.SelectedIndex = 0;
            } catch (Exception ex) {
                MessageBox.Show(d, "Не удалось открыть папку:\n" + ex.Message, d.Text);
            } finally {
                d.Cursor = Cursors.Default;
            }
        };
        show.Click += fill;
        browse.Click += delegate {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (Directory.Exists(folderBox.Text)) fb.SelectedPath = folderBox.Text;
            if (fb.ShowDialog(d) == DialogResult.OK) { folderBox.Text = fb.SelectedPath; fill(null, EventArgs.Empty); }
        };
        list.DoubleClick += delegate { if (list.SelectedItem != null) d.DialogResult = DialogResult.OK; };
        d.Shown += delegate { fill(null, EventArgs.Empty); };

        if (d.ShowDialog(this) != DialogResult.OK || list.SelectedItem == null) return;
        folder = folderBox.Text.Trim();
        tableName = ((TableInfo)list.SelectedItem).Name;
        LoadTable();
    }
}

// Выгрузка таблицы строк в Excel ("Таблица XML 2003": все значения - текст, нули в начале не теряются)
public static class Spreadsheet {
    public static void Open(string name, string[] headers, List<string[]> rows) {
        string file = Path.Combine(Path.GetTempPath(), name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xml");
        using (StreamWriter w = new StreamWriter(file, false, new UTF8Encoding(false))) {
            w.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            w.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            w.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            w.WriteLine("<Styles><Style ss:ID=\"h\"><Font ss:Bold=\"1\"/></Style></Styles>");
            w.WriteLine("<Worksheet ss:Name=\"" + DbfTable.Xml(name) + "\"><Table>");
            w.Write("<Row>");
            foreach (string h in headers) w.Write("<Cell ss:StyleID=\"h\"><Data ss:Type=\"String\">" + DbfTable.Xml(h) + "</Data></Cell>");
            w.WriteLine("</Row>");
            foreach (string[] row in rows) {
                w.Write("<Row>");
                foreach (string v in row)
                    w.Write(v == null || v.Length == 0 ? "<Cell/>" : "<Cell><Data ss:Type=\"String\">" + DbfTable.Xml(v) + "</Data></Cell>");
                w.WriteLine("</Row>");
            }
            w.WriteLine("</Table></Worksheet></Workbook>");
        }
        try { Process.Start("excel.exe", "\"" + file + "\""); }
        catch { Process.Start(file); }
    }
}

// Поиск по списку: вставляете столбик лицевых, счётчиков, ФИО или адресов - получаете итог по каждой строке
public class BatchSearch : Form {
    readonly DbfTable table;
    readonly AbonentQuery query;
    SearchMode mode;
    readonly RadioButton[] modeButtons = new RadioButton[4];
    internal readonly TextBox input = new TextBox();
    readonly DataGridView grid = new DataGridView();
    internal readonly Label status = new Label();
    readonly Button findButton = new Button();
    readonly List<Item> items = new List<Item>();
    bool running;

    const int MaxPerLine = 20;   // если по строке нашлось много - показываем первые

    class Item {
        public int Line;
        public string Text;
        public string Result;
        public int Row = -1;      // запись абонента или -1, если не найден
        public int Kind;          // 0 - найден один, 1 - несколько, 2 - не найден
    }

    static readonly string[] ModeNames = { "&Лицевые счета", "&Адреса", "Номера &счётчиков", "&ФИО" };

    public BatchSearch(DbfTable table, AbonentQuery query, SearchMode mode) {
        this.table = table;
        this.query = query;
        this.mode = mode;
        Text = "Поиск по списку";
        Size = new Size(1000, 700);
        StartPosition = FormStartPosition.CenterScreen;

        SplitContainer split = new SplitContainer();
        split.Dock = DockStyle.Fill;
        split.Orientation = Orientation.Horizontal;

        Panel head = new Panel();
        head.Dock = DockStyle.Top;
        head.Height = 58;
        Label l1 = new Label();
        l1.Text = "В списке:";
        l1.AutoSize = true;
        l1.Location = new Point(10, 10);
        head.Controls.Add(l1);
        int x = 80;
        for (int i = 0; i < modeButtons.Length; i++) {
            RadioButton rb = new RadioButton();
            rb.Text = ModeNames[i];
            rb.AutoSize = true;
            rb.Location = new Point(x, 8);
            rb.Tag = (SearchMode)i;
            rb.Checked = i == (int)mode;
            rb.CheckedChanged += delegate(object s, EventArgs e) { if (((RadioButton)s).Checked) this.mode = (SearchMode)((RadioButton)s).Tag; };
            modeButtons[i] = rb;
            head.Controls.Add(rb);
            x += 150;
        }
        Label l2 = new Label();
        l2.Text = "Вставьте список - по одному значению в строке (можно столбик из Excel), и нажмите \"Найти\":";
        l2.AutoSize = true;
        l2.Location = new Point(10, 36);
        head.Controls.Add(l2);

        input.Multiline = true;
        input.AcceptsReturn = true;
        input.ScrollBars = ScrollBars.Both;
        input.WordWrap = false;
        input.Dock = DockStyle.Fill;
        input.Font = new Font(FontFamily.GenericMonospace, 10f);

        Panel buttons = new Panel();
        buttons.Dock = DockStyle.Right;
        buttons.Width = 130;
        findButton.Text = "Найти";
        findButton.Location = new Point(10, 6);
        findButton.Size = new Size(110, 32);
        findButton.Click += delegate { Run(); };
        Button clear = new Button();
        clear.Text = "Очистить";
        clear.Location = new Point(10, 44);
        clear.Size = new Size(110, 28);
        clear.Click += delegate { input.Clear(); input.Focus(); };
        buttons.Controls.Add(findButton);
        buttons.Controls.Add(clear);

        split.Panel1.Controls.Add(input);
        split.Panel1.Controls.Add(buttons);
        split.Panel1.Controls.Add(head);

        grid.Dock = DockStyle.Fill;
        grid.VirtualMode = true;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.CellValueNeeded += new DataGridViewCellValueEventHandler(OnCellValueNeeded);
        grid.CellFormatting += new DataGridViewCellFormattingEventHandler(OnCellFormatting);
        grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e) { OpenCard(e.RowIndex); };
        AddColumn("№", 35);
        AddColumn("Искали", 200);
        AddColumn("Результат", 110);
        foreach (int f in query.Columns) AddColumn(query.ColumnTitle(f), 100);

        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 34;
        status.Dock = DockStyle.Fill;
        status.TextAlign = ContentAlignment.MiddleLeft;
        status.Text = "Найденные - белым, несколько абонентов на одну строку - жёлтым, не найденные - красным. Двойной щелчок - карточка.";
        Button excel = new Button();
        excel.Text = "В Excel";
        excel.Dock = DockStyle.Right;
        excel.Width = 120;
        excel.Click += delegate { Export(); };
        bottom.Controls.Add(status);
        bottom.Controls.Add(excel);

        split.Panel2.Controls.Add(grid);
        split.Panel2.Controls.Add(bottom);
        Controls.Add(split);
        Shown += delegate { split.SplitterDistance = 230; input.Focus(); };
    }

    void AddColumn(string title, int weight) {
        DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
        c.HeaderText = title;
        c.FillWeight = weight;
        c.SortMode = DataGridViewColumnSortMode.NotSortable;
        grid.Columns.Add(c);
    }

    internal void Run() {
        if (running) return;
        List<string> lines = new List<string>();
        foreach (string raw in input.Text.Replace("\r\n", "\n").Split('\n')) {
            string line = raw.Replace('\t', ' ').Trim();
            if (line.Length > 0) lines.Add(line);
        }
        if (lines.Count == 0) { MessageBox.Show(this, "Вставьте список: по одному значению в строке.", Text); return; }
        running = true;
        findButton.Enabled = false;
        Cursor = Cursors.WaitCursor;
        items.Clear();
        grid.RowCount = 0;
        DateTime started = DateTime.Now;
        int single = 0, several = 0, none = 0;
        try {
            for (int i = 0; i < lines.Count; i++) {
                if (i % 5 == 0) {
                    status.Text = "Ищу: строка " + (i + 1) + " из " + lines.Count + "...";
                    Application.DoEvents();
                    if (IsDisposed) return;
                }
                List<int> found = query.Find(lines[i], mode);
                if (found.Count == 0) {
                    Add(i + 1, lines[i], "не найден", -1, 2);
                    none++;
                    continue;
                }
                string result = found.Count == 1 ? "найден" : "найдено " + found.Count + (found.Count > MaxPerLine ? " (показаны " + MaxPerLine + ")" : "");
                if (found.Count == 1) single++; else several++;
                for (int k = 0; k < found.Count && k < MaxPerLine; k++) Add(i + 1, lines[i], result, found[k], found.Count == 1 ? 0 : 1);
            }
        } finally {
            if (!IsDisposed) {
                running = false;
                findButton.Enabled = true;
                Cursor = Cursors.Default;
            }
        }
        grid.RowCount = items.Count;
        grid.Invalidate();
        status.Text = "Строк в списке: " + lines.Count + ". Найдено по одному абоненту: " + single
            + ", по несколько: " + several + ", не найдено: " + none + ". Двойной щелчок - карточка.";
        Dictionary<int, int> lengths = new Dictionary<int, int>();
        foreach (string line in lines) {
            int digits = 0;
            foreach (char c in line) if (c >= '0' && c <= '9') digits++;
            int n;
            lengths.TryGetValue(digits, out n);
            lengths[digits] = n + 1;
        }
        List<string> shape = new List<string>();
        foreach (KeyValuePair<int, int> kv in lengths) shape.Add(kv.Key + " цифр - " + kv.Value);
        Report.Search("Поиск по списку, в списке: " + ModeNames[(int)mode].Replace("&", "").ToLowerInvariant() + "; строк " + lines.Count
            + ", найдено по одному " + single + ", по несколько " + several + ", не найдено " + none
            + ", время " + (DateTime.Now - started).TotalSeconds.ToString("0.0") + " с; цифр в строках: " + string.Join(", ", shape.ToArray()));
    }

    void Add(int line, string text, string result, int row, int kind) {
        Item it = new Item();
        it.Line = line;
        it.Text = text;
        it.Result = result;
        it.Row = row;
        it.Kind = kind;
        items.Add(it);
    }

    string Cell(Item it, int column) {
        if (column == 0) return it.Line.ToString();
        if (column == 1) return it.Text;
        if (column == 2) return it.Result;
        if (it.Row < 0) return "";
        return query.CellText(it.Row, query.Columns[column - 3]);
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        if (e.RowIndex < items.Count) e.Value = Cell(items[e.RowIndex], e.ColumnIndex);
    }

    void OnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
        if (e.RowIndex < 0 || e.RowIndex >= items.Count) return;
        int kind = items[e.RowIndex].Kind;
        if (kind == 2) e.CellStyle.BackColor = Color.FromArgb(255, 220, 220);
        else if (kind == 1) e.CellStyle.BackColor = Color.FromArgb(255, 245, 200);
    }

    void OpenCard(int index) {
        if (index < 0 || index >= items.Count || items[index].Row < 0) return;
        Card.Show(this, table, items[index].Row, query.CardExtra(items[index].Row));
    }

    void Export() {
        if (items.Count == 0) { MessageBox.Show(this, "Сначала найдите что-нибудь.", Text); return; }
        string[] headers = new string[grid.Columns.Count];
        for (int c = 0; c < headers.Length; c++) headers[c] = grid.Columns[c].HeaderText;
        List<string[]> rows = new List<string[]>();
        foreach (Item it in items) {
            string[] r = new string[headers.Length];
            for (int c = 0; c < r.Length; c++) r[c] = Cell(it, c);
            rows.Add(r);
        }
        try { Spreadsheet.Open("Поиск_по_списку", headers, rows); }
        catch (Exception ex) { MessageBox.Show(this, "Не удалось выгрузить в Excel:\n" + ex.Message, Text); }
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
    [AbonentSearch]::Run($here)
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("$_", 'Просмотр базы - ошибка')
}
