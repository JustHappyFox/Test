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
            DecodeLeaf(p, n, keyLen, tag.KeyType == 'C' ? (byte)' ' : (byte)0, node);
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
        List<IndexTag> tags = new List<IndexTag>();
        using (IndexReader r = new IndexReader(path)) {
            bool cdx = string.Compare(Path.GetExtension(path), ".cdx", StringComparison.OrdinalIgnoreCase) == 0;
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
    public string Note = "";    // пояснение для строки состояния
    public string How = "";     // как искали - для отчёта
    public double Seconds;
}

public delegate bool CancelCheck();

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

    public const int FetchLimit = 300;    // больше записей за один поиск не показываем (адрес - до 5000: потом отбор по дому)
    const int SeekLimit = 5000;

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
        foreach (string f in Directory.GetFiles(folder)) {
            string ext = Path.GetExtension(f).ToLowerInvariant();
            if (ext == ".cdx" || ext == ".idx") files.Add(f);
        }
        files.Sort(delegate(string a, string b) {
            bool oa = string.Compare(Path.GetFileNameWithoutExtension(a), own, StringComparison.OrdinalIgnoreCase) == 0;
            bool ob = string.Compare(Path.GetFileNameWithoutExtension(b), own, StringComparison.OrdinalIgnoreCase) == 0;
            if (oa != ob) return oa ? -1 : 1;
            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        });
        foreach (string f in files) {
            IndexFiles.Add(Path.GetFileName(f));
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
            bool[] variants = t.Key.UsesCase() ? new bool[] { true, false } : new bool[] { true };
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
                    byte[] key = FoxExpr.KeyBytes(t.Key.Eval(c), t.KeyLen);
                    if (t.KeyType == 'C' && IsBlank(key)) continue;   // пустой ключ есть у многих - не показательно
                    bool more;
                    List<int> found = r.Seek(t, key, SeekLimit, out more);
                    if (more && !found.Contains(recNo)) continue;      // слишком много одинаковых - не показательно
                    tried++;
                    if (found.Contains(recNo)) ok++;
                    else if (sample.Length == 0) sample = "; ключ записи: " + Report.Shape(Header.Enc.GetString(key), false)
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

    public SearchResult Search(string text, SearchMode mode, Dictionary<string, bool> selsovety) {
        return Search(text, mode, selsovety, FetchLimit);
    }

    // Ошибка чтения бывает, если программа базы как раз дописывает индекс - тогда пробуем ещё раз
    public SearchResult Search(string text, SearchMode mode, Dictionary<string, bool> selsovety, int limit) {
        SearchResult r = SearchOnce(text, mode, selsovety, limit);
        if (r.Error != null) {
            System.Threading.Thread.Sleep(300);
            SearchResult again = SearchOnce(text, mode, selsovety, limit);
            again.How += "; (со второй попытки, первая: " + r.Error + ")";
            return again;
        }
        return r;
    }

    SearchResult SearchOnce(string text, SearchMode mode, Dictionary<string, bool> selsovety, int limit) {
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        Seeker s = new Seeker();
        try {
            bool ok;
            switch (mode) {
                case SearchMode.Account: ok = SeekAccount(text, s, res); break;
                case SearchMode.Meter: ok = SeekMeter(text, s, res); break;
                case SearchMode.Name: ok = SeekName(text, s, res); break;
                default: ok = SeekAddress(text, s, res, selsovety); break;
            }
            if (!ok) { res.NoIndex = res.Error == null && res.Note.Length == 0; return Finish(res, null, selsovety, started); }
            List<int> recNos = new List<int>(s.Found.Keys);
            if (mode == SearchMode.Address) limit = Math.Max(limit, SeekLimit);
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
    public SearchResult Scan(string text, SearchMode mode, Dictionary<string, bool> selsovety, LoadProgress progress) {
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Scanned = true;
        DbfTable part;
        if (mode == SearchMode.Address && Dir != null && Header.FieldIndex("UL") >= 0) {
            AddressParts a = AddressParts.Parse(text);
            Dictionary<string, bool> codes = new Dictionary<string, bool>();
            int best;
            foreach (string c in Dir.BestStreets(a.Words, null, out best)) codes[c] = true;
            DbfField ul = Header.Fields[Header.FieldIndex("UL")];
            part = DbfTable.Load(FilePath, null, delegate(byte[] rec, int start) {
                return codes.ContainsKey(StreetDirectory.NormalizeCode(Encoding.ASCII.GetString(rec, start + ul.Offset, ul.Length)));
            }, progress, 20000);
        } else {
            string words = mode == SearchMode.Account ? Digits(text) : text;
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
        if (part.Count > 0) {
            List<int> rows = res.Query.Find(res.Text, res.Mode);
            int ul = part.FieldIndex("UL");
            foreach (int row in rows) {
                if (Dir != null && ul >= 0 && selsovety != null
                    && !Dir.InSelsovety(DbfTable.Format(part.GetValue(row, ul, null), true), selsovety)) { res.Hidden++; continue; }
                res.Rows.Add(row);
            }
        }
        res.Seconds = (DateTime.Now - started).TotalSeconds;
        return res;
    }

    static string Digits(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s) if (c >= '0' && c <= '9') sb.Append(c);
        return sb.ToString();
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

    // Лицевой: точный номер в нескольких видах (с нулями впереди и без), потом - номера, начинающиеся так
    bool SeekAccount(string text, Seeker s, SearchResult res) {
        string digits = Digits(text);
        if (digits.Length == 0) { res.Note = "Впишите номер."; return false; }
        string core = digits.TrimStart('0');
        if (core.Length == 0) core = "0";
        bool any = false;
        foreach (string field in AccountFields) {
            IndexTag t = TagFor(field);
            if (t == null) continue;
            any = true;
            Style st = StyleOf(field);
            int f = Header.FieldIndex(field);
            int width = Header.Fields[f].Length;
            List<string> exact = new List<string>();
            AddVariant(exact, digits);
            AddVariant(exact, core);
            if (st.Digits > core.Length) AddVariant(exact, core.PadLeft(st.Digits, '0'));
            if (Header.Fields[f].Type == 'C' && width > core.Length) AddVariant(exact, core.PadLeft(width, '0'));
            foreach (string v in exact) {
                s.Seek(t, KeyFor(t, new string[] { v }, false));
                if (Header.Fields[f].Type == 'C') s.Seek(t, KeyFor(t, new string[] { v.PadLeft(width) }, false));
            }
            if (s.Found.Count > 0) { res.How = "точный номер (поле " + field + ")"; return true; }
        }
        if (!any) return false;
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

    // Счётчик: точный номер, потом начало номера; если не нашлось - то же по снятому счётчику (NSC)
    bool SeekMeter(string text, Seeker s, SearchResult res) {
        string typed = text.Trim();
        if (typed.Length == 0) { res.Note = "Впишите номер."; return false; }
        string digits = Digits(typed);
        bool any = false;
        foreach (string field in new string[] { "NCH", "NSC" }) {
            IndexTag t = TagFor(field);
            if (t == null) continue;
            any = true;
            Style st = StyleOf(field);
            List<string> variants = new List<string>();
            AddVariant(variants, typed);
            AddVariant(variants, typed.ToUpperInvariant());
            if (digits.Length >= 4) { AddVariant(variants, digits); AddVariant(variants, digits.TrimStart('0')); }
            foreach (string v in variants) s.Seek(t, KeyFor(t, new string[] { Aligned(field, v, st) }, false));
            if (s.Found.Count == 0 && !st.Right) foreach (string v in variants) if (v.Length >= 3) s.Seek(t, KeyFor(t, new string[] { v }, true));
            if (s.Found.Count > 0) { res.How = field == "NSC" ? "по снятому счётчику (NSC)" : "поле " + field; return true; }
        }
        return any;
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

    // ФИО: начало ФИО целиком ("иванов иван"), если не нашлось - по фамилии
    bool SeekName(string text, Seeker s, SearchResult res) {
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
            foreach (string v in variants) s.Seek(t, KeyFor(t, new string[] { v }, true));
            if (s.Found.Count > 0) { res.How = n == words.Length ? "начало ФИО" : "по фамилии"; return true; }
        }
        return true;
    }

    // Адрес: по справочнику улиц находим улицы, подходящие по словам (населённый пункт, улица),
    // и по индексу берём абонентов этих улиц (сразу с домом, если индекс по улице и дому)
    bool SeekAddress(string text, Seeker s, SearchResult res, Dictionary<string, bool> selsovety) {
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

// Адрес, разобранный на части: слова (населённый пункт, улица), дом, корпус, квартира
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
    public Dictionary<string, string> Links = new Dictionary<string, string>();   // поле-ссылка -> таблица (IDNP -> SPRNP.DBF)
    readonly Dictionary<string, string> names = new Dictionary<string, string>();
    readonly Dictionary<string, string> lowerNames = new Dictionary<string, string>();
    readonly Dictionary<string, string> ssOf = new Dictionary<string, string>();  // код улицы -> сельсовет

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
public class AbonentSearch : Form {
    readonly string settingsFile;
    readonly string viewerSettings;
    string folder;
    string tableName;
    LiveBase live;                       // открытая база (null - ещё не открыта)
    bool opening;
    SearchResult current;
    bool searching, again;               // идёт поиск; пока он шёл, запрос поменялся
    bool scanning;
    volatile bool cancelScan;
    bool cardWhenDone;                   // Enter: нашёлся один - сразу карточка
    Dictionary<string, bool> selected;   // выбранные сельсоветы (null - все)
    string baseInfo = "";

    static readonly string[] DefaultSelsovety = { "Шершунский", "Папернянский", "Юзуфовский", "Острошицкий", "Боровлянский", "Колодищанский" };
    internal SearchMode mode = SearchMode.Account;
    internal readonly RadioButton[] modeButtons = new RadioButton[4];
    readonly Label hint = new Label();
    readonly Button ssButton = new Button();
    readonly Button scanButton = new Button();

    static readonly string[] ModeNames = { "&Лицевой счёт", "&Адрес", "№ &счётчика", "&ФИО" };
    static readonly string[] Prompts = { "лицевой счёт", "адрес", "номер счётчика", "фамилию" };
    static readonly string[] Hints = {
        "Например:  044060216   (нули впереди можно не писать; можно начало номера)",
        "Вставьте адрес целиком:  Острошицкий Городок Заводская ул. 18   или коротко:  заводская 18 кв 5",
        "Номер счётчика целиком или его начало",
        "С фамилии:  иванов   или   иванов иван" };

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
        Button listButton = new Button();
        listButton.Text = "Поиск по списку...";
        listButton.AutoSize = true;
        listButton.Margin = new Padding(12, 3, 3, 0);
        listButton.Click += delegate { OpenBatch(); };
        row.Controls.Add(listButton);
        ssButton.Text = "Сельсоветы ▼";
        ssButton.AutoSize = true;
        ssButton.Margin = new Padding(6, 3, 3, 0);
        ssButton.Click += delegate { ShowSelsovety(); };
        row.Controls.Add(ssButton);
        top.Controls.Add(row);
        Load += delegate { Width = Math.Max(Width, row.PreferredSize.Width + 60); };

        queryBox.Font = new Font(Font.FontFamily, 14f);
        queryBox.Location = new Point(12, 38);
        queryBox.Width = top.Width - 24;
        queryBox.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        queryBox.TextChanged += new EventHandler(OnQueryChanged);
        queryBox.KeyDown += new KeyEventHandler(OnQueryKey);
        hint.ForeColor = SystemColors.GrayText;
        hint.AutoSize = true;
        hint.Location = new Point(12, 76);
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
        scanButton.Text = "Искать перебором";
        scanButton.Dock = DockStyle.Right;
        scanButton.Width = 150;
        scanButton.Visible = false;
        scanButton.Click += delegate { if (scanning) cancelScan = true; else StartScan(); };
        bottom.Controls.Add(scanButton);

        timer.Interval = 400;
        timer.Tick += delegate { timer.Stop(); StartSearch(); };
        logTimer.Interval = 2000;
        logTimer.Tick += delegate { logTimer.Stop(); LogSearch(); };

        Controls.Add(grid);
        Controls.Add(bottom);
        Controls.Add(top);

        Shown += delegate { Start(); };
        FormClosing += delegate { cancelScan = true; };
    }

    Button BottomButton(string text, EventHandler click) {
        Button b = new Button();
        b.Text = text;
        b.Dock = DockStyle.Right;
        b.Width = 120;
        b.Click += click;
        return b;
    }

    // из фонового потока - в окно
    void Post(MethodInvoker m) {
        try { if (!IsDisposed) BeginInvoke(m); } catch (Exception) { }
    }

    void PostStatus(string text) {
        Post(delegate { status.Text = text; });
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
        StartSearch();
    }

    void SaveSettings() {
        if (folder == null) return;
        StreetDirectory d = live == null ? null : live.Dir;
        List<string> links = new List<string>();
        if (d != null) foreach (KeyValuePair<string, string> kv in d.Links) links.Add(kv.Key + "=" + kv.Value);
        Places.SaveSettings(settingsFile, new string[] {
            folder, tableName == null ? "" : tableName, ((int)mode).ToString(),
            d == null ? "" : d.TableName, d == null ? "" : d.KeyField, string.Join(";", links.ToArray()) });
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
        SearchResult r = current;
        if (r == null) return;
        string line = "Поиск по: " + Prompts[(int)r.Mode] + (r.Scanned ? " (перебором)" : "") + "; "
            + (r.NoIndex ? "нет подходящего индекса" : r.How) + (r.Query != null && r.Query.LastExplain.Length > 0 ? "; " + r.Query.LastExplain : "")
            + (r.Note.Length > 0 ? "; " + r.Note : "") + " -> найдено " + r.Rows.Count + (r.Hidden > 0 ? ", скрыто (другие сельсоветы) " + r.Hidden : "")
            + (r.TooMany ? ", показана часть" : "") + ", " + r.Seconds.ToString("0.00") + " с";
        if (line == lastLogged) return;
        lastLogged = line;
        Report.Search(line);
    }

    string Dir { get { return Path.GetDirectoryName(Path.GetFullPath(settingsFile)); } }

    void Start() {
        Report.Start(Path.Combine(Dir, "otchet.txt"));
        // выборка абонентов прошлой версии больше не нужна (в ней были данные абонентов)
        foreach (string old in new string[] { "kesh_abonenty.dbf", "kesh_abonenty.txt", "kesh_abonenty.dbf.tmp" }) {
            try { if (File.Exists(Path.Combine(Dir, old))) File.Delete(Path.Combine(Dir, old)); } catch (Exception) { }
        }
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
        OpenBase();
    }

    // ---- открытие базы: заголовок, индексы и справочники; самих абонентов не читаем ----

    internal void OpenBase() {
        if (opening || tableName == null) return;
        opening = true;
        live = null;
        current = null;
        grid.RowCount = 0;
        string path = Path.Combine(folder, tableName);
        string dirTable = Places.ReadSetting(settingsFile, 3), dirKey = Places.ReadSetting(settingsFile, 4);
        Dictionary<string, string> links = new Dictionary<string, string>();
        string linkLine = Places.ReadSetting(settingsFile, 5);
        if (linkLine != null)
            foreach (string pair in linkLine.Split(';')) {
                int eq = pair.IndexOf('=');
                if (eq > 0) links[pair.Substring(0, eq).Trim().ToUpperInvariant()] = pair.Substring(eq + 1).Trim();
            }
        string table = tableName, dirFolder = folder;
        SetBusy(true, "Открываю базу " + tableName + "...");
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            LiveBase b = null;
            string error = null;
            Dictionary<string, bool> sel = null;
            DateTime t0 = DateTime.Now;
            try {
                b = LiveBase.Open(path);
                double opened = (DateTime.Now - t0).TotalSeconds;
                ReportBase(b);
                PostStatus("Читаю справочники улиц...");
                DateTime t1 = DateTime.Now;
                b.Dir = PrepareDirectory(b, dirFolder, table, dirTable, dirKey, links);
                Report.Line("Время открытия: таблица и индексы - " + opened.ToString("0.0") + " с, справочники - "
                    + (DateTime.Now - t1).TotalSeconds.ToString("0.0") + " с");
                sel = LoadSelection(b);
            } catch (Exception ex) {
                error = ex.Message;
                Report.Line("ОШИБКА открытия " + path + ": " + ex.Message);
            }
            double seconds = (DateTime.Now - t0).TotalSeconds;
            Post(delegate { AfterOpen(b, sel, error, seconds); });
            if (error != null) return;
            // заранее проверяем индексы по нужным полям - итог в отчёт
            try {
                b.VerifyMain();
                Report.Line("Проверка индексов (ищем по индексу записи из разных мест таблицы):");
                foreach (IndexTag t in b.Tags)
                    if (t.State != 0 || t.Check.Length > 0) Report.Line("    " + t.Label + ": " + t.Check);
                Report.Line("Чем ищется: лицевой - " + TagName(b, "NAB", false) + ", Л/С ЕРИП - " + TagName(b, "LS_ERIP", false)
                    + ", счётчик - " + TagName(b, "NCH", false) + ", снятый счётчик - " + TagName(b, "NSC", false)
                    + ", ФИО - " + TagName(b, "FIO", false) + ", адрес - " + TagName(b, "UL", true));
                Report.Line("--- поиски (сами слова не записываются) ---");
            } catch (Exception ex) {
                Report.Line("Проверка индексов прервалась: " + ex.Message);
            }
        });
        th.IsBackground = true;
        th.Start();
    }

    static string TagName(LiveBase b, string field, bool preferLong) {
        if (b.Header.FieldIndex(field) < 0) return "нет поля";
        IndexTag t = b.TagFor(field, preferLong);
        return t == null ? "нет индекса (только перебором)" : t.Label;
    }

    void ReportBase(LiveBase b) {
        long size = 0;
        try { size = new FileInfo(b.FilePath).Length; } catch (Exception) { }
        Report.Line("Таблица абонентов " + Path.GetFileName(b.FilePath) + ": записей по заголовку " + b.Header.Expected + ", длина записи "
            + b.Header.RecordLength + " байт, файл " + (size / 1048576.0).ToString("0") + " МБ");
        Report.Line("Поля (по первым " + b.Sample.Count + " записям):");
        Report.Fields(b.Sample, "    ");
        List<string> files = b.IndexFiles;
        Report.Line("Индексных файлов (.CDX, .IDX) в папке: " + files.Count + (files.Count > 0 ? ": " + string.Join(", ", files.GetRange(0, Math.Min(files.Count, 80)).ToArray())
            + (files.Count > 80 ? "..." : "") : ""));
        Report.Line("Индексы, подходящие к таблице " + Path.GetFileName(b.FilePath) + " (" + b.Tags.Count + "):");
        if (b.Tags.Count == 0) Report.Line("    нет - поиск возможен только перебором всей таблицы");
        foreach (IndexTag t in b.Tags) {
            StringBuilder sb = new StringBuilder("    " + t.Label + (t.Structural ? " (свой .CDX)" : "") + "  ключ: " + t.Expr
                + "  (длина " + t.KeyLen + ", тип " + t.KeyType + (t.Compact ? "" : ", обычный .IDX") + (t.Descending ? ", по убыванию" : "") + ")");
            if (t.ForText != null) sb.Append("  FOR " + t.ForText);
            if (t.Raw.Length > 0) sb.Append("  [" + t.Raw + "]");
            if (t.Problem != null) sb.Append("  - не годится: " + t.Problem);
            else {
                List<string> fields = new List<string>();
                foreach (string f in t.TermFields) fields.Add(f == null ? "?" : f);
                sb.Append("  -> поля " + string.Join("+", fields.ToArray()));
            }
            Report.Line(sb.ToString());
        }
    }

    // Справочник улиц: запомненный с прошлого раза или ищем по кодам улиц из начала базы
    StreetDirectory PrepareDirectory(LiveBase b, string dirFolder, string table, string savedTable, string savedKey, Dictionary<string, string> links) {
        int ul = b.Header.FieldIndex("UL");
        if (ul < 0) { Report.Line("Поля UL (улица) в таблице нет - справочник улиц не нужен"); return null; }
        if (savedTable != null && savedKey != null && File.Exists(Path.Combine(dirFolder, savedTable))) {
            try {
                Report.Line("Справочник улиц с прошлого раза: " + savedTable + "." + savedKey);
                StreetDirectory d = StreetDirectory.Load(dirFolder, savedTable, savedKey, links);
                if (d != null && d.Count > 0) return d;
            } catch (Exception ex) {
                Report.Line("    не открылся: " + ex.Message);
            }
        }
        List<string> codes = StreetDirectory.Codes(b.Sample, b.Sample.FieldIndex("UL"));
        if (codes == null) return null;
        Report.Line("Ищу справочник улиц среди таблиц папки:");
        try {
            return StreetDirectory.Detect(dirFolder, table, codes, delegate(int done, int total) {
                PostStatus("В поле улицы - коды. Ищу справочник улиц: проверено " + done + " из " + total + " таблиц...");
            });
        } catch (Exception ex) {
            Report.Line("    ошибка: " + ex.Message);
            return null;
        }
    }

    void AfterOpen(LiveBase b, Dictionary<string, bool> sel, string error, double seconds) {
        opening = false;
        SetBusy(false, null);
        if (error != null) {
            status.Text = "Не удалось открыть " + Path.Combine(folder, tableName) + ": " + error + "   Нажмите \"Где искать...\".";
            return;
        }
        live = b;
        baseInfo = tableName + ", " + b.Header.Expected.ToString("#,0") + " записей";
        selected = sel;
        UpdateSsButton();
        SaveSettings();
        BuildColumns();
        status.Text = "База открыта за " + seconds.ToString("0.0") + " с.";
        if (DbfTable.SplitWords(queryBox.Text).Length > 0) StartSearch();
        else UpdateStatus();
    }

    void SetBusy(bool busy, string text) {
        foreach (Control c in bottom.Controls) if (c is Button && c != scanButton) c.Enabled = !busy;
        Cursor = busy ? Cursors.AppStarting : Cursors.Default;
        if (text != null) { status.Text = text; status.Refresh(); }
    }

    AbonentQuery columns;

    void BuildColumns() {
        columns = new AbonentQuery(DbfTable.CreateLike(live.Header), live.Dir);
        grid.RowCount = 0;
        grid.Columns.Clear();
        foreach (int f in columns.Columns) {
            DbfField field = live.Header.Fields[f];
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = columns.ColumnTitle(f);
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            string n = field.Name.ToUpperInvariant();
            c.FillWeight = n == "FIO" ? 300 : n == "UL" ? 200 : (n == "NAB" || n == "NCH" || n == "TIPCH") ? 130 : 70;
            grid.Columns.Add(c);
        }
    }

    // ---- поиск (в фоне, чтобы окно не замирало, пока читается сеть) ----

    internal void StartSearch() {
        timer.Stop();
        if (live == null) return;
        if (scanning) { cancelScan = true; again = true; return; }
        if (searching) { again = true; return; }
        string text = queryBox.Text;
        if (DbfTable.SplitWords(text).Length == 0) { Show(null); return; }
        if (!cardWhenDone && !EnoughTyped(text, mode)) { Show(null); status.Text = "Впишите дальше или нажмите Enter."; return; }
        searching = true;
        SearchMode m = mode;
        Dictionary<string, bool> ss = selected;
        LiveBase b = live;
        status.Text = "Ищу...";
        Cursor = Cursors.AppStarting;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            SearchResult r = b.Search(text, m, ss);
            Post(delegate { SearchDone(r, b); });
        });
        th.IsBackground = true;
        th.Start();
    }

    // пока набирается одна-две буквы, не ищем (иначе по сети читается очень много лишнего)
    static bool EnoughTyped(string text, SearchMode m) {
        int digits = 0, letters = 0;
        foreach (char c in text) { if (char.IsDigit(c)) digits++; else if (char.IsLetter(c)) letters++; }
        switch (m) {
            case SearchMode.Account: return digits >= 4;
            case SearchMode.Meter: return digits + letters >= 4;
            case SearchMode.Name: return letters >= 3;
            default: return letters >= 3;
        }
    }

    void SearchDone(SearchResult r, LiveBase b) {
        searching = false;
        Cursor = Cursors.Default;
        if (b != live) return;
        if (again || r.Text != queryBox.Text || r.Mode != mode) { again = false; StartSearch(); return; }
        Show(r);
        if (cardWhenDone) {
            cardWhenDone = false;
            LogSearch();
            if (r.Rows.Count == 1) OpenCard(0);   // нашёлся один - сразу карточка
            else if (r.Rows.Count > 1) grid.Focus();
        }
    }

    internal void StartScan() {
        if (live == null || searching || scanning) return;
        string text = queryBox.Text;
        if (DbfTable.SplitWords(text).Length == 0) return;
        scanning = true;
        cancelScan = false;
        scanButton.Text = "Остановить перебор";
        SearchMode m = mode;
        Dictionary<string, bool> ss = selected;
        LiveBase b = live;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            SearchResult r = null;
            DateTime started = DateTime.Now;
            int shown = -1;
            try {
                r = b.Scan(text, m, ss, delegate(int done, int expected) {
                    if (cancelScan) throw new OperationCanceledException();
                    int pct = (int)(done * 100L / Math.Max(expected, 1));
                    if (pct == shown) return;
                    shown = pct;
                    double spent = (DateTime.Now - started).TotalSeconds;
                    string left = pct > 0 ? ", осталось примерно " + Math.Max(1, (int)(spent * (100 - pct) / pct / 60)) + " мин" : "";
                    PostStatus("Перебираю всю базу: " + pct + "%" + left + "...   Можно остановить кнопкой справа.");
                });
            } catch (OperationCanceledException) {
                r = null;
            } catch (Exception ex) {
                r = new SearchResult();
                r.Text = text;
                r.Mode = m;
                r.Error = ex.Message;
            }
            Post(delegate { ScanDone(r, b); });
        });
        th.IsBackground = true;
        th.Start();
    }

    void ScanDone(SearchResult r, LiveBase b) {
        scanning = false;
        scanButton.Text = "Искать перебором";
        if (b != live) return;
        if (again) { again = false; StartSearch(); return; }
        if (r == null) { UpdateStatus(); return; }
        Show(r);
    }

    void Show(SearchResult r) {
        current = r;
        grid.RowCount = 0;
        grid.RowCount = r == null ? 0 : r.Rows.Count;
        grid.Invalidate();
        scanButton.Visible = scanning || (r != null && r.NoIndex);
        logTimer.Stop();
        if (r != null) logTimer.Start();
        UpdateStatus();
    }

    void UpdateStatus() {
        if (live == null) return;
        string where = "База: " + baseInfo;
        if (selected != null) where += ", сельсоветов выбрано " + selected.Count + " из " + live.Dir.Selsovety.Count;
        SearchResult r = current;
        string took = r == null ? "" : "   (" + r.Seconds.ToString("0.0") + " с)";
        if (r == null)
            status.Text = "Впишите " + Prompts[(int)mode] + ".   " + where;
        else if (r.Error != null)
            status.Text = "Ошибка при поиске: " + r.Error + "   Нажмите \"Отчёт\" и пришлите его.";
        else if (r.NoIndex)
            status.Text = "Для поиска по полю \"" + Prompts[(int)r.Mode] + "\" в базе нет подходящего индекса. Можно искать перебором всей базы - это долго (минуты). Кнопка справа.";
        else if (r.Note.Length > 0 && r.Rows.Count == 0)
            status.Text = r.Note + "   " + where;
        else if (r.Rows.Count == 0)
            status.Text = (r.Hidden > 0 ? "В выбранных сельсоветах не найдено, а в других нашлось: " + r.Hidden + " - отметьте их в списке \"Сельсоветы\"."
                : "Ничего не найдено.") + took + "   " + where;
        else
            status.Text = "Найдено: " + r.Rows.Count + (r.TooMany ? " (нашлось очень много, показана часть - уточните запрос)" : "")
                + (r.Hidden > 0 ? ", ещё " + r.Hidden + " в других сельсоветах скрыто" : "")
                + ". Enter или двойной щелчок - карточка." + took + "   " + where;
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        SearchResult r = current;
        if (r == null || e.RowIndex >= r.Rows.Count || e.ColumnIndex >= r.Query.Columns.Length) return;
        e.Value = r.Query.CellText(r.Rows[e.RowIndex], r.Query.Columns[e.ColumnIndex]);
    }

    void OnQueryKey(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) {
            e.Handled = true;
            e.SuppressKeyPress = true;
            cardWhenDone = true;
            StartSearch();
        } else if (e.KeyCode == Keys.Down && current != null && current.Rows.Count > 0) {
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
        SearchResult r = current;
        if (r == null || index < 0 || index >= r.Rows.Count) return;
        int row = r.Rows[index];
        Card.Show(this, r.Table, row, r.Query.CardExtra(row));
    }

    // ---- сельсоветы: выпадающий список с галочками, только отбирает показанное ----

    string SelsovetyFile { get { return Path.Combine(Dir, "selsovety.txt"); } }

    // список сельсоветов из файла selsovety.txt (если его нет - ваши шесть)
    string[] ReadSelsovety() {
        string file = SelsovetyFile;
        if (!File.Exists(file)) return DefaultSelsovety;
        List<string> names = new List<string>();
        try {
            foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                string n = line.Trim();
                if (n.Length > 0 && !n.StartsWith("#")) names.Add(n);
            }
        } catch (Exception) {
            return DefaultSelsovety;
        }
        return names.ToArray();
    }

    void SaveSelsovety(ICollection<string> names) {
        List<string> lines = new List<string>();
        lines.Add("# Сельсоветы, абоненты которых показываются в поиске - по одному в строке (пусто - все)");
        lines.AddRange(names);
        Places.SaveSettings(SelsovetyFile, lines.ToArray());
    }

    // названия из файла -> сельсоветы справочника ("Боровлянский" находит "Боровлянский с/с")
    Dictionary<string, bool> LoadSelection(LiveBase b) {
        if (b.Dir == null) return null;
        List<string> all = b.Dir.Selsovety;
        string[] wanted = ReadSelsovety();
        Dictionary<string, bool> sel = new Dictionary<string, bool>();
        List<string> missing = new List<string>();
        foreach (string w in wanted) {
            string stem = StreetDirectory.Stem(w);
            bool hit = false;
            foreach (string name in all) {
                string l = name.ToLowerInvariant();
                if (stem.Length > 0 && (l.Contains(stem) || stem.Contains(StreetDirectory.Stem(name)))) { sel[name] = true; hit = true; }
            }
            if (!hit) missing.Add(w);
        }
        Report.Line("Сельсоветы в справочниках (" + all.Count + "): " + string.Join(", ", all.ToArray()));
        Report.Line("Выбраны для показа: " + (sel.Count == 0 || sel.Count == all.Count ? "все" : string.Join(", ", new List<string>(sel.Keys).ToArray()))
            + (missing.Count > 0 ? "; не нашлись в справочниках: " + string.Join(", ", missing.ToArray()) : ""));
        if (sel.Count == 0 || sel.Count == all.Count) return null;
        return sel;
    }

    void UpdateSsButton() {
        if (live == null || live.Dir == null || live.Dir.Selsovety.Count == 0) {
            ssButton.Text = "Сельсоветы ▼";
            return;
        }
        int total = live.Dir.Selsovety.Count;
        ssButton.Text = selected == null ? "Сельсоветы: все ▼" : "Сельсоветы: " + selected.Count + " из " + total + " ▼";
    }

    internal void ShowSelsovety() {
        if (live == null) { MessageBox.Show(this, "Подождите, пока откроется база.", Text); return; }
        if (live.Dir == null || live.Dir.Selsovety.Count == 0) {
            MessageBox.Show(this, "В справочниках улиц не нашлось сельсоветов - отбирать не по чему. Нажмите \"Отчёт\" и пришлите его.", Text);
            return;
        }
        List<string> all = live.Dir.Selsovety;
        ToolStripDropDown dd = new ToolStripDropDown();
        Panel panel = new Panel();
        panel.Font = Font;
        CheckedListBox list = new CheckedListBox();
        list.Font = Font;
        list.CheckOnClick = true;
        list.IntegralHeight = false;
        foreach (string name in all) list.Items.Add(name, selected == null || selected.ContainsKey(name));
        int listHeight = Math.Min(all.Count, 20) * list.ItemHeight + 6;
        int width = 240;
        foreach (string name in all) width = Math.Max(width, TextRenderer.MeasureText(name, Font).Width + 50);
        list.Location = new Point(0, 0);
        list.Size = new Size(width, listHeight);
        Button allButton = new Button();
        allButton.Font = Font;
        allButton.Text = "Все";
        allButton.Location = new Point(4, listHeight + 4);
        allButton.Size = new Size(80, 26);
        Button noneButton = new Button();
        noneButton.Font = Font;
        noneButton.Text = "Снять все";
        noneButton.Location = new Point(90, listHeight + 4);
        noneButton.Size = new Size(100, 26);
        panel.Size = new Size(width, listHeight + 34);
        panel.Controls.Add(list);
        panel.Controls.Add(allButton);
        panel.Controls.Add(noneButton);
        bool bulk = false;
        list.ItemCheck += delegate { if (!bulk) BeginInvoke(new MethodInvoker(delegate { ApplySelection(list); })); };
        EventHandler setAll = delegate(object s, EventArgs e) {
            bulk = true;
            for (int i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, s == allButton);
            bulk = false;
            ApplySelection(list);
        };
        allButton.Click += setAll;
        noneButton.Click += setAll;
        ToolStripControlHost host = new ToolStripControlHost(panel);
        host.Margin = Padding.Empty;
        host.Padding = Padding.Empty;
        host.AutoSize = false;
        host.Size = panel.Size;
        dd.Padding = Padding.Empty;
        dd.Items.Add(host);
        dd.Show(ssButton, new Point(0, ssButton.Height));
        list.Focus();
    }

    void ApplySelection(CheckedListBox list) {
        List<string> names = new List<string>();
        foreach (object item in list.CheckedItems) names.Add((string)item);
        Dictionary<string, bool> sel = new Dictionary<string, bool>();
        foreach (string n in names) sel[n] = true;
        selected = sel.Count == 0 || sel.Count == list.Items.Count ? null : sel;
        SaveSelsovety(names);
        UpdateSsButton();
        Report.Search("Сельсоветы для показа: " + (selected == null ? "все" : string.Join(", ", names.ToArray())));
        if (current != null) StartSearch(); else UpdateStatus();
    }

    internal BatchSearch LastBatch;

    internal void OpenBatch() {
        if (live == null) {
            MessageBox.Show(this, "Подождите, пока откроется база.", Text);
            return;
        }
        LastBatch = new BatchSearch(live, selected, mode, columns);
        LastBatch.Show(this);
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
        string newFolder = folderBox.Text.Trim(), newTable = ((TableInfo)list.SelectedItem).Name;
        if (string.Compare(newFolder, folder, StringComparison.OrdinalIgnoreCase) != 0 || string.Compare(newTable, tableName, StringComparison.OrdinalIgnoreCase) != 0) {
            folder = newFolder;
            tableName = newTable;
            Places.SaveSettings(settingsFile, new string[] { folder, tableName, ((int)mode).ToString() });   // справочник искать заново
        }
        OpenBase();
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
    readonly LiveBase live;
    readonly Dictionary<string, bool> selected;
    SearchMode mode;
    internal readonly RadioButton[] modeButtons = new RadioButton[4];
    internal readonly TextBox input = new TextBox();
    readonly DataGridView grid = new DataGridView();
    internal readonly Label status = new Label();
    readonly Button findButton = new Button();
    readonly List<Item> items = new List<Item>();
    bool running;
    volatile bool stop;

    const int MaxPerLine = 20;   // если по строке нашлось много - показываем первые

    class Item {
        public int Line;
        public string Text;
        public string Result;
        public SearchResult Res;  // где лежит найденная запись
        public int Row = -1;      // запись абонента или -1, если не найден
        public int Kind;          // 0 - найден один, 1 - несколько, 2 - не найден
    }

    static readonly string[] ModeNames = { "&Лицевые счета", "&Адреса", "Номера &счётчиков", "&ФИО" };

    public BatchSearch(LiveBase live, Dictionary<string, bool> selected, SearchMode mode, AbonentQuery query) {
        this.live = live;
        this.selected = selected;
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
        FormClosing += delegate { stop = true; };
    }

    void AddColumn(string title, int weight) {
        DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
        c.HeaderText = title;
        c.FillWeight = weight;
        c.SortMode = DataGridViewColumnSortMode.NotSortable;
        grid.Columns.Add(c);
    }

    // Ищем строку за строкой в фоне; кнопка "Найти" на это время - "Стоп"
    internal void Run() {
        if (running) { stop = true; return; }
        List<string> lines = new List<string>();
        foreach (string raw in input.Text.Replace("\r\n", "\n").Split('\n')) {
            string line = raw.Replace('\t', ' ').Trim();
            if (line.Length > 0) lines.Add(line);
        }
        if (lines.Count == 0) { MessageBox.Show(this, "Вставьте список: по одному значению в строке.", Text); return; }
        running = true;
        stop = false;
        findButton.Text = "Стоп";
        items.Clear();
        grid.RowCount = 0;
        SearchMode m = mode;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            DateTime started = DateTime.Now;
            int single = 0, several = 0, none = 0, done = 0;
            for (int i = 0; i < lines.Count && !stop; i++) {
                SearchResult r = live.Search(lines[i], m, selected, 100);
                List<Item> add = new List<Item>();
                if (r.Rows.Count == 0) {
                    string why = r.Error != null ? "ошибка: " + r.Error : r.NoIndex ? "нет индекса для поиска"
                        : r.Hidden > 0 ? "в выбранных сельсоветах нет (в других: " + r.Hidden + ")" : "не найден";
                    add.Add(NewItem(i + 1, lines[i], why, null, -1, 2));
                    none++;
                } else {
                    string result = r.Rows.Count == 1 ? "найден" : "найдено " + r.Rows.Count + (r.Rows.Count > MaxPerLine ? " (показаны " + MaxPerLine + ")" : "");
                    if (r.Hidden > 0) result += " (+" + r.Hidden + " в других с/с)";
                    if (r.Rows.Count == 1) single++; else several++;
                    for (int k = 0; k < r.Rows.Count && k < MaxPerLine; k++)
                        add.Add(NewItem(i + 1, lines[i], result, r, r.Rows[k], r.Rows.Count == 1 ? 0 : 1));
                }
                done = i + 1;
                int line = done;
                Post(delegate {
                    items.AddRange(add);
                    grid.RowCount = items.Count;
                    status.Text = "Ищу: строка " + line + " из " + lines.Count + "...   (кнопка \"Стоп\" - прервать)";
                });
            }
            double seconds = (DateTime.Now - started).TotalSeconds;
            int s1 = single, s2 = several, s3 = none, d = done;
            Post(delegate { Finish(lines, d, s1, s2, s3, seconds); });
        });
        th.IsBackground = true;
        th.Start();
    }

    void Post(MethodInvoker m) {
        try { if (!IsDisposed) BeginInvoke(m); } catch (Exception) { }
    }

    void Finish(List<string> lines, int done, int single, int several, int none, double seconds) {
        running = false;
        findButton.Text = "Найти";
        grid.RowCount = items.Count;
        grid.Invalidate();
        status.Text = (done < lines.Count ? "Остановлено на строке " + done + " из " + lines.Count + ". " : "Строк в списке: " + lines.Count + ". ")
            + "Найдено по одному абоненту: " + single + ", по несколько: " + several + ", не найдено: " + none + ". Двойной щелчок - карточка.";
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
            + (done < lines.Count ? " (остановлен на " + done + ")" : "") + ", найдено по одному " + single + ", по несколько " + several + ", не найдено " + none
            + ", время " + seconds.ToString("0.0") + " с; цифр в строках: " + string.Join(", ", shape.ToArray()));
    }

    static Item NewItem(int line, string text, string result, SearchResult res, int row, int kind) {
        Item it = new Item();
        it.Line = line;
        it.Text = text;
        it.Result = result;
        it.Res = res;
        it.Row = row;
        it.Kind = kind;
        return it;
    }

    string Cell(Item it, int column) {
        if (column == 0) return it.Line.ToString();
        if (column == 1) return it.Text;
        if (column == 2) return it.Result;
        if (it.Row < 0 || it.Res == null) return "";
        return it.Res.Query.CellText(it.Row, it.Res.Query.Columns[column - 3]);
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
        Item it = items[index];
        Card.Show(this, it.Res.Table, it.Row, it.Res.Query.CardExtra(it.Row));
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
