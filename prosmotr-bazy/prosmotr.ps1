# Поиск абонента в базе FoxPro (таблицы .dbf) - ТОЛЬКО ЧТЕНИЕ.
# Таблицы открываются только на чтение, в базу ничего не записывается.
# Запуск: zapusk.vbs (без чёрного окна) или zapusk.bat.

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

    // Индекс по ФИО программы базы может лежать не в папке базы: программа базы запускается с диска D:
    // (d:\Energ, временные файлы - d:\temp), бывают и вложенные папки рядом с базой. Ищем там .IDX/.CDX
    // с ФИО в выражении; подходящий к этой таблице (проверяется на её записях) - подключаем для поиска по ФИО.
    // где может лежать что-то от программы базы: папка базы (с вложенными), папка над ней,
    // на локальных дисках - корень, Energ, temp, fpd26 (с вложенными)
    public static List<string> ProgramDirs(string baseFolder) {
        List<string> dirs = new List<string>();
        AddDir(dirs, baseFolder, 50);
        try { DirectoryInfo up = Directory.GetParent(baseFolder.TrimEnd('\\', '/')); if (up != null) AddDir(dirs, up.FullName, 0); } catch (Exception) { }
        foreach (string drive in new string[] { "C", "D", "E" }) {
            try {
                DriveInfo di = new DriveInfo(drive);
                if (di.DriveType != DriveType.Fixed || !di.IsReady) continue;
                string root = drive + ":\\";
                AddDir(dirs, root, 0);
                foreach (string name in new string[] { "Energ", "temp", "fpd26", "BASAFOX" }) AddDir(dirs, Path.Combine(root, name), 30);
            } catch (Exception) { }
        }
        return dirs;
    }

    // Исходники программы базы (.PRG, .SPR, .MPR - обычный текст): строки, где поиск или индекс связан с ФИО.
    // В отчёт - только эти строки кода (данных абонентов в них нет), чтобы искать по ФИО так же, как она.
    public static void ReportSources(List<string> dirs, List<string> notes) {
        int files = 0, shown = 0;
        string[] verbs = { "INDE", "SEEK", "ORDER", "LOCATE", "SELECT", "FILTER", "FIND", ".IDX", ".CDX", "TAG", "LIKE", "SCAN" };
        Encoding enc = Encoding.GetEncoding(866);
        foreach (string dir in dirs) {
            string[] list;
            try { list = Directory.GetFiles(dir); } catch (Exception) { continue; }
            foreach (string f in list) {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".prg" && ext != ".spr" && ext != ".mpr") continue;
                files++;
                string[] lines;
                try {
                    if (new FileInfo(f).Length > 4 << 20) continue;
                    lines = File.ReadAllLines(f, enc);
                } catch (Exception) { continue; }
                for (int i = 0; i < lines.Length && shown < 80; i++) {
                    string u = lines[i].ToUpperInvariant();
                    if (u.IndexOf("FIO") < 0 && u.IndexOf("FAM") < 0) continue;
                    bool verb = false;
                    foreach (string v in verbs) if (u.IndexOf(v) >= 0) { verb = true; break; }
                    if (!verb) continue;
                    string line = lines[i].Trim();
                    if (line.Length > 200) line = line.Substring(0, 200) + "...";
                    notes.Add(Path.GetFileName(f) + ":" + (i + 1) + ": " + line);
                    shown++;
                }
            }
        }
        notes.Insert(0, "исходников программы базы (.PRG/.SPR/.MPR) найдено: " + files + (shown > 0 ? ", строки про поиск по ФИО:" : ", строк про поиск по ФИО нет"));
    }

    public IndexTag FindNameIndexElsewhere(string baseFolder, List<string> notes) {
        List<string> dirs = ProgramDirs(baseFolder);
        int files = 0;
        IndexTag found = null;
        Dictionary<string, bool> known = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (IndexTag t in Tags) known[t.FilePath + "|" + t.Name] = true;
        foreach (string dir in dirs) {
            string[] list;
            try { list = Directory.GetFiles(dir); } catch (Exception) { continue; }
            foreach (string f in list) {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".idx" && ext != ".cdx" && ext != ".ndx" && ext != ".ntx" && ext != ".mdx" && ext != ".tmp") continue;
                if (ext == ".tmp" && dir.TrimEnd('\\', '/').ToLowerInvariant().EndsWith("temp") == false) continue;
                files++;
                string upper = Path.GetFileNameWithoutExtension(f).ToUpperInvariant();
                if (ext == ".tmp") { notes.Add(DescribeTemp(f)); continue; }
                if (ext != ".idx" && ext != ".cdx") {
                    if (upper.Contains("FIO") || upper.Contains("FAM")) notes.Add(f + " - индекс не FoxPro (не прочитать)");
                    continue;
                }
                List<IndexTag> tags;
                try { tags = IndexReader.ReadTags(f, Header.Enc); } catch (Exception) { continue; }
                foreach (IndexTag t in tags) {
                    string e = (t.Expr ?? "").ToUpperInvariant();
                    if (!e.Contains("FIO") && !e.Contains("FAM")) continue;
                    if (known.ContainsKey(t.FilePath + "|" + t.Name)) continue;
                    string where = f + (t.Name != null && t.Name.Length > 0 && ext == ".cdx" ? ":" + t.Name : "") + " (" + t.Expr + ")";
                    t.Structural = false;
                    Prepare(t);
                    if (t.Key == null || string.Compare(t.FirstField, "FIO", StringComparison.OrdinalIgnoreCase) != 0) { notes.Add(where + " - не от таблицы абонентов"); continue; }
                    if (found != null) { notes.Add(where); continue; }
                    if (Verify(t)) {
                        List<IndexTag> l = new List<IndexTag>(Tags);
                        l.Add(t);
                        Tags = l;      // новый список, а не добавление: поиски в других потоках перебирают старый
                        found = t;
                        notes.Add(where + " - ПОДХОДИТ к таблице абонентов, поиск по ФИО идёт через него");
                    } else notes.Add(where + " - не подходит: " + t.Check);
                }
            }
        }
        notes.Insert(0, "просмотрено индексных файлов: " + files + " в папках: " + string.Join(", ", dirs.ToArray()));
        return found;
    }

    // Временный файл программы базы (d:\temp\*.TMP): что это - таблица (результат запроса) или индекс и по чему.
    // Только описание для отчёта, без данных: размер, поля, выражение индекса.
    string DescribeTemp(string f) {
        string head = f;
        try {
            FileInfo fi = new FileInfo(f);
            head += " (" + (fi.Length >= 1048576 ? (fi.Length / 1048576.0).ToString("0.0") + " МБ" : (fi.Length / 1024) + " КБ")
                + ", изменён " + fi.LastWriteTime.ToString("dd.MM HH:mm") + ")";
            byte[] h = new byte[32];
            using (FileStream fs = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                if (DbfTable.ReadUpTo(fs, h, 32) < 32) return head + ": пустой";
            int count = BitConverter.ToInt32(h, 4), hl = BitConverter.ToUInt16(h, 8), rl = BitConverter.ToUInt16(h, 10);
            if (hl >= 65 && rl >= 2 && count >= 0 && (long)hl + (long)count * rl <= fi.Length + 1 && fi.Length - hl - (long)count * rl <= rl + 2) {
                DbfTable t = DbfTable.LoadHeader(f);
                List<string> names = new List<string>();
                foreach (DbfField fld in t.Fields) names.Add(fld.Name);
                return head + ": таблица, записей " + count + ", поля " + string.Join(", ", names.ToArray());
            }
            foreach (bool cdx in new bool[] { true, false }) {
                try {
                    List<string> exprs = new List<string>();
                    foreach (IndexTag t in IndexReader.ReadTags(f, Header.Enc, cdx))
                        if (t.Expr.Length > 0 && t.Expr.Length < 200 && FoxExprLooksValid(t.Expr)) exprs.Add((cdx ? t.Name + ": " : "") + t.Expr);
                    if (exprs.Count > 0) return head + ": индекс" + (cdx ? " (как .CDX)" : " (как .IDX)") + " - " + string.Join("; ", exprs.ToArray());
                } catch (Exception) { }
            }
            return head + ": не таблица и не индекс";
        } catch (Exception ex) {
            return head + ": не открылся - " + ex.Message;
        }
    }

    // выражение индекса похоже на настоящее: печатные символы и хотя бы одна буква
    static bool FoxExprLooksValid(string e) {
        bool letter = false;
        foreach (char c in e) {
            if (c < ' ' || c == '\uFFFD') return false;
            if (char.IsLetter(c)) letter = true;
        }
        return letter;
    }

    static void AddDir(List<string> dirs, string dir, int subdirs) {
        try {
            if (!Directory.Exists(dir)) return;
            foreach (string d in dirs) if (string.Compare(d.TrimEnd('\\', '/'), dir.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase) == 0) return;
            dirs.Add(dir);
            if (subdirs <= 0) return;
            string[] subs = Directory.GetDirectories(dir);
            for (int i = 0; i < subs.Length && i < subdirs; i++) AddDir(dirs, subs[i], 0);
        } catch (Exception) { }
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
        if (mode == SearchMode.Name && r.Rows.Count == 0 && r.Hidden == 0 && r.NameText == null && !r.NoIndex && TagFor("FIO") != null) {
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
    bool SeekNameByAddress(string text, Seeker s, SearchResult res, bool partial) {
        if (Dir == null) return false;
        string name = null, addr = null;
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
        if (addr == null || name == null || name.Length == 0) return false;
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
public class BaseCard : Form {
    public static readonly Color Back = Color.FromArgb(0, 0, 160);
    static readonly Color LabelColor = Color.White;
    static readonly Color BoxColor = Color.FromArgb(0, 128, 168);
    static readonly Color TitleColor = Color.FromArgb(255, 255, 85);
    static readonly Color MarkColor = Color.FromArgb(255, 130, 130);
    const int Cols = 80;

    // Экран "АБОНЕНТ" из программы базы. {ПОЛЕ,ширина} - значение поля, {#...} - из справочников,
    // {@N} - с колонки N, {-} - черта. Надпись перед полем показывается, только если поле есть в таблице.
    static readonly string[] MainScreen = {
        "№ Абонента:{NAB,10}{@35}{#T1,6}{@43}{#T2,6}{@51}{#T3,6}",
        "Ф.И.О.:{FIO,71}",
        "Н.П.:{#NP,32}{@44}{#TNP,4}{@50}с/с:{#SS,24}",
        "Код:{UL,5} Ул.:{#ULN,32}{@59}Л\\С:{LS_ERIP,14}",
        "Дом:{DOM,7}{@14}Кор.:{KOR,3}{@24}Кв.:{KV,5}{@34}Тел.:{TEL,7}{@50}№ Книги:{KNIGA,4}",
        " Код лг.:{LGI,3} Наимен.:{#LGN,30}{@57}% Ск.:{SKID,3}",
        " Код тар.:{KTAR,2} {#TARN,34}{@58}Жил.:{GIL,3} Лг.Жил.:{KLG,3}",
        " Тип пом.:{TPO,3} Розеток:{ROZ,3} Точек осв.:{TOS,3}",
        "{-}",
        "Сч.:{NCH,20}{@28}{TIPCH,30}{@60}Разр.:{RAS,1} Г.пов.:{GODPR,4}",
        "Кв.пов.:{KVP,2}",
        "{@26}Место уст.ПУ:{#MESTO,30}",
        "Пок.уст.:{NPOK,8} Д.пр.:{DPR,10} Пок.:{POK,8} Пр.пров.:{PPR,3} Опл.:{OPOK,8}",
        "{-}",
        "Дата зам.:{DZA,10} Пок.зам.сч.:{PZS,8} № Сч.:{NSC,20}",
        "{@43}Тип:{TSC,30}",
        "Пр.архива:{PRA,3} Дата откл.:{DOT,10} Пр.пл.:{PRP,5}{@62}№ уч-ка:{N_UCH,4}",
        "Дата уч.без пр.:{DUC,10} Дата оконч.:{DOK,10} № Фидера:{N_F,15}",
        "{-}",
        "Мощность:{#POWER,8}",
    };

    // экран "ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ" (Ctrl+D): сверху - абонент, ниже - поля таблиц DOP и DOGOVOR
    static readonly string[] ExtraTop = { "№ Абонента:{NAB,10}", "Ф.И.О.:{FIO,71}" };

    class CardValue {
        public string Text = "";
        public bool Right, Mark;
    }

    SearchResult res;
    int row;
    object[] values;
    List<ExtraRecord> extra;
    Dictionary<int, bool> used = new Dictionary<int, bool>();
    bool extraPage;
    readonly LiveBase live;
    readonly CardMove move;
    readonly ScreenPanel screen;
    readonly Button prevButton, nextButton, extraButton, closeButton;
    readonly FlowLayoutPanel bar;

    public BaseCard(SearchResult r, int row, LiveBase live, CardMove move) {
        this.live = live;
        this.move = move;
        Text = "Карточка абонента";
        BackColor = Back;
        StartPosition = FormStartPosition.CenterScreen;
        KeyPreview = true;
        MinimizeBox = false;
        screen = new ScreenPanel(PickFont(11f));
        screen.Dock = DockStyle.Fill;
        bar = new FlowLayoutPanel();
        bar.Dock = DockStyle.Bottom;
        bar.Height = 36;
        bar.WrapContents = false;
        bar.Padding = new Padding(4, 4, 4, 0);
        bar.BackColor = SystemColors.Control;
        prevButton = AddButton("< Пред. (PgUp)", delegate { Step(-1); });
        nextButton = AddButton("След. (PgDn) >", delegate { Step(1); });
        extraButton = AddButton("Доп. сведения (Ctrl+D)", delegate { ToggleExtra(); });
        AddButton("Все поля (F2)", delegate { ResultColumns.ShowList(this, res, this.row, live == null ? null : live.Extra); });
        closeButton = AddButton("Закрыть (Esc)", delegate { Close(); });
        prevButton.Enabled = nextButton.Enabled = move != null;
        Controls.Add(bar);
        Controls.Add(screen);
        screen.BringToFront();
        FormClosed += delegate { screen.CellFont.Dispose(); };
        Open(r, row);
    }

    Button AddButton(string text, EventHandler click) {
        Button b = new Button();
        b.Text = text;
        b.AutoSize = true;
        b.UseVisualStyleBackColor = true;
        b.Click += click;
        bar.Controls.Add(b);
        return b;
    }

    static Font PickFont(float size) {
        foreach (string name in new string[] { "Consolas", "Lucida Console", "Courier New" }) {
            Font f = new Font(name, size);
            if (string.Compare(f.Name, name, StringComparison.OrdinalIgnoreCase) == 0) return f;
            f.Dispose();
        }
        return new Font(FontFamily.GenericMonospace, size);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
        switch (keyData) {
            case Keys.Escape: if (extraPage) ToggleExtra(); else Close(); return true;
            case Keys.Control | Keys.D: ToggleExtra(); return true;
            case Keys.F2: ResultColumns.ShowList(this, res, row, live == null ? null : live.Extra); return true;
            case Keys.PageUp: Step(-1); return true;
            case Keys.PageDown: Step(1); return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    public void Step(int delta) {
        SearchResult r;
        int rw;
        if (move != null && move(delta, out r, out rw)) Open(r, rw);
        else System.Media.SystemSounds.Beep.Play();
    }

    public void ToggleExtra() {
        extraPage = !extraPage;
        Build();
    }

    void Open(SearchResult r, int rw) {
        res = r;
        row = rw;
        using (MemoReader memo = r.Table.OpenMemo()) values = r.Table.GetRow(rw, memo);
        extra = null;
        Build();
    }

    void Build() {
        screen.Begin(extraPage ? "ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ" : "АБОНЕНТ");
        used = new Dictionary<int, bool>();
        if (extraPage) BuildExtra(); else BuildMain();
        screen.End();
        extraButton.Text = extraPage ? "Карточка (Ctrl+D)" : "Доп. сведения (Ctrl+D)";
        Text = "Абонент " + Field("NAB") + (extraPage ? " - доп. сведения" : "") + " (только просмотр)";
        if (WindowState == FormWindowState.Normal) {
            Size content = screen.ContentSize;
            Rectangle area = Screen.FromControl(this).WorkingArea;
            int h = Math.Min(content.Height + bar.Height, area.Height - 60);
            int w = content.Width + (h < content.Height + bar.Height ? SystemInformation.VerticalScrollBarWidth : 0);
            ClientSize = new Size(Math.Min(w, area.Width - 40), h);
        }
        ActiveControl = closeButton;
    }

    void BuildMain() {
        int r = 0;
        bool line = false;
        foreach (string t in MainScreen) {
            if (t == "{-}") { line = r > 0; continue; }
            int at = line ? r + 1 : r;
            if (!Template(t, at)) continue;
            if (line) screen.Line(r, null);
            r = at + 1;
            line = false;
        }
        // поля таблицы, которых нет на экране базы
        List<int> rest = new List<int>();
        for (int j = 0; j < values.Length; j++) if (!used.ContainsKey(j)) rest.Add(j);
        if (rest.Count > 0) r = Flow(r, res.Table, values, rest);
        List<string> notes = Notes();
        if (notes.Count > 0) {
            screen.Line(r++, null);
            foreach (string n in notes) screen.Label(r++, 0, n, TitleColor);
        }
    }

    void BuildExtra() {
        int r = 0;
        foreach (string t in ExtraTop) if (Template(t, r)) r++;
        ExtraInfo x = live == null ? null : live.Extra;
        if (x == null || x.Tables.Count == 0) {
            screen.Line(r++, null);
            screen.Label(r++, 0, "Таблиц доп. сведений (DOP.DBF, DOGOVOR.DBF) рядом с базой нет", LabelColor);
            return;
        }
        if (extra == null) {
            Cursor old = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try { extra = x.Records(Field("NAB")); }
            catch (Exception ex) {
                extra = new List<ExtraRecord>();
                ExtraRecord e = new ExtraRecord();
                e.Name = "Доп. сведения";
                e.Problem = "не прочитались: " + ex.Message;
                extra.Add(e);
            }
            finally { Cursor.Current = old; }
        }
        foreach (ExtraRecord e in extra) {
            screen.Line(r++, Path.GetFileNameWithoutExtension(e.Name).ToUpperInvariant());
            if (e.Problem != null) { screen.Label(r++, 0, e.Problem, LabelColor); continue; }
            List<int> fields = new List<int>();
            for (int j = 0; j < e.Values.Length; j++)
                if (string.Compare(e.Table.Fields[j].Name, "NAB", StringComparison.OrdinalIgnoreCase) != 0) fields.Add(j);
            r = Flow(r, e.Table, e.Values, fields);
        }
    }

    List<string> Notes() {
        List<string> l = new List<string>();
        string why;
        if (res.Related.TryGetValue(row, out why)) l.Add("Добавлен к найденному: " + why);
        else if (res.Added.ContainsKey(row)) l.Add("Добавлен к найденному: другой тариф того же счётчика");
        if (!res.Related.ContainsKey(row) && res.Orange.ContainsKey(row)) l.Add("Оранжевым: на адресе больше одного счётчика");
        int t;
        if (res.Tariff.TryGetValue(row, out t)) l.Add("Тариф T" + t + ": у счётчика несколько лицевых (таблица ABDIF_T)");
        return l;
    }

    // одна строка экрана по шаблону; false - на ней ничего не оказалось (полей нет в таблице)
    bool Template(string line, int r) {
        bool placed = false;
        int col = 0, textCol = 0, i = 0;
        StringBuilder text = new StringBuilder();
        while (i < line.Length) {
            if (line[i] != '{') {
                if (text.Length == 0) textCol = col;
                text.Append(line[i]);
                col++;
                i++;
                continue;
            }
            int end = line.IndexOf('}', i);
            string tok = line.Substring(i + 1, end - i - 1);
            i = end + 1;
            string label = text.ToString().TrimEnd();
            text.Length = 0;
            if (tok[0] == '@') {
                if (label.Trim().Length > 0) { screen.Label(r, textCol, label, LabelColor); placed = true; }
                col = int.Parse(tok.Substring(1));
                continue;
            }
            int comma = tok.IndexOf(',');
            int w = int.Parse(tok.Substring(comma + 1));
            CardValue v = Resolve(tok.Substring(0, comma));
            if (v != null) {
                if (label.Trim().Length > 0) screen.Label(r, textCol, label, LabelColor);
                if (v.Mark) screen.Label(r, col, v.Text, MarkColor);
                else screen.Box(r, col, w, v.Text, v.Right);
                placed = true;
            }
            col += w;
        }
        if (text.ToString().Trim().Length > 0) { screen.Label(r, textCol, text.ToString().TrimEnd(), LabelColor); placed = true; }
        return placed;
    }

    // поля одно за другим, по строкам ("Надпись:[значение]")
    int Flow(int r, DbfTable t, object[] vals, List<int> fields) {
        int col = 0;
        foreach (int j in fields) {
            DbfField f = t.Fields[j];
            string label = (FieldNames.Label(f.Name) ?? f.Name) + ":";
            CardValue v = FieldValue(f, vals[j]);
            int w = BoxWidth(f);
            if (label.Length + w > Cols) w = Cols - label.Length;
            if (col > 0 && col + label.Length + w > Cols) { r++; col = 0; }
            screen.Label(r, col, label, LabelColor);
            screen.Box(r, col + label.Length, w, v.Text, v.Right);
            col += label.Length + w + 2;
        }
        return col > 0 ? r + 1 : r;
    }

    static int BoxWidth(DbfField f) {
        switch (f.Type) {
            case 'D': return 10;
            case 'T': return 19;
            case 'L': return 3;
            case 'M': case 'G': return Cols;
            case 'C': return Math.Max(1, Math.Min(f.Length, 60));
            default: return Math.Max(1, f.Length);
        }
    }

    CardValue Resolve(string key) {
        if (key[0] == '#') return Computed(key.Substring(1));
        int i = res.Table.FieldIndex(key);
        if (i < 0) return null;
        used[i] = true;
        return FieldValue(res.Table.Fields[i], values[i]);
    }

    static CardValue FieldValue(DbfField f, object v) {
        CardValue c = new CardValue();
        c.Right = f.Type == 'N' || f.Type == 'F' || f.Type == 'I' || f.Type == 'Y' || f.Type == 'B';
        if (v is bool) c.Text = (bool)v ? "да" : "нет";
        else c.Text = DbfTable.Format(v, true).Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
        return c;
    }

    string Field(string name) {
        int i = res.Table.FieldIndex(name);
        return i < 0 ? "" : DbfTable.Format(values[i], true).Trim();
    }

    static CardValue Plain(string text) {
        CardValue c = new CardValue();
        c.Text = text ?? "";
        return c;
    }

    CardValue Computed(string key) {
        StreetDirectory dir = live == null ? null : live.Dir;
        string ul = Field("UL");
        switch (key) {
            case "T1": case "T2": case "T3": {
                int t;
                res.Tariff.TryGetValue(row, out t);
                CardValue c = Plain((t == key[1] - '0' ? "(•) " : "( ) ") + key);
                c.Mark = true;
                return c;
            }
            case "NP": case "TNP": case "SS": case "ULN": {
                if (res.Table.FieldIndex("UL") < 0) return null;
                if (dir == null) return Plain("");
                if (key == "SS") return Plain(dir.Selsovet(ul));
                if (key == "ULN") return Plain(dir.Street(ul));
                string town = dir.Town(ul) ?? "", type = "", name = town;
                int sp = town.IndexOf(' ');
                if (sp > 0 && town[sp - 1] == '.') { type = town.Substring(0, sp); name = town.Substring(sp + 1); }
                return Plain(key == "NP" ? name : type);
            }
            case "LGN": return Lookup("LGI", "SPRLGOT.DBF", new string[] { "LGI", "KODLG", "KOD", "KLG", "KL" });
            case "TARN": return Lookup("KTAR", "SPRTAR.DBF", new string[] { "KTAR", "KOD", "KT" });
            case "MESTO": return Lookup("IDTPSCH", "TPSCH.DBF", new string[] { "IDTPSCH", "ID", "KOD" });
            case "POWER": {
                ExtraInfo x = live == null ? null : live.Extra;
                if (x == null || x.PowerBase == null) return null;
                string p = res.PowerText(row);
                if (p.Length == 0)
                    try {
                        Dictionary<string, string> d = x.Power(new List<string>(new string[] { Field("NAB") }));
                        if (!d.TryGetValue(ExtraInfo.Nab(Field("NAB")), out p)) p = "";
                    } catch (Exception) { p = ""; }
                CardValue c = Plain(p);
                c.Right = true;
                return c;
            }
        }
        return null;
    }

    // название по коду из справочника базы; не нашлось - пусто (а поле с кодом показывается среди прочих)
    CardValue Lookup(string field, string table, string[] keys) {
        int i = res.Table.FieldIndex(field);
        if (i < 0) return null;
        string code = Field(field);
        if (code.Length == 0 || live == null) return Plain("");
        string name = CodeBook.Get(Path.GetDirectoryName(live.FilePath), table, keys).Name(code);
        if (name.Length > 0) used[i] = true;
        return Plain(name);
    }

    // Экран в знакоместах: надписи рисуются, значения - поля только для чтения (можно выделить и скопировать)
    class ScreenPanel : Panel {
        public readonly Font CellFont;
        readonly float charW;
        readonly int lineH;
        string title = "";
        int rows;
        readonly List<KeyValuePair<Point, KeyValuePair<string, Color>>> texts = new List<KeyValuePair<Point, KeyValuePair<string, Color>>>();
        readonly List<KeyValuePair<int, string>> lines = new List<KeyValuePair<int, string>>();
        const int Edge = 2;

        public ScreenPanel(Font font) {
            CellFont = font;
            DoubleBuffered = true;
            ResizeRedraw = true;
            AutoScroll = true;
            BackColor = Back;
            charW = TextRenderer.MeasureText(new string('Ж', 40), font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Width / 40f;
            lineH = font.Height + 7;
        }

        int X(int col) { return (int)Math.Round((Edge + col) * charW); }
        int Y(int row) { return lineH + lineH / 2 + row * lineH; }

        public Size ContentSize { get { return new Size(X(Cols + Edge), Y(rows) + lineH); } }

        public void Begin(string caption) {
            SuspendLayout();
            AutoScrollPosition = new Point(0, 0);
            while (Controls.Count > 0) { Control c = Controls[0]; Controls.RemoveAt(0); c.Dispose(); }
            texts.Clear();
            lines.Clear();
            rows = 0;
            title = caption;
        }

        public void End() {
            AutoScrollMinSize = ContentSize;
            ResumeLayout();
            Invalidate();
        }

        public void Label(int row, int col, string text, Color color) {
            texts.Add(new KeyValuePair<Point, KeyValuePair<string, Color>>(new Point(col, row), new KeyValuePair<string, Color>(text, color)));
            rows = Math.Max(rows, row + 1);
        }

        public void Line(int row, string caption) {
            lines.Add(new KeyValuePair<int, string>(row, caption));
            rows = Math.Max(rows, row + 1);
        }

        public void Box(int row, int col, int width, string value, bool right) {
            TextBox b = new TextBox();
            b.ReadOnly = true;
            b.BorderStyle = BorderStyle.None;
            b.BackColor = BoxColor;
            b.ForeColor = Color.White;
            b.Font = CellFont;
            b.Text = value;
            b.TextAlign = right ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            b.Width = (int)Math.Ceiling(width * charW) + 3;
            b.Location = new Point(X(col) - 1, Y(row) + (lineH - b.Height) / 2);
            Controls.Add(b);
            rows = Math.Max(rows, row + 1);
        }

        protected override void OnScroll(ScrollEventArgs se) {
            base.OnScroll(se);
            Invalidate();
        }

        protected override void OnMouseWheel(MouseEventArgs e) {
            base.OnMouseWheel(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e) {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            int dx = AutoScrollPosition.X, dy = AutoScrollPosition.Y;
            Size size = ContentSize;
            using (Pen pen = new Pen(Color.FromArgb(220, 220, 220))) {
                Rectangle frame = new Rectangle(dx + (int)(charW * 0.7f), dy + lineH / 2, size.Width - (int)(charW * 1.4f), size.Height - lineH / 2 - lineH / 3);
                g.DrawRectangle(pen, frame);
                frame.Inflate(-3, -3);
                g.DrawRectangle(pen, frame);
                foreach (KeyValuePair<int, string> l in lines) {
                    int y = dy + Y(l.Key) + lineH / 2;
                    g.DrawLine(pen, dx + X(0), y, dx + X(Cols), y);
                    if (l.Value != null) Caption(g, " " + l.Value + " ", dx + size.Width / 2, y);
                }
            }
            Caption(g, " " + title + " ", dx + size.Width / 2, dy + lineH / 2);
            int pad = (lineH - CellFont.Height) / 2;
            foreach (KeyValuePair<Point, KeyValuePair<string, Color>> t in texts)
                TextRenderer.DrawText(g, t.Value.Key, CellFont, new Point(dx + X(t.Key.X), dy + Y(t.Key.Y) + pad), t.Value.Value, TextFormatFlags.NoPadding);
        }

        void Caption(Graphics g, string text, int centerX, int centerY) {
            Size s = TextRenderer.MeasureText(text, CellFont, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
            Rectangle r = new Rectangle(centerX - s.Width / 2, centerY - s.Height / 2, s.Width, s.Height);
            using (SolidBrush b = new SolidBrush(Back)) g.FillRectangle(b, r);
            TextRenderer.DrawText(g, text, CellFont, r.Location, TitleColor, TextFormatFlags.NoPadding);
        }
    }
}

public static class Card {
    public static void Show(IWin32Window owner, DbfTable table, int row) {
        Show(owner, table, row, null);
    }

    // extra - пары "подпись, значение", показываются первыми
    public static void Show(IWin32Window owner, DbfTable table, int row, string[] extra) {
        Show(owner, table, row, extra, null);
    }

    // after - пары, которые показываются после полей записи (доп. сведения из других таблиц)
    public static void Show(IWin32Window owner, DbfTable table, int row, string[] extra, List<string> after) {
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
        if (after != null)
            for (int i = 0; i + 1 < after.Count; i += 2) {
                int r = g.Rows.Add(after[i], after[i + 1]);
                if (after[i].StartsWith("---")) {
                    g.Rows[r].DefaultCellStyle.BackColor = SystemColors.Control;
                    g.Rows[r].DefaultCellStyle.Font = new Font(g.Font, FontStyle.Bold);
                }
            }
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
    internal readonly CheckBox partialBox = new CheckBox();

    static readonly string[] ModeNames = { "&Лицевой счёт", "&Адрес", "№ &счётчика", "&ФИО" };
    static readonly string[] Prompts = { "лицевой счёт", "адрес", "номер счётчика", "фамилию" };
    static readonly string[] Hints = {
        "Номер целиком:  044060216   (нули впереди можно не писать)",
        "Населённый пункт, улица, дом, корпус, квартира:  Острошицкий Городок Заводская ул. 18 кв 5   или  заводская 18 кв 5",
        "Номер счётчика целиком:  50001  (\"№\", пробелы и нули впереди не важны)",
        "С фамилии, словами целиком, и населённый пункт или улица:  иванов боровляны   или   иванов иван, боровляны заводская" };

    internal readonly TextBox queryBox = new TextBox();
    readonly DataGridView grid = new DataGridView();
    readonly Label status = new Label();
    readonly Panel bottom = new Panel();
    readonly Timer timer = new Timer();
    readonly Timer logTimer = new Timer();   // поиск пишем в отчёт, когда ввод успокоился

    public static void Run(string dir) {
        Run(dir, null);
    }

    // launch - как скрипт подготовил программу (готовая из кэша или собрана, за сколько секунд) - для отчёта
    public static void Run(string dir, string launch) {
        launchInfo = launch;
        Application.EnableVisualStyles();
        Application.Run(new AbonentSearch(Path.Combine(dir, "poisk.txt"), Path.Combine(dir, "nastroiki.txt")));
    }

    static string launchInfo;

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
        partialBox.Text = "Частичное со&впадение";
        partialBox.AutoSize = true;
        partialBox.Margin = new Padding(10, 7, 3, 0);
        partialBox.CheckedChanged += delegate { StartSearch(); };
        new ToolTip().SetToolTip(partialBox, "Без галочки - только точное совпадение.\nС галочкой - ещё номера, которые так начинаются, ФИО по началу слов, адрес по похожим словам.");
        row.Controls.Add(partialBox);
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
        ResultColumns.SelectAllKey(queryBox);
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
        grid.CellFormatting += new DataGridViewCellFormattingEventHandler(OnCellFormatting);
        grid.CellPainting += new DataGridViewCellPaintingEventHandler(OnCellPainting);
        grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0) OpenCard(e.RowIndex); };
        grid.KeyDown += new KeyEventHandler(OnGridKey);

        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 34;
        status.Dock = DockStyle.Fill;
        status.TextAlign = ContentAlignment.MiddleLeft;
        bottom.Controls.Add(status);
        bottom.Controls.Add(BottomButton("В Excel", delegate { ExportResult(null); }));
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

        Shown += delegate { BringUp(); Start(); };
        FormClosing += delegate { cancelScan = true; };
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int cmd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    // Запущенное из свёрнутого или скрытого окна PowerShell окно Windows показывает свёрнутым -
    // разворачиваем и выводим на передний план сами
    void BringUp() {
        try {
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            ShowWindow(Handle, 9);   // SW_RESTORE
            SetForegroundWindow(Handle);
        } catch (Exception) { }
        Activate();
        TopMost = true;
        TopMost = false;
    }

    Button BottomButton(string text, EventHandler click) {
        Button b = new Button();
        b.Text = text;
        b.Dock = DockStyle.Right;
        b.Width = 120;
        b.Click += click;
        return b;
    }

    // найденное - в Excel: те же столбцы, что в списке, цвета строк и примечание, почему строка добавлена;
    // folder - только сохранить файл туда, не открывая Excel (для проверки)
    internal void ExportResult(string folder) {
        SearchResult r = current;
        if (r == null || r.Rows.Count == 0) { MessageBox.Show(this, "Сначала найдите что-нибудь.", Text); return; }
        string[] headers = new string[kinds.Count + 1];
        for (int c = 0; c < kinds.Count; c++) headers[c] = grid.Columns[c].HeaderText;
        headers[kinds.Count] = "Примечание";
        List<string[]> rows = new List<string[]>();
        List<string> styles = new List<string>();
        foreach (int row in r.Rows) {
            string[] v = new string[headers.Length];
            for (int c = 0; c < kinds.Count; c++) v[c] = ResultColumns.Text(r, row, kinds[c]);
            v[kinds.Count] = ResultColumns.NoteText(r, row);
            rows.Add(v);
            styles.Add(ResultColumns.ExcelStyle(r, row));
        }
        try {
            if (folder != null) Spreadsheet.Write(folder, "Поиск_абонента", headers, rows, styles);
            else Spreadsheet.Open("Поиск_абонента", headers, rows, styles);
        }
        catch (Exception ex) { MessageBox.Show(this, "Не удалось выгрузить в Excel:\n" + ex.Message, Text); }
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
        string line = "Поиск по: " + Prompts[(int)r.Mode] + (r.Partial ? " (частичное совпадение)" : "") + (r.Scanned ? " (перебором)" : "") + "; "
            + (r.NoIndex ? "нет подходящего индекса" : r.How) + (r.Query != null && r.Query.LastExplain.Length > 0 ? "; " + r.Query.LastExplain : "")
            + (r.Note.Length > 0 ? "; " + r.Note : "") + " -> найдено " + r.Searched + (r.Added.Count > 0 ? " (+" + (r.Added.Count - r.Related.Count) + " других тарифов, +" + r.Related.Count + " по адресу/ФИО)" : "") + (r.Hidden > 0 ? ", скрыто (другие сельсоветы) " + r.Hidden : "")
            + (r.TooMany ? ", показана часть" : "") + ", " + r.Seconds.ToString("0.00") + " с";
        if (line == lastLogged) return;
        lastLogged = line;
        Report.Search(line);
    }

    string Dir { get { return Path.GetDirectoryName(Path.GetFullPath(settingsFile)); } }

    void Start() {
        Report.Start(Path.Combine(Dir, "otchet.txt"));
        if (launchInfo != null) Report.Line("Подготовка программы: " + launchInfo);
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
                double dirs = (DateTime.Now - t1).TotalSeconds;
                Report.Line("Время открытия: таблица и индексы - " + opened.ToString("0.0") + " с, справочники - " + dirs.ToString("0.0") + " с");
                sel = LoadSelection(b);
            } catch (Exception ex) {
                error = ex.Message;
                Report.Line("ОШИБКА открытия " + path + ": " + ex.Message);
            }
            double seconds = (DateTime.Now - t0).TotalSeconds;
            Post(delegate { AfterOpen(b, sel, error, seconds); });
            if (error != null) return;
            // тарифы (T1, T2...) и доп. сведения - уже когда окно работает: искать можно сразу
            DateTime t2 = DateTime.Now;
            try { b.Extra = ExtraInfo.Open(dirFolder); }
            catch (Exception ex) { Report.Line("Тарифы и доп. сведения не прочитались: " + ex.Message); }
            Report.Line("Тарифы и доп. сведения прочитаны в фоне за " + (DateTime.Now - t2).TotalSeconds.ToString("0.0") + " с");
            Post(delegate { ExtraReady(b); });
            // заранее проверяем индексы по нужным полям - итог в отчёт
            try {
                b.VerifyMain();
                Report.Line("Проверка индексов (ищем по индексу записи из разных мест таблицы):");
                foreach (IndexTag t in b.Tags)
                    if (t.State != 0 || t.Check.Length > 0) Report.Line("    " + t.Label + ": " + t.Check + (t.Learned.Length > 0 ? "; " + t.Learned : ""));
                string learned = b.WeightsReport();
                if (learned.Length > 0) Report.Line(learned);
                if (b.Extra != null)
                    foreach (LiveBase x in b.Extra.Tables) {
                        IndexTag t = x.TagFor("NAB");
                        Report.Line("    доп. сведения " + Path.GetFileName(x.FilePath) + ": индекс по NAB - "
                            + (t == null ? "нет подходящего" : t.Label + ": " + t.Check + (t.Learned.Length > 0 ? "; " + t.Learned : "")));
                    }
                Report.Line("Чем ищется: лицевой - " + TagName(b, "NAB", false) + ", Л/С ЕРИП - " + TagName(b, "LS_ERIP", false)
                    + ", счётчик - " + TagName(b, "NCH", false) + ", снятый счётчик - " + TagName(b, "NSC", false)
                    + ", ФИО - " + TagName(b, "FIO", false) + ", адрес - " + TagName(b, "UL", true));
                if (b.Header.FieldIndex("FIO") >= 0 && b.TagFor("FIO") == null) {
                    List<string> notes = new List<string>();
                    IndexTag fio = null;
                    try { fio = b.FindNameIndexElsewhere(dirFolder, notes); }
                    catch (Exception ex) { notes.Add("поиск прервался: " + ex.Message); }
                    Report.Line("Индекс по ФИО вне таблицы абонентов: " + (fio == null ? "не найден" : "найден " + fio.FilePath));
                    foreach (string n in notes) Report.Line("    " + n);
                    ReportNameIndexes(dirFolder, b.Header.Enc);
                    List<string> code = new List<string>();
                    try { LiveBase.ReportSources(LiveBase.ProgramDirs(dirFolder), code); }
                    catch (Exception ex) { code.Add("просмотр исходников прервался: " + ex.Message); }
                    foreach (string n in code) Report.Line((n.StartsWith("исходников") ? "" : "    ") + n);
                    if (fio != null) {
                        string where = fio.FilePath;
                        Post(delegate { if (b == live) { status.Text = "Найден индекс по ФИО программы базы (" + where + ") - теперь по ФИО ищется сразу."; } });
                    }
                }
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
            if (t.Collated) sb.Append("  сортировка " + t.Collation);
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
                if (d != null && d.Count > 0 && (d.HasTowns || d.Selsovety.Count > 0)) return d;
                Report.Line("    в нём нет населённых пунктов и сельсоветов - ищу справочник заново");
            } catch (Exception ex) {
                Report.Line("    не открылся: " + ex.Message);
            }
        }
        // коды улиц - из записей по всей таблице (в начале таблицы улиц мало)
        PostStatus("Смотрю, как записаны улицы...");
        DbfTable sample = b.SpreadSample(300);
        List<string> codes = StreetDirectory.Codes(sample, sample.FieldIndex("UL"));
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

    // Индекса по ФИО у таблицы абонентов нет. Есть ли он где-то ещё в папке базы (у другой таблицы, в отдельном
    // файле)? Программа базы может искать по ФИО через него - в отчёт, чтобы искать так же.
    static void ReportNameIndexes(string folder, Encoding enc) {
        List<string> found = new List<string>();
        int files = 0;
        try {
            foreach (string f in Directory.GetFiles(folder)) {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".cdx" && ext != ".idx") continue;
                files++;
                try {
                    foreach (IndexTag t in IndexReader.ReadTags(f, enc)) {
                        string e = (t.Expr ?? "").ToUpperInvariant();
                        if (e.Contains("FIO") || e.Contains("FAM") || e.Contains("NAME") || e.Contains("NAIM"))
                            found.Add(Path.GetFileName(f) + ":" + t.Name + " (" + t.Expr + (t.ForText != null && t.ForText.Length > 0 ? " FOR " + t.ForText : "") + ")");
                    }
                } catch (Exception) { }
            }
            Report.Line("Индексы по ФИО во всех файлах папки базы (просмотрено " + files + "): " + (found.Count == 0 ? "нет" : string.Join("; ", found.ToArray())));
            List<string> tables = new List<string>();
            foreach (string f in Directory.GetFiles(folder)) {
                if (string.Compare(Path.GetExtension(f), ".dbf", StringComparison.OrdinalIgnoreCase) != 0) continue;
                try {
                    DbfTable h = DbfTable.LoadHeader(f);
                    int fio = h.FieldIndex("FIO");
                    if (fio >= 0) tables.Add(Path.GetFileName(f) + " (" + h.Expected + " зап." + (h.FieldIndex("NAB") >= 0 ? ", есть NAB" : "") + ")");
                } catch (Exception) { }
            }
            Report.Line("Таблицы с полем FIO: " + (tables.Count == 0 ? "нет" : string.Join(", ", tables.ToArray())));
            List<string> subs = new List<string>();
            foreach (string d in Directory.GetDirectories(folder)) subs.Add(Path.GetFileName(d));
            Report.Line("Вложенные папки: " + (subs.Count == 0 ? "нет" : string.Join(", ", subs.ToArray())));
        } catch (Exception ex) {
            Report.Line("Поиск индексов по ФИО прервался: " + ex.Message);
        }
    }

    // тарифы и доп. сведения дочитаны: появляются столбцы "Т" и "Мощность", найденное ищется заново
    void ExtraReady(LiveBase b) {
        if (b != live || b.Extra == null) return;
        BuildColumns();
        if (current != null && DbfTable.SplitWords(queryBox.Text).Length > 0) StartSearch();
        else Show(current);
    }

    void SetBusy(bool busy, string text) {
        foreach (Control c in bottom.Controls) if (c is Button && c != scanButton) c.Enabled = !busy;
        Cursor = busy ? Cursors.AppStarting : Cursors.Default;
        if (text != null) { status.Text = text; status.Refresh(); }
    }

    AbonentQuery columns;

    List<int> kinds = new List<int>();

    void BuildColumns() {
        columns = new AbonentQuery(DbfTable.CreateLike(live.Header), live.Dir);
        kinds = ResultColumns.Kinds(columns, live.Extra);
        grid.RowCount = 0;
        grid.Columns.Clear();
        foreach (int k in kinds) {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = ResultColumns.Title(columns, k);
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            c.FillWeight = ResultColumns.Weight(live.Header, k);
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
        bool partial = partialBox.Checked;
        status.Text = "Ищу...";
        Cursor = Cursors.AppStarting;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            SearchResult r = b.Search(text, m, ss, partial);
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
        if (again || r.Text != queryBox.Text || r.Mode != mode || r.Partial != partialBox.Checked) { again = false; StartSearch(); return; }
        Show(r);
        if (cardWhenDone) {
            cardWhenDone = false;
            LogSearch();
            // нашёлся один (кроме добавленных тарифов того же счётчика) - сразу его карточка
            if (r.Searched == 1) { for (int i = 0; i < r.Rows.Count; i++) if (!r.Added.ContainsKey(r.Rows[i])) { grid.CurrentCell = grid.Rows[i].Cells[0]; OpenCard(i); break; } }
            else if (r.Rows.Count > 0) grid.Focus();
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
        bool partial = partialBox.Checked;
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
                }, partial);
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
        // как понят адрес: "аг. Боровляны -> Заводская ул. -> дом 18"
        if (r != null && r.Address != null && r.Address.Said.Length > 0) took = "   Искал: " + r.Address.Said + "." + took;
        string tip = r != null && !r.Partial && r.Mode != SearchMode.Address ? " Нужны и похожие номера - поставьте галочку \"Частичное совпадение\"." : "";
        if (r == null)
            status.Text = "Впишите " + Prompts[(int)mode] + ".   " + where;
        else if (r.Error != null)
            status.Text = "Ошибка при поиске: " + r.Error + "   Нажмите \"Отчёт\" и пришлите его.";
        else if (r.NoIndex && r.Mode == SearchMode.Name)
            status.Text = "Допишите населённый пункт или улицу: \"Иванов Боровляны\" - тогда найдётся сразу. По одной фамилии - только перебором всей базы (кнопка справа, это долго): индекса по ФИО в базе нет.";
        else if (r.NoIndex)
            status.Text = "Для поиска по полю \"" + Prompts[(int)r.Mode] + "\" в базе нет подходящего индекса. Можно искать перебором всей базы - это долго (минуты). Кнопка справа.";
        else if (r.Note.Length > 0 && r.Rows.Count == 0)
            status.Text = r.Note + "   " + where;
        else if (r.Rows.Count == 0)
            status.Text = (r.Hidden > 0 ? "В выбранных сельсоветах не найдено, а в других нашлось: " + r.Hidden + " - отметьте их в списке \"Сельсоветы\"."
                : "Ничего не найдено. " + (!r.TooMany && r.Query != null ? r.Query.Miss : "") + tip) + took + "   " + where;
        else
            status.Text = "Найдено: " + r.Searched + (r.TooMany ? " (нашлось очень много, показана часть - уточните запрос)" : "")
                + (r.Added.Count - r.Related.Count > 0 ? ", ещё " + (r.Added.Count - r.Related.Count) + " - другие тарифы тех же счётчиков" : "")
                + (r.Related.Count > 0 ? ", ещё " + r.Related.Count + " - на том же адресе или с тем же ФИО в доме" : "")
                + (r.Added.Count > 0 ? " (их номер красным)" : "")
                + (r.Hidden > 0 ? ", ещё " + r.Hidden + " в других сельсоветах скрыто" : "")
                + ". Enter или двойной щелчок - карточка." + took + "   " + where;
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        SearchResult r = current;
        if (r == null || e.RowIndex >= r.Rows.Count || e.ColumnIndex >= kinds.Count) return;
        e.Value = ResultColumns.Text(r, r.Rows[e.RowIndex], kinds[e.ColumnIndex]);
    }

    string GroupAt(int i) {
        SearchResult r = current;
        return r == null || i < 0 || i >= r.Rows.Count ? null : r.GroupOf(r.Rows[i]);
    }

    void OnCellPainting(object sender, DataGridViewCellPaintingEventArgs e) {
        if (e.RowIndex < 0) return;
        ResultColumns.PaintGroup(grid, e, GroupAt(e.RowIndex - 1), GroupAt(e.RowIndex), GroupAt(e.RowIndex + 1));
    }

    // лицевые тарифов (T1, T2...) - зеленоватым; не искавшиеся тарифы того же счётчика - номер красным
    void OnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
        SearchResult r = current;
        if (r == null || e.RowIndex < 0 || e.RowIndex >= r.Rows.Count || e.ColumnIndex >= kinds.Count) return;
        ResultColumns.Format(r, r.Rows[e.RowIndex], kinds[e.ColumnIndex], e.CellStyle);
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
        int pos = index;
        // PgUp/PgDn в карточке - соседний абонент этого же списка
        CardMove move = delegate(int delta, out SearchResult rr, out int row) {
            rr = null;
            row = -1;
            if (current != r || pos + delta < 0 || pos + delta >= r.Rows.Count) return false;
            pos += delta;
            rr = r;
            row = r.Rows[pos];
            if (pos < grid.RowCount) grid.CurrentCell = grid.Rows[pos].Cells[0];
            return true;
        };
        ResultColumns.ShowCard(this, r, r.Rows[index], live, move);
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
        LastBatch = new BatchSearch(live, selected, mode, columns, kinds);
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

// Столбцы списка найденного: тариф (T1, T2...), поля абонента, мощность (из доп. сведений)
public static class ResultColumns {
    public const int Tariff = -1, Power = -2;
    public static readonly Color TariffColor = Color.FromArgb(214, 245, 214);

    public static List<int> Kinds(AbonentQuery q, ExtraInfo x) {
        List<int> list = new List<int>();
        list.Add(Tariff);
        list.AddRange(q.Columns);
        if (x != null && x.PowerBase != null) list.Add(Power);
        return list;
    }

    public static string Title(AbonentQuery q, int kind) {
        if (kind == Tariff) return "Т";
        if (kind == Power) return "Мощность";
        return q.ColumnTitle(kind);
    }

    public static int Weight(DbfTable h, int kind) {
        if (kind == Tariff) return 30;
        if (kind == Power) return 70;
        string n = h.Fields[kind].Name.ToUpperInvariant();
        return n == "FIO" ? 300 : n == "UL" ? 200 : (n == "NAB" || n == "NCH" || n == "TIPCH") ? 130 : 70;
    }

    public static readonly Color AddedColor = Color.FromArgb(210, 0, 0);

    // строка - другой тариф того же счётчика (её не искали): номер абонента красным
    public static readonly Color OrangeColor = Color.FromArgb(255, 205, 150);

    public static void Format(SearchResult r, int row, int kind, DataGridViewCellStyle style) {
        if (r.Tariff.ContainsKey(row)) style.BackColor = TariffColor;
        else if (r.Orange.ContainsKey(row)) style.BackColor = OrangeColor;
        if (r.Added.ContainsKey(row) && kind >= 0 && string.Compare(r.Table.Fields[kind].Name, "NAB", StringComparison.OrdinalIgnoreCase) == 0) {
            style.ForeColor = AddedColor;
            style.SelectionForeColor = Color.FromArgb(255, 190, 190);
        }
    }

    // Жирная чёрная рамка вокруг группы тарифов одного счётчика (T1, T2, T3), чтобы группы не сливались.
    // prev, cur, next - группа предыдущей, этой и следующей строки (null - строка не в группе)
    public static void PaintGroup(DataGridView g, DataGridViewCellPaintingEventArgs e, string prev, string cur, string next) {
        if (cur == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
        e.Paint(e.ClipBounds, DataGridViewPaintParts.All);
        Rectangle b = e.CellBounds;
        const int w = 3;
        if (prev != cur) e.Graphics.FillRectangle(Brushes.Black, b.Left, b.Top, b.Width, w);
        if (next != cur) e.Graphics.FillRectangle(Brushes.Black, b.Left, b.Bottom - w, b.Width, w);
        if (e.ColumnIndex == 0) e.Graphics.FillRectangle(Brushes.Black, b.Left, b.Top, w, b.Height);
        if (e.ColumnIndex == g.Columns.Count - 1) e.Graphics.FillRectangle(Brushes.Black, b.Right - w, b.Top, w, b.Height);
        e.Handled = true;
    }

    // примечание для выгрузки в Excel (цвета там есть, но почему строка в списке - видно только так)
    public static string NoteText(SearchResult r, int row) {
        string why;
        if (r.Related.TryGetValue(row, out why)) return "добавлен: " + why;
        if (r.Added.ContainsKey(row)) return "добавлен: другой тариф того же счётчика";
        if (r.Orange.ContainsKey(row)) return "на адресе больше одного счётчика";
        return "";
    }

    // цвет строки в Excel: t - тариф (зелёный), o - несколько счётчиков на адресе (оранжевый)
    public static string ExcelStyle(SearchResult r, int row) {
        if (r.Tariff.ContainsKey(row)) return "t";
        if (r.Orange.ContainsKey(row)) return "o";
        return null;
    }

    // Ctrl+A в поле ввода - выделить всё
    public static void SelectAllKey(TextBox box) {
        box.KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.A) { box.SelectAll(); e.Handled = true; e.SuppressKeyPress = true; }
        };
    }

    public static string Text(SearchResult r, int row, int kind) {
        if (kind == Tariff) return r.TariffText(row);
        if (kind == Power) return r.PowerText(row);
        return r.Query.CellText(row, kind);
    }

    // карточка как в программе базы; не открылась - все поля списком
    public static void ShowCard(IWin32Window owner, SearchResult r, int row, LiveBase live, CardMove move) {
        BaseCard card;
        try { card = new BaseCard(r, row, live, move); }
        catch (Exception ex) {
            Report.Line("Карточка не открылась: " + ex.Message);
            ShowList(owner, r, row, live == null ? null : live.Extra);
            return;
        }
        card.Show(owner);
    }

    // все поля списком: сначала адрес по справочнику и тариф, после полей - доп. сведения (Ctrl+D в программе базы)
    public static void ShowList(IWin32Window owner, SearchResult r, int row, ExtraInfo x) {
        List<string> first = new List<string>();
        string[] street = r.Query.CardExtra(row);
        if (street != null) first.AddRange(street);
        int t;
        if (r.Tariff.TryGetValue(row, out t)) { first.Add("Тариф"); first.Add("T" + t + " (лицевые одного счётчика: таблица ABDIF_T)"); }
        string why;
        if (r.Related.TryGetValue(row, out why)) { first.Add("Добавлен к найденному"); first.Add(why); }
        else if (r.Orange.ContainsKey(row)) { first.Add("На адресе"); first.Add("больше одного счётчика (подсвечено оранжевым)"); }
        List<string> after = null;
        int nab = r.Table.FieldIndex("NAB");
        if (x != null && nab >= 0) {
            Cursor old = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try { after = x.Details(DbfTable.Format(r.Table.GetValue(row, nab, null), true)); }
            catch (Exception ex) { after = new List<string>(new string[] { "--- Доп. сведения", "не прочитались: " + ex.Message }); }
            finally { Cursor.Current = old; }
        }
        Card.Show(owner, r.Table, row, first.ToArray(), after);
    }
}

// Выгрузка таблицы строк в Excel ("Таблица XML 2003": все значения - текст, нули в начале не теряются)
public static class Spreadsheet {
    public static void Open(string name, string[] headers, List<string[]> rows) {
        Open(name, headers, rows, null);
    }

    // styles - цвет каждой строки (null - без цвета): t - зелёный, o - оранжевый, y - жёлтый, n - красный
    public static void Open(string name, string[] headers, List<string[]> rows, List<string> styles) {
        string file = Write(Path.GetTempPath(), name, headers, rows, styles);
        try { Process.Start("excel.exe", "\"" + file + "\""); }
        catch { Process.Start(file); }
    }

    public static string Write(string folder, string name, string[] headers, List<string[]> rows, List<string> styles) {
        string file = Path.Combine(folder, name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xml");
        using (StreamWriter w = new StreamWriter(file, false, new UTF8Encoding(false))) {
            w.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            w.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            w.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            w.WriteLine("<Styles><Style ss:ID=\"h\"><Font ss:Bold=\"1\"/></Style>" + Fill("t", "#D6F5D6") + Fill("o", "#FFCD96")
                + Fill("y", "#FFF5C8") + Fill("n", "#FFDCDC") + "</Styles>");
            w.WriteLine("<Worksheet ss:Name=\"" + DbfTable.Xml(name) + "\"><Table>");
            w.Write("<Row>");
            foreach (string h in headers) w.Write("<Cell ss:StyleID=\"h\"><Data ss:Type=\"String\">" + DbfTable.Xml(h) + "</Data></Cell>");
            w.WriteLine("</Row>");
            for (int i = 0; i < rows.Count; i++) {
                string st = styles != null && i < styles.Count && styles[i] != null ? " ss:StyleID=\"" + styles[i] + "\"" : "";
                w.Write("<Row>");
                foreach (string v in rows[i])
                    w.Write(v == null || v.Length == 0 ? "<Cell" + st + "/>" : "<Cell" + st + "><Data ss:Type=\"String\">" + DbfTable.Xml(v) + "</Data></Cell>");
                w.WriteLine("</Row>");
            }
            w.WriteLine("</Table></Worksheet></Workbook>");
        }
        return file;
    }

    static string Fill(string id, string color) {
        return "<Style ss:ID=\"" + id + "\"><Interior ss:Color=\"" + color + "\" ss:Pattern=\"Solid\"/></Style>";
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
    internal readonly CheckBox partialBox = new CheckBox();
    readonly List<Item> items = new List<Item>();
    bool running;
    volatile bool stop;

    class Item {
        public int Line;
        public string Text;
        public string Result;
        public SearchResult Res;  // где лежит найденная запись
        public int Row = -1;      // запись абонента или -1, если не найден
        public int Kind;          // 0 - найден один, 1 - несколько, 2 - не найден
    }

    static readonly string[] ModeNames = { "&Лицевые счета", "&Адреса", "Номера &счётчиков", "&ФИО" };

    readonly List<int> kinds;

    public BatchSearch(LiveBase live, Dictionary<string, bool> selected, SearchMode mode, AbonentQuery query, List<int> kinds) {
        this.kinds = kinds;
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
        partialBox.Text = "Частичное со&впадение";
        partialBox.AutoSize = true;
        partialBox.Location = new Point(x + 20, 8);
        new ToolTip().SetToolTip(partialBox, "Без галочки - только точное совпадение.\nС галочкой - ещё номера, которые так начинаются, ФИО по началу слов, адрес по похожим словам.");
        head.Controls.Add(partialBox);
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
        ResultColumns.SelectAllKey(input);

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
        grid.CellPainting += delegate(object s, DataGridViewCellPaintingEventArgs e) {
            if (e.RowIndex >= 0) ResultColumns.PaintGroup(grid, e, GroupAt(e.RowIndex - 1), GroupAt(e.RowIndex), GroupAt(e.RowIndex + 1));
        };
        grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e) { OpenCard(e.RowIndex); };
        AddColumn("№", 35);
        AddColumn("Искали", 200);
        AddColumn("Результат", 220);
        foreach (int k in kinds) AddColumn(ResultColumns.Title(query, k), k == ResultColumns.Tariff ? 30 : 100);

        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 34;
        status.Dock = DockStyle.Fill;
        status.TextAlign = ContentAlignment.MiddleLeft;
        status.Text = "Найденные - белым, несколько - жёлтым, тарифы (T1, T2...) - зелёным, 2 счётчика на адресе - оранжевым, не найденные - красным; номер красным - добавлено к найденному. Двойной щелчок - карточка.";
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
        bool partial = partialBox.Checked;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            DateTime started = DateTime.Now;
            int single = 0, several = 0, none = 0, done = 0;
            for (int i = 0; i < lines.Count && !stop; i++) {
                SearchResult r = live.Search(lines[i], m, selected, partial);
                List<Item> add = new List<Item>();
                if (r.Rows.Count == 0) {
                    string why = r.Error != null ? "ошибка: " + r.Error
                        : r.NoIndex && m == SearchMode.Name ? "допишите населённый пункт или улицу (Иванов Боровляны)"
                        : r.NoIndex ? "нет индекса для поиска"
                        : r.Hidden > 0 ? "в выбранных сельсоветах нет (в других: " + r.Hidden + ")"
                        : r.Note.Length > 0 ? "не найден: " + r.Note
                        : !r.TooMany && r.Query != null && r.Query.Miss.Length > 0 ? "не найден: " + r.Query.Miss : "не найден";
                    add.Add(NewItem(i + 1, lines[i], why, null, -1, 2));
                    none++;
                } else {
                    int n = r.Searched;
                    string result = n == 1 ? "найден" : "найдено " + n + (r.TooMany ? " (очень много - показана часть)" : "");
                    if (r.Added.Count > r.Related.Count) result += " (+" + (r.Added.Count - r.Related.Count) + " др. тариф)";
                    if (r.Related.Count > 0) result += " (+" + r.Related.Count + " адрес/ФИО)";
                    if (r.Hidden > 0) result += " (+" + r.Hidden + " в других с/с)";
                    if (n == 1) single++; else several++;
                    for (int k = 0; k < r.Rows.Count; k++)
                        add.Add(NewItem(i + 1, lines[i], result, r, r.Rows[k], n == 1 ? 0 : 1));
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
            Post(delegate { Finish(lines, d, s1, s2, s3, seconds, partial); });
        });
        th.IsBackground = true;
        th.Start();
    }

    void Post(MethodInvoker m) {
        try { if (!IsDisposed) BeginInvoke(m); } catch (Exception) { }
    }

    void Finish(List<string> lines, int done, int single, int several, int none, double seconds, bool partial) {
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
        Report.Search("Поиск по списку, в списке: " + ModeNames[(int)mode].Replace("&", "").ToLowerInvariant() + (partial ? " (частичное совпадение)" : "") + "; строк " + lines.Count
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
        if (it.Row < 0 || it.Res == null || column - 3 >= kinds.Count) return "";
        return ResultColumns.Text(it.Res, it.Row, kinds[column - 3]);
    }

    void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e) {
        if (e.RowIndex < items.Count) e.Value = Cell(items[e.RowIndex], e.ColumnIndex);
    }

    // группа тарифов строки итога: своя для каждой строки списка
    string GroupAt(int i) {
        if (i < 0 || i >= items.Count) return null;
        Item it = items[i];
        if (it.Res == null || it.Row < 0) return null;
        string g = it.Res.GroupOf(it.Row);
        return g == null ? null : it.Line + "|" + g;
    }

    void OnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
        if (e.RowIndex < 0 || e.RowIndex >= items.Count) return;
        Item it = items[e.RowIndex];
        int kind = it.Kind;
        if (kind == 2) e.CellStyle.BackColor = Color.FromArgb(255, 220, 220);
        else if (kind == 1) e.CellStyle.BackColor = Color.FromArgb(255, 245, 200);
        if (it.Res != null && it.Row >= 0 && e.ColumnIndex >= 3 && e.ColumnIndex - 3 < kinds.Count)
            ResultColumns.Format(it.Res, it.Row, kinds[e.ColumnIndex - 3], e.CellStyle);
        else if (it.Res != null && it.Res.Tariff.ContainsKey(it.Row)) e.CellStyle.BackColor = ResultColumns.TariffColor;
    }

    internal void OpenCardAt(int index) { OpenCard(index); }

    void OpenCard(int index) {
        if (index < 0 || index >= items.Count || items[index].Row < 0) return;
        Item it = items[index];
        int pos = index;
        // PgUp/PgDn в карточке - следующий найденный абонент итога
        CardMove move = delegate(int delta, out SearchResult rr, out int row) {
            rr = null;
            row = -1;
            int p = pos;
            do {
                p += delta;
                if (p < 0 || p >= items.Count) return false;
            } while (items[p].Row < 0 || items[p].Res == null);
            pos = p;
            rr = items[p].Res;
            row = items[p].Row;
            if (p < grid.RowCount) grid.CurrentCell = grid.Rows[p].Cells[0];
            return true;
        };
        ResultColumns.ShowCard(this, it.Res, it.Row, live, move);
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
        List<string> styles = new List<string>();
        foreach (Item it in items) {
            string st = it.Res != null && it.Row >= 0 ? ResultColumns.ExcelStyle(it.Res, it.Row) : null;
            styles.Add(st ?? (it.Kind == 2 ? "n" : it.Kind == 1 ? "y" : null));
        }
        try { Spreadsheet.Open("Поиск_по_списку", headers, rows, styles); }
        catch (Exception ex) { MessageBox.Show(this, "Не удалось выгрузить в Excel:\n" + ex.Message, Text); }
    }
}
'@
# Программа собирается из этого скрипта ОДИН раз после обновления и сохраняется готовой (файл poisk_*.dll
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
$name = "poisk_$stamp.dll"
$cacheDirs = @()
if ($env:LOCALAPPDATA) { $cacheDirs += (Join-Path $env:LOCALAPPDATA 'poisk-abonenta') }
if ($env:PUBLIC) { $cacheDirs += (Join-Path $env:PUBLIC 'poisk-abonenta') }
$cacheDirs += (Join-Path $here 'kesh-programmy')

function Test-Ready { return ($null -ne ('AbonentSearch' -as [type])) }

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
        $splash.Text = 'Поиск абонента'
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
            Get-ChildItem -Path $d -Filter 'poisk_*.dll' -ErrorAction SilentlyContinue |
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
    [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n" + ($errors -join "`n`n") + "`n`nСфотографируйте это окно и пришлите.", 'Поиск абонента - ошибка')
    exit
}

try {
    [AbonentSearch]::Run($here, $launch)
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("$_", 'Поиск абонента - ошибка')
}
