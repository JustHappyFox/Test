# Режим Excel в движке базы (LiveBase, StreetDirectory, FieldNames) - одинаково для «Поиска» и «Договора» (у «Договора» копия движка).
# python3 patch_engine.py prosmotr-bazy/prosmotr.ps1 [dogovor/dogovor.ps1]
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from pyedit import Doc
for path in sys.argv[1:]:
    d = Doc(path)
    if 'public DbfTable Mem;' in d.t: print(path, 'уже'); continue
    d.rep('''    public static LiveBase Open(string path) {
        LiveBase b = new LiveBase();''', '''    // Режим Excel: своя копия выгрузки быта (byt.dbf) целиком в памяти, поиск - AbonentQuery по всей таблице (индексов нет)
    public DbfTable Mem;
    int memLimit = FetchLimit;      // сколько найденных показывать (больше - «показана часть»)

    public static LiveBase OpenMemory(string path, StreetDirectory dir) {
        LiveBase b = new LiveBase();
        b.FilePath = path;
        b.Mem = DbfTable.Load(path, false);
        b.Mem.FilePath = path;
        b.Header = b.Mem;
        b.Sample = b.Mem;
        b.RecordCount = b.Mem.Count;
        b.Dir = dir;
        return b;
    }

    // поиск в памяти (режим Excel): как перебор, только таблица уже прочитана; ФИО с адресом - «Иванов Боровляны»
    SearchResult SearchMemory(string text, SearchMode mode, Dictionary<string, bool> selsovety, int limit, bool partial) {
        DateTime started = DateTime.Now;
        memLimit = mode == SearchMode.Address ? Math.Max(limit, SeekLimit) : Math.Max(limit, FetchLimit);
        SearchResult res = MemoryOnce(text, mode, selsovety, partial, null, null);
        if (mode == SearchMode.Name && res.Rows.Count == 0 && res.Hidden == 0) {
            string[] w = DbfTable.SplitWords(text);
            for (int k = w.Length - 1; k >= 1; k--) {
                string name = string.Join(" ", w, 0, k), place = string.Join(" ", w, k, w.Length - k);
                SearchResult a = MemoryOnce(text, mode, selsovety, partial, name, place);
                if (a.Rows.Count > 0 || a.Hidden > 0) { res = a; break; }
            }
        }
        res.How = "выгрузка в памяти (" + Mem.Count + " записей)" + (res.How.Length > 0 ? "; " + res.How : "");
        res.Seconds = (DateTime.Now - started).TotalSeconds;
        return res;
    }

    SearchResult MemoryOnce(string text, SearchMode mode, Dictionary<string, bool> selsovety, bool partial, string nameText, string addressText) {
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Partial = partial;
        res.NameText = nameText;
        res.AddressText = addressText;
        if (mode == SearchMode.Address && !partial && Dir != null) {
            res.Address = Dir.Match(text);
            if (res.Address.Problem != null) { res.Note = res.Address.Problem; return Finish(res, null, selsovety, started); }
        }
        return Finish(res, Mem, selsovety, started);
    }

    public static LiveBase Open(string path) {
        LiveBase b = new LiveBase();''')
    d.rep('''        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Partial = partial;
        if (mode == SearchMode.Name && names != null && TagFor("FIO") == null) {''', '''        if (Mem != null) return SearchMemory(text, mode, selsovety, limit, partial);
        DateTime started = DateTime.Now;
        SearchResult res = new SearchResult();
        res.Text = text;
        res.Mode = mode;
        res.Partial = partial;
        if (mode == SearchMode.Name && names != null && TagFor("FIO") == null) {''')
    d.rep('''    public bool NamesNeeded { get { return Header.FieldIndex("FIO") >= 0 && TagFor("FIO") == null; } }''',
          '''    public bool NamesNeeded { get { return Mem == null && Header.FieldIndex("FIO") >= 0 && TagFor("FIO") == null; } }''')
    # тот же адрес / то же ФИО в доме - в памяти без индекса
    d.rep('''        IndexTag t = TagFor("UL", true);
        if (t == null || t.Terms.Count < 2 || t.TermFields[1] != "DOM") return;
        List<int> searched = new List<int>();''', '''        IndexTag t = Mem != null ? null : TagFor("UL", true);
        if (Mem == null && (t == null || t.Terms.Count < 2 || t.TermFields[1] != "DOM")) return;
        List<int> searched = new List<int>();''')
    d.rep('''        Seeker s = new Seeker();
        try {
            foreach (List<int> list in houses.Values) {
                int r0 = list[0];
                s.Seek(t, KeyFor(t, new string[] { Field(tb, r0, ulF), Field(tb, r0, domF) }, false));
            }
        } finally {
            s.Close();
        }
        List<int> missing = new List<int>();
        foreach (int rec in s.Found.Keys) if (!rowOfRec.ContainsKey(rec)) missing.Add(rec);''', '''        List<int> foundRecs = new List<int>();
        if (Mem != null) {
            // режим Excel: вся таблица в памяти - дома найденных ищем перебором строк
            for (int row = 0; row < tb.Count; row++) if (houses.ContainsKey(HouseKey(tb, row))) foundRecs.Add(tb.RecNos[row]);
        } else {
            Seeker s = new Seeker();
            try {
                foreach (List<int> list in houses.Values) {
                    int r0 = list[0];
                    s.Seek(t, KeyFor(t, new string[] { Field(tb, r0, ulF), Field(tb, r0, domF) }, false));
                }
            } finally {
                s.Close();
            }
            foundRecs.AddRange(s.Found.Keys);
        }
        List<int> missing = new List<int>();
        foreach (int rec in foundRecs) if (!rowOfRec.ContainsKey(rec)) missing.Add(rec);''')
    d.rep('''        int added = 0;
        foreach (int rec in s.Found.Keys) {''', '''        int added = 0;
        foreach (int rec in foundRecs) {''')
    # справочник улиц из копии выгрузки
    d.rep('''    public bool HasTowns {''', '''    // режим Excel: справочник улиц своей копии выгрузки (код -> н.п., улица, сельсовет; VygruzkaByt.ReadStreets)
    public static StreetDirectory FromCopy(string file) {
        StreetDirectory dir = new StreetDirectory();
        dir.TableName = Path.GetFileName(file);
        dir.KeyField = "UL";
        foreach (KeyValuePair<string, string[]> kv in VygruzkaByt.ReadStreets(file)) {
            string code = NormalizeCode(kv.Key), town = kv.Value[0].Trim(), own = kv.Value[1].Trim(), ss = kv.Value[2].Trim();
            StringBuilder display = new StringBuilder(town);
            if (own.Length > 0) display.Append((display.Length > 0 ? ", " : "") + own);
            if (ss.Length > 0) {
                string l = ss.ToLowerInvariant();
                display.Append((display.Length > 0 ? ", " : "") + ss + (l.Contains("с/с") || l.Contains("сельсовет") ? "" : " с/с"));
            }
            dir.names[code] = display.ToString();
            dir.lowerNames[code] = display.ToString().ToLowerInvariant().Replace('ё', 'е');
            dir.ssOf[code] = CleanSs(ss);
            dir.townOf[code] = town;
            dir.streetOf[code] = own;
        }
        return dir;
    }

    public bool HasTowns {''')
    # подписи полей копии выгрузки (NP, ULICA, SELSOVET...) - как в заголовке выгрузки
    d.rep('''    public static string Label(string field) {''', '''    public static Dictionary<string, string> Excel;      // режим Excel: поле копии -> заголовок выгрузки

    public static string Label(string field) {
        if (Excel != null) { string x; if (Excel.TryGetValue(field, out x) && x.Length > 0) return x; }''')
    d.rep('            int ul = part.FieldIndex("UL");\n            foreach (int row in rows) {', '            if (Mem != null && part == Mem && rows.Count > memLimit) { rows.RemoveRange(memLimit, rows.Count - memLimit); res.TooMany = true; }\n            int ul = part.FieldIndex("UL");\n            foreach (int row in rows) {')
    d.save()
    print(path, 'ок')
