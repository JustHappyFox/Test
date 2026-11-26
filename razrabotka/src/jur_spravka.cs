// ==== БЛОК СПРАВКИ ЮР (исходник - razrabotka/src/jur_spravka.cs; здесь не править - вставляется скриптом inject.py) ====
// 26.11: «Справка юр» (справка для производства расчетов юр лица: расторжение, ввод в расчёты...; шаблон jur_spravka.docx, по фото)
// и «Поиск по юр базе». Юр база - выгрузка юр лиц (таблица точек учёта FTU: KP договор, NAB карточка, NP наименование договора,
// NAB_NAIM наименование абонента с адресом, NAIM точка учёта, NST № счётчика, TIPSCH, GODEN, AMPER, VOLT, FAZA, KTT, RAEN).
// В справке - два места подтягивания: договор + карточка (все счётчики карточки) или № счётчика (этот счётчик и его карточка).

public class JurRow {
    public string Kp = "", Nab = "", Contract = "", Abon = "", Point = "", Meter = "", Type = "", Poverka = "", Amper = "", Volt = "",
        Faza = "", Ktt = "", Raen = "";
    public string Hay = "";       // для поиска по словам: всё строчными

    public string Card { get { return Kp + (Nab.Length > 0 ? "/" + Nab : ""); } }
}

public static class JurIndex {
    static List<JurRow> rows;
    static DbfTable from;

    // все строки юр базы (читается один раз; заменили файл выгрузки - заново)
    public static List<JurRow> All() {
        JurBase b = JurBase.OpenFtu();
        if (rows != null && from == b.Ftu) return rows;
        DbfTable t = b.Ftu;
        List<JurRow> r = new List<JurRow>(t.Count);
        int kp = t.FieldIndex("KP"), nab = t.FieldIndex("NAB"), np = t.FieldIndex("NP"), nn = t.FieldIndex("NAB_NAIM"), naim = t.FieldIndex("NAIM"),
            nst = t.FieldIndex("NST"), tip = t.FieldIndex("TIPSCH"), god = t.FieldIndex("GODEN"), amp = t.FieldIndex("AMPER"), volt = t.FieldIndex("VOLT"),
            faza = t.FieldIndex("FAZA"), ktt = t.FieldIndex("KTT"), raen = t.FieldIndex("RAEN");
        for (int i = 0; i < t.Count; i++) {
            JurRow x = new JurRow();
            x.Kp = Num(F(t, i, kp)); x.Nab = Num(F(t, i, nab)); x.Contract = Quotes(F(t, i, np)); x.Abon = F(t, i, nn); x.Point = F(t, i, naim);
            x.Meter = Num(F(t, i, nst)); x.Type = F(t, i, tip); x.Poverka = F(t, i, god); x.Amper = F(t, i, amp); x.Volt = Num(F(t, i, volt));
            x.Faza = Num(F(t, i, faza)); x.Ktt = Num(F(t, i, ktt)); x.Raen = F(t, i, raen);
            if (x.Poverka == "-") x.Poverka = "";
            if (x.Kp.Length == 0 && x.Meter.Length == 0) continue;
            x.Hay = Low(x.Contract + " " + x.Abon + " " + x.Point + " " + x.Raen + " " + x.Type + " " + x.Meter + " " + x.Kp + "/" + x.Nab);
            r.Add(x);
        }
        rows = r;
        from = t;
        return r;
    }

    static string F(DbfTable t, int row, int f) { return f < 0 ? "" : DbfTable.Format(t.GetValue(row, f, null), true).Trim(); }

    // «11118.00» -> «11118»
    static string Num(string s) {
        s = (s ?? "").Trim();
        Match m = Regex.Match(s, @"^(\d+)[.,]0*$");
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
        string d = Regex.Replace(no ?? "", @"\D", "");
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
        Match m = Regex.Match(num, @"^(\d{1,7})\s*[/\\\-\s]\s*(\d{1,6})$");
        if (m.Success) { kp = m.Groups[1].Value; nab = m.Groups[2].Value; }
        else if (Regex.IsMatch(num, @"^\d+$")) digits = num;
        else if (num.Length > 0) text = (num + " " + text).Trim();
        List<string> words = new List<string>();
        foreach (string w in Low(text).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) words.Add(w);
        if (kp.Length == 0 && digits.Length == 0 && words.Count == 0) return r;
        foreach (JurRow x in All()) {
            if (kp.Length > 0 && !(SameNum(x.Kp, kp) && SameNum(x.Nab, nab))) continue;
            // одни цифры: договор с таким номером или счётчик с этими цифрами (от 4)
            if (digits.Length > 0 && !SameNum(x.Kp, digits) && !(digits.Length >= 4 && x.Meter.Contains(digits.TrimStart('0').Length > 0 ? digits.TrimStart('0') : digits))) continue;
            bool all = true;
            foreach (string w in words) if (x.Hay.IndexOf(w, StringComparison.Ordinal) < 0) { all = false; break; }
            if (!all) continue;
            r.Add(x);
            if (r.Count >= max) break;
        }
        return r;
    }

    // «Нераспроданные квартиры п.Макетино,ул.Полевая,34» -> «п. Макетино, ул. Полевая, 34»; район - спереди
    public static string AddressOf(JurRow x) {
        string a = x.Abon ?? "";
        Match m = Regex.Match(a, @"(?:^|[\s,])((?:п|д|аг|г|гп|г\.п|пос|дер|ул|пр|пер)\.\s*\S.*)$", RegexOptions.IgnoreCase);
        string tail = m.Success ? m.Groups[1].Value : "";
        tail = Regex.Replace(tail, @"\b(п|д|аг|г|гп|пос|дер|ул|пр|пер)\.\s*", "$1. ");
        tail = Regex.Replace(tail, @"\s*,\s*", ", ").Trim().TrimEnd(',', '.');
        string r = x.Raen.Trim();
        if (r.Length > 0) r = Regex.Replace(r, @"\s+район$", " р-н", RegexOptions.IgnoreCase);
        return (r.Length > 0 && tail.Length > 0 && tail.IndexOf(r, StringComparison.OrdinalIgnoreCase) < 0 ? r + ", " : "") + tail;
    }

    // строка прибора учёта для справки: точка|тип|№|U|I|показания|пов. кв.|пов. год|Т/Т|К расч
    public static string PuLine(JurRow x, string zeros) {
        string u = x.Volt == "400" || x.Faza == "3" ? "3х230/400" : x.Volt.Length > 0 ? x.Volt : "";
        string i = Regex.Replace(x.Amper ?? "", @"^\s*(\d+(?:[.,]\d+)?)\s*[-/]\s*(\d+(?:[.,]\d+)?)\s*$", "$1($2)");
        string kv = "", god = "";
        Match p = Regex.Match(x.Poverka ?? "", @"(\d{4})\D*(\d)?");
        if (p.Success) { god = p.Groups[1].Value; kv = p.Groups[2].Value; }
        double k;
        bool tt = double.TryParse(x.Ktt, NumberStyles.Float, CultureInfo.InvariantCulture, out k) && k > 1;
        string tts = tt ? (k * 5).ToString(CultureInfo.InvariantCulture) + "/5" : "-";
        return string.Join("|", new string[] { x.Point, x.Type, x.Meter, u, i, zeros, kv, god, tts, tt ? x.Ktt : "1" });
    }
}

public class JurSpravkaForm : Form {
    readonly string here, templateFile;
    readonly Lists lists;
    readonly Memory memory;
    readonly Dictionary<string, Control> inp = new Dictionary<string, Control>();
    readonly List<Card> cards = new List<Card>();
    readonly Panel content = new Panel();
    readonly Label head = new Label(), headSub = new Label(), status = new Label(), found = new Label();
    readonly ToolTip tip = new ToolTip();
    RowList pu, att;
    public StartMenu Home;
    public Form Back;
    bool loading, dirty;
    string savedName, autoAtt = "";
    static readonly string[] Needed = { "kp", "abonent", "pu", "inzh", "nach" };

    public bool Unsaved { get { return dirty; } }
    public static string Folder(string here) { return Path.Combine(here, "Справки юр"); }
    public static string DataFolder(string here) { return Path.Combine(Folder(here), "данные"); }

    public JurSpravkaForm(string here) {
        Load += delegate { St.Dress(this); };
        this.here = here;
        templateFile = Path.Combine(here, "jur_spravka.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = Memory.For(Path.Combine(here, "nastroiki.txt"));
        Text = "Справка юр - для производства расчетов";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1180, wa.Width), Math.Min(840, wa.Height));
        MinimumSize = new Size(860, 560);
        Load += delegate { St.Fit(this); };
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
        FillLists();
        Clear();
        Shown += delegate { LayoutCards(); FillLists(); };
        Activated += delegate { if (!lists.Changed()) return; try { lists.Load(); FillLists(); } catch { } };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.S) { e.SuppressKeyPress = true; Save(true, true); }
            else if (e.Control && e.KeyCode == Keys.P) { e.SuppressKeyPress = true; Preview(); }
            else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; GoBack(); }
            else if (e.KeyCode == Keys.Enter && ActiveControl != null && (ActiveControl.Name == "kp" || ActiveControl.Name == "nab")) { e.SuppressKeyPress = true; ByCard(); }
            else if (e.KeyCode == Keys.Enter && ActiveControl != null && ActiveControl.Name == "meter") { e.SuppressKeyPress = true; ByMeter(); }
        };
        FormClosing += delegate(object s, FormClosingEventArgs e) {
            if (Home != null && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; GoBack(); }
        };
    }

    TextBox T(string key, string cue) { TextBox t = Ui.Text(1); Reg(key, t); if (cue != null) t.HandleCreated += delegate { Ui.Cue(t, cue); }; return t; }
    Combo C(string key, string cue) { Combo c = new Combo(true, false); Reg(key, c); if (cue != null) c.HandleCreated += delegate { Ui.Cue(c, cue); }; return c; }

    void Reg(string key, Control c) {
        inp[key] = c;
        c.Name = key;
        EventHandler h = delegate {
            if (loading) return;
            Crash.Mark("справка юр, поле " + key);
            dirty = true;
            if (key == "prichina") AttFor();
            RefreshState();
        };
        if (c is ComboBox) { ((ComboBox)c).TextChanged += h; ((ComboBox)c).SelectedIndexChanged += h; }
        else if (c is DateBox) ((DateBox)c).ValueChanged += h;
        else c.TextChanged += h;
        if (c is TextBox || c is ComboBox) {
            c.Enter += delegate { c.BackColor = St.Focus; };
            c.Leave += delegate { c.BackColor = c.Enabled ? St.Input : St.Disabled; };
        }
    }

    public string V(string key) {
        if (key == "pu") return pu.Value;
        if (key == "att") return att.Value;
        Control c;
        if (!inp.TryGetValue(key, out c)) return "";
        if (c is DateBox) return ((DateBox)c).Text2;
        return c.Text.Trim();
    }

    public void SetV(string key, string v) {
        v = v ?? "";
        if (key == "pu") { pu.Value = v; return; }
        if (key == "att") { att.Value = v; return; }
        Control c;
        if (!inp.TryGetValue(key, out c)) return;
        if (c is DateBox) {
            DateTime d;
            if (DateTime.TryParseExact(v, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) ((DateBox)c).Value = d;
        } else if (c is Combo) ((Combo)c).Value = v;
        else c.Text = v;
    }

    Cell F(string caption, Control c, int chars) { Cell x = Cell.F(caption, c, chars == 0 ? 4 : 1); x.W = chars; return x; }

    void BuildCards() {
        Card c = new Card("Из юр базы", "впишите договор и карточку ИЛИ № счётчика и нажмите Enter (или «Найти») - остальное подтянется из юр базы", 4);
        cards.Add(c);
        Button byCard = Ui.Button("Найти", false);
        byCard.Click += delegate { ByCard(); };
        Button byMeter = Ui.Button("Найти", false);
        byMeter.Click += delegate { ByMeter(); };
        tip.SetToolTip(byCard, "Все счётчики карточки договора - из юр базы (Enter)");
        tip.SetToolTip(byMeter, "Этот счётчик и его договор, карточка - из юр базы (Enter)");
        c.Grid.Row(F("Договор №", T("kp", "11118"), 10), F("Карточка №", T("nab", "26"), 8), F(" ", byCard, 8),
            F("или № счётчика", T("meter", "80536165"), 14), F(" ", byMeter, 8));
        found.Font = St.Small;
        found.ForeColor = St.Hint;
        c.Grid.Row(F(null, found, 0));

        c = new Card("Абонент и объект", "наименование, объект и адрес - из юр базы, можно поправить", 4);
        cards.Add(c);
        c.Grid.Row(F("Наименование плательщика (абонент)", C("abonent", "ООО «...»"), 0));
        c.Grid.Row(F("Объект", T("obekt", "«...»"), 0));
        c.Grid.Row(F("Адрес", T("adres", "Минский р-н, д. ..., ул. ..., 7"), 0));

        c = new Card("Причина и приложения", "что требуется (строка в справке жирным) и что прилагается", 4);
        cards.Add(c);
        c.Grid.Row(F("Причина (строка в справке)", C("prichina", null), 0));
        att = new RowList(new string[] { "К справке прилагается (строка на документ, «+» - ещё)" }, new int[] { 0 }, new bool[] { true }, null);
        att.Name = "att";
        att.Changed += delegate { if (loading) return; dirty = true; RefreshState(); };
        att.Resized += delegate { LayoutCards(); };
        c.Grid.Row(F(null, att, 0));

        c = new Card("Приборы учёта", "строка на счётчик, «+» - ещё; показания - нулями по разрядности или как есть", 4);
        cards.Add(c);
        pu = new RowList(new string[] { "Точка учёта", "Тип счётчика", "№ счётчика", "U, В", "I, А", "Показания", "Пов. кв.", "Пов. год", "Т/Т", "К расч" },
            new int[] { 10, 0, 11, 9, 7, 10, 5, 5, 8, 5 },
            new bool[] { false, true, false, true, true, false, true, false, false, false },
            new string[] { null, null, null, "3х230/400", "5(100)", "000000,00", null, null, "-", "1" });
        pu.Name = "pu";
        pu.Changed += delegate { if (loading) return; dirty = true; RefreshState(); };
        pu.Resized += delegate { LayoutCards(); };
        c.Grid.Row(F(null, pu, 0));
        DateBox dpu = new DateBox();
        Reg("data_pu", dpu);
        c.Grid.Row(F("Данные приборов учёта на", dpu, 14));

        c = new Card("Подписи и даты", "инженер, начальник, дата подписей", 4);
        cards.Add(c);
        DateBox dp = new DateBox();
        Reg("data_podp", dp);
        c.Grid.Row(F("Должность инженера", C("inzh_dolzh", null), 16), F("Инженер", C("inzh", null), 18), F("Должность начальника", C("nach_dolzh", null), 16),
            F("Начальник", C("nach", null), 18), F("Дата подписей", dp, 14));
    }

    void FillLists() {
        bool was = loading;
        loading = true;
        try {
            Items("abonent", memory.Hist("юр справка абонент"));
            Items("prichina", With(lists.Get("Юр справка: причина"), memory.Hist("юр справка причина")));
            Items("inzh_dolzh", lists.Get("Юр справка: должность инженера"));
            Items("inzh", lists.Get("Оформил"));
            List<string> n = new List<string>();
            foreach (string[] s in lists.Signers()) n.Add(s[0]);
            Items("nach", n);
            Items("nach_dolzh", lists.Get("Должность начальника"));
            att.SetItems(0, lists.Get("Юр справка: приложение"));
            pu.SetItems(1, Types());
            pu.SetItems(3, lists.Get("БС: напряжение"));
            pu.SetItems(4, lists.Get("БС: ток счётчика"));
            pu.SetItems(6, new List<string>(new string[] { "1", "2", "3", "4" }));
        } finally { loading = was; }
    }

    List<string> Types() {
        List<string> r = new List<string>();
        foreach (MeterType m in lists.Meters()) { string t = (m.Model + " " + m.Case + " " + m.Mod).Trim(); if (!r.Contains(t)) r.Add(t); }
        return r;
    }

    void Items(string key, List<string> items) { Combo c = inp[key] as Combo; if (c != null) c.SetItems(items); }

    static List<string> With(List<string> a, List<string> b) {
        List<string> r = new List<string>(a);
        foreach (string s in b) if (!r.Contains(s)) r.Add(s);
        return r;
    }

    static string First(List<string> l) { return l.Count > 0 ? l[0] : ""; }

    // ---------- подтягивание из юр базы ----------

    void Found(string text, bool ok) { found.Text = text; found.ForeColor = ok ? St.Good : St.Bad; LayoutCards(); }

    List<JurRow> Safe(string what, string a, string b) {
        try { return what == "card" ? JurIndex.Card(a, b) : JurIndex.ByMeter(a); }
        catch (Exception ex) { Found("Юр база не открылась: " + ex.Message, false); return null; }
    }

    // договор + карточка -> все счётчики карточки; без карточки: одна карточка - она, несколько - выбрать
    public void ByCard() {
        string kp = V("kp"), nab = V("nab");
        if (kp.Length == 0) { Found("Впишите номер договора (и карточку).", false); return; }
        List<JurRow> r = Safe("card", kp, nab);
        if (r == null) return;
        if (r.Count == 0) { Found("В юр базе нет договора " + kp + (nab.Length > 0 ? " с карточкой " + nab : "") + ".", false); return; }
        if (nab.Length == 0) {
            List<string> cardsOf = new List<string>();
            foreach (JurRow x in r) if (!cardsOf.Contains(x.Nab)) cardsOf.Add(x.Nab);
            if (cardsOf.Count > 1) { PickCard(r, cardsOf); return; }
        }
        FromRows(r, "договор " + r[0].Card);
    }

    void PickCard(List<JurRow> r, List<string> cardsOf) {
        ContextMenuStrip m = new ContextMenuStrip();
        foreach (string n in cardsOf) {
            JurRow first = null;
            int cnt = 0;
            foreach (JurRow x in r) if (x.Nab == n) { if (first == null) first = x; cnt++; }
            string label = "карточка " + n + " - " + first.Abon + " (счётчиков: " + cnt + ")";
            string nn = n;
            m.Items.Add(label.Length > 110 ? label.Substring(0, 110) + "…" : label, null, delegate { SetV("nab", nn); ByCard(); });
        }
        Found("У договора " + V("kp") + " карточек: " + cardsOf.Count + " - выберите карточку.", true);
        Control k = inp["nab"];
        m.Show(k, new Point(0, k.Height));
    }

    // № счётчика -> его договор и карточка, этот счётчик
    public void ByMeter() {
        string no = V("meter");
        if (no.Length == 0) { Found("Впишите номер счётчика.", false); return; }
        List<JurRow> r = Safe("meter", no, null);
        if (r == null) return;
        if (r.Count == 0) { Found("Счётчика " + no + " в юр базе нет.", false); return; }
        if (r.Count > 1) {
            ContextMenuStrip m = new ContextMenuStrip();
            foreach (JurRow x in r) {
                JurRow xx = x;
                string label = "№ " + x.Meter + " - договор " + x.Card + ", " + x.Abon + ", " + x.Point;
                m.Items.Add(label.Length > 120 ? label.Substring(0, 120) + "…" : label, null, delegate { FromRows(new List<JurRow>(new JurRow[] { xx }), "счётчик № " + xx.Meter); });
            }
            Found("Номер подходит к " + r.Count + " счётчикам - выберите.", true);
            Control k = inp["meter"];
            m.Show(k, new Point(0, k.Height));
            return;
        }
        FromRows(r, "счётчик № " + r[0].Meter);
    }

    // строки юр базы -> поля справки (договор, карточка, абонент, объект, адрес, счётчики)
    public void FromRows(List<JurRow> r, string what) {
        JurRow a = r[0];
        loading = true;
        try {
            SetV("kp", a.Kp);
            SetV("nab", a.Nab);
            if (r.Count == 1) SetV("meter", a.Meter);
            SetV("abonent", a.Contract);
            string ob = UdalenieForm.ObjectOf(a.Abon);
            SetV("obekt", ob.Length > 0 ? "«" + ob.Trim('«', '»') + "»" : "");
            SetV("adres", JurIndex.AddressOf(a));
            List<string> lines = new List<string>();
            foreach (JurRow x in r) lines.Add(JurIndex.PuLine(x, ZerosFor(x.Type)));
            SetV("pu", string.Join("\n", lines.ToArray()));
        } finally { loading = false; }
        dirty = true;
        Found("Из юр базы (" + what + "): " + a.Contract + "; счётчиков: " + r.Count + ". Проверьте объект, адрес и показания.", true);
        head.Text = "Справка юр " + V("kp") + (V("nab").Length > 0 ? "/" + V("nab") : "");
        LayoutCards();
        RefreshState();
    }

    // нули по разрядности типа (справочник [БС: разрядность счётчика]), иначе 000000,00
    string ZerosFor(string type) {
        string key = MeterKey(type), r = "6+2";
        foreach (string s in lists.Get("БС: разрядность счётчика")) {
            int e = s.IndexOf('=');
            if (e > 0 && key.StartsWith(MeterKey(s.Substring(0, e)))) { r = s.Substring(e + 1).Trim(); break; }
        }
        string[] p = r.Replace(" ", "").Split('+', ',');
        int x, y = 0;
        if (!int.TryParse(p[0], out x) || x <= 0 || x > 12) { x = 6; y = 2; }
        else if (p.Length > 1) int.TryParse(p[1], out y);
        return new string('0', x) + (y > 0 ? "," + new string('0', Math.Min(y, 4)) : "");
    }

    static string MeterKey(string t) {
        StringBuilder b = new StringBuilder();
        foreach (char ch in (t ?? "").ToUpperInvariant()) {
            if (ch == ' ' || ch == '-' || ch == '.' || ch == '_') continue;
            int i = "ABCEHKMOPTXY".IndexOf(ch);
            b.Append(i >= 0 ? "АВСЕНКМОРТХУ"[i] : ch);
        }
        return b.ToString();
    }

    // причина -> что прилагается (из справочника «причина = приложение; приложение»), пока приложения не правили руками
    void AttFor() {
        string p = V("prichina");
        string add = "";
        foreach (string s in lists.Get("Юр справка: что прилагается к причине")) {
            int e = s.IndexOf('=');
            if (e > 0 && p.ToLowerInvariant().Contains(s.Substring(0, e).Trim().ToLowerInvariant())) { add = s.Substring(e + 1).Trim(); break; }
        }
        if (att.Value.Trim() != autoAtt.Trim() && att.Value.Trim().Length > 0) return;
        List<string> l = new List<string>();
        foreach (string x in add.Split(';')) if (x.Trim().Length > 0) l.Add(x.Trim());
        bool was = loading;
        loading = true;
        try { att.Value = string.Join("\n", l.ToArray()); } finally { loading = was; }
        autoAtt = att.Value;
        LayoutCards();
    }

    // ---------- окно ----------

    Control BuildNav() {
        Panel nav = new Panel();
        nav.Dock = DockStyle.Left;
        nav.Width = St.NavWidth;
        nav.BackColor = St.Nav;
        Label title = new Label();
        title.Text = "Справка юр";
        title.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
        title.ForeColor = Color.White;
        title.SetBounds(22, 22, 220, 34);
        Label sub = new Label();
        sub.Text = "для производства расчетов";
        sub.Font = St.Small;
        sub.ForeColor = St.NavText;
        sub.SetBounds(24, 58, 220, 20);
        nav.Controls.Add(title);
        nav.Controls.Add(sub);
        NavItem back = new NavItem("← Назад", false);
        back.SetBounds(0, 100, nav.Width, back.Height);
        back.Click += delegate { GoBack(); };
        nav.Controls.Add(back);
        int y = 150;
        foreach (Card c in cards) {
            NavItem n = new NavItem(c.Name, false);
            n.Mark = St.TintOf(c.Name).Mark;
            n.SetBounds(0, y, nav.Width, n.Height);
            Card card = c;
            n.Click += delegate { content.AutoScrollPosition = new Point(0, Math.Max(0, card.Panel.Top - content.AutoScrollPosition.Y - 4)); };
            nav.Controls.Add(n);
            y += n.Height;
        }
        NavItem openItem = new NavItem("Открыть сохранённую…", true);
        openItem.Click += delegate { OpenSaved(); };
        NavItem folder = new NavItem("Папка справок юр", true);
        folder.Click += delegate {
            try { Directory.CreateDirectory(Folder(here)); Process.Start("explorer.exe", "\"" + Folder(here) + "\""); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text); }
        };
        NavItem dirs = new NavItem("Справочники (списки)", true);
        dirs.Click += delegate { try { Process.Start("notepad.exe", "\"" + lists.File + "\""); } catch (Exception ex) { MessageBox.Show(this, ex.Message, Text); } };
        NavItem[] low = { openItem, folder, dirs };
        foreach (NavItem n in low) { n.Anchor = AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right; nav.Controls.Add(n); }
        nav.Resize += delegate {
            int yy = nav.ClientSize.Height - 16;
            for (int i = low.Length - 1; i >= 0; i--) { yy -= low[i].Height; low[i].SetBounds(0, yy, nav.Width, low[i].Height); }
        };
        St.FitNav(nav);
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
        p.BackColor = St.Surface;
        p.Paint += delegate(object s, PaintEventArgs e) { using (Pen pen = new Pen(St.Border)) e.Graphics.DrawLine(pen, 0, 0, p.Width, 0); };
        Button clear = Ui.Button("Очистить", false);
        clear.SetBounds(24, 14, 110, 38);
        clear.Click += delegate {
            if (dirty && MessageBox.Show(this, "Очистить справку? Введённое пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Clear();
        };
        Button look = Ui.Button("Предпросмотр", false);
        look.Size = new Size(140, 38);
        look.Click += delegate { Preview(); };
        Button save = Ui.Button("Сохранить справку", true);
        save.Size = new Size(210, 40);
        save.Click += delegate { Save(true, true); };
        tip.SetToolTip(look, "Посмотреть справку в Word, ничего не сохраняя (Ctrl+P)");
        tip.SetToolTip(save, "Сохранить в папку «Справки юр» и открыть в Word (Ctrl+S)");
        status.AutoSize = false;
        status.Font = St.Base;
        status.TextAlign = ContentAlignment.MiddleLeft;
        p.Controls.Add(clear);
        p.Controls.Add(look);
        p.Controls.Add(save);
        p.Controls.Add(status);
        p.Resize += delegate {
            save.Location = new Point(p.ClientSize.Width - 24 - save.Width, 13);
            look.Location = new Point(save.Left - 10 - look.Width, 14);
            status.SetBounds(clear.Right + 16, 14, Math.Max(10, look.Left - clear.Right - 32), 38);
            status.Font = status.Width < 300 ? St.Small : St.Base;
        };
        return p;
    }

    void LayoutCards() {
        content.SuspendLayout();
        int full = content.Width - SystemInformation.VerticalScrollBarWidth;
        int w = Math.Max(600, Math.Min(full - 32, 1100));
        int x = Math.Max(16, (full - w) / 2);
        int y = 4 + content.AutoScrollPosition.Y;
        foreach (Card c in cards) {
            int h = c.Layout(w);
            c.Panel.Location = new Point(x + content.AutoScrollPosition.X, y);
            y += h + 10;
        }
        content.AutoScrollMinSize = new Size(0, y - content.AutoScrollPosition.Y + 6);
        content.ResumeLayout();
    }

    public void GoBack() {
        Hide();
        if (Home != null) Home.Return(Back);
        else if (Back != null && !Back.IsDisposed) { Back.Show(); Back.Activate(); }
    }

    public void Clear() {
        loading = true;
        try {
            foreach (KeyValuePair<string, Control> kv in inp) if (!(kv.Value is DateBox)) SetV(kv.Key, "");
            pu.Value = "";
            att.Value = "";
            ((DateBox)inp["data_pu"]).Value = DateTime.Today;
            ((DateBox)inp["data_podp"]).Value = DateTime.Today;
            SetV("prichina", memory.Get("ЮрСправкаПричина", First(lists.Get("Юр справка: причина"))));
            SetV("inzh_dolzh", memory.Get("ЮрСправкаДолжностьИнженера", First(lists.Get("Юр справка: должность инженера"))));
            string eng = memory.Get("ЮрСправкаИнженер", memory.Get("СправкаИнженер", ""));
            SetV("inzh", lists.Get("Оформил").Contains(eng) ? eng : First(lists.Get("Оформил")));
            List<string> names = new List<string>();
            foreach (string[] sg in lists.Signers()) names.Add(sg[0]);
            SetV("nach", memory.Get("ЮрСправкаНачальник", memory.Get("СправкаНачальник", First(names))));
            SetV("nach_dolzh", memory.Get("ЮрСправкаДолжностьНачальника", First(lists.Get("Должность начальника"))));
            savedName = null;
            autoAtt = "";
        } finally { loading = false; }
        AttFor();
        dirty = false;
        found.Text = "Договор + карточка - все счётчики карточки; № счётчика - этот счётчик, его договор и карточка.";
        found.ForeColor = St.Hint;
        head.Text = "Новая справка юр";
        headSub.ForeColor = St.Hint;
        headSub.Text = "Справка для производства расчетов юр лица (расторжение, ввод в расчёты...)";
        RefreshState();
    }

    void OpenSaved() {
        string dir = DataFolder(here);
        if (!Directory.Exists(dir) || Directory.GetFiles(dir, "*.txt").Length == 0) { MessageBox.Show(this, "Сохранённых справок юр пока нет.", Text); return; }
        if (dirty && MessageBox.Show(this, "Справка не сохранена. Открыть другую? Введённое здесь пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        using (SavedDialog dlg = new SavedDialog(dir, 4)) {
            if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Chosen == null) return;
            try { LoadSaved(dlg.Chosen); }
            catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть:\n" + ex.Message, Text); }
        }
    }

    public void LoadSaved(string file) {
        Dictionary<string, string> d = DogovorForm.ReadData(file);
        string name = Path.GetFileNameWithoutExtension(file);
        Clear();
        loading = true;
        try {
            foreach (string k in new List<string>(inp.Keys)) { string v; if (d.TryGetValue(k, out v)) SetV(k, v); }
            string pv;
            if (d.TryGetValue("pu", out pv)) SetV("pu", pv);
            if (d.TryGetValue("att", out pv)) { SetV("att", pv); autoAtt = ""; }
            savedName = name;
        } finally { loading = false; }
        dirty = false;
        LayoutCards();
        head.Text = "Справка юр «" + name + "»";
        headSub.Text = "открыта сохранённая - исправьте и сохраните заново";
        RefreshState();
    }

    public Dictionary<string, string> Compose(List<Dictionary<string, string>> rows) {
        Dictionary<string, string> v = new Dictionary<string, string>();
        v["ДОГНОМ"] = V("kp") + (V("nab").Length > 0 ? "/" + V("nab") : "");
        v["АБОНЕНТ"] = V("abonent");
        v["ОБЪЕКТ"] = V("obekt");
        v["АДРЕС"] = V("adres").TrimEnd('.');
        string p = V("prichina");
        v["ПРИЧИНА"] = p.Length > 0 && !p.EndsWith(".") ? p + "." : p;
        List<string> a = new List<string>();
        for (int i = 0; i < att.Count; i++) if (att.Cell(i, 0).Length > 0) a.Add(att.Cell(i, 0).TrimEnd('.') + ".");
        v["ПРИЛОЖЕНИЕ"] = string.Join("\n", a.ToArray());
        v["ДАТАПУ"] = V("data_pu");
        v["ДАТАПОДП"] = V("data_podp");
        v["ИНЖДОЛЖН"] = V("inzh_dolzh").Length > 0 ? V("inzh_dolzh") : "Инженер по РПЭ";
        v["ИНЖЕНЕР"] = V("inzh");
        v["НАЧДОЛЖН"] = V("nach_dolzh").Length > 0 ? V("nach_dolzh") : "Начальник ГРП";
        v["НАЧАЛЬНИК"] = V("nach");
        for (int i = 0; i < pu.Count; i++) {
            if (pu.Empty(i) && i > 0) continue;
            Dictionary<string, string> r = new Dictionary<string, string>();
            string kv = pu.Cell(i, 6);
            r["ПУН"] = (rows.Count + 1).ToString();
            r["ПУТОЧКА"] = pu.Cell(i, 0);
            r["ПУТИП"] = pu.Cell(i, 1);
            r["ПУНОМЕР"] = pu.Cell(i, 2);
            r["ПУВ"] = pu.Cell(i, 3);
            r["ПУА"] = pu.Cell(i, 4);
            r["ПУПОКАЗ"] = pu.Cell(i, 5);
            r["ПУКВ"] = kv.Length > 0 && char.IsDigit(kv[kv.Length - 1]) ? kv + " кв." : kv;
            r["ПУГОДГ"] = pu.Cell(i, 7).Length > 0 ? pu.Cell(i, 7) + " г." : "";
            r["ПУТТ"] = pu.Cell(i, 8);
            r["ПУК"] = pu.Cell(i, 9);
            rows.Add(r);
        }
        return v;
    }

    List<string> Missing() {
        List<string> m = new List<string>();
        foreach (string k in Needed) if (k == "pu" ? pu.Cell(0, 2).Length == 0 && pu.Cell(0, 1).Length == 0 : V(k).Length == 0) m.Add(k);
        return m;
    }

    static string Caption(string k) {
        switch (k) {
            case "kp": return "договор";
            case "abonent": return "наименование плательщика";
            case "pu": return "прибор учёта";
            case "inzh": return "инженер";
            case "nach": return "начальник";
            default: return k;
        }
    }

    void RefreshState() {
        List<string> m = Missing();
        if (m.Count == 0) { status.Text = savedName != null && !dirty ? "Справка сохранена" : "Всё основное заполнено"; status.ForeColor = St.Good; return; }
        List<string> names = new List<string>();
        foreach (string k in m) names.Add(Caption(k));
        status.Text = "Не заполнено: " + string.Join(", ", names.ToArray());
        status.ForeColor = St.Bad;
    }

    // «Дог 11118-26 расторжение 26.11.2026»
    public string FileName() {
        string p = V("prichina").ToLowerInvariant();
        string what = p.Contains("расторг") ? "расторжение" : p.Contains("ввести") || p.Contains("допуск") ? "ввод" : p.Contains("исключ") ? "исключение" : "справка";
        string n = ("Дог " + V("kp") + (V("nab").Length > 0 ? "-" + V("nab") : "") + " " + what + " " + V("data_podp")).Trim();
        foreach (char ch in Path.GetInvalidFileNameChars()) n = n.Replace(ch, '_');
        while (n.Contains("  ")) n = n.Replace("  ", " ");
        return V("kp").Length > 0 ? n : "Справка юр " + DateTime.Now.ToString("yyyy-MM-dd HH-mm");
    }

    bool HasTemplate() {
        if (File.Exists(templateFile)) return true;
        MessageBox.Show(this, "Нет файла шаблона:\n" + templateFile + "\n\nОн должен лежать рядом с программой (есть в архиве программы).", Text,
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return false;
    }

    void Fill(string output) {
        List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
        Dictionary<string, string> v = Compose(rows);
        if (rows.Count == 0) rows.Add(new Dictionary<string, string>());
        string tmp = output + ".pu.tmp";
        try {
            DocxTemplate.ExpandRows(templateFile, tmp, "ПУТОЧКА", null, rows, new string[0], new string[0]);
            for (int i = 0; i < rows.Count; i++) foreach (KeyValuePair<string, string> kv in rows[i]) v[kv.Key + "_" + (i + 1)] = kv.Value;
            DocxTemplate.Fill(tmp, output, v);
        } finally { try { File.Delete(tmp); } catch { } }
    }

    public string Preview() {
        if (!HasTemplate()) return null;
        try {
            string f = DocPreview.NewFile(here, FileName());
            Fill(f);
            DocPreview.Show(this, f);
            return f;
        } catch (Exception ex) {
            MessageBox.Show(this, "Предпросмотр не получился:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }
    }

    public string Save(bool openWord, bool ask) {
        List<string> m = Missing();
        if (ask && m.Count > 0) {
            List<string> names = new List<string>();
            foreach (string k in m) names.Add("  • " + Caption(k));
            if (MessageBox.Show(this, "Не заполнено:\n" + string.Join("\n", names.ToArray()) + "\n\nВсё равно сохранить справку?", "Проверка перед сохранением",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
        }
        if (!HasTemplate()) return null;
        string name = FileName();
        string dir = Folder(here);
        string docx = Path.Combine(dir, name + ".docx");
        try {
            Directory.CreateDirectory(dir);
            if (ask && savedName != name && File.Exists(docx) && MessageBox.Show(this, "Справка «" + name + "» уже есть. Заменить её?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
            string tmp = docx + ".new";
            Fill(tmp);
            if (File.Exists(docx)) File.Delete(docx);
            File.Move(tmp, docx);
            Directory.CreateDirectory(DataFolder(here));
            StringBuilder sb = new StringBuilder("# Данные справки юр для программы «Оформление договора»\r\n");
            sb.Append("сохранён=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
            foreach (string k in inp.Keys) sb.Append(k).Append('=').Append(Memory.Escape(V(k))).Append("\r\n");
            sb.Append("pu=").Append(Memory.Escape(pu.Value)).Append("\r\n");
            sb.Append("att=").Append(Memory.Escape(att.Value)).Append("\r\n");
            File.WriteAllText(Path.Combine(DataFolder(here), name + ".txt"), sb.ToString(), new UTF8Encoding(true));
        } catch (IOException ex) {
            MessageBox.Show(this, "Справка не сохранена (открыта в Word?):\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        memory.Remember("юр справка абонент", V("abonent"));
        memory.Remember("юр справка причина", V("prichina"));
        memory.Values["ЮрСправкаПричина"] = V("prichina");
        memory.Values["ЮрСправкаДолжностьИнженера"] = V("inzh_dolzh");
        memory.Values["ЮрСправкаИнженер"] = V("inzh");
        memory.Values["ЮрСправкаНачальник"] = V("nach");
        memory.Values["ЮрСправкаДолжностьНачальника"] = V("nach_dolzh");
        memory.Save();
        savedName = name;
        dirty = false;
        head.Text = "Справка юр «" + name + "»";
        RefreshState();
        if (openWord) try { DocPreview.Open(docx); } catch { }
        return docx;
    }

    public bool AskSaveBeforeExit() {
        if (!dirty) return true;
        DialogResult r = MessageBox.Show(this, "Справка юр не сохранена. Сохранить её?\n\nДа - сохранить и выйти\nНет - выйти без сохранения\nОтмена - вернуться", Text,
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) return Save(false, true) != null;
        return true;
    }

    public string StatusText { get { return status.Text; } }
}

// Поиск по юр базе: номер (договор, договор/карточка, № счётчика) и слова (название, адрес, точка учёта); двойной щелчок или
// «Справка юр…» - справка по карточке выбранной строки
public class JurSearchForm : Form {
    readonly TextBox num = Ui.Text(1), words = Ui.Text(1);
    readonly ListView list = new ListView();
    readonly Label info = new Label();
    readonly Timer delay = new Timer();
    public StartMenu Home;
    List<JurRow> shown = new List<JurRow>();

    public JurSearchForm() {
        Load += delegate { St.Dress(this); };
        Text = "Поиск по юр базе";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1240, wa.Width), Math.Min(780, wa.Height));
        MinimumSize = new Size(Math.Min(860, wa.Width), Math.Min(500, wa.Height));
        Load += delegate { St.Fit(this); };
        KeyPreview = true;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

        Panel top = new Panel();
        top.Dock = DockStyle.Top;
        top.Height = 92;
        top.BackColor = St.Surface;
        Label t = new Label();
        t.Text = "Поиск по юр базе";
        t.Font = St.Big;
        t.ForeColor = St.Text;
        t.AutoSize = true;
        t.Location = new Point(20, 8);
        Label l1 = Ui.Caption("Договор, договор/карточка (11118/26) или № счётчика");
        Label l2 = Ui.Caption("Название, адрес, точка учёта - слова через пробел");
        Ui.Cue(num, "11118 или 11118/26 или 80536165");
        Ui.Cue(words, "Тестстрой Полевая 34");
        top.Controls.Add(t);
        top.Controls.Add(l1);
        top.Controls.Add(l2);
        top.Controls.Add(num);
        top.Controls.Add(words);
        top.Resize += delegate {
            int w = (top.ClientSize.Width - 60) / 2;
            l1.SetBounds(20, 40, w, 18);
            num.SetBounds(20, 58, w, 26);
            l2.SetBounds(40 + w, 40, w, 18);
            words.SetBounds(40 + w, 58, w, 26);
        };

        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        foreach (object[] c in new object[][] {
                new object[] { "Договор", 70 }, new object[] { "Карт.", 50 }, new object[] { "Наименование договора", 200 }, new object[] { "Абонент (объект, адрес)", 260 },
                new object[] { "Точка учёта", 140 }, new object[] { "Тип", 120 }, new object[] { "№ счётчика", 100 }, new object[] { "Поверка", 70 },
                new object[] { "I, А", 60 }, new object[] { "U", 45 }, new object[] { "Ктт", 45 }, new object[] { "Район", 110 } })
            list.Columns.Add((string)c[0], (int)c[1]);
        list.DoubleClick += delegate { OpenSpravka(); };

        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 60;
        bottom.BackColor = St.Surface;
        Button back = Ui.Button("Назад", false);
        back.Click += delegate { GoBack(); };
        Button copy = Ui.Button("Копировать строку", false);
        copy.Click += delegate { CopyRow(); };
        Button spr = Ui.Button("Справка юр по карточке", true);
        spr.Click += delegate { OpenSpravka(); };
        info.Font = St.Small;
        info.ForeColor = St.Hint;
        info.AutoEllipsis = true;
        bottom.Controls.Add(back);
        bottom.Controls.Add(copy);
        bottom.Controls.Add(spr);
        bottom.Controls.Add(info);
        bottom.Resize += delegate {
            back.SetBounds(20, 12, 100, 36);
            spr.SetBounds(bottom.ClientSize.Width - 20 - 230, 11, 230, 38);
            copy.SetBounds(spr.Left - 10 - 170, 12, 170, 36);
            info.SetBounds(back.Right + 14, 12, Math.Max(10, copy.Left - back.Right - 28), 36);
        };
        Controls.Add(list);
        Controls.Add(top);
        Controls.Add(bottom);

        delay.Interval = 350;
        delay.Tick += delegate { delay.Stop(); Search(); };
        num.TextChanged += delegate { delay.Stop(); delay.Start(); };
        words.TextChanged += delegate { delay.Stop(); delay.Start(); };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; GoBack(); }
            else if (e.KeyCode == Keys.Enter && (ActiveControl == num || ActiveControl == words)) { e.SuppressKeyPress = true; delay.Stop(); Search(); if (list.Items.Count > 0) { list.Focus(); list.Items[0].Selected = true; } }
            else if (e.KeyCode == Keys.Enter && ActiveControl == list) { e.SuppressKeyPress = true; OpenSpravka(); }
        };
        FormClosing += delegate(object s, FormClosingEventArgs e) {
            if (Home != null && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; GoBack(); }
        };
        Shown += delegate { Hint(); num.Focus(); };
    }

    public string Num { get { return num.Text; } set { num.Text = value; } }
    public string Words { get { return words.Text; } set { words.Text = value; } }
    public int Found { get { return shown.Count; } }

    void Hint() {
        try { int n = JurIndex.All().Count; info.Text = "В юр базе точек учёта: " + n + ". Впишите номер или слова - список сразу."; info.ForeColor = St.Hint; }
        catch (Exception ex) { info.Text = "Юр база не открылась: " + ex.Message; info.ForeColor = St.Bad; }
    }

    public void Search() {
        list.BeginUpdate();
        try {
            list.Items.Clear();
            shown = new List<JurRow>();
            if (num.Text.Trim().Length == 0 && words.Text.Trim().Length == 0) { Hint(); return; }
            try { shown = JurIndex.Find(num.Text, words.Text, 500); }
            catch (Exception ex) { info.Text = "Юр база не открылась: " + ex.Message; info.ForeColor = St.Bad; return; }
            Dictionary<string, bool> dogs = new Dictionary<string, bool>();
            foreach (JurRow x in shown) {
                ListViewItem it = new ListViewItem(new string[] { x.Kp, x.Nab, x.Contract, x.Abon, x.Point, x.Type, x.Meter, x.Poverka, x.Amper, x.Volt, x.Ktt, x.Raen });
                it.Tag = x;
                list.Items.Add(it);
                dogs[x.Kp] = true;
            }
            info.Text = shown.Count == 0 ? "Ничего не нашлось - проверьте номер или уберите лишнее слово." :
                "Найдено точек учёта: " + shown.Count + (shown.Count >= 500 ? " (показаны первые 500)" : "") + ", договоров: " + dogs.Count +
                ". Двойной щелчок - справка юр по карточке.";
            info.ForeColor = shown.Count == 0 ? St.Bad : St.Hint;
        } finally { list.EndUpdate(); }
    }

    JurRow Selected { get { return list.SelectedItems.Count > 0 ? list.SelectedItems[0].Tag as JurRow : list.Items.Count == 1 ? list.Items[0].Tag as JurRow : null; } }

    void CopyRow() {
        JurRow x = Selected;
        if (x == null) return;
        try { Clipboard.SetText("Договор " + x.Card + "; " + x.Contract + "; " + x.Abon + "; " + x.Point + "; " + x.Type + " № " + x.Meter); } catch { }
    }

    void OpenSpravka() {
        JurRow x = Selected;
        if (x == null) { info.Text = "Выберите строку в списке."; info.ForeColor = St.Bad; return; }
        if (Home != null) Home.OpenJurSpravka(this, x.Kp, x.Nab);
    }

    public void GoBack() {
        Hide();
        if (Home != null) Home.Return(null);
    }
}
// ==== КОНЕЦ БЛОКА СПРАВКИ ЮР ====
