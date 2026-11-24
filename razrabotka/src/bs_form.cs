// ==== БЛОК ФОРМЫ БС (исходник - razrabotka/src/bs_form.cs; здесь не править - вставляется скриптом inject.py) ====
// Пакет документов юр лица (объект вроде базовой станции): справка для производства расчетов, акт приемки приборов учёта,
// доп. соглашение (Приложение 3), наряд на включение, «потери» (Приложение 8 - технологический расход по линии 0,4 кВ).
// Одна форма: данные вводятся один раз и расходятся по всем документам (шаблоны bs_*.docx рядом с программой).
// Пакет - в папку «БС\<название пакета>» (все 5 файлов или выбранные), данные - «БС\данные\<название>.txt» (открыть заново,
// пломбы - в отчёты пломб).

// Таблица строк с «+»: у каждой строки свои поля (текст или список с вариантами); значение - строки через \n, поля через |
public class RowList : Panel {
    public readonly string[] Caps;
    readonly int[] widths;            // ширина поля в знаках; 0 - делят остаток
    readonly bool[] combo;
    readonly List<string>[] items;
    readonly string[] defaults;
    class Row { public Control[] Cells; public Button Remove; }
    readonly List<Row> rows = new List<Row>();
    readonly Label[] heads;
    readonly Button add = new Button();
    bool quiet, ready;                // ready - конструктор закончен (раскладка раньше - не нужна)
    public event EventHandler Changed;
    public event EventHandler Resized;
    public delegate void CellEdit(int row, int col);
    public event CellEdit Edited;      // поле строки поменяли руками (для подстановок вроде r0 по марке кабеля)

    public RowList(string[] caps, int[] widths, bool[] combo, string[] defaults) {
        Caps = caps;
        this.widths = widths;
        this.combo = combo;
        this.defaults = defaults;
        items = new List<string>[caps.Length];
        for (int i = 0; i < caps.Length; i++) items[i] = new List<string>();
        BackColor = Color.Transparent;
        heads = new Label[caps.Length];
        for (int i = 0; i < caps.Length; i++) {
            Label l = Ui.Caption(caps[i]);
            l.AutoEllipsis = true;
            heads[i] = l;
            Controls.Add(l);
        }
        add.Text = "+";
        add.Font = St.Bold;
        add.FlatStyle = FlatStyle.Flat;
        add.FlatAppearance.BorderColor = St.Accent;
        add.ForeColor = St.Link;
        add.BackColor = St.Surface;
        add.UseVisualStyleBackColor = false;
        add.Cursor = Cursors.Hand;
        add.Click += delegate { AddRow(true); Relayout(); Fire(); rows[rows.Count - 1].Cells[0].Focus(); };
        Controls.Add(add);
        AddRow(true);
        ready = true;
        PerformLayout();
    }

    int RowH { get { return Ui.InputHeightCached; } }
    int CapH { get { return St.Small.Height + 3; } }
    public int NeededHeight { get { return CapH + rows.Count * (RowH + 6) - 6; } }
    public int Count { get { return rows.Count; } }

    public void SetItems(int col, List<string> list) {
        items[col] = new List<string>(list);
        quiet = true;
        try {
            foreach (Row r in rows) {
                Combo c = r.Cells[col] as Combo;
                if (c == null) continue;
                string t = c.Text;
                c.SetItems(items[col]);
                if (c.Text != t) c.Text = t;
            }
        } finally { quiet = false; }
    }

    Row AddRow(bool withDefaults) {
        Row r = new Row();
        r.Cells = new Control[Caps.Length];
        for (int i = 0; i < Caps.Length; i++) {
            Control c;
            if (combo[i]) { Combo cb = new Combo(true, false); cb.SetItems(items[i]); c = cb; }
            else c = Ui.Text(1);
            if (withDefaults && defaults != null && i < defaults.Length && defaults[i] != null) c.Text = defaults[i];
            Control cc = c;
            int col = i;
            Row row = r;
            EventHandler h = delegate {
                if (quiet) return;
                if (Edited != null) Edited(rows.IndexOf(row), col);
                Fire();
            };
            c.TextChanged += h;
            if (c is ComboBox) ((ComboBox)c).SelectedIndexChanged += h;
            cc.Enter += delegate { cc.BackColor = St.Focus; };
            cc.Leave += delegate { cc.BackColor = St.Input; };
            r.Cells[i] = c;
            Controls.Add(c);
        }
        r.Remove = new Button();
        r.Remove.Text = "×";
        r.Remove.Font = St.Bold;
        r.Remove.FlatStyle = FlatStyle.Flat;
        r.Remove.FlatAppearance.BorderColor = St.InputBorder;
        r.Remove.ForeColor = St.Hint;
        r.Remove.BackColor = St.Surface;
        r.Remove.UseVisualStyleBackColor = false;
        r.Remove.Cursor = Cursors.Hand;
        r.Remove.TabStop = false;
        Row me = r;
        r.Remove.Click += delegate { RemoveRow(me); };
        Controls.Add(r.Remove);
        rows.Add(r);
        return r;
    }

    void RemoveRow(Row r) {
        if (rows.Count == 1) {
            quiet = true;
            try { for (int i = 0; i < r.Cells.Length; i++) r.Cells[i].Text = defaults != null && i < defaults.Length && defaults[i] != null ? defaults[i] : ""; }
            finally { quiet = false; }
        } else {
            rows.Remove(r);
            foreach (Control c in r.Cells) { Controls.Remove(c); c.Dispose(); }
            Controls.Remove(r.Remove);
            r.Remove.Dispose();
            Relayout();
        }
        Fire();
    }

    void Relayout() {
        PerformLayout();
        if (Resized != null) Resized(this, EventArgs.Empty);
    }

    void Fire() { if (!quiet && Changed != null) Changed(this, EventArgs.Empty); }

    protected override void OnLayout(LayoutEventArgs e) {
        base.OnLayout(e);
        if (!ready) return;
        int h = RowH, gap = 6, w = ClientSize.Width, bw = h;
        int fixedW = 0, fills = 0;
        int[] px = new int[Caps.Length];
        for (int i = 0; i < Caps.Length; i++) {
            if (widths[i] > 0) { px[i] = (int)(widths[i] * Grid.CharW + 10); fixedW += px[i]; } else fills++;
        }
        // не влезает - узкие поля ужимаются поровну, широкому (тип счётчика, марка кабеля) остаётся не меньше 150 точек
        int avail = w - gap * Caps.Length - bw * 2 - gap, minFill = 150 * fills;
        if (fixedW > 0 && fixedW > avail - minFill) {
            double k = Math.Max(0.4, (avail - minFill) / (double)fixedW);
            fixedW = 0;
            for (int i = 0; i < Caps.Length; i++) if (widths[i] > 0) { px[i] = Math.Max(28, (int)(px[i] * k)); fixedW += px[i]; }
        }
        int rest = avail - fixedW;
        for (int i = 0; i < Caps.Length; i++) if (widths[i] <= 0) px[i] = Math.Max(60, fills > 0 ? rest / fills : 60);
        int x = 0;
        for (int i = 0; i < Caps.Length; i++) { heads[i].SetBounds(x, 0, px[i], CapH); x += px[i] + gap; }
        for (int r = 0; r < rows.Count; r++) {
            int y = CapH + r * (h + gap);
            x = 0;
            for (int i = 0; i < Caps.Length; i++) { rows[r].Cells[i].SetBounds(x, y, px[i], h); x += px[i] + gap; }
            rows[r].Remove.SetBounds(x, y, bw, h);
            if (r == rows.Count - 1) add.SetBounds(x + bw + gap, y, bw, h);
        }
    }

    public string Cell(int row, int col) { return row < rows.Count ? rows[row].Cells[col].Text.Trim() : ""; }

    public void SetCell(int row, int col, string v) {
        if (row >= rows.Count) return;
        Control c = rows[row].Cells[col];
        if (c.Text == v) return;
        quiet = true;
        try { c.Text = v; } finally { quiet = false; }
    }

    // строка пустая - все поля пусты или как по умолчанию
    public bool Empty(int row) {
        for (int i = 0; i < Caps.Length; i++) {
            string d = defaults != null && i < defaults.Length && defaults[i] != null ? defaults[i] : "";
            if (Cell(row, i).Length > 0 && Cell(row, i) != d) return false;
        }
        return true;
    }

    public string Value {
        get {
            List<string> lines = new List<string>();
            for (int r = 0; r < rows.Count; r++) {
                string[] c = new string[Caps.Length];
                for (int i = 0; i < Caps.Length; i++) c[i] = Cell(r, i).Replace("|", "/").Replace("\n", " ");
                lines.Add(string.Join("|", c));
            }
            return string.Join("\n", lines.ToArray());
        }
        set {
            quiet = true;
            try {
                while (rows.Count > 1) {
                    Row r = rows[rows.Count - 1];
                    rows.Remove(r);
                    foreach (Control c in r.Cells) { Controls.Remove(c); c.Dispose(); }
                    Controls.Remove(r.Remove);
                    r.Remove.Dispose();
                }
                string[] lines = (value ?? "").Split('\n');
                for (int k = 0; k < lines.Length; k++) {
                    if (k > 0) AddRow(false);
                    string[] c = lines[k].Split('|');
                    for (int i = 0; i < Caps.Length; i++)
                        rows[k].Cells[i].Text = i < c.Length ? c[i] : (k == 0 && lines[k].Length == 0 && defaults != null && i < defaults.Length && defaults[i] != null ? defaults[i] : "");
                }
            } finally { quiet = false; }
            Relayout();
        }
    }
}

public class BsForm : Form {
    readonly string here;
    readonly Lists lists;
    readonly Memory memory;
    readonly Dictionary<string, Control> inp = new Dictionary<string, Control>();
    readonly List<Card> cards = new List<Card>();
    readonly Panel content = new Panel();
    readonly Label head = new Label(), headSub = new Label(), status = new Label(), shemaNote = new Label();
    readonly ToolTip tip = new ToolTip();
    public StartMenu Home;
    public Form Back;
    bool loading, dirty;
    string savedName;
    RowList pu, moshn, linii;
    Seals seals;
    readonly Dictionary<string, string> autos = new Dictionary<string, string>();
    static readonly string[] Composed = { "adres", "adres_kr", "obekt_full", "obekt_adres", "potr", "plomby", "minus", "paket", "inzh_fio" };
    static readonly string[] AddrKeys = { "obl", "rayon", "selsovet", "np_tip", "np", "ulica_tip", "ulica", "dom" };
    // документы пакета: ключ галочки, шаблон, имя файла в папке пакета
    public static readonly string[][] Docs = {
        new string[] { "doc_spr", "bs_spravka.docx", "Справка для производства расчетов" },
        new string[] { "doc_akt", "bs_akt.docx", "Акт приемки приборов учета" },
        new string[] { "doc_dop", "bs_dopsogl.docx", "Доп. соглашение" },
        new string[] { "doc_nar", "bs_naryad.docx", "Наряд" },
        new string[] { "doc_pot", "bs_poteri.docx", "Потери (Приложение 8)" } };
    static readonly string[] Months = { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };
    // обязательные поля: ключ, подпись, раздел
    readonly List<string[]> fieldList = new List<string[]>();
    Dictionary<string, bool> requiredSet;
    static readonly string[] DefaultRequired = { "dog_nom", "abonent", "obekt", "np", "dom", "pitanie", "liniya", "mesto", "inzh", "nach", "pu", "moshn", "linii", "shema" };

    public bool Unsaved { get { return dirty; } }
    public static string Folder(string here) { return Path.Combine(here, "БС"); }
    public static string DataFolder(string here) { return Path.Combine(Folder(here), "данные"); }

    public BsForm(string here) {
        Load += delegate { St.Dress(this); };
        this.here = here;
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = Memory.For(Path.Combine(here, "nastroiki.txt"));
        Text = "БС - пакет документов юр лица";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1240, wa.Width), Math.Min(880, wa.Height));
        MinimumSize = new Size(900, 560);
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
        LoadRequired();
        Clear();
        Shown += delegate { LayoutCards(); FillLists(); };
        Activated += delegate {
            if (!lists.Changed()) return;
            try { lists.Load(); FillLists(); } catch { }
        };
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.Control && e.KeyCode == Keys.S) { e.SuppressKeyPress = true; Save(true, true); }
            else if (e.Control && e.KeyCode == Keys.P) { e.SuppressKeyPress = true; Preview(); }
            else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; GoBack(); }
            else if (e.KeyCode == Keys.Enter && !e.Control && !e.Alt) {
                Control a = ActiveControl;
                if (a == null || (a is TextBox && ((TextBox)a).Multiline) || a is Button) return;
                if (a is ComboBox && ((ComboBox)a).DroppedDown) return;
                e.SuppressKeyPress = true;
                SelectNextControl(a, true, true, true, true);
            }
        };
        FormClosing += delegate(object s, FormClosingEventArgs e) {
            if (Home != null && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; GoBack(); }
        };
    }

    // ---------- поля ----------

    TextBox T(string key, string cue) {
        TextBox t = Ui.Text(1);
        Reg(key, t);
        if (cue != null) t.HandleCreated += delegate { Ui.Cue(t, cue); };
        return t;
    }

    Combo C(string key, string cue) {
        Combo c = new Combo(true, false);
        Reg(key, c);
        if (cue != null) c.HandleCreated += delegate { Ui.Cue(c, cue); };
        return c;
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

    DateBox D(string key) {
        DateBox d = new DateBox();
        Reg(key, d);
        return d;
    }

    void Reg(string key, Control c) {
        inp[key] = c;
        c.Name = key;
        EventHandler h = delegate {
            if (loading) return;
            Crash.Mark("БС, поле " + key);
            dirty = true;
            OnChanged(key);
            RefreshState();
        };
        if (c is Seg) ((Seg)c).Changed += h;
        else if (c is ComboBox) { ((ComboBox)c).TextChanged += h; ((ComboBox)c).SelectedIndexChanged += h; }
        else if (c is DateBox) ((DateBox)c).ValueChanged += h;
        else if (c is CheckBox) ((CheckBox)c).CheckedChanged += delegate { Enable(); h(c, EventArgs.Empty); };
        else c.TextChanged += h;
        if (c is TextBox || c is ComboBox) {
            c.Enter += delegate { c.BackColor = St.Focus; };
            c.Leave += delegate { c.BackColor = c.Enabled ? St.Input : St.Disabled; };
        }
        if (key == "np" || key == "ulica") c.Leave += delegate { string v = c.Text.Trim(); if (v.Length > 0 && Up(v) != v) c.Text = Up(v); };
    }

    void Field(string key, string caption, string card) { fieldList.Add(new string[] { key, caption, card }); }

    public string V(string key) {
        if (key == "pu") return pu.Value;
        if (key == "moshn") return moshn.Value;
        if (key == "linii") return linii.Value;
        if (key == "plomba_drugoe") return seals.Value;
        Control c;
        if (!inp.TryGetValue(key, out c)) return "";
        if (c is Seg) return ((Seg)c).Value;
        if (c is DateBox) return ((DateBox)c).Text2;
        if (c is CheckBox) return ((CheckBox)c).Checked ? "1" : "";
        return c.Text.Trim();
    }

    public void SetV(string key, string v) {
        v = v ?? "";
        if (key == "pu") { pu.Value = v; if (!loading) RefreshState(); return; }
        if (key == "moshn") { moshn.Value = v; if (!loading) RefreshState(); return; }
        if (key == "linii") { linii.Value = v; if (!loading) RefreshState(); return; }
        if (key == "plomba_drugoe") { seals.Value = v; return; }
        Control c;
        if (!inp.TryGetValue(key, out c)) return;
        if (c is Seg) ((Seg)c).Value = v;
        else if (c is CheckBox) ((CheckBox)c).Checked = v == "1";
        else if (c is DateBox) {
            DateTime d;
            if (DateTime.TryParseExact(v, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) ((DateBox)c).Value = d;
        } else if (c is Combo) ((Combo)c).Value = v;
        else c.Text = v;
    }

    Cell F(string caption, Control c, int chars) {
        Cell x = Cell.F(caption, c, chars == 0 ? 4 : 1);
        x.W = chars;
        return x;
    }

    // ---------- разделы ----------

    void BuildCards() {
        Card c = new Card("Договор и абонент", "юр лицо: договор, наименование, телефон", 4);
        cards.Add(c);
        c.Grid.Row(F("Договор №", T("dog_nom", "11098"), 10), F("от (дата договора)", D("dog_data"), 14),
            F("Абонент (наименование юр лица)", C("abonent", "ООО «...»"), 0), F("Телефон абонента", T("tel", "+375 29 ..."), 18));
        Field("dog_nom", "Договор №", c.Name); Field("dog_data", "Дата договора", c.Name); Field("abonent", "Абонент", c.Name); Field("tel", "Телефон", c.Name);

        c = new Card("Объект", "наименование, код и адрес - строки собираются сами, их можно поправить", 4);
        cards.Add(c);
        c.Grid.Row(F("Объект (как в документах)", C("obekt", "Возведение мобильной базовой станции БС-..."), 0),
            F("Вид сооружения и код", C("kod", "сооружение специализированное связи, код 3 07 00"), 36));
        c.Grid.Row(F("Область", C("obl", null), 16), F("Район", C("rayon", null), 14), F("Сельсовет", C("selsovet", null), 16),
            F("Тип н.п.", C("np_tip", null), 7), F("Населённый пункт", C("np", null), 0));
        c.Grid.Row(F("Тип улицы", C("ulica_tip", null), 7), F("Улица", C("ulica", null), 22),
            F("Дом (или «в районе дома №43»)", T("dom", null), 0));
        c.Grid.Row(F("Адрес полностью (справка, акт, доп. соглашение, потери)", T("adres", null), 0));
        c.Grid.Row(F("Адрес коротко (наряд)", T("adres_kr", null), 0));
        c.Grid.Row(F("Объект в справке, акте, потерях: «наименование по адресу: ...(код)»", T("obekt_full", null), 0));
        c.Grid.Row(F("Объект в доп. соглашении (ячейка таблицы)", T("obekt_adres", null), 0));
        Field("obekt", "Объект", c.Name); Field("kod", "Вид сооружения и код", c.Name); Field("np", "Населённый пункт", c.Name); Field("dom", "Дом", c.Name);

        c = new Card("Электроснабжение", "откуда питание, линия, место включения, акт разграничения", 4);
        cards.Add(c);
        c.Grid.Row(F("Питание от (подстанция)", C("pitanie", "КТПБ-4726"), 18), F("Акт разграничения от", D("data_razgr"), 14),
            F("Место включения (наряд)", C("mesto", "РП-301"), 0));
        c.Grid.Row(F("Отходящая линия (акт): «КЛ-0,4 кВ: РП-... - ЩУНИ ... марка L=..м.»", T("liniya", null), 0));
        Field("pitanie", "Питание от", c.Name); Field("data_razgr", "Акт разграничения", c.Name); Field("liniya", "Отходящая линия", c.Name); Field("mesto", "Место включения", c.Name);

        c = new Card("Приборы учёта", "строка на каждый прибор учёта, «+» - ещё один; в справку, акт и доп. соглашение", 4);
        cards.Add(c);
        pu = new RowList(new string[] { "Точка учёта", "Тип счётчика", "№ счётчика", "В", "А", "Показания", "Пов. кв.", "Пов. год", "ТТ тока", "ТТ напр.", "К расч" },
            new int[] { 7, 0, 12, 9, 7, 10, 5, 5, 5, 5, 4 },
            new bool[] { true, true, false, true, true, false, true, false, false, false, false },
            new string[] { "ЩУЭ", null, null, "3х230/400", "5(100)", "000000,00", null, null, "-", "-", "1" });
        pu.Name = "pu";
        Rows(pu);
        c.Grid.Row(F(null, pu, 0));
        c.Grid.Row(F(null, X("minus_on", "Строка в акте: «Данные прибора учёта № ... минусуются от данных прибора учёта № ...»"), 0));
        c.Grid.Row(F("минусуется ПУ №", T("minus_x", null), 14), F("от данных ПУ №", T("minus_y", null), 14), F("Строка в акте (собирается сама)", T("minus", null), 0));
        Field("pu", "Приборы учёта (тип и номер)", c.Name);

        c = new Card("Мощность", "расчёт по установленной мощности (справка, акт, доп. соглашение); строка на точку учёта", 4);
        cards.Add(c);
        moshn = new RowList(new string[] { "Нагрузка (точка учёта)", "Мощность", "Ки", "Время работы", "Выходные" },
            new int[] { 16, 10, 6, 14, 0 },
            new bool[] { true, true, true, true, true },
            new string[] { "ЩУЭ", null, "-", "24 ч/сут.", "Без выходных" });
        moshn.Name = "moshn";
        Rows(moshn);
        c.Grid.Row(F(null, moshn, 0));
        Field("moshn", "Мощность", c.Name);

        c = new Card("Пломбы", "в акт («Клемная крышка ... опломбирована пломбой ...») и в отчёт пломб", 4);
        cards.Add(c);
        seals = new Seals();
        seals.Name = "plomba_drugoe";
        seals.Changed += delegate { if (loading) return; dirty = true; Compose("plomby"); RefreshState(); };
        seals.Resized += delegate { LayoutCards(); };
        c.Grid.Row(F("Маркировка", C("plomba_marka", "МЭСБ"), 8), F("КК - клеммная крышка (номера через запятую)", T("plomba_klemm", null), 0),
            F("ВА - вводной автомат (номера)", T("plomba_vvod", null), 0), F("Ток ВА, А", C("va_tok", "63"), 6));
        c.Grid.Row(F("Другие пломбы (место | номер)", seals, 0));
        c.Grid.Row(F("Пломбы в акте (собирается само, можно поправить)", T("plomby", null), 0));

        c = new Card("Доп. соглашение", "пункт 1, тарифная группа, ТРЭТ, примечание - варианты пополняются в справочнике", 4);
        cards.Add(c);
        c.Grid.Row(F("Пункт 1 соглашения", C("punkt1", null), 0));
        c.Grid.Row(F("Тарифная группа", C("tarif", null), 0), F("ТРЭТ", T("tret", null), 8), F("Примечание", T("prim", null), 16));

        c = new Card("Потери", "Приложение 8: линии 0,4 кВ (Rл = r0 · L считается само) и схема питания", 4);
        cards.Add(c);
        linii = new RowList(new string[] { "Uн, кВ", "Марка, сечение", "L, км", "r0, Ом/км", "Rл = r0·L, Ом" },
            new int[] { 6, 0, 9, 9, 11 },
            new bool[] { true, true, false, false, false },
            new string[] { "0,4", null, null, null, null });
        linii.Name = "linii";
        Rows(linii);
        linii.Edited += delegate(int row, int col) { if (col == 1) R0For(row); if (col == 1 || col == 2 || col == 3) RFor(row); };
        c.Grid.Row(F(null, linii, 0));
        Button pick = Ui.Button("Выбрать схему…", false);
        pick.Name = "shemaPick";
        pick.Click += delegate { PickShema(); };
        tip.SetToolTip(pick, "Фото или скан схемы питания из проекта (JPG, PNG) - встанет в «потери» на место схемы");
        c.Grid.Row(F("Схема питания (картинка)", T("shema", "файл JPG/PNG"), 0), F(" ", pick, 16));
        shemaNote.Font = St.Small;
        shemaNote.ForeColor = St.Hint;
        c.Grid.Row(F(null, shemaNote, 0));
        Field("linii", "Линии (марка, L)", c.Name); Field("shema", "Схема питания", c.Name);

        c = new Card("Наряд", "наряд на включение (отключение): вид, причина, потребитель", 4);
        cards.Add(c);
        Seg vid = new Seg("включение", "отключение");
        Reg("vid", vid);
        c.Grid.Row(F("Наряд на", vid, 24), F("Причина", C("prichina", "Допуск"), 18), F("Потребитель в наряде (собирается сам)", T("potr", null), 0));
        c.Grid.Row(F("Объект в наряде (коротко)", T("obekt_kr", null), 0));

        c = new Card("Подписи и даты", "инженер и начальник - из справочника; даты документов", 4);
        cards.Add(c);
        c.Grid.Row(F("Инженер (Фамилия И.О.)", C("inzh", null), 18), F("Инженер полностью (акт)", T("inzh_fio", null), 0),
            F("Должность начальника", C("nach_dolzh", null), 16), F("Начальник (Фамилия И.О.)", C("nach", null), 16));
        c.Grid.Row(F("Данные ПУ на (справка)", D("data_pu"), 14), F("Дата подписей (справка)", D("data_podp"), 14), F("Дата акта", D("data_akt"), 14));
        Field("inzh", "Инженер", c.Name); Field("inzh_fio", "Инженер полностью", c.Name); Field("nach", "Начальник", c.Name);

        c = new Card("Документы", "что сформировать в папку пакета", 4);
        cards.Add(c);
        c.Grid.Row(F(null, X(Docs[0][0], Docs[0][2]), 26), F(null, X(Docs[1][0], Docs[1][2]), 26), F(null, X(Docs[2][0], Docs[2][2]), 0));
        c.Grid.Row(F(null, X(Docs[3][0], Docs[3][2]), 26), F(null, X(Docs[4][0], Docs[4][2]), 26));
        c.Grid.Row(F("Название пакета (папка в «БС»; собирается само)", T("paket", null), 0));
    }

    void Rows(RowList r) {
        r.Changed += delegate { if (loading) return; dirty = true; Compose("plomby"); RefreshState(); };
        r.Resized += delegate { LayoutCards(); };
    }

    void FillLists() {
        bool was = loading;
        loading = true;
        try {
            Items("abonent", memory.Hist("бс абонент"));
            Items("obekt", With(lists.Get("БС: объект"), memory.Hist("бс объект")));
            Items("kod", With(lists.Get("БС: вид сооружения и код"), memory.Hist("бс код")));
            Items("obl", lists.Get("БС: область"));
            Items("rayon", lists.Get("БС: район"));
            Items("selsovet", With(lists.Get("Сельсовет"), memory.Hist("сельсовет")));
            Items("np_tip", lists.Get("Тип населённого пункта"));
            Items("ulica_tip", lists.Get("Тип улицы"));
            Items("np", memory.Hist("нп"));
            Items("ulica", memory.Hist("улица"));
            Items("pitanie", memory.Hist("бс питание"));
            Items("mesto", memory.Hist("бс место"));
            Items("plomba_marka", lists.Get("БС: маркировка пломб"));
            Items("va_tok", lists.Get("БС: ток ВА"));
            Items("punkt1", lists.Get("БС: доп. соглашение, пункт 1"));
            Items("tarif", lists.Get("БС: тарифная группа"));
            Items("prichina", With(lists.Get("Наряд: причина"), memory.Hist("наряд причина")));
            Items("inzh", Engineers());
            Items("nach", Names());
            Items("nach_dolzh", lists.Get("Должность начальника"));
            pu.SetItems(0, lists.Get("БС: точка учёта"));
            pu.SetItems(1, With(lists.Get("БС: тип счётчика"), memory.Hist("бс счётчик")));
            pu.SetItems(3, lists.Get("БС: напряжение"));
            pu.SetItems(4, lists.Get("БС: ток счётчика"));
            pu.SetItems(6, new List<string>(new string[] { "1", "2", "3", "4" }));
            moshn.SetItems(0, lists.Get("БС: точка учёта"));
            moshn.SetItems(1, lists.Get("БС: мощность"));
            moshn.SetItems(2, new List<string>(new string[] { "-" }));
            moshn.SetItems(3, lists.Get("БС: время работы"));
            moshn.SetItems(4, lists.Get("БС: выходные"));
            linii.SetItems(0, new List<string>(new string[] { "0,4", "10" }));
            linii.SetItems(1, Cables());
            seals.SetPlaces(lists.Get("Места пломб"));
        } finally { loading = was; }
    }

    void Items(string key, List<string> items) { Combo c = inp[key] as Combo; if (c != null) c.SetItems(items); }

    static List<string> With(List<string> a, List<string> b) {
        List<string> r = new List<string>(a);
        foreach (string s in b) {
            bool found = false;
            foreach (string x in r) if (string.Compare(x, s, StringComparison.OrdinalIgnoreCase) == 0) found = true;
            if (!found) r.Add(s);
        }
        return r;
    }

    // инженеры: «Фамилия И.О. = Фамилия Имя Отчество» (полностью - для акта)
    List<string> Engineers() {
        List<string> r = new List<string>();
        foreach (string s in lists.Get("БС: инженер")) { int e = s.IndexOf('='); r.Add((e > 0 ? s.Substring(0, e) : s).Trim()); }
        foreach (string s in lists.Get("Оформил")) if (!r.Contains(s)) r.Add(s);
        return r;
    }

    string EngineerFull(string shortName) {
        foreach (string s in lists.Get("БС: инженер")) {
            int e = s.IndexOf('=');
            if (e > 0 && string.Compare(s.Substring(0, e).Trim(), shortName.Trim(), StringComparison.OrdinalIgnoreCase) == 0) return s.Substring(e + 1).Trim();
        }
        return "";
    }

    List<string> Names() {
        List<string> r = new List<string>();
        foreach (string[] s in lists.Signers()) r.Add(s[0]);
        if (r.Count == 0) r.Add("Зенченко Д.И.");
        return r;
    }

    // кабели: «АВБбШв-5х16 = 1,98» (марка, сечение = r0, Ом/км)
    List<string> Cables() {
        List<string> r = new List<string>();
        foreach (string s in lists.Get("БС: кабель и r0")) { int e = s.IndexOf('='); r.Add((e > 0 ? s.Substring(0, e) : s).Trim()); }
        return r;
    }

    string R0Of(string cable) {
        foreach (string s in lists.Get("БС: кабель и r0")) {
            int e = s.IndexOf('=');
            if (e > 0 && string.Compare(s.Substring(0, e).Trim(), cable.Trim(), StringComparison.OrdinalIgnoreCase) == 0) return s.Substring(e + 1).Trim();
        }
        return "";
    }

    void R0For(int row) {
        string r0 = R0Of(linii.Cell(row, 1));
        if (r0.Length > 0) linii.SetCell(row, 3, r0);
    }

    void RFor(int row) {
        double r0, l;
        if (Num(linii.Cell(row, 3), out r0) && Num(linii.Cell(row, 2), out l)) linii.SetCell(row, 4, Fmt(r0 * l));
        else linii.SetCell(row, 4, "");
    }

    static bool Num(string s, out double d) {
        return double.TryParse(s.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d);
    }

    // 0,12276 - запятая, до 5 знаков, без лишних нулей
    static string Fmt(double d) {
        return Math.Round(d, 5).ToString("0.#####", CultureInfo.InvariantCulture).Replace('.', ',');
    }

    void Enable() {
        bool m = V("minus_on") == "1";
        foreach (string k in new string[] { "minus_x", "minus_y", "minus" }) { Control c = inp[k]; c.Enabled = m; c.BackColor = m ? St.Input : St.Disabled; }
    }

    // ---------- строки из частей ----------

    void OnChanged(string key) {
        if (Array.IndexOf(AddrKeys, key) >= 0) { Compose("adres"); Compose("adres_kr"); Compose("obekt_full"); Compose("obekt_adres"); Compose("paket"); }
        else if (key == "obekt" || key == "kod") { Compose("obekt_full"); Compose("obekt_adres"); if (key == "obekt") { SetAuto("obekt_kr", V("obekt")); Compose("paket"); } }
        else if (key == "abonent" || key == "tel") Compose("potr");
        else if (key == "minus_x" || key == "minus_y") Compose("minus");
        else if (key.StartsWith("plomba") || key == "va_tok") Compose("plomby");
        else if (key == "inzh") Compose("inzh_fio");
        else if (key == "dog_nom") Compose("paket");
        else if (key == "shema") ShemaNote();
    }

    void SetAuto(string key, string v) {
        // поле, которое не правили руками, - следует за источником
        string old;
        if (autos.TryGetValue(key, out old) && V(key) != old) return;
        bool was = loading;
        loading = true;
        try { SetV(key, v); } finally { loading = was; }
        autos[key] = v;
    }

    void Compose(string key) { SetAuto(key, AutoOf(key)); }

    void ComposeAll() { foreach (string k in Composed) Compose(k); SetAuto("obekt_kr", V("obekt")); }

    void ResetAutos() { foreach (string k in Composed) autos[k] = ""; autos["obekt_kr"] = ""; }

    string Street() {
        string ul = (V("ulica_tip") + " " + V("ulica")).Trim();
        string dom = V("dom");
        if (dom.Length == 0) return ul;
        // «43» -> «ул. Авиационная, д. 43»; «в районе дома №43» -> «ул. Авиационная в районе дома №43»
        bool num = char.IsDigit(dom[0]);
        return (ul + (num ? (ul.Length > 0 ? ", " : "") + "д. " + dom : " " + dom)).Trim();
    }

    string AutoOf(string key) {
        switch (key) {
            case "adres": {
                List<string> p = new List<string>();
                if (V("obl").Length > 0) p.Add(V("obl"));
                if (V("rayon").Length > 0) p.Add(V("rayon"));
                if (V("selsovet").Length > 0) p.Add(Ss(V("selsovet")));
                if (V("np").Length > 0) p.Add((V("np_tip") + " " + V("np")).Trim());
                if (Street().Length > 0) p.Add(Street());
                return string.Join(", ", p.ToArray());
            }
            case "adres_kr": {
                List<string> p = new List<string>();
                if (V("np").Length > 0) p.Add((V("np_tip") + " " + V("np")).Trim());
                if (Street().Length > 0) p.Add(Street());
                string a = string.Join(", ", p.ToArray());
                return a;
            }
            case "obekt_full":
                return V("obekt").Length == 0 ? "" : "«" + V("obekt") + (V("adres").Length > 0 ? " по адресу: " + V("adres") : "") + (V("kod").Length > 0 ? " (" + V("kod") + ")" : "") + "»";
            case "obekt_adres":
                return V("obekt").Length == 0 ? "" : V("obekt") + (V("adres").Length > 0 ? " по адресу: " + V("adres") + "." : "") + (V("kod").Length > 0 ? " (" + V("kod") + ")" : "");
            case "potr":
                return V("abonent").Length == 0 ? "" : V("abonent") + (V("tel").Length > 0 ? " (" + V("tel") + ")" : "");
            case "minus":
                return V("minus_x").Length == 0 && V("minus_y").Length == 0 ? "" : "Данные прибора учёта № " + V("minus_x") + " минусуются от данных прибора учёта № " + V("minus_y") + ".";
            case "plomby": return SealsText();
            case "inzh_fio": return EngineerFull(V("inzh"));
            case "paket": {
                string n = (V("dog_nom").Length > 0 ? "Дог " + V("dog_nom") + " " : "") + V("np") + " " + ShortObject();
                foreach (char ch in Path.GetInvalidFileNameChars()) n = n.Replace(ch, '_');
                while (n.Contains("  ")) n = n.Replace("  ", " ");
                return n.Trim().TrimEnd('.');
            }
        }
        return "";
    }

    static string Ss(string s) {
        string l = s.ToLowerInvariant();
        return l.Contains("с/с") || l.Contains("сельсовет") ? s : s + " с/с";
    }

    // «Возведение мобильной базовой станции БС-6966» -> «БС-6966» (для названия папки), иначе - первые слова
    string ShortObject() {
        string o = V("obekt");
        Match m = Regex.Match(o, @"[А-ЯA-Z]{1,4}\s*-\s*\d+");
        if (m.Success) return m.Value.Replace(" ", "");
        string[] w = o.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return string.Join(" ", w, 0, Math.Min(3, w.Length));
    }

    // «МЭСБ 4948093; ВА(63)-МЭСБ 4948094, МЭСБ 4948095;»
    string SealsText() {
        string mk = V("plomba_marka");
        string pre = mk.Length > 0 ? mk + " " : "";
        List<string> parts = new List<string>();
        List<string> kk = SealRegistry.Numbers(V("plomba_klemm")), va = SealRegistry.Numbers(V("plomba_vvod"));
        if (kk.Count > 0) parts.Add(Join(pre, kk));
        if (va.Count > 0) parts.Add("ВА" + (V("va_tok").Length > 0 ? "(" + V("va_tok") + ")" : "") + "-" + Join(pre, va));
        foreach (string[] s in Seals.Parse(V("plomba_drugoe"))) {
            List<string> n = SealRegistry.Numbers(s[1]);
            if (n.Count > 0) parts.Add(s[0] + "-" + Join(pre, n));
        }
        return parts.Count == 0 ? "" : string.Join("; ", parts.ToArray()) + ";";
    }

    static string Join(string pre, List<string> nums) {
        List<string> r = new List<string>();
        foreach (string n in nums) r.Add(pre + n);
        return string.Join(", ", r.ToArray());
    }

    // ---------- схема ----------

    void PickShema() {
        using (OpenFileDialog d = new OpenFileDialog()) {
            d.Title = "Схема питания (фото или скан)";
            d.Filter = "Картинки (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|Все файлы (*.*)|*.*";
            try { if (File.Exists(V("shema"))) d.InitialDirectory = Path.GetDirectoryName(V("shema")); } catch { }
            if (d.ShowDialog(this) == DialogResult.OK) SetV("shema", d.FileName);
        }
    }

    void ShemaNote() {
        string f = V("shema");
        if (f.Length == 0) { shemaNote.Text = "схема не выбрана - в «потерях» останется место под схему (вставите в Word)"; shemaNote.ForeColor = St.Hint; return; }
        try {
            using (Image im = Image.FromFile(f)) { shemaNote.Text = "схема: " + Path.GetFileName(f) + ", " + im.Width + "×" + im.Height; shemaNote.ForeColor = St.Good; }
        } catch (Exception ex) { shemaNote.Text = "схема не открывается: " + ex.Message; shemaNote.ForeColor = St.Bad; }
    }

    // ---------- окно ----------

    Control BuildNav() {
        Panel nav = new Panel();
        nav.Dock = DockStyle.Left;
        nav.Width = St.NavWidth;
        nav.BackColor = St.Nav;
        Label title = new Label();
        title.Text = "БС";
        title.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
        title.ForeColor = Color.White;
        title.SetBounds(22, 22, 210, 34);
        Label sub = new Label();
        sub.Text = "пакет документов юр лица";
        sub.Font = St.Small;
        sub.ForeColor = St.NavText;
        sub.SetBounds(24, 58, 210, 20);
        nav.Controls.Add(title);
        nav.Controls.Add(sub);
        NavItem back = new NavItem("← Назад", false);
        back.SetBounds(0, 90, nav.Width, back.Height);
        back.Click += delegate { GoBack(); };
        nav.Controls.Add(back);
        int y = 130;
        foreach (Card c in cards) {
            NavItem n = new NavItem(c.Name, false);
            n.Mark = St.TintOf(c.Name).Mark;
            n.SetBounds(0, y, nav.Width, n.Height);
            Card card = c;
            n.Click += delegate { content.AutoScrollPosition = new Point(0, Math.Max(0, card.Panel.Top - content.AutoScrollPosition.Y - 4)); };
            nav.Controls.Add(n);
            y += n.Height;
        }
        NavItem req = new NavItem("Обязательные поля", true);
        req.Click += delegate { SetupRequired(); };
        NavItem openItem = new NavItem("Открыть сохранённый…", true);
        openItem.Click += delegate { OpenSaved(); };
        NavItem folder = new NavItem("Папка БС", true);
        folder.Click += delegate {
            try { Directory.CreateDirectory(Folder(here)); Process.Start("explorer.exe", "\"" + Folder(here) + "\""); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text); }
        };
        NavItem dirs = new NavItem("Справочники (списки)", true);
        dirs.Click += delegate {
            try { Process.Start("notepad.exe", "\"" + lists.File + "\""); } catch (Exception ex) { MessageBox.Show(this, ex.Message, Text); }
        };
        NavItem[] low = { req, openItem, folder, dirs };
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
            if (dirty && MessageBox.Show(this, "Очистить форму? Введённое пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Clear();
        };
        Button look = Ui.Button("Предпросмотр", false);
        look.Size = new Size(140, 38);
        look.Click += delegate { Preview(); };
        tip.SetToolTip(look, "Отмеченные документы - в Word, ничего не сохраняя (Ctrl+P)");
        Button save = Ui.Button("Сформировать пакет", true);
        save.Size = new Size(220, 40);
        save.Click += delegate { Save(true, true); };
        tip.SetToolTip(save, "Отмеченные документы - в папку «БС\\<название пакета>», данные - чтобы открыть заново (Ctrl+S)");
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
            status.SetBounds(clear.Right + 20, 14, Math.Max(10, look.Left - clear.Right - 40), 38);
            status.Font = status.Width < 360 ? St.Small : St.Base;
        };
        return p;
    }

    void LayoutCards() {
        content.SuspendLayout();
        int full = content.Width - SystemInformation.VerticalScrollBarWidth;
        int w = Math.Max(640, Math.Min(full - 32, 1120));
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

    // ---------- данные ----------

    static string First(List<string> l) { return l.Count > 0 ? l[0] : ""; }
    static string Up(string s) { return s.Length == 0 ? s : char.ToUpper(s[0]) + s.Substring(1); }

    public void Clear() {
        loading = true;
        try {
            foreach (KeyValuePair<string, Control> kv in inp) if (!(kv.Value is DateBox) && !(kv.Value is Seg)) SetV(kv.Key, "");
            foreach (string k in new string[] { "dog_data", "data_razgr", "data_pu", "data_podp", "data_akt" }) ((DateBox)inp[k]).Value = DateTime.Today;
            pu.Value = "";
            moshn.Value = "";
            linii.Value = "";
            seals.Value = "";
            SetV("obl", memory.Get("БСОбласть", First(lists.Get("БС: область"))));
            SetV("rayon", memory.Get("БСРайон", First(lists.Get("БС: район"))));
            SetV("np_tip", "д.");
            SetV("ulica_tip", "ул.");
            SetV("kod", First(lists.Get("БС: вид сооружения и код")));
            SetV("plomba_marka", memory.Get("БСМаркировка", First(lists.Get("БС: маркировка пломб"))));
            SetV("punkt1", First(lists.Get("БС: доп. соглашение, пункт 1")));
            SetV("tarif", First(lists.Get("БС: тарифная группа")));
            SetV("tret", "-");
            SetV("prim", "-");
            SetV("vid", "включение");
            SetV("prichina", First(lists.Get("Наряд: причина")));
            string eng = memory.Get("БСИнженер", "");
            SetV("inzh", Engineers().Contains(eng) ? eng : First(Engineers()));
            SetV("nach", memory.Get("БСНачальник", First(Names())));
            SetV("nach_dolzh", memory.Get("БСДолжностьНачальника", First(lists.Get("Должность начальника"))));
            string docs = memory.Get("БСДокументы", "11111");
            for (int i = 0; i < Docs.Length; i++) SetV(Docs[i][0], i < docs.Length && docs[i] == '0' ? "" : "1");
            ResetAutos();
            ComposeAll();
            savedName = null;
        } finally { loading = false; }
        dirty = false;
        Enable();
        ShemaNote();
        head.Text = "Новый пакет документов";
        headSub.ForeColor = St.Hint;
        headSub.Text = "Заполните договор, объект, приборы учёта, мощность, линии - строки документов собираются сами";
        RefreshState();
    }

    public Dictionary<string, string> Data() {
        Dictionary<string, string> d = new Dictionary<string, string>();
        foreach (string k in inp.Keys) d[k] = V(k);
        d["pu"] = pu.Value;
        d["moshn"] = moshn.Value;
        d["linii"] = linii.Value;
        d["plomba_drugoe"] = seals.Value;
        return d;
    }

    void OpenSaved() {
        string dir = DataFolder(here);
        if (!Directory.Exists(dir) || Directory.GetFiles(dir, "*.txt").Length == 0) { MessageBox.Show(this, "Сохранённых пакетов документов пока нет.", Text); return; }
        if (dirty && MessageBox.Show(this, "Пакет не сохранён. Открыть другой? Введённое здесь пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        using (SavedDialog dlg = new SavedDialog(dir, 3)) {
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
            foreach (string k in new string[] { "pu", "moshn", "linii", "plomba_drugoe" }) { string v; if (d.TryGetValue(k, out v)) SetV(k, v); }
            foreach (string k in Composed) autos[k] = AutoOf(k);
            autos["obekt_kr"] = V("obekt");
            savedName = name;
        } finally { loading = false; }
        dirty = false;
        Enable();
        ShemaNote();
        LayoutCards();
        head.Text = "Пакет «" + name + "»";
        headSub.ForeColor = St.Hint;
        string when;
        d.TryGetValue("сохранён", out when);
        headSub.Text = "открыт сохранённый (" + (when ?? "") + ") - исправьте и сформируйте заново";
        RefreshState();
    }

    void SaveData(string name) {
        string dir = DataFolder(here);
        Directory.CreateDirectory(dir);
        StringBuilder sb = new StringBuilder();
        sb.Append("# Данные пакета БС для программы «Оформление договора» (открыть заново, пломбы - в отчёт пломб)\r\n");
        sb.Append("сохранён=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
        foreach (KeyValuePair<string, string> kv in Data()) sb.Append(kv.Key).Append('=').Append(Memory.Escape(kv.Value)).Append("\r\n");
        // для отчёта пломб: дата, договор, № учёта (первый счётчик), адрес - как у договоров
        sb.Append("data=").Append(Memory.Escape(V("data_akt"))).Append("\r\n");
        sb.Append("pu1_zav=").Append(Memory.Escape(pu.Cell(0, 2))).Append("\r\n");
        string file = Path.Combine(dir, name + ".txt"), tmp = file + ".tmp";
        File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(true));
        if (File.Exists(file)) File.Replace(tmp, file, null);
        else File.Move(tmp, file);
    }

    // ---------- значения для шаблонов ----------

    public Dictionary<string, string> Compose(List<Dictionary<string, string>> puRows, List<Dictionary<string, string>> mRows, List<Dictionary<string, string>> lRows) {
        Dictionary<string, string> v = new Dictionary<string, string>();
        v["ДОГНОМ"] = V("dog_nom");
        v["ДОГДАТА"] = V("dog_data");
        v["АБОНЕНТ"] = V("abonent");
        v["ТЕЛЕФОН"] = V("tel");
        v["ОБЪЕКТ"] = V("obekt_full");
        v["ОБЪЕКТАДРЕС"] = V("obekt_adres");
        v["ОБЪЕКТКР"] = V("obekt_kr");
        v["АДРЕС"] = V("adres").TrimEnd('.');
        v["АДРЕСКР"] = V("adres_kr");
        v["АДРЕСНАР"] = V("adres_kr") + (V("selsovet").Length > 0 ? " (" + Ss(V("selsovet")) + ")" : "");
        v["ПИТАНИЕ"] = V("pitanie");
        v["ДАТАРАЗГР"] = V("data_razgr");
        v["ЛИНИЯ"] = V("liniya");
        v["МЕСТО"] = V("mesto");
        v["МИНУС"] = V("minus_on") == "1" && V("minus").Length > 0 ? V("minus") : "";
        v["ПЛОМБЫ"] = V("plomby");
        v["ПУНКТ1"] = V("punkt1");
        v["ТАРИФГР"] = V("tarif");
        v["ТРЭТ"] = V("tret");
        v["ПРИМ"] = V("prim");
        v["ВИД"] = V("vid") == "отключение" ? "отключение" : "включение";
        v["ПРИЧИНА"] = V("prichina");
        v["ПОТРЕБИТЕЛЬ"] = V("potr");
        v["ИНЖЕНЕР"] = V("inzh");
        v["ИНЖЕНЕРФИО"] = V("inzh_fio").Length > 0 ? V("inzh_fio") : V("inzh");
        v["НАЧАЛЬНИК"] = V("nach");
        // 24.11: должность начальника - на выбор (ГРП / ОПСУ); в доп. соглашении у ГРП - «Начальник ГРП ОПСУ», как в образце
        string pos = V("nach_dolzh").Length > 0 ? V("nach_dolzh") : "Начальник ГРП";
        v["НАЧДОЛЖН"] = pos;
        v["НАЧДОЛЖНДС"] = pos == "Начальник ГРП" ? "Начальник ГРП  ОПСУ" : pos;
        v["ДАТАПУ"] = V("data_pu");
        v["ДАТАПОДП"] = V("data_podp");
        DateTime a = ((DateBox)inp["data_akt"]).Value;
        v["ДАТААКТ"] = "«" + a.Day + "» " + Months[a.Month - 1] + " " + a.Year + " г";
        // строки таблиц: ПУ, мощность (в доп. соглашении - у строки ПУ своя мощность), линии
        for (int i = 0; i < pu.Count; i++) {
            if (pu.Empty(i) && i > 0) continue;
            Dictionary<string, string> r = new Dictionary<string, string>();
            string kv = pu.Cell(i, 6);
            r["ПУН"] = (puRows.Count + 1).ToString();
            r["ПУТОЧКА"] = pu.Cell(i, 0);
            r["ПУТИП"] = pu.Cell(i, 1);
            r["ПУНОМЕР"] = pu.Cell(i, 2);
            r["ПУВ"] = pu.Cell(i, 3);
            r["ПУА"] = pu.Cell(i, 4);
            r["ПУПОКАЗ"] = pu.Cell(i, 5);
            r["ПУКВ"] = kv.Length > 0 && char.IsDigit(kv[kv.Length - 1]) ? kv + " кв." : kv;
            r["ПУГОД"] = pu.Cell(i, 7);
            r["ПУТТ"] = pu.Cell(i, 8);
            r["ПУТН"] = pu.Cell(i, 9);
            r["ПУК"] = pu.Cell(i, 10);
            r["ПУТИПНОМ"] = pu.Cell(i, 1) + (pu.Cell(i, 2).Length > 0 ? "\n№ " + pu.Cell(i, 2) : "");
            int m = Math.Min(puRows.Count, moshn.Count - 1);
            r["ММОЩН"] = Power(moshn.Cell(m, 1));
            r["МРЕЖИМ"] = Mode(moshn.Cell(m, 3), moshn.Cell(m, 4));
            r["ОБЪЕКТАДРЕС"] = v["ОБЪЕКТАДРЕС"];
            r["ТАРИФГР"] = v["ТАРИФГР"];
            r["ТРЭТ"] = v["ТРЭТ"];
            r["ПРИМ"] = v["ПРИМ"];
            puRows.Add(r);
        }
        if (puRows.Count > 0) v["ПУТИПНОМ1"] = puRows[0]["ПУТИП"] + (puRows[0]["ПУНОМЕР"].Length > 0 ? " № " + puRows[0]["ПУНОМЕР"] : "");
        else v["ПУТИПНОМ1"] = "";
        for (int i = 0; i < moshn.Count; i++) {
            if (moshn.Empty(i) && i > 0) continue;
            Dictionary<string, string> r = new Dictionary<string, string>();
            r["МН"] = (mRows.Count + 1).ToString();
            r["МНАГР"] = moshn.Cell(i, 0);
            r["ММОЩН"] = Power(moshn.Cell(i, 1));
            r["МКИ"] = moshn.Cell(i, 2);
            r["МВРЕМЯ"] = moshn.Cell(i, 3);
            r["МВЫХ"] = moshn.Cell(i, 4);
            mRows.Add(r);
        }
        for (int i = 0; i < linii.Count; i++) {
            if (linii.Empty(i) && i > 0) continue;
            Dictionary<string, string> r = new Dictionary<string, string>();
            r["ЛН"] = (lRows.Count + 1).ToString();
            r["ЛU"] = linii.Cell(i, 0);
            r["ЛМАРКА"] = linii.Cell(i, 1);
            r["ЛДЛИНА"] = linii.Cell(i, 2);
            r["ЛR0"] = linii.Cell(i, 3);
            r["ЛR"] = linii.Cell(i, 4);
            lRows.Add(r);
        }
        return v;
    }

    // «20» -> «20 кВт»
    static string Power(string s) {
        s = s.Trim();
        if (s.Length == 0) return s;
        return char.IsDigit(s[s.Length - 1]) ? s + " кВт" : s;
    }

    // доп. соглашение: «24 ч/сут, без выходных»
    static string Mode(string time, string days) {
        string t = time.Trim().TrimEnd('.');
        string d = days.Trim();
        if (d.Length > 0) d = char.ToLower(d[0]) + d.Substring(1);
        return (t + (t.Length > 0 && d.Length > 0 ? ", " : "") + d).Trim();
    }

    // заполнить один документ: строки таблиц (ПУ, мощность, линии) -> поля «ИМЯ_1», «ИМЯ_2»...; потери - со своей схемой
    public string Fill(string[] doc, string output) {
        string tpl = Path.Combine(here, doc[1]);
        if (!File.Exists(tpl)) throw new FileNotFoundException("Нет шаблона " + doc[1] + " рядом с программой");
        List<Dictionary<string, string>> puRows = new List<Dictionary<string, string>>(), mRows = new List<Dictionary<string, string>>(), lRows = new List<Dictionary<string, string>>();
        Dictionary<string, string> v = Compose(puRows, mRows, lRows);
        List<string> fields = DocxTemplate.Fields(tpl);
        string cur = tpl;
        List<string> temps = new List<string>();
        try {
            string[][] tables = { new string[] { "ПУТОЧКА", "pu" }, new string[] { "МНАГР", "m" }, new string[] { "ЛМАРКА", "l" } };
            foreach (string[] t in tables) {
                if (!fields.Contains(t[0])) continue;
                List<Dictionary<string, string>> rows = t[1] == "pu" ? puRows : t[1] == "m" ? mRows : lRows;
                if (rows.Count == 0) rows.Add(new Dictionary<string, string>());
                string next = output + "." + t[1] + ".tmp";
                temps.Add(next);
                DocxTemplate.ExpandRows(cur, next, t[0], null, rows, new string[0], new string[] { "ОБЪЕКТАДРЕС", "ТАРИФГР", "ТРЭТ", "ПРИМ" });
                cur = next;
                for (int i = 0; i < rows.Count; i++) foreach (KeyValuePair<string, string> kv in rows[i]) v[kv.Key + "_" + (i + 1)] = kv.Value;
            }
            DocxTemplate.Fill(cur, output, v);
            if (doc[1] == "bs_poteri.docx" && V("shema").Length > 0) PutShema(output, V("shema"));
        } finally {
            foreach (string t in temps) try { File.Delete(t); } catch { }
        }
        return output;
    }

    // схема питания: картинка-заглушка шаблона (единственная .jpeg в документе) заменяется фото пользователя (в JPEG,
    // не шире 1600 точек); размер на листе - ширина как у заглушки, высота - по пропорциям (не выше 20 см)
    public static void PutShema(string docx, string picture) {
        byte[] jpg;
        int w, h;
        using (Image src = Image.FromFile(picture)) {
            double k = Math.Min(1.0, 1600.0 / Math.Max(src.Width, src.Height));
            w = Math.Max(1, (int)(src.Width * k));
            h = Math.Max(1, (int)(src.Height * k));
            using (Bitmap b = new Bitmap(w, h)) {
                using (Graphics g = Graphics.FromImage(b)) {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(src, 0, 0, w, h);
                }
                using (MemoryStream ms = new MemoryStream()) { b.Save(ms, ImageFormat.Jpeg); jpg = ms.ToArray(); }
            }
        }
        List<ZipEntry> entries = Zip.Read(docx);
        foreach (ZipEntry e in entries) {
            if (e.Name.StartsWith("word/media/") && e.Name.EndsWith(".jpeg")) e.SetData(jpg);
            if (e.Name != "word/document.xml") continue;
            string xml = Encoding.UTF8.GetString(e.GetData());
            Match m = Regex.Match(xml, "<wp:extent cx=\"(\\d+)\" cy=\"(\\d+)\"/>");
            if (!m.Success) continue;
            long cx = long.Parse(m.Groups[1].Value), cy = cx * h / w, maxCy = 7200000;
            if (cy > maxCy) { cx = cx * maxCy / cy; cy = maxCy; }
            string oldCx = m.Groups[1].Value, oldCy = m.Groups[2].Value;
            xml = xml.Replace("<wp:extent cx=\"" + oldCx + "\" cy=\"" + oldCy + "\"/>", "<wp:extent cx=\"" + cx + "\" cy=\"" + cy + "\"/>");
            xml = xml.Replace("<a:ext cx=\"" + oldCx + "\" cy=\"" + oldCy + "\"/>", "<a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>");
            e.SetData(Encoding.UTF8.GetBytes(xml));
        }
        Zip.Write(docx, entries);
    }

    List<string[]> Chosen() {
        List<string[]> r = new List<string[]>();
        foreach (string[] d in Docs) if (V(d[0]) == "1") r.Add(d);
        return r;
    }

    public List<string> Preview() {
        List<string> made = new List<string>();
        List<string[]> docs = Chosen();
        if (docs.Count == 0) { MessageBox.Show(this, "Не отмечено ни одного документа (раздел «Документы»).", Text); return made; }
        try {
            foreach (string[] d in docs) {
                string f = DocPreview.NewFile(here, d[2]);
                Fill(d, f);
                made.Add(f);
            }
            foreach (string f in made) DocPreview.Show(this, f);
        } catch (Exception ex) {
            MessageBox.Show(this, "Предпросмотр не получился:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return made;
    }

    public string PackageName() {
        string n = V("paket");
        foreach (char ch in Path.GetInvalidFileNameChars()) n = n.Replace(ch, '_');
        n = n.Trim().TrimEnd('.');
        return n.Length > 0 ? n : "Пакет документов " + DateTime.Now.ToString("yyyy-MM-dd HH-mm");
    }

    // сформировать пакет: отмеченные документы - в папку «БС\<название>»; openFolder - открыть папку. Возвращает папку или null
    public string Save(bool openFolder, bool ask) {
        List<string> m = Missing();
        if (ask && m.Count > 0) {
            List<string> names = new List<string>();
            foreach (string k in m) names.Add("  • " + Caption(k));
            if (MessageBox.Show(this, "Не заполнено:\n" + string.Join("\n", names.ToArray()) + "\n\nВсё равно сформировать? (Незаполненное можно дописать в Word.)",
                    "Проверка перед сохранением", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
        }
        List<string[]> docs = Chosen();
        if (docs.Count == 0) { if (ask) MessageBox.Show(this, "Не отмечено ни одного документа (раздел «Документы»).", Text); return null; }
        string name = PackageName();
        string dir = Path.Combine(Folder(here), name);
        if (ask && savedName != name && Directory.Exists(dir) && Directory.GetFiles(dir, "*.docx").Length > 0
            && MessageBox.Show(this, "Пакет «" + name + "» уже есть. Заменить его документы?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
        List<string> busy = new List<string>();
        try {
            Directory.CreateDirectory(dir);
            foreach (string[] d in docs) {
                string f = Path.Combine(dir, d[2] + ".docx");
                string tmp = f + ".new";
                Fill(d, tmp);
                try { if (File.Exists(f)) File.Delete(f); File.Move(tmp, f); }
                catch (IOException) { busy.Add(d[2]); try { File.Delete(tmp); } catch { } }
            }
            SaveData(name);
        } catch (Exception ex) {
            MessageBox.Show(this, "Пакет не сформирован:\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }
        Remember();
        savedName = name;
        dirty = false;
        head.Text = "Пакет «" + name + "»";
        headSub.ForeColor = St.Good;
        headSub.Text = "Сформирован: " + docs.Count + " документ(ов) в папке БС\\" + name;
        if (busy.Count > 0)
            MessageBox.Show(this, "Открыты в Word и не заменены:\n" + string.Join("\n", busy.ToArray()) + "\n\nЗакройте их в Word и сформируйте ещё раз.", Text,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        RefreshState();
        if (openFolder) try { Process.Start("explorer.exe", "\"" + dir + "\""); } catch { }
        return dir;
    }

    void Remember() {
        memory.Remember("бс абонент", V("abonent"));
        memory.Remember("бс объект", V("obekt"));
        memory.Remember("бс код", V("kod"));
        memory.Remember("бс питание", V("pitanie"));
        memory.Remember("бс место", V("mesto"));
        memory.Remember("нп", V("np"));
        memory.Remember("улица", V("ulica"));
        memory.Remember("сельсовет", V("selsovet"));
        for (int i = 0; i < pu.Count; i++) memory.Remember("бс счётчик", pu.Cell(i, 1));
        memory.Values["БСОбласть"] = V("obl");
        memory.Values["БСРайон"] = V("rayon");
        memory.Values["БСМаркировка"] = V("plomba_marka");
        memory.Values["БСИнженер"] = V("inzh");
        memory.Values["БСНачальник"] = V("nach");
        memory.Values["БСДолжностьНачальника"] = V("nach_dolzh");
        StringBuilder docs = new StringBuilder();
        foreach (string[] d in Docs) docs.Append(V(d[0]) == "1" ? '1' : '0');
        memory.Values["БСДокументы"] = docs.ToString();
        memory.Save();
    }

    // ---------- обязательные поля и состояние ----------

    void LoadRequired() {
        requiredSet = new Dictionary<string, bool>();
        string v = memory.Get("БСОбязательные", null);
        string[] keys = v == null ? DefaultRequired : v.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string k in keys) requiredSet[k.Trim()] = true;
    }

    void SetupRequired() {
        RequiredDialog d = new RequiredDialog(fieldList, requiredSet, DefaultRequired);
        if (d.ShowDialog(this) != DialogResult.OK) return;
        requiredSet = d.Result;
        List<string> keys = new List<string>();
        foreach (string[] f in fieldList) if (requiredSet.ContainsKey(f[0])) keys.Add(f[0]);
        memory.Values["БСОбязательные"] = string.Join(",", keys.ToArray());
        memory.Save();
        RefreshState();
    }

    bool Filled(string key) {
        if (key == "pu") return pu.Cell(0, 1).Length > 0 && pu.Cell(0, 2).Length > 0;
        if (key == "moshn") return moshn.Cell(0, 1).Length > 0;
        if (key == "linii") return linii.Cell(0, 1).Length > 0 && linii.Cell(0, 2).Length > 0;
        if (key == "shema") return V("shema").Length > 0 && File.Exists(V("shema"));
        return V(key).Length > 0;
    }

    List<string> Missing() {
        List<string> m = new List<string>();
        foreach (string[] f in fieldList) if (requiredSet.ContainsKey(f[0]) && !Filled(f[0])) m.Add(f[0]);
        return m;
    }

    string Caption(string key) {
        foreach (string[] f in fieldList) if (f[0] == key) return f[1];
        return key;
    }

    void RefreshState() {
        List<string> m = Missing();
        if (m.Count == 0) { status.Text = savedName != null && !dirty ? "Пакет сформирован" : "Обязательное заполнено"; status.ForeColor = St.Good; }
        else {
            List<string> names = new List<string>();
            foreach (string k in m) names.Add(Caption(k));
            status.Text = "Не заполнено: " + string.Join(", ", names.ToArray());
            status.ForeColor = St.Bad;
        }
    }

    public string StatusText { get { return status.Text; } }

    public bool AskSaveBeforeExit() {
        if (!dirty) return true;
        DialogResult r = MessageBox.Show(this, "Пакет документов не сформирован. Сформировать его?\n\nДа - сформировать и выйти\nНет - выйти без сохранения\nОтмена - вернуться", Text,
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) return Save(false, true) != null;
        return true;
    }
}
// ==== КОНЕЦ БЛОКА ФОРМЫ БС ====
