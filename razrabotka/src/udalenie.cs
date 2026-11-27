// ==== БЛОК СПРАВКИ НА УДАЛЕНИЕ (исходник - razrabotka/src/udalenie.cs; здесь не править - вставляется скриптом inject.py) ====
// «Справка в расчетную группу» - исключить ПУ проданной квартиры из расчётов юр лица (шаблон udalenie.docx, по фото 24.11).
// Заполняется из «Перевода из юр в физ» (договор/карточка, наименование, объект, адрес, лицевой, счётчики) или из сохранённой
// справки (галочка «+ удалить из юр»). Два экземпляра на листе - как справка. Папка «Справки на удаление» (+ данные).
public class UdalenieForm : Form {
    readonly string here, templateFile;
    readonly Lists lists;
    readonly Memory memory;
    readonly Dictionary<string, Control> inp = new Dictionary<string, Control>();
    readonly List<Card> cards = new List<Card>();
    readonly Panel content = new Panel();
    readonly Label head = new Label(), headSub = new Label(), status = new Label();
    readonly ToolTip tip = new ToolTip();
    RowList pu;
    public StartMenu Home;
    public Form Back;
    bool loading, dirty;
    string savedName;
    static readonly string[] Needed = { "kp", "nab", "dog_name", "ls", "adres", "pu", "inzh", "nach", "opsu" };

    public bool Unsaved { get { return dirty; } }
    public static string Folder(string here) { return Path.Combine(here, "Справки на удаление"); }
    public static string DataFolder(string here) { return Path.Combine(Folder(here), "данные"); }

    public UdalenieForm(string here) {
        Load += delegate { St.Dress(this); };
        this.here = here;
        templateFile = Path.Combine(here, "udalenie.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = Memory.For(Path.Combine(here, "nastroiki.txt"));
        Text = "Справка в расчетную группу - удаление ПУ из юр";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1180, wa.Width), Math.Min(820, wa.Height));
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
            Crash.Mark("справка на удаление, поле " + key);
            dirty = true;
            RefreshState();
        };
        if (c is ComboBox) { ((ComboBox)c).TextChanged += h; ((ComboBox)c).SelectedIndexChanged += h; }
        else if (c is DateBox) ((DateBox)c).ValueChanged += h;
        else if (c is CheckBox) ((CheckBox)c).CheckedChanged += h;
        else c.TextChanged += h;
        if (c is TextBox || c is ComboBox) {
            c.Enter += delegate { c.BackColor = St.Focus; };
            c.Leave += delegate { c.BackColor = c.Enabled ? St.Input : St.Disabled; };
        }
    }

    public string V(string key) {
        if (key == "pu") return pu.Value;
        Control c;
        if (!inp.TryGetValue(key, out c)) return "";
        if (c is DateBox) return ((DateBox)c).Text2;
        if (c is CheckBox) return ((CheckBox)c).Checked ? "1" : "";
        return c.Text.Trim();
    }

    public void SetV(string key, string v) {
        v = v ?? "";
        if (key == "pu") { pu.Value = v; return; }
        Control c;
        if (!inp.TryGetValue(key, out c)) return;
        if (c is CheckBox) ((CheckBox)c).Checked = v == "1";
        else if (c is DateBox) {
            DateTime d;
            if (DateTime.TryParseExact(v, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) ((DateBox)c).Value = d;
        } else if (c is Combo) ((Combo)c).Value = v;
        else c.Text = v;
    }

    Cell F(string caption, Control c, int chars) { Cell x = Cell.F(caption, c, chars == 0 ? 4 : 1); x.W = chars; return x; }

    void BuildCards() {
        Card c = new Card("Договор юр лица", "договор и карточка - из «Перевода из юр в физ» или вписать", 4);
        cards.Add(c);
        c.Grid.Row(F("Договор №", T("kp", "11118"), 10), F("Карточка №", T("nab", "52"), 8), F("Наименование юр лица", C("dog_name", "ООО «...»"), 0));
        c.Grid.Row(F("Объект", C("obekt", "Нераспроданные квартиры"), 0), F("В связи с продажей", C("vid", "квартиры"), 16));

        c = new Card("Потребитель", "новый договор с собственником и адрес", 4);
        cards.Add(c);
        c.Grid.Row(F("Лицевой нового договора (9 цифр)", T("ls", "052950737"), 16), F("Расположенный по адресу", T("adres", "д. ..., ул. ..., д. ..., кв. ..."), 0));

        c = new Card("Приборы учёта", "данные ПУ на момент продажи; строка на счётчик, «+» - ещё", 4);
        cards.Add(c);
        pu = new RowList(new string[] { "Тип счётчика", "№ счётчика", "Поверка (г.п.)", "Показания" }, new int[] { 0, 14, 10, 14 },
            new bool[] { true, false, false, false }, null);
        pu.Name = "pu";
        pu.Changed += delegate { if (loading) return; dirty = true; RefreshState(); };
        pu.Resized += delegate { LayoutCards(); };
        c.Grid.Row(F(null, pu, 0));

        c = new Card("Подписи и даты", "инженер, начальник ГРП, начальник ОПСУ, дата; два экземпляра на листе - как справка", 4);
        cards.Add(c);
        DateBox date = new DateBox();
        Reg("data", date);
        CheckBox two = new CheckBox();
        two.Text = "Два экземпляра на листе (линия отреза посередине)";
        two.Font = St.Base;
        two.ForeColor = St.Text;
        two.BackColor = Color.Transparent;
        Reg("dve", two);
        c.Grid.Row(F("Инженер", C("inzh", null), 18), F("Дата", date, 14), F("Должность начальника", C("nach_dolzh", null), 16), F("Начальник", C("nach", null), 18));
        // 24.11: подписывает и начальник ОПСУ (строка под начальником ГРП)
        c.Grid.Row(F("Начальник ОПСУ", C("opsu", null), 20), F(null, two, 0));
    }

    void FillLists() {
        bool was = loading;
        loading = true;
        try {
            Items("dog_name", memory.Hist("удаление юр лицо"));
            Items("obekt", With(lists.Get("Удаление: объект"), memory.Hist("удаление объект")));
            Items("vid", lists.Get("Удаление: в связи с продажей"));
            Items("inzh", lists.Get("Оформил"));
            List<string> n = new List<string>();
            foreach (string[] s in lists.Signers()) n.Add(s[0]);
            Items("nach", n);
            Items("nach_dolzh", lists.Get("Должность начальника"));
            Items("opsu", n);
            pu.SetItems(0, Types());
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

    static string Get(Dictionary<string, string> d, string k) {
        string v;
        return d != null && d.TryGetValue(k, out v) && v != null ? v.Trim() : "";
    }

    Control BuildNav() {
        Panel nav = new Panel();
        nav.Dock = DockStyle.Left;
        nav.Width = St.NavWidth;
        nav.BackColor = St.Nav;
        Label title = new Label();
        title.Text = "Справка в РГ";
        title.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
        title.ForeColor = Color.White;
        title.SetBounds(22, 22, 220, 34);
        Label sub = new Label();
        sub.Text = "удаление ПУ из юр (в расчетную группу)";
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
        NavItem folder = new NavItem("Папка справок на удаление", true);
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
        tip.SetToolTip(save, "Сохранить в папку «Справки на удаление» и открыть в Word (Ctrl+S)");
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
        int w = Math.Max(600, Math.Min(full - 32, 1060));
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
            ((DateBox)inp["data"]).Value = DateTime.Today;
            SetV("obekt", First(lists.Get("Удаление: объект")));
            SetV("vid", First(lists.Get("Удаление: в связи с продажей")));
            SetV("dve", memory.Get("УдалениеДвеКопии", "1"));
            string eng = memory.Get("УдалениеИнженер", memory.Get("СправкаИнженер", ""));
            SetV("inzh", lists.Get("Оформил").Contains(eng) ? eng : First(lists.Get("Оформил")));
            List<string> names = new List<string>();
            foreach (string[] sg in lists.Signers()) names.Add(sg[0]);
            SetV("nach", memory.Get("УдалениеНачальник", memory.Get("СправкаНачальник", First(names))));
            SetV("opsu", memory.Get("УдалениеНачальникОПСУ", ""));
            SetV("nach_dolzh", memory.Get("УдалениеДолжностьНачальника", First(lists.Get("Должность начальника"))));
            savedName = null;
        } finally { loading = false; }
        dirty = false;
        head.Text = "Новая справка в расчетную группу";
        headSub.ForeColor = St.Hint;
        headSub.Text = "Исключить ПУ проданной квартиры из расчётов юр лица";
        RefreshState();
    }

    // «Нераспроданные квартиры п.Солнечный,ул.Широкая,10» -> «Нераспроданные квартиры» (объект без адреса)
    public static string ObjectOf(string name) {
        Match m = Regex.Match(name ?? "", @"^(.*?)(?:[\s,]+(?:п|д|аг|г|гп|ул|пос|дер)\.|,|\s+\d)");
        string s = (m.Success ? m.Groups[1].Value : name ?? "").Trim().Trim(',', '.', ' ');
        return s;
    }

    // из «Перевода из юр в физ» (или данных юр лица, запомненных справкой): договор, карточка, наименование, объект, адрес, лицевой, счётчики
    public void FromData(Dictionary<string, string> jur, Dictionary<string, string> spravka) {
        Clear();
        string found = null;
        loading = true;
        try {
            if (spravka != null) {
                int n;
                if (!int.TryParse(Get(spravka, "pu_n"), out n)) n = 1;
                // 27.11: лицевые всех приборов учёта справки - через запятую (одинаковые - один раз)
                List<string> lss = new List<string>();
                for (int i = 1; i <= n; i++) { string l = BaseLink.Nab(Get(spravka, "pu" + i + "_ls")); if (l.Length > 0 && !lss.Contains(l)) lss.Add(l); }
                SetV("ls", string.Join(", ", lss.ToArray()));
                SetV("adres", Get(spravka, "pu1_adres"));
                List<string> lines = new List<string>();
                for (int i = 1; i <= n; i++) {
                    string p = "pu" + i + "_";
                    if (Get(spravka, p + "tip").Length == 0 && Get(spravka, p + "zav").Length == 0) continue;
                    lines.Add(Get(spravka, p + "tip") + "|" + Get(spravka, p + "zav") + "|" + Get(spravka, p + "pov") + "|" + Get(spravka, p + "pokaz"));
                }
                SetV("pu", string.Join("\n", lines.ToArray()));
                if (Get(spravka, "inzh").Length > 0) SetV("inzh", Get(spravka, "inzh"));
                if (Get(spravka, "nach").Length > 0) SetV("nach", Get(spravka, "nach"));
            }
            if (jur != null) {
                foreach (string k in new string[] { "kp", "nab", "dog_name" }) if (Get(jur, k).Length > 0) SetV(k, Get(jur, k));
                if (Get(jur, "obekt").Length > 0) SetV("obekt", ObjectOf(Get(jur, "obekt")));
                if (spravka == null) {
                    SetV("ls", Get(jur, "ls"));
                    SetV("adres", Get(jur, "adres"));
                    if (Get(jur, "pu").Length > 0) SetV("pu", Get(jur, "pu"));
                }
            }
            // 27.11: данных юр лица нет (справку открыли из сохранённых, вписали руками) - договор, карточка, наименование, объект
            // из юр базы по номерам счётчиков
            if (V("kp").Length == 0) found = FromJurBase();
        } finally { loading = false; }
        dirty = true;
        LayoutCards();
        head.Text = ("Справка в РГ " + V("kp") + (V("nab").Length > 0 ? "/" + V("nab") : "")).Trim();
        headSub.ForeColor = St.Hint;
        headSub.Text = (jur != null ? "Заполнено из перевода из юр в физ" : found != null ? "Заполнено из справки, юр лицо - из юр базы (" + found + ")" : "Заполнено из справки") +
            " - проверьте показания и наименование юр лица";
        RefreshState();
    }

    // договор, карточка, наименование юр лица, объект - из юр базы по номеру счётчика (первый найденный); null - не нашлось
    string FromJurBase() {
        try {
            for (int i = 0; i < pu.Count; i++) {
                string no = pu.Cell(i, 1);
                if (no.Length < 4) continue;
                List<JurRow> r = JurIndex.ByMeter(no);
                if (r.Count == 0) continue;
                JurRow x = r[0];
                SetV("kp", x.Kp);
                SetV("nab", x.Nab);
                SetV("dog_name", x.Contract);
                string ob = ObjectOf(x.Abon);
                if (ob.Length > 0) SetV("obekt", ob);
                return "счётчик № " + x.Meter;
            }
        } catch { }
        return null;
    }

    void OpenSaved() {
        string dir = DataFolder(here);
        if (!Directory.Exists(dir) || Directory.GetFiles(dir, "*.txt").Length == 0) { MessageBox.Show(this, "Сохранённых справок на удаление пока нет.", Text); return; }
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
            savedName = name;
        } finally { loading = false; }
        dirty = false;
        LayoutCards();
        head.Text = "Справка в РГ «" + name + "»";
        headSub.Text = "открыта сохранённая - исправьте и сохраните заново";
        RefreshState();
    }

    // «CE208BY S53 145.1 JPR.UKVFLZ, № 61811220, г.п. 2023-4; показание; 00003671 кВт*ч» - строка на счётчик
    string PuText() {
        List<string> lines = new List<string>();
        for (int i = 0; i < pu.Count; i++) {
            string tip = pu.Cell(i, 0), nom = pu.Cell(i, 1), pov = pu.Cell(i, 2), pok = pu.Cell(i, 3);
            if (tip.Length == 0 && nom.Length == 0) continue;
            List<string> s = new List<string>();
            if (tip.Length > 0) s.Add(tip);
            if (nom.Length > 0) s.Add("№ " + nom);
            if (pov.Length > 0) s.Add("г.п. " + pov);
            string line = string.Join(", ", s.ToArray());
            if (pok.Length > 0) line += "; показание; " + pok + " кВт*ч";
            lines.Add(line);
        }
        return string.Join("\n", lines.ToArray());
    }

    public Dictionary<string, string> Compose() {
        Dictionary<string, string> v = new Dictionary<string, string>();
        v["ДОГКАРТ"] = V("kp") + (V("nab").Length > 0 ? "/" + V("nab") : "");
        v["ДОГНОМ"] = V("kp");
        v["КАРТОЧКА"] = V("nab");
        v["ДОГНАИМ"] = V("dog_name");
        string ob = V("obekt");
        v["ОБЪЕКТ"] = ob.Length == 0 || ob.StartsWith("«") ? ob : "«" + ob + "»";
        v["ВИД"] = V("vid").Length > 0 ? V("vid") : "квартиры";
        v["ЛС"] = V("ls");
        v["АДРЕС"] = V("adres");
        v["ПУ"] = PuText();
        v["ИНЖЕНЕР"] = V("inzh");
        v["НАЧАЛЬНИК"] = V("nach");
        v["ОПСУ"] = V("opsu");
        v["НАЧДОЛЖН"] = V("nach_dolzh").Length > 0 ? V("nach_dolzh") : "Начальник ГРП";
        v["ДАТА"] = V("data");
        v["ГОД"] = ((DateBox)inp["data"]).Value.Year.ToString();
        return v;
    }

    List<string> Missing() {
        List<string> m = new List<string>();
        foreach (string k in Needed) if (k == "pu" ? PuText().Length == 0 : V(k).Length == 0) m.Add(k);
        return m;
    }

    static string Caption(string k) {
        switch (k) {
            case "kp": return "договор";
            case "nab": return "карточка";
            case "dog_name": return "наименование юр лица";
            case "ls": return "лицевой";
            case "adres": return "адрес";
            case "pu": return "прибор учёта";
            case "inzh": return "инженер";
            case "nach": return "начальник";
            case "opsu": return "начальник ОПСУ";
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

    // как у справок: «п. Макетино, ул. Полевая, д. 34, кв. 12» -> «Макетино Полевая 34-12 удаление»
    public string FileName() {
        string[] skip = { "г", "д", "аг", "п", "гп", "пос", "ул", "пер", "пр", "пр-т", "б-р", "пл", "корп", "к", "дом", "с/с", "р-н", "обл", "уч" };
        List<string> w = new List<string>();
        bool kv = false;
        foreach (string raw in V("adres").Replace(",", " ").Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
            string t = raw.Trim('.', ' ');
            string low = t.ToLowerInvariant();
            if (low == "кв") { kv = true; continue; }
            if (Array.IndexOf(skip, low) >= 0 || t.Length == 0) continue;
            if (kv && w.Count > 0) { w[w.Count - 1] += "-" + t; kv = false; continue; }
            w.Add(t);
        }
        string n = (string.Join(" ", w.ToArray()) + " удаление").Trim();
        foreach (char ch in Path.GetInvalidFileNameChars()) n = n.Replace(ch, '_');
        n = n.Replace(",", "").Replace("«", "").Replace("»", "");
        while (n.Contains("  ")) n = n.Replace("  ", " ");
        n = n.Trim().TrimEnd('.');
        return n.Length > "удаление".Length ? n : "Удаление " + DateTime.Now.ToString("yyyy-MM-dd HH-mm");
    }

    bool HasTemplate() {
        if (File.Exists(templateFile)) return true;
        MessageBox.Show(this, "Нет файла шаблона:\n" + templateFile + "\n\nОн должен лежать рядом с программой (есть в архиве программы).", Text,
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return false;
    }

    void Fill(string output) {
        DocxTemplate.Fill(templateFile, output, Compose());
        if (V("dve") == "1") DocxTemplate.TwoCopies(output);
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
            StringBuilder sb = new StringBuilder("# Данные справки в расчетную группу (удаление ПУ из юр) для программы «Оформление договора»\r\n");
            sb.Append("сохранён=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
            foreach (string k in inp.Keys) sb.Append(k).Append('=').Append(Memory.Escape(V(k))).Append("\r\n");
            sb.Append("pu=").Append(Memory.Escape(pu.Value)).Append("\r\n");
            File.WriteAllText(Path.Combine(DataFolder(here), name + ".txt"), sb.ToString(), new UTF8Encoding(true));
        } catch (IOException ex) {
            MessageBox.Show(this, "Справка не сохранена (открыта в Word?):\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        memory.Remember("удаление юр лицо", V("dog_name"));
        memory.Remember("удаление объект", V("obekt"));
        memory.Values["УдалениеДвеКопии"] = V("dve");
        memory.Values["УдалениеИнженер"] = V("inzh");
        memory.Values["УдалениеНачальник"] = V("nach");
        memory.Values["УдалениеНачальникОПСУ"] = V("opsu");
        memory.Values["УдалениеДолжностьНачальника"] = V("nach_dolzh");
        memory.Save();
        savedName = name;
        dirty = false;
        head.Text = "Справка в РГ «" + name + "»";
        RefreshState();
        if (openWord) try { DocPreview.Open(docx); } catch { }
        return docx;
    }

    public bool AskSaveBeforeExit() {
        if (!dirty) return true;
        DialogResult r = MessageBox.Show(this, "Справка в расчетную группу не сохранена. Сохранить её?\n\nДа - сохранить и выйти\nНет - выйти без сохранения\nОтмена - вернуться", Text,
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) return Save(false, true) != null;
        return true;
    }

    public string StatusText { get { return status.Text; } }
}
// ==== КОНЕЦ БЛОКА СПРАВКИ НА УДАЛЕНИЕ ====
