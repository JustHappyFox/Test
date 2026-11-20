// ==== БЛОК ЮР НАРЯДА (исходник - razrabotka/src/jur_naryad.cs; здесь не править - вставляется скриптом inject.py) ====
// Наряд юр лица на включение (отключение) - лист A4, по фото бланка 20.11 (шаблон jur_naryad.docx рядом с программой).
// Пока только заполнение: поля, списки, предпросмотр, сохранение в папку «Юр наряды», открыть сохранённый; данные можно
// взять из сохранённого пакета БС.
public class JurNaryadForm : Form {
    readonly string here, templateFile;
    readonly Lists lists;
    readonly Memory memory;
    readonly Dictionary<string, Control> inp = new Dictionary<string, Control>();
    readonly List<Card> cards = new List<Card>();
    readonly Panel content = new Panel();
    readonly Label head = new Label(), headSub = new Label(), status = new Label();
    readonly ToolTip tip = new ToolTip();
    public StartMenu Home;
    public Form Back;
    bool loading, dirty;
    string savedName;
    static readonly string[] Months = { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };
    static readonly string[] Needed = { "potr", "mesto", "prichina", "adres", "obekty", "nach", "inzh" };

    public bool Unsaved { get { return dirty; } }
    public static string Folder(string here) { return Path.Combine(here, "Юр наряды"); }
    public static string DataFolder(string here) { return Path.Combine(Folder(here), "данные"); }

    public JurNaryadForm(string here) {
        Load += delegate { St.Dress(this); };
        this.here = here;
        templateFile = Path.Combine(here, "jur_naryad.docx");
        lists = new Lists(Path.Combine(here, "spravochniki.txt"));
        memory = Memory.For(Path.Combine(here, "nastroiki.txt"));
        Text = "Юр наряд на включение / отключение";
        Font = St.Base;
        BackColor = St.Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1180, wa.Width), Math.Min(820, wa.Height));
        MinimumSize = new Size(860, 560);
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
            Crash.Mark("юр наряд, поле " + key);
            dirty = true;
            if (key == "s_datoy") Enable();
            RefreshState();
        };
        if (c is Seg) ((Seg)c).Changed += h;
        else if (c is ComboBox) { ((ComboBox)c).TextChanged += h; ((ComboBox)c).SelectedIndexChanged += h; }
        else if (c is DateBox) ((DateBox)c).ValueChanged += h;
        else if (c is CheckBox) ((CheckBox)c).CheckedChanged += h;
        else c.TextChanged += h;
        if (c is TextBox || c is ComboBox) {
            c.Enter += delegate { c.BackColor = St.Focus; };
            c.Leave += delegate { c.BackColor = c.Enabled ? St.Input : St.Disabled; };
        }
    }

    public string V(string key) {
        Control c;
        if (!inp.TryGetValue(key, out c)) return "";
        if (c is Seg) return ((Seg)c).Value;
        if (c is DateBox) return ((DateBox)c).Text2;
        if (c is CheckBox) return ((CheckBox)c).Checked ? "1" : "";
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

    Cell F(string caption, Control c, int chars) { Cell x = Cell.F(caption, c, chars == 0 ? 4 : 1); x.W = chars; return x; }

    void BuildCards() {
        Card c = new Card("Наряд", "№ можно не писать - будет линия, впишут от руки; дату - по галочке", 4);
        cards.Add(c);
        Seg vid = new Seg("включение", "отключение");
        Reg("vid", vid);
        DateBox date = new DateBox();
        Reg("data", date);
        CheckBox dated = new CheckBox();
        dated.Text = "указать дату";
        dated.Font = St.Base;
        dated.ForeColor = St.Text;
        dated.BackColor = Color.Transparent;
        Reg("s_datoy", dated);
        c.Grid.Row(F("Наряд на", vid, 24), F("№ наряда", T("nomer", "пусто - линия"), 12), F(null, dated, 15), F("Дата", date, 14), F("РЭС", C("res", "РЭС-3"), 0));

        c = new Card("Потребитель и объект", "наименование, место и причина включения, адрес, объекты", 4);
        cards.Add(c);
        c.Grid.Row(F("Наименование предприятия потребителя", C("potr", "ООО «...»"), 0));
        c.Grid.Row(F("Место включения (отключения)", C("mesto", "РП-301 с.1 яч.17, с.2 яч.18"), 0), F("Причина", C("prichina", "Допуск в эксплуатацию"), 30));
        c.Grid.Row(F("Адрес (с подстанцией)", T("adres", "д. ..., ул. ... (КТПБ-...)"), 0));
        TextBox ob = T("obekty", null);
        ob.Multiline = true;
        ob.ScrollBars = ScrollBars.Vertical;
        ob.Tag = 4;
        c.Grid.Row(F("Следующих объектов (жирным в наряде)", ob, 0));

        c = new Card("Подписи", "должностные лица, выдавшие наряд", 4);
        cards.Add(c);
        c.Grid.Row(F("Должность начальника", C("nach_dolzh", null), 0), F("Ф.И.О.", C("nach", null), 20));
        c.Grid.Row(F("Должность инженера", C("inzh_dolzh", null), 0), F("Ф.И.О.", C("inzh", null), 20));
    }

    void FillLists() {
        bool was = loading;
        loading = true;
        try {
            Items("res", lists.Get("Юр наряд: РЭС"));
            Items("potr", memory.Hist("юр наряд потребитель"));
            Items("mesto", memory.Hist("юр наряд место"));
            Items("prichina", With(lists.Get("Наряд: причина"), memory.Hist("наряд причина")));
            Items("nach_dolzh", lists.Get("Юр наряд: должность начальника"));
            Items("inzh_dolzh", lists.Get("Юр наряд: должность инженера"));
            List<string> n = new List<string>();
            foreach (string[] s in lists.Signers()) n.Add(s[0]);
            Items("nach", n);
            List<string> e = new List<string>();
            foreach (string s in lists.Get("БС: инженер")) { int k = s.IndexOf('='); e.Add((k > 0 ? s.Substring(0, k) : s).Trim()); }
            foreach (string s in lists.Get("Оформил")) if (!e.Contains(s)) e.Add(s);
            Items("inzh", e);
        } finally { loading = was; }
    }

    void Items(string key, List<string> items) { Combo c = inp[key] as Combo; if (c != null) c.SetItems(items); }

    static List<string> With(List<string> a, List<string> b) {
        List<string> r = new List<string>(a);
        foreach (string s in b) if (!r.Contains(s)) r.Add(s);
        return r;
    }

    static string First(List<string> l) { return l.Count > 0 ? l[0] : ""; }

    void Enable() { inp["data"].Enabled = V("s_datoy") == "1"; }

    Control BuildNav() {
        Panel nav = new Panel();
        nav.Dock = DockStyle.Left;
        nav.Width = 252;
        nav.BackColor = St.Nav;
        Label title = new Label();
        title.Text = "Юр наряд";
        title.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
        title.ForeColor = Color.White;
        title.SetBounds(22, 22, 210, 34);
        Label sub = new Label();
        sub.Text = "на включение / отключение, A4";
        sub.Font = St.Small;
        sub.ForeColor = St.NavText;
        sub.SetBounds(24, 58, 210, 20);
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
        NavItem openItem = new NavItem("Открыть сохранённый…", true);
        openItem.Click += delegate { OpenSaved(); };
        NavItem folder = new NavItem("Папка юр нарядов", true);
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
        Button fromBs = Ui.Button("Из пакета БС…", false);
        fromBs.SetBounds(24, 14, 150, 38);
        fromBs.Click += delegate { FromBs(); };
        tip.SetToolTip(fromBs, "Взять потребителя, место, адрес, объект и подписи из сохранённого пакета БС");
        Button clear = Ui.Button("Очистить", false);
        clear.SetBounds(184, 14, 100, 38);
        clear.Click += delegate {
            if (dirty && MessageBox.Show(this, "Очистить наряд? Введённое пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Clear();
        };
        Button look = Ui.Button("Предпросмотр", false);
        look.Size = new Size(140, 38);
        look.Click += delegate { Preview(); };
        Button save = Ui.Button("Сохранить наряд", true);
        save.Size = new Size(200, 40);
        save.Click += delegate { Save(true, true); };
        status.AutoSize = false;
        status.Font = St.Base;
        status.TextAlign = ContentAlignment.MiddleLeft;
        p.Controls.Add(fromBs);
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
            foreach (KeyValuePair<string, Control> kv in inp) if (!(kv.Value is DateBox) && !(kv.Value is Seg)) SetV(kv.Key, "");
            ((DateBox)inp["data"]).Value = DateTime.Today;
            SetV("s_datoy", memory.Get("ЮрНарядСДатой", "1"));
            SetV("vid", "включение");
            SetV("res", memory.Get("ЮрНарядРЭС", First(lists.Get("Юр наряд: РЭС"))));
            SetV("prichina", First(lists.Get("Наряд: причина")));
            SetV("nach_dolzh", memory.Get("ЮрНарядДолжностьНачальника", First(lists.Get("Юр наряд: должность начальника"))));
            SetV("inzh_dolzh", memory.Get("ЮрНарядДолжностьИнженера", First(lists.Get("Юр наряд: должность инженера"))));
            SetV("nach", memory.Get("ЮрНарядНачальник", memory.Get("БСНачальник", "")));
            SetV("inzh", memory.Get("ЮрНарядИнженер", memory.Get("БСИнженер", "")));
            savedName = null;
        } finally { loading = false; }
        dirty = false;
        Enable();
        head.Text = "Новый юр наряд";
        headSub.ForeColor = St.Hint;
        headSub.Text = "Наряд юр лица на включение (отключение) - лист A4";
        RefreshState();
    }

    // из сохранённого пакета БС: потребитель, место включения, адрес с подстанцией, объект, подписи
    void FromBs() {
        string dir = BsForm.DataFolder(here);
        if (!Directory.Exists(dir) || Directory.GetFiles(dir, "*.txt").Length == 0) { MessageBox.Show(this, "Сохранённых пакетов БС пока нет.", Text); return; }
        if (dirty && MessageBox.Show(this, "Заменить введённое данными пакета БС?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        using (SavedDialog dlg = new SavedDialog(dir, 3)) {
            if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Chosen == null) return;
            try { FromBsData(DogovorForm.ReadData(dlg.Chosen)); }
            catch (Exception ex) { MessageBox.Show(this, "Не удалось открыть:\n" + ex.Message, Text); }
        }
    }

    public void FromBsData(Dictionary<string, string> d) {
        string v;
        loading = true;
        try {
            if (d.TryGetValue("abonent", out v)) SetV("potr", v);
            if (d.TryGetValue("mesto", out v)) SetV("mesto", v);
            string adres = d.TryGetValue("adres_kr", out v) ? v : "", pit = d.TryGetValue("pitanie", out v) ? v : "";
            SetV("adres", adres + (pit.Length > 0 ? " (" + pit + ")" : ""));
            string ob = d.TryGetValue("obekt_full", out v) ? v.Trim('«', '»', ' ') : "";
            if (ob.Length > 0) SetV("obekty", ob);
            if (d.TryGetValue("inzh", out v) && v.Length > 0) SetV("inzh", v);
            if (d.TryGetValue("nach", out v) && v.Length > 0) SetV("nach", v);
            if (d.TryGetValue("prichina", out v) && v.Length > 0) SetV("prichina", v);
        } finally { loading = false; }
        dirty = true;
        headSub.ForeColor = St.Hint;
        headSub.Text = "Заполнено из пакета БС - проверьте место включения и объекты";
        RefreshState();
    }

    void OpenSaved() {
        string dir = DataFolder(here);
        if (!Directory.Exists(dir) || Directory.GetFiles(dir, "*.txt").Length == 0) { MessageBox.Show(this, "Сохранённых юр нарядов пока нет.", Text); return; }
        if (dirty && MessageBox.Show(this, "Наряд не сохранён. Открыть другой? Введённое здесь пропадёт.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        using (SavedDialog dlg = new SavedDialog(dir, 2)) {
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
            savedName = name;
        } finally { loading = false; }
        dirty = false;
        Enable();
        head.Text = "Юр наряд «" + name + "»";
        headSub.Text = "открыт сохранённый - исправьте и сохраните заново";
        RefreshState();
    }

    string DateText(bool corner) {
        DateTime d = ((DateBox)inp["data"]).Value;
        if (V("s_datoy") != "1") return corner ? "«____» ______________ 20____ г." : "«____» ______________ " + DateTime.Today.Year + " г.";
        return "« " + d.Day + " » " + Months[d.Month - 1] + " " + d.Year + " г.";
    }

    public Dictionary<string, string> Compose() {
        Dictionary<string, string> v = new Dictionary<string, string>();
        bool on = V("vid") != "отключение";
        v["ДАТАН"] = DateText(false);
        v["ДАТАНК"] = DateText(true);
        v["НОМЕР"] = V("nomer").Length > 0 ? V("nomer") : "____________";
        v["РЭС"] = V("res");
        v["ВИД"] = on ? "включение" : "отключение";
        v["ВИДА"] = on ? "включения" : "отключения";
        v["ВИДП"] = on ? "подключение" : "отключение";
        v["ВИДА2"] = on ? "подключения" : "отключения";
        v["ПОТРЕБИТЕЛЬ"] = V("potr");
        v["МЕСТО"] = V("mesto");
        v["ПРИЧИНА"] = V("prichina");
        v["АДРЕС"] = V("adres");
        v["ОБЪЕКТЫ"] = V("obekty").Replace("\r\n", " ").Replace("\n", " ");
        v["НАЧДОЛЖН"] = V("nach_dolzh");
        v["НАЧАЛЬНИК"] = V("nach");
        v["ИНЖДОЛЖН"] = V("inzh_dolzh");
        v["ИНЖЕНЕР"] = V("inzh");
        return v;
    }

    List<string> Missing() {
        List<string> m = new List<string>();
        foreach (string k in Needed) if (V(k).Length == 0) m.Add(k);
        return m;
    }

    static string Caption(string k) {
        switch (k) {
            case "potr": return "потребитель";
            case "mesto": return "место включения";
            case "prichina": return "причина";
            case "adres": return "адрес";
            case "obekty": return "объекты";
            case "nach": return "начальник";
            case "inzh": return "инженер";
            default: return k;
        }
    }

    void RefreshState() {
        List<string> m = Missing();
        if (m.Count == 0) { status.Text = savedName != null && !dirty ? "Наряд сохранён" : "Всё основное заполнено"; status.ForeColor = St.Good; return; }
        List<string> names = new List<string>();
        foreach (string k in m) names.Add(Caption(k));
        status.Text = "Не заполнено: " + string.Join(", ", names.ToArray());
        status.ForeColor = St.Bad;
    }

    public string FileName() {
        // «ООО Тестовая связь д. Тестовка, ул. Лесная 5» - адрес без подстанции в скобках, не длиннее 80 знаков (по слову)
        string a = V("adres");
        int br = a.LastIndexOf(" (");
        if (br > 0 && a.EndsWith(")")) a = a.Substring(0, br);
        string n = (V("potr") + " " + a).Trim();
        if (n.Length > 80) { n = n.Substring(0, 80); int sp = n.LastIndexOf(' '); if (sp > 40) n = n.Substring(0, sp); }
        if (V("vid") == "отключение") n += " отключение";
        foreach (char ch in Path.GetInvalidFileNameChars()) n = n.Replace(ch, '_');
        n = n.Replace('«', ' ').Replace('»', ' ').Replace('"', ' ');
        while (n.Contains("  ")) n = n.Replace("  ", " ");
        n = n.Trim().TrimEnd('.');
        return n.Length > 0 ? n : "Юр наряд " + DateTime.Now.ToString("yyyy-MM-dd HH-mm");
    }

    bool HasTemplate() {
        if (File.Exists(templateFile)) return true;
        MessageBox.Show(this, "Нет файла шаблона юр наряда:\n" + templateFile + "\n\nОн должен лежать рядом с программой (есть в архиве программы).", Text,
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return false;
    }

    public string Preview() {
        if (!HasTemplate()) return null;
        try {
            string f = DocPreview.NewFile(here, FileName());
            DocxTemplate.Fill(templateFile, f, Compose());
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
            if (MessageBox.Show(this, "Не заполнено:\n" + string.Join("\n", names.ToArray()) + "\n\nВсё равно сохранить наряд?", "Проверка перед сохранением",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
        }
        if (!HasTemplate()) return null;
        string name = FileName();
        string dir = Folder(here);
        string docx = Path.Combine(dir, name + ".docx");
        try {
            Directory.CreateDirectory(dir);
            if (ask && savedName != name && File.Exists(docx) && MessageBox.Show(this, "Наряд «" + name + "» уже есть. Заменить его?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
            string tmp = docx + ".new";
            DocxTemplate.Fill(templateFile, tmp, Compose());
            if (File.Exists(docx)) File.Delete(docx);
            File.Move(tmp, docx);
            // данные - открыть заново
            Directory.CreateDirectory(DataFolder(here));
            StringBuilder sb = new StringBuilder("# Данные юр наряда для программы «Оформление договора»\r\n");
            sb.Append("сохранён=").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append("\r\n");
            foreach (string k in inp.Keys) sb.Append(k).Append('=').Append(Memory.Escape(V(k))).Append("\r\n");
            File.WriteAllText(Path.Combine(DataFolder(here), name + ".txt"), sb.ToString(), new UTF8Encoding(true));
        } catch (IOException ex) {
            MessageBox.Show(this, "Наряд не сохранён (открыт в Word?):\n" + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        memory.Remember("юр наряд потребитель", V("potr"));
        memory.Remember("юр наряд место", V("mesto"));
        memory.Values["ЮрНарядСДатой"] = V("s_datoy");
        memory.Values["ЮрНарядРЭС"] = V("res");
        memory.Values["ЮрНарядДолжностьНачальника"] = V("nach_dolzh");
        memory.Values["ЮрНарядДолжностьИнженера"] = V("inzh_dolzh");
        memory.Values["ЮрНарядНачальник"] = V("nach");
        memory.Values["ЮрНарядИнженер"] = V("inzh");
        memory.Save();
        savedName = name;
        dirty = false;
        head.Text = "Юр наряд «" + name + "»";
        RefreshState();
        if (openWord) try { DocPreview.Open(docx); } catch { }
        return docx;
    }

    public bool AskSaveBeforeExit() {
        if (!dirty) return true;
        DialogResult r = MessageBox.Show(this, "Юр наряд не сохранён. Сохранить его?\n\nДа - сохранить и выйти\nНет - выйти без сохранения\nОтмена - вернуться", Text,
            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) return Save(false, true) != null;
        return true;
    }

    public string StatusText { get { return status.Text; } }
}
// ==== КОНЕЦ БЛОКА ЮР НАРЯДА ====
