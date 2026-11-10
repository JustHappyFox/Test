# Меню программ: одно окно, из которого запускаются «Поиск абонента», «Оформление договора» и «Фото с телефона».
# Запуск: zapusk.vbs (без чёрного окна) или zapusk.bat. Ярлык на рабочий стол - в самом меню.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @([System.Windows.Forms.Form].Assembly.Location, [System.Drawing.Point].Assembly.Location)

$code = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

// Меню программ: одно окно (и один ярлык на рабочем столе), из которого запускаются «Поиск абонента», «Оформление договора»
// и «Фото с телефона». Меню только запускает их zapusk.vbs - в папках программ ничего не пишет. Где какая программа -
// находит само (рядом с собой, на рабочем столе, в «Документах», на дисках) или указывается вручную; запоминается в nastroiki.txt
public static class MenuApp {
    public static string Here = "";

    public static void Run(string here, string launch) {
        Here = here;
        try { Application.EnableVisualStyles(); } catch { }
        try { Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException); } catch { }
        Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e) { Fail(e.Exception); };
        Application.Run(new MenuForm(here));
    }

    // ошибка - в oshibki.txt рядом с меню и окно (номера от 6 цифр скрыты)
    public static void Fail(Exception ex) {
        string text = System.Text.RegularExpressions.Regex.Replace(ex.ToString(), @"\d{6,}", "######");
        try {
            File.AppendAllText(Path.Combine(Here, "oshibki.txt"), "=== " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " ===\r\nWindows " + Environment.OSVersion.Version
                + ", .NET " + Environment.Version + "\r\n" + text + "\r\n\r\n", new UTF8Encoding(true));
        } catch { }
        MessageBox.Show("Ошибка: " + ex.Message + "\n\nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", "Программы - ошибка");
    }
}

// цвета - общей темы программ (Look: тёмная / светлая, как в «Договоре»)
public static class Pal {
    public static Color Back { get { return Look.Back; } }
    public static Color Tile { get { return Look.Surface; } }
    public static Color TileHover { get { return Look.Hover; } }
    public static Color Border { get { return Look.Border; } }
    public static Color Text { get { return Look.Text; } }
    public static Color Hint { get { return Look.Hint; } }
    public static Color Bad { get { return Look.Bad; } }
    public static Color Good { get { return Look.Good; } }
    public static Color Link { get { return Look.Link; } }
    public static readonly Font Title = new Font("Segoe UI", 17f, FontStyle.Bold), Big = new Font("Segoe UI", 13.5f, FontStyle.Bold),
        Base = new Font("Segoe UI", 10f), Small = new Font("Segoe UI", 9f), Number = new Font("Segoe UI", 22f, FontStyle.Bold);
}

// программа меню: скрипт, по которому узнаётся её папка, и файлы, которые появляются, когда ею пользуются
public class Prog {
    public string Key, Title, About, Script;
    public string[] Work;            // рабочие файлы/папки: такая папка - настоящая, а не свежераспакованная копия
    public Color Mark;
    public string Dir = "";          // где лежит (пусто - не найдена)
    public bool Auto;                // найдена сама

    public static Prog[] All() {
        Prog a = new Prog(), b = new Prog(), c = new Prog();
        a.Key = "Поиск"; a.Title = "Поиск абонента"; a.Script = "prosmotr.ps1"; a.Mark = Color.FromArgb(99, 160, 235);
        a.About = "лицевой, адрес, счётчик, Ф.И.О.; карточка, наряды, обходы, оплаты";
        a.Work = new string[] { "poisk.txt", "indeks_fio.dat" };
        b.Key = "Договор"; b.Title = "Оформление договора"; b.Script = "dogovor.ps1"; b.Mark = Color.FromArgb(112, 200, 142);
        b.About = "договор, справка, наряд на подключение, отчёты пломб";
        b.Work = new string[] { "nastroiki.txt", "Договоры", "Справки" };
        c.Key = "Фото"; c.Title = "Фото с телефона"; c.Script = "foto.ps1"; c.Mark = Color.FromArgb(230, 170, 90);
        c.About = "фото, скрины, видео с телефона - в папки год\\месяц\\число";
        c.Work = new string[] { "nastroiki.txt", "telefony.txt" };
        return new Prog[] { a, b, c };
    }

    public bool Fits(string dir) {
        try { return dir != null && dir.Length > 1 && File.Exists(Path.Combine(dir, Script)); } catch { return false; }
    }

    // чем больше рабочих файлов и чем свежее - тем вероятнее, что это та папка, где работают
    public long Score(string dir) {
        int work = 0;
        DateTime t = DateTime.MinValue;
        foreach (string w in Work) {
            try {
                string p = Path.Combine(dir, w);
                DateTime m = DateTime.MinValue;
                if (File.Exists(p)) m = File.GetLastWriteTime(p);
                else if (Directory.Exists(p)) m = Directory.GetLastWriteTime(p);
                else continue;
                work++;
                if (m > t) t = m;
            } catch { }
        }
        try { DateTime s = File.GetLastWriteTime(Path.Combine(dir, Script)); if (work == 0 && s > t) t = s; } catch { }
        return work * 100000000000000L + Math.Max(0L, t.Ticks / 10000000L);
    }
}

// Поиск папок программ: папка меню и соседние, на ступень выше, рабочий стол, «Документы», корни дисков - по две ступени вглубь
public static class Finder {
    public static void Find(string here, Prog[] progs) {
        List<string> roots = new List<string>();
        try {
            string full = Path.GetFullPath(here);
            roots.Add(full);
            string up = Path.GetDirectoryName(full);
            if (up != null) { roots.Add(up); string up2 = Path.GetDirectoryName(up); if (up2 != null) roots.Add(up2); }
        } catch { }
        try { roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)); } catch { }
        try { roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.Personal)); } catch { }
        try { foreach (DriveInfo d in DriveInfo.GetDrives()) if (d.DriveType == DriveType.Fixed && d.IsReady) roots.Add(d.RootDirectory.FullName); } catch { }
        long[] best = new long[progs.Length];
        string[] found = new string[progs.Length];
        Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        int budget = 6000;
        foreach (string r in roots) if (r != null && r.Length > 0) Walk(r, 2, progs, found, best, seen, ref budget);
        for (int i = 0; i < progs.Length; i++) {
            if (found[i] == null) continue;
            progs[i].Dir = found[i];
            progs[i].Auto = true;
        }
    }

    static void Walk(string dir, int depth, Prog[] progs, string[] found, long[] best, Dictionary<string, bool> seen, ref int budget) {
        if (budget-- <= 0 || seen.ContainsKey(dir)) return;
        seen[dir] = true;
        for (int i = 0; i < progs.Length; i++) {
            if (progs[i].Dir.Length > 0 || !progs[i].Fits(dir)) continue;
            long s = progs[i].Score(dir);
            if (found[i] == null || s > best[i]) { found[i] = dir; best[i] = s; }
        }
        if (depth == 0) return;
        string[] subs;
        try { subs = Directory.GetDirectories(dir); } catch { return; }
        foreach (string sub in subs) {
            string n = Path.GetFileName(sub).ToLowerInvariant();
            if (n.StartsWith("$") || n.StartsWith(".") || n == "windows" || n == "program files" || n == "program files (x86)" || n == "programdata"
                || n == "system volume information" || n == "recovery" || n == "appdata" || n == "proc" || n == "sys"
                || n == "договоры" || n == "справки" || n == "наряды" || n == "пломбы" || n == "kesh-programmy") continue;
            try { if ((File.GetAttributes(sub) & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue; } catch { continue; }
            Walk(sub, depth - 1, progs, found, best, seen, ref budget);
        }
    }
}

// настройки меню: «ключ=значение» в nastroiki.txt
public class MenuSettings {
    readonly string file;
    public readonly Dictionary<string, string> Values = new Dictionary<string, string>();

    public MenuSettings(string file) {
        this.file = file;
        try {
            if (File.Exists(file))
                foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                    int e = line.IndexOf('=');
                    if (e > 0 && !line.StartsWith("#")) Values[line.Substring(0, e).Trim()] = line.Substring(e + 1).Trim();
                }
        } catch { }
    }

    public string Get(string k) { string v; return Values.TryGetValue(k, out v) ? v : ""; }

    public void Save() {
        try {
            List<string> lines = new List<string>();
            lines.Add("# Меню программ: где какая программа. Файл можно удалить - меню найдёт программы заново.");
            foreach (KeyValuePair<string, string> kv in Values) lines.Add(kv.Key + "=" + kv.Value);
            File.WriteAllLines(file, lines.ToArray(), new UTF8Encoding(true));
        } catch { }
    }
}

// плитка программы: номер, название, что умеет, где лежит; щелчок - запуск
public class ProgTile : Control {
    public Prog P;
    public int Index;
    public string State = "";
    public Color StateColor = Pal.Hint;
    bool hover;
    public readonly LinkLabel Folder = new LinkLabel();

    public ProgTile(Prog p, int index) {
        P = p;
        Index = index;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        Cursor = Cursors.Hand;
        BackColor = Pal.Tile;
        Folder.Text = "папка…";
        Folder.Font = Pal.Small;
        Folder.AutoSize = true;
        Folder.LinkColor = Folder.ActiveLinkColor = Pal.Link;
        Folder.BackColor = Color.Transparent;
        Folder.LinkBehavior = LinkBehavior.HoverUnderline;
        Controls.Add(Folder);
        Resize += delegate { Folder.Location = new Point(Width - Folder.Width - 14, Height - Folder.Height - 10); };
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.Clear(hover ? Pal.TileHover : Pal.Tile);
        using (Pen pen = new Pen(Pal.Border)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        using (SolidBrush b = new SolidBrush(P.Mark)) g.FillRectangle(b, 0, 0, 6, Height);
        TextRenderer.DrawText(g, (Index + 1).ToString(), Pal.Number, new Point(20, (Height - 40) / 2), P.Mark);
        TextRenderer.DrawText(g, P.Title, Pal.Big, new Point(62, 10), Pal.Text);
        TextRenderer.DrawText(g, P.About, Pal.Base, new Point(64, 40), Pal.Hint);
        Rectangle st = new Rectangle(64, Height - 28, Width - 64 - Folder.Width - 30, 20);
        TextRenderer.DrawText(g, State, Pal.Small, st, StateColor, TextFormatFlags.Left | TextFormatFlags.PathEllipsis | TextFormatFlags.SingleLine);
    }
}

public class MenuForm : Form {
    readonly string here;
    readonly MenuSettings settings;
    readonly Prog[] progs = Prog.All();
    readonly List<ProgTile> tiles = new List<ProgTile>();
    readonly CheckBox closeAfter = new CheckBox();
    readonly Label status = new Label();
    readonly LinkLabel shortcut = new LinkLabel(), again = new LinkLabel();
    LinkLabel theme;
    bool searching;

    public delegate void Starter(string file, string args);
    public static Starter Start;          // для проверки: вместо запуска - сюда

    public MenuForm(string here) {
        this.here = here;
        settings = new MenuSettings(Path.Combine(here, "nastroiki.txt"));
        // тема (nastroiki.txt «ТёмнаяТема»; нет записи - тёмная) - до создания элементов
        Look.Apply(settings.Get("ТёмнаяТема") != "0");
        Text = "Программы";
        Font = Pal.Base;
        BackColor = Pal.Back;
        ForeColor = Pal.Text;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(640, 486);
        KeyPreview = true;
        try { Icon = MenuIcon.Make(); } catch { }
        Label title = new Label();
        title.Text = "Программы";
        title.Font = Pal.Title;
        title.ForeColor = Pal.Text;
        title.SetBounds(28, 18, 400, 36);
        Label sub = new Label();
        sub.Text = "щёлкните программу (или нажмите 1, 2, 3)";
        sub.Font = Pal.Small;
        sub.ForeColor = Pal.Hint;
        sub.SetBounds(30, 54, 500, 20);
        Controls.Add(title);
        Controls.Add(sub);
        for (int i = 0; i < progs.Length; i++) {
            ProgTile t = new ProgTile(progs[i], i);
            t.SetBounds(28, 84 + i * 104, 584, 92);
            int k = i;
            t.Click += delegate { Launch(k); };
            t.Folder.LinkClicked += delegate { PickFolder(k); };
            tiles.Add(t);
            Controls.Add(t);
        }
        closeAfter.Text = "Закрывать меню после запуска программы";
        closeAfter.ForeColor = Pal.Text;
        closeAfter.Font = Pal.Small;
        closeAfter.SetBounds(28, 396, 330, 24);
        closeAfter.Checked = settings.Get("Закрывать") != "0";
        closeAfter.CheckedChanged += delegate { settings.Values["Закрывать"] = closeAfter.Checked ? "1" : "0"; settings.Save(); };
        Controls.Add(closeAfter);
        foreach (LinkLabel l in new LinkLabel[] { shortcut, again }) {
            l.Font = Pal.Small;
            l.AutoSize = true;
            l.LinkColor = l.ActiveLinkColor = Pal.Link;
            l.LinkBehavior = LinkBehavior.HoverUnderline;
            Controls.Add(l);
        }
        shortcut.LinkClicked += delegate { MakeShortcut(); };
        theme = Look.Switch(delegate(bool dark) { settings.Values["ТёмнаяТема"] = dark ? "1" : "0"; settings.Save(); ThemePlace(); });
        Controls.Add(theme);
        again.Text = "Искать программы заново";
        again.LinkClicked += delegate { Search(true); };
        status.Font = Pal.Small;
        status.ForeColor = Pal.Hint;
        status.SetBounds(28, 454, 584, 22);
        Controls.Add(status);
        ShortcutText();
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) Close();
            int n = e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D3 ? e.KeyCode - Keys.D1 : e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad3 ? e.KeyCode - Keys.NumPad1 : -1;
            if (n >= 0) { e.SuppressKeyPress = true; Launch(n); }
        };
        // запомненные папки; нет или программы там уже нет - поискать
        bool missing = false;
        foreach (Prog p in progs) {
            string d = settings.Get(p.Key);
            if (p.Fits(d)) p.Dir = d;
            else missing = true;
        }
        Show2();
        Shown += delegate { if (missing) Search(false); };
    }

    void Place(LinkLabel l, int x, int y) { l.Location = new Point(x, y); }

    void ShortcutText() {
        shortcut.Text = File.Exists(ShortcutPath) ? "Ярлык «Программы» на рабочем столе есть - обновить" : "Создать ярлык «Программы» на рабочем столе";
        Place(shortcut, 30, 426);
        ThemePlace();
        Place(again, ClientSize.Width - 28 - TextRenderer.MeasureText(again.Text, Pal.Small).Width - 4, 400);
    }

    void ThemePlace() { if (theme != null) Place(theme, ClientSize.Width - 28 - TextRenderer.MeasureText(theme.Text, Look.Small).Width - 4, 426); }

    // плитки: где лежит программа
    void Show2() {
        for (int i = 0; i < progs.Length; i++) {
            Prog p = progs[i];
            ProgTile t = tiles[i];
            if (p.Dir.Length > 0) { t.State = "папка: " + p.Dir + (p.Auto ? " (найдена сама)" : ""); t.StateColor = Pal.Hint; }
            else if (searching) { t.State = "ищу папку программы..."; t.StateColor = Pal.Hint; }
            else { t.State = "не найдена - укажите папку (где " + p.Script + "): «папка…»"; t.StateColor = Pal.Bad; }
            t.Invalidate();
        }
    }

    // поиск папок программ - в фоне (диски - секунды)
    public void Search(bool all) {
        if (searching) return;
        searching = true;
        if (all) foreach (Prog p in progs) { p.Dir = ""; p.Auto = false; }
        status.Text = "Ищу программы: рядом с меню, на рабочем столе, в «Документах», на дисках...";
        status.ForeColor = Pal.Hint;
        Show2();
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            try { Finder.Find(here, progs); } catch { }
            try {
                BeginInvoke((MethodInvoker)delegate {
                    searching = false;
                    int n = 0;
                    foreach (Prog p in progs) if (p.Dir.Length > 0) { settings.Values[p.Key] = p.Dir; n++; }
                    settings.Save();
                    status.Text = n == progs.Length ? "Все программы найдены." : "Найдено программ: " + n + " из " + progs.Length + " - для остальных укажите папку («папка…»).";
                    status.ForeColor = n == progs.Length ? Pal.Good : Pal.Bad;
                    Show2();
                });
            } catch { }
        });
        th.IsBackground = true;
        th.Start();
    }

    void PickFolder(int i) {
        Prog p = progs[i];
        using (FolderBrowserDialog d = new FolderBrowserDialog()) {
            d.Description = "Папка программы «" + p.Title + "» (где " + p.Script + " и zapusk.vbs)";
            d.ShowNewFolderButton = false;
            try { if (Directory.Exists(p.Dir)) d.SelectedPath = p.Dir; } catch { }
            if (d.ShowDialog(this) != DialogResult.OK) return;
            SetFolder(i, d.SelectedPath);
        }
    }

    public bool SetFolder(int i, string dir) {
        Prog p = progs[i];
        if (!p.Fits(dir)) {
            status.Text = "В папке " + dir + " нет " + p.Script + " - это не «" + p.Title + "».";
            status.ForeColor = Pal.Bad;
            return false;
        }
        p.Dir = dir;
        p.Auto = false;
        settings.Values[p.Key] = dir;
        settings.Save();
        status.Text = "«" + p.Title + "» - папка запомнена.";
        status.ForeColor = Pal.Good;
        Show2();
        return true;
    }

    // запуск: zapusk.vbs программы (без чёрного окна), нет его - сам скрипт через PowerShell
    public void Launch(int i) {
        if (i < 0 || i >= progs.Length) return;
        Prog p = progs[i];
        if (p.Dir.Length == 0 || !p.Fits(p.Dir)) {
            status.Text = searching ? "Ещё ищу, где «" + p.Title + "», - секунду..." : "«" + p.Title + "» не найдена - укажите папку («папка…»).";
            status.ForeColor = Pal.Bad;
            if (!searching) { p.Dir = ""; Show2(); }
            return;
        }
        string vbs = Path.Combine(p.Dir, "zapusk.vbs"), file, args;
        if (File.Exists(vbs)) { file = "wscript.exe"; args = "\"" + vbs + "\""; }
        else { file = "powershell.exe"; args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + Path.Combine(p.Dir, p.Script) + "\""; }
        try {
            if (Start != null) Start(file, args);
            else {
                ProcessStartInfo si = new ProcessStartInfo(file, args);
                si.WorkingDirectory = p.Dir;
                si.UseShellExecute = false;
                Process.Start(si);
            }
        } catch (Exception ex) {
            status.Text = "«" + p.Title + "» не запустилась: " + ex.Message;
            status.ForeColor = Pal.Bad;
            return;
        }
        status.Text = "Запускаю «" + p.Title + "»...";
        status.ForeColor = Pal.Good;
        if (closeAfter.Checked) {
            Timer t = new Timer();
            t.Interval = 700;
            t.Tick += delegate { t.Stop(); Close(); };
            t.Start();
        }
    }

    public static string ShortcutPath {
        get {
            string desk = "";
            try { desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); } catch { }
            return Path.Combine(desk, "Программы.lnk");
        }
    }

    // ярлык на рабочем столе: wscript.exe "<папка меню>\zapusk.vbs", значок - menu.ico (рисуется сам)
    public void MakeShortcut() {
        try {
            string ico = Path.Combine(here, "menu.ico");
            try { MenuIcon.Save(ico); } catch { ico = ""; }
            Type t = Type.GetTypeFromProgID("WScript.Shell");
            if (t == null) throw new Exception("в Windows нет WScript.Shell");
            object sh = Activator.CreateInstance(t);
            object lnk = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, sh, new object[] { ShortcutPath });
            Type lt = lnk.GetType();
            string wscript = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "wscript.exe");
            lt.InvokeMember("TargetPath", BindingFlags.SetProperty, null, lnk, new object[] { wscript });
            lt.InvokeMember("Arguments", BindingFlags.SetProperty, null, lnk, new object[] { "\"" + Path.Combine(here, "zapusk.vbs") + "\"" });
            lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, lnk, new object[] { here });
            lt.InvokeMember("Description", BindingFlags.SetProperty, null, lnk, new object[] { "Поиск абонента, Оформление договора, Фото с телефона" });
            if (ico.Length > 0) lt.InvokeMember("IconLocation", BindingFlags.SetProperty, null, lnk, new object[] { ico + ",0" });
            lt.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
            status.Text = "Ярлык «Программы» - на рабочем столе.";
            status.ForeColor = Pal.Good;
        } catch (Exception ex) {
            Exception e = ex.InnerException ?? ex;
            status.Text = "Ярлык не создался: " + e.Message;
            status.ForeColor = Pal.Bad;
        }
        ShortcutText();
    }

    // для проверки
    public string Dump() {
        StringBuilder sb = new StringBuilder();
        foreach (ProgTile t in tiles) sb.Append("  ").Append(t.Index + 1).Append(". ").Append(t.P.Title).Append(": ").Append(t.State).Append("\n");
        sb.Append("  внизу: ").Append(status.Text).Append("\n");
        return sb.ToString();
    }
    public bool Busy { get { return searching; } }
}

// значок меню: синий квадрат со скруглёнными углами и тремя полосками разных цветов (три программы)
public static class MenuIcon {
    static Bitmap Draw(int size) {
        Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp)) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            float r = size * 0.22f, m = size * 0.04f, w = size - 2 * m;
            using (GraphicsPath p = new GraphicsPath()) {
                p.AddArc(m, m, r * 2, r * 2, 180, 90);
                p.AddArc(m + w - r * 2, m, r * 2, r * 2, 270, 90);
                p.AddArc(m + w - r * 2, m + w - r * 2, r * 2, r * 2, 0, 90);
                p.AddArc(m, m + w - r * 2, r * 2, r * 2, 90, 90);
                p.CloseFigure();
                using (SolidBrush b = new SolidBrush(Color.FromArgb(44, 98, 165))) g.FillPath(b, p);
            }
            Color[] bars = { Color.FromArgb(150, 200, 255), Color.FromArgb(140, 225, 165), Color.FromArgb(250, 195, 110) };
            float h = size * 0.13f, x = size * 0.22f, bw = size * 0.56f;
            for (int i = 0; i < 3; i++)
                using (SolidBrush b = new SolidBrush(bars[i])) g.FillRectangle(b, x, size * 0.24f + i * size * 0.2f, bw, h);
        }
        return bmp;
    }

    public static Icon Make() {
        using (MemoryStream ms = new MemoryStream(Bytes())) return new Icon(ms);
    }

    public static void Save(string file) { File.WriteAllBytes(file, Bytes()); }

    // файл .ico: значки 48, 32 и 16 точек, 32 бита (BMP внутри)
    public static byte[] Bytes() {
        int[] sizes = { 48, 32, 16 };
        List<byte[]> images = new List<byte[]>();
        foreach (int s in sizes) images.Add(Image(Draw(s)));
        MemoryStream ms = new MemoryStream();
        BinaryWriter w = new BinaryWriter(ms);
        w.Write((short)0);
        w.Write((short)1);
        w.Write((short)sizes.Length);
        int offset = 6 + 16 * sizes.Length;
        for (int i = 0; i < sizes.Length; i++) {
            w.Write((byte)sizes[i]);
            w.Write((byte)sizes[i]);
            w.Write((byte)0);
            w.Write((byte)0);
            w.Write((short)1);
            w.Write((short)32);
            w.Write(images[i].Length);
            w.Write(offset);
            offset += images[i].Length;
        }
        foreach (byte[] img in images) w.Write(img);
        w.Flush();
        return ms.ToArray();
    }

    static byte[] Image(Bitmap bmp) {
        int s = bmp.Width;
        MemoryStream ms = new MemoryStream();
        BinaryWriter w = new BinaryWriter(ms);
        w.Write(40);
        w.Write(s);
        w.Write(s * 2);           // высота - вместе с маской
        w.Write((short)1);
        w.Write((short)32);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        w.Write(0);
        for (int y = s - 1; y >= 0; y--)
            for (int x = 0; x < s; x++) {
                Color c = bmp.GetPixel(x, y);
                w.Write(c.B);
                w.Write(c.G);
                w.Write(c.R);
                w.Write(c.A);
            }
        int row = ((s + 31) / 32) * 4;      // маска - всё видно (прозрачность - в альфе)
        w.Write(new byte[row * s]);
        bmp.Dispose();
        w.Flush();
        return ms.ToArray();
    }
}

// Общее оформление программ - как у «Оформления договора»: тёмная тема по умолчанию (глаза меньше устают), светлая - по выбору
// (переключатель в окне, меняется при следующем запуске). Цвета ставятся один раз при запуске (Look.Apply), до создания окон;
// Look.Dress(окно) перекрашивает окно и всё, что в нём (и что добавится потом): системное светлое - в цвета темы, кнопки -
// плоские с рамкой. Цвета, заданные программой нарочно (например, синие окна как в программе базы), не трогаются
public static class Look {
    public static bool Dark = true;
    public static System.Drawing.Color Back, Surface, Input, Header, Border, InputBorder, Text, Caption, Hint, Accent, AccentHover, Hover, Pressed,
        Link, Good, Bad, Warn, Info, Strong, RowGood, RowWarn, RowBad, RowOther, SelGood, SelBad, SelOther, RowPick;
    public static readonly System.Drawing.Font Base = new System.Drawing.Font("Segoe UI", 9f), Small = new System.Drawing.Font("Segoe UI", 8.5f),
        Bold = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold), Title = new System.Drawing.Font("Segoe UI", 15f, System.Drawing.FontStyle.Bold);

    static System.Drawing.Color C(int r, int g, int b) { return System.Drawing.Color.FromArgb(r, g, b); }

    static Look() { Apply(true); }

    public static void Apply(bool dark) {
        Dark = dark;
        if (dark) {
            Back = C(30, 33, 38); Surface = C(40, 44, 51); Input = C(48, 53, 61); Header = C(44, 48, 55);
            Border = C(66, 72, 82); InputBorder = C(98, 106, 118);
            Text = C(222, 226, 232); Caption = C(172, 180, 192); Hint = C(138, 146, 158);
            Accent = C(44, 98, 165); AccentHover = C(54, 114, 188); Hover = C(56, 63, 74); Pressed = C(66, 76, 92);
            Link = C(118, 172, 236); Good = C(112, 200, 142); Bad = C(240, 120, 120); Warn = C(232, 164, 72); Info = C(130, 182, 240);
            Strong = C(205, 211, 220);
            RowGood = C(38, 72, 50); RowWarn = C(84, 64, 34); RowBad = C(88, 42, 46); RowOther = C(52, 56, 62);
            SelGood = C(46, 120, 70); SelBad = C(150, 50, 56); SelOther = C(84, 90, 100); RowPick = C(38, 60, 92);
        } else {
            Back = C(242, 244, 247); Surface = System.Drawing.Color.White; Input = System.Drawing.Color.White; Header = C(232, 236, 241);
            Border = C(221, 226, 233); InputBorder = C(196, 204, 214);
            Text = C(31, 36, 43); Caption = C(58, 66, 78); Hint = C(128, 137, 150);
            Accent = C(35, 99, 178); AccentHover = C(28, 84, 153); Hover = C(236, 242, 250); Pressed = C(222, 233, 247);
            Link = C(35, 99, 178); Good = C(46, 140, 87); Bad = C(200, 55, 55); Warn = C(190, 105, 0); Info = C(0, 70, 140);
            Strong = System.Drawing.Color.Black;
            RowGood = C(214, 241, 219); RowWarn = C(255, 228, 190); RowBad = C(255, 218, 218); RowOther = C(236, 236, 236);
            SelGood = C(40, 140, 60); SelBad = C(190, 40, 40); SelOther = C(110, 110, 110); RowPick = C(226, 240, 255);
        }
    }

    public static string SwitchText { get { return Dark ? "Тема: тёмная - сделать светлой" : "Тема: светлая - сделать тёмной"; } }
    public static string SwitchedText { get { return Dark ? "будет светлая - после перезапуска" : "будет тёмная - после перезапуска"; } }

    public delegate void SaveTheme(bool dark);

    // ссылка «Тема: ...» - щелчок запоминает другую тему (save), она будет при следующем запуске
    public static System.Windows.Forms.LinkLabel Switch(SaveTheme save) {
        System.Windows.Forms.LinkLabel l = new System.Windows.Forms.LinkLabel();
        l.Text = SwitchText;
        l.Font = Small;
        l.AutoSize = true;
        l.LinkColor = l.ActiveLinkColor = Link;
        l.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
        bool next = !Dark;
        l.LinkClicked += delegate {
            save(next);
            l.Text = next == Dark ? SwitchText : SwitchedText;
            next = !next;
        };
        return l;
    }

    static bool Same(System.Drawing.Color a, System.Drawing.Color b) { return a.ToArgb() == b.ToArgb(); }

    // светлое системное (его и красим)
    static bool Plain(System.Drawing.Color c) {
        return c.A == 255 && (Same(c, System.Drawing.Color.White) || Same(c, System.Drawing.SystemColors.Window) || Same(c, System.Drawing.SystemColors.Control)
            || Same(c, System.Drawing.SystemColors.ControlLight) || Same(c, System.Drawing.SystemColors.ButtonFace) || Same(c, System.Drawing.SystemColors.AppWorkspace));
    }

    static bool PlainText(System.Drawing.Color c) {
        return Same(c, System.Drawing.Color.Black) || Same(c, System.Drawing.SystemColors.ControlText) || Same(c, System.Drawing.SystemColors.WindowText);
    }

    // главная кнопка окна (синяя): AcceptButton окна или Tag «главная»
    public static void MainButton(System.Windows.Forms.Button b) { b.Tag = "главная"; Button(b); }

    public static void Button(System.Windows.Forms.Button b) {
        System.Windows.Forms.Form f = b.FindForm();
        bool main = "главная".Equals(b.Tag) || f != null && f.AcceptButton == b;
        bool custom = !main && !Plain(b.BackColor) && b.BackColor.A == 255;     // цвет задан программой - оставить
        b.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
        b.UseVisualStyleBackColor = false;
        b.Cursor = System.Windows.Forms.Cursors.Hand;
        if (main) {
            b.BackColor = Accent;
            b.ForeColor = System.Drawing.Color.White;
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = b.FlatAppearance.MouseDownBackColor = AccentHover;
        } else if (!custom) {
            b.BackColor = Surface;
            b.ForeColor = Text;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = InputBorder;
            b.FlatAppearance.MouseOverBackColor = Hover;
            b.FlatAppearance.MouseDownBackColor = Pressed;
        }
    }

    public static void Grid(System.Windows.Forms.DataGridView g) {
        g.BackgroundColor = Surface;
        g.GridColor = Border;
        g.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        g.DefaultCellStyle.BackColor = Input;
        g.DefaultCellStyle.ForeColor = Text;
        g.DefaultCellStyle.SelectionBackColor = Accent;
        g.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
        g.EnableHeadersVisualStyles = false;
        g.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
        g.ColumnHeadersDefaultCellStyle.BackColor = Header;
        g.ColumnHeadersDefaultCellStyle.ForeColor = Text;
        g.ColumnHeadersDefaultCellStyle.SelectionBackColor = Header;
        g.ColumnHeadersDefaultCellStyle.SelectionForeColor = Text;
        g.RowHeadersDefaultCellStyle.BackColor = Header;
        g.RowHeadersDefaultCellStyle.ForeColor = Text;
    }

    // окно и всё в нём (и что добавится потом)
    public static void Dress(System.Windows.Forms.Control c) {
        if (c == null || "свой цвет".Equals(c.Tag)) return;
        if (c is System.Windows.Forms.Form) {
            c.BackColor = Back;
            c.ForeColor = Text;
        } else if (c is System.Windows.Forms.TextBoxBase || c is System.Windows.Forms.ComboBox || c is System.Windows.Forms.ListBox
            || c is System.Windows.Forms.ListView || c is System.Windows.Forms.NumericUpDown || c is System.Windows.Forms.TreeView) {
            if (Plain(c.BackColor)) c.BackColor = c is System.Windows.Forms.TextBoxBase && ((System.Windows.Forms.TextBoxBase)c).ReadOnly ? Surface : Input;
            if (PlainText(c.ForeColor)) c.ForeColor = Text;
            System.Windows.Forms.TextBox t = c as System.Windows.Forms.TextBox;
            if (t != null && t.BorderStyle == System.Windows.Forms.BorderStyle.Fixed3D) t.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            System.Windows.Forms.ComboBox cb = c as System.Windows.Forms.ComboBox;
            if (cb != null && Dark) cb.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            System.Windows.Forms.ListView lv = c as System.Windows.Forms.ListView;
            if (lv != null && Dark) lv.GridLines = false;           // линии сетки Windows рисует светлыми - на тёмном режут глаз
        } else if (c is System.Windows.Forms.DataGridView) {
            Grid((System.Windows.Forms.DataGridView)c);
        } else if (c is System.Windows.Forms.Button) {
            Button((System.Windows.Forms.Button)c);
        } else if (c is System.Windows.Forms.LinkLabel) {
            System.Windows.Forms.LinkLabel l = (System.Windows.Forms.LinkLabel)c;
            l.LinkColor = Link;
            l.ActiveLinkColor = AccentHover;
            l.VisitedLinkColor = Link;
            if (PlainText(l.ForeColor)) l.ForeColor = Text;
        } else if (c is System.Windows.Forms.DateTimePicker) {
            System.Windows.Forms.DateTimePicker d = (System.Windows.Forms.DateTimePicker)c;
            d.CalendarMonthBackground = Input;
            d.CalendarForeColor = Text;
            d.CalendarTitleBackColor = Accent;
            d.CalendarTitleForeColor = System.Drawing.Color.White;
            d.CalendarTrailingForeColor = Hint;
        } else if (c is System.Windows.Forms.MonthCalendar) {
            System.Windows.Forms.MonthCalendar m = (System.Windows.Forms.MonthCalendar)c;
            // календарь Windows в стиле системы свои цвета не берёт - в тёмной теме рисуем его без стиля системы
            if (Dark) { if (m.IsHandleCreated) Plain(m); else m.HandleCreated += delegate { Plain(m); }; }
            m.BackColor = Input;
            m.ForeColor = Text;
            m.TitleBackColor = Accent;
            m.TitleForeColor = System.Drawing.Color.White;
            m.TrailingForeColor = Hint;
        } else if (c is System.Windows.Forms.ProgressBar) {
        } else {
            // подписи, галочки, панели, разделители
            if (!(c is System.Windows.Forms.Label) && !(c is System.Windows.Forms.CheckBox) && !(c is System.Windows.Forms.RadioButton) && Plain(c.BackColor)) c.BackColor = Back;
            if (PlainText(c.ForeColor)) c.ForeColor = Text;
            else if (Same(c.ForeColor, System.Drawing.SystemColors.GrayText)) c.ForeColor = Hint;
        }
        foreach (System.Windows.Forms.Control x in c.Controls) Dress(x);
        if (!hooked.ContainsKey(c) && (c is System.Windows.Forms.ContainerControl || c is System.Windows.Forms.Panel || c is System.Windows.Forms.GroupBox
            || c is System.Windows.Forms.SplitterPanel || c is System.Windows.Forms.TabPage)) {
            hooked[c] = true;
            System.Windows.Forms.Control me = c;
            c.ControlAdded += delegate(object s, System.Windows.Forms.ControlEventArgs e) { Dress(e.Control); };
            c.Disposed += delegate { hooked.Remove(me); };
        }
    }

    [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    static extern int SetWindowTheme(System.IntPtr hwnd, string app, string idList);

    static void Plain(System.Windows.Forms.Control c) {
        if (System.Environment.OSVersion.Platform != System.PlatformID.Win32NT) return;
        try { SetWindowTheme(c.Handle, "", ""); } catch { }
    }

    static readonly System.Collections.Generic.Dictionary<System.Windows.Forms.Control, bool> hooked = new System.Collections.Generic.Dictionary<System.Windows.Forms.Control, bool>();

    // окно - в теме при показе (и шрифт программы)
    public static void Wear(System.Windows.Forms.Form f) {
        f.Load += delegate { Dress(f); };
    }
}
'@
# Программа собирается из этого скрипта ОДИН раз после обновления и сохраняется готовой (файл menu_*.dll
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
$name = "menu_$stamp.dll"
$cacheDirs = @()
if ($env:LOCALAPPDATA) { $cacheDirs += (Join-Path $env:LOCALAPPDATA 'menu-programm') }
if ($env:PUBLIC) { $cacheDirs += (Join-Path $env:PUBLIC 'menu-programm') }
$cacheDirs += (Join-Path $here 'kesh-programmy')

function Test-Ready { return ($null -ne ('MenuApp' -as [type])) }

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
        $splash.Text = 'Программы'
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
            Get-ChildItem -Path $d -Filter 'menu_*.dll' -ErrorAction SilentlyContinue |
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

# ошибка запуска - в oshibki.txt рядом с программой (что за Windows, PowerShell, .NET, памяти) - его можно прислать
function Write-Crash($what, $text) {
    try {
        $mem = ''
        try { $os = Get-WmiObject Win32_OperatingSystem; $mem = ', памяти свободно ' + [int]($os.FreePhysicalMemory / 1024) + ' из ' + [int]($os.TotalVisibleMemorySize / 1024) + ' МБ' } catch { }
        $info = 'Windows ' + [Environment]::OSVersion.Version + ', PowerShell ' + $PSVersionTable.PSVersion + ', .NET ' + [Environment]::Version + ', ' + ([IntPtr]::Size * 8) + '-битный' + $mem
        $text = [regex]::Replace("$text", '\d{6,}', '######')
        $line = '=== ' + (Get-Date).ToString('dd.MM.yyyy HH:mm:ss') + ' (' + $what + ") ===`r`n" + $info + "`r`n" + $text + "`r`n`r`n"
        [IO.File]::AppendAllText((Join-Path $here 'oshibki.txt'), $line, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}

if (-not (Test-Ready)) {
    Write-Crash 'сборка' ($errors -join "`r`n`r`n")
    [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n" + ($errors -join "`n`n") + "`n`nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", 'Программы - ошибка')
    exit
}

try {
    [MenuApp]::Run($here, $launch)
} catch {
    $ex = $_.Exception
    while ($ex.InnerException -ne $null) { $ex = $ex.InnerException }
    Write-Crash 'работа' $ex.ToString()
    [void][System.Windows.Forms.MessageBox]::Show("Ошибка: " + $ex.Message + "`n`nПодробности - в файле oshibki.txt рядом с меню. Пришлите его (или фото этого окна).", 'Программы - ошибка')
}
