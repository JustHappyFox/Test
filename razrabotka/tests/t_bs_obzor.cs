// Обзор окна пакета документов: вымышленные данные, снимок каждого раздела (для проработки окна). Аргументы: папка, png-префикс.
using System; using System.Windows.Forms; using System.Reflection; using System.Diagnostics; using System.Collections;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            BsForm b = (BsForm)F(m, "bs");
            if (step == 1) { m.OpenBs(m); return; }
            if (step == 2) {
                b.WindowState = FormWindowState.Maximized;
                if (Environment.GetEnvironmentVariable("T_FILL") == "1") {
                    // вымышленный пример: вид объекта подставляет код, деятельность, тариф, точку учёта, режим
                    b.SetV("vid_obekta", "Зарядная станция электромобилей");
                    b.SetV("dog_nom", "12345"); b.SetV("abonent", "ООО «Тестовая зарядка»"); b.SetV("np", "Тестовка"); b.SetV("ulica", "Лесная"); b.SetV("dom", "5");
                    b.SetV("pitanie", "КТПБ-0001"); b.SetV("pot_uchet", "через ТТ");
                }
                Snap(png, 0); return;
            }
            IList cards = (IList)F(b, "cards");
            int i = step - 3;
            if (i >= cards.Count) { t.Stop(); Environment.Exit(0); }
            Panel content = (Panel)F(b, "content");
            Control p = (Control)F(cards[i], "Panel");
            content.AutoScrollPosition = new System.Drawing.Point(0, p.Top - content.AutoScrollPosition.Y - 4);
            Application.DoEvents();
            Snap(png, i + 1);
        };
        t.Start();
        Application.Run(m);
    }
    static void Snap(string png, int n) { Process.Start("import", "-window root " + png + "_" + n.ToString("00") + ".png").WaitForExit(); }
}
