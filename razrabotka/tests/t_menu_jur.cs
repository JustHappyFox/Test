// Проверка 26.11: меню программ - строка «Поиск быт | Поиск юр», окно «Поиск юр» по вымышленной юр базе (tests/gen_ftu.py, T_FTU).
// Аргументы: папка меню, png-префикс.
using System; using System.Windows.Forms; using System.Reflection;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    static void Snap(string file) { Application.DoEvents(); try { System.Diagnostics.Process.Start("import", "-window root " + file).WaitForExit(); } catch { } }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1];
        Istochnik.Save(true, "", Environment.GetEnvironmentVariable("T_FTU"), "");
        Application.EnableVisualStyles();
        MenuForm m = new MenuForm(here);
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) {
                Control jt = (Control)F(m, "jurTile");
                Eq("плитка «Поиск юр» справа от «Поиск быт»", jt.Left > 300 && jt.Top < 100 ? "да" : "нет", "да");
                Snap(png + "_menu.png");
                m.GetType().GetMethod("OpenJurFind", BF).Invoke(m, null);
            } else if (step == 2) {
                JurFindForm j = (JurFindForm)F(m, "jurFind");
                string[][] q = { new string[] { "11118/26", "", "7" }, new string[] { "11118", "", "8" }, new string[] { "72220001", "", "1" },
                    new string[] { "", "Полевая 34", "7" }, new string[] { "", "пример садовая", "1" }, new string[] { "11118", "кв 12", "2" } };
                foreach (string[] x in q) { j.Num = x[0]; j.Words = x[1]; j.Search(); Eq("поиск «" + x[0] + "» + «" + x[1] + "»", j.Found.ToString(), x[2]); }
                j.Num = ""; j.Words = "Полевая 34"; j.Search();
                Snap(png + "_find.png");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        t.Start();
        Application.Run(m);
    }
}
