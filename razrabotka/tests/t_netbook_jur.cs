// Проверка 24.11: «Перевод из юр в физ» на нетбуке (1024x560): режим «база как раньше», базы «Долгов района» нет, указана выгрузка
// юр лиц FTU .dbf - ищется по выгрузке; затем режим Excel. Аргументы: папка, png-префикс; T_BYT, T_JUR - пути выгрузок.
using System; using System.Windows.Forms; using System.Reflection; using System.Diagnostics;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1], byt = Environment.GetEnvironmentVariable("T_BYT"), jur = Environment.GetEnvironmentVariable("T_JUR");
        Istochnik.Save(false, byt, jur, "");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step == 2 || step == 5) {
                if (step == 5) Istochnik.Save(true, byt, jur, "");
                p.NoQuestions = true; p.SetAddress("п.", "Макетино", "ул.", "Полевая", "34"); p.SetQuery("", "12"); p.Search();
            } else if (step == 4 || step == 7) {
                Console.WriteLine((step == 4 ? "Режим «база»: " : "Режим Excel: ") + p.StatusText.Replace("\n", " "));
                Process.Start("import", "-window root " + png + "_" + step + ".png").WaitForExit();
            } else if (step == 8) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
