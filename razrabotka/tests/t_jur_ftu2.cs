// Проверка 24.11: юр база FTU как на фото - «кв.№12 - быт» + «кв№12 - элекроотопление» -> два счётчика; «38 - быт» -> один
using System; using System.Windows.Forms; using System.Reflection;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        string[] flats = { "12", "38", "8" };
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step >= 3 && step < 3 + flats.Length * 2) {
                int k = (step - 3) / 2;
                if ((step - 3) % 2 == 0) { p.NoQuestions = true; p.SetAddress("п.", "Макетино", "ул.", "Полевая", "34"); p.SetQuery("", flats[k]); p.Search(); }
                else {
                    string s = p.StatusText;
                    int i = s.IndexOf("| ТУ:"), j = s.IndexOf("| Лицевые");
                    Console.WriteLine("кв." + flats[k] + ": " + s.Substring(0, s.IndexOf('|')).Trim() + "\n    " + s.Substring(i, j - i).Trim());
                    int w = s.IndexOf("| В договор");
                    Console.WriteLine("    " + s.Substring(w).Trim());
                }
            } else if (step == 4 + flats.Length * 2) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
