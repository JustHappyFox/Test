// Проверка 26.11: «Перевод из юр в физ» - улица «Михаила Миля»: в бытовой базе полностью, в юр базе «М.Миля», «М Миля», полностью
// (gen_byt.py, gen_ftu.py). Пишем улицу полностью, «М миля», «М. Миля» - находятся все три карточки юр базы и дом 5 в бытовой.
// Аргументы: папка программы; T_BYT - xlsx быта, T_FTU - dbf юр базы.
using System; using System.Windows.Forms; using System.Reflection;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], byt = Environment.GetEnvironmentVariable("T_BYT"), ftu = Environment.GetEnvironmentVariable("T_FTU");
        Istochnik.Save(true, byt, ftu, "");
        Console.WriteLine("   индексы быта: " + VygruzkaByt.Build(VygKind.Byt, byt, null).Replace("\n", " | "));
        foreach (string[] x in new string[][] { new string[] { "Михаила Миля", "миля" }, new string[] { "М. Миля", "миля" }, new string[] { "М.Миля", "миля" }, new string[] { "Миля М.", "миля" } })
            Eq("главное слово «" + x[0] + "»", JurBase.StreetWord(x[0]), x[1]);
        foreach (string[] x in new string[][] { new string[] { "М. Миля", "Миля" }, new string[] { "М.Миля", "Миля" }, new string[] { "М миля", "миля" },
                new string[] { "Михаила Миля", "Михаила Миля" }, new string[] { "1-я Набережная", "1-я Набережная" }, new string[] { "Я. Купалы", "Купалы" } })
            Eq("для быта «" + x[0] + "»", JurBase.NoInitials(x[0]), x[1]);
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        string[] ways = { "Михаила Миля", "М миля", "М. Миля" };
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step >= 3 && step < 3 + ways.Length * 2) {
                int k = (step - 3) / 2;
                if ((step - 3) % 2 == 0) { p.NoQuestions = true; p.SetAddress("п.", "Макетино", "ул.", ways[k], "5"); p.SetQuery("", "3"); p.Search(); }
                else {
                    string s = p.StatusText;
                    Console.WriteLine("«" + ways[k] + "»: " + s.Substring(0, Math.Min(s.Length, 400)));
                    ListView cards = null;
                    foreach (string n in new string[] { "abList", "cards", "abons" }) { FieldInfo fi = p.GetType().GetField(n, BF); if (fi != null && fi.GetValue(p) is ListView) cards = (ListView)fi.GetValue(p); }
                    if (cards != null) Eq("«" + ways[k] + "»: карточек юр базы", cards.Items.Count.ToString(), "3");
                    Eq("«" + ways[k] + "»: в бытовой базе дом найден", s.Contains("в доме 5:") ? "да" : "нет", "да");
                }
            } else if (step == 4 + ways.Length * 2) { Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad); t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
