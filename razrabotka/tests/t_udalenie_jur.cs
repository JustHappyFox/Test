// Проверка 27.11: справка в РГ из справки (быт) без данных перевода: лицевые двух ПУ - через запятую; договор, карточка, наименование,
// объект - из юр базы по номерам счётчиков (вымышленная юр база gen_ftu.py, T_FTU). Аргумент: папка программы.
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
        Istochnik.Save(true, "", Environment.GetEnvironmentVariable("T_FTU"), "");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(a[0], "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) { m.OpenSpravka(m, null); return; }
            SpravkaForm s = (SpravkaForm)F(m, "spr");
            MethodInfo sv = s.GetType().GetMethod("SetV", BF);
            if (step == 2) {
                s.AddMeter();
                string[][] v = { new string[] { "pu1_ls", "052950737" }, new string[] { "pu1_zav", "80504002" }, new string[] { "pu1_tip", "CE318BY" },
                    new string[] { "pu2_ls", "052950745" }, new string[] { "pu2_zav", "80504003" }, new string[] { "pu2_tip", "CE318BY" },
                    new string[] { "pu1_adres", "п. Макетино, ул. Полевая, д. 34, кв. 12" } };
                foreach (string[] x in v) sv.Invoke(s, new object[] { x[0], x[1] });
                m.OpenUdalenie(s, null, s.Data());
            } else if (step == 3) {
                UdalenieForm u = (UdalenieForm)F(m, "udal");
                Eq("лицевые", u.V("ls"), "052950737, 052950745");
                Eq("договор/карточка", u.V("kp") + "/" + u.V("nab"), "11118/26");
                Eq("наименование", u.V("dog_name"), "ООО «Тестстрой»");
                Eq("объект", u.V("obekt"), "Нераспроданные квартиры");
                Console.WriteLine("   " + ((Label)F(u, "headSub")).Text);
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        t.Start(); Application.Run(m);
    }
}
