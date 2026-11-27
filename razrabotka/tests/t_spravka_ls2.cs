// Проверка 27.11: справка из перевода из юр в физ (квартира с двумя счётчиками) - лицевой проверяется у КАЖДОГО ПУ (и у второго):
// не занят ли: нет в базе - «свободен», есть - «ЗАНЯТ» (справка из перевода - нужен свободный). Вымышленные базы: T_BYT, T_FTU.
using System; using System.Windows.Forms; using System.Reflection;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, bool ok, string got) { if (!ok) bad++; Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»"); }
    [STAThread] public static void Main(string[] a) {
        string byt = Environment.GetEnvironmentVariable("T_BYT");
        Istochnik.Save(true, byt, Environment.GetEnvironmentVariable("T_FTU"), "");
        VygruzkaByt.Build(VygKind.Byt, byt, null);
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(a[0], "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step == 3) { p.NoQuestions = true; p.SetAddress("п.", "Макетино", "ул.", "Полевая", "34"); p.SetQuery("", "12"); p.Search(); }
            else if (step == 5) p.CreateSpravka();
            else if (step == 6) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                MethodInfo sv = s.GetType().GetMethod("SetV", BF);
                Console.WriteLine("   ПУ в справке: " + ((System.Collections.IList)F(s, "meters")).Count);
                sv.Invoke(s, new object[] { "pu1_ls", "099000001" });   // новый лицевой
                sv.Invoke(s, new object[] { "pu2_ls", "000990000" });   // второй - уже есть в базе (Михаила Миля, 5)
            } else if (step == 8) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                string h = ((Label)F(s, "headSub")).Text;
                Console.WriteLine("   " + h);
                Eq("проверен лицевой ПУ 1", h.Contains("ПУ 1:"), h);
                Eq("проверен лицевой ПУ 2", h.Contains("ПУ 2:"), h);
                Eq("ПУ 1 - свободен", h.Contains("ПУ 1: лицевой 099000001 свободен"), h);
                Eq("ПУ 2 - занят", h.Contains("ПУ 2: ЛИЦЕВОЙ 000990000 ЗАНЯТ"), h);
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        t.Start(); Application.Run(m);
    }
}
