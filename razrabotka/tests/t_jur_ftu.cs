// Проверка 24.11: перевод из юр в физ по юр базе FTU (вымышленный DBF gen_ftu.py) + соседи из выгрузки быта (gen_byt.py)
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step == 3) {
                p.NoQuestions = true;
                p.SetAddress("д.", "Макетино", "ул.", "Полевая", "34");
                p.SetQuery("", "12");
                p.Search();
            } else if (step == 5) {
                Console.WriteLine(p.StatusText);
            } else if (step == 6) {
                p.CreateSpravka();
            } else if (step == 7) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                foreach (string k in new string[] { "pu1_ls", "pu1_np", "pu1_ulica", "pu1_dom", "pu1_kv", "pu1_zav", "pu1_tip", "pu1_pov" })
                    Console.WriteLine("  справка " + k + " = " + s.GetType().GetMethod("V", BF).Invoke(s, new object[] { k }));
            } else if (step == 9) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
