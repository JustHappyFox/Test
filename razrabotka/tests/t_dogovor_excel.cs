// Проверка 20.11: «Договор» в режиме Excel-выгрузки быта (вымышленная выгрузка gen_byt.py, лицевой 010019999)
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static object Call(object o, string n, params object[] a) {
        foreach (MethodInfo m in o.GetType().GetMethods(BF)) if (m.Name == n && m.GetParameters().Length == a.Length) return m.Invoke(o, a);
        throw new Exception("нет " + n);
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) m.OpenDogovor();
            else if (step == 3) {
                DogovorForm d = (DogovorForm)F(m, "dog");
                Console.WriteLine("Договор, надпись о данных: " + d.BaseText.Replace("\n", " / "));
                Call(d, "SetV", "ls", "5002010019999");
            } else if (step == 5) {
                DogovorForm d = (DogovorForm)F(m, "dog");
                Console.WriteLine("Договор, лицевой: " + ((Control)F(d, "lsNote")).Text);
                Call(m, "OpenSpravka", d, null);
            } else if (step == 6) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                Call(s, "SetV", "pu1_ls", "010019999");
            } else if (step == 8) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                foreach (string k in new string[] { "pu1_fio", "pu1_tel", "pu1_np", "pu1_ulica", "pu1_dom", "pu1_adres", "pu1_zav", "pu1_tip", "pu1_pov" })
                    Console.WriteLine("  справка " + k + " = " + Call(s, "V", k));
            } else if (step == 10) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
