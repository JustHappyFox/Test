// Проверка 24.11: потери по схеме (ТКП 460-2017) - вымышленные данные. Вариант 1: учёт 0,4 кВ через ТТ, ТМ-250/10 абонента,
// кабель 0,4 кВ и кабель 10 кВ до границы. Вариант 2: учёт 10 кВ через ТТ и ТН (Арона), кабель 10 кВ. Пакет - в папку программы.
using System; using System.Windows.Forms; using System.Reflection; using System.IO; using System.Diagnostics;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static void M(object o, string n, params object[] a) {
        foreach (MethodInfo mi in o.GetType().GetMethods(BF)) if (mi.Name == n && mi.GetParameters().Length == a.Length) { mi.Invoke(o, a); return; }
    }
    static void Lines(BsForm b) {
        RowList l = (RowList)F(b, "linii");
        for (int i = 0; i < l.Count; i++) { M(b, "R0For", i); M(b, "RFor", i); }
        M(b, "KabelNote"); M(b, "PotNote");
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1], shema = Environment.GetEnvironmentVariable("T_SHEMA");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            BsForm b = (BsForm)F(m, "bs");
            if (step == 1) { m.OpenBs(m); }
            else if (step == 2) {
                b.WindowState = FormWindowState.Maximized;
                foreach (string[] x in new string[][] {
                    new string[] { "dog_nom", "12345" }, new string[] { "abonent", "ООО «Тестовый завод»" }, new string[] { "obekt", "Цех" },
                    new string[] { "np", "Тестовка" }, new string[] { "ulica", "Лесная" }, new string[] { "dom", "5" }, new string[] { "pitanie", "ТП-0001" },
                    new string[] { "pu", "ЩУ|Тестовый-3Ф||3х230/400|5(10)||1|2026|ТОП-0,66 300/5|-||" }, new string[] { "moshn", "ЩУ|150|-|8 ч|2 выходных" },
                    new string[] { "shema", shema },
                    new string[] { "pot_u", "0,4 кВ" }, new string[] { "pot_uchet", "через ТТ" }, new string[] { "balans", "абонента" },
                    new string[] { "tr_tip", "ТМ-250/10" }, new string[] { "rezhim", "в две смены" }, new string[] { "pot_react", "1" }, new string[] { "pot_test", "60000" } })
                    b.SetV(x[0], x[1]);
                b.SetV("linii", "0,4|АВБбШв 4х120|0,05||\n10|ААБл 3х95|0,6||");
                Lines(b);
                M(b, "ZerosFor", 0); M(b, "KFor", 0);
                Console.WriteLine("1) " + ((Label)F(b, "potNote")).Text.Replace("\n", " | "));
                Console.WriteLine("   ПУ: " + b.V("pu") + " | линия: " + b.V("liniya"));
            } else if (step == 3) {
                Panel content = (Panel)F(b, "content");
                System.Collections.IList cards = (System.Collections.IList)F(b, "cards");
                Control p = (Control)F(cards[7], "Panel");
                content.AutoScrollPosition = new System.Drawing.Point(0, p.Top - content.AutoScrollPosition.Y - 4);
            } else if (step == 4) {
                Process.Start("import", "-window root " + png + "_1.png").WaitForExit();
                Console.WriteLine("   пакет: " + b.Save(false, false));
            } else if (step == 5) {
                b.SetV("dog_nom", "12346");
                b.SetV("pot_u", "10 кВ"); b.SetV("pot_uchet", "через ТТ и ТН"); b.SetV("pot_aron", "Арона"); b.SetV("tr_tip", "нет");
                b.SetV("pu", "ЯЧ-1|Тестовый-3Ф||3х100|5(10)||1|2026|ТОЛ-10 100/5|НАМИ-10 10000/100||");
                b.SetV("linii", "10|АПвБбШп 3х120|0,3||");
                Lines(b);
                M(b, "ZerosFor", 0); M(b, "KFor", 0);
                Console.WriteLine("2) " + ((Label)F(b, "potNote")).Text.Replace("\n", " | "));
                Console.WriteLine("   ПУ: " + b.V("pu") + " | линия: " + b.V("liniya"));
                Console.WriteLine("   пакет: " + b.Save(false, false));
            } else if (step == 7) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
