// Проверка 26.11: населённый пункт с приписанным сельсоветом («Лесной Боровлянский с/с» в поле н.п. выгрузки быта, gen_byt.py) -
// адрес «п. Лесной, ул. Садовая, д. 9» находится. Аргументы: папка «Договора»; T_BYT - xlsx быта.
using System; using System.Windows.Forms; using System.Reflection;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    [STAThread] public static void Main(string[] a) {
        string byt = Environment.GetEnvironmentVariable("T_BYT");
        Istochnik.Save(true, byt, "", "");
        Console.WriteLine("   " + VygruzkaByt.Build(VygKind.Byt, byt, null).Replace("\n", " | "));
        foreach (string[] x in new string[][] { new string[] { "Лесной Боровлянский с/с", "", "Лесной|Боровлянский с/с" },
                new string[] { "Лесной, Боровлянский сельсовет", "", "Лесной|Боровлянский с/с" }, new string[] { "Лесной Боровлянский с/с", "Боровлянский", "Лесной|Боровлянский" },
                new string[] { "Новая Выдумка", "", "Новая Выдумка|" } }) {
            string town = x[0], ss = x[1];
            StreetDirectory.SplitTown(ref town, ref ss);
            Eq("«" + x[0] + "»", town + "|" + ss, x[2]);
        }
        StartMenu m = new StartMenu(a[0], "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) { m.OpenJur(); return; }
            if (step < 3) return;
            JurPick p = (JurPick)m.GetType().GetField("jurPick", BF).GetValue(m);
            BaseLink b = (BaseLink)p.GetType().GetMethod("Base", BF).Invoke(p, null);
            foreach (string q in new string[] { "п. Лесной, ул. Садовая, д. 9", "Лесной Садовая 9", "Лесной Боровлянский с/с Садовая 9", "Садовая 9 Боровлянский с/с" }) {
                string note;
                int n = -1;
                try { n = b.Address(q, out note).Count; Console.WriteLine("   " + q + " -> " + n + (note != null ? " (" + note + ")" : "")); } catch (Exception ex) { Console.WriteLine(q + " ! " + ex.Message); }
                Eq("«" + q + "» найдено", n.ToString(), "2");
            }
            Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
            t.Stop(); Environment.Exit(0);
        };
        t.Start(); Application.Run(m);
    }
}
