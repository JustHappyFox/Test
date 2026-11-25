// Проверка 20.11: форма БС - вымышленные данные, 2 ПУ, линия, пломбы, схема; формирование пакета из 5 документов
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], shema = Environment.GetEnvironmentVariable("T_SHEMA");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            BsForm b = (BsForm)F(m, "bs");
            if (step == 1) m.OpenBs(m);
            else if (step == 2) {
                string[][] v = {
                    new string[] { "dog_nom", "12345" }, new string[] { "dog_data", "01.02.2015" }, new string[] { "abonent", "ООО «Тестовая связь»" },
                    new string[] { "tel", "+375 29 000-00-00" }, new string[] { "obekt", "Возведение мобильной базовой станции БС-0001" },
                    new string[] { "selsovet", "Пробный" }, new string[] { "np", "Тестовка" }, new string[] { "ulica", "Лесная" }, new string[] { "dom", "в районе дома №5" },
                    new string[] { "pitanie", "КТПБ-0001" }, new string[] { "data_razgr", "21.09.2026" }, new string[] { "mesto", "РП-100" },
                    new string[] { "liniya", "КЛ-0,4 кВ: РП-100 - ЩУНИ БС-0001 АВБбШв-1 5х16 L=50м." },
                    new string[] { "pu", "ЩУЭ|МЭС-3-5/100-1/3/А|1234567890|3х230/400|5(100)|000000,00|1|2026|-|-|1\nЩУНИ|МЭС-3-5/100-1/3/А|1234567891|3х230/400|5(100)|000010,50|2|2025|-|-|1" },
                    new string[] { "moshn", "ЩУЭ|20|-|24 ч/сут.|Без выходных\nЩУНИ|5|-|24 ч/сут.|Без выходных" },
                    new string[] { "plomba_klemm", "1111111" }, new string[] { "plomba_vvod", "2222222, 2222223" }, new string[] { "va_tok", "63" },
                    new string[] { "minus_on", "1" }, new string[] { "minus_x", "33333333" }, new string[] { "minus_y", "44444444" },
                    new string[] { "shema", shema }, new string[] { "nach_dolzh", Environment.GetEnvironmentVariable("T_DOLZH") ?? "Начальник ГРП" } };
                foreach (string[] x in v) b.SetV(x[0], x[1]);
                b.SetV("linii", "0,4|АВБбШв 4х120|0,18||\n10|АПвБбШп 3х95|0,35||");
                // как будто марку выбрали руками: r0 и Rл - сами
                for (int i = 0; i < 2; i++) { b.GetType().GetMethod("R0For", BF).Invoke(b, new object[] { i }); b.GetType().GetMethod("RFor", BF).Invoke(b, new object[] { i }); }
                b.GetType().GetMethod("KabelNote", BF).Invoke(b, new object[0]);
                Console.WriteLine("адрес: " + b.V("adres"));
                Console.WriteLine("коротко: " + b.V("adres_kr") + " | объект: " + b.V("obekt_full"));
                Console.WriteLine("пломбы: " + b.V("plomby") + " | минус: " + b.V("minus") + " | линии: " + b.V("linii"));
                Console.WriteLine("пакет: " + b.V("paket") + " | инженер полностью: " + b.V("inzh_fio") + " | состояние: " + b.StatusText);
            } else if (step == 3) {
                Panel content = (Panel)F(b, "content");
                System.Collections.IList cards = (System.Collections.IList)F(b, "cards");
                Control p = (Control)F(cards[3], "Panel");
                content.AutoScrollPosition = new System.Drawing.Point(0, p.Top - content.AutoScrollPosition.Y - 4);
            } else if (step == 4) {
                string dir = b.Save(false, false);
                Console.WriteLine("Пакет: " + dir);
                if (dir != null) foreach (string f in Directory.GetFiles(dir)) Console.WriteLine("  " + Path.GetFileName(f) + " " + new FileInfo(f).Length);
                int seals = SealRegistry.All(here, Memory.For(Path.Combine(here, "nastroiki.txt"))).Count;
                Console.WriteLine("Пломб в отчёт: " + seals);
            } else if (step == 6) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
