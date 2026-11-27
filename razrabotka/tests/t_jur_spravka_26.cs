// Проверка 26.11: «Справка юр» и «Поиск по юр базе» в «Договоре» на вымышленной юр базе (tests/gen_ftu.py). Аргументы: папка программы,
// png-префикс; T_FTU - путь к ftu.dbf. Подтягивание: договор + карточка (все счётчики), № счётчика (один, с ТТ); причина -> приложения;
// сохранение справки; поиск: договор/карточка, договор, счётчик, слова.
using System; using System.Windows.Forms; using System.Reflection; using System.IO; using System.Collections;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    static void Snap(string file) { Application.DoEvents(); try { System.Diagnostics.Process.Start("import", "-window root " + file).WaitForExit(); } catch { } }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1];
        Istochnik.Save(true, "", Environment.GetEnvironmentVariable("T_FTU"), "");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) { Snap(png + "_menu.png"); m.OpenJurSpravka(m, null, null); return; }
            JurSpravkaForm s = (JurSpravkaForm)F(m, "jspr");
            RowList pu = (RowList)F(s, "pu");
            if (step == 2) {
                Console.WriteLine("--- договор + карточка");
                s.SetV("kp", "11118"); s.SetV("nab", "26"); s.ByCard();
                Eq("счётчиков", pu.Count.ToString(), "7");
                Eq("абонент", s.V("abonent"), "ООО «Тестстрой»");
                Eq("объект", s.V("obekt"), "«Нераспроданные квартиры»");
                Eq("адрес", s.V("adres"), "Минский р-н, п. Макетино, ул. Полевая, 34");
                Eq("1-я строка", pu.Value.Split('\n')[0], "кв.№8 - быт|СЕ 318 BY|80504000|3х230/400|5(60)|000000,00|3|2025|-|1");
                s.SetV("prichina", "Требуется расторгнуть договор по заявлению абонента");
                Application.DoEvents();
                Eq("приложения сами", ((RowList)F(s, "att")).Value, "Дополнительное соглашение на расторжение договора\nЗаявление абонента");
                Snap(png + "_spr1.png");
                string f = s.Save(false, false);
                Console.WriteLine("   сохранена: " + (f == null ? "НЕТ" : Path.GetFileName(f)));
            } else if (step == 3) {
                Console.WriteLine("--- № счётчика (с ТТ)");
                s.Clear();
                s.SetV("meter", "72220001"); s.ByMeter();
                Eq("договор/карточка", s.V("kp") + "/" + s.V("nab"), "22222/1");
                Eq("абонент", s.V("abonent"), "ЧУП «Пример»");
                Eq("строка", pu.Value, "ЩУ магазина|СЕ 318 BY|72220001|3х230/400|5(10)|000000,00|2|2023|150/5|30");
                Console.WriteLine("--- договор без карточки - выбор карточки");
                s.Clear(); s.SetV("kp", "11118"); s.ByCard();
                Console.WriteLine("   " + ((Label)F(s, "found")).Text);
                Console.WriteLine("--- нет такого счётчика");
                s.SetV("meter", "99999999"); s.ByMeter();
                Console.WriteLine("   " + ((Label)F(s, "found")).Text);
                s.Clear();
                s.GoBack();
                m.OpenJurSearch();
            } else if (step == 4) {
                JurBrowser j = (JurBrowser)F(m, "jfind");
                string[][] q = { new string[] { "11118", "", "0, 2" }, new string[] { "11118/26", "", "1, 7" }, new string[] { "72220001", "", "1, 1" },
                    new string[] { "", "Полевая 34", "0, 1" }, new string[] { "", "тестстрой офис", "0, 1" }, new string[] { "", "нетакого", "0, 0" } };
                foreach (string[] x in q) { j.Num = x[0]; j.Words = x[1]; j.Search(); Eq("поиск «" + x[0] + "» + «" + x[1] + "»: уровень, строк", j.Level + ", " + j.Rows, x[2]); }
                j.Num = ""; j.Words = "Полевая 34"; j.Search();
                Snap(png + "_find.png");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        t.Start();
        Application.Run(m);
    }
}
