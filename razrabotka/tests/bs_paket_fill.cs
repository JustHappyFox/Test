// ЗАГОТОВКА: пакет документов юр лица по бумагам пользователя (см. razrabotka/BS-PAKET.md). Данные здесь ВЫМЫШЛЕННЫЕ -
// копию с настоящими держать только в scratchpad (в git не класть). Запуск:
// T_SHEMA=схема.jpg runtest.sh dogovor bs_paket_fill.cs <папка-копия программы с шаблонами> <png-префикс> 40
using System; using System.Windows.Forms; using System.Reflection; using System.IO; using System.Diagnostics; using System.Collections;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1], shema = Environment.GetEnvironmentVariable("T_SHEMA");
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 1800; int step = 0; int shot = 0;
        t.Tick += delegate {
            step++;
            BsForm b = (BsForm)F(m, "bs");
            if (step == 1) m.OpenBs(m);
            else if (step == 2) {
                b.WindowState = FormWindowState.Maximized;
                string[][] v = {
                    new string[] { "dog_nom", "12345" }, new string[] { "dog_data", "01.02.2020" }, new string[] { "abonent", "ООО «Тестовая зарядка»" },
                    new string[] { "tel", "+375 29 000-00-00" }, new string[] { "obekt", "Установка зарядных станций для электромобилей в паркинге" },
                    new string[] { "kod", "" }, new string[] { "deyat", "электрическая энергия используется для зарядки электромобилей" },
                    new string[] { "obl", "Минская обл." }, new string[] { "rayon", "Минский р-н" }, new string[] { "selsovet", "Пробный" },
                    new string[] { "np_tip", "д." }, new string[] { "np", "Тестовка" }, new string[] { "ulica_tip", "ул." }, new string[] { "ulica", "Лесная" }, new string[] { "dom", "10" },
                    new string[] { "pitanie", "КТПБ-0001" }, new string[] { "data_razgr", "01.10.2026" }, new string[] { "mesto", "КТПБ-0001, РУ-0,4 кВ, 2 секция, руб. 1" },
                    new string[] { "liniya", "КЛ-0,4 кВ: КТПБ-0001 - ЩУР-1 АВБбШв 4х95 L=44м." },
                    new string[] { "pu", "ЭЗС|СС-302-05T4U1-AP1|11111111|3х230/400|5(10)|без напряжения|3|2026|ТОП-0,66-1-5 УЗ 150/5 кл.т. 0,5S № 2222221, 2222222, 2222223 (пов. 2026 2 кв.)|-|30" },
                    new string[] { "moshn", "ЭЗС|80|-|24 ч/сут.|Без выходных" },
                    new string[] { "plomba_drugoe", "ИК 3333332; ТТ 3333333, 3333334, 3333335; ФП 3333336" }, new string[] { "plomba_marka", "МЭСБ" }, new string[] { "plomba_vvod", "" }, new string[] { "plomba_klemm", "3333331" },
                    new string[] { "tarif", "Электрическая энергия, используемая электрозарядными станциями, предназначенными для зарядки электромобилей (одноставочный)" },
                    new string[] { "tret", "-" }, new string[] { "prim", "-" },
                    new string[] { "prichina", "Допуск в эксплуатацию" }, new string[] { "inzh", "Тестов Т.Т." }, new string[] { "inzh_fio", "Тестов Тест Тестович" }, new string[] { "nach_dolzh", "Начальник ОПСУ" }, new string[] { "nach", "Примеров П.П." }, new string[] { "data_pu", "01.11.2026" }, new string[] { "data_podp", "01.11.2026" }, new string[] { "data_akt", "01.11.2026" },
                    new string[] { "shema", shema } };
                foreach (string[] x in v) b.SetV(x[0], x[1]);
                b.SetV("linii", "0,4|АВБбШв-4х95|0,044||");
                b.GetType().GetMethod("R0For", BF).Invoke(b, new object[] { 0 });
                b.GetType().GetMethod("RFor", BF).Invoke(b, new object[] { 0 });
                foreach (string k in new string[] { "plomby", "pu", "adres", "adres_kr", "obekt_full", "obekt_adres", "obekt_kr", "potr", "linii", "paket", "inzh_fio" }) Console.WriteLine(k + " = " + b.V(k));
                Console.WriteLine("состояние: " + b.StatusText);
            } else if (step >= 3 && step < 3 + 12) {
                Panel content = (Panel)F(b, "content");
                IList cards = (IList)F(b, "cards");
                int i = step - 3;
                if (i >= cards.Count) return;
                Control p = (Control)F(cards[i], "Panel");
                content.AutoScrollPosition = new System.Drawing.Point(0, p.Top - content.AutoScrollPosition.Y - 4);
                Application.DoEvents();
                shot++;
                Process.Start("import", "-window root " + png + "_" + shot.ToString("00") + ".png").WaitForExit();
            } else if (step == 16) {
                string dir = b.Save(false, false);
                Console.WriteLine("Пакет: " + dir);
            } else if (step == 18) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
