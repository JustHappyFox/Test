// Проверка 24.11: наряды из выгрузки (вымышленная, gen_nar.py) - окно «Наряды…» «Поиска»: список номеров, наряды лицевого
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        AbonentSearch f = new AbonentSearch(Path.Combine(here, "poisk.txt"), Path.Combine(here, "nastroiki.txt"));
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 2) f.GetType().GetMethod("OpenNaryady", BF).Invoke(f, null);
            else if (step == 3) {
                NarExcelForm n = (NarExcelForm)F(f, "LastNarExcel");
                n.SetInput("202610001\n2025-10000\n202699999\n10002");     // выполнен, ..., нет такого, без года
                n.Find();
                Console.WriteLine("Итог: " + n.InfoText + " | строк: " + n.FoundCount);
                Console.WriteLine("Лицевой первой строки: " + n.LsText + " | нарядов лицевого: " + n.LsCount);
            } else if (step == 4) {
                NarExcelForm c = new NarExcelForm("010019999", "Шаблонов М.А.");
                c.Show(f);
            } else if (step == 5) {
                foreach (Form x in Application.OpenForms) if (x is NarExcelForm && x != F(f, "LastNarExcel")) Console.WriteLine("Ctrl+Z: " + ((NarExcelForm)x).LsText);
            } else if (step == 7) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(f);
    }
}
