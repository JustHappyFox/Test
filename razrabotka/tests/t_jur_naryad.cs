// Проверка 20.11: юр наряд - из пакета БС (вымышленный, от t_bs.cs), сохранение; меню с плитками БС и юр наряда
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 2) m.OpenJurNaryad(m);
            else if (step == 3) {
                JurNaryadForm j = (JurNaryadForm)F(m, "jnar");
                string[] f = Directory.GetFiles(BsForm.DataFolder(here), "*.txt");
                j.FromBsData(DogovorForm.ReadData(f[0]));
                j.SetV("mesto", "РП-100 с.1 яч.1, с.2 яч.2");
                Console.WriteLine("Юр наряд: " + j.StatusText + " | адрес: " + j.V("adres") + " | объекты: " + j.V("obekty"));
            } else if (step == 5) {
                JurNaryadForm j = (JurNaryadForm)F(m, "jnar");
                Console.WriteLine("Сохранён: " + j.Save(false, false));
            } else if (step == 6) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
