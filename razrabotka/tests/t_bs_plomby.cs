// Проверка 20.11: пломбы пакета документов (БС) - в отчёте пломб партии с такими номерами (вымышленный пакет от t_bs.cs)
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
            if (step == 1) m.OpenPlomby();
            else if (step == 2) {
                PlombyForm p = (PlombyForm)F(m, "pl");
                p.AddBatch("пластиковые", "1111101", 100, "18.11.2026", "Тестов Т.Т.", "Инженер по РПЭ", "");
                p.AddBatch("пластиковые", "2222201", 100, "18.11.2026", "Тестов Т.Т.", "Инженер по РПЭ", "");
            } else if (step == 3) {
                PlombyForm p = (PlombyForm)F(m, "pl");
                Console.WriteLine(p.Status);
                foreach (string line in p.Dump().Split('\n')) if (line.Contains("1111111") || line.Contains("222222")) Console.WriteLine("  " + line.Trim());
                Console.WriteLine("Печать: " + p.Print(false));
            } else if (step == 5) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
