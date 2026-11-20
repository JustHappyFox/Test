// Проверка 20.11: «Поиск» - история 30 последних поисков (режим Excel, вымышленная выгрузка). Уточнение адреса заменяет запись.
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static void Call(object o, string n, params object[] a) { o.GetType().GetMethod(n, BF).Invoke(o, a); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        try { File.Delete(Path.Combine(here, "istoriya.txt")); } catch { }
        Application.EnableVisualStyles();
        AbonentSearch f = new AbonentSearch(Path.Combine(here, "poisk.txt"), Path.Combine(here, "nastroiki.txt"));
        string[][] q = { new string[] { "0", "010019999" }, new string[] { "1", "Макетино" }, new string[] { "1", "Макетино Полевая 34" }, new string[] { "3", "Шаблонов Мария" } };
        Timer t = new Timer(); t.Interval = 3000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step >= 2 && step < 2 + q.Length) {
                string[] x = q[step - 2];
                ((RadioButton[])F(f, "modeButtons"))[int.Parse(x[0])].Checked = true;
                ((TextBox)F(f, "queryBox")).Text = x[1];
                Call(f, "StartSearch");
            } else if (step == 2 + q.Length) {
                Call(f, "OpenHistory");
                HistoryForm h = (HistoryForm)F(f, "LastHistory");
                Console.WriteLine("В истории поисков: " + h.Count);
                foreach (string line in File.ReadAllLines(Path.Combine(here, "istoriya.txt"))) if (line.StartsWith("П")) Console.WriteLine("  " + line.Replace("\t", " | "));
            } else if (step == 4 + q.Length) {
                HistoryForm h = (HistoryForm)F(f, "LastHistory");
                Call(h, "Again", 0);
                Console.WriteLine("Повтор: искать [" + ((TextBox)F(f, "queryBox")).Text + "]");
            } else if (step == 5 + q.Length) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(f);
    }
}
