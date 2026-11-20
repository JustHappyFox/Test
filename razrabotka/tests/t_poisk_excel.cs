// Проверка 20.11: «Поиск» в режиме Excel-выгрузки (вымышленная выгрузка gen_byt.py). Шаги - через отражение.
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static void Call(object o, string n) { o.GetType().GetMethod(n, BF, null, new Type[0], null).Invoke(o, null); }
    static string Status(AbonentSearch f) { return ((Control)F(f, "status")).Text; }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        AbonentSearch f = new AbonentSearch(Path.Combine(here, "poisk.txt"), Path.Combine(here, "nastroiki.txt"));
        string[][] q = { new string[] { "0", "010019999" }, new string[] { "3", "Шаблонов Мария" }, new string[] { "1", "Макетино Полевая 34" },
                         new string[] { "2", "50000999" }, new string[] { "3", "Тестов Иван Иванович Тестовка" } };
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) Console.WriteLine("Открытие: " + Status(f));
            else if (step >= 2 && step < 2 + q.Length) {
                if (step > 2) Console.WriteLine("  -> " + Status(f) + " | строк в списке: " + ((DataGridView)F(f, "grid")).RowCount);
                string[] x = q[step - 2];
                ((RadioButton[])F(f, "modeButtons"))[int.Parse(x[0])].Checked = true;
                ((TextBox)F(f, "queryBox")).Text = x[1];
                Call(f, "StartSearch");
                Console.WriteLine("Ищу [" + x[0] + "] " + x[1]);
            } else if (step == 2 + q.Length) {
                Console.WriteLine("  -> " + Status(f) + " | строк в списке: " + ((DataGridView)F(f, "grid")).RowCount);
                ((RadioButton[])F(f, "modeButtons"))[0].Checked = true;
                ((TextBox)F(f, "queryBox")).Text = "010019999";
                Call(f, "StartSearch");
            } else if (step == 3 + q.Length) {
                f.GetType().GetMethod("OpenCard", BF).Invoke(f, new object[] { 0 });
            } else if (step == 5 + q.Length) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(f);
    }
}
