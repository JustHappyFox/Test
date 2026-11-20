// Проверка 20.11: «Поиск» - история = только открытые карточки (30 последних), тот же абонент не дублируется
using System; using System.Windows.Forms; using System.Reflection; using System.IO;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static void Call(object o, string n, params object[] a) { o.GetType().GetMethod(n, BF).Invoke(o, a); }
    static void Find(AbonentSearch f, int mode, string q) {
        ((RadioButton[])F(f, "modeButtons"))[mode].Checked = true;
        ((TextBox)F(f, "queryBox")).Text = q;
        Call(f, "StartSearch");
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        try { File.Delete(Path.Combine(here, "istoriya.txt")); } catch { }
        Application.EnableVisualStyles();
        AbonentSearch f = new AbonentSearch(Path.Combine(here, "poisk.txt"), Path.Combine(here, "nastroiki.txt"));
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 2) Find(f, 0, "010019999");
            else if (step == 3) Call(f, "OpenCard", 0);
            else if (step == 4) Find(f, 1, "Макетино Полевая 34");          // без карточки - в историю не идёт
            else if (step == 5) Find(f, 3, "Шаблонов Мария");
            else if (step == 6) Call(f, "OpenCard", 1);
            else if (step == 7) Call(f, "OpenCard", 1);                      // тот же абонент ещё раз - одна запись
            else if (step == 8) {
                Call(f, "OpenHistory");
                HistoryForm h = (HistoryForm)F(f, "LastHistory");
                Console.WriteLine("В истории: " + h.Count + " (ждём 2)");
                foreach (string line in File.ReadAllLines(Path.Combine(here, "istoriya.txt"))) if (line.StartsWith("П") || line.StartsWith("С")) Console.WriteLine("  " + line.Replace("\t", " | "));
            } else if (step == 10) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(f);
    }
}
