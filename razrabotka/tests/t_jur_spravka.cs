// Проверка 20.11: «Перевод из юр в физ» -> «Создать справку» (вымышленные данные, без баз)
using System; using System.Windows.Forms; using System.Reflection; using System.Drawing;
public static class T {
    static object F(object o, string n) { return o.GetType().GetField(n, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public).GetValue(o); }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) {
                m.OpenJur();
                JurPick p = (JurPick)F(m, "jurPick");
                p.NoQuestions = true;
                p.SetAddress("д.", "Тестовая", "ул.", "Лесная", "5");
                p.SetQuery("", "12");
                p.SetLs("045671234");
            } else if (step == 3) {
                JurPick p = (JurPick)F(m, "jurPick");
                Console.WriteLine("ЮР: " + p.StatusText);
                p.CreateSpravka();
            } else if (step == 5) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                Console.WriteLine("Справка видна: " + s.Visible + ", перевод скрыт: " + !((Form)F(m, "jurPick")).Visible + ", Back=перевод: " + (F(s, "Back") == F(m, "jurPick")));
                Console.WriteLine("Заголовок: " + ((Control)F(s, "head")).Text);
            } else if (step == 7) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
