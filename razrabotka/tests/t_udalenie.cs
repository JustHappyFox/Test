// Проверка 24.11: справка в расчетную группу (удаление ПУ из юр) - из перевода (кв.12: быт + отопление) и из сохранённой справки
using System; using System.Windows.Forms; using System.Reflection; using System.Collections.Generic;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static void Dump(string title, UdalenieForm u) {
        Console.WriteLine(title + ": " + u.StatusText);
        foreach (KeyValuePair<string, string> kv in u.Compose()) Console.WriteLine("   " + kv.Key + " = " + kv.Value.Replace("\n", " / "));
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0];
        DocPreview.Open = delegate(string p) { };
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 2500; int step = 0;
        t.Tick += delegate {
            step++;
            JurPick p = (JurPick)F(m, "jurPick");
            if (step == 1) m.OpenJur();
            else if (step == 2) { p.NoQuestions = true; p.SetAddress("п.", "Макетино", "ул.", "Полевая", "34"); p.SetQuery("", "12"); p.Search(); }
            else if (step == 4) p.CreateUdalenie();
            else if (step == 5) {
                UdalenieForm u = (UdalenieForm)F(m, "udal");
                Dump("Из перевода", u);
                u.SetV("dog_name", "ООО «Тестстрой»");
                Console.WriteLine("Сохранена: " + u.Save(false, false));
                u.GoBack();
            } else if (step == 6) p.CreateSpravka();
            else if (step == 7) {
                SpravkaForm s = (SpravkaForm)F(m, "spr");
                Console.WriteLine("Справка: данные юр лица запомнены - " + (s.JurInfo != null) + ", галочка «+ удалить из юр» есть - " + (F(s, "withUdal") != null));
                s.GetType().GetMethod("SetV", BF).Invoke(s, new object[] { "pu1_pokaz", "00003671" });
                s.Save(false, false);
                m.OpenUdalenie(s, s.JurInfo, s.Data());
            } else if (step == 8) Dump("Из справки", (UdalenieForm)F(m, "udal"));
            else if (step == 10) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
