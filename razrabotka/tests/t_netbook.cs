// Проверка 24.11: маленький экран (нетбук 1024x600) - каждое окно «Договора» не больше экрана, пункты колонки слева не налезают
// друг на друга. Запуск: SCREEN=1024x560 runtest.sh dogovor t_netbook.cs <папка> <png> 1 (скриншоты делает сам тест: png-префикс - a[1])
using System; using System.Drawing; using System.Windows.Forms; using System.Reflection; using System.Collections.Generic; using System.Diagnostics;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static void Call(object o, string n, params object[] a) {
        foreach (MethodInfo mi in o.GetType().GetMethods(BF)) if (mi.Name == n && mi.GetParameters().Length == a.Length) { mi.Invoke(o, a); return; }
        Console.WriteLine("нет метода " + n);
    }
    static void Check(Form f, string png) {
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        bool fits = wa.Contains(f.Bounds);
        Console.WriteLine(f.GetType().Name + " " + f.Bounds + (fits ? " - в экране" : " - НЕ ПОМЕЩАЕТСЯ в " + wa));
        Check(f, f);
        Process.Start("import", "-window root " + png).WaitForExit();
    }
    // всё видимое внутри окна (кнопки внизу, колонка слева) - в пределах окна; пункты колонки не перекрываются
    static void Check(Form f, Control c) {
        foreach (Control k in c.Controls) {
            if (!k.Visible) continue;
            Rectangle r = f.RectangleToClient(k.Parent.RectangleToScreen(k.Bounds));
            if ((k is Button || k is NavItem) && (r.Bottom > f.ClientSize.Height + 1 || r.Right > f.ClientSize.Width + 1) && !(c is ScrollableControl && ((ScrollableControl)c).AutoScroll))
                Console.WriteLine("   за краем: " + k.GetType().Name + " «" + k.Text + "» " + r);
            if (c.GetType() == typeof(Panel) && k is NavItem)
                foreach (Control o in c.Controls) if (o != k && o.Visible && (o is NavItem || o is Label) && o.Bounds.IntersectsWith(k.Bounds) && o.Text.Length > 0)
                    Console.WriteLine("   перекрытие в колонке: «" + k.Text + "» " + k.Bounds + " и «" + o.Text + "» " + o.Bounds + ", колонка " + c.ClientSize);
            Check(f, k);
        }
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1];
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        string[][] steps = { new string[] { "OpenDogovor" }, new string[] { "OpenSpravka" }, new string[] { "OpenNaryad" }, new string[] { "OpenJurNaryad" },
            new string[] { "OpenLog" }, new string[] { "OpenDogLog" }, new string[] { "OpenJur" }, new string[] { "OpenBs" }, new string[] { "OpenPlomby" },
            new string[] { "OpenUdalenie" }, new string[] { "OpenJurSettings" }, new string[] { "OpenBaseSetup" } };
        Timer t = new Timer(); t.Interval = 2500; int step = -1; List<Form> seen = new List<Form>();
        t.Tick += delegate {
            step++;
            if (step == 0) { Console.WriteLine("Экран: " + Screen.PrimaryScreen.WorkingArea + ", маленький - " + St.Compact); Check(m, png + "_00.png"); return; }
            int k = (step - 1) / 2;
            if (k >= steps.Length) { t.Stop(); Environment.Exit(0); }
            if ((step - 1) % 2 == 0) {
                foreach (Form f in Application.OpenForms) seen.Add(f);
                string n = steps[k][0];
                if (n == "OpenDogovor" || n == "OpenJur" || n == "OpenPlomby") Call(m, n);
                else if (n == "OpenNaryad") Call(m, n, m, null, null);
                else if (n == "OpenSpravka") Call(m, n, m, null);
                else if (n == "OpenUdalenie") Call(m, n, m, null, null);
                else if (n == "OpenJurSettings" || n == "OpenBaseSetup") { string nn = n; m.BeginInvoke((MethodInvoker)delegate { Call(m, nn, m); }); }
                else Call(m, n, m);
            } else {
                Form top = null;
                foreach (Form f in Application.OpenForms) if (f.Visible && !seen.Contains(f)) top = f;
                if (top == null) { Console.WriteLine(steps[k][0] + ": окно не открылось"); return; }
                Check(top, png + "_" + (k + 1).ToString("00") + ".png");
                if (top.Modal) top.Close(); else top.Hide();
            }
        };
        t.Start();
        Application.Run(m);
    }
}
