// Проверка 26.11: «Фото с телефона» на экране нетбука (SCREEN=1024x560): свой календарь (один месяц, тёмный), под ним «с ... по»,
// «Сегодня / Вчера / Все дни» - в окне; вымышленный телефон (результат осмотра подаётся сразу, без оболочки Windows):
// 8 фото за 3 дня; папка выключена - подсказка «поставьте галочку». Аргументы: папка программы, png-префикс.
using System; using System.Windows.Forms; using System.Reflection; using System.Drawing; using System.Collections.Generic;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    static void Snap(string file) { Application.DoEvents(); try { System.Diagnostics.Process.Start("import", "-window root " + file).WaitForExit(); } catch { } }
    static bool Inside(Control c, Form f) {
        Rectangle r = f.RectangleToClient(c.RectangleToScreen(c.ClientRectangle));
        return r.Bottom <= f.ClientSize.Height && r.Right <= f.ClientSize.Width && c.Visible;
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a[1];
        Application.EnableVisualStyles();
        FotoForm f = new FotoForm(here);
        f.WindowState = FormWindowState.Maximized;
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) {
                ScanResult r = new ScanResult();
                r.Devices.Add("Redmi 8A");
                string src = "Redmi 8A: Внутренний общий накопитель\\DCIM\\Camera";
                r.Sources.Add(src);
                DateTime[] days = { DateTime.Today, DateTime.Today, DateTime.Today.AddDays(-1), DateTime.Today.AddDays(-1), DateTime.Today.AddDays(-1),
                    DateTime.Today.AddDays(-9), DateTime.Today.AddDays(-9), DateTime.Today.AddDays(-9) };
                for (int i = 0; i < days.Length; i++) {
                    PhoneFile p = new PhoneFile();
                    p.Device = "Redmi 8A"; p.Folder = "Внутренний общий накопитель\\DCIM\\Camera"; p.Name = "IMG_" + i + ".jpg";
                    p.Modified = p.Taken = p.Date = days[i].AddHours(10 + i); p.Size = 3000000; p.Kind = "фото";
                    r.Files.Add(p);
                }
                f.GetType().GetMethod("ScanDone", BF).Invoke(f, new object[] { r });
                Application.DoEvents();
                ListView dl = (ListView)F(f, "days");
                Eq("дней с фото", dl.Items.Count.ToString(), "3");
                Control cal = (Control)F(f, "cal");
                Eq("календарь в окне", Inside(cal, f).ToString(), "True");
                foreach (Control c in cal.Parent.Controls)
                    if (c is Button) Eq("кнопка «" + c.Text + "» в окне", Inside(c, f).ToString(), "True");
                Snap(png + "_1.png");
            } else if (step == 2) {
                // «Все дни»
                foreach (Control c in ((Control)F(f, "cal")).Parent.Controls) if (c is Button && c.Text == "Все дни") ((Button)c).PerformClick();
                Application.DoEvents();
                Console.WriteLine("   " + ((Label)F(f, "chosen")).Text);
                Snap(png + "_2.png");
            } else if (step == 3) {
                CheckedListBox s = (CheckedListBox)F(f, "sources");
                s.SetItemChecked(0, false);
                Application.DoEvents();
                Eq("папка выключена - подсказка", ((Label)F(f, "chosen")).Text, "Папки телефона выключены - поставьте галочку в списке «Откуда»");
                Snap(png + "_3.png");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        f.Shown += delegate { t.Start(); };
        Application.Run(f);
    }
}
