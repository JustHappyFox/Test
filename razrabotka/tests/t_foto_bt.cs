// Проверка 26.11: «Фото с телефона» - Bluetooth. Имитация (здесь нет Bluetooth и Windows): устройства - вымышленный список,
// мастер Windows «Получение файлов» - «кладёт» вымышленные фото в папку «Документы». Аргументы: папка, png-префикс.
// Снимки: нет адаптера, нет драйвера, отключён, работает, после приёма. Фото для приёма: T_BT_SRC (IMG_20261103_*.jpg,
// без даты в имени - с EXIF 15.10.2026).
using System; using System.Windows.Forms; using System.IO; using System.Collections.Generic; using System.Drawing;
public static class T {
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    static List<string[]> devs = new List<string[]>();
    static Dictionary<string, string> hw = new Dictionary<string, string>();
    static void Snap(Form f, string file) {
        Application.DoEvents();
        // снимок экрана (DrawToBitmap у диалога под Mono пустой)
        try { System.Diagnostics.Process.Start("import", "-window root " + file).WaitForExit(); } catch { }
    }
    [STAThread] public static void Main(string[] a) {
        string dir = a[0], png = a[1], src = Environment.GetEnvironmentVariable("T_BT_SRC");
        string docs = Path.Combine(dir, "Документы"), root = Path.Combine(dir, "Фото");
        Directory.CreateDirectory(docs);
        Bt.Query = delegate { return new List<string[]>(devs); };
        Bt.Ids = delegate(string id) { string v; return hw.TryGetValue(id, out v) ? v : ""; };
        Bt.Folders = new List<string>(new string[] { docs });
        int ticks = 0;
        Bt.Start = delegate(string exe, string args) {
            Console.WriteLine("   запуск: " + Path.GetFileName(exe) + " " + args);
            if (args != "-receive") return delegate { return true; };
            // как Windows: принятый файл создан сейчас
            foreach (string f in Directory.GetFiles(src)) { string g = Path.Combine(docs, Path.GetFileName(f)); File.Copy(f, g, true); File.SetCreationTime(g, DateTime.Now); File.SetLastWriteTime(g, DateTime.Now); }
            ticks = 0;
            return delegate { return ++ticks > 2; };
        };
        Application.EnableVisualStyles();
        BtDialog d = new BtDialog(root);
        Timer t = new Timer(); t.Interval = 1200; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) {
                Console.WriteLine("--- нет адаптера (выключен кнопкой)");
                devs.Clear(); devs.Add(new string[] { "Тачпад", "ACPI\\SYN0001\\4&1", "0", "{4d36e96f-e325-11ce-bfc1-08002be10318}" });
                d.Recheck();
                Eq("работает", d.State.Works.ToString(), "False");
                Console.WriteLine("   " + d.State.Text.Substring(0, 60) + "…");
                Snap(d, png + "_1.png");
            } else if (step == 2) {
                Console.WriteLine("--- адаптер без драйвера (код 28, виден по USB-классу E0)");
                devs.Add(new string[] { "Неизвестное устройство", "USB\\VID_0CF3&PID_3005\\5&2", "28", "" });
                hw["USB\\VID_0CF3&PID_3005\\5&2"] = "USB\\VID_0CF3&PID_3005&REV_0001|USB\\VID_0CF3&PID_3005|USB\\Class_E0&SubClass_01&Prot_01|USB\\Class_E0";
                d.Recheck();
                Eq("работает", d.State.Works.ToString(), "False");
                Eq("для поиска драйвера", d.State.Ids, "Неизвестное устройство - USB\\VID_0CF3&PID_3005");
                Snap(d, png + "_2.png");
            } else if (step == 3) {
                Console.WriteLine("--- отключён в диспетчере (код 22)");
                devs[1] = new string[] { "Generic Bluetooth Adapter", "USB\\VID_0CF3&PID_3005\\5&2", "22", Bt.Class };
                d.Recheck();
                Eq("отключён", d.State.Text.Contains("отключён").ToString(), "True");
            } else if (step == 4) {
                Console.WriteLine("--- работает (дети BTHENUM не считаются адаптером)");
                devs[1] = new string[] { "Generic Bluetooth Adapter", "USB\\VID_0CF3&PID_3005\\5&2", "0", Bt.Class };
                devs.Add(new string[] { "Bluetooth Peripheral Device", "BTHENUM\\{0000110a}_LOCALMFG&0000\\7&1", "28", "" });
                d.Recheck();
                Eq("работает", d.State.Works.ToString(), "True");
                Snap(d, png + "_3.png");
                Button rcv = null;
                foreach (Control c in d.Controls) if (c is Button && c.Text == "Принять фото…") rcv = (Button)c;
                rcv.PerformClick();
            } else if (step == 7) {
                Console.WriteLine("--- после приёма");
                Console.WriteLine("   " + d.ResultText);
                Eq("разложено", d.Last == null ? "нет" : d.Last.Placed.Count.ToString(), "3");
                Eq("по имени 03.11", Directory.Exists(Path.Combine(root, "2026\\11\\03".Replace('\\', Path.DirectorySeparatorChar))) ? Directory.GetFiles(Path.Combine(Path.Combine(Path.Combine(root, "2026"), "11"), "03")).Length.ToString() : "0", "2");
                Eq("по EXIF 15.10", Directory.Exists(Path.Combine(Path.Combine(Path.Combine(root, "2026"), "10"), "15")).ToString(), "True");
                Eq("в «Документах» не осталось", Directory.GetFiles(docs).Length.ToString(), "0");
                Snap(d, png + "_4.png");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        d.Shown += delegate { t.Start(); };
        Application.Run(d);
    }
}
