// Проверка 20.11: меню -> «Настройки…»: пути, «Обновить индексы» (вымышленная выгрузка), сохранение. Аргумент 2 - путь к xlsx.
using System; using System.Windows.Forms;
public static class T {
    [STAThread] public static void Main(string[] a) {
        string here = a[0], xlsx = Environment.GetEnvironmentVariable("T_XLSX");
        Application.EnableVisualStyles();
        MenuForm m = new MenuForm(here);
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) m.BeginInvoke((MethodInvoker)delegate { m.OpenDataSettings(); });           // модальное окно - дальше шаги идут, пока оно открыто
            else if (step == 2) { m.DataForm.SetPaths(xlsx, Environment.GetEnvironmentVariable("T_JUR") ?? "", Environment.GetEnvironmentVariable("T_NAR") ?? "/нет/такого/файла.xlsx", true); }
            else if (step == 4) { m.DataForm.Rebuild(); }
            else if (step > 4 && m.DataForm != null && !m.DataForm.Building && m.DataForm.LastResult.Length > 0 && step < 100) {
                Console.WriteLine("Итог: " + m.DataForm.LastResult);
                step = 100;
            } else if (step == 102) { m.DataForm.AcceptButton.PerformClick(); }
            else if (step == 103) {
                Console.WriteLine("Режим Excel: " + Istochnik.UseExcel + "; быт: " + (Istochnik.BytPath == xlsx) + "; копия: " + (Istochnik.BytCacheProblem() ?? "готова"));
                Console.WriteLine("Юр: " + Istochnik.Missing("юр лиц", Istochnik.JurPath));
                Console.WriteLine("Наряды: " + Istochnik.Missing("нарядов", Istochnik.NaryadyPath));
            } else if (step == 105) { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
}
