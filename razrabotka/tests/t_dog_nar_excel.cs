// Проверка 24.11: «Договор» - «Найти в базе…» по № наряда из выгрузки нарядов (вымышленные выгрузки)
using System; using System.Collections.Generic; using System.Windows.Forms;
public static class T {
    [STAThread] public static void Main(string[] a) {
        string err, note;
        BaseLink b = BaseShare.Get(new Dictionary<string, string>(), out err);
        Console.WriteLine("база: " + (b == null ? err : "открыта, Excel=" + b.Excel));
        foreach (string q in new string[] { "2025-10000", "202610001", "202699999", "10002" }) {
            List<Dictionary<string, string>> r = b.Find("nar", q, out note);
            Console.WriteLine(q + ": найдено " + r.Count + (note != null ? " (" + note + ")" : ""));
            foreach (Dictionary<string, string> d in r) {
                string nar, fio;
                d.TryGetValue("naryad", out nar); d.TryGetValue("FIO", out fio);
                Console.WriteLine("   " + fio + " | " + nar);
            }
        }
        Environment.Exit(0);
    }
}
