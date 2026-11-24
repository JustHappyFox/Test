// Проверка 24.11: копии выгрузок нарядов и юр лиц (xlsx) общим блоком. Аргументы: наряды.xlsx юр.xlsx
using System; using System.IO; using System.Text; using System.Collections.Generic; using System.Xml;
public static class T {
    public static void Main(string[] a) {
        Istochnik.Save(true, Istochnik.BytPath, a[1], a[0]);
        Console.WriteLine("до: " + VygKind.Nar.Problem());
        Console.WriteLine(VygruzkaByt.Build(VygKind.Nar, a[0], null));
        Console.WriteLine(VygruzkaByt.Build(VygKind.Jur, a[1], null));
        Console.WriteLine("наряды: " + (VygKind.Nar.Problem() ?? "готово") + " | " + VygKind.Nar.About());
        Console.WriteLine("юр: " + (VygKind.Jur.Problem() ?? "готово") + " | " + VygKind.Jur.About());
        Console.WriteLine("юр .dbf: " + (VygKind.Jur.Problem("/нет/такого.dbf") ?? "готово"));
    }
}
