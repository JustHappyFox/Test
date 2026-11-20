// Проверка общего блока выгрузок отдельно от программ: создать копию выгрузки быта и прочитать её заголовок.
// Собирается вместе с razrabotka/shared/vygruzka.cs (см. razrabotka/tools/vygtest.sh)
using System; using System.IO; using System.Text; using System.Collections.Generic; using System.Xml;
public static class T {
    public static void Main(string[] a) {
        Istochnik.Save(true, a[0], "", "");
        Console.WriteLine("Режим Excel: " + Istochnik.UseExcel + ", до создания: " + Istochnik.BytCacheProblem());
        string r = VygruzkaByt.Build(a[0], delegate(string s) { Console.WriteLine("  " + s); });
        Console.WriteLine(r);
        Console.WriteLine("после: " + (Istochnik.BytCacheProblem() ?? "готово") + "; " + Istochnik.BytAbout());
        Console.WriteLine("Пустой путь юр: " + Istochnik.Missing("юр лиц", Istochnik.JurPath));
    }
}
