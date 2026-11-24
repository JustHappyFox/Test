// Проверка 24.11: копия выгрузки из файлов разного вида (.xlsx, .xls 97/95/2.1, XML 2003, HTML, текст) - должна быть одинаковой.
// vygtest.sh t_xls_formats.cs файл..., затем сравнить <файл>.dbf (tests/gen_xls_old.py делает файлы из gen_byt.py)
using System; using System.IO;
public static class T {
    public static void Main(string[] a) {
        foreach (string f in a) {
            DateTime t0 = DateTime.Now;
            try {
                string r = VygruzkaByt.Build(VygKind.Byt, f, null);
                File.Copy(VygKind.Byt.Table, f + ".dbf", true);
                Console.WriteLine(Path.GetFileName(f) + ": " + r.Replace("\n", " | ") + " (" + (DateTime.Now - t0).TotalSeconds.ToString("0.0") + " с)");
            } catch (Exception ex) { Console.WriteLine(Path.GetFileName(f) + ": ОШИБКА " + ex); }
        }
    }
}
