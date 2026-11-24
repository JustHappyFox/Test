// ==== БЛОК НАРЯДОВ ИЗ ВЫГРУЗКИ (исходник - razrabotka/src/nar_excel.cs; здесь не править - вставляется скриптом inject.py) ====
// Наряды из Excel-выгрузки (меню программ «Настройки…»; своя копия nar.dbf - «Обновить индексы»): по № наряда (ГодНомер) и по
// лицевому. Поля копии: NZ № наряда, DOF дата оформления, RABOTA работа, NAB лицевой, FIO, NP, ULICA, DOM, KOR, KV, FIO_PR принял,
// FIO_VIP выполнил (пусто - не выполнен), DVIP дата выполнения, TIPCH, NCH счётчик, PRIM примечание. Только чтение; одна загрузка.
public class NarExcel {
    public DbfTable T;
    public string About = "";
    readonly Dictionary<string, List<int>> byNum = new Dictionary<string, List<int>>();
    readonly Dictionary<string, List<int>> byNab = new Dictionary<string, List<int>>();
    static NarExcel cache;
    static string cacheStamp = "";
    static readonly object gate = new object();

    public static NarExcel Open() {
        VygKind k = VygKind.Nar;
        string problem = k.Problem();
        if (problem != null) throw new Exception(problem);
        lock (gate) {
            string stamp = "";
            try { FileInfo fi = new FileInfo(k.Table); stamp = fi.Length + "|" + fi.LastWriteTimeUtc.Ticks; } catch { }
            if (cache != null && stamp == cacheStamp) return cache;
            NarExcel n = new NarExcel();
            n.T = DbfTable.Load(k.Table, false);
            n.About = k.About();
            for (int r = 0; r < n.T.Count; r++) {
                Add(n.byNum, Digits(n.F(r, "NZ")).TrimStart('0'), r);
                Add(n.byNab, Digits(n.F(r, "NAB")).TrimStart('0'), r);
            }
            cache = n;
            cacheStamp = stamp;
            return n;
        }
    }

    static void Add(Dictionary<string, List<int>> d, string key, int row) {
        if (key.Length == 0) return;
        List<int> l;
        if (!d.TryGetValue(key, out l)) { l = new List<int>(); d[key] = l; }
        l.Add(row);
    }

    public string F(int row, string field) {
        int i = T.FieldIndex(field);
        return i < 0 ? "" : DbfTable.Format(T.GetValue(row, i, null), true).Trim();
    }

    static string Digits(string s) {
        StringBuilder sb = new StringBuilder();
        foreach (char c in s ?? "") if (c >= '0' && c <= '9') sb.Append(c);
        return sb.ToString();
    }

    // № наряда полностью, с годом: 202620936 / 2026-20936 / 2026 20936 -> «202620936»; без года (короткий) - null
    public static string Number(string text) {
        string d = Digits(text);
        if (d.Length < 6 || d.Length > 11) return null;
        int y;
        if (!int.TryParse(d.Substring(0, 4), out y) || y < 2000 || y > DateTime.Today.Year + 1) return null;
        return d;
    }

    public List<int> ByNumber(string number) {
        List<int> l;
        return byNum.TryGetValue(Digits(number).TrimStart('0'), out l) ? l : new List<int>();
    }

    // наряды лицевого, новые сверху
    public List<int> ByNab(string nab) {
        List<int> l;
        List<int> r = byNab.TryGetValue(Digits(nab).TrimStart('0'), out l) ? new List<int>(l) : new List<int>();
        int dof = T.FieldIndex("DOF");
        r.Sort(delegate(int a, int b) {
            int c = dof < 0 ? 0 : -DbfTable.CompareValues(T.GetValue(a, dof, null), T.GetValue(b, dof, null));
            return c != 0 ? c : -string.CompareOrdinal(F(a, "NZ"), F(b, "NZ"));
        });
        return r;
    }

    public bool Done(int row) { return F(row, "FIO_VIP").Length > 0 || F(row, "DVIP").Length > 0; }

    // «Макетино, Полевая ул., д.34 кв.12»
    public string Address(int row) {
        List<string> a = new List<string>();
        if (F(row, "NP").Length > 0) a.Add(F(row, "NP"));
        if (F(row, "ULICA").Length > 0) a.Add(F(row, "ULICA"));
        string h = (F(row, "DOM").Length > 0 ? "д." + F(row, "DOM") : "") + (F(row, "KOR").Length > 0 ? " корп." + F(row, "KOR") : "") + (F(row, "KV").Length > 0 ? " кв." + F(row, "KV") : "");
        if (h.Trim().Length > 0) a.Add(h.Trim());
        return string.Join(", ", a.ToArray());
    }

    // строка для списков: «наряд 202620936 от 12.01.2026, замена счётчика - ВЫПОЛНЕН 15.01.2026 (Иванов И.И.)»
    public string Text(int row) {
        string s = "наряд " + F(row, "NZ") + (F(row, "DOF").Length > 0 ? " от " + F(row, "DOF") : "") + (F(row, "RABOTA").Length > 0 ? ", " + F(row, "RABOTA") : "");
        return s + (Done(row) ? " - ВЫПОЛНЕН" + (F(row, "DVIP").Length > 0 ? " " + F(row, "DVIP") : "") + (F(row, "FIO_VIP").Length > 0 ? " (" + F(row, "FIO_VIP") + ")" : "") : " - НЕ ВЫПОЛНЕН");
    }

    // столбцы окон: подпись, поле копии ("" - своё: итог, адрес, счётчик)
    public static readonly string[][] Columns = {
        new string[] { "№ Н-З", "NZ" }, new string[] { "Итог", "" }, new string[] { "Дата оф.", "DOF" }, new string[] { "Работа", "RABOTA" },
        new string[] { "Лицевой", "NAB" }, new string[] { "Ф.И.О.", "FIO" }, new string[] { "Адрес", "" }, new string[] { "Принял", "FIO_PR" },
        new string[] { "Выполнил", "FIO_VIP" }, new string[] { "Дата вып.", "DVIP" }, new string[] { "Счётчик", "" }, new string[] { "Примечание", "PRIM" } };

    public string[] Cells(int row) {
        string[] c = new string[Columns.Length];
        for (int i = 0; i < c.Length; i++) {
            string f = Columns[i][1];
            if (f.Length > 0) c[i] = F(row, f);
            else if (Columns[i][0] == "Итог") c[i] = Done(row) ? "ВЫПОЛНЕН" : "НЕ ВЫПОЛНЕН";
            else if (Columns[i][0] == "Адрес") c[i] = Address(row);
            else c[i] = (F(row, "TIPCH") + " " + F(row, "NCH")).Trim();
        }
        return c;
    }
}
// ==== КОНЕЦ БЛОКА НАРЯДОВ ИЗ ВЫГРУЗКИ ====
