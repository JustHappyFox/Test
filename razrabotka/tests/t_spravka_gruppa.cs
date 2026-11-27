// Проверка 27.11: справка (быт) - (1) тарифная группа у каждого прибора учёта: два ПУ -> «ПУ 1 - ...», «ПУ 2 - ...»; (2) тип из перевода
// из юр в физ («CE318BY» без корпуса и модификации) - выбор корпуса и модификации меняет «Тип в справке». Аргумент: папка программы.
using System; using System.Windows.Forms; using System.Reflection; using System.Collections.Generic;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want; if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    [STAThread] public static void Main(string[] a) {
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(a[0], "проверка");
        Timer t = new Timer(); t.Interval = 2000; int step = 0;
        t.Tick += delegate {
            step++;
            if (step == 1) { m.OpenSpravka(m, null); return; }
            SpravkaForm s = (SpravkaForm)F(m, "spr");
            MethodInfo sv = s.GetType().GetMethod("SetV", BF), v = s.GetType().GetMethod("V", BF);
            Action<string, string> Set = delegate(string k, string x) { sv.Invoke(s, new object[] { k, x }); };
            Func<string, string> Get = delegate(string k) { return (string)v.Invoke(s, new object[] { k }); };
            if (step == 2) {
                Console.WriteLine("--- тарифная группа по ПУ");
                Set("on_gruppa", "1");
                Set("pu1_gruppa", "(26) \"Жилой дом, ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\"");
                Eq("один ПУ", (string)s.Compose()["ГРУППА"], "(26) \"Жилой дом, ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\".");
                s.AddMeter();
                Set("pu2_gruppa", "(31) отопление");
                Eq("два ПУ", ((string)s.Compose()["ГРУППА"]).Replace("\n", " / "), " / ПУ 1 - (26) \"Жилой дом, ОБОРУДОВАННЫЙ в установленном порядке электрической плитой\". / ПУ 2 - (31) отопление.");
                Set("on_gruppa", "");
                Eq("галочка снята - строки нет", (string)s.Compose()["ГРУППА"] == DocxTemplate.Drop ? "нет" : "есть", "нет");
            } else if (step == 3) {
                Console.WriteLine("--- тип из перевода (только модель) -> выбор корпуса и модификации");
                Dictionary<string, string> d = new Dictionary<string, string>();
                d["pu1_model"] = "CE318BY"; d["pu1_korpus"] = ""; d["pu1_modif"] = ""; d["pu1_tip"] = "CE318BY"; d["schetchikov"] = "1 счётчик";
                s.Clear();
                s.FromContract(d);
                Eq("тип после перевода", Get("pu1_tip"), "CE318BY");
                Set("pu1_korpus", "R32");
                Set("pu1_modif", "146 JP.UVFL");
                Eq("выбрали корпус и модификацию - тип", Get("pu1_tip"), "CE318BY R32 146 JP.UVFL");
                Set("pu1_modif", "043 JA.UVFL");
                Eq("сменили модификацию - тип", Get("pu1_tip"), "CE318BY R32 043 JA.UVFL");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
                t.Stop(); Environment.Exit(0);
            }
        };
        t.Start(); Application.Run(m);
    }
}
