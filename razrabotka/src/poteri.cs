// ==== БЛОК РАСЧЁТА ПОТЕРЬ (исходник - razrabotka/src/poteri.cs; вставляется скриптом inject.py) ====
// «Потери (Приложение 8)» пакета документов: схема от границы балансовой принадлежности (ГБП) до учёта -> порядок расчёта по
// ТКП 460-2017 (33240) с формулами Word (OMML) и подставленными числами. Справочник силовых трансформаторов - табл. Б.1-Б.3.
// Элементы - от учёта к границе (п. 4.12): ТТ/ТН учёта, линии на напряжении учёта, силовой трансформатор, линии за ним.
public class TrafoInfo {
    public string Type = "";
    public double Unt, R, X, DP, DQ;        // кВ, Ом, Ом, кВт, квар
}

public static class Trafo {
    static readonly System.Globalization.CultureInfo Inv = System.Globalization.CultureInfo.InvariantCulture;
    // тип Rт Xт ΔPхх ΔQхх (Б.1 - 6 кВ, Б.2 - 10 кВ, Б.3 - 35 кВ)
    static readonly string B1 = "ОМ-1,25/6-У1 1382 1584 0.02 0.238;ОМ-2,5/6-У1 501 514 0.04 0.375;ОМП-4/6-У1 269 116 0.045 0.2;ОМП-10/6-У1 64 73 0.07 0.55;" +
        "ТМ-10/6 133 218 0.105 1;ТМ-20/6 60 109 0.18 1.8;ТМ-25/6 40 73 0.115 0.8;ТМ-30/6 37.6 72.6 0.25 2.4;ТМ-40/6 23 46 0.165 1.2;ТМ-50/6 21 43.6 0.35 3.5;" +
        "ТМ-63/6 13.7 29 0.243 1.7;ТМ-100/6 8.4 18.4 0.338 2.6;ТМ-160/6 4.5 11.5 0.5 3.8;ТМ-180/6 4.9 12.1 0.7 7.2;ТМ-250/6 2.4 7.9 1.04 8.8;" +
        "ТМ-320/6 2.4 6.8 1.6 12.8;ТМ-400/6 1.4 4.5 1.0 8.4;ТМ-530/6 0.8 3.4 1.55 12.6;ТМ-1000/6 0.47 2.2 2.3 14;ТМ-1600/6 0.27 1.4 3.05 20.8;" +
        "ТМ-2500/6 0.15 0.87 4.25 25;ТМ-4000/6 0.08 0.64 5.975 36;ТМ-6300/6 0.05 0.4 8.325 50.4;ТМ-10000/6 0.04 0.3 29.0 300";
    static readonly string B2 = "ОМ-1,25/10-У1 3840 4400 0.02 0.238;ОМ-2,5/10-У1 1361 1504 0.04 0.375;ОМП-4/10-У1 775 328 0.045 0.2;ОМП-10/10-У1 180 203 0.07 0.55;" +
        "ТМ-10/10 369 605 0.14 1;ТМ-20/10 165 302 0.22 2;ТМ-25/10 110 202 0.12 0.8;ТМ-30/10 103 202 0.3 2.7;ТМ-40/10 64.7 126 0.165 1.2;" +
        "ТМ-50/10 58.5 121 0.44 4.0;ТМ-63/10 38.1 80.4 0.243 1.8;ТМ-100/10 23.4 50.5 0.338 2.6;ТМ-160/10 12.4 31.6 0.5 3.8;ТМ-180/10 13.8 33.6 0.7 7.2;" +
        "ТМ-250/10 6.9 20.2 0.72 5.8;ТМ-320/10 6.6 19.0 1.125 12.8;ТМ-400/10 3.9 12.6 1.0 8.4;ТМ-530/10 3.2 10.8 2.9 33.6;ТМ-630/10 2.2 9.6 1.55 12.6;" +
        "ТМ-1000/10 1.3 6.0 2.275 14;ТМ-1600/10 1.03 3.8 8.0 72.0;ТМ-1800/10 0.8 3.4 8.0 81;ТМ-3200/10 0.4 1.9 11.0 128;ТМ-5600/10 0.2 1.1 18 224;" +
        "ТМ-6300/10 0.13 1.14 8.3 50.4;ТМ-7500/10 0.14 1.1 24 262";
    static readonly string B3 = "ТМ-100/35 266 875 0.47 2.6;ТМ-160/35 140 549 0.66 3.84;ТМ-180/35 170 481 1.5 14.4;ТМ-250/35 80 350 0.96 5.75;ТМ-320/35 82 273 2.3 24;" +
        "ТМ-400/35 46.6 219 1.35 8.4;ТМ-630/35 26 140 2 12.6;ТМ-1000/35 16 87.5 2.75 15;ТМ-1600/35 9.1 54.9 3.65 22.4;ТМ-4000/35 2.8 25.2 6.7 4;" +
        "ТМ-6300/35 1.6 16.1 9.4 57;ТМ-7500/35 1.81 13.6 24 262;ТД-10000/35 1.25 10.1 29 300;ТД-16000/35 0.48 6.75 21 120;ТД-20000/35 0.41 5.75 39 600;" +
        "ТДН-25000/35 0.27 4.3 29 175;ТД-40000/35 0.15 2.87 39 260;ТД-80000/35 0.07 1.53 65 480";

    public static List<string> Types() {
        List<string> r = new List<string>();
        foreach (string t in new string[] { B2, B1, B3 }) foreach (string row in t.Split(';')) r.Add(row.Split(' ')[0]);
        return r;
    }

    // «ТМ-250/10», «тм 250/10», «ТМГ-250/10» (ТМГ - как ТМ той же мощности, сверить с паспортом) -> данные таблицы Б
    public static TrafoInfo Find(string type) {
        string t = (type ?? "").Trim().ToUpperInvariant().Replace(" ", "-").Replace("ТМГ", "ТМ").Replace("ТМЗ", "ТМ");
        if (t.Length == 0) return null;
        string[][] tabs = { new string[] { B1, "6" }, new string[] { B2, "10" }, new string[] { B3, "35" } };
        foreach (string[] tab in tabs)
            foreach (string row in tab[0].Split(';')) {
                string[] p = row.Split(' ');
                if (p[0].ToUpperInvariant() != t) continue;
                TrafoInfo i = new TrafoInfo();
                i.Type = p[0];
                i.Unt = double.Parse(tab[1], Inv);
                i.R = double.Parse(p[1], Inv); i.X = double.Parse(p[2], Inv); i.DP = double.Parse(p[3], Inv); i.DQ = double.Parse(p[4], Inv);
                return i;
            }
        return null;
    }
}

public class PotLine {
    public string Name = "", Mark = "", Kind = "КЛ";
    public double U, L, R0, R, X0 = -1, X = -1, Q30 = -1, TgD;
    public bool Low { get { return U < 1; } }
    public bool OnePhase { get { return U < 0.3; } }
}

// один шаг расчёта: пояснение и формулы (запись - как TeX: \frac{a}{b}, x_{a}, x^{2}, \text{слово})
public class PotStep {
    public string Text = "";
    public List<string> Math = new List<string>();
    public PotStep(string text) { Text = text; }
}

public class PotScheme {
    static readonly System.Globalization.CultureInfo Inv = System.Globalization.CultureInfo.InvariantCulture;
    public double Uuch = 0.38;                   // напряжение учёта, кВ
    public int Uchet;                            // 0 - прямого включения, 1 - через ТТ, 2 - через ТТ и ТН
    public bool Aron;                            // ТТ по схеме Арона (неполнофазная)
    public TrafoInfo Tr;                         // силовой трансформатор между ГБП и учётом (null - нет)
    public List<PotLine> Lines = new List<PotLine>();
    public bool Abonent = true;                  // участок на балансе абонента: потери прибавляются к учёту, иначе вычитаются
    public double D = 1.01, Tg = 0.33;
    public string Rezhim = "непрерывная работа", TgWhat = "остальные потребители";
    public bool WpMeter;                         // реактивная - по показаниям счётчика
    public bool Chain = true;                    // потери предыдущих элементов - к энергии следующего (п. 4.12)
    public bool Reactive;                        // считать реактивную составляющую ТРЭТ
    public bool Popr;                            // строка «корректирующая поправка и учёт с коррекцией» (24.11: по умолчанию нет)

    public static string N(double d) {
        if (Math.Abs(d) >= 100) return Math.Round(d, 1).ToString("0.#", Inv).Replace('.', ',');
        return Math.Round(d, 5).ToString("0.#####", Inv).Replace('.', ',');
    }
    public static string Uv(double u) { return u < 0.3 ? "0,22" : u < 1 ? "0,38" : N(u); }

    // таблица 6.2 (ТТ) и 6.3 (ТН), кВт·ч в месяц
    public double TT() {
        if (Uchet == 0) return 0;
        double[] u = { 0.38, 6, 10, 35, 110 }, full = { 4.2, 5.0, 8.3, 33.3, 91.7 }, aron = { 2.8, 3.3, 5.6, 22.2, 61.1 };
        for (int i = 0; i < u.Length; i++) if (Math.Abs(Uuch - u[i]) < 0.5 || (Uuch < 1 && i == 0)) return Aron ? aron[i] : full[i];
        return 0;
    }
    public double TN() {
        if (Uchet != 2) return 0;
        double[] u = { 6, 10, 35, 110 }, w = { 128, 158, 300, 917 };
        for (int i = 0; i < u.Length; i++) if (Math.Abs(Uuch - u[i]) < 0.5) return w[i];
        return 0;
    }

    // линии на напряжении учёта (до трансформатора, считая от учёта) и за трансформатором
    public List<PotLine> Near() { List<PotLine> r = new List<PotLine>(); foreach (PotLine l in Lines) if (Tr == null || Same(l.U, Uuch)) r.Add(l); return r; }
    public List<PotLine> Far() { List<PotLine> r = new List<PotLine>(); if (Tr != null) foreach (PotLine l in Lines) if (!Same(l.U, Uuch)) r.Add(l); return r; }
    static bool Same(double a, double b) { return (a < 1 && b < 1) || Math.Abs(a - b) < 0.5; }

    string Sg { get { return Abonent ? "+" : "-"; } }

    // ---- порядок расчёта: шаги с формулами (числа подставлены, Wa и Wp - за расчётный месяц) ----
    public List<PotStep> Steps() {
        List<PotStep> s = new List<PotStep>();
        PotStep p;
        if (!WpMeter) {
            p = new PotStep("Потребление реактивной энергии за расчетный месяц при отсутствии показаний расчетных приборов учета (формула (7), tgφ - таблица 6.4, " + TgWhat + "), квар·ч:");
            p.Math.Add("W_{р}=W_{а}·tgφ=W_{а}·" + N(Tg));
            s.Add(p);
        }
        List<string> sumA = new List<string>(), sumP = new List<string>();
        string prevA = "", prevP = "";             // накопленные потери предыдущих элементов (символами)
        if (TT() > 0) {
            p = new PotStep("Потери в трехфазной группе измерительных трансформаторов тока " + Uv(Uuch) + " кВ (" + (Aron ? "неполнофазная схема подключения счетчика (схема Арона)" : "полнофазная схема подключения счетчика") + ", таблица 6.2), кВт·ч:");
            p.Math.Add("ΔW_{а ТТ}=" + N(TT()));
            s.Add(p);
            sumA.Add("ΔW_{а ТТ}");
            if (Chain) prevA += Sg + "ΔW_{а ТТ}";
        }
        if (TN() > 0) {
            p = new PotStep("Потери в измерительных трансформаторах напряжения " + N(Uuch) + " кВ (таблица 6.3), кВт·ч:");
            p.Math.Add("ΔW_{а ТН}=" + N(TN()));
            s.Add(p);
            sumA.Add("ΔW_{а ТН}");
            if (Chain) prevA += Sg + "ΔW_{а ТН}";
        }
        foreach (PotLine l in Near()) LineSteps(s, l, sumA, sumP, ref prevA, ref prevP);
        if (Tr != null) {
            p = new PotStep("Постоянные потери (потери холостого хода) в силовом трансформаторе " + Tr.Type + " (формулы (1), (2)), кВт·ч и квар·ч:");
            p.Math.Add("Δ\\bar{W}_{а Т}=T·ΔP_{хх}=T·" + N(Tr.DP));
            if (Reactive) p.Math.Add("Δ\\bar{W}_{р Т}=T·ΔQ_{хх}=T·" + N(Tr.DQ));
            s.Add(p);
            p = new PotStep("Переменные (нагрузочные) потери в силовом трансформаторе " + Tr.Type + " (формулы (8), (9)), кВт·ч и квар·ч:");
            Through(p, prevA, prevP);
            string f = "\\frac{" + Wsum("а", prevA) + "^{2}+" + Wsum("р", prevP) + "^{2}}";
            Two(p, "Δ\\tilde{W}_{а Т}", f + "{T·U_{нт}^{2}}·R_{т}·d·10^{-3}", f + "{T·" + N(Tr.Unt) + "^{2}}·" + N(Tr.R) + "·" + N(D) + "·10^{-3}");
            if (Reactive) Two(p, "Δ\\tilde{W}_{р Т}", f + "{T·U_{нт}^{2}}·X_{т}·d·10^{-3}", f + "{T·" + N(Tr.Unt) + "^{2}}·" + N(Tr.X) + "·" + N(D) + "·10^{-3}");
            s.Add(p);
            sumA.Add("Δ\\bar{W}_{а Т}"); sumA.Add("Δ\\tilde{W}_{а Т}");
            if (Reactive) { sumP.Add("Δ\\bar{W}_{р Т}"); sumP.Add("Δ\\tilde{W}_{р Т}"); }
            if (Chain) { prevA += Sg + "Δ\\bar{W}_{а Т}" + Sg + "Δ\\tilde{W}_{а Т}"; if (Reactive) prevP += Sg + "Δ\\bar{W}_{р Т}" + Sg + "Δ\\tilde{W}_{р Т}"; }
        }
        foreach (PotLine l in Far()) LineSteps(s, l, sumA, sumP, ref prevA, ref prevP);
        p = new PotStep("Суммарная величина технологического расхода электрической энергии (п. 6.5), кВт·ч" + (Reactive && sumP.Count > 0 ? " и квар·ч" : "") + ":");
        p.Math.Add("ΔW_{а Σ}=" + (sumA.Count == 0 ? "0" : string.Join("+", sumA.ToArray())));
        if (Reactive && sumP.Count > 0) p.Math.Add("ΔW_{р Σ}=" + string.Join("+", sumP.ToArray()));
        s.Add(p);
        if (!Popr) return s;
        p = new PotStep("Корректирующая поправка и данные учета с учетом коррекции (раздел 7; участок между точкой учета и ГБП - на балансе " + (Abonent ? "абонента: поправка прибавляется" : "РУП «МИНСКЭНЕРГО»: поправка вычитается") + "), кВт·ч:");
        p.Math.Add("δW_{а}=ΔW_{а Σ};   W'_{а}=W_{а}" + Sg + "δW_{а}");
        if (Reactive && sumP.Count > 0) p.Math.Add("δW_{р}=ΔW_{р Σ};   W'_{р}=W_{р}" + Sg + "δW_{р}");
        s.Add(p);
        return s;
    }

    static string Wsum(string c, string prev) { return prev.Length == 0 ? "W_{" + c + "}" : "W_{" + c + " вх}"; }

    // энергия на входе элемента с учётом потерь предыдущих (п. 4.12) - отдельной строкой перед формулой элемента
    static void Through(PotStep p, string prevA, string prevP) {
        if (prevA.Length > 0) p.Math.Add("W_{а вх}=W_{а}" + prevA);
        if (prevP.Length > 0) p.Math.Add("W_{р вх}=W_{р}" + prevP);
    }

    // «A=формула=числа» -> две строки: «A=формула» и «A=числа»
    static void Two(PotStep p, string lhs, string sym, string num) { p.Math.Add(lhs + "=" + sym); p.Math.Add(lhs + "=" + num); }

    void LineSteps(List<PotStep> s, PotLine l, List<string> sumA, List<string> sumP, ref string prevA, ref string prevP) {
        string n = l.Name;
        PotStep p;
        if (!l.Low && l.Kind == "КЛ" && l.Q30 > 0) {
            p = new PotStep("Постоянные потери в кабельной линии " + n + " " + Uv(l.U) + " кВ (формулы (3)-(5), q30 = " + N(l.Q30) + " квар/км, tgδ = " + N(l.TgD) + " - таблица 6.1), кВт·ч" + (Reactive ? " и квар·ч" : "") + ":");
            p.Math.Add("Q_{з " + n + "}=q_{30}·L=" + N(l.Q30) + "·" + N(l.L) + "=" + N(l.Q30 * l.L) + " \\text{квар}");
            p.Math.Add("Δ\\bar{W}_{а " + n + "}=Q_{з}·tgδ·T=" + N(l.Q30 * l.L) + "·" + N(l.TgD) + "·T=" + N(l.Q30 * l.L * l.TgD) + "·T");
            if (Reactive) p.Math.Add("Δ\\bar{W}_{р " + n + "}=-Q_{з}·T=-" + N(l.Q30 * l.L) + "·T");
            s.Add(p);
            sumA.Add("Δ\\bar{W}_{а " + n + "}");
            if (Reactive) sumP.Add("Δ\\bar{W}_{р " + n + "}");
        }
        if (l.OnePhase) {
            p = new PotStep("Величина технологического расхода электрической энергии в однофазной линии " + n + " 0,22 кВ (формулы (12), (13): Rл = Rф + R0, cosφ = 0,995; реактивные потери равны нулю), кВт·ч:");
            Through(p, prevA, "");
            Two(p, "ΔW_{а " + n + "}", "\\frac{" + Wsum("а", prevA) + "^{2}}{cos^{2}φ·U_{н}^{2}·T}·R_{" + n + "}·d·10^{-3}", "\\frac{" + Wsum("а", prevA) + "^{2}}{0,995^{2}·" + N(l.U) + "^{2}·T}·" + N(l.R) + "·" + N(D) + "·10^{-3}");
        } else {
            p = new PotStep("Величина технологического расхода электрической энергии на участке " + (l.Kind == "ВЛ" ? "воздушной" : "кабельной") + " линии " + n + " " + N(l.U) + " кВ (формула (10)" + (l.Low ? "; реактивные потери в линиях до 1 кВ равны нулю" : "") + "), кВт·ч:");
            Through(p, prevA, prevP);
            string f = "\\frac{" + Wsum("а", prevA) + "^{2}+" + Wsum("р", prevP) + "^{2}}";
            Two(p, "ΔW_{а " + n + "}", f + "{T·U_{н}^{2}}·R_{" + n + "}·d·10^{-3}", f + "{T·" + N(l.U) + "^{2}}·" + N(l.R) + "·" + N(D) + "·10^{-3}");
            if (!l.Low && Reactive && l.X > 0)
                Two(p, "ΔW_{р " + n + "}", f + "{T·U_{н}^{2}}·X_{" + n + "}·d·10^{-3}", f + "{T·" + N(l.U) + "^{2}}·" + N(l.X) + "·" + N(D) + "·10^{-3}");
        }
        s.Add(p);
        sumA.Add("ΔW_{а " + n + "}");
        if (!l.Low && Reactive && l.X > 0) sumP.Add("ΔW_{р " + n + "}");
        if (Chain) { prevA += Sg + "ΔW_{а " + n + "}"; if (!l.Low && Reactive && l.X > 0) prevP += Sg + "ΔW_{р " + n + "}"; }
    }

    // ---- проверочный расчёт за месяц (числа): итог активной составляющей, кВт·ч ----
    public double Calc(double wa, double wp, double t, List<string> log) {
        if (!WpMeter) wp = wa * Tg;
        double sg = Abonent ? 1 : -1, ca = 0, cp = 0, sum = 0;
        double tt = TT(), tn = TN();
        if (tt > 0) { sum += tt; if (Chain) ca += sg * tt; if (log != null) log.Add("ТТ " + N(tt)); }
        if (tn > 0) { sum += tn; if (Chain) ca += sg * tn; if (log != null) log.Add("ТН " + N(tn)); }
        foreach (PotLine l in Near()) sum += CalcLine(l, wa, wp, t, ref ca, ref cp, sg, log);
        if (Tr != null) {
            double c = t * Tr.DP, v = (Sq(wa + ca) + Sq(wp + cp)) / (t * Tr.Unt * Tr.Unt) * Tr.R * D * 1e-3;
            sum += c + v;
            if (log != null) log.Add(Tr.Type + " " + N(c) + " + " + N(v));
            if (Chain) { ca += sg * (c + v); if (Reactive) cp += sg * (t * Tr.DQ + (Sq(wa + ca) + Sq(wp + cp)) / (t * Tr.Unt * Tr.Unt) * Tr.X * D * 1e-3); }
        }
        foreach (PotLine l in Far()) sum += CalcLine(l, wa, wp, t, ref ca, ref cp, sg, log);
        return sum;
    }

    static double Sq(double x) { return x * x; }

    double CalcLine(PotLine l, double wa, double wp, double t, ref double ca, ref double cp, double sg, List<string> log) {
        double c = !l.Low && l.Kind == "КЛ" && l.Q30 > 0 ? l.Q30 * l.L * l.TgD * t : 0;
        // Uн - как у линии (0,4 / 0,38; 0,22 / 0,23): примечание к 3.1.10 допускает 0,4 и 0,23 кВ; пример А.3 книги считает с 0,4
        double v = l.OnePhase ? Sq(wa + ca) / (0.995 * 0.995 * l.U * l.U * t) * l.R * D * 1e-3
            : (Sq(wa + ca) + Sq(wp + cp)) / (t * Sq(l.U)) * l.R * D * 1e-3;
        if (log != null) log.Add(l.Name + " " + N(c + v));
        if (Chain) { ca += sg * v; if (!l.Low && Reactive && l.X > 0) cp += sg * (Sq(wa + ca) + Sq(wp + cp)) / (t * Sq(l.U)) * l.X * D * 1e-3; }
        return c + v;
    }

    // ---- описание схемы одной строкой (для формы) ----
    public string Summary() {
        List<string> e = new List<string>();
        e.Add("учёт " + (Uuch < 0.3 ? "0,23" : Uuch < 1 ? "0,4" : N(Uuch)) + " кВ " + (Uchet == 0 ? "прямого включения" : Uchet == 1 ? "через ТТ" : "через ТТ и ТН"));
        if (TT() > 0) e.Add("ТТ " + N(TT()) + " кВт·ч/мес");
        if (TN() > 0) e.Add("ТН " + N(TN()) + " кВт·ч/мес");
        foreach (PotLine l in Near()) e.Add(l.Name + " " + l.Kind + " " + N(l.U) + " кВ, R = " + N(l.R) + " Ом");
        if (Tr != null) e.Add("трансформатор " + Tr.Type);
        foreach (PotLine l in Far()) e.Add(l.Name + " " + l.Kind + " " + Uv(l.U) + " кВ, R = " + N(l.R) + " Ом");
        e.Add("d = " + N(D));
        e.Add(Abonent ? "потери прибавляются (баланс абонента)" : "потери вычитаются (баланс РУП)");
        return string.Join("; ", e.ToArray());
    }
}

// формулы в Word (OMML) из записи вида TeX и абзацы/таблицы WordprocessingML для вставки в документ
public static class Omml {
    const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    static string Esc(string s) { return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;"); }
    static string Run(string t, bool plain) {
        // числа с запятой («0,33») - обычным текстом: иначе Word ставит после запятой пробел, как в перечислении
        bool num = false;
        for (int i = 1; i + 1 < t.Length; i++) if (t[i] == ',' && char.IsDigit(t[i - 1]) && char.IsDigit(t[i + 1])) num = true;
        return "<m:r>" + (num ? "<m:rPr><m:nor/></m:rPr>" : plain ? "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" : "") + "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:sz w:val=\"24\"/></w:rPr><m:t xml:space=\"preserve\">" + Esc(t) + "</m:t></m:r>";
    }

    // запись -> OMML: \frac{a}{b}, x_{..}, x^{..}, \bar{W}, \tilde{W}, \text{..}; остальное - текст
    public static string Math(string tex) {
        int i = 0;
        return "<m:oMath>" + Seq(tex, ref i) + "</m:oMath>";
    }

    static string Seq(string s, ref int i) {
        StringBuilder sb = new StringBuilder(), plain = new StringBuilder();
        while (i < s.Length && s[i] != '}') {
            string atom = null;
            if (s[i] == '\\') {
                int j = i + 1;
                while (j < s.Length && char.IsLetter(s[j])) j++;
                string cmd = s.Substring(i + 1, j - i - 1);
                i = j;
                if (cmd == "frac") { string a = Arg(s, ref i), b = Arg(s, ref i); atom = "<m:f><m:num>" + a + "</m:num><m:den>" + b + "</m:den></m:f>"; }
                else if (cmd == "bar" || cmd == "tilde") {
                    string a = Arg(s, ref i);
                    atom = "<m:acc><m:accPr><m:chr m:val=\"" + (cmd == "bar" ? "̅" : "̃") + "\"/></m:accPr><m:e>" + a + "</m:e></m:acc>";
                } else if (cmd == "text") { int k = s.IndexOf('}', i); string t = s.Substring(i + 1, k - i - 1); i = k + 1; atom = Run(t, true); }
                else atom = Run(cmd, false);
            } else if (s[i] == '{') { i++; atom = Seq(s, ref i); i++; }
            else if (s[i] == '_' || s[i] == '^') { atom = null; }
            else {
                // простой текст: копим до значка индекса или команды
                char ch = s[i];
                bool next = i + 1 < s.Length && (s[i + 1] == '_' || s[i + 1] == '^');
                if (!next) { plain.Append(ch); i++; continue; }
                atom = Run(ch.ToString(), !char.IsLetter(ch) || ch > 'z');
                i++;
            }
            if (plain.Length > 0) { sb.Append(Run(plain.ToString(), Cyr(plain.ToString()))); plain.Length = 0; }
            if (atom == null) { atom = ""; }
            string sub = null, sup = null;
            while (i < s.Length && (s[i] == '_' || s[i] == '^')) {
                bool isSub = s[i] == '_';
                i++;
                string a = Arg(s, ref i);
                if (isSub) sub = a; else sup = a;
            }
            if (sub != null && sup != null) sb.Append("<m:sSubSup><m:e>" + atom + "</m:e><m:sub>" + sub + "</m:sub><m:sup>" + sup + "</m:sup></m:sSubSup>");
            else if (sub != null) sb.Append("<m:sSub><m:e>" + atom + "</m:e><m:sub>" + sub + "</m:sub></m:sSub>");
            else if (sup != null) sb.Append("<m:sSup><m:e>" + atom + "</m:e><m:sup>" + sup + "</m:sup></m:sSup>");
            else sb.Append(atom);
        }
        if (plain.Length > 0) sb.Append(Run(plain.ToString(), Cyr(plain.ToString())));
        return sb.ToString();
    }

    // русские буквы (индексы «а», «р», «ТТ», «хх») - прямым шрифтом
    static bool Cyr(string t) { foreach (char c in t) if (c >= '\u0400' && c <= '\u04FF') return true; return false; }

    static string Arg(string s, ref int i) {
        if (i < s.Length && s[i] == '{') {
            i++;
            string r = Seq(s, ref i);
            i++;
            // индекс из русских букв - прямым шрифтом
            return r;
        }
        if (i < s.Length) { string r = Run(s[i].ToString(), false); i++; return r; }
        return "";
    }

    // ---- WordprocessingML ----
    public static string P(string text, bool bold, bool indent) {
        return "<w:p><w:pPr><w:spacing w:before=\"60\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>" + (indent ? "<w:ind w:firstLine=\"567\"/>" : "") +
            "<w:jc w:val=\"both\"/></w:pPr><w:r><w:rPr>" + (bold ? "<w:b/>" : "") + "<w:sz w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">" + Esc(text) + "</w:t></w:r></w:p>";
    }

    public static string MathP(string tex) {
        return "<w:p><w:pPr><w:spacing w:before=\"60\" w:after=\"60\"/><w:jc w:val=\"center\"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr>" + Math(tex) + "</m:oMathPara></w:p>";
    }

    // таблица по ширине текста (twips), первая строка - заголовок
    public static string Table(string[] head, List<string[]> rows, int[] widths) { return Table(head, rows, widths, 20); }

    public static string Table(string[] head, List<string[]> rows, int[] widths, int size) {
        StringBuilder sb = new StringBuilder();
        int total = 0;
        foreach (int w in widths) total += w;
        sb.Append("<w:tbl><w:tblPr><w:tblW w:w=\"" + total + "\" w:type=\"dxa\"/><w:jc w:val=\"center\"/><w:tblBorders>");
        foreach (string b in new string[] { "top", "left", "bottom", "right", "insideH", "insideV" }) sb.Append("<w:" + b + " w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        sb.Append("</w:tblBorders><w:tblLayout w:type=\"fixed\"/><w:tblCellMar><w:left w:w=\"57\" w:type=\"dxa\"/><w:right w:w=\"57\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGrid>");
        foreach (int w in widths) sb.Append("<w:gridCol w:w=\"" + w + "\"/>");
        sb.Append("</w:tblGrid>");
        List<string[]> all = new List<string[]>();
        all.Add(head);
        all.AddRange(rows);
        for (int r = 0; r < all.Count; r++) {
            sb.Append("<w:tr><w:trPr><w:cantSplit/></w:trPr>");
            for (int c = 0; c < widths.Length; c++) {
                string t = c < all[r].Length ? all[r][c] : "";
                sb.Append("<w:tc><w:tcPr><w:tcW w:w=\"" + widths[c] + "\" w:type=\"dxa\"/><w:vAlign w:val=\"center\"/></w:tcPr><w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr>" +
                    (r == 0 ? "<w:b/>" : "") + "<w:sz w:val=\"" + size + "\"/></w:rPr><w:t xml:space=\"preserve\">" + Esc(t) + "</w:t></w:r></w:p></w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
        return sb.ToString();
    }

    // абзац с меткой mark заменяется фрагментом xml (абзацы, таблицы); ns - пространства имён w и m
    public static void Insert(string docx, string mark, string xml) {
        List<ZipEntry> entries = Zip.Read(docx);
        foreach (ZipEntry e in entries) {
            if (e.Name != "word/document.xml") continue;
            string doc = Encoding.UTF8.GetString(e.GetData());
            int at = doc.IndexOf(mark);
            if (at < 0) return;
            int ps = doc.LastIndexOf("<w:p ", at), ps2 = doc.LastIndexOf("<w:p>", at);
            ps = Math2(ps, ps2);
            int pe = doc.IndexOf("</w:p>", at) + 6;
            string root = doc.Substring(0, doc.IndexOf('>', doc.IndexOf("<w:document")) + 1);
            if (root.IndexOf("xmlns:m=") < 0) doc = doc.Replace(root, root.Replace("<w:document ", "<w:document xmlns:m=\"" + M + "\" "));
            at = doc.IndexOf(mark);
            ps = Math2(doc.LastIndexOf("<w:p ", at), doc.LastIndexOf("<w:p>", at));
            pe = doc.IndexOf("</w:p>", at) + 6;
            doc = doc.Substring(0, ps) + xml + doc.Substring(pe);
            e.SetData(new UTF8Encoding(false).GetBytes(doc));
        }
        Zip.Write(docx, entries);
    }

    static int Math2(int a, int b) { return a > b ? a : b; }
}
// ==== КОНЕЦ БЛОКА РАСЧЁТА ПОТЕРЬ ====
