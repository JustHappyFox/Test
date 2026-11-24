// Проверка расчёта потерь (PotScheme.Calc) на примерах из книги ТКП 460-2017 - числа должны совпасть с книгой.
// mcs -r:dogovor_mono.dll t_poteri_kniga.cs; mono t_poteri_kniga.exe
using System; using System.Collections.Generic;
public static class T {
    static void Check(string name, double got, double want) {
        Console.WriteLine((Math.Abs(got - want) <= Math.Max(1, want * 0.002) ? "ОК   " : "ОШИБКА ") + name + ": " + Math.Round(got, 1) + " (в книге " + want + ")");
    }
    public static void Main() {
        // А.1: трансформатор 25 МВА 110/10, учёт 10 кВ через ТТ и ТН, Wa = 7 000 000, Wp = 6 160 000 (по счётчику), d = 1,7, T = 744
        PotScheme a1 = new PotScheme();
        a1.Uuch = 10; a1.Uchet = 2; a1.D = 1.7; a1.WpMeter = true; a1.Chain = false;
        a1.Tr = new TrafoInfo(); a1.Tr.Type = "ТРДН-25000/110"; a1.Tr.Unt = 110; a1.Tr.R = 2.54; a1.Tr.X = 55.5; a1.Tr.DP = 25; a1.Tr.DQ = 175;
        Check("А.1 активная составляющая ТРЭТ", a1.Calc(7000000, 6160000, 744, null), 60469);
        // А.3 без группы: одна ВЛ 0,38 кВ с Rэ = 0,017 Ом (эквивалент трёх ВЛ), учёт через ТТ, Wa = 35 000, tgφ = 0,88, d = 1,7
        PotScheme a3 = new PotScheme();
        a3.Uuch = 0.38; a3.Uchet = 1; a3.D = 1.7; a3.Tg = 0.88; a3.Chain = false;
        PotLine l = new PotLine(); l.Name = "Lэ"; l.U = 0.4; l.L = 1; l.R0 = 0.017; l.R = 0.017; l.Kind = "ВЛ";
        a3.Lines.Add(l);
        Check("А.3 активная составляющая ТРЭТ", a3.Calc(35000, 0, 744, null), 532);
        // образец пользователя (0,4 кВ, ТТ, АВБбШв 4х120 0,180 км, d = 1,70, tgφ = 0,33, ТТ к энергии кабеля) - для Wa = 35 000
        PotScheme u = new PotScheme();
        u.Uuch = 0.38; u.Uchet = 1; u.D = 1.7; u.Tg = 0.33; u.Chain = true;
        PotLine k = new PotLine(); k.Name = "L1"; k.U = 0.4; k.L = 0.18; k.R0 = 0.258; k.R = 0.04644;
        u.Lines.Add(k);
        double wa = 35000, wp = wa * 0.33, ruchn = (Math.Pow(wa + 4.2, 2) + wp * wp) / (744 * 0.4 * 0.4) * 0.04644 * 1.7e-3 + 4.2;
        Check("образец 0,4 кВ с ТТ (ручной пересчёт)", u.Calc(wa, 0, 744, null), Math.Round(ruchn, 1));
    }
}
