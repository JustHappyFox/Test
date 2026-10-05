package by.mobilemeter.protocol.mirtek;

import java.util.Locale;

/** Decoded answer of command 0x05 (Read Status Counter). */
public final class EnergyReading {
    public final boolean newGeneration;
    public final int decimals;
    public final int tariffsUsed;
    public final int activeTariff;
    public final long total;
    public final long totalByTariffs;
    public final long[] tariffs;
    public final int ku;
    public final int ki;

    private EnergyReading(boolean newGeneration, int config, long total, long totalByTariffs, long[] tariffs, int ku, int ki) {
        this.newGeneration = newGeneration;
        this.decimals = config & 0x03;
        this.activeTariff = ((config >> 2) & 0x03) + 1;
        this.tariffsUsed = ((config >> 6) & 0x03) + 1;
        this.total = total;
        this.totalByTariffs = totalByTariffs;
        this.tariffs = tariffs;
        this.ku = ku;
        this.ki = ki;
    }

    public static EnergyReading parse(byte[] d) {
        long[] t = new long[4];
        if (d.length >= 30) {
            // type(1) cfg(1) Ku(2) Ki(2) full(4) byTariffs(4) T1..T4(16)
            int cfg = d[1] & 0xFF;
            for (int i = 0; i < 4; i++) {
                t[i] = MirtekFrame.u32(d, 14 + 4 * i);
            }
            return new EnergyReading(true, cfg, MirtekFrame.u32(d, 6), MirtekFrame.u32(d, 10), t,
                    MirtekFrame.u16(d, 2), MirtekFrame.u16(d, 4));
        }
        if (d.length >= 26) {
            // old generation: sum(4) cfg(1) kd(1) role(1) kmul(3) T1..T4(16)
            int cfg = d[4] & 0xFF;
            for (int i = 0; i < 4; i++) {
                t[i] = MirtekFrame.u32(d, 10 + 4 * i);
            }
            long sum = MirtekFrame.u32(d, 0);
            return new EnergyReading(false, cfg, sum, sum, t, 1, 1);
        }
        throw new IllegalArgumentException("Неожиданная длина ответа 0x05: " + d.length);
    }

    public String format(long raw) {
        double v = raw / Math.pow(10, decimals);
        return String.format(Locale.US, "%." + decimals + "f", v);
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("Сумма A+ = ").append(format(total)).append(" кВт·ч");
        if (newGeneration && totalByTariffs != total) {
            sb.append(" (по задействованным тарифам ").append(format(totalByTariffs)).append(')');
        }
        for (int i = 0; i < tariffsUsed && i < 4; i++) {
            sb.append("; T").append(i + 1).append(" = ").append(format(tariffs[i]));
        }
        sb.append("; действующий тариф T").append(activeTariff);
        if (ku != 1 || ki != 1) {
            sb.append("; Kн=").append(ku).append(" Kт=").append(ki);
        }
        return sb.toString();
    }
}
