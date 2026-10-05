package by.mobilemeter.protocol.mirtek;

import by.mobilemeter.transport.SerialLink;
import by.mobilemeter.util.Hex;
import by.mobilemeter.util.LogSink;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Locale;

/**
 * МИРТЕК protocol session over a SerialLink (optical port: 9600 8N1).
 * Pure Java: no Android types, unit-testable on a PC.
 */
public final class MirtekSession {
    public static final int CMD_PING = 0x01;
    public static final int CMD_READ_ENERGY = 0x05;
    public static final int CMD_READ_FACTORY_STRING = 0x0A;
    public static final int CMD_READ_CONFIGURE = 0x10;
    public static final int CMD_READ_DATETIME = 0x1C;
    public static final int CMD_READ_INSTANT = 0x2B;
    public static final int CMD_GET_INFO = 0x30;
    public static final int CMD_RELAY = 0x3A;

    private static final String[] DOW = {"Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"};

    private final SerialLink link;
    private final LogSink log;
    private int address;
    private long password;
    private int timeoutMs = 3000;
    private boolean portReady;

    public MirtekSession(SerialLink link, LogSink log) {
        this.link = link;
        this.log = log;
    }

    public void configure(int address, long password) {
        this.address = address & 0xFFFF;
        this.password = password & 0xFFFFFFFFL;
    }

    public void setTimeout(int ms) {
        timeoutMs = ms;
    }

    public void resetPort() {
        portReady = false;
    }

    // ---- commands -----------------------------------------------------------

    /** 0x01 Ping: firmware version and the meter's own address. */
    public String ping() throws IOException {
        MirtekFrame f = exchange(CMD_PING, new byte[0]);
        String text;
        if (f.data.length >= 4) {
            text = String.format(Locale.US, "версия ПО %d.%d, адрес %d",
                    f.data[1] & 0xFF, f.data[0] & 0xFF, MirtekFrame.u16(f.data, 2));
        } else {
            text = "данные: " + Hex.of(f.data);
        }
        log.line("Ping: " + text);
        return text;
    }

    /** 0x05 Read Status Counter, active forward energy (type 0). */
    public EnergyReading readEnergy() throws IOException {
        MirtekFrame f;
        try {
            f = exchange(CMD_READ_ENERGY, new byte[]{0x00});
        } catch (MeterError e) {
            if (e.code != 0x04) {
                throw e;
            }
            log.line("Счётчик старого поколения: повторяю 0x05 без параметра");
            f = exchange(CMD_READ_ENERGY, new byte[0]);
        }
        EnergyReading r = EnergyReading.parse(f.data);
        log.line("Показания: " + r);
        return r;
    }

    /** 0x1C ReadDateTime. */
    public String readDateTime() throws IOException {
        MirtekFrame f = exchange(CMD_READ_DATETIME, new byte[0]);
        if (f.data.length < 7) {
            throw new IOException("Короткий ответ на 0x1C: " + f.data.length + " байт");
        }
        int dow = f.data[3] & 0xFF;
        String text = String.format(Locale.US, "%02d.%02d.20%02d %02d:%02d:%02d (%s)",
                f.data[4] & 0xFF, f.data[5] & 0xFF, f.data[6] & 0xFF,
                f.data[2] & 0xFF, f.data[1] & 0xFF, f.data[0] & 0xFF,
                dow < DOW.length ? DOW[dow] : "?");
        log.line("Время счётчика: " + text);
        return text;
    }

    /** 0x0A ReadFactoryString: 1 serial, 2 production date, 3 plant, 4 and 5 device name. */
    public String readFactoryString(int field) throws IOException {
        MirtekFrame f = exchange(CMD_READ_FACTORY_STRING, new byte[]{(byte) field});
        String s = f.data.length > 1 ? ascii(f.data, 1, f.data.length - 1) : "";
        log.line("Заводское поле " + field + ": " + s);
        return s;
    }

    /** 0x30 GetInfo: board id, firmware version, uptime. */
    public String getInfo() throws IOException {
        MirtekFrame f = exchange(CMD_GET_INFO, new byte[0]);
        byte[] d = f.data;
        StringBuilder sb = new StringBuilder();
        if (d.length >= 3) {
            sb.append(String.format(Locale.US, "ID платы 0x%02X, версия прошивки %d.%d", d[0] & 0xFF, d[2] & 0xFF, d[1] & 0xFF));
        }
        if (d.length >= 9) {
            sb.append(", наработка ").append(MirtekFrame.u32(d, 5) / 3600).append(" ч");
        }
        if (d.length >= 15) {
            sb.append(", 100А: ").append((d[14] & 0x80) != 0 ? "да" : "нет");
        }
        log.line("GetInfo: " + sb);
        return sb.toString();
    }

    /** 0x10 ReadConfigure: role and relay bits (new generation). */
    public String readConfigure() throws IOException {
        MirtekFrame f = exchange(CMD_READ_CONFIGURE, new byte[0]);
        byte[] d = f.data;
        StringBuilder sb = new StringBuilder(d.length + " байт");
        if (d.length >= 19) {
            sb.append(String.format(Locale.US, "; роль 0x%02X; реле в наличии: %s; статус реле (1=разомкнуто): %s",
                    d[10] & 0xFF, bits(d[11] & 0xFF), bits(d[13] & 0xFF)));
        } else if (d.length >= 7) {
            sb.append(String.format(Locale.US, "; роль 0x%02X", d[d.length - 1] & 0xFF));
        }
        log.line("Конфигурация: " + sb);
        return sb.toString();
    }

    /**
     * 0x3A OperateOverload: relay 0 is the load disconnect relay.
     * disconnect=true opens the relay (load off), false closes it (load on). Needs the password.
     */
    public MirtekFrame relay(int relayNo, boolean disconnect) throws IOException {
        log.line((disconnect ? "ОТКЛЮЧЕНИЕ" : "ВКЛЮЧЕНИЕ") + " нагрузки, реле " + relayNo + ", адрес " + address);
        MirtekFrame f = exchange(CMD_RELAY, new byte[]{(byte) relayNo, (byte) (disconnect ? 1 : 0)});
        log.line("Реле: команда принята. Состояние по статусу: нагрузка "
                + (MirtekStatus.loadDisconnected(f.statusOrPassword) ? "ОТКЛЮЧЕНА" : "включена"));
        return f;
    }

    /** Any command with raw data bytes; the reply is logged in full. */
    public MirtekFrame raw(int command, byte[] data) throws IOException {
        MirtekFrame f = exchange(command, data);
        log.line(String.format(Locale.US, "Ответ на 0x%02X: %d байт: %s", command, f.data.length, Hex.of(f.data)));
        return f;
    }

    // ---- exchange -----------------------------------------------------------

    public MirtekFrame exchange(int command, byte[] data) throws IOException {
        ensurePort();
        drain();
        byte[] req = MirtekFrame.buildRequest(address, command, password, data);
        log.tx(req);
        link.write(req);
        MirtekFrame f = readFrame(timeoutMs);
        if (f == null) {
            throw new IOException(String.format(Locale.US,
                    "Счётчик с адресом %d не ответил на команду 0x%02X за %d мс", address, command, timeoutMs));
        }
        if (f.request) {
            throw new IOException("Принят собственный запрос (эхо): проверьте подключение");
        }
        log.line("Статус: " + MirtekStatus.describe(f.statusOrPassword));
        if (f.command != command) {
            log.line(String.format(Locale.US, "Предупреждение: в ответе код команды 0x%02X вместо 0x%02X", f.command, command));
        }
        int err = f.errorCode();
        if (err != 0) {
            throw new MeterError(err, String.format(Locale.US, "Счётчик вернул ошибку 0x%02X: %s", err, MirtekStatus.errorText(err)));
        }
        return f;
    }

    /** Error reported by the meter in the status field. */
    public static final class MeterError extends IOException {
        public final int code;

        MeterError(int code, String message) {
            super(message);
            this.code = code;
        }
    }

    private void ensurePort() throws IOException {
        if (!portReady) {
            link.setParameters(9600, 8, 1, SerialLink.PARITY_NONE);
            log.line("Порт: 9600 бод, 8N1 (протокол МИРТЕК)");
            portReady = true;
        }
    }

    private void drain() throws IOException {
        link.purge();
        byte[] buf = new byte[256];
        int n;
        while ((n = link.read(buf, 30)) > 0) {
            log.rx(buf, 0, n);
        }
    }

    private MirtekFrame readFrame(int totalTimeoutMs) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[512];
        long deadline = System.currentTimeMillis() + totalTimeoutMs;
        while (System.currentTimeMillis() < deadline) {
            int wait = (int) Math.max(50, Math.min(500, deadline - System.currentTimeMillis()));
            int n = link.read(buf, wait);
            if (n <= 0) {
                continue;
            }
            log.rx(buf, 0, n);
            bos.write(buf, 0, n);
            MirtekFrame f = MirtekFrame.parse(bos.toByteArray());
            if (f != null) {
                return f;
            }
        }
        return null;
    }

    private static String ascii(byte[] d, int off, int len) {
        StringBuilder sb = new StringBuilder();
        for (int i = off; i < off + len; i++) {
            int c = d[i] & 0xFF;
            if (c == 0) {
                break;
            }
            sb.append(c >= 0x20 && c < 0x7F ? (char) c : '.');
        }
        return sb.toString().trim();
    }

    private static String bits(int v) {
        return String.format(Locale.US, "%5s", Integer.toBinaryString(v & 0x1F)).replace(' ', '0');
    }
}
