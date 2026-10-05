package by.mobilemeter.protocol.iec62056;

import by.mobilemeter.transport.SerialLink;
import by.mobilemeter.util.Hex;
import by.mobilemeter.util.LogSink;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * IEC 62056-21 (ГОСТ IEC 61107-2011) session: identification, data readout,
 * programming mode with password, R1/W1 commands. Pure Java, no Android types,
 * so it can be unit-tested on a PC and reused over any SerialLink.
 */
public final class Iec62056Session {
    public static final byte SOH = 0x01;
    public static final byte STX = 0x02;
    public static final byte ETX = 0x03;
    public static final byte EOT = 0x04;
    public static final byte ACK = 0x06;
    public static final byte NAK = 0x15;

    private final SerialLink link;
    private final LogSink log;

    private int initialBaud = 300;
    private boolean switchBaud = true;
    private int dataBits = 7;
    private int parity = SerialLink.PARITY_EVEN;
    private int responseTimeoutMs = 3000;
    private int interCharTimeoutMs = 1500;

    private int currentBaud = -1;
    private Identification lastId;

    public Iec62056Session(SerialLink link, LogSink log) {
        this.link = link;
        this.log = log;
    }

    public void configure(int initialBaud, boolean switchBaud, boolean sevenE1) {
        this.initialBaud = initialBaud;
        this.switchBaud = switchBaud;
        this.dataBits = sevenE1 ? 7 : 8;
        this.parity = sevenE1 ? SerialLink.PARITY_EVEN : SerialLink.PARITY_NONE;
        this.currentBaud = -1;
    }

    public Identification lastIdentification() {
        return lastId;
    }

    // ---- public operations -------------------------------------------------

    /** Sends "/?address!" and parses the identification message. */
    public Identification identify(String address) throws IOException {
        applyBaud(initialBaud);
        drain();
        String addr = address == null ? "" : address.trim();
        send(Hex.ascii("/?" + addr + "!\r\n"));
        byte[] line = readUntil((byte) '\n', responseTimeoutMs);
        if (line.length == 0) {
            throw new IOException("Счётчик не ответил на запрос идентификации (" + currentBaud
                    + " бод, " + formatName() + "). Проверьте положение оптоголовки, начальную скорость и формат.");
        }
        String text = Hex.ascii(line, 0, line.length).trim();
        Identification id = Identification.parse(text);
        lastId = id;
        log.line("Идентификация: " + id.summary());
        return id;
    }

    /** Identification followed by the standard data readout (all registers). */
    public List<Register> readout(String address) throws IOException {
        Identification id = identify(address);
        if (id.modeE) {
            log.line("Счётчик объявляет режим E (DLMS/COSEM по HDLC). Пробуем обычное считывание, "
                    + "многие счётчики его всё равно поддерживают.");
        }
        acknowledge(id, '0');
        byte[] frame = readFrame(60000);
        String body = body(frame, STX);
        List<Register> regs = ReadoutParser.parse(body);
        log.line("Считано строк: " + regs.size());
        for (Register r : regs) {
            log.line("  " + r);
        }
        return regs;
    }

    /** Identification, then programming mode with the P1 password. */
    public void enterProgrammingMode(String address, String password) throws IOException {
        Identification id = identify(address);
        acknowledge(id, '1');
        byte[] frame = readFrame(responseTimeoutMs);
        if (frame.length == 0) {
            throw new IOException("Нет ответа на запрос режима программирования");
        }
        if (isSingle(frame, ACK)) {
            log.line("Счётчик сразу ответил ACK: пароль не требуется");
            return;
        }
        String seed = body(frame, SOH);
        log.line("Запрос пароля от счётчика: " + seed);
        sendFrame(SOH, "P1", "(" + (password == null ? "" : password.trim()) + ")");
        expectAck("на пароль");
        log.line("Пароль принят, режим программирования открыт");
    }

    /** R1 read of one parameter, e.g. "ET0PE" or "1.8.0". Returns the body "name(value)". */
    public String readRegister(String name) throws IOException {
        String n = name.trim();
        if (!n.endsWith(")")) {
            n = n + "()";
        }
        sendFrame(SOH, "R1", n);
        byte[] frame = readFrame(responseTimeoutMs);
        if (frame.length == 0) {
            throw new IOException("Нет ответа на R1 " + n);
        }
        if (isSingle(frame, NAK)) {
            throw new IOException("NAK на R1 " + n + " (команда отвергнута)");
        }
        String b = body(frame, STX);
        log.line("Ответ: " + b);
        return b;
    }

    /** W1 write of one parameter: name(value). */
    public String writeRegister(String name, String value) throws IOException {
        sendFrame(SOH, "W1", name.trim() + "(" + value + ")");
        byte[] frame = readFrame(responseTimeoutMs);
        if (frame.length == 0) {
            throw new IOException("Нет ответа на W1 " + name);
        }
        if (isSingle(frame, ACK)) {
            log.line("W1 " + name + ": ACK (принято)");
            return "ACK";
        }
        if (isSingle(frame, NAK)) {
            throw new IOException("NAK на W1 " + name + " (команда отвергнута)");
        }
        String b = body(frame, STX);
        log.line("Ответ на W1: " + b);
        return b;
    }

    /** Sends text verbatim plus CR LF and shows whatever arrives within waitMs. */
    public String raw(String text, int waitMs) throws IOException {
        if (currentBaud < 0) {
            applyBaud(initialBaud);
        }
        send(Hex.ascii(text + "\r\n"));
        byte[] reply = collect(waitMs);
        return Hex.printable(reply, 0, reply.length);
    }

    /** B0 sign-off: ends the programming session. */
    public void signOff() throws IOException {
        sendFrame(SOH, "B0", null);
        byte[] reply = collect(500);
        if (reply.length > 0) {
            log.line("Ответ на B0: " + Hex.printable(reply, 0, reply.length));
        }
        currentBaud = -1;
    }

    // ---- protocol helpers ---------------------------------------------------

    /** ACK + option select; switches the host baud rate when the meter proposes one. */
    private void acknowledge(Identification id, char mode) throws IOException {
        if (id.modeB) {
            // Mode B: the meter switches on its own right after the identification.
            if (switchBaud && id.baud > 0) {
                sleep(250);
                applyBaud(id.baud);
            }
            return;
        }
        boolean doSwitch = switchBaud && id.baud > 0 && id.baud != currentBaud;
        char z = doSwitch ? id.baudId : Identification.baudIdFor(currentBaud);
        byte[] ack = {ACK, (byte) '0', (byte) z, (byte) mode, (byte) '\r', (byte) '\n'};
        send(ack);
        // Let the acknowledgement leave the wire before changing the rate (6 chars, 10 bits each).
        sleep(300 + 6 * 10 * 1000 / Math.max(currentBaud, 300));
        if (doSwitch) {
            applyBaud(id.baud);
        }
    }

    private void expectAck(String what) throws IOException {
        byte[] reply = readFrame(responseTimeoutMs);
        if (reply.length == 0) {
            throw new IOException("Нет ответа " + what);
        }
        if (isSingle(reply, ACK)) {
            return;
        }
        if (isSingle(reply, NAK)) {
            throw new IOException("NAK " + what + ": отвергнуто (неверный пароль?)");
        }
        throw new IOException("Неожиданный ответ " + what + ": " + Hex.printable(reply, 0, reply.length));
    }

    private void sendFrame(byte start, String command, String data) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        bos.write(start);
        byte[] cmd = Hex.ascii(command);
        bos.write(cmd, 0, cmd.length);
        if (data != null) {
            bos.write(STX);
            byte[] d = Hex.ascii(data);
            bos.write(d, 0, d.length);
        }
        bos.write(ETX);
        byte[] frame = bos.toByteArray();
        bos.write(bcc(frame, 1, frame.length));
        send(bos.toByteArray());
    }

    /**
     * Extracts the text between the start byte (SOH or STX) and ETX, checking the BCC.
     * A BCC mismatch is logged, not thrown: some meters compute it differently.
     */
    private String body(byte[] frame, byte expectedStart) throws IOException {
        if (frame.length == 0) {
            throw new IOException("Пустой ответ");
        }
        if (isSingle(frame, NAK)) {
            throw new IOException("Счётчик ответил NAK");
        }
        if (isSingle(frame, EOT)) {
            throw new IOException("Счётчик завершил сеанс (EOT)");
        }
        int start = -1;
        for (int i = 0; i < frame.length; i++) {
            byte b = (byte) (frame[i] & 0x7F);
            if (b == SOH || b == STX) {
                start = i;
                break;
            }
        }
        if (start < 0) {
            // Plain text without framing: return as is.
            return Hex.ascii(frame, 0, frame.length);
        }
        int etx = -1;
        for (int i = start + 1; i < frame.length; i++) {
            if ((frame[i] & 0x7F) == ETX) {
                etx = i;
                break;
            }
        }
        if (etx < 0) {
            log.line("Предупреждение: ETX не получен, ответ неполный");
            return Hex.ascii(frame, start + 1, frame.length - start - 1);
        }
        if (etx + 1 < frame.length) {
            int expected = bcc(frame, start + 1, etx + 1) & 0x7F;
            int actual = frame[etx + 1] & 0x7F;
            if (expected != actual) {
                log.line(String.format("Предупреждение: BCC не совпал (ожидали %02X, получили %02X)", expected, actual));
            }
        } else {
            log.line("Предупреждение: BCC не получен");
        }
        // Skip the STX that follows a SOH command echo, if any.
        int bodyStart = start + 1;
        StringBuilder sb = new StringBuilder();
        for (int i = bodyStart; i < etx; i++) {
            int c = frame[i] & 0x7F;
            if (c == STX && expectedStart == SOH) {
                continue;
            }
            sb.append((char) c);
        }
        return sb.toString();
    }

    /** XOR of bytes [from, to) — the IEC 62056-21 block check character. */
    public static byte bcc(byte[] data, int from, int to) {
        int x = 0;
        for (int i = from; i < to; i++) {
            x ^= (data[i] & 0x7F);
        }
        return (byte) x;
    }

    private static boolean isSingle(byte[] frame, byte b) {
        return frame.length == 1 && (frame[0] & 0x7F) == b;
    }

    // ---- I/O helpers --------------------------------------------------------

    private void applyBaud(int baud) throws IOException {
        if (baud == currentBaud) {
            return;
        }
        link.setParameters(baud, dataBits, 1, parity);
        currentBaud = baud;
        log.line("Порт: " + baud + " бод, " + formatName());
    }

    private String formatName() {
        return dataBits + (parity == SerialLink.PARITY_EVEN ? "E" : parity == SerialLink.PARITY_ODD ? "O" : "N") + "1";
    }

    private void send(byte[] data) throws IOException {
        log.tx(data);
        link.write(data);
    }

    private void drain() throws IOException {
        link.purge();
        byte[] buf = new byte[256];
        int n;
        while ((n = link.read(buf, 50)) > 0) {
            log.rx(buf, 0, n);
        }
    }

    /** Reads until the terminator or until the timeouts expire. */
    private byte[] readUntil(byte terminator, int firstByteTimeout) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[256];
        long deadline = System.currentTimeMillis() + firstByteTimeout;
        while (System.currentTimeMillis() < deadline) {
            int n = link.read(buf, bos.size() == 0 ? firstByteTimeout : interCharTimeoutMs);
            if (n <= 0) {
                if (bos.size() > 0) {
                    break;
                }
                continue;
            }
            log.rx(buf, 0, n);
            bos.write(buf, 0, n);
            if (indexOf(bos.toByteArray(), terminator) >= 0) {
                break;
            }
            deadline = System.currentTimeMillis() + interCharTimeoutMs;
        }
        return bos.toByteArray();
    }

    /**
     * Reads one frame: a single ACK/NAK/EOT, or bytes up to ETX plus the BCC byte.
     * Stops early when nothing arrives for interCharTimeoutMs.
     */
    private byte[] readFrame(int totalTimeoutMs) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[1024];
        long deadline = System.currentTimeMillis() + totalTimeoutMs;
        long quietDeadline = deadline;
        while (System.currentTimeMillis() < Math.min(deadline, quietDeadline)) {
            int wait = (int) Math.max(50, Math.min(interCharTimeoutMs, deadline - System.currentTimeMillis()));
            int n = link.read(buf, wait);
            if (n <= 0) {
                if (bos.size() > 0) {
                    break;
                }
                continue;
            }
            log.rx(buf, 0, n);
            bos.write(buf, 0, n);
            byte[] cur = bos.toByteArray();
            if (cur.length == 1) {
                int c = cur[0] & 0x7F;
                if (c == ACK || c == NAK || c == EOT) {
                    // Give a possible continuation a moment, then accept the single byte.
                    int more = link.read(buf, 100);
                    if (more <= 0) {
                        return cur;
                    }
                    log.rx(buf, 0, more);
                    bos.write(buf, 0, more);
                    cur = bos.toByteArray();
                }
            }
            int etx = indexOf(cur, ETX);
            if (etx >= 0 && cur.length >= etx + 2) {
                return cur;
            }
            quietDeadline = System.currentTimeMillis() + interCharTimeoutMs;
        }
        return bos.toByteArray();
    }

    private byte[] collect(int waitMs) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[1024];
        long deadline = System.currentTimeMillis() + waitMs;
        while (System.currentTimeMillis() < deadline) {
            int n = link.read(buf, (int) Math.max(50, deadline - System.currentTimeMillis()));
            if (n > 0) {
                log.rx(buf, 0, n);
                bos.write(buf, 0, n);
            }
        }
        return bos.toByteArray();
    }

    private static int indexOf(byte[] data, byte b) {
        for (int i = 0; i < data.length; i++) {
            if ((data[i] & 0x7F) == b) {
                return i;
            }
        }
        return -1;
    }

    private static void sleep(long ms) {
        try {
            Thread.sleep(ms);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }
}
