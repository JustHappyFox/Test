package by.mobilemeter.protocol.mirtek;

import by.mobilemeter.transport.SerialLink;
import by.mobilemeter.util.Hex;
import by.mobilemeter.util.LogSink;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/** Plain-JVM self test of the МИРТЕК framing and session against a scripted fake meter. */
public final class MirtekSelfTest {

    static void check(boolean cond, String msg) {
        if (!cond) {
            throw new AssertionError(msg);
        }
    }

    /** Table-driven CRC8 (poly 0xA9, MSB first) to cross-check the bitwise transcription. */
    static int crc8Table(byte[] d) {
        int[] table = new int[256];
        for (int i = 0; i < 256; i++) {
            int c = i;
            for (int k = 0; k < 8; k++) {
                c = (c & 0x80) != 0 ? ((c << 1) ^ 0xA9) & 0xFF : (c << 1) & 0xFF;
            }
            table[i] = c;
        }
        int crc = 0;
        for (byte b : d) {
            crc = table[(crc ^ (b & 0xFF)) & 0xFF];
        }
        return crc;
    }

    /** Builds a reply frame the way a meter would (D=0, status instead of password). */
    static byte[] reply(int address, int command, byte[] status, byte[] data) {
        ByteArrayOutputStream body = new ByteArrayOutputStream();
        body.write(data.length);
        body.write(0);
        body.write(address & 0xFF);
        body.write(address >>> 8);
        body.write(0xFF);
        body.write(0xFF);
        body.write(command);
        body.write(status, 0, 4);
        body.write(data, 0, data.length);
        byte[] b = body.toByteArray();
        int crc = MirtekFrame.crc8(b, 0, b.length);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        out.write(0x73);
        out.write(0x55);
        for (byte x : b) {
            stuff(out, x & 0xFF);
        }
        stuff(out, crc);
        out.write(0x55);
        return out.toByteArray();
    }

    static void stuff(ByteArrayOutputStream out, int b) {
        if (b == 0x55) { out.write(0x73); out.write(0x11); }
        else if (b == 0x73) { out.write(0x73); out.write(0x22); }
        else out.write(b);
    }

    static final class FakeMeter implements SerialLink {
        final int address;
        final long password;
        final ByteArrayOutputStream pending = new ByteArrayOutputStream();
        boolean relayOff;
        boolean oldGeneration;
        int lastCommand = -1;

        FakeMeter(int address, long password) {
            this.address = address;
            this.password = password;
        }

        public void open() { }
        public void setParameters(int baud, int dataBits, int stopBits, int parity) {
            check(baud == 9600 && dataBits == 8 && parity == SerialLink.PARITY_NONE, "port 9600 8N1");
        }
        public void purge() { }
        public boolean isOpen() { return true; }
        public String describe() { return "fake mirtek"; }
        public void close() { }

        public void write(byte[] raw) throws IOException {
            MirtekFrame f = MirtekFrame.parse(raw);
            check(f != null && f.request, "request frame parsed");
            check(f.source == 0xFFFF, "source address");
            if (f.destination != address) {
                return; // not for us: silence
            }
            long pw = MirtekFrame.u32(f.statusOrPassword, 0);
            lastCommand = f.command;
            byte[] status = {0x39, 0x00, (byte) (0x02 | (relayOff ? 0x08 : 0x00) | 0x01), 0x00};
            byte[] data = new byte[0];
            switch (f.command) {
                case 0x01:
                    data = new byte[]{0x05, 0x01, (byte) address, (byte) (address >>> 8)};
                    break;
                case 0x05:
                    if (oldGeneration) {
                        if (f.data.length != 0) { status[3] = 0x04; break; }
                        data = new byte[26];
                        put32(data, 0, 123456); data[4] = 0x02; // 2 decimals, 1 tariff
                        put32(data, 10, 123456);
                    } else {
                        if (f.data.length != 1) { status[3] = 0x04; break; }
                        data = new byte[30];
                        data[0] = 0; data[1] = (byte) (0x03 | 0x40); // 3 decimals, 2 tariffs
                        data[2] = 1; data[4] = 1;
                        put32(data, 6, 2576123); put32(data, 10, 2576123);
                        put32(data, 14, 2000000); put32(data, 18, 576123);
                    }
                    break;
                case 0x1C:
                    data = new byte[]{30, 15, 14, 1, 5, 10, 26};
                    break;
                case 0x3A:
                    if (pw != password) { status[3] = 0x01; break; }
                    if (f.data.length != 2) { status[3] = 0x04; break; }
                    relayOff = f.data[1] == 1;
                    status[2] = (byte) (0x02 | (relayOff ? 0x08 : 0x00) | 0x01);
                    break;
                default:
                    status[3] = 0x08;
            }
            byte[] r = reply(address, f.command, status, data);
            pending.write(r, 0, r.length);
        }

        public int read(byte[] buf, int timeoutMs) {
            byte[] p = pending.toByteArray();
            if (p.length == 0) {
                return 0;
            }
            // deliver in two chunks to exercise reassembly
            int n = Math.min(buf.length, Math.max(1, p.length / 2));
            System.arraycopy(p, 0, buf, 0, n);
            pending.reset();
            pending.write(p, n, p.length - n);
            return n;
        }
    }

    static void put32(byte[] d, int off, long v) {
        for (int i = 0; i < 4; i++) {
            d[off + i] = (byte) (v >>> (8 * i));
        }
    }

    public static void main(String[] args) throws Exception {
        // 1. CRC8 transcription agrees with the table-driven implementation
        byte[] sample = {0x20, 0x00, 0x39, 0x30, (byte) 0xFF, (byte) 0xFF, 0x05, 0, 0, 0, 0, 0x55, 0x73};
        check(MirtekFrame.crc8(sample, 0, sample.length) == crc8Table(sample), "crc8 bitwise vs table");
        check(MirtekFrame.crc8(new byte[]{0}, 0, 1) == 0, "crc of zero byte");

        // 2. request layout and stuffing: address 0x5573 forces both escapes
        byte[] req = MirtekFrame.buildRequest(0x5573, 0x05, 0x55735573L, new byte[]{0x00});
        check(req[0] == 0x73 && (req[1] & 0xFF) == 0x55 && (req[req.length - 1] & 0xFF) == 0x55, "start/stop");
        String hex = Hex.of(req);
        check(hex.startsWith("73 55 21 00 73 22 73 11 FF FF 05 73 22 73 11 73 22 73 11 00"), "stuffed request: " + hex);
        // no unescaped 0x55 or 0x73 inside
        for (int i = 2; i < req.length - 1; i++) {
            int b = req[i] & 0xFF;
            if (b == 0x73) {
                int nx = req[i + 1] & 0xFF;
                check(nx == 0x11 || nx == 0x22, "escape followed by 11/22");
                i++;
            } else {
                check(b != 0x55, "no bare 0x55 inside");
            }
        }
        // 3. round trip through the parser
        MirtekFrame f = MirtekFrame.parse(req);
        check(f != null && f.request && f.destination == 0x5573 && f.command == 0x05, "parsed request");
        check(MirtekFrame.u32(f.statusOrPassword, 0) == 0x55735573L, "password round trip");
        check(f.data.length == 1 && f.data[0] == 0, "data round trip");
        check(f.rawEnd == req.length, "rawEnd");
        // incomplete buffer -> null, corrupted CRC -> exception
        byte[] cut = new byte[req.length - 1];
        System.arraycopy(req, 0, cut, 0, cut.length);
        check(MirtekFrame.parse(cut) == null, "incomplete frame is null");
        byte[] bad = req.clone();
        bad[10] ^= 0x01; // command byte
        boolean threw = false;
        try { MirtekFrame.parse(bad); } catch (IOException e) { threw = e.getMessage().contains("CRC8"); }
        check(threw, "crc error detected");
        // garbage before the frame is skipped
        byte[] noisy = new byte[req.length + 3];
        noisy[0] = 1; noisy[1] = 2; noisy[2] = 0x55;
        System.arraycopy(req, 0, noisy, 3, req.length);
        check(MirtekFrame.parse(noisy) != null, "leading garbage skipped");

        // 4. session against the fake meter
        FakeMeter meter = new FakeMeter(1234, 1934979925L);
        LogSink log = new LogSink();
        MirtekSession s = new MirtekSession(meter, log);
        s.setTimeout(300);
        s.configure(1234, 1934979925L);
        String ping = s.ping();
        check(ping.contains("адрес 1234") && ping.contains("1.5"), "ping: " + ping);
        EnergyReading e = s.readEnergy();
        check(e.newGeneration && e.decimals == 3 && e.tariffsUsed == 2, "energy cfg");
        check("2576.123".equals(e.format(e.total)) && "2000.000".equals(e.format(e.tariffs[0])), "energy values " + e);
        String dt = s.readDateTime();
        check(dt.startsWith("05.10.2026 14:15:30"), "datetime " + dt);
        MirtekFrame r = s.relay(0, true);
        check(meter.relayOff && MirtekStatus.loadDisconnected(r.statusOrPassword), "relay off");
        s.relay(0, false);
        check(!meter.relayOff, "relay on");

        // 5. wrong password -> MeterError 0x01, wrong address -> timeout
        s.configure(1234, 1);
        try { s.relay(0, true); check(false, "expected error"); }
        catch (MirtekSession.MeterError me) { check(me.code == 0x01 && !meter.relayOff, "wrong password rejected"); }
        s.configure(4321, 1934979925L);
        try { s.ping(); check(false, "expected timeout"); }
        catch (IOException ex) { check(ex.getMessage().contains("не ответил"), "timeout message: " + ex.getMessage()); }

        // 6. old generation meter: 0x05 with a parameter is refused (0x04), session retries without it
        FakeMeter old = new FakeMeter(7, 0);
        old.oldGeneration = true;
        MirtekSession s2 = new MirtekSession(old, new LogSink());
        s2.setTimeout(300);
        s2.configure(7, 0);
        EnergyReading e2 = s2.readEnergy();
        check(!e2.newGeneration && "1234.56".equals(e2.format(e2.total)), "old generation energy " + e2);

        System.out.println("MIRTEK SELFTEST OK");
    }
}
