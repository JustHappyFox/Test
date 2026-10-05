package by.mobilemeter.protocol.iec62056;

import by.mobilemeter.transport.SerialLink;
import by.mobilemeter.util.Hex;
import by.mobilemeter.util.LogSink;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Plain-JVM self test of the protocol layer with a scripted fake meter.
 * Run: tools/run-selftest.sh (no JUnit needed).
 */
public final class SessionSelfTest {

    /** Fake meter: answers each expected request with a canned reply. */
    static final class FakeLink implements SerialLink {
        final List<byte[]> expect = new ArrayList<byte[]>();
        final List<byte[]> reply = new ArrayList<byte[]>();
        final ByteArrayOutputStream pending = new ByteArrayOutputStream();
        final List<Integer> bauds = new ArrayList<Integer>();
        int step = 0;
        /** When set, the fake meter only hears requests sent at exactly these settings. */
        int onlyBaud = -1;
        int onlyParity = -1;
        int curBaud = -1;
        int curParity = -1;

        void script(String request, byte[] answer) {
            script(Hex.ascii(request), answer);
        }

        void script(byte[] request, byte[] answer) {
            expect.add(request);
            reply.add(answer);
        }

        public void open() {
        }

        public void setParameters(int baud, int dataBits, int stopBits, int parity) {
            bauds.add(baud);
            curBaud = baud;
            curParity = parity;
        }

        public void write(byte[] data) throws IOException {
            if (onlyBaud >= 0 && (curBaud != onlyBaud || curParity != onlyParity)) {
                return; // wrong settings: the meter sees garbage and stays silent
            }
            if (step >= expect.size()) {
                throw new IOException("unexpected write: " + Hex.printable(data, 0, data.length));
            }
            byte[] e = expect.get(step);
            if (!java.util.Arrays.equals(e, data)) {
                throw new IOException("step " + step + ": expected " + Hex.printable(e, 0, e.length)
                        + " got " + Hex.printable(data, 0, data.length));
            }
            pending.write(reply.get(step), 0, reply.get(step).length);
            step++;
        }

        public int read(byte[] buf, int timeoutMs) {
            byte[] p = pending.toByteArray();
            if (p.length == 0) {
                return 0;
            }
            int n = Math.min(buf.length, p.length);
            System.arraycopy(p, 0, buf, 0, n);
            pending.reset();
            pending.write(p, n, p.length - n);
            return n;
        }

        public void purge() {
        }

        public boolean isOpen() {
            return true;
        }

        public String describe() {
            return "fake";
        }

        public void close() {
        }
    }

    static byte[] frame(byte start, String cmd, String data) {
        ByteArrayOutputStream b = new ByteArrayOutputStream();
        b.write(start);
        byte[] c = Hex.ascii(cmd);
        b.write(c, 0, c.length);
        if (data != null) {
            b.write(Iec62056Session.STX);
            byte[] d = Hex.ascii(data);
            b.write(d, 0, d.length);
        }
        b.write(Iec62056Session.ETX);
        byte[] f = b.toByteArray();
        b.write(Iec62056Session.bcc(f, 1, f.length));
        return b.toByteArray();
    }

    static byte[] dataBlock(String text) {
        ByteArrayOutputStream b = new ByteArrayOutputStream();
        b.write(Iec62056Session.STX);
        byte[] d = Hex.ascii(text);
        b.write(d, 0, d.length);
        b.write(Iec62056Session.ETX);
        byte[] f = b.toByteArray();
        b.write(Iec62056Session.bcc(f, 1, f.length));
        return b.toByteArray();
    }

    static void check(boolean cond, String msg) {
        if (!cond) {
            throw new AssertionError(msg);
        }
    }

    public static void main(String[] args) throws Exception {
        // 1. identification parsing
        Identification id = Identification.parse("/EKT5CE208v3\r\n");
        check("EKT".equals(id.manufacturer), "manufacturer");
        check(id.baud == 9600 && id.baudId == '5', "baud 9600");
        check(!id.modeE && !id.modeB, "mode C");
        check("CE208v3".equals(id.ident), "ident " + id.ident);

        Identification idE = Identification.parse("/EKT5\\2CE318BY");
        check(idE.modeE, "mode E flag");
        check("CE318BY".equals(idE.ident), "ident E " + idE.ident);

        Identification idB = Identification.parse("/MIRE2301");
        check(idB.modeB && idB.baud == 9600, "mode B");

        // 2. BCC known value: bytes "R1" STX "ET0PE()" ETX
        byte[] f = frame(Iec62056Session.SOH, "R1", "ET0PE()");
        int x = 0;
        for (int i = 1; i < f.length - 1; i++) {
            x ^= f[i];
        }
        check((f[f.length - 1] & 0x7F) == (x & 0x7F), "bcc");

        // 3. readout parser
        List<Register> regs = ReadoutParser.parse("1.8.0(05801.969*kWh)\r\n1.8.1(01000.5*kWh)(0.1)\r\n0.9.2(23-10-05)\r\n!\r\n");
        check(regs.size() == 3, "3 registers, got " + regs.size());
        check("1.8.0".equals(regs.get(0).name) && "05801.969".equals(regs.get(0).values.get(0))
                && "kWh".equals(regs.get(0).unit), "reg 0");
        check(regs.get(1).values.size() == 2, "two values");
        check(regs.get(2).unit == null, "no unit");

        // 4. full readout session against the fake meter (initial 300, switch to 9600)
        FakeLink link = new FakeLink();
        link.script("/?!\r\n", Hex.ascii("/EKT5CE208v3\r\n"));
        link.script("\u0006050\r\n", dataBlock("1.8.0(05801.969*kWh)\r\n1.8.1(01000.5*kWh)\r\n!\r\n"));
        LogSink log = new LogSink();
        Iec62056Session s = new Iec62056Session(link, log);
        s.configure(300, true, true);
        List<Register> got = s.readout("");
        check(got.size() == 2, "readout size " + got.size());
        check(link.bauds.size() == 2 && link.bauds.get(0) == 300 && link.bauds.get(1) == 9600, "baud switch " + link.bauds);

        // 5. programming mode + R1 + B0
        FakeLink l2 = new FakeLink();
        l2.script("/?80430692!\r\n", Hex.ascii("/EKT5CE318BY\r\n"));
        l2.script("\u0006051\r\n", frame(Iec62056Session.SOH, "P0", "(00000000)"));
        l2.script(frame(Iec62056Session.SOH, "P1", "(777777)"), new byte[]{Iec62056Session.ACK});
        l2.script(frame(Iec62056Session.SOH, "R1", "ET0PE()"), dataBlock("ET0PE(05801.969)(01000.5)(0.0)(0.0)(0.0)"));
        l2.script(frame(Iec62056Session.SOH, "B0", null), new byte[0]);
        Iec62056Session s2 = new Iec62056Session(l2, new LogSink());
        s2.configure(300, true, true);
        s2.enterProgrammingMode("80430692", "777777");
        String body = s2.readRegister("ET0PE");
        check(body.startsWith("ET0PE(05801.969)"), "R1 body " + body);
        s2.signOff();
        check(l2.step == 5, "all steps used: " + l2.step);

        // 6. no-switch mode keeps the initial baud and sends '0' as rate id at 300
        FakeLink l3 = new FakeLink();
        l3.script("/?!\r\n", Hex.ascii("/EKT5CE208v3\r\n"));
        l3.script("\u0006000\r\n", dataBlock("!\r\n"));
        Iec62056Session s3 = new Iec62056Session(l3, new LogSink());
        s3.configure(300, false, true);
        s3.readout("");
        check(l3.bauds.size() == 1, "no baud switch");

        // 7. auto mode: the meter answers only at 9600 8N1; the scan must find it and then reuse it
        FakeLink l4 = new FakeLink();
        l4.onlyBaud = 9600;
        l4.onlyParity = SerialLink.PARITY_NONE;
        l4.script("/?!\r\n", Hex.ascii("/EKT5CE318BY\r\n"));
        l4.script("/?!\r\n", Hex.ascii("/EKT5CE318BY\r\n"));
        l4.script("\u0006050\r\n", dataBlock("1.8.0(00000.18*kWh)\r\n!\r\n"));
        Iec62056Session s4 = new Iec62056Session(l4, new LogSink());
        s4.configureAuto(0, null, true);
        Identification found = s4.identify("");
        check("CE318BY".equals(found.ident), "auto ident");
        check("9600 8N1".equals(s4.lockedDescription()), "locked " + s4.lockedDescription());
        int baudsAfterScan = l4.bauds.size();
        check(baudsAfterScan == 3, "scan order 9600 7E1, 300 7E1, 9600 8N1: " + l4.bauds);
        List<Register> r4 = s4.readout("");
        check(r4.size() == 1, "readout after auto");
        check(l4.bauds.size() == baudsAfterScan + 1, "locked settings reused without rescanning: " + l4.bauds);

        // 8. auto mode with nothing answering: one clear error listing the attempts
        FakeLink l5 = new FakeLink();
        l5.onlyBaud = 1; // never matches
        Iec62056Session s5 = new Iec62056Session(l5, new LogSink());
        s5.configureAuto(0, null, true);
        try {
            s5.identify("");
            check(false, "expected failure");
        } catch (IOException e) {
            check(e.getMessage().contains("300 7E1") && e.getMessage().contains("19200 7E1"), "attempts listed: " + e.getMessage());
        }

        // 9. auto format only, fixed 300 baud
        FakeLink l6 = new FakeLink();
        l6.onlyBaud = 300;
        l6.onlyParity = SerialLink.PARITY_NONE;
        l6.script("/?!\r\n", Hex.ascii("/EKT5CE208v3\r\n"));
        Iec62056Session s6 = new Iec62056Session(l6, new LogSink());
        s6.configureAuto(300, null, true);
        s6.identify("");
        check("300 8N1".equals(s6.lockedDescription()), "fixed baud, auto format: " + s6.lockedDescription());

        System.out.println("SELFTEST OK: " + log.text().split("\n").length + " log lines in scenario 4");
    }
}
