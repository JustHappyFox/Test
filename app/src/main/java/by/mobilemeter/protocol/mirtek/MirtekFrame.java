package by.mobilemeter.protocol.mirtek;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * One packet of the МИРТЕК protocol (document "Описание протокола", v1.20, 2015):
 * <pre>
 * 0x73 0x55 | Parameters | 0x00 | AddrDst(2, LE) | AddrSrc(2, LE) | Command | Password or Status (4) | Data (0..31) | CRC8 | 0x55
 * </pre>
 * Parameters: bit7 C (encoded data), bit6 V0 (0 = meter), bit5 D (1 = request), bits 4..0 = data length.
 * CRC8: polynomial 0xA9, initial 0x00, over everything between the start bytes and the CRC itself.
 * Byte stuffing between start and stop: 0x55 -> 0x73 0x11, 0x73 -> 0x73 0x22.
 */
public final class MirtekFrame {
    public static final int START1 = 0x73;
    public static final int START2 = 0x55;
    public static final int STOP = 0x55;
    public static final int ESCAPE = 0x73;
    public static final int ESCAPED_55 = 0x11;
    public static final int ESCAPED_73 = 0x22;
    public static final int SOURCE_ADDRESS = 0xFFFF;
    public static final int BROADCAST_ADDRESS = 0xFFFF;
    public static final int MAX_DATA = 31;

    public final boolean request;
    public final boolean encoded;
    public final boolean uspd;
    public final int destination;
    public final int source;
    public final int command;
    /** Password in a request, status in a reply: 4 bytes. */
    public final byte[] statusOrPassword;
    public final byte[] data;
    /** Index in the raw buffer just after the stop byte (set by parse()). */
    public final int rawEnd;

    private MirtekFrame(boolean request, boolean encoded, boolean uspd, int destination, int source,
                        int command, byte[] statusOrPassword, byte[] data, int rawEnd) {
        this.request = request;
        this.encoded = encoded;
        this.uspd = uspd;
        this.destination = destination;
        this.source = source;
        this.command = command;
        this.statusOrPassword = statusOrPassword;
        this.data = data;
        this.rawEnd = rawEnd;
    }

    /** Builds a stuffed request packet ready to be written to the port. */
    public static byte[] buildRequest(int destination, int command, long password, byte[] data) {
        if (data == null) {
            data = new byte[0];
        }
        if (data.length > MAX_DATA) {
            throw new IllegalArgumentException("Поле данных длиннее 31 байта: " + data.length);
        }
        ByteArrayOutputStream body = new ByteArrayOutputStream();
        body.write(0x20 | data.length); // C=0, V0=0, D=1
        body.write(0x00);
        body.write(destination & 0xFF);
        body.write((destination >>> 8) & 0xFF);
        body.write(SOURCE_ADDRESS & 0xFF);
        body.write((SOURCE_ADDRESS >>> 8) & 0xFF);
        body.write(command & 0xFF);
        for (int i = 0; i < 4; i++) {
            body.write((int) ((password >>> (8 * i)) & 0xFF));
        }
        body.write(data, 0, data.length);
        byte[] b = body.toByteArray();
        int crc = crc8(b, 0, b.length);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        out.write(START1);
        out.write(START2);
        for (byte x : b) {
            stuff(out, x & 0xFF);
        }
        stuff(out, crc);
        out.write(STOP);
        return out.toByteArray();
    }

    private static void stuff(ByteArrayOutputStream out, int b) {
        if (b == 0x55) {
            out.write(ESCAPE);
            out.write(ESCAPED_55);
        } else if (b == 0x73) {
            out.write(ESCAPE);
            out.write(ESCAPED_73);
        } else {
            out.write(b);
        }
    }

    /** CRC8, polynomial 0xA9, initial value 0 — a transcription of the vendor's Delphi sample. */
    public static int crc8(byte[] data, int off, int len) {
        int crc = 0;
        for (int i = off; i < off + len; i++) {
            int b = data[i] & 0xFF;
            for (int k = 0; k < 8; k++) {
                if (((b ^ crc) & 0x80) == 0) {
                    crc = (crc << 1) & 0xFF;
                } else {
                    crc = ((crc << 1) ^ 0xA9) & 0xFF;
                }
                b = (b << 1) & 0xFF;
            }
        }
        return crc;
    }

    /**
     * Parses the first complete frame found in raw. Returns null when the buffer does not yet
     * hold a whole frame; throws on a framing or CRC error.
     */
    public static MirtekFrame parse(byte[] raw) throws IOException {
        int start = -1;
        for (int i = 0; i + 1 < raw.length; i++) {
            if ((raw[i] & 0xFF) == START1 && (raw[i + 1] & 0xFF) == START2) {
                start = i;
                break;
            }
        }
        if (start < 0) {
            return null;
        }
        ByteArrayOutputStream un = new ByteArrayOutputStream();
        boolean escape = false;
        int end = -1;
        for (int i = start + 2; i < raw.length; i++) {
            int b = raw[i] & 0xFF;
            if (escape) {
                if (b == ESCAPED_55) {
                    un.write(0x55);
                } else if (b == ESCAPED_73) {
                    un.write(0x73);
                } else {
                    throw new IOException(String.format("Ошибка байтстаффинга: 0x73 0x%02X", b));
                }
                escape = false;
                continue;
            }
            if (b == ESCAPE) {
                escape = true;
                continue;
            }
            if (b == STOP) {
                end = i + 1;
                break;
            }
            un.write(b);
        }
        if (end < 0) {
            return null;
        }
        byte[] u = un.toByteArray();
        final int header = 1 + 1 + 2 + 2 + 1 + 4;
        if (u.length < header + 1) {
            throw new IOException("Слишком короткий пакет: " + u.length + " байт");
        }
        int crc = u[u.length - 1] & 0xFF;
        int calc = crc8(u, 0, u.length - 1);
        if (crc != calc) {
            throw new IOException(String.format("CRC8 не совпал: ожидали %02X, получили %02X", calc, crc));
        }
        int params = u[0] & 0xFF;
        int len = params & 0x1F;
        if (header + len != u.length - 1) {
            throw new IOException("Длина данных " + len + " не соответствует длине пакета " + u.length);
        }
        byte[] status = new byte[4];
        System.arraycopy(u, 7, status, 0, 4);
        byte[] data = new byte[len];
        System.arraycopy(u, header, data, 0, len);
        return new MirtekFrame((params & 0x20) != 0, (params & 0x80) != 0, (params & 0x40) != 0,
                (u[2] & 0xFF) | ((u[3] & 0xFF) << 8), (u[4] & 0xFF) | ((u[5] & 0xFF) << 8),
                u[6] & 0xFF, status, data, end);
    }

    /** Error code from the status field of a reply (4th byte). */
    public int errorCode() {
        return statusOrPassword[3] & 0xFF;
    }

    public static int u16(byte[] d, int off) {
        return (d[off] & 0xFF) | ((d[off + 1] & 0xFF) << 8);
    }

    public static long u32(byte[] d, int off) {
        return (d[off] & 0xFFL) | ((d[off + 1] & 0xFFL) << 8) | ((d[off + 2] & 0xFFL) << 16) | ((d[off + 3] & 0xFFL) << 24);
    }
}
