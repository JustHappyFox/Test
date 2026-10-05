package by.mobilemeter.util;

/** Hex and printable renderings of byte buffers for the log. */
public final class Hex {
    private static final char[] DIGITS = "0123456789ABCDEF".toCharArray();

    private Hex() {
    }

    public static String of(byte[] data) {
        return of(data, 0, data.length);
    }

    public static String of(byte[] data, int off, int len) {
        StringBuilder sb = new StringBuilder(len * 3);
        for (int i = off; i < off + len; i++) {
            if (sb.length() > 0) {
                sb.append(' ');
            }
            int v = data[i] & 0xFF;
            sb.append(DIGITS[v >>> 4]).append(DIGITS[v & 0x0F]);
        }
        return sb.toString();
    }

    /** ASCII rendering where control characters are shown as mnemonics. */
    public static String printable(byte[] data, int off, int len) {
        StringBuilder sb = new StringBuilder(len + 16);
        for (int i = off; i < off + len; i++) {
            int v = data[i] & 0x7F; // 7-bit protocols: strip parity bit if the driver passes it through
            switch (v) {
                case 0x01: sb.append("<SOH>"); break;
                case 0x02: sb.append("<STX>"); break;
                case 0x03: sb.append("<ETX>"); break;
                case 0x04: sb.append("<EOT>"); break;
                case 0x06: sb.append("<ACK>"); break;
                case 0x15: sb.append("<NAK>"); break;
                case 0x0D: sb.append("<CR>"); break;
                case 0x0A: sb.append("<LF>"); break;
                default:
                    if (v >= 0x20 && v < 0x7F) {
                        sb.append((char) v);
                    } else {
                        sb.append('<').append(DIGITS[v >>> 4]).append(DIGITS[v & 0x0F]).append('>');
                    }
            }
        }
        return sb.toString();
    }

    public static byte[] ascii(String s) {
        byte[] out = new byte[s.length()];
        for (int i = 0; i < s.length(); i++) {
            out[i] = (byte) (s.charAt(i) & 0x7F);
        }
        return out;
    }

    public static String ascii(byte[] data, int off, int len) {
        StringBuilder sb = new StringBuilder(len);
        for (int i = off; i < off + len; i++) {
            sb.append((char) (data[i] & 0x7F));
        }
        return sb.toString();
    }
}
