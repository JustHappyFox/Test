package by.mobilemeter.protocol.iec62056;

/** Parsed identification message "/XXXZ\W<ident>" (IEC 62056-21, ГОСТ IEC 61107). */
public final class Identification {
    public final String raw;
    public final String manufacturer;
    public final char baudId;
    /** Baud rate proposed by the meter, or -1 if the character is not a standard rate id. */
    public final int baud;
    /** True for mode B meters (letters A..I): they switch the rate without an acknowledgement. */
    public final boolean modeB;
    /** True when the meter announces mode E ("\2" = HDLC, i.e. DLMS/COSEM). */
    public final boolean modeE;
    public final char modeChar;
    public final String ident;

    private Identification(String raw, String manufacturer, char baudId, int baud, boolean modeB,
                           boolean modeE, char modeChar, String ident) {
        this.raw = raw;
        this.manufacturer = manufacturer;
        this.baudId = baudId;
        this.baud = baud;
        this.modeB = modeB;
        this.modeE = modeE;
        this.modeChar = modeChar;
        this.ident = ident;
    }

    public static Identification parse(String line) {
        String s = line.trim();
        int slash = s.indexOf('/');
        if (slash < 0 || s.length() < slash + 5) {
            throw new IllegalArgumentException("Это не идентификационное сообщение: \"" + s + "\"");
        }
        s = s.substring(slash);
        String manufacturer = s.substring(1, 4);
        char z = s.charAt(4);
        int pos = 5;
        char modeChar = 0;
        if (s.length() > pos + 1 && s.charAt(pos) == '\\') {
            modeChar = s.charAt(pos + 1);
            pos += 2;
        }
        String ident = s.substring(pos);
        return new Identification(s, manufacturer, z, baudFor(z), z >= 'A' && z <= 'I',
                modeChar == '2', modeChar, ident);
    }

    public static int baudFor(char z) {
        switch (z) {
            case '0': return 300;
            case '1': return 600;
            case '2': return 1200;
            case '3': return 2400;
            case '4': return 4800;
            case '5': return 9600;
            case '6': return 19200;
            case 'A': return 600;
            case 'B': return 1200;
            case 'C': return 2400;
            case 'D': return 4800;
            case 'E': return 9600;
            case 'F': return 19200;
            default: return -1;
        }
    }

    public static char baudIdFor(int baud) {
        switch (baud) {
            case 300: return '0';
            case 600: return '1';
            case 1200: return '2';
            case 2400: return '3';
            case 4800: return '4';
            case 9600: return '5';
            case 19200: return '6';
            default: return '0';
        }
    }

    public String summary() {
        StringBuilder sb = new StringBuilder();
        sb.append("изготовитель=").append(manufacturer);
        sb.append(", модель=").append(ident);
        sb.append(", код скорости='").append(baudId).append('\'');
        if (baud > 0) {
            sb.append(" (").append(baud).append(" бод)");
        }
        if (modeB) {
            sb.append(", режим B");
        } else if (modeE) {
            sb.append(", режим E: HDLC/DLMS");
        } else if (modeChar != 0) {
            sb.append(", признак '\\").append(modeChar).append('\'');
        } else {
            sb.append(", режим C");
        }
        return sb.toString();
    }
}
