package by.mobilemeter.protocol.iec62056;

import java.util.ArrayList;
import java.util.List;

/** Parses the data block of a readout or an R1 answer into registers. */
public final class ReadoutParser {
    private ReadoutParser() {
    }

    public static List<Register> parse(String block) {
        List<Register> out = new ArrayList<Register>();
        String[] lines = block.split("\r?\n");
        for (String line : lines) {
            String s = line.trim();
            if (s.isEmpty() || s.equals("!")) {
                continue;
            }
            Register r = parseLine(s);
            if (r != null) {
                out.add(r);
            }
        }
        return out;
    }

    /** Parses "name(v1*unit)(v2)"; returns null for lines without parentheses. */
    public static Register parseLine(String s) {
        int open = s.indexOf('(');
        if (open < 0) {
            return null;
        }
        String name = s.substring(0, open).trim();
        List<String> values = new ArrayList<String>();
        String unit = null;
        int pos = open;
        while (pos < s.length() && s.charAt(pos) == '(') {
            int close = s.indexOf(')', pos);
            if (close < 0) {
                values.add(s.substring(pos + 1));
                break;
            }
            String v = s.substring(pos + 1, close);
            int star = v.indexOf('*');
            if (star >= 0) {
                if (unit == null) {
                    unit = v.substring(star + 1);
                }
                v = v.substring(0, star);
            }
            values.add(v);
            pos = close + 1;
        }
        return new Register(name, values, unit);
    }
}
