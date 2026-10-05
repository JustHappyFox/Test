package by.mobilemeter.protocol.iec62056;

import java.util.Collections;
import java.util.List;

/** One data line of a readout: name(value*unit)(value2)... */
public final class Register {
    public final String name;
    public final List<String> values;
    public final String unit;

    public Register(String name, List<String> values, String unit) {
        this.name = name;
        this.values = Collections.unmodifiableList(values);
        this.unit = unit;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder(name).append(" = ");
        for (int i = 0; i < values.size(); i++) {
            if (i > 0) {
                sb.append(" | ");
            }
            sb.append(values.get(i));
        }
        if (unit != null) {
            sb.append(' ').append(unit);
        }
        return sb.toString();
    }
}
