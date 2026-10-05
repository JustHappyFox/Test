package by.mobilemeter.util;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

/** Thread-safe text log with optional hex dump of the exchanged bytes. */
public final class LogSink {

    public interface Listener {
        void onLine(String line);
    }

    private final StringBuilder all = new StringBuilder();
    private final SimpleDateFormat fmt = new SimpleDateFormat("HH:mm:ss.SSS", Locale.US);
    private volatile boolean hex = true;
    private volatile Listener listener;

    public void setListener(Listener l) {
        listener = l;
    }

    public void setHex(boolean enabled) {
        hex = enabled;
    }

    public void line(String text) {
        String stamped;
        synchronized (this) {
            stamped = fmt.format(new Date()) + " " + text;
            all.append(stamped).append('\n');
        }
        Listener l = listener;
        if (l != null) {
            l.onLine(stamped);
        }
    }

    public void tx(byte[] data) {
        line("TX " + render(data, 0, data.length));
    }

    public void rx(byte[] data, int off, int len) {
        line("RX " + render(data, off, len));
    }

    private String render(byte[] data, int off, int len) {
        String p = Hex.printable(data, off, len);
        if (!hex) {
            return p;
        }
        return p + "   [" + Hex.of(data, off, len) + "]";
    }

    public synchronized String text() {
        return all.toString();
    }

    public synchronized void clear() {
        all.setLength(0);
    }
}
