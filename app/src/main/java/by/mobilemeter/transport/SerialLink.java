package by.mobilemeter.transport;

import java.io.Closeable;
import java.io.IOException;

/**
 * Byte-oriented serial channel to a meter. Today it is a USB-serial port
 * (optical head or radio module over OTG); later a radio-module framing
 * layer can implement the same interface.
 */
public interface SerialLink extends Closeable {
    int PARITY_NONE = 0;
    int PARITY_ODD = 1;
    int PARITY_EVEN = 2;

    void open() throws IOException;

    void setParameters(int baud, int dataBits, int stopBits, int parity) throws IOException;

    void write(byte[] data) throws IOException;

    /** Reads up to buf.length bytes. Returns the number read, 0 on timeout. */
    int read(byte[] buf, int timeoutMs) throws IOException;

    /** Discards anything pending in the receive buffer. */
    void purge() throws IOException;

    boolean isOpen();

    String describe();
}
