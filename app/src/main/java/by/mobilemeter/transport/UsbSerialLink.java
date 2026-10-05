package by.mobilemeter.transport;

import android.hardware.usb.UsbDeviceConnection;
import android.hardware.usb.UsbManager;

import com.hoho.android.usbserial.driver.UsbSerialDriver;
import com.hoho.android.usbserial.driver.UsbSerialPort;

import java.io.IOException;

/** SerialLink over a USB-serial chip (FTDI, CP210x, CH34x, PL2303, CDC-ACM). */
public final class UsbSerialLink implements SerialLink {
    private final UsbManager manager;
    private final UsbSerialDriver driver;
    private final boolean driveDtrRts;
    private UsbDeviceConnection connection;
    private UsbSerialPort port;

    public UsbSerialLink(UsbManager manager, UsbSerialDriver driver, boolean driveDtrRts) {
        this.manager = manager;
        this.driver = driver;
        this.driveDtrRts = driveDtrRts;
    }

    @Override
    public void open() throws IOException {
        connection = manager.openDevice(driver.getDevice());
        if (connection == null) {
            throw new IOException("Система не открыла USB-устройство (нет разрешения или устройство занято)");
        }
        port = driver.getPorts().get(0);
        port.open(connection);
        port.setParameters(9600, 8, UsbSerialPort.STOPBITS_1, UsbSerialPort.PARITY_NONE);
        if (driveDtrRts) {
            // Many optical heads are powered from the DTR/RTS lines.
            try {
                port.setDTR(true);
                port.setRTS(true);
            } catch (UnsupportedOperationException ignored) {
                // the chip has no modem lines
            }
        }
    }

    @Override
    public void setParameters(int baud, int dataBits, int stopBits, int parity) throws IOException {
        int p;
        switch (parity) {
            case PARITY_ODD: p = UsbSerialPort.PARITY_ODD; break;
            case PARITY_EVEN: p = UsbSerialPort.PARITY_EVEN; break;
            default: p = UsbSerialPort.PARITY_NONE;
        }
        int s = stopBits == 2 ? UsbSerialPort.STOPBITS_2 : UsbSerialPort.STOPBITS_1;
        port.setParameters(baud, dataBits, s, p);
    }

    @Override
    public void write(byte[] data) throws IOException {
        port.write(data, 2000);
    }

    @Override
    public int read(byte[] buf, int timeoutMs) throws IOException {
        return port.read(buf, timeoutMs);
    }

    @Override
    public void purge() throws IOException {
        try {
            port.purgeHwBuffers(true, true);
        } catch (UnsupportedOperationException ignored) {
            // not every chip supports it; the session drains the buffer by reading
        }
    }

    @Override
    public boolean isOpen() {
        return port != null && port.isOpen();
    }

    @Override
    public String describe() {
        return UsbDevices.driverName(driver) + " " + driver.getDevice().getDeviceName();
    }

    @Override
    public void close() throws IOException {
        try {
            if (port != null) {
                port.close();
            }
        } finally {
            port = null;
            if (connection != null) {
                connection.close();
                connection = null;
            }
        }
    }
}
