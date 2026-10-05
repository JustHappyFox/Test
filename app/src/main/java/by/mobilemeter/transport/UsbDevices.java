package by.mobilemeter.transport;

import android.hardware.usb.UsbConstants;
import android.hardware.usb.UsbDevice;
import android.hardware.usb.UsbEndpoint;
import android.hardware.usb.UsbInterface;
import android.hardware.usb.UsbManager;
import android.os.Build;

import com.hoho.android.usbserial.driver.CdcAcmSerialDriver;
import com.hoho.android.usbserial.driver.Ch34xSerialDriver;
import com.hoho.android.usbserial.driver.Cp21xxSerialDriver;
import com.hoho.android.usbserial.driver.FtdiSerialDriver;
import com.hoho.android.usbserial.driver.ProlificSerialDriver;
import com.hoho.android.usbserial.driver.UsbSerialDriver;
import com.hoho.android.usbserial.driver.UsbSerialProber;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/** Enumerates attached USB devices and describes them for diagnostics. */
public final class UsbDevices {

    /** Index 0 means "auto"; the rest force a driver for a device the prober does not know. */
    public static final String[] FORCE_KINDS = {"авто", "CDC-ACM", "FTDI", "CP210x", "CH34x", "PL2303"};

    public static final class Entry {
        public final UsbDevice device;
        public final UsbSerialDriver driver;
        public final String label;

        Entry(UsbDevice device, UsbSerialDriver driver, String label) {
            this.device = device;
            this.driver = driver;
            this.label = label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    private UsbDevices() {
    }

    public static List<Entry> list(UsbManager manager) {
        List<Entry> out = new ArrayList<Entry>();
        Map<String, UsbDevice> devices = manager.getDeviceList();
        UsbSerialProber prober = UsbSerialProber.getDefaultProber();
        for (UsbDevice d : devices.values()) {
            UsbSerialDriver drv = prober.probeDevice(d);
            String name = productName(d);
            String label = String.format(Locale.US, "%04X:%04X %s [%s]",
                    d.getVendorId(), d.getProductId(), name,
                    drv != null ? driverName(drv) : "драйвер не найден");
            out.add(new Entry(d, drv, label));
        }
        return out;
    }

    public static UsbSerialDriver force(UsbDevice d, int kind) {
        switch (kind) {
            case 1: return new CdcAcmSerialDriver(d);
            case 2: return new FtdiSerialDriver(d);
            case 3: return new Cp21xxSerialDriver(d);
            case 4: return new Ch34xSerialDriver(d);
            case 5: return new ProlificSerialDriver(d);
            default: return null;
        }
    }

    public static String driverName(UsbSerialDriver drv) {
        String n = drv.getClass().getSimpleName();
        return n.endsWith("SerialDriver") ? n.substring(0, n.length() - "SerialDriver".length()) : n;
    }

    private static String productName(UsbDevice d) {
        String p = null;
        String m = null;
        if (Build.VERSION.SDK_INT >= 21) {
            p = d.getProductName();
            m = d.getManufacturerName();
        }
        if (p == null && m == null) {
            return d.getDeviceName();
        }
        if (m == null) {
            return p;
        }
        if (p == null) {
            return m;
        }
        return m + " " + p;
    }

    /** Multi-line description of descriptors, interfaces and endpoints. */
    public static String describe(UsbDevice d) {
        StringBuilder sb = new StringBuilder();
        sb.append(String.format(Locale.US, "Устройство %s  VID=%04X PID=%04X\n",
                d.getDeviceName(), d.getVendorId(), d.getProductId()));
        sb.append("  Изготовитель/продукт: ").append(productName(d)).append('\n');
        if (Build.VERSION.SDK_INT >= 23) {
            sb.append("  Версия устройства: ").append(d.getVersion()).append('\n');
        }
        sb.append(String.format(Locale.US, "  Класс устройства: %d / %d / %d\n",
                d.getDeviceClass(), d.getDeviceSubclass(), d.getDeviceProtocol()));
        for (int i = 0; i < d.getInterfaceCount(); i++) {
            UsbInterface itf = d.getInterface(i);
            sb.append(String.format(Locale.US, "  Интерфейс %d: класс %d (%s) / подкласс %d / протокол %d\n",
                    itf.getId(), itf.getInterfaceClass(), className(itf.getInterfaceClass()),
                    itf.getInterfaceSubclass(), itf.getInterfaceProtocol()));
            for (int j = 0; j < itf.getEndpointCount(); j++) {
                UsbEndpoint ep = itf.getEndpoint(j);
                sb.append(String.format(Locale.US, "    EP 0x%02X %s %s, пакет %d\n",
                        ep.getAddress(),
                        ep.getDirection() == UsbConstants.USB_DIR_IN ? "IN " : "OUT",
                        epType(ep.getType()), ep.getMaxPacketSize()));
            }
        }
        UsbSerialDriver drv = UsbSerialProber.getDefaultProber().probeDevice(d);
        sb.append("  Драйвер USB-serial: ").append(drv != null ? driverName(drv) : "не определён автоматически");
        return sb.toString();
    }

    private static String className(int c) {
        switch (c) {
            case UsbConstants.USB_CLASS_COMM: return "CDC управление";
            case UsbConstants.USB_CLASS_CDC_DATA: return "CDC данные";
            case UsbConstants.USB_CLASS_HID: return "HID";
            case UsbConstants.USB_CLASS_MASS_STORAGE: return "накопитель";
            case UsbConstants.USB_CLASS_VENDOR_SPEC: return "vendor-specific";
            case UsbConstants.USB_CLASS_HUB: return "хаб";
            default: return "другой";
        }
    }

    private static String epType(int t) {
        switch (t) {
            case UsbConstants.USB_ENDPOINT_XFER_BULK: return "bulk";
            case UsbConstants.USB_ENDPOINT_XFER_INT: return "interrupt";
            case UsbConstants.USB_ENDPOINT_XFER_ISOC: return "isochronous";
            default: return "control";
        }
    }
}
