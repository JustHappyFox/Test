package by.mobilemeter;

import android.app.Activity;
import android.app.PendingIntent;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.hardware.usb.UsbDevice;
import android.hardware.usb.UsbManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.CompoundButton;
import android.widget.EditText;
import android.widget.ScrollView;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import by.mobilemeter.protocol.iec62056.Iec62056Session;
import by.mobilemeter.transport.UsbDevices;
import by.mobilemeter.transport.UsbSerialLink;
import by.mobilemeter.util.LogSink;

import com.hoho.android.usbserial.driver.UsbSerialDriver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Diagnostic screen: lists USB devices on the OTG port, opens a serial link to the
 * optical head or radio module, and runs IEC 62056-21 requests against the meter.
 */
public class MainActivity extends Activity {
    private static final String ACTION_USB_PERMISSION = "by.mobilemeter.USB_PERMISSION";
    private static final String[] BAUDS = {"300", "600", "1200", "2400", "4800", "9600", "19200"};
    private static final String[] FORMATS = {"7E1", "8N1"};

    private Spinner spinnerDevice;
    private Spinner spinnerBaud;
    private Spinner spinnerFormat;
    private Spinner spinnerDriver;
    private CheckBox cbSwitchBaud;
    private CheckBox cbHex;
    private CheckBox cbDtr;
    private EditText editAddress;
    private EditText editPassword;
    private EditText editCommand;
    private Button btnConnect;
    private TextView textStatus;
    private TextView textLog;
    private ScrollView scrollLog;

    private final ExecutorService exec = Executors.newSingleThreadExecutor();
    private final Handler ui = new Handler(Looper.getMainLooper());
    private final LogSink log = new LogSink();

    private UsbManager usbManager;
    private List<UsbDevices.Entry> entries = new ArrayList<UsbDevices.Entry>();
    private UsbSerialLink link;
    private Iec62056Session session;
    private UsbDevices.Entry pendingPermission;
    private UsbSerialDriver pendingDriver;

    private final BroadcastReceiver usbReceiver = new BroadcastReceiver() {
        @Override
        public void onReceive(Context context, Intent intent) {
            String action = intent.getAction();
            if (ACTION_USB_PERMISSION.equals(action)) {
                boolean granted = intent.getBooleanExtra(UsbManager.EXTRA_PERMISSION_GRANTED, false);
                UsbDevices.Entry e = pendingPermission;
                UsbSerialDriver d = pendingDriver;
                pendingPermission = null;
                pendingDriver = null;
                if (granted && e != null && d != null) {
                    log.line("Разрешение на USB получено");
                    connectNow(e, d);
                } else {
                    log.line("Разрешение на USB-устройство отклонено");
                }
            } else if (UsbManager.ACTION_USB_DEVICE_DETACHED.equals(action)) {
                log.line("USB-устройство отключено");
                disconnect();
                refreshDevices();
            } else if (UsbManager.ACTION_USB_DEVICE_ATTACHED.equals(action)) {
                log.line("USB-устройство подключено");
                refreshDevices();
            }
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        usbManager = (UsbManager) getSystemService(Context.USB_SERVICE);

        spinnerDevice = (Spinner) findViewById(R.id.spinnerDevice);
        spinnerBaud = (Spinner) findViewById(R.id.spinnerBaud);
        spinnerFormat = (Spinner) findViewById(R.id.spinnerFormat);
        spinnerDriver = (Spinner) findViewById(R.id.spinnerDriver);
        cbSwitchBaud = (CheckBox) findViewById(R.id.cbSwitchBaud);
        cbHex = (CheckBox) findViewById(R.id.cbHex);
        cbDtr = (CheckBox) findViewById(R.id.cbDtr);
        editAddress = (EditText) findViewById(R.id.editAddress);
        editPassword = (EditText) findViewById(R.id.editPassword);
        editCommand = (EditText) findViewById(R.id.editCommand);
        btnConnect = (Button) findViewById(R.id.btnConnect);
        textStatus = (TextView) findViewById(R.id.textStatus);
        textLog = (TextView) findViewById(R.id.textLog);
        scrollLog = (ScrollView) findViewById(R.id.scrollLog);

        spinnerBaud.setAdapter(simpleAdapter(BAUDS));
        spinnerBaud.setSelection(0);
        spinnerFormat.setAdapter(simpleAdapter(FORMATS));
        spinnerDriver.setAdapter(simpleAdapter(UsbDevices.FORCE_KINDS));

        log.setListener(new LogSink.Listener() {
            @Override
            public void onLine(final String line) {
                ui.post(new Runnable() {
                    @Override
                    public void run() {
                        textLog.append(line + "\n");
                        scrollLog.post(new Runnable() {
                            @Override
                            public void run() {
                                scrollLog.fullScroll(View.FOCUS_DOWN);
                            }
                        });
                    }
                });
            }
        });
        cbHex.setOnCheckedChangeListener(new CompoundButton.OnCheckedChangeListener() {
            @Override
            public void onCheckedChanged(CompoundButton b, boolean checked) {
                log.setHex(checked);
            }
        });

        bind(R.id.btnRefresh, new Runnable() { public void run() { refreshDevices(); } });
        bind(R.id.btnConnect, new Runnable() { public void run() { toggleConnect(); } });
        bind(R.id.btnInfo, new Runnable() { public void run() { showDeviceInfo(); } });
        bind(R.id.btnIdentify, new Runnable() { public void run() { runIdentify(); } });
        bind(R.id.btnReadout, new Runnable() { public void run() { runReadout(); } });
        bind(R.id.btnProgram, new Runnable() { public void run() { runProgram(); } });
        bind(R.id.btnReadReg, new Runnable() { public void run() { runReadRegister(); } });
        bind(R.id.btnRaw, new Runnable() { public void run() { runRaw(); } });
        bind(R.id.btnSignOff, new Runnable() { public void run() { runSignOff(); } });
        bind(R.id.btnShare, new Runnable() { public void run() { shareLog(); } });
        bind(R.id.btnClear, new Runnable() { public void run() { textLog.setText(""); log.clear(); } });

        IntentFilter filter = new IntentFilter();
        filter.addAction(ACTION_USB_PERMISSION);
        filter.addAction(UsbManager.ACTION_USB_DEVICE_ATTACHED);
        filter.addAction(UsbManager.ACTION_USB_DEVICE_DETACHED);
        if (Build.VERSION.SDK_INT >= 33) {
            registerReceiver(usbReceiver, filter, Context.RECEIVER_EXPORTED);
        } else {
            registerReceiver(usbReceiver, filter);
        }

        log.line("Мобильный контролёр, пробник IEC 61107. Android " + Build.VERSION.RELEASE
                + " (API " + Build.VERSION.SDK_INT + "), " + Build.MANUFACTURER + " " + Build.MODEL);
        log.line("Подключите оптоголовку или радиомодуль через OTG, нажмите «Обновить», затем «Подключить».");
        log.line("Если счётчик молчит на 300 бод, попробуйте начальную скорость 9600.");
        refreshDevices();
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        if (UsbManager.ACTION_USB_DEVICE_ATTACHED.equals(intent.getAction())) {
            refreshDevices();
        }
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        try {
            unregisterReceiver(usbReceiver);
        } catch (IllegalArgumentException ignored) {
            // not registered
        }
        disconnect();
        exec.shutdownNow();
    }

    // ---- device handling ----------------------------------------------------

    private void refreshDevices() {
        entries = UsbDevices.list(usbManager);
        List<String> labels = new ArrayList<String>();
        for (UsbDevices.Entry e : entries) {
            labels.add(e.label);
        }
        if (labels.isEmpty()) {
            labels.add(getString(R.string.no_devices));
        }
        spinnerDevice.setAdapter(simpleAdapter(labels.toArray(new String[labels.size()])));
        log.line("Найдено USB-устройств: " + entries.size());
        for (UsbDevices.Entry e : entries) {
            log.line("  " + e.label);
        }
    }

    private UsbDevices.Entry selectedEntry() {
        int i = spinnerDevice.getSelectedItemPosition();
        if (i < 0 || i >= entries.size()) {
            log.line("Сначала выберите USB-устройство");
            return null;
        }
        return entries.get(i);
    }

    private void showDeviceInfo() {
        UsbDevices.Entry e = selectedEntry();
        if (e != null) {
            log.line(UsbDevices.describe(e.device));
        }
    }

    private void toggleConnect() {
        if (link != null) {
            disconnect();
            return;
        }
        UsbDevices.Entry e = selectedEntry();
        if (e == null) {
            return;
        }
        UsbSerialDriver driver = e.driver;
        int forced = spinnerDriver.getSelectedItemPosition();
        if (forced > 0) {
            driver = UsbDevices.force(e.device, forced);
            log.line("Драйвер выбран принудительно: " + UsbDevices.FORCE_KINDS[forced]);
        }
        if (driver == null) {
            log.line("Для этого устройства нет драйвера USB-serial. Нажмите «Инфо USB», пришлите лог, "
                    + "либо попробуйте выбрать драйвер вручную (чаще всего CDC-ACM).");
            return;
        }
        if (!usbManager.hasPermission(e.device)) {
            pendingPermission = e;
            pendingDriver = driver;
            Intent intent = new Intent(ACTION_USB_PERMISSION);
            intent.setPackage(getPackageName());
            int flags = Build.VERSION.SDK_INT >= 31 ? PendingIntent.FLAG_MUTABLE : 0;
            PendingIntent pi = PendingIntent.getBroadcast(this, 0, intent, flags);
            log.line("Запрашиваю разрешение на USB-устройство…");
            usbManager.requestPermission(e.device, pi);
            return;
        }
        connectNow(e, driver);
    }

    private void connectNow(final UsbDevices.Entry e, final UsbSerialDriver driver) {
        final boolean dtr = cbDtr.isChecked();
        exec.submit(new Runnable() {
            @Override
            public void run() {
                try {
                    UsbSerialLink l = new UsbSerialLink(usbManager, driver, dtr);
                    l.open();
                    link = l;
                    session = new Iec62056Session(l, log);
                    log.line("Порт открыт: " + l.describe());
                    setStatus("Подключено: " + e.label, true);
                } catch (IOException ex) {
                    log.line("Ошибка открытия порта: " + ex.getMessage());
                    setStatus(getString(R.string.status_idle), false);
                } catch (RuntimeException ex) {
                    log.line("Ошибка открытия порта: " + ex);
                    setStatus(getString(R.string.status_idle), false);
                }
            }
        });
    }

    private void disconnect() {
        final UsbSerialLink l = link;
        link = null;
        session = null;
        setStatus(getString(R.string.status_idle), false);
        if (l == null) {
            return;
        }
        exec.submit(new Runnable() {
            @Override
            public void run() {
                try {
                    l.close();
                    log.line("Порт закрыт");
                } catch (IOException ex) {
                    log.line("Ошибка закрытия порта: " + ex.getMessage());
                }
            }
        });
    }

    private void setStatus(final String text, final boolean connected) {
        ui.post(new Runnable() {
            @Override
            public void run() {
                textStatus.setText(text);
                btnConnect.setText(connected ? R.string.btn_disconnect : R.string.btn_connect);
            }
        });
    }

    // ---- meter operations ---------------------------------------------------

    private boolean prepareSession() {
        if (session == null || link == null || !link.isOpen()) {
            log.line("Порт не открыт. Сначала нажмите «Подключить».");
            return false;
        }
        int baud = Integer.parseInt(BAUDS[spinnerBaud.getSelectedItemPosition()]);
        boolean sevenE1 = spinnerFormat.getSelectedItemPosition() == 0;
        session.configure(baud, cbSwitchBaud.isChecked(), sevenE1);
        return true;
    }

    private void runIdentify() {
        if (!prepareSession()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        submit("идентификация", new Op() {
            @Override
            public void run() throws IOException {
                session.identify(addr);
            }
        });
    }

    private void runReadout() {
        if (!prepareSession()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        submit("считывание показаний", new Op() {
            @Override
            public void run() throws IOException {
                session.readout(addr);
            }
        });
    }

    private void runProgram() {
        if (!prepareSession()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        final String pwd = editPassword.getText().toString();
        submit("вход в режим программирования", new Op() {
            @Override
            public void run() throws IOException {
                session.enterProgrammingMode(addr, pwd);
            }
        });
    }

    private void runReadRegister() {
        if (!prepareSession()) {
            return;
        }
        final String name = editCommand.getText().toString().trim();
        if (name.isEmpty()) {
            log.line("Введите имя параметра, например ET0PE или 1.8.0");
            return;
        }
        submit("чтение R1 " + name, new Op() {
            @Override
            public void run() throws IOException {
                session.readRegister(name);
            }
        });
    }

    private void runRaw() {
        if (!prepareSession()) {
            return;
        }
        final String text = editCommand.getText().toString();
        submit("сырой обмен", new Op() {
            @Override
            public void run() throws IOException {
                String reply = session.raw(text, 2500);
                log.line("Получено: " + (reply.isEmpty() ? "(ничего)" : reply));
            }
        });
    }

    private void runSignOff() {
        if (!prepareSession()) {
            return;
        }
        submit("выход B0", new Op() {
            @Override
            public void run() throws IOException {
                session.signOff();
            }
        });
    }

    private interface Op {
        void run() throws IOException;
    }

    private void submit(final String title, final Op op) {
        exec.submit(new Runnable() {
            @Override
            public void run() {
                log.line("== " + title + " ==");
                try {
                    op.run();
                    log.line("== готово ==");
                } catch (IOException ex) {
                    log.line("Ошибка: " + ex.getMessage());
                } catch (RuntimeException ex) {
                    log.line("Ошибка: " + ex);
                }
            }
        });
    }

    // ---- misc ---------------------------------------------------------------

    private void shareLog() {
        Intent send = new Intent(Intent.ACTION_SEND);
        send.setType("text/plain");
        send.putExtra(Intent.EXTRA_SUBJECT, "Лог пробника счётчика");
        send.putExtra(Intent.EXTRA_TEXT, log.text());
        try {
            startActivity(Intent.createChooser(send, getString(R.string.btn_share)));
        } catch (RuntimeException ex) {
            Toast.makeText(this, "Нет приложения для отправки", Toast.LENGTH_SHORT).show();
        }
    }

    private ArrayAdapter<String> simpleAdapter(String[] items) {
        ArrayAdapter<String> a = new ArrayAdapter<String>(this, android.R.layout.simple_spinner_item, items);
        a.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        return a;
    }

    private void bind(int id, final Runnable action) {
        findViewById(id).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                action.run();
            }
        });
    }
}
