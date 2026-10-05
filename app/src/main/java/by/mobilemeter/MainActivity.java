package by.mobilemeter;

import android.app.Activity;
import android.app.AlertDialog;
import android.app.PendingIntent;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.IntentFilter;
import android.hardware.usb.UsbManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.View;
import android.widget.AdapterView;
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
import by.mobilemeter.protocol.mirtek.MirtekSession;
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
 * optical head or radio module, and talks to the meter over IEC 62056-21 or МИРТЕК.
 */
public class MainActivity extends Activity {
    private static final String ACTION_USB_PERMISSION = "by.mobilemeter.USB_PERMISSION";
    private static final String[] PROTOCOLS = {"IEC 61107 (Энергомера CE)", "МИРТЕК"};
    private static final String[] BAUDS = {"авто", "300", "600", "1200", "2400", "4800", "9600", "19200"};
    private static final String[] FORMATS = {"авто", "7E1", "8N1"};
    private static final int PROTOCOL_IEC = 0;
    private static final int PROTOCOL_MIRTEK = 1;

    private Spinner spinnerDevice;
    private Spinner spinnerProtocol;
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
    private View groupIec;
    private View groupMirtek;
    private View rowIecPort;

    private final ExecutorService exec = Executors.newSingleThreadExecutor();
    private final Handler ui = new Handler(Looper.getMainLooper());
    private final LogSink log = new LogSink();

    private UsbManager usbManager;
    private List<UsbDevices.Entry> entries = new ArrayList<UsbDevices.Entry>();
    private UsbSerialLink link;
    private Iec62056Session iec;
    private MirtekSession mirtek;
    private int configuredKey = -1;
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
        spinnerProtocol = (Spinner) findViewById(R.id.spinnerProtocol);
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
        groupIec = findViewById(R.id.groupIec);
        groupMirtek = findViewById(R.id.groupMirtek);
        rowIecPort = findViewById(R.id.rowIecPort);

        spinnerProtocol.setAdapter(simpleAdapter(PROTOCOLS));
        spinnerBaud.setAdapter(simpleAdapter(BAUDS));
        spinnerFormat.setAdapter(simpleAdapter(FORMATS));
        spinnerDriver.setAdapter(simpleAdapter(UsbDevices.FORCE_KINDS));
        spinnerProtocol.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                applyProtocolUi(position);
            }

            @Override
            public void onNothingSelected(AdapterView<?> parent) {
            }
        });

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
        bind(R.id.btnShare, new Runnable() { public void run() { shareLog(); } });
        bind(R.id.btnClear, new Runnable() { public void run() { textLog.setText(""); log.clear(); } });
        // IEC 61107
        bind(R.id.btnIdentify, new Runnable() { public void run() { runIdentify(); } });
        bind(R.id.btnReadout, new Runnable() { public void run() { runReadout(); } });
        bind(R.id.btnProgram, new Runnable() { public void run() { runProgram(); } });
        bind(R.id.btnReadReg, new Runnable() { public void run() { runReadRegister(); } });
        bind(R.id.btnRaw, new Runnable() { public void run() { runRaw(); } });
        bind(R.id.btnSignOff, new Runnable() { public void run() { runSignOff(); } });
        // МИРТЕК
        bind(R.id.btnMPing, new Runnable() { public void run() { runMirtekPing(); } });
        bind(R.id.btnMEnergy, new Runnable() { public void run() { runMirtekEnergy(); } });
        bind(R.id.btnMDateTime, new Runnable() { public void run() { runMirtekDateTime(); } });
        bind(R.id.btnMInfo, new Runnable() { public void run() { runMirtekInfo(); } });
        bind(R.id.btnMRelayOn, new Runnable() { public void run() { confirmRelay(false); } });
        bind(R.id.btnMRelayOff, new Runnable() { public void run() { confirmRelay(true); } });
        bind(R.id.btnMRaw, new Runnable() { public void run() { runMirtekRaw(); } });

        IntentFilter filter = new IntentFilter();
        filter.addAction(ACTION_USB_PERMISSION);
        filter.addAction(UsbManager.ACTION_USB_DEVICE_ATTACHED);
        filter.addAction(UsbManager.ACTION_USB_DEVICE_DETACHED);
        if (Build.VERSION.SDK_INT >= 33) {
            registerReceiver(usbReceiver, filter, Context.RECEIVER_EXPORTED);
        } else {
            registerReceiver(usbReceiver, filter);
        }

        log.line("Мобильный контролёр " + BuildInfo.VERSION + ". Android " + Build.VERSION.RELEASE
                + " (API " + Build.VERSION.SDK_INT + "), " + Build.MANUFACTURER + " " + Build.MODEL);
        log.line("Подключите оптоголовку или радиомодуль через OTG, нажмите «Обновить», затем «Подключить».");
        log.line("Энергомера: скорость и формат подбираются автоматически. Миртек: 9600 8N1, нужен адрес счётчика.");
        applyProtocolUi(PROTOCOL_IEC);
        refreshDevices();
    }

    private void applyProtocolUi(int protocol) {
        boolean m = protocol == PROTOCOL_MIRTEK;
        groupIec.setVisibility(m ? View.GONE : View.VISIBLE);
        rowIecPort.setVisibility(m ? View.GONE : View.VISIBLE);
        groupMirtek.setVisibility(m ? View.VISIBLE : View.GONE);
        editAddress.setHint(m ? R.string.hint_address_mirtek : R.string.hint_address_iec);
        editPassword.setHint(m ? R.string.hint_password_mirtek : R.string.hint_password_iec);
        editCommand.setHint(m ? R.string.hint_command_mirtek : R.string.hint_command_iec);
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
                    iec = new Iec62056Session(l, log);
                    mirtek = new MirtekSession(l, log);
                    configuredKey = -1;
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
        iec = null;
        mirtek = null;
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

    private boolean portOpen() {
        if (link == null || !link.isOpen()) {
            log.line("Порт не открыт. Сначала нажмите «Подключить».");
            return false;
        }
        return true;
    }

    // ---- IEC 61107 ----------------------------------------------------------

    private boolean prepareIec() {
        if (!portOpen()) {
            return false;
        }
        int baudPos = spinnerBaud.getSelectedItemPosition();
        int fmtPos = spinnerFormat.getSelectedItemPosition();
        int key = 1000 + baudPos * 100 + fmtPos * 10 + (cbSwitchBaud.isChecked() ? 1 : 0);
        if (key == configuredKey) {
            return true; // settings unchanged: keep the auto-detected parameters
        }
        configuredKey = key;
        mirtek.resetPort();
        int baud = baudPos == 0 ? 0 : Integer.parseInt(BAUDS[baudPos]);
        Boolean sevenE1 = fmtPos == 0 ? null : Boolean.valueOf(fmtPos == 1);
        if (baudPos == 0 || fmtPos == 0) {
            iec.configureAuto(baud, sevenE1, cbSwitchBaud.isChecked());
        } else {
            iec.configure(baud, cbSwitchBaud.isChecked(), sevenE1.booleanValue());
        }
        return true;
    }

    private void runIdentify() {
        if (!prepareIec()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        submit("идентификация", new Op() {
            @Override
            public void run() throws IOException {
                iec.identify(addr);
            }
        });
    }

    private void runReadout() {
        if (!prepareIec()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        submit("считывание показаний", new Op() {
            @Override
            public void run() throws IOException {
                iec.readout(addr);
            }
        });
    }

    private void runProgram() {
        if (!prepareIec()) {
            return;
        }
        final String addr = editAddress.getText().toString();
        final String pwd = editPassword.getText().toString();
        submit("вход в режим программирования", new Op() {
            @Override
            public void run() throws IOException {
                iec.enterProgrammingMode(addr, pwd);
            }
        });
    }

    private void runReadRegister() {
        if (!prepareIec()) {
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
                iec.readRegister(name);
            }
        });
    }

    private void runRaw() {
        if (!prepareIec()) {
            return;
        }
        final String text = editCommand.getText().toString();
        submit("сырой обмен", new Op() {
            @Override
            public void run() throws IOException {
                String reply = iec.raw(text, 2500);
                log.line("Получено: " + (reply.isEmpty() ? "(ничего)" : reply));
            }
        });
    }

    private void runSignOff() {
        if (!prepareIec()) {
            return;
        }
        submit("выход B0", new Op() {
            @Override
            public void run() throws IOException {
                iec.signOff();
            }
        });
    }

    // ---- МИРТЕК -------------------------------------------------------------

    private boolean prepareMirtek() {
        if (!portOpen()) {
            return false;
        }
        String a = editAddress.getText().toString().trim();
        String p = editPassword.getText().toString().trim();
        int address;
        long password;
        try {
            address = a.isEmpty() ? 0 : Integer.parseInt(a);
            password = p.isEmpty() ? 0 : Long.parseLong(p);
        } catch (NumberFormatException ex) {
            log.line("Для Миртека адрес и пароль должны быть числами (адрес 0…65535, пароль 0…4294967295)");
            return false;
        }
        if (address < 0 || address > 0xFFFF || password < 0 || password > 0xFFFFFFFFL) {
            log.line("Адрес должен быть 0…65535, пароль 0…4294967295");
            return false;
        }
        int key = 2000;
        if (key != configuredKey) {
            configuredKey = key;
            mirtek.resetPort();
        }
        mirtek.configure(address, password);
        return true;
    }

    private void runMirtekPing() {
        if (!prepareMirtek()) {
            return;
        }
        submit("Миртек: Ping", new Op() {
            @Override
            public void run() throws IOException {
                mirtek.ping();
            }
        });
    }

    private void runMirtekEnergy() {
        if (!prepareMirtek()) {
            return;
        }
        submit("Миртек: показания", new Op() {
            @Override
            public void run() throws IOException {
                mirtek.readEnergy();
            }
        });
    }

    private void runMirtekDateTime() {
        if (!prepareMirtek()) {
            return;
        }
        submit("Миртек: дата и время", new Op() {
            @Override
            public void run() throws IOException {
                mirtek.readDateTime();
            }
        });
    }

    private void runMirtekInfo() {
        if (!prepareMirtek()) {
            return;
        }
        submit("Миртек: информация о счётчике", new Op() {
            @Override
            public void run() throws IOException {
                mirtek.ping();
                mirtek.getInfo();
                mirtek.readConfigure();
                for (int i = 1; i <= 4; i++) {
                    try {
                        mirtek.readFactoryString(i);
                    } catch (IOException ex) {
                        log.line("Поле " + i + ": " + ex.getMessage());
                    }
                }
            }
        });
    }

    private void confirmRelay(final boolean disconnect) {
        if (!prepareMirtek()) {
            return;
        }
        final int address = Integer.parseInt(editAddress.getText().toString().trim().isEmpty() ? "0"
                : editAddress.getText().toString().trim());
        String text = getString(disconnect ? R.string.relay_off_text : R.string.relay_on_text, address);
        new AlertDialog.Builder(this)
                .setTitle(R.string.relay_confirm_title)
                .setMessage(text)
                .setPositiveButton(R.string.yes, new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        submit(disconnect ? "Миртек: ОТКЛЮЧЕНИЕ нагрузки" : "Миртек: включение нагрузки", new Op() {
                            @Override
                            public void run() throws IOException {
                                mirtek.relay(0, disconnect);
                            }
                        });
                    }
                })
                .setNegativeButton(R.string.no, null)
                .show();
    }

    private void runMirtekRaw() {
        if (!prepareMirtek()) {
            return;
        }
        final byte[] bytes;
        try {
            bytes = parseHex(editCommand.getText().toString());
        } catch (IllegalArgumentException ex) {
            log.line("Команда задаётся в HEX: первый байт код команды, далее данные, например «2B 00»");
            return;
        }
        if (bytes.length == 0) {
            log.line("Введите код команды в HEX, например «01» для Ping или «2B 00» для мгновенных значений");
            return;
        }
        final byte[] data = new byte[bytes.length - 1];
        System.arraycopy(bytes, 1, data, 0, data.length);
        final int cmd = bytes[0] & 0xFF;
        submit(String.format("Миртек: команда 0x%02X", cmd), new Op() {
            @Override
            public void run() throws IOException {
                mirtek.raw(cmd, data);
            }
        });
    }

    static byte[] parseHex(String text) {
        String s = text.replaceAll("[^0-9A-Fa-f]", "");
        if (s.length() % 2 != 0) {
            throw new IllegalArgumentException("odd hex length");
        }
        byte[] out = new byte[s.length() / 2];
        for (int i = 0; i < out.length; i++) {
            out[i] = (byte) Integer.parseInt(s.substring(2 * i, 2 * i + 2), 16);
        }
        return out;
    }

    // ---- common -------------------------------------------------------------

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
                Iec62056Session s = iec;
                UsbSerialLink l = link;
                if (s != null && l != null && s.lockedDescription() != null) {
                    setStatus("Подключено: " + l.describe() + " | " + s.lockedDescription(), true);
                }
            }
        });
    }

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
