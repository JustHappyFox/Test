# Фото с телефона: фото и скрины с телефона (Android, iPhone) по кабелю переносятся (вырезаются) в папки год\месяц\число.
# Запуск: zapusk.vbs (без чёрного окна) или zapusk.bat.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @([System.Windows.Forms.Form].Assembly.Location, [System.Drawing.Point].Assembly.Location)

$code = @'
using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Forms;
using System;

// Ошибки программы: вместо стандартного окна Windows - своё, с подробностями (их можно скопировать),
// и запись в oshibki.txt рядом с программой - его можно прислать. Номера (лицевые и т.п.) и имена файлов
// фото в записи скрыты.
public static class Crash {
    static string log;
    static bool showing;

    public static string LogFile { get { return log; } }

    public static void Install(string here) {
        log = Path.Combine(here, "oshibki.txt");
        try { Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException); } catch { }
        Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e) { Show(e.Exception, true); };
        AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e) {
            Exception ex = e.ExceptionObject as Exception;
            Write(ex != null ? ex.ToString() : "" + e.ExceptionObject, "в фоне");
        };
    }

    // что за компьютер: Windows, .NET, разрядность, память - без данных абонентов
    public static string Environment2() {
        StringBuilder sb = new StringBuilder();
        sb.Append("Windows ").Append(Environment.OSVersion.Version).Append(" ").Append(Environment.OSVersion.ServicePack);
        sb.Append(", .NET ").Append(Environment.Version).Append(", ").Append(IntPtr.Size == 8 ? "64" : "32").Append("-битный процесс");
        try {
            Process p = Process.GetCurrentProcess();
            sb.Append(", память программы ").Append(p.WorkingSet64 / 1048576).Append(" МБ (пик ").Append(p.PeakWorkingSet64 / 1048576).Append(" МБ)");
            sb.Append(", объектов GDI ").Append(GdiCount(p)).Append(", окон USER ").Append(UserCount(p));
        } catch { }
        return sb.ToString();
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern int GetGuiResources(IntPtr process, int flags);

    static string GdiCount(Process p) { try { return GetGuiResources(p.Handle, 0).ToString(); } catch { return "?"; } }
    static string UserCount(Process p) { try { return GetGuiResources(p.Handle, 1).ToString(); } catch { return "?"; } }

    // скрыть данные: длинные номера и имена файлов фото
    public static string Clean(string s) {
        s = Regex.Replace(s ?? "", @"\d{6,}", "######");
        s = Regex.Replace(s, @"[^\\/\r\n""':]+\.(jpe?g|png|heic|heif|mp4|mov|dng|webp)", "…", RegexOptions.IgnoreCase);
        return s;
    }

    public static void Write(string text, string where) {
        try {
            File.AppendAllText(log, "=== " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " (" + where + ") ===\r\n" + Environment2() + "\r\n" + Clean(text) + "\r\n\r\n", new UTF8Encoding(true));
        } catch { }
    }

    static Button Btn(string text) { Button b = new Button(); b.Text = text; return b; }

    public static void Show(Exception ex, bool canContinue) {
        Write(ex.ToString(), "окно");
        if (showing) return;
        showing = true;
        try {
            using (Form f = new Form()) {
                f.Text = "Фото с телефона - ошибка";
                f.Font = new Font("Segoe UI", 10f);
                f.BackColor = Color.White;
                f.StartPosition = FormStartPosition.CenterScreen;
                f.Size = new Size(760, 480);
                f.TopMost = true;
                Label l = new Label();
                l.Text = "Произошла ошибка: " + Clean(ex.Message) + "\n\nПодробности записаны в файл " + log +
                    " - пришлите его (или фото этого окна). Имена фото в нём скрыты.";
                l.Font = f.Font;
                l.SetBounds(16, 12, 712, 86);
                TextBox t = new TextBox();
                t.Multiline = true;
                t.ReadOnly = true;
                t.ScrollBars = ScrollBars.Both;
                t.WordWrap = false;
                t.Font = new Font("Consolas", 8.5f);
                t.Text = Environment2() + "\r\n\r\n" + Clean(ex.ToString());
                t.SetBounds(16, 104, 712, 268);
                Button copy = Btn("Скопировать");
                copy.SetBounds(16, 384, 140, 36);
                copy.Click += delegate { try { Clipboard.SetText(t.Text); } catch { } };
                Button go = Btn(canContinue ? "Продолжить" : "Закрыть");
                go.SetBounds(canContinue ? 432 : 588, 384, 140, 36);
                go.DialogResult = DialogResult.OK;
                Button quit = Btn("Закрыть программу");
                quit.SetBounds(580, 384, 148, 36);
                quit.Click += delegate { Environment.Exit(1); };
                f.Controls.AddRange(new Control[] { l, t, copy, go });
                if (canContinue) f.Controls.Add(quit);
                f.AcceptButton = go;
                f.ShowDialog();
            }
        } catch {
            try { MessageBox.Show("Ошибка: " + Clean(ex.Message) + "\n\nПодробности - в " + log, "Фото с телефона"); } catch { }
        } finally { showing = false; }
    }
}

// Телефон по кабелю (Android - «Передача файлов», iPhone - «Доверять этому компьютеру») виден в «Этот компьютер» как
// устройство, а не как диск: обычными файловыми функциями его не прочитать. Работаем через оболочку Windows
// (Shell.Application, как Проводник), в отдельном потоке STA. Фото ищутся в DCIM (Camera, Screenshots, у iPhone - 100APPLE...)
// и в Pictures\Screenshots; дата - «Изменён» (как в Проводнике).

public delegate object ShellMaker();
public delegate void TextHandler(string text);

// файл на телефоне
public class PhoneFile {
    public string Device = "";       // «Redmi 8A», «Apple iPhone»
    public string DevicePath = "";   // путь устройства в оболочке (два одинаковых телефона различаются по нему)
    public string Folder = "";       // путь в телефоне: «Внутренний общий накопитель\DCIM\Camera»
    public string Name = "";         // имя с расширением
    public DateTime Date = DateTime.MinValue;   // по какой дате раскладывать (выбор в окне); MinValue - неизвестно
    public DateTime Modified = DateTime.MinValue;   // «Изменён» по телефону (как в Проводнике)
    public DateTime Taken = DateTime.MinValue;      // дата съёмки (из имени файла или свойства «Дата съёмки»)
    public string TakenFrom = "";                   // имя, свойство
    public long Size;
    public string Kind = "";         // фото, скрин, видео
    public string Source { get { return Device + ": " + Folder; } }
}

// подключённый телефон: имя, путь в оболочке, номер устройства (серийный номер USB; IMEI Windows не видит)
public class PhoneDevice {
    public string Name = "", Path = "", Id = "";
}

public class ScanResult {
    public string Error;
    public readonly List<string> Devices = new List<string>();      // найденные телефоны
    public readonly List<string> DeviceInfo = new List<string>();   // для отчёта: имя и номер устройства
    public readonly List<string> Sources = new List<string>();      // папки с фото («Redmi 8A: ...\DCIM\Camera»)
    public readonly List<PhoneFile> Files = new List<PhoneFile>();
}

public class MoveResult {
    public string Error;
    public bool Stopped;
    public int Duplicates;                                          // такой же файл уже был в папке - второй не кладётся
    public readonly List<PhoneFile> Done = new List<PhoneFile>();   // ушли с телефона и лежат в папках
    public readonly List<string> Placed = new List<string>();       // где лежат (полный путь)
    public readonly List<PhoneFile> Failed = new List<PhoneFile>(); // не перенеслись (остались на телефоне)
    public readonly List<PhoneFile> NotRemoved = new List<PhoneFile>();  // скопированы, но на телефоне остались
}

// перенос можно остановить (после текущего файла)
public class StopFlag { public volatile bool Stop; }

public static class PhoneShell {
    // оболочку можно подменить (проверка программы без телефона)
    public static ShellMaker MakeShell = delegate { return Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application")); };
    public static int StartSeconds = 90;     // сколько ждать, пока файл начнёт появляться
    public const string StagingName = "_перенос с телефона";

    static readonly string[] PhotoExt = { ".jpg", ".jpeg", ".png", ".heic", ".heif", ".dng", ".webp", ".gif", ".bmp", ".tif", ".tiff" };
    static readonly string[] VideoExt = { ".mp4", ".mov", ".3gp", ".m4v", ".avi", ".mkv", ".mts", ".hevc" };

    // фото, скрин, видео или null (не фото: .nomedia, .thumbnails и т.п.)
    public static string KindOf(string name, string folder) {
        string e = Path.GetExtension(name).ToLowerInvariant();
        bool screen = name.ToLowerInvariant().Contains("screenshot") || folder.ToLowerInvariant().Contains("screenshot");
        if (Array.IndexOf(PhotoExt, e) >= 0) return screen ? "скрин" : "фото";
        if (Array.IndexOf(VideoExt, e) >= 0) return "видео";
        return null;
    }

    static object Call(object o, string m, params object[] a) { return o.GetType().InvokeMember(m, System.Reflection.BindingFlags.InvokeMethod, null, o, a); }
    static object Prop(object o, string p) { return o.GetType().InvokeMember(p, System.Reflection.BindingFlags.GetProperty, null, o, null); }

    static List<object> Items(object folder) {
        List<object> r = new List<object>();
        if (folder == null) return r;
        object items = Call(folder, "Items");
        if (items == null) return r;
        int n = Convert.ToInt32(Prop(items, "Count"));
        for (int i = 0; i < n; i++) {
            object it = Call(items, "Item", i);
            if (it != null) r.Add(it);
        }
        return r;
    }

    static string Str(object o, string p) { try { object v = Prop(o, p); return v == null ? "" : v.ToString(); } catch { return ""; } }
    static bool Flag(object o, string p) { try { return Convert.ToBoolean(Prop(o, p)); } catch { return false; } }
    static object Sub(object item) { try { return Prop(item, "GetFolder"); } catch { return null; } }

    static object Ext(object item, string key) { try { return Call(item, "ExtendedProperty", key); } catch { return null; } }

    // имя файла с расширением (Проводник может прятать расширения)
    static string FileName(object item) {
        object o = Ext(item, "System.FileName");
        string n = o == null ? "" : o.ToString();
        if (n.Length == 0) n = Str(item, "Name");
        if (Path.GetExtension(n).Length == 0) {
            string t = Str(item, "Type").ToUpperInvariant();
            foreach (string e in new string[] { "JPG", "JPEG", "PNG", "HEIC", "MP4", "MOV" })
                if (t.Contains(e)) { n += "." + e.ToLowerInvariant(); break; }
        }
        return n;
    }

    // раскладывать по дате съёмки (как в галерее телефона), иначе - по «Изменён» (как в Проводнике)
    public static bool ByTaken;

    static DateTime Local(object d) {
        if (!(d is DateTime) || ((DateTime)d).Year <= 1990) return DateTime.MinValue;
        DateTime t = (DateTime)d;
        return t.Kind == DateTimeKind.Utc ? t.ToLocalTime() : t;
    }

    // «Изменён» (как в Проводнике); нет - из свойств файла
    static DateTime Modified(object item) {
        try {
            DateTime t = Local(Prop(item, "ModifyDate"));
            if (t != DateTime.MinValue) return t;
        } catch { }
        return Local(Ext(item, "System.DateModified"));
    }

    // дата съёмки: из имени (IMG_20261103_154512.jpg, Screenshot_2026-11-03-15-45-12-345_...png, IMG-20261103-WA0001.jpg);
    // в имени даты нет (iPhone: IMG_0001) - свойство «Дата съёмки» (у видео - «Дата записи»)
    static DateTime Taken(object item, string name, out string from) {
        from = "имя";
        DateTime t = NameDate(name);
        if (t != DateTime.MinValue) return t;
        from = "свойство";
        foreach (string k in new string[] { "System.Photo.DateTaken", "System.Media.DateEncoded" }) {
            t = Local(Ext(item, k));
            if (t != DateTime.MinValue) return t;
        }
        from = "";
        return DateTime.MinValue;
    }

    public static DateTime NameDate(string name) {
        Match m = Regex.Match(name ?? "", @"(?<!\d)(20\d\d)[-_.]?(\d\d)[-_.]?(\d\d)(?:[-_ .T]?(\d\d)[-_.]?(\d\d)[-_.]?(\d\d))?(?!\d)");
        if (!m.Success) return DateTime.MinValue;
        try {
            DateTime d = new DateTime(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value));
            if (d > DateTime.Today.AddDays(2)) return DateTime.MinValue;
            if (m.Groups[4].Success && m.Groups[4].Value.Length > 0) {
                int h = int.Parse(m.Groups[4].Value), mi = int.Parse(m.Groups[5].Value), se = int.Parse(m.Groups[6].Value);
                if (h < 24 && mi < 60 && se < 60) d = d.AddHours(h).AddMinutes(mi).AddSeconds(se);
            }
            return d;
        } catch { return DateTime.MinValue; }
    }

    // по какой дате раскладывать: выбранная, нет её - другая
    public static void UseDate(PhoneFile f) {
        DateTime a = ByTaken ? f.Taken : f.Modified, b = ByTaken ? f.Modified : f.Taken;
        f.Date = a != DateTime.MinValue ? a : b;
    }

    static object Named(List<object> items, string name) {
        foreach (object it in items) if (string.Compare(Str(it, "Name"), name, StringComparison.OrdinalIgnoreCase) == 0 && Flag(it, "IsFolder")) return it;
        return null;
    }

    // телефоны: в «Этот компьютер» - не диски, а устройства
    static List<object> Devices(object shell) {
        List<object> r = new List<object>();
        object computer = Call(shell, "NameSpace", 17);
        foreach (object it in Items(computer))
            if (!Flag(it, "IsFileSystem") && Flag(it, "IsFolder")) r.Add(it);
        return r;
    }

    // какие телефоны подключены (быстро: только «Этот компьютер», без фото)
    public static List<PhoneDevice> ListDevices() {
        List<PhoneDevice> r = new List<PhoneDevice>();
        object shell = MakeShell();
        foreach (object d in Devices(shell)) r.Add(Describe(d));
        return r;
    }

    static PhoneDevice Describe(object d) {
        PhoneDevice p = new PhoneDevice();
        p.Name = Str(d, "Name");
        p.Path = Str(d, "Path");
        p.Id = DeviceId(p.Path, p.Name);
        return p;
    }

    // Номер устройства - серийный номер USB (у Android - серийный номер телефона, у iPhone - UDID): IMEI по кабелю Windows
    // не видит. Путь устройства в оболочке: ...\\?\usb#vid_2717&pid_ff48&mi_00#7&1a2b3c4d&0&0000#{6ac27878-...}; у составного
    // устройства (mi_00) серийный номер - имя раздела родителя в реестре Enum\USB\VID_2717&PID_FF48 с тем же ParentIdPrefix.
    // Не нашёлся - по имени телефона.
    public static string DeviceId(string path, string name) {
        try {
            Match m = Regex.Match(path ?? "", @"usb#(vid_[0-9a-f]{4})&(pid_[0-9a-f]{4})(&mi_[0-9a-f]{2})?#([^#\\]+)#", RegexOptions.IgnoreCase);
            if (m.Success) {
                string vid = m.Groups[1].Value.ToUpperInvariant(), pid = m.Groups[2].Value.ToUpperInvariant(), inst = m.Groups[4].Value;
                if (m.Groups[3].Success && m.Groups[3].Value.Length > 0) {
                    string serial = ParentSerial(vid + "&" + pid, inst);
                    if (serial != null) return vid.Substring(4) + "-" + serial.ToUpperInvariant();
                } else if (inst.IndexOf('&') < 0) return vid.Substring(4) + "-" + inst.ToUpperInvariant();
            }
        } catch { }
        return "имя:" + name;
    }

    static string ParentSerial(string vidPid, string inst) {
        int k = inst.LastIndexOf('&');
        if (k <= 0) return null;
        string prefix = inst.Substring(0, k);
        using (RegistryKey usb = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\USB\\" + vidPid)) {
            if (usb == null) return null;
            foreach (string sub in usb.GetSubKeyNames()) {
                using (RegistryKey k2 = usb.OpenSubKey(sub)) {
                    object v = k2 == null ? null : k2.GetValue("ParentIdPrefix");
                    if (v != null && string.Compare(v.ToString(), prefix, StringComparison.OrdinalIgnoreCase) == 0) return sub;
                }
            }
        }
        return null;
    }

    // папки с фото на устройстве: [путь, папка]. DCIM - на самом устройстве или в его памяти («Внутренний общий накопитель»,
    // «Internal Storage», карта); в DCIM - все папки, кроме скрытых (.thumbnails): Camera, Screenshots, 100APPLE...;
    // и Pictures\Screenshots (скрины на многих Android)
    static List<KeyValuePair<string, object>> PhotoFolders(object dev) {
        List<KeyValuePair<string, object>> r = new List<KeyValuePair<string, object>>();
        List<KeyValuePair<string, object>> roots = new List<KeyValuePair<string, object>>();
        roots.Add(new KeyValuePair<string, object>("", Sub(dev)));
        foreach (object st in Items(Sub(dev)))
            if (Flag(st, "IsFolder") && string.Compare(Str(st, "Name"), "DCIM", StringComparison.OrdinalIgnoreCase) != 0)
                roots.Add(new KeyValuePair<string, object>(Str(st, "Name"), Sub(st)));
        foreach (KeyValuePair<string, object> root in roots) {
            string pre = root.Key.Length > 0 ? root.Key + "\\" : "";
            List<object> top = Items(root.Value);
            object dcim = Named(top, "DCIM");
            if (dcim != null) {
                r.Add(new KeyValuePair<string, object>(pre + "DCIM", Sub(dcim)));
                foreach (object sub in Items(Sub(dcim)))
                    if (Flag(sub, "IsFolder") && !Str(sub, "Name").StartsWith("."))
                        r.Add(new KeyValuePair<string, object>(pre + "DCIM\\" + Str(sub, "Name"), Sub(sub)));
            }
            object pics = Named(top, "Pictures");
            if (pics != null) {
                object shots = Named(Items(Sub(pics)), "Screenshots");
                if (shots != null) r.Add(new KeyValuePair<string, object>(pre + "Pictures\\Screenshots", Sub(shots)));
            }
        }
        return r;
    }

    static Exception Inner(Exception ex) { return ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex; }

    // какие фото есть на подключённых телефонах
    public static ScanResult Scan(TextHandler progress) { return Scan(progress, null); }

    // onlyPath - только этот телефон (путь устройства)
    public static ScanResult Scan(TextHandler progress, string onlyPath) {
        ScanResult r = new ScanResult();
        try {
            object shell = MakeShell();
            List<object> devs = Devices(shell);
            if (onlyPath != null) devs = devs.FindAll(delegate(object d) { return Str(d, "Path") == onlyPath; });
            if (devs.Count == 0) {
                r.Error = "Телефон не найден. Подключите его кабелем и разблокируйте. Android: в уведомлении «USB» выберите «Передача файлов». " +
                    "iPhone: на вопрос «Доверять этому компьютеру?» - «Доверять».";
                return r;
            }
            foreach (object dev in devs) {
                string devName = Str(dev, "Name"), devPath = Str(dev, "Path");
                r.Devices.Add(devName);
                r.DeviceInfo.Add(devName + " - номер устройства " + DeviceId(devPath, devName));
                if (progress != null) progress("Телефон «" + devName + "»: ищу папки с фото...");
                foreach (KeyValuePair<string, object> f in PhotoFolders(dev)) {
                    int n = 0;
                    List<object> items = Items(f.Value);
                    foreach (object it in items) {
                        if (Flag(it, "IsFolder")) continue;
                        string name = FileName(it);
                        string kind = KindOf(name, f.Key);
                        if (kind == null) continue;
                        PhoneFile p = new PhoneFile();
                        p.Device = devName;
                        p.DevicePath = devPath;
                        p.Folder = f.Key;
                        p.Name = name;
                        p.Kind = kind;
                        p.Modified = Modified(it);
                        p.Taken = Taken(it, name, out p.TakenFrom);
                        UseDate(p);
                        try { p.Size = Convert.ToInt64(Prop(it, "Size")); } catch { p.Size = 0; }
                        r.Files.Add(p);
                        n++;
                        if (progress != null && n % 100 == 0) progress("«" + devName + "»: " + f.Key + " - просмотрено " + n + " из " + items.Count + "...");
                    }
                    if (n > 0) r.Sources.Add(devName + ": " + f.Key);
                }
            }
            if (r.Sources.Count == 0)
                r.Error = "Телефон «" + string.Join("», «", r.Devices.ToArray()) + "» виден, но фото на нём не открываются (или их нет). " +
                    "Разблокируйте телефон: Android - выберите «Передача файлов» и «Разрешить доступ к данным»; iPhone - «Доверять».";
        } catch (Exception ex) {
            r.Error = "Не получилось посмотреть телефон: " + Inner(ex).Message;
        }
        return r;
    }

    // папка в телефоне по пути (заново: объекты оболочки живут только в своём потоке)
    static object FolderOf(object shell, PhoneFile f) {
        object dev = null;
        foreach (object d in Devices(shell)) if (f.DevicePath.Length > 0 ? Str(d, "Path") == f.DevicePath : Str(d, "Name") == f.Device) { dev = d; break; }
        if (dev == null) return null;
        string path = f.Folder;
        object cur = Sub(dev);
        foreach (string part in path.Split('\\')) {
            if (part.Length == 0) continue;
            object next = Named(Items(cur), part);
            if (next == null) return null;
            cur = Sub(next);
        }
        return cur;
    }

    static Dictionary<string, object> ByName(object folder) {
        Dictionary<string, object> d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (object it in Items(folder)) {
            if (Flag(it, "IsFolder")) continue;
            string n = FileName(it);
            if (!d.ContainsKey(n)) d[n] = it;
        }
        return d;
    }

    // папка дня: root\2026\08\21; без даты - root\без даты
    public static string DayFolder(string root, DateTime d) {
        if (d == DateTime.MinValue) return Path.Combine(root, "без даты");
        return Path.Combine(Path.Combine(Path.Combine(root, d.ToString("yyyy")), d.ToString("MM")), d.ToString("dd"));
    }

    // Перенести (вырезать) файлы с телефона в папки дней под root. Каждый файл сначала приходит в промежуточную папку
    // (там пусто - ничего не перезапишется), потом кладётся в папку своего дня: такое имя уже есть - тот же файл
    // (не кладётся второй раз) или другое фото - с номером «IMG_0001 (2).jpg». В конце - проверка, что с телефона ушло.
    public static MoveResult Move(List<PhoneFile> files, string root, StopFlag stop, TextHandler progress) {
        MoveResult r = new MoveResult();
        try {
            Directory.CreateDirectory(root);
            string staging = Staging(root);
            object shell = MakeShell();
            object target = Call(shell, "NameSpace", staging);
            if (target == null) { r.Error = "Не открывается папка: " + staging; return r; }
            List<PhoneFile> order = new List<PhoneFile>(files);
            order.Sort(delegate(PhoneFile a, PhoneFile b) {
                int c = string.Compare(a.Source, b.Source, StringComparison.Ordinal);
                if (c != 0) return c;
                c = a.Date.CompareTo(b.Date);
                return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
            });
            string curSource = null;
            Dictionary<string, object> items = null;
            List<string> touched = new List<string>();
            for (int i = 0; i < order.Count; i++) {
                PhoneFile f = order[i];
                if (stop != null && stop.Stop) { r.Stopped = true; break; }
                if (f.Source != curSource) {
                    curSource = f.Source;
                    if (progress != null) progress("Открываю на телефоне " + f.Source + "...");
                    object folder = FolderOf(shell, f);
                    items = folder == null ? new Dictionary<string, object>() : ByName(folder);
                    touched.Add(f.Source);
                }
                object item;
                if (!items.TryGetValue(f.Name, out item)) { r.Failed.Add(f); continue; }
                if (progress != null) progress("Переношу " + (i + 1) + " из " + order.Count + ": " + f.Name + " (" + Mb(f.Size) + ")");
                Dictionary<string, bool> before = Snapshot(staging);
                try { Call(target, "MoveHere", item, 4 | 16 | 512 | 1024); }   // без вопросов и окон
                catch { r.Failed.Add(f); continue; }
                string got = WaitNew(staging, before, f);
                if (got == null) { r.Failed.Add(f); continue; }
                try {
                    string placed = Place(got, f.Date, f.Modified, root, r);
                    r.Done.Add(f);
                    if (placed != null) r.Placed.Add(placed);
                } catch { r.Failed.Add(f); }
            }
            // ушло ли с телефона (это перенос, Ctrl+X): осталось - файл уже лежит на компьютере, удаляем его с телефона сами,
            // без вопросов; не удалился и так - телефон запрещает удаление (r.NotRemoved)
            if (progress != null && r.Done.Count > 0) progress("Проверяю, что перенесённого на телефоне больше нет...");
            foreach (string src in touched) {
                PhoneFile any = null;
                foreach (PhoneFile f in r.Done) if (f.Source == src) { any = f; break; }
                if (any == null) continue;
                object folder = FolderOf(shell, any);
                if (folder == null) continue;
                Dictionary<string, object> left = ByName(folder);
                bool deleted = false;
                foreach (PhoneFile f in r.Done) {
                    object item;
                    if (f.Source != src || !left.TryGetValue(f.Name, out item)) continue;
                    if (progress != null) progress("Удаляю с телефона перенесённое: " + f.Name);
                    DeleteItem(item);
                    deleted = true;
                }
                if (!deleted) continue;
                System.Threading.Thread.Sleep(500);
                folder = FolderOf(shell, any);
                if (folder == null) continue;
                left = ByName(folder);
                foreach (PhoneFile f in r.Done) if (f.Source == src && left.ContainsKey(f.Name)) r.NotRemoved.Add(f);
            }
            TryDeleteDir(staging);
        } catch (Exception ex) {
            r.Error = "Перенос прервался: " + Inner(ex).Message;
        }
        return r;
    }

    // удалить файл с телефона без вопросов: как Проводник (IFileOperation, без подтверждения); не вышло - командой «Удалить»
    static void DeleteItem(object item) {
        try {
            string path = Str(item, "Path");
            if (path.Length > 0 && ShellDelete.Delete(path)) return;
        } catch { }
        try { Call(item, "InvokeVerb", "delete"); } catch { }
    }

    // отчёт: что программа видит на телефоне - папки, сколько файлов, по каким дням (без имён файлов)
    public static string Report(ScanResult r) {
        StringBuilder sb = new StringBuilder();
        sb.Append("Отчёт программы «Фото с телефона» ").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append(" (имён файлов здесь нет)\r\n");
        sb.Append("Раскладывать по: ").Append(ByTaken ? "дате съёмки" : "«Изменён»").Append("\r\n");
        foreach (string d in r.DeviceInfo) sb.Append("Телефон: ").Append(d).Append("\r\n");
        if (r.Error != null) sb.Append("Ошибка: ").Append(r.Error).Append("\r\n");
        foreach (string src in r.Sources) {
            int n = 0, mod = 0, byName = 0, byProp = 0, diff = 0;
            Dictionary<string, int> kinds = new Dictionary<string, int>(), ext = new Dictionary<string, int>();
            Dictionary<DateTime, int> dm = new Dictionary<DateTime, int>(), dt = new Dictionary<DateTime, int>();
            Dictionary<string, int> pairs = new Dictionary<string, int>();
            foreach (PhoneFile f in r.Files) {
                if (f.Source != src) continue;
                n++;
                Count(kinds, f.Kind);
                Count(ext, Path.GetExtension(f.Name).ToLowerInvariant());
                if (f.Modified != DateTime.MinValue) { mod++; Count(dm, f.Modified.Date); }
                if (f.TakenFrom == "имя") byName++;
                if (f.TakenFrom == "свойство") byProp++;
                if (f.Taken != DateTime.MinValue) Count(dt, f.Taken.Date);
                if (f.Modified != DateTime.MinValue && f.Taken != DateTime.MinValue && f.Modified.Date != f.Taken.Date) {
                    diff++;
                    Count(pairs, "снято " + f.Taken.ToString("dd.MM.yyyy") + " -> «Изменён» " + f.Modified.ToString("dd.MM.yyyy"));
                }
            }
            sb.Append("\r\n").Append(src).Append(": файлов ").Append(n).Append(" (").Append(Join(kinds)).Append("; ").Append(Join(ext)).Append(")\r\n");
            sb.Append("  «Изменён» известен у ").Append(mod).Append("; дата съёмки - из имени у ").Append(byName).Append(", из свойства у ").Append(byProp)
              .Append("; дни разные у ").Append(diff).Append("\r\n");
            sb.Append("  по «Изменён»: ").Append(Days(dm)).Append("\r\n");
            sb.Append("  по дате съёмки: ").Append(Days(dt)).Append("\r\n");
            if (pairs.Count > 0) {
                List<string> l = new List<string>();
                foreach (KeyValuePair<string, int> kv in pairs) l.Add(kv.Key + ": " + kv.Value);
                l.Sort();
                if (l.Count > 15) l.RemoveRange(15, l.Count - 15);
                sb.Append("  разные дни: ").Append(string.Join("; ", l.ToArray())).Append("\r\n");
            }
        }
        return sb.ToString();
    }

    static void Count<T>(Dictionary<T, int> d, T k) { int v; d.TryGetValue(k, out v); d[k] = v + 1; }

    static string Join(Dictionary<string, int> d) {
        List<string> l = new List<string>();
        foreach (KeyValuePair<string, int> kv in d) l.Add((kv.Key.Length > 0 ? kv.Key : "без расширения") + " " + kv.Value);
        return string.Join(", ", l.ToArray());
    }

    // последние 40 дней, где есть файлы: «04.11.2026 - 25»
    static string Days(Dictionary<DateTime, int> d) {
        List<DateTime> k = new List<DateTime>(d.Keys);
        k.Sort();
        k.Reverse();
        List<string> l = new List<string>();
        for (int i = 0; i < k.Count && i < 40; i++) l.Add(k[i].ToString("dd.MM.yyyy") + " - " + d[k[i]]);
        if (k.Count > 40) l.Add("... ещё дней " + (k.Count - 40));
        return l.Count == 0 ? "нет" : string.Join(", ", l.ToArray());
    }

    public static string Mb(long size) { return size <= 0 ? "?" : (size / 1048576.0).ToString(size < 10485760 ? "0.0" : "0") + " МБ"; }

    public static string Staging(string root) {
        string s = Path.Combine(root, StagingName);
        if (!Directory.Exists(s)) {
            DirectoryInfo d = Directory.CreateDirectory(s);
            try { d.Attributes |= FileAttributes.Hidden; } catch { }
        }
        return s;
    }

    static void TryDeleteDir(string d) { try { if (Directory.GetFileSystemEntries(d).Length == 0) Directory.Delete(d); } catch { } }

    // пришедший файл - в папку дня date; дата «Изменён» - как на телефоне. Возвращает путь (null - такой же файл уже был)
    static string Place(string got, DateTime date, DateTime modified, string root, MoveResult r) {
        DateTime m = modified != DateTime.MinValue ? modified : date;
        if (m != DateTime.MinValue) { try { File.SetLastWriteTime(got, m); } catch { } }
        string dir = DayFolder(root, date);
        Directory.CreateDirectory(dir);
        string name = Path.GetFileName(got), dst = Path.Combine(dir, name);
        if (File.Exists(dst)) {
            if (SameFile(got, dst)) { File.Delete(got); r.Duplicates++; return null; }
            string b = Path.GetFileNameWithoutExtension(name), e = Path.GetExtension(name);
            for (int k = 2; File.Exists(dst); k++) dst = Path.Combine(dir, b + " (" + k + ")" + e);
        }
        File.Move(got, dst);
        return dst;
    }

    static bool SameFile(string a, string b) {
        FileInfo fa = new FileInfo(a), fb = new FileInfo(b);
        if (fa.Length != fb.Length) return false;
        using (FileStream x = File.OpenRead(a)) using (FileStream y = File.OpenRead(b)) {
            byte[] p = new byte[65536], q = new byte[65536];
            while (true) {
                int n = x.Read(p, 0, p.Length), m = y.Read(q, 0, q.Length);
                if (n != m) return false;
                if (n == 0) return true;
                for (int i = 0; i < n; i++) if (p[i] != q[i]) return false;
            }
        }
    }

    // остатки прерванного переноса (свет погас, кабель выдернули) - разложить по датам «Изменён». Сколько разложено
    public static int Recover(string root) {
        string s = Path.Combine(root, StagingName);
        if (!Directory.Exists(s)) return 0;
        int n = 0;
        MoveResult r = new MoveResult();
        foreach (string f in Directory.GetFiles(s)) {
            try { DateTime t = File.GetLastWriteTime(f); Place(f, t, t, root, r); n++; } catch { }
        }
        TryDeleteDir(s);
        return n;
    }

    static Dictionary<string, bool> Snapshot(string dir) {
        Dictionary<string, bool> d = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (string f in Directory.GetFiles(dir)) d[f] = true;
        return d;
    }

    // перенос идёт в фоне: ждём новый файл, пока он не допишется (размер не растёт и файл свободен).
    // Пока размер растёт - ждём сколько нужно (большое видео); не появился и не растёт StartSeconds - не вышло
    static string WaitNew(string dir, Dictionary<string, bool> before, PhoneFile f) {
        DateTime end = DateTime.Now.AddSeconds(StartSeconds);
        string cand = null;
        long len = -1;
        int same = 0;
        while (DateTime.Now < end) {
            System.Threading.Thread.Sleep(250);
            if (cand == null)
                foreach (string x in Directory.GetFiles(dir))
                    if (!before.ContainsKey(x)) { cand = x; break; }
            if (cand == null) continue;
            long l;
            try { l = new FileInfo(cand).Length; } catch { continue; }
            if (l > len) end = DateTime.Now.AddSeconds(StartSeconds);
            // то же расширение - файл не меньше, чем на телефоне (iPhone может отдать HEIC как JPG - тогда размер другой)
            bool full = f.Size <= 0 || !Path.GetExtension(cand).Equals(Path.GetExtension(f.Name), StringComparison.OrdinalIgnoreCase) || l >= f.Size;
            if (l == len && l > 0 && full) same++;
            else same = 0;
            len = l;
            if (same < 2) continue;
            try { using (FileStream fs = new FileStream(cand, FileMode.Open, FileAccess.Read, FileShare.None)) { } }
            catch (IOException) { same = 0; continue; }
            return cand;
        }
        // не дождались: недописанный кусок убрать (на телефоне файл остался)
        if (cand != null) { try { File.Delete(cand); } catch { } }
        return null;
    }
}

// Удаление через IFileOperation (Windows Vista и новее) - так удаляет Проводник; флаги: без окна, без подтверждения, без
// сообщений об ошибках. Объявления COM - вручную (порядок методов как в shobjidl.h).
public static class ShellDelete {
    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItem { }

    [ComImport, Guid("947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IFileOperation {
        void Advise(IntPtr sink, out uint cookie);
        void Unadvise(uint cookie);
        void SetOperationFlags(uint flags);
        void SetProgressMessage([MarshalAs(UnmanagedType.LPWStr)] string message);
        void SetProgressDialog(IntPtr dialog);
        void SetProperties(IntPtr properties);
        void SetOwnerWindow(IntPtr hwnd);
        void ApplyPropertiesToItem(IShellItem item);
        void ApplyPropertiesToItems(IntPtr items);
        void RenameItem(IShellItem item, [MarshalAs(UnmanagedType.LPWStr)] string newName, IntPtr sink);
        void RenameItems(IntPtr items, [MarshalAs(UnmanagedType.LPWStr)] string newName);
        void MoveItem(IShellItem item, IShellItem folder, [MarshalAs(UnmanagedType.LPWStr)] string newName, IntPtr sink);
        void MoveItems(IntPtr items, IShellItem folder);
        void CopyItem(IShellItem item, IShellItem folder, [MarshalAs(UnmanagedType.LPWStr)] string newName, IntPtr sink);
        void CopyItems(IntPtr items, IShellItem folder);
        void DeleteItem(IShellItem item, IntPtr sink);
        void DeleteItems(IntPtr items);
        void NewItem(IShellItem folder, uint attributes, [MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] string template, IntPtr sink);
        void PerformOperations();
        [return: MarshalAs(UnmanagedType.Bool)]
        bool GetAnyOperationsAborted();
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string path, IntPtr bindContext, ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItem item);

    // true - удалено (операция прошла)
    public static bool Delete(string parsingName) {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT || Environment.OSVersion.Version.Major < 6) return false;
        Guid iid = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
        IShellItem item;
        if (SHCreateItemFromParsingName(parsingName, IntPtr.Zero, ref iid, out item) != 0 || item == null) return false;
        Type t = Type.GetTypeFromCLSID(new Guid("3ad05575-8857-4850-9277-11b85bdb8e09"));
        IFileOperation op = (IFileOperation)Activator.CreateInstance(t);
        try {
            op.SetOperationFlags(0x4 | 0x10 | 0x400);   // FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI
            op.DeleteItem(item, IntPtr.Zero);
            op.PerformOperations();
            return !op.GetAnyOperationsAborted();
        } finally {
            Marshal.ReleaseComObject(op);
            Marshal.ReleaseComObject(item);
        }
    }
}

// Телефоны для автопереноса: подключили телефон из списка - фото уходят с него сами, по его настройке (какие: все, за сегодня,
// за вчера, за неделю; что: фото, скрины, видео). Телефон узнаётся по номеру устройства (IMEI по кабелю Windows не видит);
// IMEI пишется рядом - для себя. Список - telefony.txt рядом с программой.
public class KnownPhone {
    public string Id = "", Name = "", Imei = "";
    public string Period = "сегодня";            // все, сегодня, вчера, неделя
    public bool Photo = true, Screens = true, Video, On = true;
    public string Last = "";                     // последний автоперенос: когда и сколько

    public static readonly string[] Periods = { "все", "сегодня", "вчера", "неделя" };

    public static string PeriodText(string p) {
        return p == "все" ? "все фото" : p == "вчера" ? "за вчера" : p == "неделя" ? "за неделю" : "за сегодня";
    }

    public string KindsText {
        get {
            List<string> k = new List<string>();
            if (Photo) k.Add("фото");
            if (Screens) k.Add("скрины");
            if (Video) k.Add("видео");
            return k.Count == 0 ? "ничего" : string.Join(", ", k.ToArray());
        }
    }

    // брать ли файл (today - сегодняшнее число); неделя - сегодня и 6 дней до него
    public bool Takes(PhoneFile f, DateTime today) {
        if (f.Kind == "фото" && !Photo || f.Kind == "скрин" && !Screens || f.Kind == "видео" && !Video) return false;
        if (Period == "все") return true;
        if (f.Date == DateTime.MinValue) return false;
        DateTime d = f.Date.Date;
        if (Period == "вчера") return d == today.AddDays(-1);
        if (Period == "неделя") return d <= today && d >= today.AddDays(-6);
        return d == today;
    }
}

public static class PhoneBook {
    static string Clean(string s) { return (s ?? "").Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Trim(); }

    public static List<KnownPhone> Load(string file) {
        List<KnownPhone> r = new List<KnownPhone>();
        try {
            if (!File.Exists(file)) return r;
            foreach (string line in File.ReadAllLines(file, Encoding.UTF8)) {
                if (line.StartsWith("#") || line.Trim().Length == 0) continue;
                string[] p = line.Split('\t');
                if (p.Length < 6) continue;
                KnownPhone k = new KnownPhone();
                k.Id = p[0];
                k.Name = p[1];
                k.Imei = p[2];
                k.Period = Array.IndexOf(KnownPhone.Periods, p[3]) >= 0 ? p[3] : "сегодня";
                k.Photo = p[4].Contains("фото");
                k.Screens = p[4].Contains("скрины");
                k.Video = p[4].Contains("видео");
                k.On = p[5] != "0";
                k.Last = p.Length > 6 ? p[6] : "";
                r.Add(k);
            }
        } catch { }
        return r;
    }

    public static void Save(string file, List<KnownPhone> list) {
        StringBuilder sb = new StringBuilder("# Телефоны для автопереноса (программа «Фото с телефона»): номер устройства, название, IMEI, какие, что, вкл., последний перенос\r\n");
        foreach (KnownPhone k in list)
            sb.Append(Clean(k.Id)).Append('\t').Append(Clean(k.Name)).Append('\t').Append(Clean(k.Imei)).Append('\t').Append(k.Period).Append('\t')
              .Append(k.KindsText).Append('\t').Append(k.On ? "1" : "0").Append('\t').Append(Clean(k.Last)).Append("\r\n");
        File.WriteAllText(file, sb.ToString(), new UTF8Encoding(true));
    }

    public static KnownPhone Find(List<KnownPhone> list, string id) {
        foreach (KnownPhone k in list) if (k.Id == id) return k;
        return null;
    }
}

// Запуск при входе в Windows (свёрнутой): скрипт в папке «Автозагрузка» пользователя
public static class Autostart {
    static string FileName { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), "Фото с телефона.vbs"); } }

    public static bool IsOn { get { try { return File.Exists(FileName); } catch { return false; } } }

    public static void Set(bool on, string here) {
        if (!on) { if (File.Exists(FileName)) File.Delete(FileName); return; }
        string ps1 = Path.Combine(here, "foto.ps1");
        string vbs = "' Фото с телефона: запуск при входе в Windows (свёрнутой) - телефоны из списка отдают фото сами\r\n" +
            "CreateObject(\"WScript.Shell\").Run \"powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"\"" + ps1 + "\"\" /min\", 0, False\r\n";
        // путь к программе может быть по-русски: скрипт - в Юникоде (так его понимает Windows)
        File.WriteAllText(FileName, vbs, Encoding.Unicode);
    }
}

// Окно «Телефоны для автопереноса»
public class PhonesDialog : Form {
    readonly List<KnownPhone> phones;
    readonly List<PhoneDevice> connected;
    readonly ListView list = new ListView();
    readonly CheckBox auto = new CheckBox();
    bool filling, ready;

    public PhonesDialog(List<KnownPhone> phones, List<PhoneDevice> connected, string here) {
        this.phones = phones;
        this.connected = connected;
        Text = "Телефоны для автопереноса";
        Font = new Font("Segoe UI", 10f);
        BackColor = Color.White;
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(1000, 560);
        MinimumSize = new Size(760, 420);
        MinimizeBox = false;
        ShowInTaskbar = false;
        Label note = new Label();
        note.Text = "Подключили телефон из списка - фото уходят с него САМИ (вырезаются), по его настройке, в папки год \\ месяц \\ число. " +
            "Программа должна быть запущена (можно свёрнутой). IMEI по кабелю Windows не видит - телефон узнаётся по номеру устройства " +
            "(программа берёт его сама: подключите телефон и нажмите «Добавить подключённый»); IMEI впишите рядом - для себя.";
        note.SetBounds(16, 10, 950, 60);
        note.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
        list.View = View.Details;
        list.CheckBoxes = true;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.GridLines = true;
        list.SetBounds(16, 74, 952, 340);
        list.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
        list.Columns.Add("Телефон", 160);
        list.Columns.Add("IMEI", 140);
        list.Columns.Add("Что", 150);
        list.Columns.Add("Какие", 100);
        list.Columns.Add("Сейчас", 110);
        list.Columns.Add("Последний перенос", 150);
        list.Columns.Add("Номер устройства", 200);
        // Windows, создавая список, сама «отмечает» галочки строк - до показа окна это не нажатия; у строки «Список пуст» телефона нет
        list.ItemChecked += delegate(object s, ItemCheckedEventArgs e) {
            if (filling || !ready) return;
            KnownPhone k = e.Item.Tag as KnownPhone;
            if (k != null) k.On = e.Item.Checked;
            else if (e.Item.Checked) { ListViewItem it = e.Item; BeginInvoke((MethodInvoker)delegate { filling = true; try { it.Checked = false; } finally { filling = false; } }); }
        };
        Shown += delegate { ready = true; };
        list.DoubleClick += delegate { Edit(); };
        Button add = Btn("Добавить подключённый…", true), edit = Btn("Настроить…", false), del = Btn("Убрать из списка", false), done = Btn("Готово", false);
        add.SetBounds(16, 424, 230, 36);
        edit.SetBounds(256, 424, 130, 36);
        del.SetBounds(396, 424, 160, 36);
        done.SetBounds(848, 470, 120, 36);
        foreach (Button b in new Button[] { add, edit, del }) b.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        done.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        add.Click += delegate { Add(); };
        edit.Click += delegate { Edit(); };
        del.Click += delegate { Remove(); };
        done.DialogResult = DialogResult.OK;
        AcceptButton = done;
        CancelButton = done;
        auto.Text = "Запускать программу при входе в Windows (свёрнутой) - чтобы телефоны из списка отдавали фото сами";
        auto.AutoSize = true;
        auto.Location = new Point(16, 476);
        auto.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        auto.Checked = Autostart.IsOn;
        auto.CheckedChanged += delegate {
            try { Autostart.Set(auto.Checked, here); }
            catch (Exception ex) { MessageBox.Show(this, "Не получилось: " + ex.Message, Text); }
        };
        Controls.AddRange(new Control[] { note, list, add, edit, del, done, auto });
        Fill();
    }

    static Button Btn(string text, bool main) {
        Button b = new Button();
        b.Text = text;
        b.FlatStyle = FlatStyle.Flat;
        b.UseVisualStyleBackColor = false;
        b.BackColor = main ? Color.FromArgb(32, 100, 190) : Color.White;
        b.ForeColor = main ? Color.White : Color.FromArgb(30, 40, 55);
        b.FlatAppearance.BorderColor = main ? Color.FromArgb(32, 100, 190) : Color.FromArgb(150, 160, 175);
        return b;
    }

    bool IsConnected(KnownPhone k) { foreach (PhoneDevice d in connected) if (d.Id == k.Id) return true; return false; }

    void Fill() {
        filling = true;
        try {
            list.Items.Clear();
            foreach (KnownPhone k in phones) {
                ListViewItem it = new ListViewItem(new string[] { k.Name, k.Imei, k.KindsText, KnownPhone.PeriodText(k.Period),
                    IsConnected(k) ? "подключён" : "", k.Last, k.Id });
                it.Tag = k;
                it.Checked = k.On;
                if (IsConnected(k)) it.ForeColor = Color.FromArgb(30, 120, 60);
                list.Items.Add(it);
            }
            if (phones.Count == 0) {
                ListViewItem it = new ListViewItem(new string[] { "Список пуст", "", "подключите телефон и нажмите «Добавить подключённый»", "", "", "", "" });
                it.ForeColor = Color.Gray;
                list.Items.Add(it);
            }
        } finally { filling = false; }
    }

    // подключённые телефоны, которых ещё нет в списке
    List<PhoneDevice> NotInList() {
        List<PhoneDevice> r = new List<PhoneDevice>();
        foreach (PhoneDevice d in connected) if (PhoneBook.Find(phones, d.Id) == null) r.Add(d);
        return r;
    }

    public void Add() {
        List<PhoneDevice> l = NotInList();
        if (l.Count == 0) {
            MessageBox.Show(this, connected.Count == 0
                ? "Телефон не подключён. Подключите его кабелем, разблокируйте (Android - «Передача файлов», iPhone - «Доверять») и нажмите ещё раз."
                : "Все подключённые телефоны уже в списке.", Text);
            return;
        }
        if (l.Count == 1) { AddDevice(l[0]); return; }
        ContextMenuStrip m = new ContextMenuStrip();
        foreach (PhoneDevice d in l) {
            PhoneDevice dd = d;
            m.Items.Add(d.Name + "  (" + d.Id + ")", null, delegate { AddDevice(dd); });
        }
        m.Show(this, new Point(16, 424));
    }

    void AddDevice(PhoneDevice d) {
        KnownPhone k = new KnownPhone();
        k.Id = d.Id;
        k.Name = d.Name;
        using (PhoneEditDialog e = new PhoneEditDialog(k, true)) {
            if (e.ShowDialog(this) != DialogResult.OK) return;
        }
        phones.Add(k);
        Fill();
    }

    KnownPhone Selected() { return list.SelectedItems.Count > 0 ? list.SelectedItems[0].Tag as KnownPhone : null; }

    void Edit() {
        KnownPhone k = Selected();
        if (k == null) return;
        using (PhoneEditDialog e = new PhoneEditDialog(k, false)) e.ShowDialog(this);
        Fill();
    }

    void Remove() {
        KnownPhone k = Selected();
        if (k == null) return;
        if (MessageBox.Show(this, "Убрать «" + k.Name + "» из списка? Фото с него больше не будут уходить сами.", Text, MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
        phones.Remove(k);
        Fill();
    }

    // для проверки: отметить строку, как мышью
    public void CheckRow(int i, bool on) { if (i < list.Items.Count) list.Items[i].Checked = on; }
    public string Rows() {
        StringBuilder sb = new StringBuilder();
        foreach (ListViewItem it in list.Items) sb.Append(it.Checked ? "[x] " : "[ ] ").Append(it.Text).Append("; ");
        return sb.ToString();
    }

    // для проверки без окон
    public void AddFirst(string name, string imei, string period, bool photo, bool screens, bool video) {
        List<PhoneDevice> l = NotInList();
        if (l.Count == 0) return;
        KnownPhone k = new KnownPhone();
        k.Id = l[0].Id;
        k.Name = name.Length > 0 ? name : l[0].Name;
        k.Imei = imei;
        k.Period = period;
        k.Photo = photo;
        k.Screens = screens;
        k.Video = video;
        phones.Add(k);
        Fill();
    }
}

// Настройка телефона: название, IMEI (для себя), какие фото (одно), что (несколько), включён ли автоперенос
public class PhoneEditDialog : Form {
    readonly TextBox name = new TextBox(), imei = new TextBox();
    readonly RadioButton[] periods = new RadioButton[4];
    readonly CheckBox photo = new CheckBox(), screens = new CheckBox(), video = new CheckBox(), on = new CheckBox();

    public PhoneEditDialog(KnownPhone k, bool isNew) {
        Text = isNew ? "Новый телефон для автопереноса" : "Телефон «" + k.Name + "»";
        Font = new Font("Segoe UI", 10f);
        BackColor = Color.White;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(560, 430);
        int y = 14;
        Label ln = Cap("Название (как вам удобно: чей телефон)", y); y += 22;
        name.SetBounds(16, y, 528, 26); name.Text = k.Name; y += 36;
        Label li = Cap("IMEI (для себя - Windows его по кабелю не видит)", y); y += 22;
        imei.SetBounds(16, y, 260, 26); imei.Text = k.Imei; y += 34;
        Label lid = Cap("Номер устройства (по нему программа узнаёт телефон): " + k.Id, y); y += 30;
        GroupBox gp = new GroupBox();
        gp.Text = "Какие фото переносить";
        gp.SetBounds(16, y, 260, 150);
        string[] pt = { "все фото", "за сегодня", "за вчера", "за неделю (сегодня и 6 дней до)" };
        for (int i = 0; i < 4; i++) {
            periods[i] = new RadioButton();
            periods[i].Text = pt[i];
            periods[i].SetBounds(14, 26 + i * 30, 236, 26);
            periods[i].Checked = KnownPhone.Periods[i] == k.Period;
            gp.Controls.Add(periods[i]);
        }
        GroupBox gk = new GroupBox();
        gk.Text = "Что переносить (можно несколько)";
        gk.SetBounds(290, y, 254, 150);
        photo.Text = "фото"; photo.Checked = k.Photo; photo.SetBounds(14, 26, 220, 26);
        screens.Text = "скрины"; screens.Checked = k.Screens; screens.SetBounds(14, 56, 220, 26);
        video.Text = "видео"; video.Checked = k.Video; video.SetBounds(14, 86, 220, 26);
        gk.Controls.AddRange(new Control[] { photo, screens, video });
        y += 160;
        on.Text = "Переносить сам при подключении";
        on.Checked = k.On;
        on.SetBounds(16, y, 400, 26);
        Button ok = new Button(), cancel = new Button();
        ok.Text = "Сохранить";
        cancel.Text = "Отмена";
        ok.SetBounds(316, 384, 110, 34);
        cancel.SetBounds(434, 384, 110, 34);
        cancel.DialogResult = DialogResult.Cancel;
        ok.Click += delegate {
            if (!photo.Checked && !screens.Checked && !video.Checked) { MessageBox.Show(this, "Отметьте, что переносить: фото, скрины или видео.", Text); return; }
            k.Name = name.Text.Trim().Length > 0 ? name.Text.Trim() : k.Name;
            k.Imei = imei.Text.Trim();
            for (int i = 0; i < 4; i++) if (periods[i].Checked) k.Period = KnownPhone.Periods[i];
            k.Photo = photo.Checked;
            k.Screens = screens.Checked;
            k.Video = video.Checked;
            k.On = on.Checked;
            DialogResult = DialogResult.OK;
            Close();
        };
        AcceptButton = ok;
        CancelButton = cancel;
        Controls.AddRange(new Control[] { ln, name, li, imei, lid, gp, gk, on, ok, cancel });
    }

    static readonly Font SmallFont = new Font("Segoe UI", 9f);

    Label Cap(string text, int y) {
        Label l = new Label();
        l.Text = text;
        l.Font = SmallFont;
        l.ForeColor = Color.FromArgb(100, 110, 125);
        l.SetBounds(16, y, 528, 20);
        l.AutoEllipsis = true;
        return l;
    }
}

// Окно программы: куда складывать, календарь (день или период «с ... по»), дни с фото на телефоне (галочками - и отдельные
// дни), «Перенести с телефона» - фото уходят с телефона в папки ГГГГ\ММ\ДД по дате «Изменён».
public class FotoForm : Form {
    readonly string settingsFile;
    readonly Dictionary<string, string> settings = new Dictionary<string, string>();
    readonly TextBox path = new TextBox();
    readonly MonthCalendar cal = new MonthCalendar();
    readonly DateTimePicker from = new DateTimePicker(), to = new DateTimePicker();
    readonly CheckBox kPhoto = new CheckBox(), kScreens = new CheckBox(), kVideo = new CheckBox();   // что переносить вручную
    readonly RadioButton byMod = new RadioButton(), byTaken = new RadioButton();   // по какой дате раскладывать
    readonly Label dateNote = new Label();
    readonly CheckedListBox sources = new CheckedListBox();
    readonly ListView days = new ListView();
    readonly Label phoneText = new Label(), status = new Label(), chosen = new Label(), autoText = new Label();
    readonly Button scanButton, moveButton, stopButton, phonesButton;
    readonly string here, phonesFile;
    List<KnownPhone> phones;                                        // телефоны для автопереноса
    List<PhoneDevice> present = new List<PhoneDevice>();            // подключены сейчас
    readonly Dictionary<string, bool> autoDone = new Dictionary<string, bool>();   // путь устройства -> за это подключение уже перенесено
    readonly Timer settle = new Timer(), poll = new Timer();
    bool listing, pendingCheck, firstCheck = true, needScan;
    readonly ProgressBar bar = new ProgressBar();
    static readonly Font Base = new Font("Segoe UI", 10f), Small = new Font("Segoe UI", 9f), Bold = new Font("Segoe UI", 10f, FontStyle.Bold),
        Big = new Font("Segoe UI", 11f, FontStyle.Bold);
    static readonly Color Back = Color.FromArgb(243, 245, 248), Good = Color.FromArgb(30, 120, 60), Bad = Color.FromArgb(190, 40, 40),
        Hint = Color.FromArgb(100, 110, 125), Accent = Color.FromArgb(32, 100, 190);
    static readonly DateTime NoDate = DateTime.MinValue;

    List<PhoneFile> files = new List<PhoneFile>();
    readonly Dictionary<DateTime, bool> picked = new Dictionary<DateTime, bool>();   // отмеченные дни
    bool quiet, busy;
    StopFlag stop;

    public FotoForm(string here) {
        this.here = here;
        settingsFile = Path.Combine(here, "nastroiki.txt");
        phonesFile = Path.Combine(here, "telefony.txt");
        phones = PhoneBook.Load(phonesFile);
        LoadSettings();
        // программу перенесли в другую папку - запуск при входе в Windows должен вести сюда
        try { if (Autostart.IsOn) Autostart.Set(true, here); } catch { }
        Text = "Фото с телефона - перенос в папки по датам";
        Font = Base;
        BackColor = Back;
        StartPosition = FormStartPosition.CenterScreen;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1080, wa.Width), Math.Min(760, wa.Height));
        MinimumSize = new Size(900, 600);
        KeyPreview = true;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

        // ---- верх: куда складывать ----
        Panel top = new Panel();
        top.Dock = DockStyle.Top;
        top.Height = 70;
        top.BackColor = Color.White;
        Label lp = new Label();
        lp.Text = "Куда складывать (внутри - папки год \\ месяц \\ число, например 2026\\08\\21):";
        lp.Font = Small;
        lp.ForeColor = Hint;
        lp.AutoSize = true;
        lp.Location = new Point(16, 8);
        path.Font = Base;
        path.Text = Setting("Папка", DefaultRoot());
        path.TextChanged += delegate { settings["Папка"] = path.Text.Trim(); SaveSettings(); };
        Button pick = MakeButton("Выбрать…", false);
        pick.Click += delegate { PickFolder(); };
        Button open = MakeButton("Открыть папку", false);
        open.Click += delegate { OpenRoot(); };
        top.Controls.Add(lp);
        top.Controls.Add(path);
        top.Controls.Add(pick);
        top.Controls.Add(open);
        top.Resize += delegate {
            open.SetBounds(top.ClientSize.Width - 16 - 140, 28, 140, 32);
            pick.SetBounds(open.Left - 8 - 110, 28, 110, 32);
            path.SetBounds(16, 31, Math.Max(100, pick.Left - 24), 26);
        };

        // ---- слева: календарь ----
        Panel left = new Panel();
        left.Dock = DockStyle.Left;
        left.Width = 300;
        left.BackColor = Back;
        Label lc = new Label();
        lc.Text = "Какие дни";
        lc.Font = Bold;
        lc.AutoSize = true;
        lc.Location = new Point(16, 10);
        // какая дата: «Изменён» (как в Проводнике) или дата съёмки (как в галерее телефона) - бывают разные
        byMod.Text = "по дате «Изменён» (как в Проводнике)";
        byTaken.Text = "по дате съёмки (как в галерее)";
        foreach (RadioButton rb in new RadioButton[] { byMod, byTaken }) { rb.AutoSize = true; rb.Font = Small; }
        byMod.Location = new Point(16, 32);
        byTaken.Location = new Point(16, 52);
        PhoneShell.ByTaken = Setting("Дата", "изменён") == "съёмка";
        byTaken.Checked = PhoneShell.ByTaken;
        byMod.Checked = !PhoneShell.ByTaken;
        byTaken.CheckedChanged += delegate { DateMode(byTaken.Checked); };
        dateNote.Font = Small;
        dateNote.ForeColor = Color.FromArgb(190, 105, 0);
        dateNote.Location = new Point(16, 74);
        dateNote.Size = new Size(275, 30);
        cal.Location = new Point(16, 106);
        cal.CalendarDimensions = new Size(1, 2);
        cal.MaxSelectionCount = 3660;
        cal.ShowToday = true;
        cal.ShowTodayCircle = true;
        cal.DateChanged += delegate { if (!quiet) FromCalendar(); };
        Label lf = new Label();
        lf.Text = "с";
        lf.AutoSize = true;
        Label lt = new Label();
        lt.Text = "по";
        lt.AutoSize = true;
        foreach (DateTimePicker d in new DateTimePicker[] { from, to }) {
            d.Format = DateTimePickerFormat.Short;
            d.Width = 112;
            d.ValueChanged += delegate { if (!quiet) FromPickers(); };
        }
        Button today = MakeButton("Сегодня", false), yesterday = MakeButton("Вчера", false), all = MakeButton("Все дни", false);
        today.Click += delegate { SetRange(DateTime.Today, DateTime.Today); };
        yesterday.Click += delegate { SetRange(DateTime.Today.AddDays(-1), DateTime.Today.AddDays(-1)); };
        all.Click += delegate { AllDays(); };
        Label how = new Label();
        how.Font = Small;
        how.ForeColor = Hint;
        how.Text = "Щелчок по числу - один день. Период: щелчок по первому дню, Shift+щелчок по последнему (или «с ... по ...»). " +
            "Жирным - дни с фото; отдельные дни - галочками справа.";
        left.Controls.Add(lc);
        left.Controls.Add(byMod);
        left.Controls.Add(byTaken);
        left.Controls.Add(dateNote);
        left.Controls.Add(cal);
        left.Controls.Add(lf);
        left.Controls.Add(from);
        left.Controls.Add(lt);
        left.Controls.Add(to);
        left.Controls.Add(today);
        left.Controls.Add(yesterday);
        left.Controls.Add(all);
        left.Controls.Add(how);
        left.Layout += delegate {
            int y = cal.Bottom + 10;
            lf.Location = new Point(16, y + 4);
            from.Location = new Point(34, y);
            lt.Location = new Point(from.Right + 8, y + 4);
            to.Location = new Point(lt.Right + 4, y);
            y = from.Bottom + 8;
            today.SetBounds(16, y, 84, 30);
            yesterday.SetBounds(today.Right + 6, y, 74, 30);
            all.SetBounds(yesterday.Right + 6, y, 84, 30);
            how.SetBounds(16, today.Bottom + 8, left.ClientSize.Width - 24, Math.Max(40, left.ClientSize.Height - today.Bottom - 12));
        };

        // ---- справа: телефон, папки, дни ----
        Panel right = new Panel();
        right.Dock = DockStyle.Fill;
        right.BackColor = Back;
        right.Padding = new Padding(8, 8, 16, 8);
        phoneText.Font = Base;
        phoneText.AutoEllipsis = true;
        scanButton = MakeButton("Посмотреть телефон снова", false);
        scanButton.Click += delegate { Scan(); };
        phonesButton = MakeButton("Телефоны для автопереноса…", false);
        phonesButton.Click += delegate { OpenPhones(); };
        autoText.Font = Small;
        autoText.AutoEllipsis = true;
        Label lk = new Label();
        lk.Text = "Что переносить:";
        lk.AutoSize = true;
        string[][] kinds = { new string[] { "фото", "Фото", "1" }, new string[] { "скрины", "Скрины", "1" }, new string[] { "видео", "Видео", "0" } };
        CheckBox[] kb = { kPhoto, kScreens, kVideo };
        for (int i = 0; i < 3; i++) {
            CheckBox c = kb[i];
            string key = kinds[i][1];
            c.Text = kinds[i][0];
            c.AutoSize = true;
            c.Checked = Setting(key, kinds[i][2]) == "1";
            c.CheckedChanged += delegate { settings[key] = c.Checked ? "1" : "0"; SaveSettings(); Rebuild(); };
        }
        Label ls = new Label();
        ls.Text = "Откуда (папки телефона):";
        ls.Font = Small;
        ls.ForeColor = Hint;
        ls.AutoSize = true;
        sources.CheckOnClick = true;
        sources.IntegralHeight = false;
        sources.BorderStyle = BorderStyle.FixedSingle;
        sources.ItemCheck += delegate { BeginInvoke((MethodInvoker)delegate { Rebuild(); }); };
        Label ld = new Label();
        ld.Text = "Дни с фото на телефоне (галочка - перенести этот день):";
        ld.Font = Small;
        ld.ForeColor = Hint;
        ld.AutoSize = true;
        days.View = View.Details;
        days.CheckBoxes = true;
        days.FullRowSelect = true;
        days.HideSelection = false;
        days.GridLines = true;
        days.BorderStyle = BorderStyle.FixedSingle;
        days.Columns.Add("Число", 110);
        days.Columns.Add("", 50);
        days.Columns.Add("Фото", 70, HorizontalAlignment.Right);
        days.Columns.Add("Скрины", 70, HorizontalAlignment.Right);
        days.Columns.Add("Видео", 70, HorizontalAlignment.Right);
        days.Columns.Add("Размер", 90, HorizontalAlignment.Right);
        days.Columns.Add("Куда", 260);
        days.ItemChecked += delegate(object s, ItemCheckedEventArgs e) {
            if (quiet || !(e.Item.Tag is DateTime)) return;
            DateTime d = (DateTime)e.Item.Tag;
            if (e.Item.Checked) picked[d] = true; else picked.Remove(d);
            ShowChosen();
        };
        chosen.Font = Bold;
        chosen.AutoEllipsis = true;
        right.Controls.Add(phoneText);
        right.Controls.Add(scanButton);
        right.Controls.Add(autoText);
        right.Controls.Add(phonesButton);
        right.Controls.Add(lk);
        right.Controls.Add(kPhoto);
        right.Controls.Add(kScreens);
        right.Controls.Add(kVideo);
        right.Controls.Add(ls);
        right.Controls.Add(sources);
        right.Controls.Add(ld);
        right.Controls.Add(days);
        right.Controls.Add(chosen);
        right.Layout += delegate {
            int w = right.ClientSize.Width - 24;
            scanButton.SetBounds(right.ClientSize.Width - 16 - 230, 8, 230, 30);
            phoneText.SetBounds(8, 12, Math.Max(50, scanButton.Left - 16), 22);
            phonesButton.SetBounds(right.ClientSize.Width - 16 - 230, 42, 230, 30);
            autoText.SetBounds(8, 46, Math.Max(50, phonesButton.Left - 16), 22);
            lk.Location = new Point(8, 84);
            kPhoto.Location = new Point(lk.Right + 10, 82);
            kScreens.Location = new Point(kPhoto.Right + 14, 82);
            kVideo.Location = new Point(kScreens.Right + 14, 82);
            ls.Location = new Point(8, 112);
            sources.SetBounds(8, 130, w, 74);
            ld.Location = new Point(8, sources.Bottom + 10);
            chosen.SetBounds(8, right.ClientSize.Height - 30, w, 24);
            days.SetBounds(8, ld.Bottom + 2, w, Math.Max(60, chosen.Top - ld.Bottom - 8));
        };

        // ---- низ: перенос ----
        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = 76;
        bottom.BackColor = Color.White;
        bottom.Paint += delegate(object s, PaintEventArgs e) { using (Pen p = new Pen(Color.FromArgb(214, 219, 226))) e.Graphics.DrawLine(p, 0, 0, bottom.Width, 0); };
        moveButton = MakeButton("Перенести с телефона", true);
        moveButton.Font = Big;
        moveButton.Click += delegate { MoveFiles(); };
        stopButton = MakeButton("Стоп", false);
        stopButton.Enabled = false;
        stopButton.Click += delegate { if (stop != null) { stop.Stop = true; stopButton.Enabled = false; Status("Останавливаю после текущего файла...", Hint); } };
        status.Font = Base;
        status.AutoEllipsis = false;
        bar.Style = ProgressBarStyle.Continuous;
        bottom.Controls.Add(moveButton);
        bottom.Controls.Add(stopButton);
        bottom.Controls.Add(status);
        bottom.Controls.Add(bar);
        bottom.Resize += delegate {
            moveButton.SetBounds(bottom.ClientSize.Width - 16 - 300, 14, 300, 46);
            stopButton.SetBounds(moveButton.Left - 8 - 90, 16, 90, 42);
            status.SetBounds(16, 8, Math.Max(50, stopButton.Left - 32), 44);
            bar.SetBounds(16, 54, Math.Max(50, stopButton.Left - 32), 12);
        };

        Controls.Add(right);
        Controls.Add(left);
        Controls.Add(bottom);
        Controls.Add(top);

        quiet = true;
        try {
            cal.SetDate(DateTime.Today);
            from.Value = to.Value = DateTime.Today;
        } finally { quiet = false; }
        KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.F5 && !busy) { e.SuppressKeyPress = true; Scan(); } };
        Shown += delegate {
            int n = 0;
            try { if (path.Text.Trim().Length > 0) n = PhoneShell.Recover(path.Text.Trim()); } catch { }
            // какие телефоны подключены: из списка - перенос сам, потом - фото всех телефонов в окне
            CheckDevices();
            if (n > 0 && Messages) MessageBox.Show(this, "Прошлый перенос прервался: " + n + " уже перенесённых файлов разложены по папкам дат.", Text);
        };
        // телефон подключили или отключили: Windows сообщает об изменении устройств; через 2 с (телефон «встанет») - проверить.
        // И раз в 20 с - на случай, если сообщение не пришло или телефон из списка был заблокирован
        settle.Interval = 2000;
        settle.Tick += delegate { settle.Stop(); CheckDevices(); };
        poll.Interval = 20000;
        poll.Tick += delegate { if (!busy) CheckDevices(); };
        poll.Start();
        UpdateAutoText();
        FormClosing += delegate(object s, FormClosingEventArgs e) {
            if (busy && stop != null && e.CloseReason == CloseReason.UserClosing) {
                if (MessageBox.Show(this, "Идёт перенос. Остановить и закрыть программу?\n(Текущий файл допишется, остальные останутся на телефоне.)", Text,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) e.Cancel = true;
                else stop.Stop = true;
            }
        };
        ShowChosen();
    }

    static Button MakeButton(string text, bool main) {
        Button b = new Button();
        b.Text = text;
        b.FlatStyle = FlatStyle.Flat;
        b.Cursor = Cursors.Hand;
        b.UseVisualStyleBackColor = false;
        b.BackColor = main ? Accent : Color.White;
        b.ForeColor = main ? Color.White : Color.FromArgb(30, 40, 55);
        b.FlatAppearance.BorderColor = main ? Accent : Color.FromArgb(150, 160, 175);
        return b;
    }

    // ---------- настройки (nastroiki.txt рядом с программой) ----------

    void LoadSettings() {
        try {
            if (!File.Exists(settingsFile)) return;
            foreach (string line in File.ReadAllLines(settingsFile, Encoding.UTF8)) {
                int eq = line.IndexOf('=');
                if (eq > 0 && !line.StartsWith("#")) settings[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
            }
        } catch { }
    }

    void SaveSettings() {
        try {
            StringBuilder sb = new StringBuilder("# Настройки программы «Фото с телефона»\r\n");
            foreach (KeyValuePair<string, string> kv in settings) sb.Append(kv.Key).Append('=').Append(kv.Value).Append("\r\n");
            File.WriteAllText(settingsFile, sb.ToString(), new UTF8Encoding(true));
        } catch { }
    }

    string Setting(string k, string def) { string v; return settings.TryGetValue(k, out v) && v.Length > 0 ? v : def; }

    // первый раз: на диске D (если он есть) - D:\Фото, иначе «Изображения\Фото с телефона»
    static string DefaultRoot() {
        try {
            foreach (DriveInfo d in DriveInfo.GetDrives())
                if (d.DriveType == DriveType.Fixed && d.Name.StartsWith("D", StringComparison.OrdinalIgnoreCase) && d.IsReady) return Path.Combine(d.Name, "Фото");
        } catch { }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Фото с телефона");
    }

    void PickFolder() {
        using (FolderBrowserDialog d = new FolderBrowserDialog()) {
            d.Description = "Куда складывать фото с телефона (внутри появятся папки год \\ месяц \\ число)";
            if (Directory.Exists(path.Text.Trim())) d.SelectedPath = path.Text.Trim();
            if (d.ShowDialog(this) == DialogResult.OK) path.Text = d.SelectedPath;
        }
        Rebuild();
    }

    void OpenRoot() {
        string p = path.Text.Trim();
        try {
            Directory.CreateDirectory(p);
            Process.Start("explorer.exe", "\"" + p + "\"");
        } catch (Exception ex) { MessageBox.Show(this, "Не открывается:\n" + p + "\n\n" + ex.Message, Text); }
    }

    // ---------- выбор дней ----------

    void SetRange(DateTime a, DateTime b) {
        quiet = true;
        try {
            if (b < a) { DateTime t = a; a = b; b = t; }
            from.Value = a;
            to.Value = b;
            cal.SelectionRange = new SelectionRange(a, b);
        } finally { quiet = false; }
        PickRange(a, b);
    }

    void FromCalendar() {
        SelectionRange r = cal.SelectionRange;
        quiet = true;
        try { from.Value = r.Start.Date; to.Value = r.End.Date; } finally { quiet = false; }
        PickRange(r.Start.Date, r.End.Date);
    }

    void FromPickers() {
        DateTime a = from.Value.Date, b = to.Value.Date;
        if (b < a) { quiet = true; try { to.Value = a; } finally { quiet = false; } b = a; }
        quiet = true;
        try { cal.SelectionRange = new SelectionRange(a, b); } finally { quiet = false; }
        PickRange(a, b);
    }

    // отмечены дни периода (те, за которые есть фото)
    void PickRange(DateTime a, DateTime b) {
        picked.Clear();
        foreach (DateTime d in DayList()) if (d != NoDate && d >= a && d <= b) picked[d] = true;
        ShowDays();
    }

    void AllDays() {
        List<DateTime> l = DayList();
        if (l.Count == 0) return;
        DateTime a = DateTime.MaxValue, b = DateTime.MinValue;
        foreach (DateTime d in l) if (d != NoDate) { if (d < a) a = d; if (d > b) b = d; }
        if (a == DateTime.MaxValue) { picked.Clear(); picked[NoDate] = true; ShowDays(); return; }
        SetRange(a, b);
        if (l.Contains(NoDate)) { picked[NoDate] = true; ShowDays(); }
    }

    // ---------- телефон ----------

    void Post(MethodInvoker m) { try { if (!IsDisposed && IsHandleCreated) BeginInvoke(m); } catch { } }

    void Status(string text, Color c) { status.Text = text; status.ForeColor = c; }

    void SetBusy(bool b) {
        busy = b;
        scanButton.Enabled = moveButton.Enabled = phonesButton.Enabled = !b;
        stopButton.Enabled = b && stop != null;
        Cursor = b ? Cursors.AppStarting : Cursors.Default;
        if (!b) {
            bar.Value = 0;
            ShowChosen();
            // пока было занято, телефон подключили/отключили - проверить сейчас
            if (pendingCheck) { pendingCheck = false; BeginInvoke((MethodInvoker)delegate { CheckDevices(); }); }
        }
    }

    public void Scan() {
        if (busy) return;
        stop = null;
        SetBusy(true);
        Status("Ищу телефон...", Hint);
        phoneText.Text = "Телефон: ищу...";
        phoneText.ForeColor = Hint;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            ScanResult r = PhoneShell.Scan(delegate(string t) { Post(delegate { Status(t, Hint); }); });
            Post(delegate { ScanDone(r); });
        });
        // оболочка Windows (как Проводник) работает только в таком потоке
        th.SetApartmentState(System.Threading.ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
    }

    void ScanDone(ScanResult r) {
        SetBusy(false);
        files = r.Files;
        // папки телефона: прежние отключённые - так и остаются выключены
        List<string> off = new List<string>(Setting("Без папок", "").Split('|'));
        quiet = true;
        sources.BeginUpdate();
        try {
            sources.Items.Clear();
            foreach (string s in r.Sources) {
                int n = 0;
                foreach (PhoneFile f in files) if (f.Source == s) n++;
                sources.Items.Add(new SourceItem(s, n), !off.Contains(s));
            }
        } finally { sources.EndUpdate(); quiet = false; }
        sources.Invalidate();
        if (r.Devices.Count > 0) {
            phoneText.Text = "Телефон: " + string.Join(", ", r.Devices.ToArray()) + " - файлов " + files.Count;
            phoneText.ForeColor = r.Error == null ? Good : Bad;
        } else { phoneText.Text = "Телефон не найден"; phoneText.ForeColor = Bad; }
        if (r.Error != null) Status(r.Error, Bad);
        else Status("Выберите дни в календаре (или галочками в списке) и нажмите «Перенести с телефона».", Hint);
        // отчёт рядом с программой: что видно на телефоне (папки, сколько файлов, по каким дням; без имён файлов)
        try { File.WriteAllText(Path.Combine(here, "otchet.txt"), PhoneShell.Report(r), new UTF8Encoding(true)); } catch { }
        Rebuild();
        FromPickers();
    }

    // раскладывать по «Изменён» или по дате съёмки: дни пересчитываются
    void DateMode(bool taken) {
        PhoneShell.ByTaken = taken;
        settings["Дата"] = taken ? "съёмка" : "изменён";
        SaveSettings();
        foreach (PhoneFile f in files) PhoneShell.UseDate(f);
        Rebuild();
        FromPickers();
    }

    class SourceItem {
        public readonly string Source;
        readonly int count;
        public SourceItem(string s, int n) { Source = s; count = n; }
        public override string ToString() { return Source + "  (" + count + ")"; }
    }

    bool SourceOn(string s) {
        for (int i = 0; i < sources.Items.Count; i++) {
            SourceItem it = (SourceItem)sources.Items[i];
            if (it.Source == s) return sources.GetItemChecked(i);
        }
        return false;
    }

    // какие файлы участвуют: папка телефона включена, видео - если галочка
    List<PhoneFile> Shown2() {
        List<PhoneFile> r = new List<PhoneFile>();
        foreach (PhoneFile f in files) if (SourceOn(f.Source) && KindOn(f.Kind)) r.Add(f);
        return r;
    }

    bool KindOn(string kind) { return kind == "скрин" ? kScreens.Checked : kind == "видео" ? kVideo.Checked : kPhoto.Checked; }

    List<DateTime> DayList() {
        Dictionary<DateTime, bool> d = new Dictionary<DateTime, bool>();
        foreach (PhoneFile f in Shown2()) d[f.Date == NoDate ? NoDate : f.Date.Date] = true;
        List<DateTime> l = new List<DateTime>(d.Keys);
        l.Sort();
        l.Reverse();
        return l;
    }

    // после изменения папок телефона, галочки видео, списка файлов
    void Rebuild() {
        if (quiet) return;
        // выключенные папки: этого телефона - как отмечено сейчас, отключённых телефонов - как были
        List<string> off = new List<string>(), shown = new List<string>();
        for (int i = 0; i < sources.Items.Count; i++) {
            string src = ((SourceItem)sources.Items[i]).Source;
            shown.Add(src);
            if (!sources.GetItemChecked(i)) off.Add(src);
        }
        foreach (string x in Setting("Без папок", "").Split('|')) if (x.Length > 0 && !shown.Contains(x) && !off.Contains(x)) off.Add(x);
        string offText = string.Join("|", off.ToArray());
        if (offText != Setting("Без папок", "")) { settings["Без папок"] = offText; SaveSettings(); }
        // у скольких файлов «Изменён» и дата съёмки - разные дни (галерея или облако перезаписали файл)
        int diff = 0;
        foreach (PhoneFile f in Shown2())
            if (f.Modified != NoDate && f.Taken != NoDate && f.Modified.Date != f.Taken.Date) diff++;
        dateNote.Text = diff == 0 ? "" : "У " + diff + " файлов дата съёмки и «Изменён» - разные дни";
        List<DateTime> l = DayList();
        List<DateTime> bold = new List<DateTime>();
        foreach (DateTime d in l) if (d != NoDate) bold.Add(d);
        cal.BoldedDates = bold.ToArray();
        // отмеченные дни, которых больше нет, - снять
        foreach (DateTime d in new List<DateTime>(picked.Keys)) if (!l.Contains(d)) picked.Remove(d);
        ShowDays();
    }

    void ShowDays() {
        string root = path.Text.Trim();
        Dictionary<DateTime, long[]> sum = new Dictionary<DateTime, long[]>();   // фото, скрины, видео, байты
        foreach (PhoneFile f in Shown2()) {
            DateTime d = f.Date == NoDate ? NoDate : f.Date.Date;
            long[] s;
            if (!sum.TryGetValue(d, out s)) { s = new long[4]; sum[d] = s; }
            s[f.Kind == "скрин" ? 1 : f.Kind == "видео" ? 2 : 0]++;
            s[3] += Math.Max(0, f.Size);
        }
        List<DateTime> l = new List<DateTime>(sum.Keys);
        l.Sort();
        l.Reverse();
        quiet = true;
        days.BeginUpdate();
        try {
            days.Items.Clear();
            ListViewItem first = null;
            foreach (DateTime d in l) {
                long[] s = sum[d];
                ListViewItem it = new ListViewItem(new string[] { d == NoDate ? "без даты" : d.ToString("dd.MM.yyyy"),
                    d == NoDate ? "" : CultureInfo.GetCultureInfo("ru-RU").DateTimeFormat.GetAbbreviatedDayName(d.DayOfWeek),
                    s[0] > 0 ? s[0].ToString() : "", s[1] > 0 ? s[1].ToString() : "", s[2] > 0 ? s[2].ToString() : "",
                    PhoneShell.Mb(s[3]), root.Length > 0 ? PhoneShell.DayFolder(root, d).Substring(root.TrimEnd('\\').Length).TrimStart('\\') : "" });
                it.Tag = d;
                it.Checked = picked.ContainsKey(d);
                if (it.Checked) { it.BackColor = Color.FromArgb(226, 240, 255); if (first == null) first = it; }
                days.Items.Add(it);
            }
            if (first != null) first.EnsureVisible();
        } finally { days.EndUpdate(); quiet = false; }
        ShowChosen();
    }

    List<PhoneFile> Chosen() {
        List<PhoneFile> r = new List<PhoneFile>();
        foreach (PhoneFile f in Shown2()) if (picked.ContainsKey(f.Date == NoDate ? NoDate : f.Date.Date)) r.Add(f);
        return r;
    }

    void ShowChosen() {
        List<PhoneFile> c = Chosen();
        long size = 0;
        foreach (PhoneFile f in c) size += Math.Max(0, f.Size);
        chosen.Text = c.Count == 0 ? (files.Count == 0 ? "" : "За выбранные дни фото на телефоне нет") :
            "Выбрано: дней " + picked.Count + ", файлов " + c.Count + " (" + PhoneShell.Mb(size) + ")";
        chosen.ForeColor = c.Count == 0 ? Hint : Accent;
        moveButton.Text = c.Count > 0 ? "Перенести с телефона (" + c.Count + ")" : "Перенести с телефона";
        moveButton.Enabled = !busy && c.Count > 0;
        foreach (ListViewItem it in days.Items) it.BackColor = it.Checked ? Color.FromArgb(226, 240, 255) : Color.White;
    }

    // ---------- телефоны для автопереноса ----------

    const int WM_DEVICECHANGE = 0x0219;

    protected override void WndProc(ref Message m) {
        base.WndProc(ref m);
        if (m.Msg == WM_DEVICECHANGE) { settle.Stop(); settle.Start(); }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    static extern bool FlashWindow(IntPtr hwnd, bool invert);

    void UpdateAutoText() {
        int on = 0;
        foreach (KnownPhone k in phones) if (k.On) on++;
        List<string> here2 = new List<string>();
        foreach (PhoneDevice d in present) { KnownPhone k = PhoneBook.Find(phones, d.Id); if (k != null) here2.Add(k.Name); }
        autoText.Text = phones.Count == 0 ? "Автоперенос: список телефонов пуст - «Телефоны для автопереноса…»" :
            "Автоперенос: телефонов в списке " + phones.Count + (on < phones.Count ? " (включено " + on + ")" : "") +
            (here2.Count > 0 ? "; подключён: " + string.Join(", ", here2.ToArray()) : "");
        autoText.ForeColor = here2.Count > 0 ? Good : Hint;
    }

    void OpenPhones() {
        if (busy) return;
        using (PhonesDialog d = new PhonesDialog(phones, present, here)) d.ShowDialog(this);
        SavePhones();
        UpdateAutoText();
        // добавили подключённый телефон - он отдаст фото сразу
        CheckDevices();
    }

    void SavePhones() {
        try { PhoneBook.Save(phonesFile, phones); }
        catch (Exception ex) { if (Messages) MessageBox.Show(this, "Список телефонов не сохранился:\n" + ex.Message, Text); }
    }

    // какие телефоны подключены (в фоне, быстро)
    public void CheckDevices() {
        if (busy || listing) { pendingCheck = true; return; }
        listing = true;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            List<PhoneDevice> l;
            try { l = PhoneShell.ListDevices(); } catch { l = new List<PhoneDevice>(); }
            Post(delegate { DevicesListed(l); });
        });
        th.SetApartmentState(System.Threading.ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
    }

    static bool Has(List<PhoneDevice> l, PhoneDevice d) { foreach (PhoneDevice x in l) if (Key(x) == Key(d)) return true; return false; }

    static string Key(PhoneDevice d) { return d.Path + "|" + d.Name; }

    void DevicesListed(List<PhoneDevice> l) {
        listing = false;
        if (busy) { pendingCheck = true; return; }
        bool added = false, removed = false;
        foreach (PhoneDevice d in l) if (!Has(present, d)) added = true;
        foreach (PhoneDevice d in present) {
            if (Has(l, d)) continue;
            removed = true;
            autoDone.Remove(Key(d));
            // отключили - его фото из окна убрать
            PhoneDevice gone = d;
            files = files.FindAll(delegate(PhoneFile f) { return f.DevicePath.Length > 0 ? f.DevicePath != gone.Path : f.Device != gone.Name; });
        }
        present = l;
        UpdateAutoText();
        // телефоны из списка, которые ещё не отдали фото за это подключение
        foreach (PhoneDevice d in present) {
            KnownPhone k = PhoneBook.Find(phones, d.Id);
            if (k == null || !k.On || autoDone.ContainsKey(Key(d))) continue;
            StartAuto(d, k);
            return;
        }
        // новый или отключённый телефон, после автопереноса - фото всех подключённых телефонов в окне заново
        if (added || removed || firstCheck || needScan) { firstCheck = false; needScan = false; Scan(); }
    }

    // телефон из списка: посмотреть только его, взять фото по его настройке, перенести
    void StartAuto(PhoneDevice d, KnownPhone k) {
        string root = path.Text.Trim();
        if (root.Length == 0 || !Path.IsPathRooted(root)) { Status("Автоперенос: укажите, куда складывать фото.", Bad); autoDone[Key(d)] = true; return; }
        stop = new StopFlag();
        StopFlag myStop = stop;
        SetBusy(true);
        string what = "«" + k.Name + "» (из списка): " + KnownPhone.PeriodText(k.Period) + " - " + k.KindsText;
        Status("Телефон " + what + ". Смотрю телефон...", Accent);
        DateTime today = DateTime.Today;
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            TextHandler progress = delegate(string t) { Post(delegate { Status("«" + k.Name + "»: " + t, Accent); }); };
            ScanResult s = PhoneShell.Scan(progress, d.Path);
            List<PhoneFile> pick = s.Files.FindAll(delegate(PhoneFile f) { return k.Takes(f, today); });
            MoveResult m = pick.Count > 0 ? PhoneShell.Move(pick, root, myStop, progress) : null;
            Post(delegate { AutoDone(d, k, s, pick, m, root); });
        });
        th.SetApartmentState(System.Threading.ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
    }

    void AutoDone(PhoneDevice d, KnownPhone k, ScanResult s, List<PhoneFile> pick, MoveResult m, string root) {
        stop = null;
        string head = "Автоперенос «" + k.Name + "» (" + KnownPhone.PeriodText(k.Period) + ", " + k.KindsText + "): ";
        if (s.Sources.Count == 0) {
            // телефон заблокирован или не дал доступ - попробуем ещё раз через 20 с
            Status(head + "фото на телефоне не открываются - разблокируйте его (Android - «Передача файлов», iPhone - «Доверять»).", Bad);
            SetBusy(false);
            return;
        }
        autoDone[Key(d)] = true;
        needScan = true;
        string text;
        bool bad = false;
        if (m == null) text = "на телефоне таких нет";
        else {
            text = Report(m, root);
            bad = m.Error != null || m.Failed.Count > 0 || m.NotRemoved.Count > 0;
        }
        k.Last = DateTime.Now.ToString("dd.MM.yyyy HH:mm") + " - " + (m == null ? 0 : m.Done.Count - m.NotRemoved.Count);
        SavePhones();
        LastReport = head + text;
        Status(LastReport, bad ? Bad : Good);
        try { if (!ContainsFocus) FlashWindow(Handle, true); } catch { }
        SetBusy(false);
        CheckDevices();
    }

    // ---------- перенос ----------

    public void MoveFiles() {
        if (busy) return;
        List<PhoneFile> c = Chosen();
        if (c.Count == 0) return;
        string root = path.Text.Trim();
        if (root.Length == 0 || !Path.IsPathRooted(root)) {
            MessageBox.Show(this, "Укажите, куда складывать фото (кнопка «Выбрать…»).", Text);
            return;
        }
        long size = 0;
        foreach (PhoneFile f in c) size += Math.Max(0, f.Size);
        try {
            DriveInfo d = new DriveInfo(Path.GetPathRoot(root));
            if (d.IsReady && d.AvailableFreeSpace < size + 104857600) {
                MessageBox.Show(this, "На диске " + d.Name + " мало места: свободно " + PhoneShell.Mb(d.AvailableFreeSpace) + ", а переносится " + PhoneShell.Mb(size) + ".", Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        } catch { }
        stop = new StopFlag();
        StopFlag myStop = stop;
        SetBusy(true);
        bar.Maximum = c.Count;
        bar.Value = 0;
        Status("Переношу " + c.Count + " файлов с телефона...", Hint);
        System.Threading.Thread th = new System.Threading.Thread(delegate() {
            MoveResult r = PhoneShell.Move(c, root, myStop, delegate(string t) {
                Post(delegate {
                    Status(t, Hint);
                    int k = t.StartsWith("Переношу ") ? ParseIndex(t) : -1;
                    if (k > 0 && k <= bar.Maximum) bar.Value = k;
                });
            });
            Post(delegate { MoveDone(r, root); });
        });
        th.SetApartmentState(System.Threading.ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
    }

    static int ParseIndex(string t) {
        int a = "Переношу ".Length, b = t.IndexOf(' ', a);
        int k;
        return b > a && int.TryParse(t.Substring(a, b - a), out k) ? k : -1;
    }

    public string LastReport = "";

    void MoveDone(MoveResult r, string root) {
        stop = null;
        // ушедшие с телефона - убрать из списка
        Dictionary<PhoneFile, bool> gone = new Dictionary<PhoneFile, bool>();
        foreach (PhoneFile f in r.Done) gone[f] = true;
        foreach (PhoneFile f in r.NotRemoved) gone.Remove(f);
        List<PhoneFile> keep = new List<PhoneFile>();
        foreach (PhoneFile f in files) if (!gone.ContainsKey(f)) keep.Add(f);
        files = keep;
        SetBusy(false);
        Rebuild();
        int pos = phoneText.Text.LastIndexOf(" - файлов ");
        if (pos > 0) phoneText.Text = phoneText.Text.Substring(0, pos) + " - файлов " + files.Count;
        LastReport = Report(r, root);
        bool bad = r.Error != null || r.Failed.Count > 0 || r.NotRemoved.Count > 0;
        Status(LastReport, bad ? Bad : Good);
        if (bad && Messages) {
            StringBuilder m = new StringBuilder(LastReport + ".");
            if (r.Failed.Count > 0) { m.Append("\n\nНе перенеслись:"); AppendNames(m, r.Failed); }
            if (r.NotRemoved.Count > 0) { m.Append("\n\nТелефон не дал их удалить (на компьютере они уже есть):"); AppendNames(m, r.NotRemoved); }
            MessageBox.Show(this, m.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    // итог переноса одной строкой
    static string Report(MoveResult r, string root) {
        // какие папки дней получились
        Dictionary<string, bool> dirs = new Dictionary<string, bool>();
        foreach (PhoneFile f in r.Done) dirs[PhoneShell.DayFolder(root, f.Date).Substring(root.TrimEnd('\\').Length).TrimStart('\\')] = true;
        List<string> dl = new List<string>(dirs.Keys);
        dl.Sort();
        StringBuilder sb = new StringBuilder();
        if (r.Error != null) sb.Append(r.Error).Append(". ");
        int moved = r.Done.Count - r.NotRemoved.Count;
        string where = dl.Count == 0 ? "" : " в " + (dl.Count <= 4 ? string.Join(", ", dl.ToArray()) : dl[0] + " ... " + dl[dl.Count - 1] + " (" + dl.Count + " папок)");
        if (moved > 0 || r.NotRemoved.Count == 0) sb.Append("Перенесено ").Append(moved).Append(r.NotRemoved.Count == 0 ? where : "");
        if (r.Duplicates > 0) sb.Append("; ").Append(r.Duplicates).Append(" из них уже были в папке (второй раз не положены)");
        if (r.NotRemoved.Count > 0) sb.Append(sb.Length > 0 ? ". " : "").Append("Скопированы").Append(where).Append(", но телефон НЕ ДАЛ их удалить: ")
            .Append(r.NotRemoved.Count).Append(" - удалите на телефоне сами");
        if (r.Failed.Count > 0) sb.Append(". Не перенеслись (остались на телефоне): ").Append(r.Failed.Count);
        if (r.Stopped) sb.Append(". Остановлено - остальные на телефоне");
        return sb.ToString();
    }

    // для проверки без окон - без сообщений
    public bool Messages = true;

    static void AppendNames(StringBuilder m, List<PhoneFile> l) {
        for (int i = 0; i < l.Count && i < 15; i++) m.Append("\n  ").Append(l[i].Name).Append("  (").Append(l[i].Source).Append(")");
        if (l.Count > 15) m.Append("\n  ... и ещё ").Append(l.Count - 15);
    }

    // ---------- для проверки ----------
    public bool Busy { get { return busy; } }
    public void StopNow() { if (stop != null) stop.Stop = true; }
    public string StatusText { get { return status.Text; } }
    public void SetRoot(string p) { path.Text = p; }
    public void PickDays(DateTime a, DateTime b) { SetRange(a, b); }
    public void CheckDay(DateTime d, bool on) {
        foreach (ListViewItem it in days.Items) if (it.Tag is DateTime && (DateTime)it.Tag == d) it.Checked = on;
    }
    public void AddPhone(string device, string name, string imei, string period, bool photo, bool screens, bool video) {
        foreach (PhoneDevice d in present) {
            if (d.Name != device) continue;
            KnownPhone k = new KnownPhone();
            k.Id = d.Id;
            k.Name = name;
            k.Imei = imei;
            k.Period = period;
            k.Photo = photo;
            k.Screens = screens;
            k.Video = video;
            phones.Add(k);
            SavePhones();
            UpdateAutoText();
            return;
        }
    }
    public string PhonesDump() {
        StringBuilder sb = new StringBuilder(autoText.Text);
        foreach (KnownPhone k in phones) sb.Append("\n  ").Append(k.Name).Append(" | ").Append(k.Imei).Append(" | ").Append(k.Id).Append(" | ")
            .Append(KnownPhone.PeriodText(k.Period)).Append(" | ").Append(k.KindsText).Append(k.On ? "" : " | выкл.").Append(" | ").Append(k.Last);
        return sb.ToString();
    }
    public void ShowPhones() { BeginInvoke((MethodInvoker)delegate { OpenPhones(); }); }
    public void ByTaken(bool on) { if (on) byTaken.Checked = true; else byMod.Checked = true; }
    public void Kinds(bool photo, bool screens, bool video) { kPhoto.Checked = photo; kScreens.Checked = screens; kVideo.Checked = video; }
    public string Dump() {
        StringBuilder sb = new StringBuilder();
        sb.Append(phoneText.Text).Append('\n');
        for (int i = 0; i < sources.Items.Count; i++) sb.Append("  [").Append(sources.GetItemChecked(i) ? "x" : " ").Append("] ").Append(sources.Items[i]).Append('\n');
        foreach (ListViewItem it in days.Items) {
            sb.Append(it.Checked ? "  [x] " : "  [ ] ");
            for (int c = 0; c < it.SubItems.Count; c++) sb.Append(c > 0 ? " | " : "").Append(it.SubItems[c].Text);
            sb.Append('\n');
        }
        sb.Append(chosen.Text);
        return sb.ToString();
    }
}

public static class FotoApp {
    public static void Run(string here, string launch) { Run(here, launch, false); }

    // minimized - запуск при входе в Windows: окно свёрнуто, телефоны из списка отдают фото сами
    public static void Run(string here, string launch, bool minimized) {
        Crash.Install(here);
        try { Application.EnableVisualStyles(); } catch { }
        using (FotoForm f = new FotoForm(here)) {
            if (minimized) f.WindowState = FormWindowState.Minimized;
            Application.Run(f);
        }
    }
}
'@
# Программа собирается из этого скрипта ОДИН раз после обновления и сохраняется готовой (файл foto_*.dll
# в папке профиля пользователя). Следующие запуски берут готовую - без сборки, за секунды.
# Изменился скрипт - у файла другое имя, программа соберётся заново.
$t0 = Get-Date
$stamp = '' + $code.Length
try {
    $sha = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider
    $stamp += '_' + ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($code)), 0, 6) -replace '-', '')
} catch {
    $stamp += '_' + ($code.GetHashCode() -band 0x7FFFFFFF)
}
$name = "foto_$stamp.dll"
$cacheDirs = @()
if ($env:LOCALAPPDATA) { $cacheDirs += (Join-Path $env:LOCALAPPDATA 'foto-s-telefona') }
if ($env:PUBLIC) { $cacheDirs += (Join-Path $env:PUBLIC 'foto-s-telefona') }
$cacheDirs += (Join-Path $here 'kesh-programmy')

function Test-Ready { return ($null -ne ('FotoApp' -as [type])) }

$launch = $null
foreach ($d in $cacheDirs) {
    $dll = Join-Path $d $name
    if ((-not (Test-Ready)) -and (Test-Path $dll)) {
        try {
            [void][Reflection.Assembly]::LoadFrom($dll)
            if (Test-Ready) { $launch = 'готовая программа ' + $dll + ', ' + ((Get-Date) - $t0).TotalSeconds.ToString('0.0') + ' с' }
        } catch { }
    }
}

# Предупреждения компилятора (не ошибки) не должны мешать запуску: -IgnoreWarnings, если PowerShell его знает
$quiet = @{}
try { if ((Get-Command Add-Type).Parameters.ContainsKey('IgnoreWarnings')) { $quiet['IgnoreWarnings'] = $true } } catch { }

# Окошко на время сборки: иначе несколько минут ничего не видно
$splash = $null
$splashText = $null
function Show-Splash {
    try {
        $script:splash = New-Object System.Windows.Forms.Form
        $splash.Text = 'Фото с телефона'
        $splash.FormBorderStyle = 'FixedDialog'
        $splash.ControlBox = $false
        $splash.StartPosition = 'CenterScreen'
        $splash.ClientSize = New-Object System.Drawing.Size(460, 110)
        $splash.TopMost = $true
        $script:splashText = New-Object System.Windows.Forms.Label
        $splashText.Dock = 'Fill'
        $splashText.TextAlign = 'MiddleCenter'
        $splashText.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
        $splash.Controls.Add($splashText)
        Update-Splash 0
        $splash.Show()
        [System.Windows.Forms.Application]::DoEvents()
    } catch { $script:splash = $null }
}
function Update-Splash($seconds) {
    if ($splashText -eq $null) { return }
    $splashText.Text = "Программа обновилась - Windows собирает её.`nЭто только один раз после обновления.`n`nПрошло: " + [int]$seconds + ' с'
}

# Сборка. Временные файлы компилятора - в своей папке: папку TEMP Windows компилятор на некоторых
# компьютерах открыть не может ("cannot open ...\Temp\RES....tmp for writing").
# Собираем в фоне, а окошко тем временем показывает, сколько прошло.
function Build-Program($dll, $tmp) {
    if (Test-Path $tmp) { Get-ChildItem -Path $tmp -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
    else { [void](New-Item -ItemType Directory -Path $tmp) }
    $oldTmp = $env:TMP
    $oldTemp = $env:TEMP
    $env:TMP = $tmp
    $env:TEMP = $tmp
    $ps = $null
    try {
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript({
            param($code, $refs, $dll, $quiet)
            $ErrorActionPreference = 'Stop'
            if ($dll) { Add-Type @quiet -ReferencedAssemblies $refs -TypeDefinition $code -OutputAssembly $dll -OutputType Library }
            else { Add-Type @quiet -ReferencedAssemblies $refs -TypeDefinition $code }
        }).AddArgument($code).AddArgument($refs).AddArgument($dll).AddArgument($quiet)
        $h = $ps.BeginInvoke()
        while (-not $h.IsCompleted) {
            if ($splash -ne $null) {
                Update-Splash ((Get-Date) - $t0).TotalSeconds
                [System.Windows.Forms.Application]::DoEvents()
            }
            Start-Sleep -Milliseconds 200
        }
        try { [void]$ps.EndInvoke($h) } catch {
            $ex = $_.Exception
            if ($ex.InnerException -ne $null) { $ex = $ex.InnerException }
            throw $ex
        }
        if ($ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
    } finally {
        if ($ps -ne $null) { $ps.Dispose() }
        $env:TMP = $oldTmp
        $env:TEMP = $oldTemp
        Get-ChildItem -Path $tmp -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($dll -and -not (Test-Ready)) { [void][Reflection.Assembly]::LoadFrom($dll) }
}

$errors = @()
if (-not (Test-Ready)) {
    Show-Splash
    foreach ($d in $cacheDirs) {
        try {
            if (-not (Test-Path $d)) { [void](New-Item -ItemType Directory -Path $d) }
            # прежние версии программы больше не нужны
            Get-ChildItem -Path $d -Filter 'foto_*.dll' -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne $name } | Remove-Item -ErrorAction SilentlyContinue
            $dll = Join-Path $d $name
            # файл есть, но не загрузился (прошлая сборка прервалась) - собираем заново
            if (Test-Path $dll) { Remove-Item $dll -Force -ErrorAction SilentlyContinue }
            Build-Program $dll (Join-Path $d 'tmp')
            if (Test-Ready) { $launch = 'собрана и сохранена (' + $dll + ') за ' + ((Get-Date) - $t0).TotalSeconds.ToString('0') + ' с'; break }
        } catch {
            $errors += "$d : $_"
            if ($dll -and (Test-Path $dll)) { Remove-Item $dll -Force -ErrorAction SilentlyContinue }
            # ошибка в самом коде - другая папка не поможет
            if ("$_" -match 'error CS\d+') { break }
        }
    }
    # сохранить не получилось нигде - собираем в памяти (временные файлы - рядом со скриптом)
    if ((-not (Test-Ready)) -and -not ("$errors" -match 'error CS\d+')) {
        try {
            Build-Program $null (Join-Path $here 'tmp')
            if (Test-Ready) { $launch = 'собрана в памяти (сохранить не удалось) за ' + ((Get-Date) - $t0).TotalSeconds.ToString('0') + ' с; ' + ($errors -join '; ') }
        } catch { $errors += "в памяти: $_" }
    }
    if ($splash -ne $null) { $splash.Close(); $splash.Dispose() }
}

# ошибка запуска - в oshibki.txt рядом с программой (что за Windows, PowerShell, .NET, памяти) - его можно прислать
function Write-Crash($what, $text) {
    try {
        $mem = ''
        try { $os = Get-WmiObject Win32_OperatingSystem; $mem = ', памяти свободно ' + [int]($os.FreePhysicalMemory / 1024) + ' из ' + [int]($os.TotalVisibleMemorySize / 1024) + ' МБ' } catch { }
        $info = 'Windows ' + [Environment]::OSVersion.Version + ', PowerShell ' + $PSVersionTable.PSVersion + ', .NET ' + [Environment]::Version + ', ' + ([IntPtr]::Size * 8) + '-битный' + $mem
        $text = [regex]::Replace("$text", '\d{6,}', '######')
        $line = '=== ' + (Get-Date).ToString('dd.MM.yyyy HH:mm:ss') + ' (' + $what + ") ===`r`n" + $info + "`r`n" + $text + "`r`n`r`n"
        [IO.File]::AppendAllText((Join-Path $here 'oshibki.txt'), $line, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}

if (-not (Test-Ready)) {
    Write-Crash 'сборка' ($errors -join "`r`n`r`n")
    [void][System.Windows.Forms.MessageBox]::Show("Не удалось подготовить программу к запуску:`n`n" + ($errors -join "`n`n") + "`n`nПодробности - в файле oshibki.txt рядом с программой. Пришлите его (или фото этого окна).", 'Фото с телефона - ошибка')
    exit
}

try {
    # /min - запуск при входе в Windows (окно свёрнуто, телефоны из списка отдают фото сами)
    [FotoApp]::Run($here, $launch, [bool]($args -contains '/min'))
} catch {
    $ex = $_.Exception
    while ($ex.InnerException -ne $null) { $ex = $ex.InnerException }
    Write-Crash 'работа' $ex.ToString()
    [void][System.Windows.Forms.MessageBox]::Show("Ошибка: " + $ex.Message + "`n`nПодробности - в файле oshibki.txt рядом с программой. Пришлите его (или фото этого окна).", 'Фото с телефона - ошибка')
}
