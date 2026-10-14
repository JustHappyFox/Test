# Вставка номеров: маленькое окно с кнопкой поверх всех окон.
# Встаньте в поле программы и нажмите кнопку - в поле напечатаются цифры из буфера обмена.
# Всё, кроме цифр, отбрасывается. Окно не забирает фокус у программы,
# а клавиши F1-F12 не используются, поэтому программа ничего не перехватит.

$PAUSE = 20     # пауза между цифрами в мс (если цифры теряются - поставьте 50)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

try {
    Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

public class NoFocusButton : Button {
    public NoFocusButton() { SetStyle(ControlStyles.Selectable, false); }
}

public class PastePanel : Form {
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] static extern uint MapVirtualKey(uint code, uint mapType);

    readonly int pause;
    readonly Label status;

    public PastePanel(int pause) {
        this.pause = pause;
        Text = "Вставка номеров";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        ShowInTaskbar = false;
        TopMost = true;
        ClientSize = new Size(240, 90);
        StartPosition = FormStartPosition.Manual;
        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(area.Right - Width - 20, area.Bottom - Height - 20);

        Button button = new NoFocusButton();
        button.Text = "Вставить номер";
        button.Font = new Font(Font.FontFamily, 12f, FontStyle.Bold);
        button.Dock = DockStyle.Fill;
        button.Click += delegate { Paste(); };
        button.MouseEnter += delegate {
            string d = ClipboardDigits();
            status.Text = d.Length > 0 ? "В буфере: " + d : "В буфере нет цифр";
        };

        status = new Label();
        status.Dock = DockStyle.Bottom;
        status.Height = 24;
        status.TextAlign = ContentAlignment.MiddleCenter;
        status.Text = "Встаньте в поле и нажмите кнопку";

        Controls.Add(button);
        Controls.Add(status);
    }

    // окно не становится активным, фокус остаётся в поле программы
    protected override bool ShowWithoutActivation { get { return true; } }

    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE
            return cp;
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == 0x21) { m.Result = (IntPtr)3; return; }   // WM_MOUSEACTIVATE -> MA_NOACTIVATE
        base.WndProc(ref m);
    }

    static string ClipboardDigits() {
        try { return Regex.Replace(Clipboard.GetText(), "[^0-9]", ""); } catch { return ""; }
    }

    void Paste() {
        string digits = ClipboardDigits();
        if (digits.Length == 0) { status.Text = "В буфере нет цифр"; return; }
        foreach (char c in digits) {
            byte vk = (byte)c;   // '0'..'9' совпадают с кодами клавиш 0..9
            byte sc = (byte)MapVirtualKey(vk, 0);
            keybd_event(vk, sc, 0, UIntPtr.Zero);
            keybd_event(vk, sc, 2, UIntPtr.Zero);
            Thread.Sleep(pause);
        }
        status.Text = "Вставлено: " + digits;
    }

    public static void Run(int pause) {
        Application.EnableVisualStyles();
        Application.Run(new PastePanel(pause));
    }
}
"@
    [PastePanel]::Run($PAUSE)
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("$_", 'Вставка номеров - ошибка')
}
