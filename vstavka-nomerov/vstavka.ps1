# Вставка номеров: Ctrl+V в рабочей программе печатает цифры из буфера обмена Windows.
# В остальных программах Ctrl+V работает как обычно.

$PAUSE = 20     # пауза между цифрами в мс (если цифры теряются - поставьте 50)

$Host.UI.RawUI.WindowTitle = 'Вставка номеров'

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

public static class K {
    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int x; public int y; }

    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mod, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")] public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vk);
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] static extern uint MapVirtualKey(uint code, uint mapType);

    public static string ActiveProgram() {
        uint pid;
        GetWindowThreadProcessId(GetForegroundWindow(), out pid);
        try { return Process.GetProcessById((int)pid).ProcessName; } catch { return ""; }
    }

    static void Key(byte vk, bool up) {
        keybd_event(vk, (byte)MapVirtualKey(vk, 0), up ? 2u : 0u, UIntPtr.Zero);
    }

    public static void TypeDigits(string digits, int pause) {
        // ждём, пока отпустят Ctrl, иначе программа получит Ctrl+цифру
        for (int i = 0; i < 300 && (GetAsyncKeyState(0x11) & 0x8000) != 0; i++) Thread.Sleep(10);
        foreach (char c in digits) {
            Key((byte)c, false);   // '0'..'9' совпадают с кодами клавиш 0..9
            Key((byte)c, true);
            Thread.Sleep(pause);
        }
    }

    public static void CtrlV() {
        Key(0x11, false); Key(0x56, false); Key(0x56, true); Key(0x11, true);
    }
}
"@

# Какую программу обслуживаем - запоминается при первом запуске в programma.txt
$cfg = Join-Path $PSScriptRoot 'programma.txt'

if (Test-Path $cfg) {
    $target = (Get-Content $cfg | Select-Object -First 1).Trim()
} else {
    Write-Host 'Первый запуск: нужно показать скрипту вашу рабочую программу.'
    do {
        Write-Host 'Щёлкните по окну программы - через 5 секунд я её запомню...'
        Start-Sleep -Seconds 5
        $target = [K]::ActiveProgram()
    } while (@('', 'powershell', 'explorer', 'WindowsTerminal', 'conhost') -contains $target)
    Set-Content -Path $cfg -Value $target
    Write-Host "Запомнил программу: $target"
}

# Ctrl+V: MOD_CONTROL (0x2) + MOD_NOREPEAT (0x4000), клавиша V (0x56)
if (-not [K]::RegisterHotKey([IntPtr]::Zero, 1, 0x4002, 0x56)) {
    Write-Host 'Ctrl+V уже перехвачен (возможно, скрипт уже запущен). Это окно можно закрыть.' -ForegroundColor Red
    exit
}

Write-Host ''
Write-Host "Работает. В программе $target клавиша Ctrl+V вставляет цифры из буфера Windows." -ForegroundColor Green
Write-Host 'В остальных программах Ctrl+V работает как обычно.'
Write-Host 'Окно можно свернуть. Чтобы выключить - закройте его.'
Write-Host ''

$msg = New-Object K+MSG
while ([K]::GetMessage([ref]$msg, [IntPtr]::Zero, 0, 0) -gt 0) {
    if ($msg.message -ne 0x0312) { continue }   # 0x0312 = WM_HOTKEY

    $digits = ''
    if ([K]::ActiveProgram() -eq $target) {
        try { $digits = [System.Windows.Forms.Clipboard]::GetText() -replace '[^0-9]', '' } catch { }
    }

    if ($digits) {
        [K]::TypeDigits($digits, $PAUSE)
        Write-Host "Вставлено: $digits"
    } else {
        # не та программа или в буфере нет цифр - пропускаем обычный Ctrl+V
        [void][K]::UnregisterHotKey([IntPtr]::Zero, 1)
        [K]::CtrlV()
        Start-Sleep -Milliseconds 100
        [void][K]::RegisterHotKey([IntPtr]::Zero, 1, 0x4002, 0x56)
    }
}
