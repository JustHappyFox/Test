# Вставка номеров: по нажатию клавиши печатает цифры из буфера обмена в активное окно.
# Всё, кроме цифр (пробелы, скобки, дефисы, буквы), отбрасывается.

$KEY   = 'F12'  # клавиша вставки: F12, F11, F8, Pause, Scroll, Insert ...
                # нужна клавиша, на которую ваша программа никак не реагирует
$PAUSE = 20     # пауза между цифрами в мс (если цифры теряются - поставьте 50)

$Host.UI.RawUI.WindowTitle = 'Вставка номеров'

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class K {
    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int x; public int y; }

    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mod, uint vk);
    [DllImport("user32.dll")] public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] static extern uint MapVirtualKey(uint code, uint mapType);

    public static void TypeDigits(string digits, int pause) {
        foreach (char c in digits) {
            byte vk = (byte)c;   // '0'..'9' совпадают с кодами клавиш 0..9
            byte sc = (byte)MapVirtualKey(vk, 0);
            keybd_event(vk, sc, 0, UIntPtr.Zero);
            keybd_event(vk, sc, 2, UIntPtr.Zero);
            Thread.Sleep(pause);
        }
    }
}
"@

$vk = [int][System.Windows.Forms.Keys]$KEY

# 0x4000 = MOD_NOREPEAT: удержание клавиши не повторяет вставку
if (-not [K]::RegisterHotKey([IntPtr]::Zero, 1, 0x4000, $vk)) {
    Write-Host "Клавиша $KEY уже занята (возможно, скрипт уже запущен). Это окно можно закрыть." -ForegroundColor Red
    exit
}

Write-Host 'Вставка номеров работает.' -ForegroundColor Green
Write-Host "Скопируйте номер, встаньте в нужное поле и нажмите $KEY."
Write-Host 'Чтобы выключить - просто закройте это окно.'
Write-Host ''

$msg = New-Object K+MSG
while ([K]::GetMessage([ref]$msg, [IntPtr]::Zero, 0, 0) -gt 0) {
    if ($msg.message -ne 0x0312) { continue }   # 0x0312 = WM_HOTKEY
    try   { $digits = [System.Windows.Forms.Clipboard]::GetText() -replace '[^0-9]', '' }
    catch { $digits = '' }
    if ($digits) {
        [K]::TypeDigits($digits, $PAUSE)
        Write-Host "Вставлено: $digits"
    }
}
