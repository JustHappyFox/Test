# Вставка номеров: по F9 печатает цифры из буфера обмена в активное окно.
# Всё, кроме цифр (пробелы, скобки, дефисы, буквы), отбрасывается.

$KEY   = 0x78   # клавиша: F9 = 0x78, F8 = 0x77, F11 = 0x7A, F12 = 0x7B
$PAUSE = 20     # пауза между цифрами в мс (если цифры теряются - поставьте 50)

$Host.UI.RawUI.WindowTitle = 'Вставка номеров'

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class HK {
    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int x; public int y; }
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mod, uint vk);
    [DllImport("user32.dll")] public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
}
"@

$name = 'F' + ($KEY - 0x6F)

if (-not [HK]::RegisterHotKey([IntPtr]::Zero, 1, 0, $KEY)) {
    Write-Host "Клавиша $name уже занята (возможно, скрипт уже запущен). Это окно можно закрыть."
    exit
}

Write-Host 'Вставка номеров работает.'
Write-Host "Скопируйте номер, встаньте в нужное поле и нажмите $name."
Write-Host 'Чтобы выключить - просто закройте это окно.'

$msg = New-Object HK+MSG
while ([HK]::GetMessage([ref]$msg, [IntPtr]::Zero, 0, 0) -gt 0) {
    if ($msg.message -ne 0x0312) { continue }   # 0x0312 = WM_HOTKEY
    try   { $digits = [System.Windows.Forms.Clipboard]::GetText() -replace '\D', '' }
    catch { $digits = '' }
    foreach ($c in $digits.ToCharArray()) {
        [System.Windows.Forms.SendKeys]::SendWait([string]$c)
        Start-Sleep -Milliseconds $PAUSE
    }
}
