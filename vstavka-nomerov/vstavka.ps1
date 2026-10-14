# Вставка номеров в DOS-базу: по F9 делает то же самое, что меню окна
# "Изменить -> Вставить", но из буфера берутся только цифры.

$KEY   = 'F9'     # клавиша вставки
$ENTER = $false   # $true - после номера сразу нажимать Enter

$Host.UI.RawUI.WindowTitle = 'Вставка номеров'

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class K {
    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int x; public int y; }

    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mod, uint vk);
    [DllImport("user32.dll")] public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);

    public static string ClassOf(IntPtr hWnd) {
        StringBuilder sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
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
Write-Host "Скопируйте номер, встаньте в нужное поле базы и нажмите $KEY."
Write-Host 'Чтобы выключить - просто закройте это окно.'
Write-Host ''

$msg = New-Object K+MSG
while ([K]::GetMessage([ref]$msg, [IntPtr]::Zero, 0, 0) -gt 0) {
    if ($msg.message -ne 0x0312) { continue }   # 0x0312 = WM_HOTKEY

    $win = [K]::GetForegroundWindow()
    if ([K]::ClassOf($win) -ne 'ConsoleWindowClass') {
        Write-Host "$KEY вставляет только в окно базы."
        continue
    }

    try { $text = [System.Windows.Forms.Clipboard]::GetText() } catch { $text = '' }
    $digits = $text -replace '[^0-9]', ''
    if (-not $digits) {
        Write-Host 'В буфере нет цифр.'
        continue
    }

    # оставляем в буфере только цифры, чтобы вставилось ровно это
    $paste = $digits
    if ($ENTER) { $paste += "`r`n" }
    if ($text -ne $paste) {
        try { [System.Windows.Forms.Clipboard]::SetText($paste) }
        catch { Write-Host 'Буфер занят, нажмите ещё раз.'; continue }
    }

    # 0x0112 = WM_SYSCOMMAND, 0xFFF1 = пункт меню окна "Изменить -> Вставить"
    [void][K]::PostMessage($win, 0x0112, [IntPtr]0xFFF1, [IntPtr]::Zero)
    Write-Host "Вставлено: $digits"
}
