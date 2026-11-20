#!/bin/bash
# Проверочный сценарий (C#) поверх собранной программы под Mono в Xvfb со скриншотами.
# runtest.sh <menu|foto|prosmotr|dogovor> <тест.cs> <папка программы> <png-префикс> <сек1> [сек2 ...]
# (сначала check.sh - он собирает *_mono.dll). Тест получает папку программы аргументом.
SP=/tmp/claude-0/-home-user-Test/ecddb7b9-ba2d-55d0-b8b6-3d6ae08ebda9/scratchpad; B=$SP/build
n=$1; cs=$2; dir=$3; png=$4; shift 4
t=$(basename "$cs" .cs)
mcs -nologo -out:$B/$t.exe -r:$B/${n}_mono.dll -r:System.Windows.Forms.dll -r:System.Drawing.dll "$cs" || exit 1
export DISPLAY=:77; pgrep -f "Xvfb :77" >/dev/null || (Xvfb :77 -screen 0 1366x768x24 >/dev/null 2>&1 &); sleep 1
(cd $B && MONO_IOMAP=all timeout 120 mono $t.exe "$dir" > $B/$t.log 2>&1 &)
prev=0; i=0
for s in "$@"; do sleep $(echo "$s - $prev" | bc); prev=$s; i=$((i+1)); import -window root "${png}_$i.png" 2>/dev/null; echo "скриншот ${png}_$i.png"; done
sleep 3; pkill -f "mono $t.exe"; grep -v 'Gtk not found' $B/$t.log
