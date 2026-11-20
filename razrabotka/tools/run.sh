#!/bin/bash
# Запуск собранной программы под Mono в Xvfb со скриншотом.
# run.sh <имя: menu|foto|prosmotr|dogovor> <папка программы> <секунд до скриншота> <файл.png>
SP=/tmp/claude-0/-home-user-Test/ecddb7b9-ba2d-55d0-b8b6-3d6ae08ebda9/scratchpad; B=$SP/build
declare -A CLS=([menu]=MenuApp [foto]=FotoApp [prosmotr]=AbonentSearch [dogovor]=DogovorApp)
n=$1; dir=$2; sec=${3:-8}; png=${4:-$B/$n.png}; c=${CLS[$n]}
cat > $B/run_$n.cs <<CS
public static class L { [System.STAThread] public static void Main() { $c.Run("$dir", "проверка"); } }
CS
mcs -nologo -out:$B/run_$n.exe -r:$B/${n}_mono.dll -r:System.Windows.Forms.dll $B/run_$n.cs || exit 1
cp $B/${n}_mono.dll $B/ 2>/dev/null
export DISPLAY=:77; (Xvfb :77 -screen 0 1366x768x24 >/dev/null 2>&1 &) ; sleep 1
(cd $B && MONO_IOMAP=all timeout $((sec+20)) mono run_$n.exe > $B/run_$n.log 2>&1 &)
sleep $sec; import -window root "$png" 2>/dev/null || xwd -root -silent | convert xwd:- "$png"
pkill -f "mono run_$n.exe"; echo "скриншот: $png"; tail -5 $B/run_$n.log
