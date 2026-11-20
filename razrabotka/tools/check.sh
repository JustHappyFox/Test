#!/bin/bash
# Проверка сборки программы (.ps1 со встроенным C#) по правилам репозитория.
# Использование: check.sh путь/к/программа.ps1 [...]
SP=/tmp/claude-0/-home-user-Test/ecddb7b9-ba2d-55d0-b8b6-3d6ae08ebda9/scratchpad
R=/usr/lib/dotnet/sdk/8.0.134/Roslyn/bincore/csc.dll
V2=$SP/refasm/ra/v2.0; V45=/usr/lib/mono/4.5-api
OUT=$SP/build; mkdir -p $OUT
rc=0
for ps in "$@"; do
  n=$(basename "$ps" .ps1); cs=$OUT/$n.cs
  python3 -I - "$ps" "$cs" <<'PY'
import sys,re
t=open(sys.argv[1],encoding='utf-8-sig').read()
m=re.search(r"^\$code = @'\r?\n(.*?)\r?\n'@", t, re.S|re.M)
open(sys.argv[2],'w',encoding='utf-8').write(m.group(1))
PY
  xml=""; grep -q 'System.Xml.XmlDocument' "$ps" && xml=1
  r2="-r:$V2/mscorlib.dll -r:$V2/System.dll -r:$V2/System.Drawing.dll -r:$V2/System.Windows.Forms.dll"; [ $xml ] && r2="$r2 -r:$V2/System.Xml.dll"
  r4="-r:$V45/mscorlib.dll -r:$V45/System.dll -r:$V45/System.Drawing.dll -r:$V45/System.Windows.Forms.dll"; [ $xml ] && r4="$r4 -r:$V45/System.Xml.dll"
  echo "=== $n ($(wc -l <"$cs") строк C#)"
  o=$(dotnet $R -nologo -nostdlib -noconfig -langversion:ISO-2 -warn:4 -t:library -out:$OUT/${n}_net20.dll $r2 "$cs" 2>&1); e=$(grep -cE ': (error|warning) ' <<<"$o"); echo "Roslyn .NET 2.0 / C# 2: $e ошибок/предупреждений"; [ $e != 0 ] && { grep -E ': (error|warning) ' <<<"$o" | head -20; rc=1; }
  o=$(dotnet $R -nologo -nostdlib -noconfig -langversion:ISO-2 -warn:4 -t:library -out:$OUT/${n}_net45.dll $r4 "$cs" 2>&1); e=$(grep -cE ': (error|warning) ' <<<"$o"); echo "Roslyn .NET 4.5: $e"; [ $e != 0 ] && { grep -E ': (error|warning) ' <<<"$o" | head -20; rc=1; }
  rm=""; for a in System System.Drawing System.Windows.Forms; do rm="$rm -r:$a.dll"; done; [ $xml ] && rm="$rm -r:System.Xml.dll"
  o=$(mcs -langversion:ISO-2 -warn:4 -t:library -out:$OUT/${n}_mono.dll $rm "$cs" 2>&1); e=$(grep -cE '(error|warning) CS' <<<"$o"); echo "mcs ISO-2: $e"; [ $e != 0 ] && { grep -E '(error|warning) CS' <<<"$o" | head -20; rc=1; }
  refs="$V2/mscorlib.dll $V2/System.dll $V2/System.Drawing.dll $V2/System.Windows.Forms.dll $V2/System.Xml.dll"
  o=$(dotnet $SP/tools/ternary/out/ternary.dll "$cs" $refs); echo "$o" | tail -1; grep -q ': 0$' <<<"$(echo "$o"|tail -1)" || { echo "$o" | head -20; rc=1; }
done
exit $rc
