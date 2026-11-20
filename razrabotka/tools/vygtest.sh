#!/bin/bash
# Общий блок выгрузок + тест (C#) -> exe под Mono. vygtest.sh <тест.cs> [аргументы...]; HOME/APPDATA - во временной папке
SP=/tmp/claude-0/-home-user-Test/ecddb7b9-ba2d-55d0-b8b6-3d6ae08ebda9/scratchpad; B=$SP/build; mkdir -p $B
R=$(cd "$(dirname "$0")/.." && pwd)
cs=$1; shift; t=$(basename "$cs" .cs)
{ echo "using System; using System.IO; using System.Text; using System.Collections.Generic; using System.Xml;"; cat "$R/shared/vygruzka.cs"; } > $B/vyg_block.cs
mcs -nologo -langversion:ISO-2 -warn:4 -out:$B/$t.exe -r:System.Xml.dll $B/vyg_block.cs "$cs" || exit 1
export XDG_CONFIG_HOME=$SP/test/profile/config XDG_DATA_HOME=$SP/test/profile/local; mkdir -p $XDG_CONFIG_HOME $XDG_DATA_HOME
mono $B/$t.exe "$@"
