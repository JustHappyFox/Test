#!/usr/bin/env bash
# Compiles the pure-Java protocol layer and runs the scripted self test on the JVM.
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/.." && pwd)
OUT=$ROOT/build/selftest
rm -rf "$OUT"; mkdir -p "$OUT"
javac -source 8 -target 8 -encoding UTF-8 -Xlint:-options -nowarn -d "$OUT" \
  "$ROOT"/app/src/main/java/by/mobilemeter/util/*.java \
  "$ROOT"/app/src/main/java/by/mobilemeter/transport/SerialLink.java \
  "$ROOT"/app/src/main/java/by/mobilemeter/protocol/iec62056/*.java \
  "$ROOT"/app/src/main/java/by/mobilemeter/protocol/mirtek/*.java \
  "$ROOT"/app/src/test/java/by/mobilemeter/protocol/iec62056/SessionSelfTest.java \
  "$ROOT"/app/src/test/java/by/mobilemeter/protocol/mirtek/MirtekSelfTest.java
java -cp "$OUT" by.mobilemeter.protocol.iec62056.SessionSelfTest
java -cp "$OUT" by.mobilemeter.protocol.mirtek.MirtekSelfTest
