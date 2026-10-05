#!/usr/bin/env bash
# Builds a debug APK without Android Studio / Gradle, using only:
#   aapt2, apksigner, zipalign, dx (Ubuntu packages: aapt apksigner zipalign dalvik-exchange)
#   an android.jar-compatible framework jar (ANDROID_JAR), e.g. Robolectric android-all
#   a debug keystore (KEYSTORE)
# Usage: tools/build-apk.sh  -> build/manual/app-debug.apk
set -euo pipefail

ROOT=$(cd "$(dirname "$0")/.." && pwd)
TOOLS_DIR=${ANDROID_TOOLS_DIR:-/opt/android-tools}
ANDROID_JAR=${ANDROID_JAR:-$TOOLS_DIR/android-all-34.jar}
KEYSTORE=${KEYSTORE:-$TOOLS_DIR/debug.keystore}
DX=${DX:-$(command -v dx || command -v dalvik-exchange || echo /usr/lib/android-sdk/build-tools/debian/dx)}
APP_ID=by.mobilemeter
VERSION_CODE=${VERSION_CODE:-1}
VERSION_NAME=${VERSION_NAME:-0.1.0}
MIN_SDK=24
TARGET_SDK=34

OUT=$ROOT/build/manual
rm -rf "$OUT"
mkdir -p "$OUT/classes" "$OUT/gen"

echo "[1/6] resources"
aapt2 compile --dir "$ROOT/app/src/main/res" -o "$OUT/res.zip"

echo "[2/6] manifest + link"
# AGP takes the namespace from Gradle; aapt2 alone needs the package attribute.
sed "s|<manifest |<manifest package=\"$APP_ID\" |" "$ROOT/app/src/main/AndroidManifest.xml" > "$OUT/AndroidManifest.xml"
aapt2 link -o "$OUT/base.apk" -I "$ANDROID_JAR" \
  --manifest "$OUT/AndroidManifest.xml" \
  --java "$OUT/gen" \
  --min-sdk-version $MIN_SDK --target-sdk-version $TARGET_SDK \
  --version-code "$VERSION_CODE" --version-name "$VERSION_NAME" \
  --auto-add-overlay "$OUT/res.zip"

echo "[3/6] javac"
# AGP generates BuildConfig for the library module; do the same here.
mkdir -p "$OUT/gen/com/hoho/android/usbserial"
cat > "$OUT/gen/com/hoho/android/usbserial/BuildConfig.java" <<'JAVA'
package com.hoho.android.usbserial;

public final class BuildConfig {
    public static final boolean DEBUG = true;
    public static final String LIBRARY_PACKAGE_NAME = "com.hoho.android.usbserial";
    public static final String BUILD_TYPE = "debug";
}
JAVA
find "$ROOT/app/src/main/java" "$ROOT/usbserial/src/main/java" "$OUT/gen" -name '*.java' > "$OUT/sources.txt"
javac -source 8 -target 8 -encoding UTF-8 -Xlint:-options -nowarn \
  -classpath "$ANDROID_JAR" -d "$OUT/classes" @"$OUT/sources.txt"

echo "[4/6] dex"
"$DX" --dex --min-sdk-version=$MIN_SDK --output="$OUT/classes.dex" "$OUT/classes"

echo "[5/6] package + zipalign"
cp "$OUT/base.apk" "$OUT/unsigned.apk"
(cd "$OUT" && zip -q -j unsigned.apk classes.dex)
zipalign -f -p 4 "$OUT/unsigned.apk" "$OUT/aligned.apk"

echo "[6/6] sign"
apksigner sign --ks "$KEYSTORE" --ks-pass pass:android --key-pass pass:android \
  --out "$OUT/app-debug.apk" "$OUT/aligned.apk"
apksigner verify "$OUT/app-debug.apk"
ls -la "$OUT/app-debug.apk"
