# usbserial

Копия библиотеки [usb-serial-for-android](https://github.com/mik3y/usb-serial-for-android) v3.8.1
(лицензия MIT, см. LICENSE.txt). Драйверы USB-serial для FTDI, CP210x, CH34x, PL2303 и CDC-ACM.

Отличия от оригинала:
- убрана зависимость от `androidx.annotation` (аннотация `@IntDef`);
- ссылка на метод в `ProlificSerialDriver` заменена анонимным классом, чтобы код собирался
  старым дексером `dx` из пакетов Ubuntu (см. `tools/build-apk.sh`).
