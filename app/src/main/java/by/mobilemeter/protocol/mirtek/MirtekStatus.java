package by.mobilemeter.protocol.mirtek;

import java.util.Locale;

/** Decodes the 4-byte status field of a МИРТЕК reply. */
public final class MirtekStatus {
    private MirtekStatus() {
    }

    public static String errorText(int code) {
        switch (code) {
            case 0x00: return "OK";
            case 0x01: return "попытка записи с неверным паролем";
            case 0x02: return "передан недопустимый параметр";
            case 0x03: return "попытка изменения заводского параметра";
            case 0x04: return "неверная длина данных";
            case 0x05: return "интерфейс заблокирован";
            case 0x06: return "запрашиваемых данных нет";
            case 0x07: return "попытка чтения с неверным паролем";
            case 0x08: return "невозможно выполнить команду";
            case 0x09: return "невозможно выполнить команду в данный момент";
            case 0x0A: return "действие уже выполнено";
            case 0xFE: return "пропало питающее напряжение";
            default: return String.format(Locale.US, "неизвестный код 0x%02X", code);
        }
    }

    public static boolean relayPresent(byte[] st) {
        return (st[2] & 0x02) != 0;
    }

    public static boolean loadDisconnected(byte[] st) {
        return (st[2] & 0x08) != 0;
    }

    public static String roleText(int role) {
        boolean relay = (role & 0x01) != 0;
        String base;
        switch (role & 0xFE) {
            case 0x10: base = "1ф 2-элементный активный однонаправленный"; break;
            case 0x20: base = "3ф активный однонаправленный"; break;
            case 0x28: base = "3ф трансформаторный активный однонаправленный"; break;
            case 0x30: base = "1ф 1-элементный активный однонаправленный"; break;
            case 0x38: base = "1ф 2-элементный активный двунаправленный"; break;
            case 0x40: base = "3ф активный двунаправленный"; break;
            case 0x48: base = "3ф трансформаторный активный двунаправленный"; break;
            case 0x50: base = "1ф 1-элементный активный двунаправленный"; break;
            case 0x58: base = "1ф 2-элементный активный однонаправленный с параметрами сети"; break;
            case 0x60: base = "3ф активный однонаправленный с параметрами сети"; break;
            case 0x68: base = "3ф трансформаторный активный однонаправленный с параметрами сети"; break;
            case 0x70: base = "1ф 1-элементный активный однонаправленный с параметрами сети"; break;
            case 0x78: base = "1ф 2-элементный активный двунаправленный с параметрами сети"; break;
            case 0x80: base = "3ф активный двунаправленный с параметрами сети"; break;
            case 0x88: base = "3ф трансформаторный активный двунаправленный с параметрами сети"; break;
            case 0x90: base = "1ф 1-элементный активно-реактивный"; break;
            case 0x98: base = "1ф 2-элементный активно-реактивный"; break;
            case 0xA0: base = "3ф активно-реактивный"; break;
            case 0xA8: base = "3ф трансформаторный активно-реактивный"; break;
            case 0xB0: base = "1ф 1-элементный активный двунаправленный с параметрами сети"; break;
            default: base = "роль не из таблицы";
        }
        return base + (relay ? ", с реле отключения" : "");
    }

    /** Human-readable summary of the whole status field. */
    public static String describe(byte[] st) {
        int role = st[0] & 0xFF;
        int f1 = st[1] & 0xFF;
        int f2 = st[2] & 0xFF;
        int err = st[3] & 0xFF;
        StringBuilder sb = new StringBuilder();
        sb.append(String.format(Locale.US, "роль 0x%02X: %s", role, roleText(role)));
        if ((f2 & 0x02) != 0) {
            sb.append("; реле: ").append((f2 & 0x04) != 0 ? "управления" : "отключения");
            sb.append(", нагрузка ").append((f2 & 0x08) != 0 ? "ОТКЛЮЧЕНА" : "включена");
        } else {
            sb.append("; реле нет");
        }
        if ((f1 & 0x01) != 0) sb.append("; есть новые записи в журнале");
        if ((f1 & 0x02) != 0) sb.append("; вскрыта клеммная крышка");
        if ((f1 & 0x04) != 0) sb.append("; вскрыт корпус");
        if ((f1 & 0x08) != 0) sb.append("; вскрыта крышка модуля связи");
        if ((f1 & 0x10) != 0) sb.append("; воздействие постоянного магнита");
        if ((f1 & 0x20) != 0) sb.append("; воздействие переменного магнитного поля");
        if ((f1 & 0x80) != 0) sb.append("; критический уровень баланса");
        if ((f2 & 0x20) != 0) sb.append("; аварийная ситуация");
        if ((f2 & 0x80) != 0) sb.append("; небаланс токов");
        if ((f2 & 0x01) == 0) sb.append("; заводская перемычка установлена");
        sb.append("; результат: ").append(errorText(err));
        return sb.toString();
    }
}
