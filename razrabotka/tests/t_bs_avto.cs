// Проверка 25.11: автоподстановки пакета юр лица (вымышленные данные) - строка объекта из бумаг разбирается сама (наименование, адрес
// по частям, код), вид объекта по наименованию, адрес руками -> по частям, место включения из «Питание от», схема учёта по счётчику,
// «+» прибора учёта (копия строки и строка мощности), длина «44 м», абонент из памяти, «новый по образцу»
using System; using System.Windows.Forms; using System.Reflection; using System.IO; using System.Collections;
public static class T {
    static BindingFlags BF = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public;
    static object F(object o, string n) { return o.GetType().GetField(n, BF).GetValue(o); }
    static object Call(object o, string n, params object[] a) {
        foreach (MethodInfo m in o.GetType().GetMethods(BF)) if (m.Name == n && m.GetParameters().Length == a.Length) return m.Invoke(o, a);
        throw new Exception("нет метода " + n);
    }
    static int bad = 0;
    static void Eq(string what, string got, string want) {
        bool ok = got == want;
        if (!ok) bad++;
        Console.WriteLine((ok ? "ОК   " : "ОШИБКА ") + what + ": «" + got + "»" + (ok ? "" : " (ждали «" + want + "»)"));
    }
    // ввод в ячейку строки «как руками» (TextChanged -> Edited)
    static void Type(object rowList, int row, int col, string v) {
        IList rows = (IList)F(rowList, "rows");
        Control[] cells = (Control[])F(rows[row], "Cells");
        cells[col].Text = v;
    }
    [STAThread] public static void Main(string[] a) {
        string here = a[0], png = a.Length > 1 ? a[1] : null;
        Application.EnableVisualStyles();
        StartMenu m = new StartMenu(here, "проверка");
        Timer t = new Timer(); t.Interval = 1500; int step = 0;
        t.Tick += delegate {
            step++;
            BsForm b = (BsForm)F(m, "bs");
            if (step == 1) { m.OpenBs(m); return; }
            object pu = F(b, "pu"), moshn = F(b, "moshn"), linii = F(b, "linii");
            if (step == 2) {
                Console.WriteLine("--- строка объекта целиком (вставили)");
                b.SetV("obekt", "«Установка зарядных станций для электромобилей по адресу: Минская обл., Минский р-н, Пробный с/с, д. Тестовка, ул. Лесная, 5 (сооружение тестовое, код 1 23 45)»");
            } else if (step == 3) {
                Eq("наименование", b.V("obekt"), "Установка зарядных станций для электромобилей");
                Eq("код", b.V("kod"), "сооружение тестовое, код 1 23 45");
                Eq("область", b.V("obl"), "Минская обл.");
                Eq("район", b.V("rayon"), "Минский р-н");
                Eq("сельсовет", b.V("selsovet"), "Пробный с/с");
                Eq("н.п.", b.V("np_tip") + " " + b.V("np"), "д. Тестовка");
                Eq("улица", b.V("ulica_tip") + " " + b.V("ulica"), "ул. Лесная");
                Eq("дом", b.V("dom"), "5");
                Eq("адрес", b.V("adres"), "Минская обл., Минский р-н, Пробный с/с, д. Тестовка, ул. Лесная, д. 5");
                Eq("коротко", b.V("adres_kr"), "д. Тестовка, ул. Лесная, д. 5");
                Eq("объект в справке", b.V("obekt_full"), "«Установка зарядных станций для электромобилей по адресу: Минская обл., Минский р-н, Пробный с/с, д. Тестовка, ул. Лесная, д. 5 (сооружение тестовое, код 1 23 45)»");
                Eq("вид объекта", b.V("vid_obekta"), "Зарядная станция электромобилей");
                Eq("характер деятельности", b.V("deyat"), "электрическая энергия используется для зарядки электромобилей");
                Eq("точка учёта", ((RowList)pu).Cell(0, 0) + " / " + ((RowList)moshn).Cell(0, 0), "ЭЗС / ЭЗС");
                Console.WriteLine("   подсказка: " + ((Label)F(b, "obNote")).Text);
                if (png != null) Snap(b, png + "_1.png", 1);

                Console.WriteLine("--- строка без адреса и кода (набрали руками), БС");
                b.Clear();
                b.SetV("obekt", "Возведение мобильной базовой станции БС-0002");
                Call(b, "ObjectLine");
                Eq("наименование", b.V("obekt"), "Возведение мобильной базовой станции БС-0002");
                Eq("код (нет в бумагах - пусто)", b.V("kod"), "");
                Eq("вид объекта", b.V("vid_obekta"), "Базовая станция связи");
                Eq("точка учёта", ((RowList)pu).Cell(0, 0), "ЩУЭ");

                Console.WriteLine("--- адрес вставили в «Адрес полностью»");
                b.SetV("adres", "Минская область, Минский район, г.п. Пробное, пер. Садовый в районе дома №3");
                Call(b, "AdresByHand");
                Eq("н.п.", b.V("np_tip") + " " + b.V("np"), "г.п. Пробное");
                Eq("улица", b.V("ulica_tip") + " " + b.V("ulica"), "пер. Садовый");
                Eq("дом", b.V("dom"), "в районе дома №3");
                Eq("сельсовет", b.V("selsovet"), "");
                Eq("коротко", b.V("adres_kr"), "г.п. Пробное, пер. Садовый в районе дома №3");
                Eq("пакет", b.V("paket"), "Пробное БС-0002");
                b.SetV("adres", "аг. Макетный, 12");
                Call(b, "AdresByHand");
                Eq("без улицы: коротко", b.V("adres_kr"), "аг. Макетный, д. 12");

                Console.WriteLine("--- место включения из «Питание от»");
                b.SetV("pitanie", "КТПБ-0009");
                Eq("место", b.V("mesto"), "КТПБ-0009, РУ-0,4 кВ");

                Console.WriteLine("--- схема учёта по счётчику");
                Eq("сначала", b.V("pot_uchet"), "прямой");
                Type(pu, 0, 4, "5(10)");
                Eq("ток 5(10) -> через ТТ", b.V("pot_uchet"), "через ТТ");
                Type(pu, 0, 8, "ТОП-0,66 150/5");
                Eq("К расч", ((RowList)pu).Cell(0, 10), "30");
                Type(pu, 0, 4, "5(60)");
                Eq("ток 5(60) -> прямой", b.V("pot_uchet"), "прямой");
                Eq("прямой: ТТ «-», К 1", ((RowList)pu).Cell(0, 8) + " " + ((RowList)pu).Cell(0, 10), "- 1");
                b.SetV("pot_uchet", "через ТТ");
                Eq("через ТТ: ток 5(60) свой - не трогать", ((RowList)pu).Cell(0, 4), "5(60)");
                Type(pu, 0, 4, "5(100)");
                b.SetV("pot_uchet", "прямой"); b.SetV("pot_uchet", "через ТТ");
                Eq("через ТТ: заготовка 5(100) -> 5(10)", ((RowList)pu).Cell(0, 4), "5(10)");

                Console.WriteLine("--- «+» в приборах учёта");
                Type(pu, 0, 1, "Тестовый-3Ф");
                Type(pu, 0, 2, "11111111");
                Type(moshn, 0, 1, "20");
                ((Button)F(pu, "add")).PerformClick();
                Eq("строк ПУ / мощности", ((RowList)pu).Count + " / " + ((RowList)moshn).Count, "2 / 2");
                Eq("новая строка - как первая", ((RowList)pu).Cell(1, 0) + "|" + ((RowList)pu).Cell(1, 1) + "|" + ((RowList)pu).Cell(1, 2) + "|" + ((RowList)pu).Cell(1, 4) + "|" + ((RowList)pu).Cell(1, 5),
                    "ЩУЭ|Тестовый-3Ф||5(10)|000000,00");
                Eq("мощность новой строки", ((RowList)moshn).Cell(1, 0) + "|" + ((RowList)moshn).Cell(1, 1) + "|" + ((RowList)moshn).Cell(1, 3), "ЩУЭ||24 ч/сут.");

                Console.WriteLine("--- длина линии в метрах");
                Type(linii, 0, 1, "АВБбШв 4х95");
                Type(linii, 0, 2, "44 м");
                Eq("L, км", ((RowList)linii).Cell(0, 2), "0,044");
                Eq("Rл", ((RowList)linii).Cell(0, 4), "0,01434");
                Eq("линия в акте", b.V("liniya"), "КЛ-0,4 кВ АВБбШв 4х95 L=44 м - на балансе абонента.");

                Console.WriteLine("--- абонент из памяти, сохранение, новый по образцу");
                b.SetV("abonent", "ООО «Тестовая связь»"); b.SetV("dog_nom", "12345"); b.SetV("dog_data", "01.02.2015"); b.SetV("tel", "+375 29 000-00-00");
                Type(pu, 1, 2, "22222222");
                b.SetV("plomba_klemm", "1111111");
                string dir = b.Save(false, false);
                Console.WriteLine("   сформирован: " + (dir != null ? Path.GetFileName(dir) : "НЕТ"));
                b.Clear();
                b.SetV("abonent", "ООО «Тестовая связь»");
                Eq("договор из памяти", b.V("dog_nom") + " от " + b.V("dog_data") + ", " + b.V("tel"), "12345 от 01.02.2015, +375 29 000-00-00");
                string file = Path.Combine(BsForm.DataFolder(here), Path.GetFileName(dir) + ".txt");
                b.LoadSample(file);
                Eq("образец: абонент и договор", b.V("abonent") + " " + b.V("dog_nom"), "ООО «Тестовая связь» 12345");
                Eq("образец: объект, адрес, питание - пусто", b.V("obekt") + b.V("np") + b.V("dom") + b.V("pitanie") + b.V("mesto") + b.V("plomba_klemm"), "");
                Eq("образец: счётчики - тип есть, № нет", ((RowList)pu).Cell(0, 1) + "|" + ((RowList)pu).Cell(0, 2) + "|" + ((RowList)pu).Cell(1, 2), "Тестовый-3Ф||");
                Eq("образец: схема учёта", b.V("pot_uchet"), "через ТТ");
                Eq("образец: мощность осталась", ((RowList)moshn).Cell(0, 1), "20");
                Eq("образец: длина линии - заново", ((RowList)linii).Cell(0, 1) + "|" + ((RowList)linii).Cell(0, 2), "АВБбШв 4х95|");
                Eq("образец: дата акта - сегодня", b.V("data_akt"), DateTime.Today.ToString("dd.MM.yyyy"));
                Console.WriteLine("   заголовок: " + ((Label)F(b, "headSub")).Text);
                if (png != null) Snap(b, png + "_2.png", 0);
                Console.WriteLine("--- разрядность по типу счётчика (26.11): CE318 - 6+2, CE208 - 8; буквы латинские или русские");
                b.Clear();
                Type(pu, 0, 1, "CE318BY S31 543 JAVZ");
                Eq("CE318: разр., показания", ((RowList)pu).Cell(0, 11) + " " + ((RowList)pu).Cell(0, 5), "6+2 000000,00");
                Type(pu, 0, 1, "СЕ 208 BY S7 145.1");
                Eq("СЕ 208 (русские буквы): разр., показания", ((RowList)pu).Cell(0, 11) + " " + ((RowList)pu).Cell(0, 5), "8 00000000");
                Console.WriteLine(bad == 0 ? "ВСЁ ОК" : "ОШИБОК: " + bad);
            } else { t.Stop(); Environment.Exit(0); }
        };
        t.Start();
        Application.Run(m);
    }
    static void Snap(BsForm b, string file, int card) {
        Panel content = (Panel)F(b, "content");
        IList cards = (IList)F(b, "cards");
        Control p = (Control)F(cards[card], "Panel");
        content.AutoScrollPosition = new System.Drawing.Point(0, p.Top - content.AutoScrollPosition.Y - 4);
        Application.DoEvents();
        using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(b.Width, b.Height)) {
            b.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, b.Width, b.Height));
            bmp.Save(file);
        }
    }
}
