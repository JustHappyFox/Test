// ==== БЛОК ОКНА НАРЯДОВ ИЗ ВЫГРУЗКИ (исходник - razrabotka/src/nar_excel_form.cs; вставляется скриптом inject.py) ====
// «Поиск», режим Excel: окно «Наряды…» - номера нарядов (ГодНомер, можно списком): ВЫПОЛНЕН - зелёным, НЕ ВЫПОЛНЕН - красным,
// не найден - серым; ниже - все наряды лицевого выбранной строки. Из карточки (Ctrl+Z) - сразу наряды лицевого.
public class NarExcelForm : Form {
    readonly TextBox input = new TextBox();
    readonly DataGridView found = new DataGridView(), byLs = new DataGridView();
    readonly Label info = new Label(), lsHead = new Label();
    readonly List<int> foundRows = new List<int>();     // строка копии нарядов для строки итога (-1 - не найден)
    NarExcel nar;
    string error;

    public NarExcelForm(string nab, string who) {
        Text = "Наряды - из выгрузки нарядов (только чтение)";
        Font = Look.Base;
        Look.Wear(this);
        StartPosition = FormStartPosition.CenterParent;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Size = new Size(Math.Min(1300, wa.Width - 40), Math.Min(760, wa.Height - 40));
        KeyPreview = true;
        try { nar = NarExcel.Open(); } catch (Exception ex) { error = ex.Message; }

        Panel top = new Panel();
        top.Dock = DockStyle.Top;
        top.Height = 112;
        Label cap = new Label();
        cap.Text = "№ наряда полностью, с годом: 202620936 или 2026-20936. Можно списком - по номеру в строке.";
        cap.SetBounds(12, 8, 900, 20);
        input.Multiline = true;
        input.ScrollBars = ScrollBars.Vertical;
        input.SetBounds(12, 30, 420, 66);
        Button go = new Button();
        go.Text = "Найти";
        go.SetBounds(444, 30, 110, 32);
        go.Click += delegate { Find(); };
        Button xl = new Button();
        xl.Text = "В Excel";
        xl.SetBounds(444, 66, 110, 30);
        xl.Click += delegate { Export(); };
        info.SetBounds(568, 30, 700, 66);
        info.Text = error ?? "Данные: " + nar.About + ". Выполнен - если заполнено «Выполнил» (или дата выполнения).";
        info.ForeColor = error == null ? Look.Hint : Look.Bad;
        top.Controls.AddRange(new Control[] { cap, input, go, xl, info });
        input.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter && e.Control) { e.SuppressKeyPress = true; Find(); } };

        Grid(found);
        found.Dock = DockStyle.Fill;
        Panel bottom = new Panel();
        bottom.Dock = DockStyle.Bottom;
        bottom.Height = Math.Max(220, Size.Height / 2 - 80);
        lsHead.Dock = DockStyle.Top;
        lsHead.Height = 24;
        lsHead.Padding = new Padding(8, 4, 8, 0);
        lsHead.Font = Look.Bold;
        lsHead.Text = "Наряды лицевого - выберите строку итога";
        Grid(byLs);
        byLs.Dock = DockStyle.Fill;
        bottom.Controls.Add(byLs);
        bottom.Controls.Add(lsHead);
        Controls.Add(found);
        Controls.Add(bottom);
        Controls.Add(top);
        found.SelectionChanged += delegate { ShowLs(); };
        KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; Close(); } };
        if (nab != null && nab.Length > 0) Shown += delegate { ShowNab(nab, who); };
    }

    static void Grid(DataGridView g) {
        g.ReadOnly = true;
        g.AllowUserToAddRows = false;
        g.AllowUserToDeleteRows = false;
        g.RowHeadersVisible = false;
        g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        g.MultiSelect = false;
        g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        foreach (string[] c in NarExcel.Columns) {
            DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
            col.HeaderText = c[0];
            col.SortMode = DataGridViewColumnSortMode.NotSortable;
            col.FillWeight = c[0] == "Адрес" || c[0] == "Ф.И.О." || c[0] == "Работа" || c[0] == "Примечание" ? 150 : 70;
            g.Columns.Add(col);
        }
        Look.Grid(g);
        CopyText.Grid(g);
    }

    void Color(DataGridView g, int r, int row) {
        System.Drawing.Color back = row < 0 ? Look.RowOther : nar.Done(row) ? Look.RowGood : Look.RowBad;
        g.Rows[r].DefaultCellStyle.BackColor = back;
    }

    public void SetInput(string text) { input.Text = text; }

    // номера из поля - строки итога
    public void Find() {
        found.Rows.Clear();
        foundRows.Clear();
        if (nar == null) { info.Text = error; return; }
        int ok = 0, no = 0, miss = 0;
        foreach (string raw in input.Text.Split('\n')) {
            string line = raw.Trim();
            if (line.Length == 0) continue;
            string num = NarExcel.Number(line);
            if (num == null) {
                int r0 = found.Rows.Add(new object[] { line, "впишите полностью: год и номер (2026xxxxx)" });
                found.Rows[r0].DefaultCellStyle.BackColor = Look.RowOther;
                foundRows.Add(-1);
                miss++;
                continue;
            }
            List<int> rows = nar.ByNumber(num);
            if (rows.Count == 0) {
                int r0 = found.Rows.Add(new object[] { num, "не найден" });
                found.Rows[r0].DefaultCellStyle.BackColor = Look.RowOther;
                foundRows.Add(-1);
                miss++;
                continue;
            }
            foreach (int row in rows) {
                int r = found.Rows.Add(nar.Cells(row));
                Color(found, r, row);
                foundRows.Add(row);
                if (nar.Done(row)) ok++; else no++;
            }
        }
        info.Text = "Выполнено: " + ok + ", не выполнено: " + no + ", не найдено: " + miss + ". Данные: " + nar.About + ".";
        info.ForeColor = Look.Hint;
        if (found.Rows.Count > 0) found.CurrentCell = found.Rows[0].Cells[0];
        ShowLs();
    }

    void ShowLs() {
        if (nar == null || found.CurrentRow == null || found.CurrentRow.Index >= foundRows.Count) return;
        int row = foundRows[found.CurrentRow.Index];
        if (row < 0) { byLs.Rows.Clear(); lsHead.Text = "Наряды лицевого - выберите найденный наряд"; return; }
        ShowNab(nar.F(row, "NAB"), (nar.F(row, "FIO") + " " + nar.Address(row)).Trim());
    }

    // все наряды лицевого (новые сверху)
    public void ShowNab(string nab, string who) {
        byLs.Rows.Clear();
        if (nar == null) { lsHead.Text = error; return; }
        List<int> rows = nar.ByNab(nab);
        foreach (int row in rows) { int r = byLs.Rows.Add(nar.Cells(row)); Color(byLs, r, row); }
        lsHead.Text = "Наряды лицевого " + nab + (who != null && who.Length > 0 ? " ▶ " + who : "") + ": " + (rows.Count == 0 ? "в выгрузке нарядов нет" : rows.Count.ToString());
    }

    public int FoundCount { get { return found.Rows.Count; } }
    public int LsCount { get { return byLs.Rows.Count; } }
    public string LsText { get { return lsHead.Text; } }
    public string InfoText { get { return info.Text; } }

    void Export() {
        if (found.Rows.Count == 0) { MessageBox.Show(this, "Сначала найдите наряды.", Text); return; }
        List<string[]> rows = new List<string[]>();
        List<string> styles = new List<string>();
        for (int i = 0; i < found.Rows.Count; i++) {
            string[] c = new string[NarExcel.Columns.Length];
            for (int k = 0; k < c.Length; k++) { object v = found.Rows[i].Cells[k].Value; c[k] = v == null ? "" : v.ToString(); }
            rows.Add(c);
            int row = i < foundRows.Count ? foundRows[i] : -1;
            styles.Add(row < 0 ? "y" : nar.Done(row) ? "t" : "n");
        }
        string[] headers = new string[NarExcel.Columns.Length];
        for (int k = 0; k < headers.Length; k++) headers[k] = NarExcel.Columns[k][0];
        try { Spreadsheet.Open("Наряды", headers, rows, styles); }
        catch (Exception ex) { MessageBox.Show(this, "В Excel не получилось: " + ex.Message, Text); }
    }
}
// ==== КОНЕЦ БЛОКА ОКНА НАРЯДОВ ИЗ ВЫГРУЗКИ ====
