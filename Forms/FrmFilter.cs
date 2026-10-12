namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 条件筛选：多条件筛选 + 逐行匹配 / SQL 两种方式。
/// </summary>
public class FrmFilter : Form
{
    private readonly RichTextBox _rtbConditions;
    private readonly TextBox _txtKeyColumn;
    private readonly NumericUpDown _numHeaderRow;

    public FrmFilter()
    {
        this.Text = "条件筛选";
        this.Size = new Size(420, 280);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblCond = new Label { Text = "筛选条件（每行一个条件，格式：列字母 值）：", Location = new Point(15, 15), AutoSize = true };
        _rtbConditions = new RichTextBox { Location = new Point(15, 35), Size = new Size(370, 80) };

        var lblKey = new Label { Text = "关键字段列（SQL 模式用）：", Location = new Point(15, 120), AutoSize = true };
        _txtKeyColumn = new TextBox { Location = new Point(150, 118), Width = 80, Text = "A" };

        var lblHeader = new Label { Text = "表头行号：", Location = new Point(15, 150), AutoSize = true };
        _numHeaderRow = new NumericUpDown { Location = new Point(150, 148), Width = 80, Minimum = 1, Maximum = 100, Value = 1 };

        var btnOK = new Button { Text = "逐行匹配筛选", Location = new Point(15, 185), Size = new Size(150, 30) };
        btnOK.Click += BtnRowMatch_Click;

        var btnSQL = new Button { Text = "SQL 筛选", Location = new Point(180, 185), Size = new Size(120, 30) };
        btnSQL.Click += BtnSQL_Click;

        this.Controls.Add(lblCond);
        this.Controls.Add(_rtbConditions);
        this.Controls.Add(lblKey);
        this.Controls.Add(_txtKeyColumn);
        this.Controls.Add(lblHeader);
        this.Controls.Add(_numHeaderRow);
        this.Controls.Add(btnOK);
        this.Controls.Add(btnSQL);
    }

    private void BtnRowMatch_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        var conditions = ParseConditions(_rtbConditions.Text);
        if (conditions.Count == 0) { MessageBox.Show("请输入筛选条件。"); return; }

        int headerRow = (int)_numHeaderRow.Value;

        // Create target sheet
        string tgtName = "Filtered";
        dynamic tgtWs = CreateTargetSheet(app, ws, tgtName);

        // Header row
        for (int c = 1; c <= cols; c++)
            tgtWs.Cells[1, c] = ws.Cells[headerRow, c]?.Value;

        int matchCount = 0;
        using var progress = new ProgressForm();
        progress.Show();

        for (int r = headerRow + 1; r <= rows; r++)
        {
            bool matched = true;
            foreach (var cond in conditions)
            {
                string cellVal = Convert.ToString(ws.Cells[r, cond.ColIndex].Value)?.Trim() ?? "";
                if (!cellVal.Equals(cond.Value, StringComparison.OrdinalIgnoreCase))
                { matched = false; break; }
            }

            if (matched)
            {
                matchCount++;
                for (int c = 1; c <= cols; c++)
                    tgtWs.Cells[matchCount + 1, c] = ws.Cells[r, c]?.Value;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"筛选中… {r}/{rows}");
                progress.SetProgress((int)(r * 100.0 / rows));
                Application.DoEvents();
            }
        }

        tgtWs.Columns.AutoFit();
        progress.SetMessage($"完成！共匹配 {matchCount} 行。");
        System.Threading.Thread.Sleep(1000);
    }

    private void BtnSQL_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        string keyCol = _txtKeyColumn.Text.Trim().ToUpper();
        int headerRow = (int)_numHeaderRow.Value;

        // Get unique key values
        var keyValues = new List<string>();
        for (int r = headerRow + 1; r <= rows; r++)
        {
            string val = Convert.ToString(ws.Cells[r, keyCol].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(val) && !keyValues.Contains(val))
                keyValues.Add(val);
        }

        if (keyValues.Count == 0) { MessageBox.Show("关键字段无数据。"); return; }

        string sheetName = ws.Name?.ToString().Trim() ?? "Sheet1";
        string sql = $"SELECT * FROM [{sheetName}] WHERE [{keyCol}] IS NOT NULL";

        dynamic tgtWs = CreateTargetSheet(app, ws, "Filtered");
        try
        {
            MyTools.CopyToNewSheet(app, sql, tgtWs);
            MyTools.FormatToText(tgtWs);
            tgtWs.Columns.AutoFit();
            MessageBox.Show("SQL 筛选完成！");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"SQL 筛选出错：{ex.Message}");
        }
    }

    private List<(int ColIndex, string Value)> ParseConditions(string text)
    {
        var list = new List<(int, string)>();
        foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = line.Split(new[] { ' ', '\t' }, 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) continue;
            int col = ColumnLetterToIndex(parts[0].Trim().ToUpper());
            if (col > 0)
                list.Add((col, parts[1].Trim()));
        }
        return list;
    }

    private int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col)
        {
            if (!char.IsLetter(c)) return 0;
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private dynamic CreateTargetSheet(dynamic app, dynamic sourceWs, string name)
    {
        // Check if exists
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(ws.Name?.ToString().Trim(), name, StringComparison.OrdinalIgnoreCase))
            {
                ws.UsedRange.Clear();
                return ws;
            }
        }
        var newWs = app.Worksheets.Add(After: sourceWs);
        newWs.Name = name;
        return newWs;
    }
}