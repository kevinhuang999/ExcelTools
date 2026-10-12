namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Filter: copy matching rows into a new sheet, either by row-by-row matching
/// against typed conditions, or via a SQL query over the sheet.
///
/// Step flow matches the original add-in: type the conditions, press OK or SQL.
/// The key column and header row are pre-filled from the last run.
/// </summary>
public class FrmFilter : Form
{
    private const string KeyColumn = "filter.column";
    private const string KeyHeaderRow = "filter.headerRow";

    private readonly RichTextBox _rtbConditions;
    private readonly TextBox _txtColumn;
    private readonly TextBox _txtHeaderRow;

    public FrmFilter()
    {
        Text = Loc.T("条件筛选", "Filter");
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            Loc.T("把符合条件的行复制到新的 Filtered 工作表。\n每行写一个条件，格式：列字母 + 空格 + 值，例如  A 1001", "Copy matching rows into a new Filtered sheet.\nOne condition per line: column letter then value, e.g.  A 1001"));

        _rtbConditions = layout.AddMultilineField(Loc.T("筛选条件:", "Filter Fields:"), 110);
        _txtColumn = layout.AddField(Loc.T("关键列:", "Column:"), Settings.Get(KeyColumn, "A"));
        _txtHeaderRow = layout.AddField(Loc.T("表头行号:", "HeaderRowNo:"), Settings.Get(KeyHeaderRow, "1"));

        var buttons = layout.Finish((Loc.T("筛选", "Filter"), true), (Loc.T("SQL 查询", "SQL"), false), (Loc.T("取消", "Cancel"), false));
        buttons[0].Click += BtnOK_Click;
        buttons[1].Click += BtnSQL_Click;
        buttons[2].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[2];

        UiTheme.Apply(this);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show(Loc.T("空表，或无法清理特殊符号！", "Empty table, or unable to clean up special symbols!"), Loc.T("提示", "Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show(Loc.T("数据不足。", "Not enough data."), Loc.T("提示", "Warning")); return; }

        var conditions = ParseConditions(_rtbConditions.Text);
        if (conditions.Count == 0) { MessageBox.Show(Loc.T("请至少输入一个筛选条件。", "Please enter at least one filter condition."), Loc.T("提示", "Warning")); return; }

        int headerRow = ParseHeaderRow();

        Settings.SetMany((KeyColumn, _txtColumn.Text.Trim().ToUpperInvariant()), (KeyHeaderRow, headerRow.ToString()));

        dynamic targetSheet = CreateTargetSheet(app, ws, "Filtered");

        for (int c = 1; c <= cols; c++)
            targetSheet.Cells[1, c] = ws.Cells[headerRow, c]?.Value;

        using var progress = new ProgressForm();
        progress.Show();

        int matched = 0;
        for (int r = headerRow + 1; r <= rows; r++)
        {
            bool allMatch = true;
            foreach (var condition in conditions)
            {
                string cellValue = Convert.ToString(ws.Cells[r, condition.ColumnIndex].Value)?.Trim() ?? "";
                if (!cellValue.Equals(condition.Value, StringComparison.OrdinalIgnoreCase))
                {
                    allMatch = false;
                    break;
                }
            }

            if (allMatch)
            {
                matched++;
                for (int c = 1; c <= cols; c++)
                    targetSheet.Cells[matched + 1, c] = ws.Cells[r, c]?.Value;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage(Loc.T($"正在筛选… {r}/{rows}", "Filtering... {r}/{rows}"));
                progress.SetProgress((int)(r * 100.0 / rows));
                Application.DoEvents();
            }
        }

        targetSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage(Loc.T($"完成。 {matched} row(s) matched.", "Done. {matched} row(s) matched."));
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show(Loc.T($"已完成并保存！ {matched} row(s) matched.", "Finished and saved! {matched} row(s) matched."), Loc.T("条件筛选", "Filter"));
        Close();
    }

    private void BtnSQL_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show(Loc.T("空表，或无法清理特殊符号！", "Empty table, or unable to clean up special symbols!"), Loc.T("提示", "Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show(Loc.T("数据不足。", "Not enough data."), Loc.T("提示", "Warning")); return; }

        string keyColumn = _txtColumn.Text.Trim().ToUpperInvariant();
        if (!IsColumnLetter(keyColumn))
        {
            MessageBox.Show(Loc.T("列必须是列字母，例如 A 或 AB。", "Column must be a column letter such as A or AB."), Loc.T("提示", "Warning"));
            return;
        }

        Settings.SetMany((KeyColumn, keyColumn), (KeyHeaderRow, ParseHeaderRow().ToString()));

        string sheetName = Convert.ToString(ws.Name)?.Trim() ?? "Sheet1";
        string sql = $"SELECT * FROM [{sheetName}$] WHERE [{keyColumn}] IS NOT NULL";

        dynamic targetSheet = CreateTargetSheet(app, ws, "Filtered");
        try
        {
            MyTools.CopyToNewSheet(app, sql, targetSheet);
            MyTools.FormatToText(targetSheet);
            targetSheet.Columns.AutoFit();
            app.ActiveWorkbook.Save();

            MessageBox.Show(Loc.T("已完成并保存！", "Finished and saved!"), Loc.T("条件筛选", "Filter"));
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(Loc.T($"SQL 筛选失败：{ex.Message}", "SQL filter failed: {ex.Message}"), Loc.T("错误", "Error"));
        }
    }

    private int ParseHeaderRow()
    {
        return int.TryParse(_txtHeaderRow.Text.Trim(), out int value) && value >= 1 ? value : 1;
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

    private List<(int ColumnIndex, string Value)> ParseConditions(string text)
    {
        var list = new List<(int, string)>();
        foreach (string line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string[] parts = line.Split(new[] { ' ', '\t' }, 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) continue;

            int columnIndex = ColumnLetterToIndex(parts[0].Trim().ToUpperInvariant());
            if (columnIndex > 0)
                list.Add((columnIndex, parts[1].Trim()));
        }
        return list;
    }

    private static dynamic CreateTargetSheet(dynamic app, dynamic sourceSheet, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
            {
                ws.UsedRange.Clear();
                return ws;
            }
        }

        dynamic newSheet = app.Worksheets.Add(After: sourceSheet);
        newSheet.Name = name;
        return newSheet;
    }

    private static int ColumnLetterToIndex(string column)
    {
        int result = 0;
        foreach (char c in column)
        {
            if (!char.IsLetter(c)) return 0;
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }
}
