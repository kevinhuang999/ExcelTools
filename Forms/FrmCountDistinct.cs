namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Running distinct count: for each row, count how many distinct values have
/// appeared so far in the chosen column.
///
/// Step flow matches the original add-in: open, press the action button, watch
/// the progress bar. Both columns are pre-filled from the last run.
/// </summary>
public class FrmCountDistinct : Form
{
    private const string KeyCountColumn = "countDistinct.countColumn";
    private const string KeyOutputColumn = "countDistinct.outputColumn";

    private readonly TextBox _txtCountColumn;
    private readonly TextBox _txtOutputColumn;
    private readonly ProgressBar _progressBar;
    private readonly Label _lblStatus;

    public FrmCountDistinct()
    {
        Text = Loc.T("累计去重计数", "Count Distinct");
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            Loc.T("对指定列逐行统计：截至当前行，一共出现过多少个不重复的值。\n结果写入输出列，输出列必须在数据区右侧之外。", "For each row, count how many distinct values have appeared so far.\nResults go to the output column, which must sit outside the data range."));

        _txtCountColumn = layout.AddField(Loc.T("统计列:", "Count Column:"), Settings.Get(KeyCountColumn, "A"));
        _txtOutputColumn = layout.AddField(Loc.T("输出列:", "Output Column:"), Settings.Get(KeyOutputColumn, "Z"));
        _progressBar = layout.AddProgressBar();
        _lblStatus = layout.AddStatusLabel();

        var buttons = layout.Finish((Loc.T("统计", "Sum"), true), (Loc.T("取消", "Cancel"), false));
        buttons[0].Click += BtnSum_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }

    private void BtnSum_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

        string countColumn = _txtCountColumn.Text.Trim().ToUpperInvariant();
        string outputColumn = _txtOutputColumn.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(countColumn) || !IsColumnLetter(outputColumn))
        {
            MessageBox.Show(Loc.T("列必须是列字母，例如 A 或 AB。", "Columns must be column letters such as A or AB."), Loc.T("提示", "Warning"));
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show(Loc.T("空表，或无法清理特殊符号！", "Empty table, or unable to clean up special symbols!"), Loc.T("提示", "Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show(Loc.T("数据不足。", "Not enough data."), Loc.T("提示", "Warning")); return; }

        int countColumnIndex = ColumnLetterToIndex(countColumn);
        int outputColumnIndex = ColumnLetterToIndex(outputColumn);

        if (outputColumnIndex <= cols)
        {
            MessageBox.Show(Loc.T("输出列必须在数据范围之外。", "The output column must be outside the data range."), Loc.T("提示", "Warning"));
            return;
        }

        Settings.SetMany((KeyCountColumn, countColumn), (KeyOutputColumn, outputColumn));

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ws.Cells[1, outputColumnIndex] = "DISTINCT_COUNT";

        _progressBar.Minimum = 0;
        _progressBar.Maximum = rows;
        _progressBar.Value = 0;

        for (int r = 2; r <= rows; r++)
        {
            string value = Convert.ToString(ws.Cells[r, countColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value)) seen.Add(value);

            ws.Cells[r, outputColumnIndex] = seen.Count;

            _lblStatus.Text = r + "/" + rows;
            _progressBar.Value = r;
            if (r % 200 == 0) Application.DoEvents();
        }

        ws.Columns.AutoFit();
        _progressBar.Value = rows;
        _lblStatus.Text = "";
        app.ActiveWorkbook.Save();

        if (DialogResult.OK == MessageBox.Show(Loc.T("完成！", "Success!"), Loc.T("累计去重计数", "Count Distinct")))
        {
            Close();
        }
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

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
