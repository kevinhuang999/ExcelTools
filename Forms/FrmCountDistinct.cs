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
        Text = "累计去重计数";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "对指定列逐行统计：截至当前行，一共出现过多少个不重复的值。\n结果写入输出列，输出列必须在数据区右侧之外。");

        _txtCountColumn = layout.AddField("统计列:", Settings.Get(KeyCountColumn, "A"));
        _txtOutputColumn = layout.AddField("输出列:", Settings.Get(KeyOutputColumn, "Z"));
        _progressBar = layout.AddProgressBar();
        _lblStatus = layout.AddStatusLabel();

        var buttons = layout.Finish(("统计", true), ("取消", false));
        buttons[0].Click += BtnSum_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }

    private void BtnSum_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string countColumn = _txtCountColumn.Text.Trim().ToUpperInvariant();
        string outputColumn = _txtOutputColumn.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(countColumn) || !IsColumnLetter(outputColumn))
        {
            MessageBox.Show("列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("空表，或无法清理特殊符号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("数据不足。", "提示"); return; }

        int countColumnIndex = ColumnLetterToIndex(countColumn);
        int outputColumnIndex = ColumnLetterToIndex(outputColumn);

        if (outputColumnIndex <= cols)
        {
            MessageBox.Show("输出列必须在数据范围之外。", "提示");
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

        if (DialogResult.OK == MessageBox.Show("完成！", "Count Distinct"))
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
