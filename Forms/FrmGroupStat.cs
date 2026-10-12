namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Statistics: group the active sheet by one column and produce a
/// category-by-period cross table on its own sheet.
///
/// Step flow matches the original add-in: open, press the action button. All
/// three column fields are pre-filled from the last run.
/// </summary>
public class FrmGroupStat : Form
{
    private const string KeyGroupColumn = "groupStat.groupColumn";
    private const string KeyPeriodColumn = "groupStat.periodColumn";
    private const string KeyTargetSheet = "groupStat.targetSheet";

    private readonly TextBox _txtGroupColumn;
    private readonly TextBox _txtPeriodColumn;
    private readonly TextBox _txtTargetSheet;

    public FrmGroupStat()
    {
        Text = "分组统计";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "按分类列分组、按期间列拆分，统计每组在每个期间的记录数，\n生成「分类 × 期间」交叉汇总表。");

        _txtGroupColumn = layout.AddField("分类列:", Settings.Get(KeyGroupColumn, "A"));
        _txtPeriodColumn = layout.AddField("期间列:", Settings.Get(KeyPeriodColumn, "B"));
        _txtTargetSheet = layout.AddField("输出工作表:", Settings.Get(KeyTargetSheet, "Statistics"));

        var buttons = layout.Finish(("统计", true), ("取消", false));
        buttons[0].Click += BtnStatistics_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }

    private void BtnStatistics_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string groupColumn = _txtGroupColumn.Text.Trim().ToUpperInvariant();
        string periodColumn = _txtPeriodColumn.Text.Trim().ToUpperInvariant();
        string targetSheetName = _txtTargetSheet.Text.Trim();

        if (!IsColumnLetter(groupColumn) || !IsColumnLetter(periodColumn))
        {
            MessageBox.Show("列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }
        if (string.IsNullOrEmpty(targetSheetName))
        {
            MessageBox.Show("请输入输出工作表名称。", "提示");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int _))
        {
            MessageBox.Show("空表，或无法清理特殊符号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("数据不足。", "提示"); return; }
        app.ActiveWorkbook.Save();

        Settings.SetMany(
            (KeyGroupColumn, groupColumn),
            (KeyPeriodColumn, periodColumn),
            (KeyTargetSheet, targetSheetName));

        int groupColumnIndex = ColumnLetterToIndex(groupColumn);
        int periodColumnIndex = ColumnLetterToIndex(periodColumn);

        var groups = new List<string>();
        var counts = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        using var progress = new ProgressForm();
        progress.Show();

        for (int r = 2; r <= rows; r++)
        {
            string group = Convert.ToString(ws.Cells[r, groupColumnIndex].Value)?.Trim() ?? "";
            string period = Convert.ToString(ws.Cells[r, periodColumnIndex].Value)?.Trim() ?? "";
            if (string.IsNullOrEmpty(group)) continue;

            if (!counts.TryGetValue(group, out var byPeriod))
            {
                groups.Add(group);
                byPeriod = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                counts[group] = byPeriod;
            }

            if (!string.IsNullOrEmpty(period))
            {
                byPeriod.TryGetValue(period, out int current);
                byPeriod[period] = current + 1;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"正在读取… {r}/{rows}");
                progress.SetProgress((int)(r * 50.0 / rows));
                Application.DoEvents();
            }
        }

        if (groups.Count == 0)
        {
            MessageBox.Show("未找到有效的分类数据。", "提示");
            return;
        }

        var periods = new List<string>();
        foreach (var byPeriod in counts.Values)
            foreach (string p in byPeriod.Keys)
                if (!periods.Contains(p)) periods.Add(p);
        periods.Sort(StringComparer.OrdinalIgnoreCase);

        dynamic? targetSheet = FindSheet(app, targetSheetName);
        if (targetSheet == null)
        {
            targetSheet = app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
            targetSheet.Name = targetSheetName;
        }
        else
        {
            targetSheet.UsedRange.Clear();
        }

        targetSheet.Cells[1, 1] = groupColumn;
        for (int c = 0; c < periods.Count; c++)
            targetSheet.Cells[1, c + 2] = periods[c];
        targetSheet.Cells[1, periods.Count + 2] = "Total";

        for (int g = 0; g < groups.Count; g++)
        {
            targetSheet.Cells[g + 2, 1] = groups[g];
            var byPeriod = counts[groups[g]];
            int rowTotal = 0;
            for (int c = 0; c < periods.Count; c++)
            {
                byPeriod.TryGetValue(periods[c], out int value);
                targetSheet.Cells[g + 2, c + 2] = value;
                rowTotal += value;
            }
            targetSheet.Cells[g + 2, periods.Count + 2] = rowTotal;
            progress.SetProgress(50 + (int)((g + 1) * 50.0 / groups.Count));
            Application.DoEvents();
        }

        targetSheet.Columns.AutoFit();
        progress.SetProgress(100);
        progress.SetMessage("完成。");
        app.ActiveWorkbook.Save();

        Close();

        if (DialogResult.Yes == MessageBox.Show("已完成并保存，现在关闭吗？", "确认", MessageBoxButtons.YesNo))
        {
            app.Quit();
        }
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

    private static dynamic? FindSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                return ws;
        }
        return null;
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
