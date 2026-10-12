namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Match two sheets on a key column and list rows whose comparison column
/// differs. Results go to a "DiffResult" sheet.
///
/// Step flow: open, press the action button. All fields are pre-filled from the
/// last run.
/// </summary>
public class FrmTimeCompare : Form
{
    private const string KeySheet1 = "diff.sheet1";
    private const string KeySheet2 = "diff.sheet2";
    private const string KeyMatch1 = "diff.match1";
    private const string KeyCompare1 = "diff.compare1";
    private const string KeyMatch2 = "diff.match2";
    private const string KeyCompare2 = "diff.compare2";

    private readonly TextBox _txtSheet1;
    private readonly TextBox _txtSheet2;
    private readonly TextBox _txtMatchColumn1;
    private readonly TextBox _txtCompareColumn1;
    private readonly TextBox _txtMatchColumn2;
    private readonly TextBox _txtCompareColumn2;

    public FrmTimeCompare()
    {
        Text = "差异比对";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "按匹配列关联两张表，列出比对列取值不一致的行。\n结果写入新的 DiffResult 工作表。");

        _txtSheet1 = layout.AddField("表1:", Settings.Get(KeySheet1, ""));
        _txtMatchColumn1 = layout.AddField("表1匹配列:", Settings.Get(KeyMatch1, "A"));
        _txtCompareColumn1 = layout.AddField("表1比对列:", Settings.Get(KeyCompare1, "B"));
        _txtSheet2 = layout.AddField("表2:", Settings.Get(KeySheet2, ""));
        _txtMatchColumn2 = layout.AddField("表2匹配列:", Settings.Get(KeyMatch2, "A"));
        _txtCompareColumn2 = layout.AddField("表2比对列:", Settings.Get(KeyCompare2, "B"));

        var buttons = layout.Finish(("比对", true), ("取消", false));
        buttons[0].Click += BtnMatch_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }


    private void BtnMatch_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string sheet1Name = _txtSheet1.Text.Trim();
        string sheet2Name = _txtSheet2.Text.Trim();
        string matchColumn1 = _txtMatchColumn1.Text.Trim().ToUpperInvariant();
        string compareColumn1 = _txtCompareColumn1.Text.Trim().ToUpperInvariant();
        string matchColumn2 = _txtMatchColumn2.Text.Trim().ToUpperInvariant();
        string compareColumn2 = _txtCompareColumn2.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sheet1Name) || string.IsNullOrEmpty(sheet2Name))
        {
            MessageBox.Show("请输入两个工作表名称。", "提示");
            return;
        }
        if (!IsColumnLetter(matchColumn1) || !IsColumnLetter(compareColumn1) ||
            !IsColumnLetter(matchColumn2) || !IsColumnLetter(compareColumn2))
        {
            MessageBox.Show("列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }

        dynamic? sheet1 = FindSheet(app, sheet1Name);
        dynamic? sheet2 = FindSheet(app, sheet2Name);
        if (sheet1 == null) { MessageBox.Show($"找不到工作表：{sheet1Name}", "提示"); return; }
        if (sheet2 == null) { MessageBox.Show($"找不到工作表：{sheet2Name}", "提示"); return; }

        MyTools.ClearSpecialSymbols(sheet1, out int rows1, out int _);
        MyTools.ClearSpecialSymbols(sheet2, out int rows2, out int _);
        if (rows1 < 2 || rows2 < 2) { MessageBox.Show("数据不足。", "提示"); return; }

        Settings.SetMany(
            (KeySheet1, sheet1Name), (KeySheet2, sheet2Name),
            (KeyMatch1, matchColumn1), (KeyCompare1, compareColumn1),
            (KeyMatch2, matchColumn2), (KeyCompare2, compareColumn2));

        int matchIndex1 = ColumnLetterToIndex(matchColumn1);
        int compareIndex1 = ColumnLetterToIndex(compareColumn1);
        int matchIndex2 = ColumnLetterToIndex(matchColumn2);
        int compareIndex2 = ColumnLetterToIndex(compareColumn2);

        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 2; i <= rows1; i++)
        {
            string key = Convert.ToString(sheet1.Cells[i, matchIndex1].Value)?.Trim() ?? "";
            string value = Convert.ToString(sheet1.Cells[i, compareIndex1].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(key) && !lookup.ContainsKey(key))
                lookup[key] = value;
        }

        dynamic targetSheet = CreateTargetSheet(app, "DiffResult");
        targetSheet.Cells[1, 1] = "Key";
        targetSheet.Cells[1, 2] = "Sheet1 Value";
        targetSheet.Cells[1, 3] = "Sheet2 Value";

        using var progress = new ProgressForm();
        progress.Show();

        int differences = 0;
        for (int i = 2; i <= rows2; i++)
        {
            string key = Convert.ToString(sheet2.Cells[i, matchIndex2].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(key) && lookup.TryGetValue(key, out string? value1))
            {
                string value2 = Convert.ToString(sheet2.Cells[i, compareIndex2].Value)?.Trim() ?? "";
                if (!string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase))
                {
                    differences++;
                    targetSheet.Cells[differences + 1, 1] = key;
                    targetSheet.Cells[differences + 1, 2] = value1;
                    targetSheet.Cells[differences + 1, 3] = value2;
                }
            }

            if (i % 200 == 0)
            {
                progress.SetMessage($"正在比对… {i}/{rows2}");
                progress.SetProgress((int)(i * 100.0 / rows2));
                Application.DoEvents();
            }
        }

        targetSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage($"完成。 {differences} difference(s).");
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show($"已完成并保存！ {differences} difference(s).", "Find Differences");
        Close();
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

    private static dynamic CreateTargetSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
            {
                ws.UsedRange.Clear();
                return ws;
            }
        }

        dynamic newSheet = app.Worksheets.Add(After: app.ActiveSheet);
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
