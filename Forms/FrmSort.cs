namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 排序 the active sheet by one or two columns. The original data is backed up
/// to an "Original" sheet first.
///
/// Step flow matches the original add-in: open, press the action button. The
/// sort columns are pre-filled from the last run.
/// </summary>
public class FrmSort : Form
{
    private const string KeyCol1 = "sort.column1";
    private const string KeyCol2 = "sort.column2";

    private readonly TextBox _txtSortColumn1;
    private readonly TextBox _txtSortColumn2;

    public FrmSort()
    {
        Text = "按列排序";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "按指定列对当前工作表排序。\n排序前会把原始数据整表备份到 Original 工作表。");

        _txtSortColumn1 = layout.AddField("排序列:", Settings.Get(KeyCol1, "A"));
        _txtSortColumn2 = layout.AddField("次要排序列:", Settings.Get(KeyCol2, ""));

        var buttons = layout.Finish(("排序", true), ("取消", false));
        buttons[0].Click += BtnSort_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }

    private void BtnSort_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string column1 = _txtSortColumn1.Text.Trim().ToUpperInvariant();
        string column2 = _txtSortColumn2.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(column1))
        {
            MessageBox.Show("排序列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }
        if (!string.IsNullOrEmpty(column2) && !IsColumnLetter(column2))
        {
            MessageBox.Show("次要排序列必须是列字母，或留空。", "提示");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("空表，或无法清理特殊符号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("数据不足，无法排序。", "提示"); return; }

        Settings.SetMany((KeyCol1, column1), (KeyCol2, column2));

        MyTools.CopyToNewSheet(app, Convert.ToString(ws.Name)?.Trim() ?? "Sheet1");

        int sortIndex1 = ColumnLetterToIndex(column1);
        int helperColumn = cols + 1;

        for (int r = 2; r <= rows; r++)
        {
            string key1 = Convert.ToString(ws.Cells[r, sortIndex1].Value)?.Trim() ?? "";
            string key2 = string.IsNullOrEmpty(column2)
                ? ""
                : Convert.ToString(ws.Cells[r, ColumnLetterToIndex(column2)].Value)?.Trim() ?? "";

            ws.Cells[r, helperColumn] = string.IsNullOrEmpty(column2) ? key1 : key1 + "\t" + key2;
        }

        try
        {
            dynamic sortRange = ws.Range[ws.Cells[1, 1], ws.Cells[rows, cols]];
            dynamic sortKey = ws.Range[ws.Cells[1, helperColumn], ws.Cells[rows, helperColumn]];
            const int xlAscending = 1;
            const int xlYesHeader = 1;
            const int xlTopToBottom = 1;

            // Range.Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3,
            //            Header, OrderCustom, MatchCase, Orientation, SortMethod)
            sortRange.Sort(
                sortKey,
                xlAscending,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                xlYesHeader,
                Type.Missing,
                false,
                xlTopToBottom,
                Type.Missing);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"排序失败： {ex.Message}", "错误");
            return;
        }

        ws.Columns[helperColumn].Delete();
        ws.Columns.AutoFit();
        app.ActiveWorkbook.Save();

        MessageBox.Show("已完成并保存！", "排序");
        Close();
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
