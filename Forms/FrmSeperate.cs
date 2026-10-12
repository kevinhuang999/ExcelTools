namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Seperate the active sheet into several sheets, one per distinct value of a
/// chosen column.
///
/// Step flow matches the original add-in: open, (optionally change the column),
/// press the action button. The column is pre-filled from the last run.
/// </summary>
public class FrmSeperate : Form
{
    private const string SettingsKey = "seperate.column";

    private readonly TextBox _txtColumnHeader;

    public FrmSeperate()
    {
        Text = "拆分工作表";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "按指定列的值，把当前工作表拆成多个工作表：\n每一个不同的值生成一张新表，表名即该值。");

        _txtColumnHeader = layout.AddField("列标题:", Settings.Get(SettingsKey, "A"));

        var buttons = layout.Finish(("拆分", true), ("取消", false));
        buttons[0].Click += BtnSeperate_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
        _txtColumnHeader.SelectAll();
    }

    private void BtnSeperate_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string column = _txtColumnHeader.Text.Trim().ToUpperInvariant();
        if (string.IsNullOrEmpty(column))
        {
            MessageBox.Show("请输入列标题！", "提示");
            return;
        }
        if (!System.Text.RegularExpressions.Regex.IsMatch(column, "^[A-Z]{1,3}$"))
        {
            MessageBox.Show("列标题必须是列字母，例如 A 或 AB。", "提示");
            return;
        }

        dynamic mainSheet = app.ActiveSheet;

        if (string.IsNullOrEmpty(Convert.ToString(mainSheet.Cells[1, column].Value)))
        {
            MessageBox.Show("表头不能为空！", "提示");
            return;
        }

        string mainSheetName = Convert.ToString(mainSheet.Name)?.Trim() ?? "Sheet1";

        if (!MyTools.ClearSpecialSymbols(mainSheet, out int rowsCount, out int colsCount))
        {
            MessageBox.Show("空表，或无法清理特殊符号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        app.ActiveWorkbook.Save();

        List<string>? listByField = MyTools.GetByField(app, column);
        if (listByField == null)
        {
            MessageBox.Show("未找到有效数据！", "提示");
            return;
        }
        if (!MyTools.IsSheetNameValid(listByField))
        {
            MessageBox.Show("存在不符合工作表命名规则的字符！", "提示");
            return;
        }

        Settings.Set(SettingsKey, column);

        using var progress = new ProgressForm();
        progress.Show();

        int index = 0;
        foreach (string item in listByField)
        {
            index++;
            progress.SetMessage($"正在拆分：{item}  ({index}/{listByField.Count})");
            progress.SetProgress((int)(index * 100.0 / listByField.Count));
            Application.DoEvents();

            dynamic? targetSheet = FindSheet(app, item);
            if (targetSheet == null)
            {
                targetSheet = app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
                targetSheet.Name = item;
            }
            else
            {
                targetSheet.UsedRange.Clear();
            }

            MyTools.FormatToText(targetSheet, rowsCount, colsCount);

            string sql = $"SELECT * FROM [{mainSheetName}$] WHERE [{mainSheetName}$].[{Convert.ToString(app.Cells[1, column].Value)}] = '{item.Replace("'", "''")}'";
            MyTools.CopyToNewSheet(app, sql, targetSheet);
        }

        progress.SetProgress(100);
        progress.SetMessage("完成。");
        app.ActiveWorkbook.Save();

        Close();

        if (DialogResult.Yes == MessageBox.Show("已完成并保存，现在关闭吗？", "确认", MessageBoxButtons.YesNo))
        {
            app.Quit();
        }
    }

    private static dynamic? FindSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                return ws;
        }
        return null;
    }
}
