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
        Text = Loc.T("拆分工作表", "Seperate");
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            Loc.T("按指定列的值，把当前工作表拆成多个工作表：\n每一个不同的值生成一张新表，表名即该值。", "Split the active sheet into several sheets:\none sheet per distinct value, named after that value."));

        _txtColumnHeader = layout.AddField(Loc.T("列标题:", "Column Header:"), Settings.Get(SettingsKey, "A"));

        var buttons = layout.Finish((Loc.T("拆分", "Seperate"), true), (Loc.T("取消", "Cancel"), false));
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
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

        string column = _txtColumnHeader.Text.Trim().ToUpperInvariant();
        if (string.IsNullOrEmpty(column))
        {
            MessageBox.Show(Loc.T("请输入列标题！", "Please enter the Column Header!"), Loc.T("提示", "Warning"));
            return;
        }
        if (!System.Text.RegularExpressions.Regex.IsMatch(column, "^[A-Z]{1,3}$"))
        {
            MessageBox.Show(Loc.T("列标题必须是列字母，例如 A 或 AB。", "The Column Header must be a column letter such as A or AB."), Loc.T("提示", "Warning"));
            return;
        }

        dynamic mainSheet = app.ActiveSheet;

        if (string.IsNullOrEmpty(Convert.ToString(mainSheet.Cells[1, column].Value)))
        {
            MessageBox.Show(Loc.T("表头不能为空！", "The header should not be empty!"), Loc.T("提示", "Warning"));
            return;
        }

        string mainSheetName = Convert.ToString(mainSheet.Name)?.Trim() ?? "Sheet1";

        if (!MyTools.ClearSpecialSymbols(mainSheet, out int rowsCount, out int colsCount))
        {
            MessageBox.Show(Loc.T("空表，或无法清理特殊符号！", "Empty table, or unable to clean up special symbols!"), Loc.T("提示", "Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        app.ActiveWorkbook.Save();

        List<string>? listByField = MyTools.GetByField(app, column);
        if (listByField == null)
        {
            MessageBox.Show(Loc.T("未找到有效数据！", "No usable data found!"), Loc.T("提示", "Warning"));
            return;
        }
        if (!MyTools.IsSheetNameValid(listByField))
        {
            MessageBox.Show(Loc.T("存在不符合工作表命名规则的字符！", "Contains strings that do not conform to the worksheet naming rules!"), Loc.T("提示", "Warning"));
            return;
        }

        Settings.Set(SettingsKey, column);

        using var progress = new ProgressForm();
        progress.Show();

        int index = 0;
        foreach (string item in listByField)
        {
            index++;
            progress.SetMessage(Loc.T($"正在拆分：{item}  ({index}/{listByField.Count})", "Seperating: {item}  ({index}/{listByField.Count})"));
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
        progress.SetMessage(Loc.T("完成。", "Done."));
        app.ActiveWorkbook.Save();

        Close();

        if (DialogResult.Yes == MessageBox.Show(Loc.T("已完成并保存，现在关闭吗？", "Finished and saved! Close now?"), Loc.T("确认", "Question"), MessageBoxButtons.YesNo))
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
