namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Delete rows in the source sheet whose key value also appears in the base
/// sheet.
///
/// Step flow matches the original add-in: open, press the action button. All
/// four parameters are pre-filled from the last run (and default to the
/// conventional names), so normally nothing has to be typed.
/// </summary>
public class FrmDelete : Form
{
    private const string KeySourceSheet = "delete.sourceSheet";
    private const string KeyBaseSheet = "delete.baseSheet";
    private const string KeySourceColumn = "delete.sourceColumn";
    private const string KeyBaseColumn = "delete.baseColumn";

    private readonly TextBox _txtSourceSheet;
    private readonly TextBox _txtBaseSheet;
    private readonly TextBox _txtSourceColumn;
    private readonly TextBox _txtBaseColumn;

    public FrmDelete()
    {
        Text = Loc.T("删除匹配行", "Delete");
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            Loc.T("删除「源工作表」中，键值与「基准工作表」重复的行。", "Delete rows from the Source Sheet whose key value also exists in the Base Sheet."));

        _txtSourceSheet = layout.AddField(Loc.T("源工作表:", "Source Sheet:"), Settings.Get(KeySourceSheet, "Sheet1"));
        _txtBaseSheet = layout.AddField(Loc.T("基准工作表:", "Base Sheet:"), Settings.Get(KeyBaseSheet, "Base"));
        _txtSourceColumn = layout.AddField(Loc.T("源表匹配列:", "Source Column:"), Settings.Get(KeySourceColumn, "A"));
        _txtBaseColumn = layout.AddField(Loc.T("基准表匹配列:", "Base Column:"), Settings.Get(KeyBaseColumn, "A"));

        var buttons = layout.Finish((Loc.T("删除匹配行", "Delete"), true), (Loc.T("取消", "Cancel"), false));
        buttons[0].Click += BtnDelete_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }


    private void BtnDelete_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

        string sourceSheetName = _txtSourceSheet.Text.Trim();
        string baseSheetName = _txtBaseSheet.Text.Trim();
        string sourceColumn = _txtSourceColumn.Text.Trim().ToUpperInvariant();
        string baseColumn = _txtBaseColumn.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sourceSheetName) || string.IsNullOrEmpty(baseSheetName))
        {
            MessageBox.Show(Loc.T("请输入两个工作表名称。", "Please enter both sheet names."), Loc.T("提示", "Warning"));
            return;
        }
        if (!IsColumnLetter(sourceColumn) || !IsColumnLetter(baseColumn))
        {
            MessageBox.Show(Loc.T("列必须是列字母，例如 A 或 AB。", "Columns must be column letters such as A or AB."), Loc.T("提示", "Warning"));
            return;
        }

        dynamic? sourceSheet = FindSheet(app, sourceSheetName);
        dynamic? baseSheet = FindSheet(app, baseSheetName);
        if (sourceSheet == null) { MessageBox.Show(Loc.T($"找不到工作表：{sourceSheetName}", "Sheet not found: {sourceSheetName}"), Loc.T("提示", "Warning")); return; }
        if (baseSheet == null) { MessageBox.Show(Loc.T($"找不到工作表：{baseSheetName}", "Sheet not found: {baseSheetName}"), Loc.T("提示", "Warning")); return; }

        Settings.SetMany(
            (KeySourceSheet, sourceSheetName),
            (KeyBaseSheet, baseSheetName),
            (KeySourceColumn, sourceColumn),
            (KeyBaseColumn, baseColumn));

        MyTools.ClearSpecialSymbols(sourceSheet, out int sourceRows, out int _);
        MyTools.ClearSpecialSymbols(baseSheet, out int baseRows, out int _);
        if (sourceRows < 2 || baseRows < 2) { MessageBox.Show(Loc.T("数据不足。", "Not enough data."), Loc.T("提示", "Warning")); return; }

        int sourceColumnIndex = ColumnLetterToIndex(sourceColumn);
        int baseColumnIndex = ColumnLetterToIndex(baseColumn);

        var baseValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int r = 2; r <= baseRows; r++)
        {
            string value = Convert.ToString(baseSheet.Cells[r, baseColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value)) baseValues.Add(value);
        }
        if (baseValues.Count == 0) { MessageBox.Show(Loc.T("基准工作表没有可用数据。", "The base sheet has no usable values."), Loc.T("提示", "Warning")); return; }

        using var progress = new ProgressForm();
        progress.Show();

        int deleted = 0;
        for (int r = sourceRows; r >= 2; r--)
        {
            string value = Convert.ToString(sourceSheet.Cells[r, sourceColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value) && baseValues.Contains(value))
            {
                sourceSheet.Rows[r].Delete();
                deleted++;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage(Loc.T($"正在删除… {r}/{sourceRows}", "Deleting... {r}/{sourceRows}"));
                progress.SetProgress((int)((sourceRows - r) * 100.0 / sourceRows));
                Application.DoEvents();
            }
        }

        sourceSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage(Loc.T($"完成。 {deleted} row(s) deleted.", "Done. {deleted} row(s) deleted."));
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show(Loc.T($"已完成并保存！ {deleted} row(s) deleted.", "Finished and saved! {deleted} row(s) deleted."), Loc.T("删除匹配行", "Delete"));
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
