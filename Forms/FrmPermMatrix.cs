namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Permission matrix: read role / group / function rows from the active sheet
/// and emit a role-by-function matrix marked with a star.
///
/// Step flow matches the original add-in: open, press OK. All columns are
/// pre-filled from the last run.
/// </summary>
public class FrmPermMatrix : Form
{
    private const string KeyRoleColumn = "perm.roleColumn";
    private const string KeyGroupColumn = "perm.groupColumn";
    private const string KeyFunctionColumn = "perm.functionColumn";

    private readonly TextBox _txtRoleColumn;
    private readonly TextBox _txtGroupColumn;
    private readonly TextBox _txtFunctionColumn;

    public FrmPermMatrix()
    {
        Text = "权限矩阵";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "读取当前工作表的角色/分组/功能，生成「角色 × 功能」权限矩阵，\n有权限的格子标记 *。结果写入新的 PermMatrix 工作表。");

        _txtRoleColumn = layout.AddField("角色列:", Settings.Get(KeyRoleColumn, "A"));
        _txtGroupColumn = layout.AddField("分组列:", Settings.Get(KeyGroupColumn, "B"));
        _txtFunctionColumn = layout.AddField("功能列:", Settings.Get(KeyFunctionColumn, "C"));

        var buttons = layout.Finish(("确定", true), ("取消", false));
        buttons[0].Click += BtnOK_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }


    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string roleColumn = _txtRoleColumn.Text.Trim().ToUpperInvariant();
        string groupColumn = _txtGroupColumn.Text.Trim().ToUpperInvariant();
        string functionColumn = _txtFunctionColumn.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(roleColumn) || !IsColumnLetter(groupColumn) || !IsColumnLetter(functionColumn))
        {
            MessageBox.Show("列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int _))
        {
            MessageBox.Show("空表，或无法清理特殊符号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("数据不足。", "提示"); return; }

        int roleIndex = ColumnLetterToIndex(roleColumn);
        int groupIndex = ColumnLetterToIndex(groupColumn);
        int functionIndex = ColumnLetterToIndex(functionColumn);

        Settings.SetMany(
            (KeyRoleColumn, roleColumn),
            (KeyGroupColumn, groupColumn),
            (KeyFunctionColumn, functionColumn));

        var roles = new List<string>();
        var functions = new List<string>();
        var permissions = new HashSet<string>();

        for (int i = 2; i <= rows; i++)
        {
            string role = Normalize(GetCellValue(ws, i, roleIndex));
            string group = Normalize(GetCellValue(ws, i, groupIndex));
            string function = Normalize(GetCellValue(ws, i, functionIndex));

            if (string.IsNullOrEmpty(role) || string.IsNullOrEmpty(function)) continue;

            if (!roles.Contains(role)) roles.Add(role);

            string functionKey = group + "|" + function;
            if (!functions.Contains(functionKey)) functions.Add(functionKey);

            permissions.Add($"{role}|{group}|{function}");
        }

        if (roles.Count == 0 || functions.Count == 0)
        {
            MessageBox.Show("未找到有效数据。", "提示");
            return;
        }

        dynamic targetSheet = CreateTargetSheet(app, "PermMatrix");

        targetSheet.Cells[1, 1] = "序号";
        targetSheet.Cells[1, 2] = "Group";
        targetSheet.Cells[1, 3] = "Function";
        for (int c = 0; c < roles.Count; c++)
            targetSheet.Cells[1, c + 4] = roles[c];

        for (int r = 0; r < functions.Count; r++)
        {
            int rowIndex = r + 2;
            string[] parts = functions[r].Split(new[] { '|' }, 2);
            string group = parts.Length > 0 ? parts[0] : "";
            string function = parts.Length > 1 ? parts[1] : "";

            targetSheet.Cells[rowIndex, 1] = r + 1;
            targetSheet.Cells[rowIndex, 2] = group;
            targetSheet.Cells[rowIndex, 3] = function;

            for (int c = 0; c < roles.Count; c++)
            {
                if (permissions.Contains($"{roles[c]}|{group}|{function}"))
                    targetSheet.Cells[rowIndex, c + 4] = "*";
            }
        }

        FormatMatrix(targetSheet, app, roles.Count);
        targetSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();

        MessageBox.Show($"已完成并保存！ {roles.Count} role(s), {functions.Count} function(s).", "RoleMenus");
        Close();
    }

    private string GetCellValue(dynamic ws, int row, int column)
    {
        try { return Convert.ToString(ws.Cells[row, column].Value)?.Trim() ?? ""; }
        catch { return ""; }
    }

    private static string Normalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        return System.Text.RegularExpressions.Regex.Replace(value.Trim(), @"\s+", " ");
    }

    private void FormatMatrix(dynamic ws, dynamic app, int roleCount)
    {
        try
        {
            int lastColumn = 3 + roleCount;
            dynamic used = ws.UsedRange;
            int startRow = (int)used.Row;
            int startColumn = (int)used.Column;
            int endRow = startRow + (int)used.Rows.Count - 1;
            int endColumn = startColumn + (int)used.Columns.Count - 1;

            dynamic headerRange = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[startRow, endColumn]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = -4108; // xlCenter
            headerRange.VerticalAlignment = -4108;

            dynamic fullRange = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[endRow, endColumn]];
            fullRange.VerticalAlignment = -4108;
            fullRange.HorizontalAlignment = -4131; // xlLeft

            ws.Columns[1].ColumnWidth = 6;
            ws.Columns[2].ColumnWidth = 20;
            ws.Columns[3].ColumnWidth = 30;
            for (int c = 4; c <= lastColumn; c++)
                ws.Columns[c].ColumnWidth = 15;
        }
        catch
        {
            // Formatting is cosmetic only.
        }
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

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
