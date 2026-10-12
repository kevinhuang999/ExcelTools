namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 权限矩阵生成器：从数据表生成角色×功能模块权限矩阵，用★标记。
/// </summary>
public class FrmPermMatrix : Form
{
    private readonly TextBox _txtRoleCol;
    private readonly TextBox _txtGroupCol;
    private readonly TextBox _txtFuncCol;

    public FrmPermMatrix()
    {
        this.Text = "权限矩阵生成";
        this.Size = new Size(440, 210);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblRole = new Label { Text = "角色列（字母）：", Location = new Point(15, 20), AutoSize = true };
        _txtRoleCol = new TextBox { Location = new Point(130, 18), Width = 60, Text = "B" };

        var lblGroup = new Label { Text = "分组列（字母）：", Location = new Point(15, 55), AutoSize = true };
        _txtGroupCol = new TextBox { Location = new Point(130, 53), Width = 60, Text = "D" };

        var lblFunc = new Label { Text = "功能列（字母）：", Location = new Point(15, 90), AutoSize = true };
        _txtFuncCol = new TextBox { Location = new Point(130, 88), Width = 60, Text = "E" };

        var lblHint = new Label
        {
            Text = "说明：从当前活动工作表读取角色/分组/功能数据，\n生成权限矩阵到新工作表，有权限标记★。\n角色列和分组列支持合并单元格。",
            Location = new Point(15, 120), Size = new Size(400, 40), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "生成矩阵", Location = new Point(220, 165), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lblRole); this.Controls.Add(_txtRoleCol);
        this.Controls.Add(lblGroup); this.Controls.Add(_txtGroupCol);
        this.Controls.Add(lblFunc); this.Controls.Add(_txtFuncCol);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string roleCol = _txtRoleCol.Text.Trim().ToUpper();
        string groupCol = _txtGroupCol.Text.Trim().ToUpper();
        string funcCol = _txtFuncCol.Text.Trim().ToUpper();

        if (!System.Text.RegularExpressions.Regex.IsMatch(roleCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(groupCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(funcCol, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int _);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        int roleIdx = ColumnLetterToIndex(roleCol);
        int groupIdx = ColumnLetterToIndex(groupCol);
        int funcIdx = ColumnLetterToIndex(funcCol);
        if (roleIdx <= 0 || groupIdx <= 0 || funcIdx <= 0) return;

        // Collect roles, groups, functions and permissions
        var roleNames = new List<string>();
        var funcKeys = new List<string>(); // "group|func"
        var permissionSet = new HashSet<string>(); // "role|group|func"

        for (int i = 2; i <= rows; i++)
        {
            string roleName = Normalize(GetCellValue(ws, i, roleIdx));
            string group = Normalize(GetCellValue(ws, i, groupIdx));
            string func = Normalize(GetCellValue(ws, i, funcIdx));

            if (string.IsNullOrEmpty(roleName) || string.IsNullOrEmpty(func)) continue;

            if (!roleNames.Contains(roleName))
                roleNames.Add(roleName);

            string menuKey = group + "|" + func;
            if (!funcKeys.Contains(menuKey))
                funcKeys.Add(menuKey);

            string permKey = $"{roleName}|{group}|{func}";
            permissionSet.Add(permKey);
        }

        if (roleNames.Count == 0 || funcKeys.Count == 0)
        {
            MessageBox.Show("未找到有效数据。");
            return;
        }

        // Create target sheet
        string tgtName = "PermMatrix";
        dynamic tgtWs = null;
        foreach (dynamic w in app.Worksheets)
        {
            if (string.Equals(w.Name?.ToString().Trim(), tgtName, StringComparison.OrdinalIgnoreCase))
            { tgtWs = w; w.UsedRange.Clear(); break; }
        }
        if (tgtWs == null)
        {
            tgtWs = app.Worksheets.Add(After: ws);
            tgtWs.Name = tgtName;
        }

        // Write headers
        tgtWs.Cells[1, 1] = "序号";
        tgtWs.Cells[1, 2] = "分组";
        tgtWs.Cells[1, 3] = "功能模块";
        for (int c = 0; c < roleNames.Count; c++)
        {
            tgtWs.Cells[1, c + 4] = roleNames[c];
        }

        // Write data rows
        for (int r = 0; r < funcKeys.Count; r++)
        {
            int rowIdx = r + 2;
            var parts = funcKeys[r].Split(new[] { '|' }, 2);
            string group = parts.Length > 0 ? parts[0] : "";
            string func = parts.Length > 1 ? parts[1] : "";

            tgtWs.Cells[rowIdx, 1] = r + 1;
            tgtWs.Cells[rowIdx, 2] = group;
            tgtWs.Cells[rowIdx, 3] = func;

            // Mark permissions with ★
            for (int c = 0; c < roleNames.Count; c++)
            {
                string key = $"{roleNames[c]}|{group}|{func}";
                if (permissionSet.Contains(key))
                    tgtWs.Cells[rowIdx, c + 4] = "★";
            }
        }

        // Format: header bold + gray, borders
        FormatMatrix(tgtWs, app, rows, roleNames.Count);

        tgtWs.Columns.AutoFit();
        MessageBox.Show($"权限矩阵生成完成！{roleNames.Count} 个角色，{funcKeys.Count} 个功能模块。");
    }

    private string GetCellValue(dynamic ws, int row, int col)
    {
        try
        {
            return Convert.ToString(ws.Cells[row, col].Value)?.Trim() ?? "";
        }
        catch { return ""; }
    }

    private string Normalize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        return System.Text.RegularExpressions.Regex.Replace(s.Trim(), @"\s+", " ");
    }

    private void FormatMatrix(dynamic ws, dynamic app, int dataRows, int roleCount)
    {
        try
        {
            int lastCol = 3 + roleCount;
            var used = ws.UsedRange;
            int startRow = (int)used.Row;
            int startCol = (int)used.Column;
            int endRow = startRow + (int)used.Rows.Count - 1;
            int endCol = startCol + (int)used.Columns.Count - 1;

            // Header style
            var headerRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[startRow, endCol]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = -4108; // xlCenter
            headerRange.VerticalAlignment = -4108;

            // Full range
            var fullRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[endRow, endCol]];
            fullRange.VerticalAlignment = -4108;
            fullRange.HorizontalAlignment = -4131; // xlLeft

            // Column widths
            ws.Columns[1].ColumnWidth = 6;
            ws.Columns[2].ColumnWidth = 20;
            ws.Columns[3].ColumnWidth = 30;
            for (int c = 4; c <= lastCol; c++)
                ws.Columns[c].ColumnWidth = 15;

            // Merge contiguous same-group cells
            string lastGroup = null;
            int mergeStart = -1;
            for (int r = startRow + 1; r <= endRow; r++)
            {
                string grp = GetCellValue(ws, r, 2);
                if (string.IsNullOrEmpty(grp) && lastGroup != null)
                    grp = lastGroup;

                if (!string.IsNullOrEmpty(grp) && grp != lastGroup)
                {
                    if (mergeStart > 0 && mergeStart < r - 1)
                    {
                        var mr = ws.Range[ws.Cells[mergeStart, 2], ws.Cells[r - 1, 2]];
                        try
                        {
                            bool oldAlert = (bool)app.DisplayAlerts;
                            try { app.DisplayAlerts = false; mr.Merge(); } catch { }
                            finally { app.DisplayAlerts = oldAlert; }
                        }
                        catch { }
                    }
                    mergeStart = r;
                    lastGroup = grp;
                }
            }
            // Finalize last group
            if (mergeStart > 0 && mergeStart < endRow)
            {
                var mr = ws.Range[ws.Cells[mergeStart, 2], ws.Cells[endRow, 2]];
                try { mr.Merge(); mr.VerticalAlignment = -4108; } catch { }
            }

            // Freeze panes
            app.ActiveWindow.SplitRow = 1;
            app.ActiveWindow.SplitColumn = 3;
            app.ActiveWindow.FreezePanes = true;
        }
        catch { }
    }

    private int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col)
        {
            if (!char.IsLetter(c)) return 0;
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }
}