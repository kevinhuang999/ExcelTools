namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 分组统计：按分类字段拆分，生成"分类×月份"交叉汇总表。
/// </summary>
public class FrmGroupStat : Form
{
    private readonly TextBox _txtGroupCol;
    private readonly TextBox _txtCountCol;
    private readonly TextBox _txtMonthCol;

    public FrmGroupStat()
    {
        this.Text = "分组统计";
        this.Size = new Size(400, 210);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "分类列（列字母）：", Location = new Point(15, 20), AutoSize = true };
        _txtGroupCol = new TextBox { Location = new Point(120, 18), Width = 80, Text = "A" };

        var lbl2 = new Label { Text = "统计列（列字母）：", Location = new Point(15, 55), AutoSize = true };
        _txtCountCol = new TextBox { Location = new Point(120, 53), Width = 80, Text = "B" };

        var lbl3 = new Label { Text = "月份列（列字母）：", Location = new Point(15, 90), AutoSize = true };
        _txtMonthCol = new TextBox { Location = new Point(120, 88), Width = 80, Text = "C" };

        var lblHint = new Label
        {
            Text = "说明：按分类列分组，统计每组的记录数，\n按月份列拆分为交叉表，最后一列为合计。",
            Location = new Point(15, 120), Size = new Size(360, 35), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始统计", Location = new Point(200, 165), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1); this.Controls.Add(_txtGroupCol);
        this.Controls.Add(lbl2); this.Controls.Add(_txtCountCol);
        this.Controls.Add(lbl3); this.Controls.Add(_txtMonthCol);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string groupCol = _txtGroupCol.Text.Trim().ToUpper();
        string countCol = _txtCountCol.Text.Trim().ToUpper();
        string monthCol = _txtMonthCol.Text.Trim().ToUpper();

        if (!System.Text.RegularExpressions.Regex.IsMatch(groupCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(countCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(monthCol, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        int groupColIdx = ColumnLetterToIndex(groupCol);
        int monthColIdx = ColumnLetterToIndex(monthCol);

        // Collect group values
        var groups = new List<string>();
        var groupMonthCounts = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        for (int r = 2; r <= rows; r++)
        {
            string g = Convert.ToString(ws.Cells[r, groupColIdx].Value)?.Trim() ?? "";
            string m = Convert.ToString(ws.Cells[r, monthColIdx].Value)?.Trim() ?? "";
            if (string.IsNullOrEmpty(g)) continue;

            if (!groups.Contains(g))
            {
                groups.Add(g);
                groupMonthCounts[g] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            }

            if (!string.IsNullOrEmpty(m))
            {
                if (!groupMonthCounts[g].ContainsKey(m))
                    groupMonthCounts[g][m] = 0;
                groupMonthCounts[g][m]++;
            }
        }

        if (groups.Count == 0) { MessageBox.Show("未找到有效分类数据。"); return; }

        // Collect all unique months
        var allMonths = new List<string>();
        foreach (var gd in groupMonthCounts.Values)
            foreach (var m in gd.Keys)
                if (!allMonths.Contains(m)) allMonths.Add(m);
        allMonths.Sort();

        // Create target sheet
        string tgtName = "GroupStat";
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

        // Write header: group, month1, month2, ..., total
        tgtWs.Cells[1, 1] = "分类";
        for (int c = 0; c < allMonths.Count; c++)
            tgtWs.Cells[1, c + 2] = allMonths[c];
        tgtWs.Cells[1, allMonths.Count + 2] = "合计";

        // Write data rows
        for (int g = 0; g < groups.Count; g++)
        {
            tgtWs.Cells[g + 2, 1] = groups[g];
            int rowTotal = 0;
            var gd = groupMonthCounts[groups[g]];
            for (int c = 0; c < allMonths.Count; c++)
            {
                int cnt = gd.ContainsKey(allMonths[c]) ? gd[allMonths[c]] : 0;
                tgtWs.Cells[g + 2, c + 2] = cnt;
                rowTotal += cnt;
            }
            tgtWs.Cells[g + 2, allMonths.Count + 2] = rowTotal;
        }

        // Write column totals
        int totalRow = groups.Count + 2;
        tgtWs.Cells[totalRow, 1] = "合计";
        for (int c = 0; c < allMonths.Count; c++)
        {
            int colTotal = 0;
            for (int g = 0; g < groups.Count; g++)
                if (groupMonthCounts[groups[g]].ContainsKey(allMonths[c]))
                    colTotal += groupMonthCounts[groups[g]][allMonths[c]];
            tgtWs.Cells[totalRow, c + 2] = colTotal;
        }
        tgtWs.Cells[totalRow, allMonths.Count + 2] = "总计";

        tgtWs.Columns.AutoFit();
        MessageBox.Show($"分组统计完成！共 {groups.Count} 个分类，{allMonths.Count} 个月份。");
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