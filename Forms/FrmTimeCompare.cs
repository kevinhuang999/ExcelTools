namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 两表差异比对：按匹配列联表，找出比对列不一致的行。
/// 通用化自 FrmMatch（原为温度记录比对）。
/// </summary>
public class FrmTimeCompare : Form
{
    private readonly TextBox _txtMatchCol1;
    private readonly TextBox _txtCompareCol1;
    private readonly TextBox _txtMatchCol2;
    private readonly TextBox _txtCompareCol2;
    private readonly TextBox _txtSheet1;
    private readonly TextBox _txtSheet2;

    public FrmTimeCompare()
    {
        this.Text = "两表差异比对";
        this.Size = new Size(460, 260);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblSheet1 = new Label { Text = "表1名称：", Location = new Point(15, 15), AutoSize = true };
        _txtSheet1 = new TextBox { Location = new Point(90, 13), Width = 120 };

        var lblSheet2 = new Label { Text = "表2名称：", Location = new Point(15, 42), AutoSize = true };
        _txtSheet2 = new TextBox { Location = new Point(90, 40), Width = 120 };

        var lblM1 = new Label { Text = "表1匹配列（字母）：", Location = new Point(15, 69), AutoSize = true };
        _txtMatchCol1 = new TextBox { Location = new Point(120, 67), Width = 60, Text = "A" };

        var lblC1 = new Label { Text = "表1比对列（字母）：", Location = new Point(15, 96), AutoSize = true };
        _txtCompareCol1 = new TextBox { Location = new Point(120, 94), Width = 60, Text = "B" };

        var lblM2 = new Label { Text = "表2匹配列（字母）：", Location = new Point(250, 69), AutoSize = true };
        _txtMatchCol2 = new TextBox { Location = new Point(370, 67), Width = 60, Text = "A" };

        var lblC2 = new Label { Text = "表2比对列（字母）：", Location = new Point(250, 96), AutoSize = true };
        _txtCompareCol2 = new TextBox { Location = new Point(370, 94), Width = 60, Text = "B" };

        var lblHint = new Label
        {
            Text = "说明：按匹配列联两表，找比对列值不一致的行。\n输出到当前工作簿的差异表。",
            Location = new Point(15, 130), Size = new Size(420, 30), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始比对", Location = new Point(200, 170), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lblSheet1); this.Controls.Add(_txtSheet1);
        this.Controls.Add(lblSheet2); this.Controls.Add(_txtSheet2);
        this.Controls.Add(lblM1); this.Controls.Add(_txtMatchCol1);
        this.Controls.Add(lblC1); this.Controls.Add(_txtCompareCol1);
        this.Controls.Add(lblM2); this.Controls.Add(_txtMatchCol2);
        this.Controls.Add(lblC2); this.Controls.Add(_txtCompareCol2);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string sheet1Name = _txtSheet1.Text.Trim();
        string sheet2Name = _txtSheet2.Text.Trim();
        string matchCol1 = _txtMatchCol1.Text.Trim().ToUpper();
        string compareCol1 = _txtCompareCol1.Text.Trim().ToUpper();
        string matchCol2 = _txtMatchCol2.Text.Trim().ToUpper();
        string compareCol2 = _txtCompareCol2.Text.Trim().ToUpper();

        if (string.IsNullOrEmpty(sheet1Name) || string.IsNullOrEmpty(sheet2Name))
        { MessageBox.Show("请输入两个表名。"); return; }
        if (!System.Text.RegularExpressions.Regex.IsMatch(matchCol1, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(compareCol1, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(matchCol2, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(compareCol2, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic ws1 = FindSheet(app, sheet1Name);
        dynamic ws2 = FindSheet(app, sheet2Name);
        if (ws1 == null) { MessageBox.Show($"找不到表1：{sheet1Name}"); return; }
        if (ws2 == null) { MessageBox.Show($"找不到表2：{sheet2Name}"); return; }

        MyTools.ClearSpecialSymbols(ws1, out int rows1, out int _);
        MyTools.ClearSpecialSymbols(ws2, out int rows2, out int _);
        if (rows1 < 2 || rows2 < 2) { MessageBox.Show("数据不足。"); return; }

        int mIdx1 = ColumnLetterToIndex(matchCol1);
        int cIdx1 = ColumnLetterToIndex(compareCol1);
        int mIdx2 = ColumnLetterToIndex(matchCol2);
        int cIdx2 = ColumnLetterToIndex(compareCol2);
        if (mIdx1 <= 0 || cIdx1 <= 0 || mIdx2 <= 0 || cIdx2 <= 0) return;

        // Build dict from sheet1: matchCol -> compareCol value
        var dict1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 2; i <= rows1; i++)
        {
            string key = Convert.ToString(ws1.Cells[i, mIdx1].Value)?.Trim() ?? "";
            string val = Convert.ToString(ws1.Cells[i, cIdx1].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(key) && !dict1.ContainsKey(key))
                dict1[key] = val;
        }

        // Compare with sheet2
        using var progress = new ProgressForm();
        progress.Show();

        string tgtName = "DiffResult";
        dynamic tgtWs = null;
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(ws.Name?.ToString().Trim(), tgtName, StringComparison.OrdinalIgnoreCase))
            { tgtWs = ws; ws.UsedRange.Clear(); break; }
        }
        if (tgtWs == null)
        {
            tgtWs = app.Worksheets.Add(After: app.ActiveSheet);
            tgtWs.Name = tgtName;
        }

        tgtWs.Cells[1, 1] = "匹配列";
        tgtWs.Cells[1, 2] = "表1比对值";
        tgtWs.Cells[1, 3] = "表2比对值";

        int diffCount = 0;
        for (int i = 2; i <= rows2; i++)
        {
            string key = Convert.ToString(ws2.Cells[i, mIdx2].Value)?.Trim() ?? "";
            string val2 = Convert.ToString(ws2.Cells[i, cIdx2].Value)?.Trim() ?? "";

            if (!string.IsNullOrEmpty(key) && dict1.ContainsKey(key))
            {
                string val1 = dict1[key];
                if (!string.Equals(val1, val2, StringComparison.OrdinalIgnoreCase))
                {
                    diffCount++;
                    tgtWs.Cells[diffCount + 1, 1] = key;
                    tgtWs.Cells[diffCount + 1, 2] = val1;
                    tgtWs.Cells[diffCount + 1, 3] = val2;
                }
            }

            if (i % 200 == 0)
            {
                progress.SetMessage($"比对中… {i}/{rows2}");
                progress.SetProgress((int)(i * 100.0 / rows2));
                Application.DoEvents();
            }
        }

        tgtWs.Columns.AutoFit();
        progress.SetMessage($"完成！共 {diffCount} 处差异。");
        System.Threading.Thread.Sleep(1000);
    }

    private dynamic? FindSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(ws.Name?.ToString().Trim(), name, StringComparison.OrdinalIgnoreCase))
                return ws;
        }
        return null;
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