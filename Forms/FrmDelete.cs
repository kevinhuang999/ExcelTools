namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 删除匹配行：从源表中删除与基准表匹配的行。所有参数用户输入。
/// </summary>
public class FrmDelete : Form
{
    private readonly TextBox _txtSourceSheet;
    private readonly TextBox _txtBaseSheet;
    private readonly TextBox _txtSourceCol;
    private readonly TextBox _txtBaseCol;

    public FrmDelete()
    {
        this.Text = "删除匹配行";
        this.Size = new Size(400, 200);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "源表名：", Location = new Point(15, 20), AutoSize = true };
        _txtSourceSheet = new TextBox { Location = new Point(120, 18), Width = 180 };

        var lbl2 = new Label { Text = "基准表名：", Location = new Point(15, 50), AutoSize = true };
        _txtBaseSheet = new TextBox { Location = new Point(120, 48), Width = 180 };

        var lbl3 = new Label { Text = "源表匹配列（字母）：", Location = new Point(15, 80), AutoSize = true };
        _txtSourceCol = new TextBox { Location = new Point(120, 78), Width = 80 };

        var lbl4 = new Label { Text = "基准表匹配列（字母）：", Location = new Point(15, 110), AutoSize = true };
        _txtBaseCol = new TextBox { Location = new Point(120, 108), Width = 80 };

        var btnOK = new Button { Text = "开始删除", Location = new Point(180, 150), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1); this.Controls.Add(_txtSourceSheet);
        this.Controls.Add(lbl2); this.Controls.Add(_txtBaseSheet);
        this.Controls.Add(lbl3); this.Controls.Add(_txtSourceCol);
        this.Controls.Add(lbl4); this.Controls.Add(_txtBaseCol);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string srcSheet = _txtSourceSheet.Text.Trim();
        string baseSheet = _txtBaseSheet.Text.Trim();
        string srcCol = _txtSourceCol.Text.Trim().ToUpper();
        string baseCol = _txtBaseCol.Text.Trim().ToUpper();

        if (string.IsNullOrEmpty(srcSheet) || string.IsNullOrEmpty(baseSheet))
        { MessageBox.Show("请输入表名。"); return; }
        if (!System.Text.RegularExpressions.Regex.IsMatch(srcCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(baseCol, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic srcWs = FindSheet(app, srcSheet);
        dynamic baseWs = FindSheet(app, baseSheet);
        if (srcWs == null) { MessageBox.Show($"找不到源表：{srcSheet}"); return; }
        if (baseWs == null) { MessageBox.Show($"找不到基准表：{baseSheet}"); return; }

        MyTools.ClearSpecialSymbols(srcWs, out int srcRows, out int _);
        MyTools.ClearSpecialSymbols(baseWs, out int baseRows, out int _);
        if (srcRows < 2 || baseRows < 2) { MessageBox.Show("数据不足。"); return; }

        int srcColIdx = ColumnLetterToIndex(srcCol);
        int baseColIdx = ColumnLetterToIndex(baseCol);

        // Collect base values into HashSet for fast lookup
        var baseValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int r = 2; r <= baseRows; r++)
        {
            string val = Convert.ToString(baseWs.Cells[r, baseColIdx].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(val))
                baseValues.Add(val);
        }

        if (baseValues.Count == 0) { MessageBox.Show("基准表无有效数据。"); return; }

        using var progress = new ProgressForm();
        progress.Show();

        int deleteCount = 0;
        // Traverse in reverse to safely delete rows
        for (int r = srcRows; r >= 2; r--)
        {
            string val = Convert.ToString(srcWs.Cells[r, srcColIdx].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(val) && baseValues.Contains(val))
            {
                srcWs.Rows[r].Delete();
                deleteCount++;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"处理中… {r}/{srcRows}");
                progress.SetProgress((int)((srcRows - r) * 100.0 / srcRows));
                Application.DoEvents();
            }
        }

        progress.SetMessage($"完成！删除 {deleteCount} 行。");
        srcWs.Columns.AutoFit();
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