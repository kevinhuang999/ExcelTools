namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 累计去重计数：对指定列，逐行统计截至该行出现过多少个不重复值。
/// </summary>
public class FrmCountDistinct : Form
{
    private readonly TextBox _txtCountCol;
    private readonly TextBox _txtOutputCol;

    public FrmCountDistinct()
    {
        this.Text = "累计去重计数";
        this.Size = new Size(380, 170);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "统计列（列字母，如 A）：", Location = new Point(15, 20), AutoSize = true };
        _txtCountCol = new TextBox { Location = new Point(160, 18), Width = 80, Text = "A" };

        var lbl2 = new Label { Text = "输出列（列字母，如 Z）：", Location = new Point(15, 55), AutoSize = true };
        _txtOutputCol = new TextBox { Location = new Point(160, 53), Width = 80, Text = "Z" };

        var lblHint = new Label
        {
            Text = "说明：逐行向下，统计截至该行指定列出现\n过多少个不重复值，写入输出列。",
            Location = new Point(15, 90), Size = new Size(340, 35), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始统计", Location = new Point(180, 130), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1); this.Controls.Add(_txtCountCol);
        this.Controls.Add(lbl2); this.Controls.Add(_txtOutputCol);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string countCol = _txtCountCol.Text.Trim().ToUpper();
        string outputCol = _txtOutputCol.Text.Trim().ToUpper();

        if (!System.Text.RegularExpressions.Regex.IsMatch(countCol, "^[A-Z]{1,3}$") ||
            !System.Text.RegularExpressions.Regex.IsMatch(outputCol, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        int countColIdx = ColumnLetterToIndex(countCol);
        int outputColIdx = ColumnLetterToIndex(outputCol);

        if (countColIdx <= 0 || outputColIdx <= 0) return;

        // Ensure output column is beyond data range
        if (outputColIdx <= cols)
        {
            MessageBox.Show("输出列必须在数据范围之外。请调整输出列。");
            return;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ws.Cells[1, outputColIdx] = "累计去重数";

        using var progress = new ProgressForm();
        progress.Show();

        for (int r = 2; r <= rows; r++)
        {
            string val = Convert.ToString(ws.Cells[r, countColIdx].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(val))
                seen.Add(val);

            ws.Cells[r, outputColIdx] = seen.Count;

            if (r % 500 == 0)
            {
                progress.SetMessage($"统计中… {r}/{rows}");
                progress.SetProgress((int)(r * 100.0 / rows));
                Application.DoEvents();
            }
        }

        ws.Columns.AutoFit();
        progress.SetMessage($"完成！最终不重复值：{seen.Count} 个。");
        System.Threading.Thread.Sleep(1000);
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