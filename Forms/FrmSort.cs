namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 按列排序：用户输入排序列字母，自动备份原始数据到 Original 工作表。
/// </summary>
public class FrmSort : Form
{
    private readonly TextBox _txtSortCol1;
    private readonly TextBox _txtSortCol2;

    public FrmSort()
    {
        this.Text = "按列排序";
        this.Size = new Size(360, 180);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "第一排序列（列字母，如 B）：", Location = new Point(15, 20), AutoSize = true };
        _txtSortCol1 = new TextBox { Location = new Point(200, 18), Width = 80, Text = "B" };

        var lbl2 = new Label { Text = "第二排序列（可选）：", Location = new Point(15, 55), AutoSize = true };
        _txtSortCol2 = new TextBox { Location = new Point(200, 53), Width = 80, Text = "" };

        var lblHint = new Label
        {
            Text = "说明：排序前自动备份原始数据到 Original 表。\n可选第二列做二级排序。",
            Location = new Point(15, 90), Size = new Size(320, 35), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始排序", Location = new Point(180, 135), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1);
        this.Controls.Add(_txtSortCol1);
        this.Controls.Add(lbl2);
        this.Controls.Add(_txtSortCol2);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string col1 = _txtSortCol1.Text.Trim().ToUpper();
        string col2 = _txtSortCol2.Text.Trim().ToUpper();

        if (!System.Text.RegularExpressions.Regex.IsMatch(col1, "^[A-Z]{1,3}$"))
        { MessageBox.Show("排序列格式不正确。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        // Backup original
        MyTools.CopyToNewSheet(app, ws.Name?.ToString().Trim() ?? "Sheet1");

        int colIdx1 = ColumnLetterToIndex(col1);
        if (colIdx1 <= 0) return;

        // Use Excel's built-in sort via helper columns approach
        // Create helper column with sort keys
        int helperCol = cols + 1;
        string helperColLetter = IndexToColumnLetter(helperCol);

        // Build sort key: combine both columns for composite sort
        for (int r = 2; r <= rows; r++)
        {
            string key1 = Convert.ToString(ws.Cells[r, colIdx1].Value)?.Trim() ?? "";
            string key2 = !string.IsNullOrEmpty(col2) && System.Text.RegularExpressions.Regex.IsMatch(col2, "^[A-Z]{1,3}$")
                ? Convert.ToString(ws.Cells[r, ColumnLetterToIndex(col2)].Value)?.Trim() ?? ""
                : "";
            ws.Cells[r, helperCol] = key1 + "\t" + key2;
        }

        // Sort using range.Sort
        var sortRange = ws.Range[ws.Cells[1, 1], ws.Cells[rows, cols]];
        try
        {
            sortRange.Sort(
                ws.Range[ws.Cells[1, helperCol], ws.Cells[rows, helperCol]],
                1, 1, 1, 1, 1, 1, false, false, false);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"排序失败：{ex.Message}");
            return;
        }

        // Delete helper column
        ws.Columns[helperCol].Delete();

        ws.Columns.AutoFit();
        MessageBox.Show("排序完成！原始数据已备份到 Original 工作表。");
    }

    private int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpper())
        {
            if (!char.IsLetter(c)) return 0;
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private string IndexToColumnLetter(int index)
    {
        string result = "";
        while (index > 0)
        {
            int remainder = (index - 1) % 26;
            result = (char)('A' + remainder) + result;
            index = (index - remainder) / 26;
        }
        return result;
    }
}