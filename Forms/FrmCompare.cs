namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 两表比对：比较两个 Excel 文件的第一个工作表，逐格输出差异。
/// </summary>
public class FrmCompare : Form
{
    private readonly TextBox _txtFile1;
    private readonly TextBox _txtFile2;

    public FrmCompare()
    {
        this.Text = "两表比对";
        this.Size = new Size(460, 170);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "文件 1：", Location = new Point(15, 20), AutoSize = true };
        _txtFile1 = new TextBox { Location = new Point(70, 18), Width = 300 };

        var btnFile1 = new Button { Text = "浏览…", Location = new Point(380, 16), Width = 60 };
        btnFile1.Click += (s, e) => SelectFile(_txtFile1);

        var lbl2 = new Label { Text = "文件 2：", Location = new Point(15, 50), AutoSize = true };
        _txtFile2 = new TextBox { Location = new Point(70, 48), Width = 300 };

        var btnFile2 = new Button { Text = "浏览…", Location = new Point(380, 46), Width = 60 };
        btnFile2.Click += (s, e) => SelectFile(_txtFile2);

        var lblHint = new Label
        {
            Text = "说明：比对两个文件的第一个工作表。\n差异内容写入当前工作簿的 Compare 表。",
            Location = new Point(15, 85), Size = new Size(420, 30), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始比对", Location = new Point(200, 125), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1); this.Controls.Add(_txtFile1); this.Controls.Add(btnFile1);
        this.Controls.Add(lbl2); this.Controls.Add(_txtFile2); this.Controls.Add(btnFile2);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void SelectFile(TextBox target)
    {
        using var dlg = new OpenFileDialog { Filter = "Excel 文件|*.xls;*.xlsx;*.xlsm" };
        if (dlg.ShowDialog() == DialogResult.OK)
            target.Text = dlg.FileName;
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string file1 = _txtFile1.Text.Trim();
        string file2 = _txtFile2.Text.Trim();
        if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
        { MessageBox.Show("请选择两个文件。"); return; }
        if (!File.Exists(file1)) { MessageBox.Show($"文件不存在：{file1}"); return; }
        if (!File.Exists(file2)) { MessageBox.Show($"文件不存在：{file2}"); return; }

        try
        {
            // Open both workbooks
            dynamic wb1 = app.Workbooks.Open(file1, 0, true);
            dynamic wb2 = app.Workbooks.Open(file2, 0, true);

            using var progress = new ProgressForm();
            progress.Show();

            // Get first worksheets
            dynamic ws1 = wb1.Worksheets[1];
            dynamic ws2 = wb2.Worksheets[1];

            int rows1 = ws1.UsedRange.Rows.Count;
            int cols1 = ws1.UsedRange.Columns.Count;
            int rows2 = ws2.UsedRange.Rows.Count;
            int cols2 = ws2.UsedRange.Columns.Count;

            int maxRows = Math.Max(rows1, rows2);
            int maxCols = Math.Max(cols1, cols2);

            // Prepare target sheet
            string tgtName = "Compare";
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

            // Write headers
            tgtWs.Cells[1, 1] = "单元格";
            tgtWs.Cells[1, 2] = "文件1值";
            tgtWs.Cells[1, 3] = "文件2值";
            tgtWs.Cells[1, 4] = "差异类型";

            int diffCount = 0;
            for (int r = 1; r <= maxRows; r++)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    string val1 = Convert.ToString(ws1.Cells[r, c].Value)?.Trim() ?? "";
                    string val2 = Convert.ToString(ws2.Cells[r, c].Value)?.Trim() ?? "";

                    if (!string.Equals(val1, val2, StringComparison.OrdinalIgnoreCase))
                    {
                        diffCount++;
                        tgtWs.Cells[diffCount + 1, 1] = $"R{r}C{c}";
                        tgtWs.Cells[diffCount + 1, 2] = val1;
                        tgtWs.Cells[diffCount + 1, 3] = val2;
                        tgtWs.Cells[diffCount + 1, 4] =
                            string.IsNullOrEmpty(val1) ? "仅文件2" :
                            string.IsNullOrEmpty(val2) ? "仅文件1" : "值不同";
                    }
                }

                if (r % 200 == 0)
                {
                    progress.SetMessage($"比对中… {r}/{maxRows}");
                    progress.SetProgress((int)(r * 100.0 / maxRows));
                    Application.DoEvents();
                }
            }

            tgtWs.Columns.AutoFit();

            // Close external workbooks
            wb1.Close(false);
            wb2.Close(false);

            progress.SetMessage($"比对完成！共 {diffCount} 处差异。");
            System.Threading.Thread.Sleep(1000);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"比对出错：{ex.Message}");
        }
    }
}