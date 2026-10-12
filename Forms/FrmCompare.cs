namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Compare the first worksheet of two workbooks cell by cell and write the
/// differences to a "Compare" sheet.
///
/// Step flow matches the original add-in: pick the two files, press Compare.
/// The file paths are pre-filled from the last run.
/// </summary>
public class FrmCompare : Form
{
    private const string KeyFile1 = "compare.file1";
    private const string KeyFile2 = "compare.file2";

    private readonly TextBox _txtFile1;
    private readonly TextBox _txtFile2;

    public FrmCompare()
    {
        Text = "两表比对";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "逐格比对两个文件的第一个工作表，差异写入新的 Compare 工作表。");

        var (file1Box, btnFile1) = layout.AddFieldWithButton("文件1:", Settings.Get(KeyFile1, ""), "选择文件");
        var (file2Box, btnFile2) = layout.AddFieldWithButton("文件2:", Settings.Get(KeyFile2, ""), "选择文件");
        _txtFile1 = file1Box;
        _txtFile2 = file2Box;

        btnFile1.Click += (_, _) => SelectFile(_txtFile1);
        btnFile2.Click += (_, _) => SelectFile(_txtFile2);

        var buttons = layout.Finish(("比对", true), ("取消", false));
        buttons[0].Click += BtnCompare_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }

    private void SelectFile(TextBox target)
    {
        using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };
        if (dialog.ShowDialog() == DialogResult.OK)
            target.Text = dialog.FileName;
    }

    private void BtnCompare_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string file1 = _txtFile1.Text.Trim();
        string file2 = _txtFile2.Text.Trim();

        if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
        {
            MessageBox.Show("请选择两个文件。", "提示");
            return;
        }
        if (!File.Exists(file1)) { MessageBox.Show($"文件不存在：{file1}", "提示"); return; }
        if (!File.Exists(file2)) { MessageBox.Show($"文件不存在：{file2}", "提示"); return; }

        Settings.SetMany((KeyFile1, file1), (KeyFile2, file2));

        dynamic? workbook1 = null;
        dynamic? workbook2 = null;
        try
        {
            workbook1 = app.Workbooks.Open(file1, 0, true);
            workbook2 = app.Workbooks.Open(file2, 0, true);

            dynamic sheet1 = workbook1.Worksheets[1];
            dynamic sheet2 = workbook2.Worksheets[1];

            int maxRows = Math.Max((int)sheet1.UsedRange.Rows.Count, (int)sheet2.UsedRange.Rows.Count);
            int maxCols = Math.Max((int)sheet1.UsedRange.Columns.Count, (int)sheet2.UsedRange.Columns.Count);

            dynamic targetSheet = CreateTargetSheet(app, "Compare");

            targetSheet.Cells[1, 1] = "Cell";
            targetSheet.Cells[1, 2] = "File1";
            targetSheet.Cells[1, 3] = "File2";
            targetSheet.Cells[1, 4] = "Difference";

            using var progress = new ProgressForm();
            progress.Show();

            int differences = 0;
            for (int r = 1; r <= maxRows; r++)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    string value1 = Convert.ToString(sheet1.Cells[r, c].Value)?.Trim() ?? "";
                    string value2 = Convert.ToString(sheet2.Cells[r, c].Value)?.Trim() ?? "";

                    if (!string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase))
                    {
                        differences++;
                        targetSheet.Cells[differences + 1, 1] = $"R{r}C{c}";
                        targetSheet.Cells[differences + 1, 2] = value1;
                        targetSheet.Cells[differences + 1, 3] = value2;
                        targetSheet.Cells[differences + 1, 4] =
                            string.IsNullOrEmpty(value1) ? "Only in File2" :
                            string.IsNullOrEmpty(value2) ? "Only in File1" : "Value differs";
                    }
                }

                if (r % 200 == 0)
                {
                    progress.SetMessage($"正在比对… {r}/{maxRows}");
                    progress.SetProgress((int)(r * 100.0 / maxRows));
                    Application.DoEvents();
                }
            }

            targetSheet.Columns.AutoFit();
            app.ActiveWorkbook.Save();
            progress.SetProgress(100);
            progress.SetMessage($"完成。 {differences} difference(s).");
            Application.DoEvents();
            System.Threading.Thread.Sleep(600);

            MessageBox.Show($"已完成并保存！ {differences} difference(s).", "Compare");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"比对失败：{ex.Message}", "错误");
        }
        finally
        {
            try { workbook1?.Close(false); } catch { }
            try { workbook2?.Close(false); } catch { }
        }
    }

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
}
