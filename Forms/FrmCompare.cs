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
        Text = "Compare Excel";
        Size = new Size(440, 175);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(new Label { Text = "File1", Location = new Point(15, 22), AutoSize = true });
        _txtFile1 = new TextBox { Location = new Point(70, 19), Width = 260, Text = Settings.Get(KeyFile1, "") };
        Controls.Add(_txtFile1);

        var btnSelect1 = new Button { Text = "Select File", Location = new Point(340, 17), Size = new Size(80, 26) };
        btnSelect1.Click += (_, _) => SelectFile(_txtFile1);
        Controls.Add(btnSelect1);

        Controls.Add(new Label { Text = "File2", Location = new Point(15, 52), AutoSize = true });
        _txtFile2 = new TextBox { Location = new Point(70, 49), Width = 260, Text = Settings.Get(KeyFile2, "") };
        Controls.Add(_txtFile2);

        var btnSelect2 = new Button { Text = "Select File", Location = new Point(340, 47), Size = new Size(80, 26) };
        btnSelect2.Click += (_, _) => SelectFile(_txtFile2);
        Controls.Add(btnSelect2);

        var btnCompare = new Button { Text = "Compare", Location = new Point(240, 100), Size = new Size(85, 28) };
        btnCompare.Click += BtnCompare_Click;

        var btnCancel = new Button { Text = "Cancel", Location = new Point(335, 100), Size = new Size(85, 28) };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnCompare);
        Controls.Add(btnCancel);

        AcceptButton = btnCompare;
        UiTheme.Apply(this, "Compares the first sheet of two workbooks cell by cell and\nlists every difference on a new Compare sheet.");
        CancelButton = btnCancel;
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
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string file1 = _txtFile1.Text.Trim();
        string file2 = _txtFile2.Text.Trim();

        if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
        {
            MessageBox.Show("Please select both files.", "Warning");
            return;
        }
        if (!File.Exists(file1)) { MessageBox.Show($"File not found: {file1}", "Warning"); return; }
        if (!File.Exists(file2)) { MessageBox.Show($"File not found: {file2}", "Warning"); return; }

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
                    progress.SetMessage($"Comparing... {r}/{maxRows}");
                    progress.SetProgress((int)(r * 100.0 / maxRows));
                    Application.DoEvents();
                }
            }

            targetSheet.Columns.AutoFit();
            app.ActiveWorkbook.Save();
            progress.SetProgress(100);
            progress.SetMessage($"Done. {differences} difference(s).");
            Application.DoEvents();
            System.Threading.Thread.Sleep(600);

            MessageBox.Show($"Finished and Saved! {differences} difference(s).", "Compare");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Compare failed: {ex.Message}", "Error");
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
