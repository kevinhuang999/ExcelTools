namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Filter: copy matching rows into a new sheet, either by row-by-row matching
/// against typed conditions, or via a SQL query over the sheet.
///
/// Step flow matches the original add-in: type the conditions, press OK or SQL.
/// The key column and header row are pre-filled from the last run.
/// </summary>
public class FrmFilter : Form
{
    private const string KeyColumn = "filter.column";
    private const string KeyHeaderRow = "filter.headerRow";

    private readonly RichTextBox _rtbConditions;
    private readonly TextBox _txtColumn;
    private readonly TextBox _txtHeaderRow;

    public FrmFilter()
    {
        Text = "Filter";
        Size = new Size(420, 300);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(new Label { Text = "Filter Fields:", Location = new Point(15, 15), AutoSize = true });
        _rtbConditions = new RichTextBox { Location = new Point(15, 35), Size = new Size(370, 100) };
        Controls.Add(_rtbConditions);

        Controls.Add(new Label { Text = "Column:", Location = new Point(15, 148), AutoSize = true });
        _txtColumn = new TextBox
        {
            Location = new Point(120, 145),
            Width = 70,
            Text = Settings.Get(KeyColumn, "A")
        };
        Controls.Add(_txtColumn);

        Controls.Add(new Label { Text = "HeaderRowNo:", Location = new Point(15, 178), AutoSize = true });
        _txtHeaderRow = new TextBox
        {
            Location = new Point(120, 175),
            Width = 70,
            Text = Settings.Get(KeyHeaderRow, "1")
        };
        Controls.Add(_txtHeaderRow);

        var btnOK = new Button { Text = "&OK", Location = new Point(15, 215), Size = new Size(85, 28) };
        btnOK.Click += BtnOK_Click;

        var btnSQL = new Button { Text = "&SQL", Location = new Point(110, 215), Size = new Size(85, 28) };
        btnSQL.Click += BtnSQL_Click;

        var btnCancel = new Button { Text = "&Cancel", Location = new Point(205, 215), Size = new Size(85, 28) };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnOK);
        Controls.Add(btnSQL);
        Controls.Add(btnCancel);

        AcceptButton = btnOK;
        UiTheme.Apply(this, "Copies matching rows into a new Filtered sheet. Type one\ncondition per line as:  column value");
        CancelButton = btnCancel;
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

        var conditions = ParseConditions(_rtbConditions.Text);
        if (conditions.Count == 0) { MessageBox.Show("Please enter at least one filter condition.", "Warning"); return; }

        int headerRow = ParseHeaderRow();

        Settings.SetMany((KeyColumn, _txtColumn.Text.Trim().ToUpperInvariant()), (KeyHeaderRow, headerRow.ToString()));

        dynamic targetSheet = CreateTargetSheet(app, ws, "Filtered");

        for (int c = 1; c <= cols; c++)
            targetSheet.Cells[1, c] = ws.Cells[headerRow, c]?.Value;

        using var progress = new ProgressForm();
        progress.Show();

        int matched = 0;
        for (int r = headerRow + 1; r <= rows; r++)
        {
            bool allMatch = true;
            foreach (var condition in conditions)
            {
                string cellValue = Convert.ToString(ws.Cells[r, condition.ColumnIndex].Value)?.Trim() ?? "";
                if (!cellValue.Equals(condition.Value, StringComparison.OrdinalIgnoreCase))
                {
                    allMatch = false;
                    break;
                }
            }

            if (allMatch)
            {
                matched++;
                for (int c = 1; c <= cols; c++)
                    targetSheet.Cells[matched + 1, c] = ws.Cells[r, c]?.Value;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"Filtering... {r}/{rows}");
                progress.SetProgress((int)(r * 100.0 / rows));
                Application.DoEvents();
            }
        }

        targetSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage($"Done. {matched} row(s) matched.");
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show($"Finished and Saved! {matched} row(s) matched.", "Filter");
        Close();
    }

    private void BtnSQL_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

        string keyColumn = _txtColumn.Text.Trim().ToUpperInvariant();
        if (!IsColumnLetter(keyColumn))
        {
            MessageBox.Show("Column must be a column letter such as A or AB.", "Warning");
            return;
        }

        Settings.SetMany((KeyColumn, keyColumn), (KeyHeaderRow, ParseHeaderRow().ToString()));

        string sheetName = Convert.ToString(ws.Name)?.Trim() ?? "Sheet1";
        string sql = $"SELECT * FROM [{sheetName}$] WHERE [{keyColumn}] IS NOT NULL";

        dynamic targetSheet = CreateTargetSheet(app, ws, "Filtered");
        try
        {
            MyTools.CopyToNewSheet(app, sql, targetSheet);
            MyTools.FormatToText(targetSheet);
            targetSheet.Columns.AutoFit();
            app.ActiveWorkbook.Save();

            MessageBox.Show("Finished and Saved!", "Filter");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"SQL filter failed: {ex.Message}", "Error");
        }
    }

    private int ParseHeaderRow()
    {
        return int.TryParse(_txtHeaderRow.Text.Trim(), out int value) && value >= 1 ? value : 1;
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

    private List<(int ColumnIndex, string Value)> ParseConditions(string text)
    {
        var list = new List<(int, string)>();
        foreach (string line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string[] parts = line.Split(new[] { ' ', '\t' }, 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) continue;

            int columnIndex = ColumnLetterToIndex(parts[0].Trim().ToUpperInvariant());
            if (columnIndex > 0)
                list.Add((columnIndex, parts[1].Trim()));
        }
        return list;
    }

    private static dynamic CreateTargetSheet(dynamic app, dynamic sourceSheet, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
            {
                ws.UsedRange.Clear();
                return ws;
            }
        }

        dynamic newSheet = app.Worksheets.Add(After: sourceSheet);
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
