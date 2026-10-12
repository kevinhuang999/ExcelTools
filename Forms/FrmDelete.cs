namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Delete rows in the source sheet whose key value also appears in the base
/// sheet.
///
/// Step flow matches the original add-in: open, press the action button. All
/// four parameters are pre-filled from the last run (and default to the
/// conventional names), so normally nothing has to be typed.
/// </summary>
public class FrmDelete : Form
{
    private const string KeySourceSheet = "delete.sourceSheet";
    private const string KeyBaseSheet = "delete.baseSheet";
    private const string KeySourceColumn = "delete.sourceColumn";
    private const string KeyBaseColumn = "delete.baseColumn";

    private readonly TextBox _txtSourceSheet;
    private readonly TextBox _txtBaseSheet;
    private readonly TextBox _txtSourceColumn;
    private readonly TextBox _txtBaseColumn;

    public FrmDelete()
    {
        Text = "Delete";
        Size = new Size(330, 205);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        _txtSourceSheet = AddRow("Source Sheet:", 18, Settings.Get(KeySourceSheet, "Sheet1"));
        _txtBaseSheet = AddRow("Base Sheet:", 48, Settings.Get(KeyBaseSheet, "Base"));
        _txtSourceColumn = AddRow("Source Column:", 78, Settings.Get(KeySourceColumn, "A"));
        _txtBaseColumn = AddRow("Base Column:", 108, Settings.Get(KeyBaseColumn, "A"));

        var btnDelete = new Button
        {
            Text = "Delete Imported Records",
            Location = new Point(15, 148),
            Size = new Size(170, 30)
        };
        btnDelete.Click += BtnDelete_Click;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(195, 148),
            Size = new Size(90, 30)
        };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnDelete);
        Controls.Add(btnCancel);

        AcceptButton = btnDelete;
        UiTheme.Apply(this, "Deletes rows from the source sheet whose key value also\nexists in the base sheet.");
        CancelButton = btnCancel;
    }

    private TextBox AddRow(string label, int top, string initial)
    {
        Controls.Add(new Label { Text = label, Location = new Point(15, top + 3), AutoSize = true });
        var box = new TextBox { Location = new Point(135, top), Width = 150, Text = initial };
        Controls.Add(box);
        return box;
    }

    private void BtnDelete_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string sourceSheetName = _txtSourceSheet.Text.Trim();
        string baseSheetName = _txtBaseSheet.Text.Trim();
        string sourceColumn = _txtSourceColumn.Text.Trim().ToUpperInvariant();
        string baseColumn = _txtBaseColumn.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sourceSheetName) || string.IsNullOrEmpty(baseSheetName))
        {
            MessageBox.Show("Please enter both sheet names.", "Warning");
            return;
        }
        if (!IsColumnLetter(sourceColumn) || !IsColumnLetter(baseColumn))
        {
            MessageBox.Show("Columns must be column letters such as A or AB.", "Warning");
            return;
        }

        dynamic? sourceSheet = FindSheet(app, sourceSheetName);
        dynamic? baseSheet = FindSheet(app, baseSheetName);
        if (sourceSheet == null) { MessageBox.Show($"Sheet not found: {sourceSheetName}", "Warning"); return; }
        if (baseSheet == null) { MessageBox.Show($"Sheet not found: {baseSheetName}", "Warning"); return; }

        Settings.SetMany(
            (KeySourceSheet, sourceSheetName),
            (KeyBaseSheet, baseSheetName),
            (KeySourceColumn, sourceColumn),
            (KeyBaseColumn, baseColumn));

        MyTools.ClearSpecialSymbols(sourceSheet, out int sourceRows, out int _);
        MyTools.ClearSpecialSymbols(baseSheet, out int baseRows, out int _);
        if (sourceRows < 2 || baseRows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

        int sourceColumnIndex = ColumnLetterToIndex(sourceColumn);
        int baseColumnIndex = ColumnLetterToIndex(baseColumn);

        var baseValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int r = 2; r <= baseRows; r++)
        {
            string value = Convert.ToString(baseSheet.Cells[r, baseColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value)) baseValues.Add(value);
        }
        if (baseValues.Count == 0) { MessageBox.Show("The base sheet has no usable values.", "Warning"); return; }

        using var progress = new ProgressForm();
        progress.Show();

        int deleted = 0;
        for (int r = sourceRows; r >= 2; r--)
        {
            string value = Convert.ToString(sourceSheet.Cells[r, sourceColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value) && baseValues.Contains(value))
            {
                sourceSheet.Rows[r].Delete();
                deleted++;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"Deleting... {r}/{sourceRows}");
                progress.SetProgress((int)((sourceRows - r) * 100.0 / sourceRows));
                Application.DoEvents();
            }
        }

        sourceSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage($"Done. {deleted} row(s) deleted.");
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show($"Finished and Saved! {deleted} row(s) deleted.", "Delete");
        Close();
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

    private static dynamic? FindSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                return ws;
        }
        return null;
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
