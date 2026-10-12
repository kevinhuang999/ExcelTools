namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Seperate the active sheet into several sheets, one per distinct value of a
/// chosen column.
///
/// Step flow matches the original add-in: open, (optionally change the column),
/// press the action button. The column is pre-filled from the last run.
/// </summary>
public class FrmSeperate : Form
{
    private const string SettingsKey = "seperate.column";

    private readonly TextBox _txtColumnHeader;

    public FrmSeperate()
    {
        Text = "Seperate";
        Size = new Size(370, 215);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var lblColumn = UiTheme.FieldLabel("Column Header:");
        lblColumn.Location = new Point(18, 26);

        _txtColumnHeader = new TextBox
        {
            Location = new Point(145, 23),
            Width = 80,
            Text = Settings.Get(SettingsKey, "A")
        };

        var btnSeperate = UiTheme.PrimaryButton("Seperate");
        btnSeperate.Location = new Point(120, 90);

        var btnCancel = UiTheme.SecondaryButton("Cancel");
        btnCancel.Location = new Point(240, 90);

        btnSeperate.Click += BtnSeperate_Click;
        btnCancel.Click += (_, _) => Close();

        Controls.Add(lblColumn);
        Controls.Add(_txtColumnHeader);
        Controls.Add(btnSeperate);
        Controls.Add(btnCancel);

        AcceptButton = btnSeperate;
        CancelButton = btnCancel;

        UiTheme.Apply(this,
            "Splits the active sheet, creating one sheet for each distinct\nvalue found in the column you name.",
            (lblColumn, "The column whose values decide how the sheet is split."),
            (_txtColumnHeader, "Column letter of the key field, for example A or AB."),
            (btnSeperate, "Split the active sheet now."),
            (btnCancel, "Close without changing anything."));

        _txtColumnHeader.SelectAll();
    }

    private void BtnSeperate_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string column = _txtColumnHeader.Text.Trim().ToUpperInvariant();
        if (string.IsNullOrEmpty(column))
        {
            MessageBox.Show("Please enter the Column Header!", "Warning");
            return;
        }
        if (!System.Text.RegularExpressions.Regex.IsMatch(column, "^[A-Z]{1,3}$"))
        {
            MessageBox.Show("The Column Header must be a column letter such as A or AB.", "Warning");
            return;
        }

        dynamic mainSheet = app.ActiveSheet;

        if (string.IsNullOrEmpty(Convert.ToString(mainSheet.Cells[1, column].Value)))
        {
            MessageBox.Show("The header should not be Empty!", "Warning");
            return;
        }

        string mainSheetName = Convert.ToString(mainSheet.Name)?.Trim() ?? "Sheet1";

        if (!MyTools.ClearSpecialSymbols(mainSheet, out int rowsCount, out int colsCount))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        app.ActiveWorkbook.Save();

        List<string>? listByField = MyTools.GetByField(app, column);
        if (listByField == null)
        {
            MessageBox.Show("Fields not found!", "Warning");
            return;
        }
        if (!MyTools.IsSheetNameValid(listByField))
        {
            MessageBox.Show("Contains strings that do not conform to the worksheet naming convention!", "Warning");
            return;
        }

        Settings.Set(SettingsKey, column);

        using var progress = new ProgressForm();
        progress.Show();

        int index = 0;
        foreach (string item in listByField)
        {
            index++;
            progress.SetMessage($"Seperating: {item}  ({index}/{listByField.Count})");
            progress.SetProgress((int)(index * 100.0 / listByField.Count));
            Application.DoEvents();

            dynamic? targetSheet = FindSheet(app, item);
            if (targetSheet == null)
            {
                targetSheet = app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
                targetSheet.Name = item;
            }
            else
            {
                targetSheet.UsedRange.Clear();
            }

            MyTools.FormatToText(targetSheet, rowsCount, colsCount);

            string sql = $"SELECT * FROM [{mainSheetName}$] WHERE [{mainSheetName}$].[{Convert.ToString(app.Cells[1, column].Value)}] = '{item.Replace("'", "''")}'";
            MyTools.CopyToNewSheet(app, sql, targetSheet);
        }

        progress.SetProgress(100);
        progress.SetMessage("Done.");
        app.ActiveWorkbook.Save();

        Close();

        if (DialogResult.Yes == MessageBox.Show("Finish and Saved! Close Now?", "Question", MessageBoxButtons.YesNo))
        {
            app.Quit();
        }
    }

    private static dynamic? FindSheet(dynamic app, string name)
    {
        foreach (dynamic ws in app.Worksheets)
        {
            if (string.Equals(Convert.ToString(ws.Name)?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                return ws;
        }
        return null;
    }
}
