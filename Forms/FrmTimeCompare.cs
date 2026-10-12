namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Match two sheets on a key column and list rows whose comparison column
/// differs. Results go to a "DiffResult" sheet.
///
/// Step flow: open, press the action button. All fields are pre-filled from the
/// last run.
/// </summary>
public class FrmTimeCompare : Form
{
    private const string KeySheet1 = "diff.sheet1";
    private const string KeySheet2 = "diff.sheet2";
    private const string KeyMatch1 = "diff.match1";
    private const string KeyCompare1 = "diff.compare1";
    private const string KeyMatch2 = "diff.match2";
    private const string KeyCompare2 = "diff.compare2";

    private readonly TextBox _txtSheet1;
    private readonly TextBox _txtSheet2;
    private readonly TextBox _txtMatchColumn1;
    private readonly TextBox _txtCompareColumn1;
    private readonly TextBox _txtMatchColumn2;
    private readonly TextBox _txtCompareColumn2;

    public FrmTimeCompare()
    {
        Text = "Find Differences";
        Size = new Size(380, 235);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        _txtSheet1 = AddRow("Sheet1:", 18, Settings.Get(KeySheet1, ""));
        _txtMatchColumn1 = AddRow("Match Col1:", 48, Settings.Get(KeyMatch1, "A"));
        _txtCompareColumn1 = AddRow("Compare Col1:", 78, Settings.Get(KeyCompare1, "B"), 130);
        _txtSheet2 = AddRow("Sheet2:", 108, Settings.Get(KeySheet2, ""));
        _txtMatchColumn2 = AddRow("Match Col2:", 138, Settings.Get(KeyMatch2, "A"));
        _txtCompareColumn2 = AddRow("Compare Col2:", 168, Settings.Get(KeyCompare2, "B"), 130);

        var btnMatch = new Button { Text = "Match", Location = new Point(160, 205), Size = new Size(90, 28) };
        btnMatch.Click += BtnMatch_Click;

        var btnCancel = new Button { Text = "Cancel", Location = new Point(260, 205), Size = new Size(90, 28) };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnMatch);
        Controls.Add(btnCancel);

        AcceptButton = btnMatch;
        UiTheme.Apply(this, "Matches two sheets on a key column and lists the rows whose\ncomparison column differs.");
        CancelButton = btnCancel;
    }

    private TextBox AddRow(string label, int top, string initial, int labelWidth = 95)
    {
        Controls.Add(new Label { Text = label, Location = new Point(15, top + 3), AutoSize = true });
        var box = new TextBox { Location = new Point(labelWidth, top), Width = 120, Text = initial };
        Controls.Add(box);
        return box;
    }

    private void BtnMatch_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string sheet1Name = _txtSheet1.Text.Trim();
        string sheet2Name = _txtSheet2.Text.Trim();
        string matchColumn1 = _txtMatchColumn1.Text.Trim().ToUpperInvariant();
        string compareColumn1 = _txtCompareColumn1.Text.Trim().ToUpperInvariant();
        string matchColumn2 = _txtMatchColumn2.Text.Trim().ToUpperInvariant();
        string compareColumn2 = _txtCompareColumn2.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sheet1Name) || string.IsNullOrEmpty(sheet2Name))
        {
            MessageBox.Show("Please enter both sheet names.", "Warning");
            return;
        }
        if (!IsColumnLetter(matchColumn1) || !IsColumnLetter(compareColumn1) ||
            !IsColumnLetter(matchColumn2) || !IsColumnLetter(compareColumn2))
        {
            MessageBox.Show("Columns must be column letters such as A or AB.", "Warning");
            return;
        }

        dynamic? sheet1 = FindSheet(app, sheet1Name);
        dynamic? sheet2 = FindSheet(app, sheet2Name);
        if (sheet1 == null) { MessageBox.Show($"Sheet not found: {sheet1Name}", "Warning"); return; }
        if (sheet2 == null) { MessageBox.Show($"Sheet not found: {sheet2Name}", "Warning"); return; }

        MyTools.ClearSpecialSymbols(sheet1, out int rows1, out int _);
        MyTools.ClearSpecialSymbols(sheet2, out int rows2, out int _);
        if (rows1 < 2 || rows2 < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

        Settings.SetMany(
            (KeySheet1, sheet1Name), (KeySheet2, sheet2Name),
            (KeyMatch1, matchColumn1), (KeyCompare1, compareColumn1),
            (KeyMatch2, matchColumn2), (KeyCompare2, compareColumn2));

        int matchIndex1 = ColumnLetterToIndex(matchColumn1);
        int compareIndex1 = ColumnLetterToIndex(compareColumn1);
        int matchIndex2 = ColumnLetterToIndex(matchColumn2);
        int compareIndex2 = ColumnLetterToIndex(compareColumn2);

        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 2; i <= rows1; i++)
        {
            string key = Convert.ToString(sheet1.Cells[i, matchIndex1].Value)?.Trim() ?? "";
            string value = Convert.ToString(sheet1.Cells[i, compareIndex1].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(key) && !lookup.ContainsKey(key))
                lookup[key] = value;
        }

        dynamic targetSheet = CreateTargetSheet(app, "DiffResult");
        targetSheet.Cells[1, 1] = "Key";
        targetSheet.Cells[1, 2] = "Sheet1 Value";
        targetSheet.Cells[1, 3] = "Sheet2 Value";

        using var progress = new ProgressForm();
        progress.Show();

        int differences = 0;
        for (int i = 2; i <= rows2; i++)
        {
            string key = Convert.ToString(sheet2.Cells[i, matchIndex2].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(key) && lookup.TryGetValue(key, out string? value1))
            {
                string value2 = Convert.ToString(sheet2.Cells[i, compareIndex2].Value)?.Trim() ?? "";
                if (!string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase))
                {
                    differences++;
                    targetSheet.Cells[differences + 1, 1] = key;
                    targetSheet.Cells[differences + 1, 2] = value1;
                    targetSheet.Cells[differences + 1, 3] = value2;
                }
            }

            if (i % 200 == 0)
            {
                progress.SetMessage($"Comparing... {i}/{rows2}");
                progress.SetProgress((int)(i * 100.0 / rows2));
                Application.DoEvents();
            }
        }

        targetSheet.Columns.AutoFit();
        app.ActiveWorkbook.Save();
        progress.SetProgress(100);
        progress.SetMessage($"Done. {differences} difference(s).");
        Application.DoEvents();
        System.Threading.Thread.Sleep(600);

        MessageBox.Show($"Finished and Saved! {differences} difference(s).", "Find Differences");
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
