namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Match a source workbook against the active sheet on up to three joint
/// fields, then write the source key into the target key column.
///
/// Step flow matches the original add-in: choose the source file, press OK.
/// Every field is pre-filled from the last run.
/// </summary>
public class FrmKeyMatch : Form
{
    private const string KeySourceFile = "keyMatch.sourceFile";
    private const string KeySourceKeyColumn = "keyMatch.sourceKeyColumn";
    private const string KeyTargetKeyColumn = "keyMatch.targetKeyColumn";
    private const string KeyField1 = "keyMatch.field1";
    private const string KeyField2 = "keyMatch.field2";
    private const string KeyField3 = "keyMatch.field3";

    private readonly TextBox _txtSourceFile;
    private readonly TextBox _txtSourceKeyColumn;
    private readonly TextBox _txtTargetKeyColumn;
    private readonly TextBox _txtField1;
    private readonly TextBox _txtField2;
    private readonly TextBox _txtField3;

    public FrmKeyMatch()
    {
        Text = "Get Key";
        Size = new Size(400, 260);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(new Label { Text = "Source File:", Location = new Point(15, 20), AutoSize = true });
        _txtSourceFile = new TextBox { Location = new Point(110, 17), Width = 200, Text = Settings.Get(KeySourceFile, "") };
        Controls.Add(_txtSourceFile);

        var btnOpen = new Button { Text = "Open Excel", Location = new Point(315, 15), Size = new Size(70, 26) };
        btnOpen.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };
            if (dialog.ShowDialog() == DialogResult.OK) _txtSourceFile.Text = dialog.FileName;
        };
        Controls.Add(btnOpen);

        _txtSourceKeyColumn = AddRow("Source Key Col:", 50, Settings.Get(KeySourceKeyColumn, "C"));
        _txtTargetKeyColumn = AddRow("Target Key Col:", 80, Settings.Get(KeyTargetKeyColumn, "B"));
        _txtField1 = AddRow("Field1 (src/tgt):", 110, Settings.Get(KeyField1, ""));
        _txtField2 = AddRow("Field2 (src/tgt):", 140, Settings.Get(KeyField2, ""));
        _txtField3 = AddRow("Field3 (src/tgt):", 170, Settings.Get(KeyField3, ""), 130);

        var btnOK = new Button { Text = "OK", Location = new Point(205, 205), Size = new Size(85, 28) };
        btnOK.Click += BtnOK_Click;

        var btnCancel = new Button { Text = "Cancel", Location = new Point(300, 205), Size = new Size(85, 28) };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnOK);
        Controls.Add(btnCancel);

        AcceptButton = btnOK;
        UiTheme.Apply(this, "Matches a source workbook against the active sheet on up to\nthree fields. Field format:  sourceColumn/targetColumn");
        CancelButton = btnCancel;
    }

    private TextBox AddRow(string label, int top, string initial, int labelWidth = 115)
    {
        Controls.Add(new Label { Text = label, Location = new Point(15, top + 3), AutoSize = true });
        var box = new TextBox { Location = new Point(labelWidth, top), Width = 120, Text = initial };
        Controls.Add(box);
        return box;
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string sourceFile = _txtSourceFile.Text.Trim();
        string sourceKeyColumn = _txtSourceKeyColumn.Text.Trim().ToUpperInvariant();
        string targetKeyColumn = _txtTargetKeyColumn.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sourceFile)) { MessageBox.Show("Please choose the source file.", "Warning"); return; }
        if (!File.Exists(sourceFile)) { MessageBox.Show($"File not found: {sourceFile}", "Warning"); return; }
        if (!IsColumnLetter(sourceKeyColumn) || !IsColumnLetter(targetKeyColumn))
        {
            MessageBox.Show("Key columns must be column letters such as A or AB.", "Warning");
            return;
        }

        var fields = ParseFields();
        if (fields.Count == 0) { MessageBox.Show("Please enter at least one match field.", "Warning"); return; }

        Settings.SetMany(
            (KeySourceFile, sourceFile),
            (KeySourceKeyColumn, sourceKeyColumn),
            (KeyTargetKeyColumn, targetKeyColumn),
            (KeyField1, _txtField1.Text.Trim()),
            (KeyField2, _txtField2.Text.Trim()),
            (KeyField3, _txtField3.Text.Trim()));

        dynamic? sourceWorkbook = null;
        try
        {
            sourceWorkbook = app.Workbooks.Open(sourceFile, 0, true);
            dynamic sourceSheet = sourceWorkbook.Worksheets[1];
            dynamic targetSheet = app.ActiveSheet;

            MyTools.ClearSpecialSymbols(sourceSheet, out int sourceRows, out int _);
            MyTools.ClearSpecialSymbols(targetSheet, out int targetRows, out int _);
            if (sourceRows < 2 || targetRows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

            int sourceKeyIndex = ColumnLetterToIndex(sourceKeyColumn);
            int targetKeyIndex = ColumnLetterToIndex(targetKeyColumn);

            using var progress = new ProgressForm();
            progress.Show();

            var lookup = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            for (int i = 2; i <= sourceRows; i++)
            {
                string key = Convert.ToString(sourceSheet.Cells[i, sourceKeyIndex].Value)?.Trim() ?? "";
                if (string.IsNullOrEmpty(key) || lookup.ContainsKey(key)) continue;

                var values = new string[fields.Count];
                for (int f = 0; f < fields.Count; f++)
                    values[f] = Convert.ToString(sourceSheet.Cells[i, ColumnLetterToIndex(fields[f].SourceColumn)].Value)?.Trim() ?? "";
                lookup[key] = values;
            }

            int matched = 0;
            for (int i = 2; i <= targetRows; i++)
            {
                string targetKey = Convert.ToString(targetSheet.Cells[i, targetKeyIndex].Value)?.Trim() ?? "";

                foreach (var pair in lookup)
                {
                    if (targetKey.Equals(pair.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        bool allMatch = true;
                        for (int f = 0; f < fields.Count; f++)
                        {
                            string targetValue = Convert.ToString(targetSheet.Cells[i, ColumnLetterToIndex(fields[f].TargetColumn)].Value)?.Trim() ?? "";
                            if (!string.Equals(pair.Value[f], targetValue, StringComparison.OrdinalIgnoreCase))
                            {
                                allMatch = false;
                                break;
                            }
                        }

                        if (allMatch)
                        {
                            targetSheet.Cells[i, targetKeyIndex] = pair.Key;
                            matched++;
                        }
                        break;
                    }
                }

                if (i % 500 == 0)
                {
                    progress.SetMessage($"Matching... {i}/{targetRows}");
                    progress.SetProgress((int)(i * 100.0 / targetRows));
                    Application.DoEvents();
                }
            }

            app.ActiveWorkbook.Save();
            progress.SetProgress(100);
            progress.SetMessage($"Done. {matched} row(s) matched.");
            Application.DoEvents();
            System.Threading.Thread.Sleep(600);

            MessageBox.Show($"Finished and Saved! {matched} row(s) matched.", "Get Key");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Match failed: {ex.Message}", "Error");
        }
        finally
        {
            try { sourceWorkbook?.Close(false); } catch { }
        }
    }

    private List<(string SourceColumn, string TargetColumn)> ParseFields()
    {
        var list = new List<(string, string)>();
        foreach (string input in new[] { _txtField1.Text, _txtField2.Text, _txtField3.Text })
        {
            string[] parts = input.Trim().Split('/');
            if (parts.Length >= 2 &&
                IsColumnLetter(parts[0].Trim().ToUpperInvariant()) &&
                IsColumnLetter(parts[1].Trim().ToUpperInvariant()))
            {
                list.Add((parts[0].Trim().ToUpperInvariant(), parts[1].Trim().ToUpperInvariant()));
            }
        }
        return list;
    }

    private static bool IsColumnLetter(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Z]{1,3}$");

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
