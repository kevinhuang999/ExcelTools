namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Statistics: group the active sheet by one column and produce a
/// category-by-period cross table on its own sheet.
///
/// Step flow matches the original add-in: open, press the action button. All
/// three column fields are pre-filled from the last run.
/// </summary>
public class FrmGroupStat : Form
{
    private const string KeyGroupColumn = "groupStat.groupColumn";
    private const string KeyPeriodColumn = "groupStat.periodColumn";
    private const string KeyTargetSheet = "groupStat.targetSheet";

    private readonly TextBox _txtGroupColumn;
    private readonly TextBox _txtPeriodColumn;
    private readonly TextBox _txtTargetSheet;

    public FrmGroupStat()
    {
        Text = "SM Monthly Statistics";
        Size = new Size(360, 205);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(new Label { Text = "Category Column:", Location = new Point(15, 20), AutoSize = true });
        _txtGroupColumn = new TextBox
        {
            Location = new Point(140, 17),
            Width = 70,
            Text = Settings.Get(KeyGroupColumn, "A")
        };
        Controls.Add(_txtGroupColumn);

        Controls.Add(new Label { Text = "Period Column:", Location = new Point(15, 50), AutoSize = true });
        _txtPeriodColumn = new TextBox
        {
            Location = new Point(140, 47),
            Width = 70,
            Text = Settings.Get(KeyPeriodColumn, "B")
        };
        Controls.Add(_txtPeriodColumn);

        Controls.Add(new Label { Text = "Output Sheet:", Location = new Point(15, 80), AutoSize = true });
        _txtTargetSheet = new TextBox
        {
            Location = new Point(140, 77),
            Width = 120,
            Text = Settings.Get(KeyTargetSheet, "Statistics")
        };
        Controls.Add(_txtTargetSheet);

        var btnStatistics = new Button
        {
            Text = "Statistics",
            Location = new Point(150, 125),
            Size = new Size(90, 28)
        };
        btnStatistics.Click += BtnStatistics_Click;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(250, 125),
            Size = new Size(90, 28)
        };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnStatistics);
        Controls.Add(btnCancel);

        AcceptButton = btnStatistics;
        UiTheme.Apply(this, "Groups the active sheet by one column and writes a\ncategory-by-period cross table to a new sheet.");
        CancelButton = btnCancel;
    }

    private void BtnStatistics_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string groupColumn = _txtGroupColumn.Text.Trim().ToUpperInvariant();
        string periodColumn = _txtPeriodColumn.Text.Trim().ToUpperInvariant();
        string targetSheetName = _txtTargetSheet.Text.Trim();

        if (!IsColumnLetter(groupColumn) || !IsColumnLetter(periodColumn))
        {
            MessageBox.Show("Columns must be column letters such as A or AB.", "Warning");
            return;
        }
        if (string.IsNullOrEmpty(targetSheetName))
        {
            MessageBox.Show("Please enter an output sheet name.", "Warning");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int _))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }
        app.ActiveWorkbook.Save();

        Settings.SetMany(
            (KeyGroupColumn, groupColumn),
            (KeyPeriodColumn, periodColumn),
            (KeyTargetSheet, targetSheetName));

        int groupColumnIndex = ColumnLetterToIndex(groupColumn);
        int periodColumnIndex = ColumnLetterToIndex(periodColumn);

        var groups = new List<string>();
        var counts = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        using var progress = new ProgressForm();
        progress.Show();

        for (int r = 2; r <= rows; r++)
        {
            string group = Convert.ToString(ws.Cells[r, groupColumnIndex].Value)?.Trim() ?? "";
            string period = Convert.ToString(ws.Cells[r, periodColumnIndex].Value)?.Trim() ?? "";
            if (string.IsNullOrEmpty(group)) continue;

            if (!counts.TryGetValue(group, out var byPeriod))
            {
                groups.Add(group);
                byPeriod = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                counts[group] = byPeriod;
            }

            if (!string.IsNullOrEmpty(period))
            {
                byPeriod.TryGetValue(period, out int current);
                byPeriod[period] = current + 1;
            }

            if (r % 500 == 0)
            {
                progress.SetMessage($"Reading... {r}/{rows}");
                progress.SetProgress((int)(r * 50.0 / rows));
                Application.DoEvents();
            }
        }

        if (groups.Count == 0)
        {
            MessageBox.Show("No usable category values found.", "Warning");
            return;
        }

        var periods = new List<string>();
        foreach (var byPeriod in counts.Values)
            foreach (string p in byPeriod.Keys)
                if (!periods.Contains(p)) periods.Add(p);
        periods.Sort(StringComparer.OrdinalIgnoreCase);

        dynamic? targetSheet = FindSheet(app, targetSheetName);
        if (targetSheet == null)
        {
            targetSheet = app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
            targetSheet.Name = targetSheetName;
        }
        else
        {
            targetSheet.UsedRange.Clear();
        }

        targetSheet.Cells[1, 1] = groupColumn;
        for (int c = 0; c < periods.Count; c++)
            targetSheet.Cells[1, c + 2] = periods[c];
        targetSheet.Cells[1, periods.Count + 2] = "Total";

        for (int g = 0; g < groups.Count; g++)
        {
            targetSheet.Cells[g + 2, 1] = groups[g];
            var byPeriod = counts[groups[g]];
            int rowTotal = 0;
            for (int c = 0; c < periods.Count; c++)
            {
                byPeriod.TryGetValue(periods[c], out int value);
                targetSheet.Cells[g + 2, c + 2] = value;
                rowTotal += value;
            }
            targetSheet.Cells[g + 2, periods.Count + 2] = rowTotal;
            progress.SetProgress(50 + (int)((g + 1) * 50.0 / groups.Count));
            Application.DoEvents();
        }

        targetSheet.Columns.AutoFit();
        progress.SetProgress(100);
        progress.SetMessage("Done.");
        app.ActiveWorkbook.Save();

        Close();

        if (DialogResult.Yes == MessageBox.Show("Finish and Saved! Close Now?", "Question", MessageBoxButtons.YesNo))
        {
            app.Quit();
        }
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
