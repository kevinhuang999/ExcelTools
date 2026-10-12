namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Running distinct count: for each row, count how many distinct values have
/// appeared so far in the chosen column.
///
/// Step flow matches the original add-in: open, press the action button, watch
/// the progress bar. Both columns are pre-filled from the last run.
/// </summary>
public class FrmCountDistinct : Form
{
    private const string KeyCountColumn = "countDistinct.countColumn";
    private const string KeyOutputColumn = "countDistinct.outputColumn";

    private readonly TextBox _txtCountColumn;
    private readonly TextBox _txtOutputColumn;
    private readonly ProgressBar _progressBar;
    private readonly Label _lblStatus;

    public FrmCountDistinct()
    {
        Text = "Count Distinct";
        Size = new Size(380, 200);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(new Label { Text = "Count Column:", Location = new Point(15, 20), AutoSize = true });
        _txtCountColumn = new TextBox
        {
            Location = new Point(150, 17),
            Width = 70,
            Text = Settings.Get(KeyCountColumn, "A")
        };
        Controls.Add(_txtCountColumn);

        Controls.Add(new Label { Text = "Output Column:", Location = new Point(15, 50), AutoSize = true });
        _txtOutputColumn = new TextBox
        {
            Location = new Point(150, 47),
            Width = 70,
            Text = Settings.Get(KeyOutputColumn, "Z")
        };
        Controls.Add(_txtOutputColumn);

        _progressBar = new ProgressBar
        {
            Location = new Point(15, 82),
            Size = new Size(330, 20),
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous
        };
        Controls.Add(_progressBar);

        _lblStatus = new Label { Location = new Point(15, 106), Size = new Size(330, 18), Text = "" };
        Controls.Add(_lblStatus);

        var btnSum = new Button
        {
            Text = "Sum",
            Location = new Point(165, 130),
            Size = new Size(85, 28)
        };
        btnSum.Click += BtnSum_Click;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(260, 130),
            Size = new Size(85, 28)
        };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(btnSum);
        Controls.Add(btnCancel);

        AcceptButton = btnSum;
        UiTheme.Apply(this, "For each row, counts how many distinct values have appeared\nso far in the column you name.");
        CancelButton = btnCancel;
    }

    private void BtnSum_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string countColumn = _txtCountColumn.Text.Trim().ToUpperInvariant();
        string outputColumn = _txtOutputColumn.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(countColumn) || !IsColumnLetter(outputColumn))
        {
            MessageBox.Show("Columns must be column letters such as A or AB.", "Warning");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("Not enough data.", "Warning"); return; }

        int countColumnIndex = ColumnLetterToIndex(countColumn);
        int outputColumnIndex = ColumnLetterToIndex(outputColumn);

        if (outputColumnIndex <= cols)
        {
            MessageBox.Show("The output column must be outside the data range.", "Warning");
            return;
        }

        Settings.SetMany((KeyCountColumn, countColumn), (KeyOutputColumn, outputColumn));

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ws.Cells[1, outputColumnIndex] = "DISTINCT_COUNT";

        _progressBar.Minimum = 0;
        _progressBar.Maximum = rows;
        _progressBar.Value = 0;

        for (int r = 2; r <= rows; r++)
        {
            string value = Convert.ToString(ws.Cells[r, countColumnIndex].Value)?.Trim() ?? "";
            if (!string.IsNullOrEmpty(value)) seen.Add(value);

            ws.Cells[r, outputColumnIndex] = seen.Count;

            _lblStatus.Text = r + "/" + rows;
            _progressBar.Value = r;
            if (r % 200 == 0) Application.DoEvents();
        }

        ws.Columns.AutoFit();
        _progressBar.Value = rows;
        _lblStatus.Text = "";
        app.ActiveWorkbook.Save();

        if (DialogResult.OK == MessageBox.Show("Success!", "Count Distinct"))
        {
            Close();
        }
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
