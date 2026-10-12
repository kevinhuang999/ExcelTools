namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// Sort the active sheet by one or two columns. The original data is backed up
/// to an "Original" sheet first.
///
/// Step flow matches the original add-in: open, press the action button. The
/// sort columns are pre-filled from the last run.
/// </summary>
public class FrmSort : Form
{
    private const string KeyCol1 = "sort.column1";
    private const string KeyCol2 = "sort.column2";

    private readonly TextBox _txtSortColumn1;
    private readonly TextBox _txtSortColumn2;

    public FrmSort()
    {
        Text = "Sort";
        Size = new Size(340, 165);
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var lblColumn1 = new Label
        {
            Text = "Sort Column:",
            Location = new Point(15, 22),
            AutoSize = true
        };
        _txtSortColumn1 = new TextBox
        {
            Location = new Point(135, 19),
            Width = 70,
            Text = Settings.Get(KeyCol1, "A")
        };

        var lblColumn2 = new Label
        {
            Text = "Then By:",
            Location = new Point(15, 52),
            AutoSize = true
        };
        _txtSortColumn2 = new TextBox
        {
            Location = new Point(135, 49),
            Width = 70,
            Text = Settings.Get(KeyCol2, "")
        };

        var btnSort = new Button
        {
            Text = "Sort",
            Location = new Point(120, 95),
            Size = new Size(90, 28)
        };
        btnSort.Click += BtnSort_Click;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(220, 95),
            Size = new Size(90, 28)
        };
        btnCancel.Click += (_, _) => Close();

        Controls.Add(lblColumn1);
        Controls.Add(_txtSortColumn1);
        Controls.Add(lblColumn2);
        Controls.Add(_txtSortColumn2);
        Controls.Add(btnSort);
        Controls.Add(btnCancel);

        AcceptButton = btnSort;
        UiTheme.Apply(this, "Sorts the active sheet by the column(s) you name.\nThe original data is copied to an Original sheet first.");
        CancelButton = btnCancel;
    }

    private void BtnSort_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("Unable to access Excel.", "Warning"); return; }

        string column1 = _txtSortColumn1.Text.Trim().ToUpperInvariant();
        string column2 = _txtSortColumn2.Text.Trim().ToUpperInvariant();

        if (!IsColumnLetter(column1))
        {
            MessageBox.Show("Sort Column must be a column letter such as A or AB.", "Warning");
            return;
        }
        if (!string.IsNullOrEmpty(column2) && !IsColumnLetter(column2))
        {
            MessageBox.Show("Then By must be a column letter, or left empty.", "Warning");
            return;
        }

        dynamic ws = app.ActiveSheet;
        if (!MyTools.ClearSpecialSymbols(ws, out int rows, out int cols))
        {
            MessageBox.Show("Empty table or unable to clear special symbols!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rows < 2) { MessageBox.Show("Not enough data to sort.", "Warning"); return; }

        Settings.SetMany((KeyCol1, column1), (KeyCol2, column2));

        MyTools.CopyToNewSheet(app, Convert.ToString(ws.Name)?.Trim() ?? "Sheet1");

        int sortIndex1 = ColumnLetterToIndex(column1);
        int helperColumn = cols + 1;

        for (int r = 2; r <= rows; r++)
        {
            string key1 = Convert.ToString(ws.Cells[r, sortIndex1].Value)?.Trim() ?? "";
            string key2 = string.IsNullOrEmpty(column2)
                ? ""
                : Convert.ToString(ws.Cells[r, ColumnLetterToIndex(column2)].Value)?.Trim() ?? "";

            ws.Cells[r, helperColumn] = string.IsNullOrEmpty(column2) ? key1 : key1 + "\t" + key2;
        }

        try
        {
            dynamic sortRange = ws.Range[ws.Cells[1, 1], ws.Cells[rows, cols]];
            dynamic sortKey = ws.Range[ws.Cells[1, helperColumn], ws.Cells[rows, helperColumn]];
            const int xlAscending = 1;
            const int xlYesHeader = 1;
            const int xlTopToBottom = 1;

            // Range.Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3,
            //            Header, OrderCustom, MatchCase, Orientation, SortMethod)
            sortRange.Sort(
                sortKey,
                xlAscending,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                xlYesHeader,
                Type.Missing,
                false,
                xlTopToBottom,
                Type.Missing);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sort failed: {ex.Message}", "Error");
            return;
        }

        ws.Columns[helperColumn].Delete();
        ws.Columns.AutoFit();
        app.ActiveWorkbook.Save();

        MessageBox.Show("Finished and Saved!", "Sort");
        Close();
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
