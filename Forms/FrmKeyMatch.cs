namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 多字段匹配提取：按多字段联合匹配两表，匹配成功后将源表Key值填充到目标表。
/// 通用化自 FrmGet（原为条码匹配）。
/// </summary>
public class FrmKeyMatch : Form
{
    private readonly TextBox _txtSourceFile;
    private readonly TextBox _txtKeyCol;
    private readonly TextBox _txtField1Col;
    private readonly TextBox _txtField2Col;
    private readonly TextBox _txtField3Col;
    private readonly TextBox _txtTargetKeyCol;

    public FrmKeyMatch()
    {
        this.Text = "多字段匹配提取";
        this.Size = new Size(460, 290);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblFile = new Label { Text = "源文件：", Location = new Point(15, 15), AutoSize = true };
        _txtSourceFile = new TextBox { Location = new Point(100, 13), Width = 280 };
        var btnFile = new Button { Text = "浏览…", Location = new Point(390, 11), Width = 55 };
        btnFile.Click += (s, e) =>
        {
            using var dlg = new OpenFileDialog { Filter = "Excel 文件|*.xls;*.xlsx;*.xlsm" };
            if (dlg.ShowDialog() == DialogResult.OK)
                _txtSourceFile.Text = dlg.FileName;
        };

        var lblKey = new Label { Text = "源表Key列（字母）：", Location = new Point(15, 45), AutoSize = true };
        _txtKeyCol = new TextBox { Location = new Point(130, 43), Width = 60, Text = "C" };

        var lblTKey = new Label { Text = "目标表Key列（字母）：", Location = new Point(15, 72), AutoSize = true };
        _txtTargetKeyCol = new TextBox { Location = new Point(130, 70), Width = 60, Text = "B" };

        var lblF1 = new Label { Text = "匹配字段1（源列/目标列，如 L/L）：", Location = new Point(15, 99), AutoSize = true };
        _txtField1Col = new TextBox { Location = new Point(190, 97), Width = 60, Text = "L/L" };

        var lblF2 = new Label { Text = "匹配字段2：", Location = new Point(15, 126), AutoSize = true };
        _txtField2Col = new TextBox { Location = new Point(190, 124), Width = 60, Text = "N/N" };

        var lblF3 = new Label { Text = "匹配字段3：", Location = new Point(15, 153), AutoSize = true };
        _txtField3Col = new TextBox { Location = new Point(190, 151), Width = 60, Text = "O/O" };

        var lblHint = new Label
        {
            Text = "说明：按多字段联合匹配源表与目标表，\n匹配成功后将源表Key值写入目标表Key列。\n格式：源列字母/目标列字母（如 L/L）。",
            Location = new Point(15, 185), Size = new Size(420, 40), ForeColor = System.Drawing.Color.Gray
        };

        var btnOK = new Button { Text = "开始匹配", Location = new Point(200, 235), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lblFile); this.Controls.Add(_txtSourceFile); this.Controls.Add(btnFile);
        this.Controls.Add(lblKey); this.Controls.Add(_txtKeyCol);
        this.Controls.Add(lblTKey); this.Controls.Add(_txtTargetKeyCol);
        this.Controls.Add(lblF1); this.Controls.Add(_txtField1Col);
        this.Controls.Add(lblF2); this.Controls.Add(_txtField2Col);
        this.Controls.Add(lblF3); this.Controls.Add(_txtField3Col);
        this.Controls.Add(lblHint);
        this.Controls.Add(btnOK);
    }

    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string sourceFile = _txtSourceFile.Text.Trim();
        string keyCol = _txtKeyCol.Text.Trim().ToUpper();
        string targetKeyCol = _txtTargetKeyCol.Text.Trim().ToUpper();

        if (string.IsNullOrEmpty(sourceFile)) { MessageBox.Show("请选择源文件。"); return; }
        if (!File.Exists(sourceFile)) { MessageBox.Show($"文件不存在：{sourceFile}"); return; }

        var fields = ParseFields();
        if (fields.Count == 0) { MessageBox.Show("请至少输入一个匹配字段。"); return; }

        try
        {
            // Open source workbook
            dynamic wbSrc = app.Workbooks.Open(sourceFile, 0, true);
            dynamic wsSrc = wbSrc.Worksheets[1];

            // Get target sheet (active sheet)
            dynamic wsTarget = app.ActiveSheet;

            MyTools.ClearSpecialSymbols(wsSrc, out int srcRows, out int _);
            MyTools.ClearSpecialSymbols(wsTarget, out int tgtRows, out int _);
            if (srcRows < 2 || tgtRows < 2) { MessageBox.Show("数据不足。"); return; }

            int srcKeyIdx = ColumnLetterToIndex(keyCol);
            int tgtKeyIdx = ColumnLetterToIndex(targetKeyCol);
            if (srcKeyIdx <= 0 || tgtKeyIdx <= 0) return;

            using var progress = new ProgressForm();
            progress.Show();

            // Build dict from source: key -> [field1, field2, field3]
            var dictSrc = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            for (int i = 2; i <= srcRows; i++)
            {
                string key = Convert.ToString(wsSrc.Cells[i, srcKeyIdx].Value)?.Trim() ?? "";
                if (string.IsNullOrEmpty(key) || dictSrc.ContainsKey(key)) continue;

                var vals = new string[fields.Count];
                for (int f = 0; f < fields.Count; f++)
                {
                    int colIdx = ColumnLetterToIndex(fields[f].SourceCol);
                    vals[f] = Convert.ToString(wsSrc.Cells[i, colIdx].Value)?.Trim() ?? "";
                }
                dictSrc[key] = vals;
            }

            // Match and fill target
            int matchCount = 0;
            var matchedKeys = new List<string>();
            for (int i = 2; i <= tgtRows; i++)
            {
                // Build match key from target
                for (int f = 0; f < fields.Count; f++)
                {
                    int tgtColIdx = ColumnLetterToIndex(fields[f].TargetCol);
                    string tgtVal = Convert.ToString(wsTarget.Cells[i, tgtColIdx].Value)?.Trim() ?? "";
                    // Find in source dict
                    bool found = false;
                    foreach (var item in dictSrc)
                    {
                        if (item.Value[f] == tgtVal)
                        {
                            // Check all fields
                            bool allOk = true;
                            for (int f2 = 0; f2 < fields.Count; f2++)
                            {
                                int colIdx2 = ColumnLetterToIndex(fields[f2].TargetCol);
                                string v2 = Convert.ToString(wsTarget.Cells[i, colIdx2].Value)?.Trim() ?? "";
                                if (item.Value[f2] != v2) { allOk = false; break; }
                            }
                            if (allOk && !matchedKeys.Contains(item.Key))
                            {
                                wsTarget.Cells[i, tgtKeyIdx] = item.Key;
                                matchedKeys.Add(item.Key);
                                matchCount++;
                                found = true;
                                break;
                            }
                        }
                    }
                    if (!found) break;
                }

                if (i % 500 == 0)
                {
                    progress.SetMessage($"匹配中… {i}/{tgtRows}");
                    progress.SetProgress((int)(i * 100.0 / tgtRows));
                    Application.DoEvents();
                }
            }

            wbSrc.Close(false);
            progress.SetMessage($"完成！匹配 {matchCount} 条。");
            System.Threading.Thread.Sleep(1000);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"匹配出错：{ex.Message}");
        }
    }

    private List<(string SourceCol, string TargetCol)> ParseFields()
    {
        var list = new List<(string, string)>();
        var inputs = new[] { _txtField1Col.Text, _txtField2Col.Text, _txtField3Col.Text };
        foreach (var input in inputs)
        {
            var parts = input.Trim().Split('/');
            if (parts.Length >= 2 &&
                System.Text.RegularExpressions.Regex.IsMatch(parts[0].Trim().ToUpper(), "^[A-Z]{1,3}$") &&
                System.Text.RegularExpressions.Regex.IsMatch(parts[1].Trim().ToUpper(), "^[A-Z]{1,3}$"))
            {
                list.Add((parts[0].Trim().ToUpper(), parts[1].Trim().ToUpper()));
            }
        }
        return list;
    }

    private int ColumnLetterToIndex(string col)
    {
        int result = 0;
        foreach (var c in col)
        {
            if (!char.IsLetter(c)) return 0;
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }
}