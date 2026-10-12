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
        Text = "多字段匹配";
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new FormLayout(this,
            "按最多三个字段联合匹配「源文件」与当前工作表，\n匹配成功后把源表的 Key 值写入目标表 Key 列。\n字段格式：源列/目标列，例如 L/L");

        var (sourceBox, btnOpen) = layout.AddFieldWithButton("源文件:", Settings.Get(KeySourceFile, ""), "打开");
        _txtSourceFile = sourceBox;

        btnOpen.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };
            if (dialog.ShowDialog() == DialogResult.OK) _txtSourceFile.Text = dialog.FileName;
        };

        _txtSourceKeyColumn = layout.AddField("源表Key列:", Settings.Get(KeySourceKeyColumn, "C"));
        _txtTargetKeyColumn = layout.AddField("目标表Key列:", Settings.Get(KeyTargetKeyColumn, "B"));
        _txtField1 = layout.AddField("匹配字段1:", Settings.Get(KeyField1, ""));
        _txtField2 = layout.AddField("匹配字段2:", Settings.Get(KeyField2, ""));
        _txtField3 = layout.AddField("匹配字段3:", Settings.Get(KeyField3, ""));

        var buttons = layout.Finish(("确定", true), ("取消", false));
        buttons[0].Click += BtnOK_Click;
        buttons[1].Click += (_, _) => Close();

        AcceptButton = buttons[0];
        CancelButton = buttons[1];

        UiTheme.Apply(this);
    }


    private void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        string sourceFile = _txtSourceFile.Text.Trim();
        string sourceKeyColumn = _txtSourceKeyColumn.Text.Trim().ToUpperInvariant();
        string targetKeyColumn = _txtTargetKeyColumn.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrEmpty(sourceFile)) { MessageBox.Show("请选择源文件。", "提示"); return; }
        if (!File.Exists(sourceFile)) { MessageBox.Show($"文件不存在：{sourceFile}", "提示"); return; }
        if (!IsColumnLetter(sourceKeyColumn) || !IsColumnLetter(targetKeyColumn))
        {
            MessageBox.Show("Key 列必须是列字母，例如 A 或 AB。", "提示");
            return;
        }

        var fields = ParseFields();
        if (fields.Count == 0) { MessageBox.Show("请至少输入一个匹配字段。", "提示"); return; }

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
            if (sourceRows < 2 || targetRows < 2) { MessageBox.Show("数据不足。", "提示"); return; }

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
                    progress.SetMessage($"正在匹配… {i}/{targetRows}");
                    progress.SetProgress((int)(i * 100.0 / targetRows));
                    Application.DoEvents();
                }
            }

            app.ActiveWorkbook.Save();
            progress.SetProgress(100);
            progress.SetMessage($"完成。 {matched} row(s) matched.");
            Application.DoEvents();
            System.Threading.Thread.Sleep(600);

            MessageBox.Show($"已完成并保存！ {matched} row(s) matched.", "Get Key");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"匹配失败：{ex.Message}", "错误");
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
