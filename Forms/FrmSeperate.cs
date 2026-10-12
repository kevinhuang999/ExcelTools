namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;

/// <summary>
/// 拆分工作表：按指定列的值将总表拆成多个工作表。
/// </summary>
public class FrmSeperate : Form
{
    private readonly TextBox _txtColumn;

    public FrmSeperate()
    {
        this.Text = "拆分工作表";
        this.Size = new Size(360, 170);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lbl1 = new Label { Text = "按哪列拆分（列字母，如 A、B）：", Location = new Point(15, 20), AutoSize = true };
        _txtColumn = new TextBox { Location = new Point(180, 18), Width = 100, Text = "B" };

        var btnOK = new Button { Text = "开始拆分", Location = new Point(180, 70), Size = new Size(100, 30) };
        btnOK.Click += BtnOK_Click;

        this.Controls.Add(lbl1);
        this.Controls.Add(_txtColumn);
        this.Controls.Add(btnOK);
    }

    private async void BtnOK_Click(object? sender, EventArgs e)
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        string column = _txtColumn.Text.Trim().ToUpper();
        if (string.IsNullOrEmpty(column)) { MessageBox.Show("请输入列字母。"); return; }

        if (!System.Text.RegularExpressions.Regex.IsMatch(column, "^[A-Z]{1,3}$"))
        { MessageBox.Show("列字母格式不正确。"); return; }

        dynamic ws = app.ActiveSheet;
        MyTools.ClearSpecialSymbols(ws, out int rows, out int cols);
        if (rows < 2) { MessageBox.Show("数据不足。"); return; }

        var values = MyTools.GetByField(app, column);
        if (values == null || values.Count == 0) { MessageBox.Show("未找到有效数据。"); return; }

        using var progress = new ProgressForm();
        progress.Show();

        int total = values.Count;
        int i = 0;
        foreach (var val in values)
        {
            i++;
            progress.SetMessage($"正在拆分：{val}（{i}/{total}）");
            progress.SetProgress((int)(i * 100.0 / total));

            string sheetName = val.Length > 31 ? val.Substring(0, 31) : val;
            if (!MyTools.IsSheetNameValid(sheetName))
            {
                progress.SetMessage($"跳过无效名称：{val}");
                continue;
            }

            string sheetSql = $"'[\"{app.ActiveWorkbook.Name}\"]{ws.Name.Trim()}!\"{column}:{column}\"'";
            string sql = $"SELECT * FROM [{ws.Name.Trim()}] WHERE [{column}]='{val.Replace("'", "''")}'";
            try
            {
                dynamic newWs = app.Worksheets.Add(After: ws);
                newWs.Name = sheetName;
                MyTools.CopyToNewSheet(app, sql, newWs);
                MyTools.FormatToText(newWs, rows, cols);
            }
            catch { }

            await Task.Yield();
        }

        progress.SetProgress(100);
        progress.SetMessage("拆分完成！");
        System.Threading.Thread.Sleep(1000);
    }
}