namespace ExcelTools.Forms;

using ExcelDna.Integration;
using System.IO;

/// <summary>
/// 文件名列表：弹文件夹选择器，递归获取目录下所有文件名，写入当前活动工作表。
/// </summary>
internal static class FrmGetFileNames
{
    public static void Run()
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。"); return; }

        using var dlg = new FolderBrowserDialog();
        if (dlg.ShowDialog() != DialogResult.OK) return;

        string rootPath = dlg.SelectedPath;
        if (string.IsNullOrEmpty(rootPath) || !Directory.Exists(rootPath)) return;

        try
        {
            var files = Directory.EnumerateFiles(rootPath, "*.*", SearchOption.AllDirectories)
                .Where(f => !f.EndsWith(".lnk")) // skip shortcuts
                .OrderBy(f => f)
                .ToList();

            if (files.Count == 0) { MessageBox.Show("未找到文件。"); return; }

            dynamic ws = app.ActiveSheet;
            int startRow = ws.UsedRange.Rows.Count + 1;

            // Header
            ws.Cells[startRow, 1] = "序号";
            ws.Cells[startRow, 2] = "文件名";
            ws.Cells[startRow, 3] = "相对路径";
            ws.Cells[startRow, 4] = "文件大小(KB)";

            // Data
            for (int i = 0; i < files.Count; i++)
            {
                int row = startRow + 1 + i;
                ws.Cells[row, 1] = i + 1;
                ws.Cells[row, 2] = Path.GetFileName(files[i]);
                ws.Cells[row, 3] = Path.GetRelativePath(rootPath, files[i]).Replace('\\', '/');

                try
                {
                    var fi = new FileInfo(files[i]);
                    ws.Cells[row, 4] = Math.Round(fi.Length / 1024.0, 1);
                }
                catch
                {
                    ws.Cells[row, 4] = "";
                }
            }

            ws.Columns.AutoFit();
            MessageBox.Show($"已写入 {files.Count} 个文件名。");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"获取文件列表失败：{ex.Message}");
        }
    }
}