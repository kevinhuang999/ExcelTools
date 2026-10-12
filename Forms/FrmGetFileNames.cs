namespace ExcelTools.Forms;

using ExcelDna.Integration;
using System.IO;

/// <summary>
/// List the file names under a chosen folder, recursively, into the active
/// sheet. Launched straight from the ribbon button - no dialog of its own,
/// matching the original add-in.
/// </summary>
internal static class FrmGetFileNames
{
    public static void Run()
    {
        dynamic app = (dynamic)ExcelDnaUtil.Application;
        if (app == null) { MessageBox.Show("无法获取 Excel 对象。", "提示"); return; }

        using var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() != DialogResult.OK) return;

        string rootPath = dialog.SelectedPath;
        if (string.IsNullOrEmpty(rootPath) || !Directory.Exists(rootPath)) return;

        try
        {
            var files = Directory.EnumerateFiles(rootPath, "*.*", SearchOption.AllDirectories)
                .Where(f => !f.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f)
                .ToList();

            if (files.Count == 0) { MessageBox.Show("未找到文件。", "提示"); return; }

            dynamic ws = app.ActiveSheet;
            int startRow = (int)ws.UsedRange.Rows.Count + 1;

            ws.Cells[startRow, 1] = "序号";
            ws.Cells[startRow, 2] = "文件名";
            ws.Cells[startRow, 3] = "相对路径";
            ws.Cells[startRow, 4] = "大小(KB)";

            for (int i = 0; i < files.Count; i++)
            {
                int row = startRow + 1 + i;
                ws.Cells[row, 1] = i + 1;
                ws.Cells[row, 2] = Path.GetFileName(files[i]);
                ws.Cells[row, 3] = Path.GetRelativePath(rootPath, files[i]).Replace('\\', '/');

                try
                {
                    var info = new FileInfo(files[i]);
                    ws.Cells[row, 4] = Math.Round(info.Length / 1024.0, 1);
                }
                catch
                {
                    ws.Cells[row, 4] = "";
                }
            }

            ws.Columns.AutoFit();
            app.ActiveWorkbook.Save();
            MessageBox.Show($"已完成并保存！ {files.Count} file name(s) written.", "File Names");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"获取文件列表失败：{ex.Message}", "错误");
        }
    }
}
