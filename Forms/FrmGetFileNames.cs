namespace ExcelTools.Forms;

using ExcelDna.Integration;
using ExcelTools.Common;
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
        if (app == null) { MessageBox.Show(Loc.T("无法获取 Excel 对象。", "Unable to access Excel."), Loc.T("提示", "Warning")); return; }

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

            if (files.Count == 0) { MessageBox.Show(Loc.T("未找到文件。", "No files found."), Loc.T("提示", "Warning")); return; }

            dynamic ws = app.ActiveSheet;
            int startRow = (int)ws.UsedRange.Rows.Count + 1;

            ws.Cells[startRow, 1] = Loc.T("序号", "No.");
            ws.Cells[startRow, 2] = Loc.T("文件名", "File Name");
            ws.Cells[startRow, 3] = Loc.T("相对路径", "Relative Path");
            ws.Cells[startRow, 4] = Loc.T("大小(KB)", "Size (KB)");

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
            MessageBox.Show(Loc.T($"已完成并保存！ {files.Count} file name(s) written.", "Finished and saved! {files.Count} file name(s) written."), Loc.T("文件名列表", "File Names"));
        }
        catch (Exception ex)
        {
            MessageBox.Show(Loc.T($"获取文件列表失败：{ex.Message}", "Failed to list files: {ex.Message}"), Loc.T("错误", "Error"));
        }
    }
}
