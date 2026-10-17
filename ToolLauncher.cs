namespace ExcelTools;

using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelTools.Common;
using ExcelTools.Forms;

/// <summary>
/// 统一的工具启动入口：Ribbon 按钮、WPS 工具栏按钮、ExcelCommand 菜单
/// 三条 UI 路径都走这里，保证异常处理、日志、本地化完全一致。
/// </summary>
public static class ToolLauncher
{
    /// <summary>Wrap a tool launch: catch, log and show the full exception chain.</summary>
    public static void Run(string toolName, Func<Form?> factory)
    {
        try
        {
            using Form? form = factory();
            form?.ShowDialog();
        }
        catch (Exception ex)
        {
            string detail = Flatten(ex);
            WriteLog(toolName, detail);
            MessageBox.Show(
                Loc.T($"[{toolName}] 发生错误\n\n{detail}",
                      $"[{toolName}] Error\n\n{detail}"),
                Loc.T("Excel 工具 - 错误详情", "Excel Tools - Error Details"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string Flatten(Exception ex)
    {
        var builder = new StringBuilder();
        int depth = 0;
        for (Exception? current = ex; current is not null; current = current.InnerException, depth++)
        {
            string pad = new string(' ', depth * 2);
            builder.Append(pad).AppendLine($"{current.GetType().Name}: {current.Message}");
            if (!string.IsNullOrEmpty(current.StackTrace))
            {
                string top = current.StackTrace!.Split('\n')[0].Trim();
                builder.Append(pad).AppendLine("  at " + top);
            }
        }
        return builder.ToString();
    }

    private static void WriteLog(string toolName, string detail)
    {
        try
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "ExcelTools_ErrorLog.txt");
            File.AppendAllText(
                path,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {toolName}\n{detail}" +
                new string('-', 60) + "\n",
                Encoding.UTF8);
        }
        catch
        {
            // Logging is best-effort and must never mask the original error.
        }
    }
}
