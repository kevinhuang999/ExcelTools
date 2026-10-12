namespace ExcelTools;

using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelTools.Forms;

/// <summary>
/// Ribbon button event entry points. Excel-DNA matches methods by name via onAction.
/// 每个按钮统一走 Run() 包装：捕获完整异常链 → 弹详情框 → 落盘日志，
/// 避免 TargetInvocationException 外层包装掩盖真正根因。
/// </summary>
public static class RibbonCommands
{
    public static void OnRibbonLoad()
    {
        // Ribbon 加载时无需额外初始化
    }

    public static void OnSeperateClick()      => Run("拆分工作表",   () => new FrmSeperate());
    public static void OnFilterClick()        => Run("条件筛选",     () => new FrmFilter());
    public static void OnSortClick()          => Run("按列排序",     () => new FrmSort());
    public static void OnDeleteClick()        => Run("删除匹配行",   () => new FrmDelete());
    public static void OnCompareClick()       => Run("两表比对",     () => new FrmCompare());
    public static void OnCountDistinctClick() => Run("累计去重计数", () => new FrmCountDistinct());
    public static void OnGroupStatClick()     => Run("分组统计",     () => new FrmGroupStat());
    public static void OnTimeCompareClick()   => Run("差异比对",     () => new FrmTimeCompare());
    public static void OnKeyMatchClick()      => Run("多字段匹配",   () => new FrmKeyMatch());
    public static void OnPermMatrixClick()    => Run("权限矩阵",     () => new FrmPermMatrix());
    public static void OnGetFileNamesClick()  => Run("文件名列表",   () => { FrmGetFileNames.Run(); return null; });

    /// <summary>统一执行包装：捕获异常、展开 InnerException 链、落盘日志、弹窗提示。</summary>
    private static void Run(string toolName, Func<Form> factory)
    {
        try
        {
            using Form? frm = factory();
            frm?.ShowDialog();
        }
        catch (Exception ex)
        {
            string detail = Flatten(ex);
            WriteLog(toolName, detail);
            MessageBox.Show(
                $"[{toolName}] 发生错误\n\n{detail}",
                "Excel 工具 - 错误详情",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    /// <summary>把异常链逐层展开成可读文本（外层包装 → 内层真凶）。</summary>
    private static string Flatten(Exception ex)
    {
        var sb = new StringBuilder();
        int depth = 0;
        for (Exception? cur = ex; cur is not null; cur = cur.InnerException, depth++)
        {
            string pad = new string(' ', depth * 2);
            sb.Append(pad);
            sb.AppendLine($"{cur.GetType().Name}: {cur.Message}");
            if (!string.IsNullOrEmpty(cur.StackTrace))
            {
                // 只取最内层堆栈首行，避免刷屏
                string top = cur.StackTrace!.Split('\n')[0].Trim();
                sb.Append(pad).AppendLine("  at " + top);
            }
        }
        return sb.ToString();
    }

    /// <summary>错误详情落盘，便于脱离 Excel 环境排查。</summary>
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
            // 日志失败不影响主流程
        }
    }
}
