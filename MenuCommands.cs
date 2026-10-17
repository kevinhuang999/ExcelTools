#if WPS_BUILD
namespace ExcelTools;

using ExcelDna.Integration;

/// <summary>
/// 菜单命令（WPS 兼容路径）。
///
/// Excel-DNA 的 Ribbon 注册在 WPS 里走不通（COMAddIns 集合恒为空），
/// 所以 WPS 下改用经典 CommandBars 工具栏：由 WpsToolbarAddIn 在
/// AutoOpen 里创建工具栏，按钮的 OnAction 指向这里注册的同名 XLM 命令。
///
/// 命令名必须用纯中文/纯 ASCII：含空格或全宽标点（如「｜」）会导致
/// xlfRegister 失败，把微软 Excel 侧也一起搞坏。
/// </summary>
public static class MenuCommands
{
    [ExcelCommand(Name = "拆分工作表")]
    public static void Seperate() =>
        ToolLauncher.Run("拆分工作表", () => new Forms.FrmSeperate());

    [ExcelCommand(Name = "条件筛选")]
    public static void Filter() =>
        ToolLauncher.Run("条件筛选", () => new Forms.FrmFilter());

    [ExcelCommand(Name = "按列排序")]
    public static void Sort() =>
        ToolLauncher.Run("按列排序", () => new Forms.FrmSort());

    [ExcelCommand(Name = "删除匹配行")]
    public static void Delete() =>
        ToolLauncher.Run("删除匹配行", () => new Forms.FrmDelete());

    [ExcelCommand(Name = "两表比对")]
    public static void Compare() =>
        ToolLauncher.Run("两表比对", () => new Forms.FrmCompare());

    [ExcelCommand(Name = "累计去重计数")]
    public static void CountDistinct() =>
        ToolLauncher.Run("累计去重计数", () => new Forms.FrmCountDistinct());

    [ExcelCommand(Name = "分组统计")]
    public static void GroupStat() =>
        ToolLauncher.Run("分组统计", () => new Forms.FrmGroupStat());

    [ExcelCommand(Name = "差异比对")]
    public static void TimeCompare() =>
        ToolLauncher.Run("差异比对", () => new Forms.FrmTimeCompare());

    [ExcelCommand(Name = "多字段匹配")]
    public static void KeyMatch() =>
        ToolLauncher.Run("多字段匹配", () => new Forms.FrmKeyMatch());

    [ExcelCommand(Name = "权限矩阵")]
    public static void PermMatrix() =>
        ToolLauncher.Run("权限矩阵", () => new Forms.FrmPermMatrix());

    [ExcelCommand(Name = "文件名列表")]
    public static void GetFileNames() =>
        ToolLauncher.Run("文件名列表", () => { Forms.FrmGetFileNames.Run(); return null; });
}
#endif
