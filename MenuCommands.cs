namespace ExcelTools;

using ExcelDna.Integration;
using ExcelTools.Common;
using ExcelTools.Forms;

/// <summary>
/// 菜单命令入口（WPS 专用路径）。
///
/// 为什么需要它：微软 Excel 的功能区（Ribbon）走 COMAddIns 注册，在 WPS 里
/// 实测不可用——WPS 12.1 的 Application.COMAddIns 集合恒为 0（Count=0，
/// Item() 抛 E_FAIL），原版 Excel-DNA 的 Ribbon 加载必然失败。
/// 而 WPS 的 CommandBars 对象模型完全正常（186 个菜单，可增删），
/// Excel-DNA 的 [ExcelCommand] 正是通过 CommandBars 注册的，不经过 COMAddIns。
///
/// 效果：微软 Excel 用功能区选项卡；WPS 里功能区注册失败后，工具改从菜单进入，
/// 两条路互不影响，同一套窗体代码。
/// </summary>
public static class MenuCommands
{
    [ExcelCommand(Name = "拆分工作表")]
    public static void Seperate() =>
        RibbonController.Run(Loc.T("拆分工作表", "Seperate"), () => new FrmSeperate());

    [ExcelCommand(Name = "条件筛选")]
    public static void Filter() =>
        RibbonController.Run(Loc.T("条件筛选", "Filter"), () => new FrmFilter());

    [ExcelCommand(Name = "按列排序")]
    public static void Sort() =>
        RibbonController.Run(Loc.T("按列排序", "Sort"), () => new FrmSort());

    [ExcelCommand(Name = "删除匹配行")]
    public static void Delete() =>
        RibbonController.Run(Loc.T("删除匹配行", "Delete"), () => new FrmDelete());

    [ExcelCommand(Name = "两表比对")]
    public static void Compare() =>
        RibbonController.Run(Loc.T("两表比对", "Compare Files"), () => new FrmCompare());

    [ExcelCommand(Name = "累计去重计数")]
    public static void CountDistinct() =>
        RibbonController.Run(Loc.T("累计去重计数", "Count Distinct"), () => new FrmCountDistinct());

    [ExcelCommand(Name = "分组统计")]
    public static void GroupStat() =>
        RibbonController.Run(Loc.T("分组统计", "Statistics"), () => new FrmGroupStat());

    [ExcelCommand(Name = "差异比对")]
    public static void TimeCompare() =>
        RibbonController.Run(Loc.T("差异比对", "Find Differences"), () => new FrmTimeCompare());

    [ExcelCommand(Name = "多字段匹配")]
    public static void KeyMatch() =>
        RibbonController.Run(Loc.T("多字段匹配", "Get Key"), () => new FrmKeyMatch());

    [ExcelCommand(Name = "权限矩阵")]
    public static void PermMatrix() =>
        RibbonController.Run(Loc.T("权限矩阵", "Role Menus"), () => new FrmPermMatrix());

    [ExcelCommand(Name = "文件名列表")]
    public static void GetFileNames() =>
        RibbonController.Run(Loc.T("文件名列表", "File Names"), () => { FrmGetFileNames.Run(); return null; });
}
