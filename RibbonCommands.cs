namespace ExcelTools;

using ExcelTools.Forms;

/// <summary>
/// Ribbon 按钮事件入口。Excel-DNA 通过方法名匹配 onAction。
/// </summary>
public static class RibbonCommands
{
    public static void OnRibbonLoad()
    {
        // Ribbon 加载时无需额外初始化
    }

    public static void OnSeperateClick()
    {
        using var frm = new FrmSeperate();
        frm.ShowDialog();
    }

    public static void OnFilterClick()
    {
        using var frm = new FrmFilter();
        frm.ShowDialog();
    }

    public static void OnSortClick()
    {
        using var frm = new FrmSort();
        frm.ShowDialog();
    }

    public static void OnDeleteClick()
    {
        using var frm = new FrmDelete();
        frm.ShowDialog();
    }

    public static void OnCompareClick()
    {
        using var frm = new FrmCompare();
        frm.ShowDialog();
    }

    public static void OnCountDistinctClick()
    {
        using var frm = new FrmCountDistinct();
        frm.ShowDialog();
    }

    public static void OnGroupStatClick()
    {
        using var frm = new FrmGroupStat();
        frm.ShowDialog();
    }

    public static void OnGetFileNamesClick()
    {
        FrmGetFileNames.Run();
    }

    public static void OnTimeCompareClick()
    {
        using var frm = new FrmTimeCompare();
        frm.ShowDialog();
    }

    public static void OnKeyMatchClick()
    {
        using var frm = new FrmKeyMatch();
        frm.ShowDialog();
    }

    public static void OnPermMatrixClick()
    {
        using var frm = new FrmPermMatrix();
        frm.ShowDialog();
    }
}