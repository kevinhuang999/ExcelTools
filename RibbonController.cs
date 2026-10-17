namespace ExcelTools;

using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelTools.Common;
using ExcelTools.Forms;

/// <summary>
/// Ribbon provider.
///
/// Excel-DNA only registers a class as a COM add-in - and therefore only asks
/// it for the ribbon XML - when a class deriving from <see cref="ExcelRibbon"/>
/// exists. Writing CustomUI markup in the .dna file alone is not enough.
///
/// The markup is rebuilt on every request so that all labels and screen tips
/// follow the display language of the Excel client (see <see cref="Loc"/>).
/// The handlers are public instance methods, so they are never exported as
/// worksheet functions and cannot collide with anything.
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) => BuildRibbonXml();

    private static string BuildRibbonXml()
    {
        string tab = Loc.T("Excel 工具", "Excel Tools");

        string groupData = Loc.T("数据处理", "Data");
        string groupCompare = Loc.T("比对", "Compare");
        string groupStat = Loc.T("统计", "Statistics");
        string groupMatch = Loc.T("匹配", "Match");

        return $@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='ExcelToolsTab' label='{tab}'>
        <group id='grpData' label='{groupData}'>
          <button id='btnSeperate' label='{Loc.T("拆分工作表", "Seperate")}' size='large'
                  imageMso='SplitTable' onAction='OnSeperateClick'
                  screentip='{Loc.T("按某列的值，把当前工作表拆成多个工作表", "Split the sheet into one sheet per distinct value of a column")}'/>
          <button id='btnFilter' label='{Loc.T("条件筛选", "Filter")}' size='large'
                  imageMso='Filter' onAction='OnFilterClick'
                  screentip='{Loc.T("按条件筛选行并复制到新工作表，支持逐行匹配与 SQL 两种方式", "Filter rows into a new sheet, by conditions or by SQL")}'/>
          <button id='btnSort' label='{Loc.T("按列排序", "Sort")}' size='large'
                  imageMso='SortUp' onAction='OnSortClick'
                  screentip='{Loc.T("按指定列排序，并保留原始数据副本", "Sort by a column and keep a backup of the original data")}'/>
          <button id='btnDelete' label='{Loc.T("删除匹配行", "Delete")}' size='large'
                  imageMso='Delete' onAction='OnDeleteClick'
                  screentip='{Loc.T("删除键值在另一张表中也存在的行", "Delete rows whose key also exists in another sheet")}'/>
        </group>
        <group id='grpCompare' label='{groupCompare}'>
          <button id='btnCompare' label='{Loc.T("两表比对", "Compare Files")}' size='large'
                  imageMso='CompareAndCombine' onAction='OnCompareClick'
                  screentip='{Loc.T("逐格比对两个 Excel 文件，列出全部差异", "Compare two Excel files cell by cell and list every difference")}'/>
          <button id='btnTimeCompare' label='{Loc.T("差异比对", "Find Differences")}' size='large'
                  imageMso='FindDialogExcel' onAction='OnTimeCompareClick'
                  screentip='{Loc.T("按匹配列关联两张表，找出比对列不一致的行", "Match two sheets on a key column and list the rows that differ")}'/>
          <button id='btnCountDistinct' label='{Loc.T("累计去重计数", "Count Distinct")}' size='large'
                  imageMso='RemoveDuplicates' onAction='OnCountDistinctClick'
                  screentip='{Loc.T("逐行统计截至当前行出现过多少个不重复值", "For each row, count the distinct values seen so far")}'/>
        </group>
        <group id='grpStat' label='{groupStat}'>
          <button id='btnGroupStat' label='{Loc.T("分组统计", "Statistics")}' size='large'
                  imageMso='PivotTableInsert' onAction='OnGroupStatClick'
                  screentip='{Loc.T("按分类与期间生成交叉汇总表", "Build a category by period cross table")}'/>
          <button id='btnPermMatrix' label='{Loc.T("权限矩阵", "Role Menus")}' size='large'
                  imageMso='CreateTable' onAction='OnPermMatrixClick'
                  screentip='{Loc.T("生成角色与功能的权限矩阵", "Build a role by function permission matrix")}'/>
          <button id='btnGetFileNames' label='{Loc.T("文件名列表", "File Names")}' size='large'
                  imageMso='FileNew' onAction='OnGetFileNamesClick'
                  screentip='{Loc.T("递归提取文件夹下所有文件名到当前表", "List the file names under a folder into the sheet")}'/>
        </group>
        <group id='grpMatch' label='{groupMatch}'>
          <button id='btnKeyMatch' label='{Loc.T("多字段匹配", "Get Key")}' size='large'
                  imageMso='LookUp' onAction='OnKeyMatchClick'
                  screentip='{Loc.T("按最多三个字段联合匹配两表，回填 Key 值", "Match two tables on up to three fields and fill in the key")}'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnSeperateClick(IRibbonControl control)      => Run(Loc.T("拆分工作表", "Seperate"), () => new FrmSeperate());
    public void OnFilterClick(IRibbonControl control)        => Run(Loc.T("条件筛选", "Filter"), () => new FrmFilter());
    public void OnSortClick(IRibbonControl control)          => Run(Loc.T("按列排序", "Sort"), () => new FrmSort());
    public void OnDeleteClick(IRibbonControl control)        => Run(Loc.T("删除匹配行", "Delete"), () => new FrmDelete());
    public void OnCompareClick(IRibbonControl control)       => Run(Loc.T("两表比对", "Compare Files"), () => new FrmCompare());
    public void OnCountDistinctClick(IRibbonControl control) => Run(Loc.T("累计去重计数", "Count Distinct"), () => new FrmCountDistinct());
    public void OnGroupStatClick(IRibbonControl control)     => Run(Loc.T("分组统计", "Statistics"), () => new FrmGroupStat());
    public void OnTimeCompareClick(IRibbonControl control)   => Run(Loc.T("差异比对", "Find Differences"), () => new FrmTimeCompare());
    public void OnKeyMatchClick(IRibbonControl control)      => Run(Loc.T("多字段匹配", "Get Key"), () => new FrmKeyMatch());
    public void OnPermMatrixClick(IRibbonControl control)    => Run(Loc.T("权限矩阵", "Role Menus"), () => new FrmPermMatrix());
    public void OnGetFileNamesClick(IRibbonControl control)  => Run(Loc.T("文件名列表", "File Names"), () => { FrmGetFileNames.Run(); return null; });

    /// <summary>Wrap a tool launch: catch, log and show the full exception chain.</summary>
    internal static void Run(string toolName, Func<Form?> factory)
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

    /// <summary>Flatten the exception chain, outermost wrapper first.</summary>
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
