namespace ExcelTools;

using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelTools.Forms;

/// <summary>
/// Ribbon 提供者。
///
/// ★关键：Excel-DNA 只有在程序集里存在一个继承 <see cref="ExcelRibbon"/> 的 public 类时，
/// 才会把这个类作为 COM 加载项注册进 Excel，并由它向 Excel 提供功能区 XML。
/// 只在 .dna 文件里写 &lt;CustomUI&gt; 而没有这个派生类，那段 XML 不会被任何东西读取，
/// Excel 中就不会出现自定义选项卡（且不会报任何错）。
///
/// 另外，本类的 onAction 是 public **实例**方法（COM 回调要求实例方法），
/// 不是 public static 方法，因此不会被 Excel-DNA 当成 UDF 导出，不会污染函数名称空间。
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='ExcelToolsTab' label='Excel 工具'>
        <group id='grpData' label='数据处理'>
          <button id='btnSeperate' label='拆分工作表' size='large' imageMso='SplitTable'
                  onAction='OnSeperateClick' screentip='按某列值将总表拆成多个工作表'/>
          <button id='btnFilter' label='条件筛选' size='large' imageMso='Filter'
                  onAction='OnFilterClick' screentip='按多条件筛选数据到新表'/>
          <button id='btnSort' label='按列排序' size='large' imageMso='SortUp'
                  onAction='OnSortClick' screentip='按指定列排序并保留原始数据副本'/>
          <button id='btnDelete' label='删除匹配行' size='large' imageMso='Delete'
                  onAction='OnDeleteClick' screentip='删除与另一表匹配的行'/>
        </group>
        <group id='grpCompare' label='比对'>
          <button id='btnCompare' label='两表比对' size='large' imageMso='CompareAndCombine'
                  onAction='OnCompareClick' screentip='比较两个工作表的数据差异'/>
          <button id='btnTimeCompare' label='差异比对' size='large' imageMso='FindDialogExcel'
                  onAction='OnTimeCompareClick' screentip='按匹配列联表找比对列不一致的行'/>
          <button id='btnCountDistinct' label='累计去重计数' size='large' imageMso='RemoveDuplicates'
                  onAction='OnCountDistinctClick' screentip='截至每行统计某列累计不重复值数'/>
        </group>
        <group id='grpStat' label='统计'>
          <button id='btnGroupStat' label='分组统计' size='large' imageMso='PivotTableInsert'
                  onAction='OnGroupStatClick' screentip='按分类字段生成交叉汇总表'/>
          <button id='btnPermMatrix' label='权限矩阵' size='large' imageMso='CreateTable'
                  onAction='OnPermMatrixClick' screentip='从数据生成角色×功能权限矩阵'/>
          <button id='btnGetFileNames' label='文件名列表' size='large' imageMso='FileNew'
                  onAction='OnGetFileNamesClick' screentip='递归提取目录下所有文件名到当前表'/>
        </group>
        <group id='grpMatch' label='匹配'>
          <button id='btnKeyMatch' label='多字段匹配' size='large' imageMso='LookUp'
                  onAction='OnKeyMatchClick' screentip='按多字段联合匹配两表提取Key值'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public override string GetCustomUI(string ribbonId) => RibbonXml;

    public void OnSeperateClick(IRibbonControl control)      => Run("拆分工作表",   () => new FrmSeperate());
    public void OnFilterClick(IRibbonControl control)        => Run("条件筛选",     () => new FrmFilter());
    public void OnSortClick(IRibbonControl control)          => Run("按列排序",     () => new FrmSort());
    public void OnDeleteClick(IRibbonControl control)        => Run("删除匹配行",   () => new FrmDelete());
    public void OnCompareClick(IRibbonControl control)       => Run("两表比对",     () => new FrmCompare());
    public void OnCountDistinctClick(IRibbonControl control) => Run("累计去重计数", () => new FrmCountDistinct());
    public void OnGroupStatClick(IRibbonControl control)     => Run("分组统计",     () => new FrmGroupStat());
    public void OnTimeCompareClick(IRibbonControl control)   => Run("差异比对",     () => new FrmTimeCompare());
    public void OnKeyMatchClick(IRibbonControl control)      => Run("多字段匹配",   () => new FrmKeyMatch());
    public void OnPermMatrixClick(IRibbonControl control)    => Run("权限矩阵",     () => new FrmPermMatrix());
    public void OnGetFileNamesClick(IRibbonControl control)  => Run("文件名列表",   () => { FrmGetFileNames.Run(); return null; });

    /// <summary>统一执行包装：捕获异常、展开 InnerException 链、落盘日志、弹窗提示。</summary>
    private static void Run(string toolName, Func<Form?> factory)
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
                "ExcelTools_错误Log.txt");
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
