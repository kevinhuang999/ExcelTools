#if WPS_BUILD
namespace ExcelTools;

using System;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;

/// <summary>
/// WPS 兼容层：在 WPS 里自建工具栏（CommandBar），绕开 Ribbon / COMAddIns
/// 注册路径（WPS 的 COMAddIns 集合恒为空，Excel-DNA 的 Ribbon 注册必然失败
/// 并导致宿主异常退出）。
///
/// 加载顺序（已从 Excel-DNA 源码确认）：
///   DnaLibrary.AutoOpen 先遍历所有 IExcelAddIn 调 AutoOpen()，
///   然后才调 LoadCustomUI()（Ribbon 注册在那里）。
///   所以本类的 AutoOpen 一定在 Ribbon 报错之前执行。
///
/// 按钮点击：OnAction 指向 MenuCommands 里注册的同名 XLM 命令。
/// 美观：按 Ribbon 的分组拆成多个 CommandBar，WPS 会把它们渲染进
///      「加载项」选项卡的不同组；按钮带内置图标与悬停说明。
/// </summary>
public class WpsToolbarAddIn : IExcelAddIn
{
    private static readonly string LogFile =
        Path.Combine(Path.GetTempPath(), "ExcelTools_wps.log");

    private readonly List<object> _bars = new();

    private static void Log(string msg)
    {
        try
        {
            File.AppendAllText(LogFile,
                DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg + Environment.NewLine);
        }
        catch { }
    }

    // 分组定义：组名 → 该组的按钮（标题, 悬停说明, FaceId）
    private static readonly (string Group, string Caption, string Tip, int FaceId)[] Buttons =
    {
        ("数据处理", "拆分工作表",   "按某列的值，把当前工作表拆成多个工作表", 159),
        ("数据处理", "条件筛选",     "按条件筛选行并复制到新工作表",           158),
        ("数据处理", "按列排序",     "按指定列排序，并保留原始数据副本",       197),
        ("数据处理", "删除匹配行",   "删除键值在另一张表中也存在的行",         348),

        ("比对",     "两表比对",     "按关键字比对两张表，标出差异",           474),
        ("比对",     "差异比对",     "逐单元格比对两个工作表，高亮不同处",     215),
        ("比对",     "累计去重计数", "按关键字累计计数并去重",                 487),

        ("统计",     "分组统计",     "按关键字分组并统计各字段",               187),
        ("统计",     "权限矩阵",     "生成角色与菜单的权限矩阵表",             172),
        ("统计",     "文件名列表",   "把指定文件夹的文件名导出成工作表",       165),

        ("匹配",     "多字段匹配",   "用多个字段组合进行跨表匹配",             969),
    };

    public void AutoOpen()
    {
        Log("=== AutoOpen ===");
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            string ver;
            try { ver = Convert.ToString(app.Version); } catch { ver = ""; }
            string name;
            try { name = (string)app.Name; } catch { name = ""; }
            Log("Name=" + name + " Version=" + ver);

            bool isWps =
                name.Contains("WPS", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Kingsoft", StringComparison.OrdinalIgnoreCase) ||
                ver.StartsWith("12.", StringComparison.Ordinal) ||
                ver.StartsWith("11.", StringComparison.Ordinal);

            Log("isWps=" + isWps);
            if (!isWps)
            {
                Log("微软 Excel，不建工具栏");
                return;
            }

            BuildToolbars(app);
        }
        catch (Exception ex)
        {
            Log("AutoOpen 异常: " + ex);
        }
    }

    public void AutoClose()
    {
        foreach (var bar in _bars)
        {
            try { ((dynamic)bar).Delete(); } catch { }
        }
        _bars.Clear();
        Log("工具栏已全部删除");
    }

    private void BuildToolbars(dynamic app)
    {
        dynamic cbs = app.CommandBars;
        Log("CommandBars.Count=" + cbs.Count);

        // 按组名保持出现顺序
        var groups = new List<string>();
        foreach (var b in Buttons)
            if (!groups.Contains(b.Group))
                groups.Add(b.Group);

        foreach (string group in groups)
        {
            try
            {
                // 清同名旧栏
                for (int i = cbs.Count; i >= 1; i--)
                {
                    try
                    {
                        dynamic old = cbs.Item(i);
                        if ((string)old.Name == group)
                        {
                            old.Delete();
                            break;
                        }
                    }
                    catch { }
                }

                dynamic bar = cbs.Add(group, 0, null, true);
                bar.Visible = true;
                _bars.Add(bar);
                Log("工具栏已建: " + group);

                foreach (var b in Buttons)
                {
                    if (b.Group != group)
                        continue;
                    AddButton(bar, b.Caption, b.Tip, b.FaceId);
                }
            }
            catch (Exception ex)
            {
                Log("建工具栏 '" + group + "' 异常: " + ex.Message);
            }
        }

        Log("共 " + _bars.Count + " 个工具栏");
    }

    private void AddButton(dynamic bar, string label, string tip, int faceId)
    {
        try
        {
            dynamic btn = bar.Controls.Add(1);   // msoControlButton
            btn.Caption = label;
            btn.Style = 2;                       // msoButtonIconAndCaption：图标+文字
            btn.FaceId = faceId;                 // Office 内置图标
            btn.Visible = true;
            btn.Enabled = true;
            btn.TooltipText = tip;               // 悬停说明
            btn.OnAction = label;                // → 同名 XLM 命令（MenuCommands）
            Log("按钮: " + label + " (FaceId=" + faceId + ")");
        }
        catch (Exception ex)
        {
            Log("加按钮 '" + label + "' 异常: " + ex.Message);
        }
    }
}

#endif
