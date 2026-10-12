namespace ExcelTools.Common;

using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

/// <summary>
/// 公共工具函数。从原始 MyTools 改写，去掉公司特定逻辑。
/// </summary>
internal static class MyTools
{
    private static readonly char[] InvalidSheetChars = { ':', '\\', '/', '?', '*', '[', ']' };

    public static bool IsSheetNameValid(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return false;
        if (sheetName.Length > 31) return false;
        return !InvalidSheetChars.Any(sheetName.Contains);
    }

    public static bool IsSheetNameValid(List<string> names)
    {
        foreach (var n in names)
            if (!IsSheetNameValid(n)) return false;
        return true;
    }

    /// <summary>计算工作表真实行列数。overRows>0 时从该位置起校正空行。</summary>
    public static void GetRealRows(dynamic worksheet, int overRows, out int rowsCount, out int colsCount)
    {
        rowsCount = (int)worksheet.UsedRange.Rows.Count;
        colsCount = (int)worksheet.UsedRange.Columns.Count;

        if (overRows != 0 && rowsCount >= overRows)
        {
            for (int i = 2; i <= rowsCount; i++)
            {
                var s1 = Convert.ToString(worksheet.Cells[i, 1].Value);
                var s2 = Convert.ToString(worksheet.Cells[i, 2].Value);
                var s3 = Convert.ToString(worksheet.Cells[i, 4].Value);
                var s4 = Convert.ToString(worksheet.Cells[i, 6].Value);
                var s5 = Convert.ToString(worksheet.Cells[i, 8].Value);
                if (string.Concat(s1, s2, s3, s4, s5).Length == 0)
                    rowsCount = i - 1;
            }

            for (int i = 1; i <= colsCount; i++)
            {
                var c1 = Convert.ToString(worksheet.Cells[1, i].Value);
                var c2 = Convert.ToString(worksheet.Cells[2, i].Value);
                var c3 = Convert.ToString(worksheet.Cells[3, i].Value);
                var c4 = Convert.ToString(worksheet.Cells[4, i].Value);
                var c5 = Convert.ToString(worksheet.Cells[5, i].Value);
                if (string.Concat(c1, c2, c3, c4, c5).Length == 0)
                    colsCount = i - 1;
            }
        }
    }

    /// <summary>将常用列设为文本格式</summary>
    public static void FormatToText(dynamic worksheet)
    {
        string[] cols = { "A:A", "B:B", "C:C", "D:D", "E:E", "F:F", "G:G", "H:H", "J:J", "N:N", "Q:Q", "R:R", "T:T", "V:V", "W:W", "AH:AH", "AI:AI", "AJ:AJ", "AM:AM" };
        foreach (var c in cols)
        {
            try { worksheet.Range[c].NumberFormatLocal = "@"; } catch { }
        }
        worksheet.Columns.AutoFit();
    }

    /// <summary>将指定范围设为文本格式</summary>
    public static void FormatToText(dynamic worksheet, int mainRowsCount, int mainColsCount)
    {
        var rg = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[mainRowsCount, mainColsCount]];
        rg.NumberFormatLocal = "@";
    }

    /// <summary>清理工作表中 NULL 和 \r\n 特殊符号</summary>
    public static bool ClearSpecialSymbols(dynamic worksheet, out int rowsCount, out int colsCount)
    {
        rowsCount = 0;
        colsCount = 0;
        GetRealRows(worksheet, 1, out rowsCount, out colsCount);
        if (rowsCount < 1 || colsCount < 1) return false;

        var rg = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowsCount, colsCount]];
        rg.Replace("NULL", "");
        rg.Replace("\r\n", "");
        return true;
    }

    /// <summary>用 OleDb SQL 查询 Excel 并写入指定工作表</summary>
    public static void CopyToNewSheet(dynamic excelApp, string sql, dynamic worksheet, params string[] para)
    {
        var ds = new DataSet();
        var dt = new DataTable();
        string connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelApp.ActiveWorkbook.FullName};Extended Properties=Excel 8.0;";
        using (var conn = new OleDbConnection(connStr))
        {
            conn.Open();
            using var adapter = new OleDbDataAdapter(sql, conn);
            adapter.Fill(dt);
        }

        int rows = dt.Rows.Count;
        int cols = dt.Columns.Count;
        if (rows == 0) return;

        var dataArray = new object[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
                dataArray[i, j] = dt.Rows[i][j]?.ToString() ?? "";

        int startRow = para.Length > 0 ? 3 : 2;
        var target = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[rows + startRow - 1, cols]];
        target.Value = dataArray;

        if (para.Length > 0)
        {
            int r = (int)worksheet.UsedRange.Rows.Count;
            var lastRow = worksheet.Range[worksheet.Cells[r + 1, 1], worksheet.Cells[r + 1, 2]];
            lastRow.Merge();
            lastRow.Value = Loc.T("总计", "Total");
        }

        worksheet.Columns.AutoFit();
        worksheet.Rows.AutoFit();
    }

    /// <summary>拷贝原始表到名为 Original 的副本工作表</summary>
    public static void CopyToNewSheet(dynamic excelApp, string srcSheetName)
    {
        dynamic? srcWs = null;
        dynamic? targetWs = null;

        foreach (dynamic ws in excelApp.Worksheets)
        {
            if (string.Equals(ws.Name?.ToString().Trim(), srcSheetName, StringComparison.OrdinalIgnoreCase))
                srcWs = ws;
            else if (string.Equals(ws.Name?.ToString().Trim(), "Original", StringComparison.OrdinalIgnoreCase))
                targetWs = ws;
        }

        if (srcWs == null) return;

        if (targetWs == null)
        {
            targetWs = excelApp.Worksheets.Add();
            targetWs.Name = "Original";
        }
        else
        {
            targetWs.UsedRange.Clear();
        }

        int srcRows = (int)srcWs.UsedRange.Rows.Count;
        int srcCols = (int)srcWs.UsedRange.Columns.Count;
        var srcRange = srcWs.Range[srcWs.Cells[1, 1], srcWs.Cells[srcRows, srcCols]];
        var targetRange = targetWs.Range[targetWs.Cells[1, 1], targetWs.Cells[srcRows, srcCols]];
        srcRange.Copy(targetRange);
        srcWs.Columns.AutoFit();
        targetWs.Columns.AutoFit();
        srcWs.Move(targetWs);
        srcWs.Activate();
    }

    /// <summary>获取某列所有去重值（按指定字母列）</summary>
    public static List<string>? GetByField(dynamic excelApp, string baseField)
    {
        var list = new List<string>();
        GetRealRows(excelApp.ActiveSheet, 0, out int rows, out int _cols);
        string col = baseField.ToUpper();

        for (int i = 2; i <= rows; i++)
        {
            var val = Convert.ToString(excelApp.Cells[i, col].Value);
            if (!string.IsNullOrEmpty(val) && !list.Contains(val))
                list.Add(val);
        }

        if (list.Count == 0) return null;
        list.Sort();
        list.Reverse();
        return list;
    }

    /// <summary>获取某列所有去重值（按列号）</summary>
    public static List<string>? GetByFieldCol(dynamic excelApp, int colIndex)
    {
        var list = new List<string>();
        GetRealRows(excelApp.ActiveSheet, 0, out int rows, out int _cols);

        for (int i = 2; i <= rows; i++)
        {
            var val = Convert.ToString(excelApp.Cells[i, colIndex].Value);
            if (!string.IsNullOrEmpty(val) && !list.Contains(val))
                list.Add(val);
        }

        if (list.Count == 0) return null;
        list.Sort();
        list.Reverse();
        return list;
    }
}