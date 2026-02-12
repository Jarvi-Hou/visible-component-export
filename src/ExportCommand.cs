using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using ClosedXML.Excel; // 建议使用的库，操作简单

public void ExportToExcel(List<Element> elements)
{
    // 1. 定义保存路径（直接指向桌面）
    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    string fileName = "BIM构件导出_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
    string fullPath = Path.Combine(desktopPath, fileName);

    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("可见构件列表");

        // 2. 设置表头
        worksheet.Cell(1, 1).Value = "构件ID";
        worksheet.Cell(1, 2).Value = "类别";
        worksheet.Cell(1, 3).Value = "族与类型";
        worksheet.Cell(1, 4).Value = "关联标高";

        // 3. 填充数据
        int row = 2;
        foreach (Element el in elements)
        {
            worksheet.Cell(row, 1).Value = el.Id.ToString();
            worksheet.Cell(row, 2).Value = el.Category?.Name ?? "通用";
            worksheet.Cell(row, 3).Value = el.Name;
            
            // 获取标高参数（BIM 业务核心：通常通过 Level 参数获取）
            Parameter levelParam = el.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
            worksheet.Cell(row, 4).Value = levelParam?.AsValueString() ?? "无";
            
            row++;
        }

        // 4. 美化表格：自动调整列宽
        worksheet.Columns().AdjustToContents();

        // 5. 保存并提示
        workbook.SaveAs(fullPath);
        MessageBox.Show($"导出成功！文件已保存在桌面：\n{fileName}", "提示");
    }
}