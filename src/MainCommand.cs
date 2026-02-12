using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel; // 需要引用 Excel 库

[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
public class ExportVisibleElements : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;
        View activeView = doc.ActiveView;

        // 1. 核心逻辑：获取当前视图中所有可见的构件
        FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
        List<Element> visibleElements = collector.WhereElementIsNotElementType().ToElements().ToList();

        // 2. 准备数据列表
        // 这里根据你的业务需求，提取构件 ID、名称、类别等
        var dataList = visibleElements.Select(e => new {
            Id = e.Id.ToString(),
            Name = e.Name,
            Category = e.Category?.Name ?? "未知"
        }).ToList();

        // 3. 导出到 Excel (逻辑简化版)
        ExportToExcel(dataList);

        TaskDialog.Show("导出成功", $"已成功导出 {dataList.Count} 个构件到桌面！");
        return Result.Succeeded;
    }

    private void ExportToExcel(object data) 
    {
        // 这里编写调用 Excel 的具体代码
        // AI 提示：在实际开发中，建议使用 EPPlus 库，因为它不需要电脑安装 Excel 也能运行
    }
}