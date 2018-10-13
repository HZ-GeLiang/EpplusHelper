# EpplusHelper


```
string tempPath = @"c:a.xlsx";
using (MemoryStream ms = new MemoryStream())
using (FileStream fs = System.IO.File.OpenRead(tempPath))
using (ExcelPackage excelPackage = new ExcelPackage(fs))
{
    var config = EpplusHelper.GetEmptyConfig();
    var configSource = EpplusHelper.GetEmptyConfigSource();
    EpplusHelper.SetDefaultConfigFromExcel(excelPackage, config);
    EpplusHelper.SetConfigSourceHead(configSource, dtHead, head.Rows[0]);
    configSource.SheetBody[1] =dtbody
    configSource.SheetBodySummary[1] = new Dictionary<object, object>(){....}
    configSource.SheetBody[2] = dtBody2;
    configSource.SheetBodySummary[1] = new Dictionary<object, object>(){....}
    EpplusHelper.SetConfigSourceFoot(configSource, dtFootResult, dtFootResult.Rows[0]);
    EpplusHelper.FillData(excelPackage, config, configSource, "导出测试", 1);
    EpplusHelper.DeleteWorksheet(....);//删除母版页
    excelPackage.SaveAs(ms); // 导入数据到流中 
    ms.Position = 0;
    ms.Save(@"C:\1.xlsx");
}
```
