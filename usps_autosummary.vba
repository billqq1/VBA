Sub Auto_Open()
    ' 定义源工作表和目标工作表
    Dim 对手GA报价 As Worksheet
    Dim 对手PM报价 As Worksheet
    Dim 底价_对手GA差价百分比表 As Worksheet
    Dim 底价_对手PM差价百分比表 As Worksheet
    Dim 客户GA报价 As Worksheet
    Dim 客户PM报价 As Worksheet
    
    
    ' 设置源工作表和目标工作表
    Set 对手GA报价 = ThisWorkbook.Sheets("对手GA报价")
    Set 对手PM报价 = ThisWorkbook.Sheets("对手PM报价")
    Set 底价_对手GA差价百分比表 = ThisWorkbook.Sheets("底价_对手GA差价百分比表")
    Set 底价_对手PM差价百分比表 = ThisWorkbook.Sheets("底价_对手PM差价百分比表")
    Set 客户GA报价 = ThisWorkbook.Sheets("客户GA报价")
    Set 客户PM报价 = ThisWorkbook.Sheets("客户PM报价")
    
    
    ' 复制源工作表的A1单元格到目标工作表的B1单元格
    '对手GA报价.Range("A1").Copy Destination:=底价_对手GA差价百分比表.Range("B1")
    
    ' 复制源工作表的A1到J93范围到目标工作表的A1单元格
    对手GA报价.Range("A1:J93").Copy Destination:=底价_对手GA差价百分比表.Range("A1")
    对手PM报价.Range("A1:J73").Copy Destination:=底价_对手PM差价百分比表.Range("A1")
    
    '生成Title
    
    底价_对手GA差价百分比表.Range("A1").Value = "底价-对手GA差价百分比表"
    底价_对手PM差价百分比表.Range("A1").Value = "底价-对手PM差价百分比表"
    
    
    ' 清除内容
    底价_对手GA差价百分比表.Range("B4:J20").Value = ""
    底价_对手GA差价百分比表.Range("B23:J93").Value = ""
    底价_对手PM差价百分比表.Range("B4:J73").Value = ""
    
    
    ' 设置公式
    底价_对手GA差价百分比表.Range("B4:J20").Formula = "=(对手GA报价!B4-底价GA报价!B4)/底价GA报价!B4"
    ' 设置单元格格式为百分比
    底价_对手GA差价百分比表.Range("B4:J20").NumberFormat = "0.00%"
    '底价_对手GA差价百分比表.Range("B4:J20").Cells(1, 1).AutoFill Destination:=底价_对手GA差价百分比表.Range("B4:J20")
    底价_对手GA差价百分比表.Range("B23:J93").Formula = "=(对手GA报价!B23-底价GA报价!B23)/底价GA报价!B23"
    ' 设置单元格格式为百分比
    底价_对手GA差价百分比表.Range("B23:J93").NumberFormat = "0.00%"
    
    底价_对手PM差价百分比表.Range("B4:J73").Formula = "=(对手PM报价!B4-底价PM报价!B4)/底价PM报价!B4"
    底价_对手PM差价百分比表.Range("B4:J73").NumberFormat = "0.00%"
    
    '自动生成当前日期
    底价_对手GA差价百分比表.Range("I2").Value = "报表生成日期:" & Date
    底价_对手PM差价百分比表.Range("I2").Value = "报表生成日期:" & Date
    客户GA报价.Range("I2").Value = "报价生成日期:" & Date
    客户PM报价.Range("I2").Value = "报价生成日期:" & Date
    
    
    ' 清除剪贴板中的内容
    Application.CutCopyMode = False
    Application.WindowState = xlMaximized
End Sub
