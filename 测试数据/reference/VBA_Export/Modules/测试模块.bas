Attribute VB_Name = "测试模块"
Option Explicit



Sub test_log()
    Call glo_log.clear_log
    Call glo_log.create_log_title("模块", "工作簿", "工作表", "错误类型", "行位置", "行名称", "列位置", "列名称", "批注", "错误区域", "校验描述", "提取值", "列3", "列4", "列5", "说明")
    
    Call glo_log.export_log_to_ws(ThisWorkbook.Worksheets("输出日记"), "a")
End Sub




Sub test()
    Dim arr As Variant
    arr = Split("", ",")
    Debug.Print UBound(arr)
    
    
End Sub
