Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Module 辅助程序
    Function Excel版本号验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '验证Excel表格的更新时间
        Dim ZTJC_EXCEL As Integer
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(3, 20).Value < 20200711 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            ZTJC_EXCEL = 1
        Else
            ZTJC_EXCEL = 0
        End If
        Call 自保护程序(ExcelApp)
        '返回状态监测数值
        Return ZTJC_EXCEL
    End Function
    Sub 计算循环体(ExcelApp As Object, n As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        '计算
        For i = 1 To n
            '读取计算输入量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 141)).Value
            '返回计算结果
            ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + i, 54)).Value = ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 54)).Value
            '返回内燃机及其余热利用计算结果
            ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 15), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 18)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 15), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 18)).Value
            ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 23), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 30)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 23), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 30)).Value
            ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 54), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 57)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 54), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 57)).Value
            ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 62), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 69)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 62), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 69)).Value
            '返回设备运行信息结果
            ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range(ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 63)).Value = ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range(ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(3, 63)).Value
        Next
    End Sub
    Sub 清空输入输出数据(ExcelApp As Object)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B3:EK3").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("AB8:BX207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("CL8:EK207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range("B8:BB207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range("B8:BK207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range("O9:R207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range("W9:AD207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range("BB9:BE207").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range("BJ9:BQ207").ClearContents
    End Sub
    Sub 清空制冷和蓄冷设备负荷率计算结果(ExcelApp As Object, b As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        '将已有的制冷设备运行负荷率重置为0，每一次重新计算必需重置一次0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 39)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 91)).Value = 0
        '清空已有的蓄冷设备运行计算结果，每次重新计算时都必须清空一次
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51)).Value = 0
        '全局寻优计算模式设备启动数量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 125)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 125)).Value = Nothing
    End Sub
    Sub 清空制热和蓄热设备负荷率计算结果(ExcelApp As Object, b As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        '将已有的制热设备和蓄热设备运行负荷率重置为0，每一次重新计算必需重置一次0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 70)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 70)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97)).Value = 0
        '清空梯级供热和混水供热设备负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 80), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 83)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = Nothing
        '全局寻优计算模式设备启动数量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 126), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 126), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 141)).Value = Nothing
    End Sub
    Sub 清空指定工况输入输出数据(ExcelApp As Object, b As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        '清空数据
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 76)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97)).Value = Nothing
        '清空梯级供热和混水供热设备负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = Nothing
        '修正系数改为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 109)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + b, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + b, 54)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range(ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + b, 2), ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + b, 63)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 15), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 18)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 23), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 30)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 54), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 57)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 62), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + b, 69)).Value = Nothing
        '全局寻优计算模式设备启动数量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 125)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 126), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 141)).Value = Nothing
    End Sub
    Sub 工作表设置为彻底隐藏(ExcelApp As Object)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.Worksheets("逐负荷成本计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("水蓄冷").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("水蓄热").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("内燃机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("烟气热水型溴化锂").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式冷水机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("水冷螺杆机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("风冷螺杆式冷水（热泵）机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("水（地）源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（低温出水）").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（中温出水）").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（高温出水）").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式热泵机组（超高温出水）").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("空气源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("内燃机及其余热利用计算结果").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("逐负荷电量计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("各设备运行数据计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("直燃型溴化锂机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
    End Sub
    Sub 取消工作表彻底隐藏(ExcelApp As Object)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.Worksheets("逐负荷成本计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("水蓄冷").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("水蓄热").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("内燃机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("烟气热水型溴化锂").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式冷水机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("水冷螺杆机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("风冷螺杆式冷水（热泵）机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("水（地）源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（低温出水）").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（中温出水）").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式热泵（冷水）机组（高温出水）").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式热泵机组（超高温出水）").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("空气源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("内燃机及其余热利用计算结果").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("逐负荷电量计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("各设备运行数据计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("直燃型溴化锂机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
    End Sub
    Sub 锁定工作表(ExcelApp As Object)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("逐负荷成本计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水蓄冷").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水蓄热").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本测算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入测算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("内燃机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("烟气热水型溴化锂").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式冷水机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水冷螺杆机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("风冷螺杆式冷水（热泵）机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水（地）源热泵（冷水）机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（低温出水）").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（中温出水）").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（高温出水）").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵机组（超高温出水）").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("空气源热泵（冷水）机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("供能单价的确定").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("逐负荷电量计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("各设备运行数据计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("直燃型溴化锂机组").Protect(Password:="wscjc")
    End Sub
    Sub 解锁工作表(ExcelApp As Object)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("计算输入").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("逐负荷成本计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水蓄冷").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水蓄热").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本测算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入测算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("内燃机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("烟气热水型溴化锂").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式冷水机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水冷螺杆机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("风冷螺杆式冷水（热泵）机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("水（地）源热泵（冷水）机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（低温出水）").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（中温出水）").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵（冷水）机组（高温出水）").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式热泵机组（超高温出水）").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("空气源热泵（冷水）机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("供能单价的确定").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("逐负荷电量计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("各设备运行数据计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("直燃型溴化锂机组").unProtect(Password:="wscjc")
    End Sub
    Sub 自保护程序(ExcelApp As Object)
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            '验证本地C盘是否存在白名单文件，用来区分是不是自己的电脑，是否需要进行自保护程序验证
            Dim fs As New FileStream("C:\Windows\WhiteList_CJC.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim WhiteList_PC As String = strTemp '获取TXT文本内容
            sr.Close()
            fs.Close()
            If WhiteList_PC <> "WhiteList_PC" Then '如果不是合格的白名单文件，验证失败，则进行自保护验证
                'Call 获取本地服务器版本信息并验证()
                Call 获取本机MAC地址并验证(ExcelApp)
                Call 网络时间和本地时间交替验证(ExcelApp)
            End If
        Catch ex As Exception
            '发生任何异常，则进行自保护验证
            'Call 获取本地服务器版本信息并验证()
            Call 获取本机MAC地址并验证(ExcelApp)
            Call 网络时间和本地时间交替验证(ExcelApp)
            Exit Sub
        End Try
    End Sub
    Sub 获取本地服务器版本信息并验证(ExcelApp As Object)
        'On Error Resume Next        
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            Dim fs As New FileStream("\\192.168.9.201\user1\CJC-陈俊丞\综合能源多联供项目计算工具\内燃机版计算程序\Version.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim FWQBBH_String As String = strTemp '服务器版本号，字符串格式
            sr.Close()
            fs.Close()
            Dim FWQBBH = CInt(FWQBBH_String) '将服务器版本号从字符串格式转为整数型格式
            If FWQBBH > 20190401 Then '如果服务器版本号大于设定的内置版本号信息，则退出程序
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Version Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        Catch ex As Exception
            '发生任何异常，直接退出
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Version Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
            Exit Sub
        End Try
    End Sub
    Sub 获取本机MAC地址并验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim wmiObjSet As WbemScripting.SWbemObjectSet
        Dim obj As WbemScripting.SWbemObject
        Dim MAC As String = Nothing
        Dim MACTEST As Integer = 0 'MAC地址检测，1代表通过，0代表不通过
        wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapterConfiguration")
        For Each obj In wmiObjSet
            MAC = obj.MACAddress
            'MAC地址白名单(随便写一个)
            If MAC = "45:40:C1:8F:AD:4A" Then
                MACTEST = 1
                Exit For
            Else
                MACTEST = 0
            End If
        Next
        '————————————————————————————————————————————————————————————
        '针对MAC地址检测结果，进行不同操作
        If MACTEST = 0 Then '如果MAC地址检测不通过
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local MAC Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Sub 网络时间和本地时间交替验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '程序开始
        '先验证网络时间
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = "TimeOut" '网络状态默认值
        For i = 1 To 10 '最多Ping网络10次
            A = CStr(Pings()) '反回函数状态
            If A = "Success" Then
                A = "Success"
                Exit For '只要检测出ping通，就跳出循环
            End If
        Next
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then '如果网络是通的
            With CreateObject("MSXML2.ServerXMLHTTP") '获取网络时间
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #01/01/2020# Then '月/日/年，验证网络时间
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                End If
            End With
        Else '如果网络不通，验证系统本地时间
            '进程暂停一段时间（2000分钟）
            Threading.Thread.Sleep(120000000)
            Dim Local_Time = Date.Now '获取系统本地时间
            Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
            If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Internet Ping Error and Local Date Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        End If
    End Sub
    Sub 获取系统时间并验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim Local_Time = Date.Now '获取系统本地时间
        Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
        If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local Date Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Sub 程序联网验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = CStr(Pings()) '反回函数状态
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then
            With CreateObject("MSXML2.ServerXMLHTTP")
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #01/01/2020# Then '月/日/年
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                End If
            End With
        Else
            MsgBox("网络连接异常，请检查网络连接！本程序必需联网验证才可以使用！")
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Function Pings() As String
        Dim _ping As New Net.NetworkInformation.Ping
        Dim _pingreply As Net.NetworkInformation.PingReply = _ping.Send("www.baidu.com") 'ping 百度
        Return _pingreply.Status.ToString()
        '返回以下信息
        'TimedOut 失败
        'Success  成功
    End Function
End Module
