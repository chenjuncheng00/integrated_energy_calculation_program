Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Public Class 计算模式选择
    Private Sub 常规计算模式_Click(sender As Object, e As EventArgs) Handles 常规计算模式.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化一个主计算程序
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '读取输入的负荷调节精度
        Dim precision As Double = CType(Me.FHTJJD_shuru.Text, Double)
        '计算模式设置为1
        Dim calculation_mode As Integer = 1
        '——————————————————————————————————————————————————————————————————————————————————————————————
        If FHTJJD_shuru.Text = Nothing Then
            MsgBox("必须输入负荷调节进度参数，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If precision <= 0 Then
            MsgBox("必须输入正确的负荷调节进度参数，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '隐藏窗体
        Me.Hide()
        '调用主程序进行计算
        mainprogram.计算主程序(ExcelApp, precision, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, calculation_mode)
        '关闭窗口
        Me.Close()
    End Sub

    Private Sub 全局寻优计算_Click(sender As Object, e As EventArgs) Handles 全局寻优计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化一个主计算程序
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '读取输入的负荷调节精度
        Dim precision As Double = CType(Me.FHTJJD_shuru.Text, Double)
        '输入的购电单价和天然气单价参数
        Dim TRQ_price As Double = CType(TRQDJ.Text, Double)
        Dim D_price_GF1 As Double = CType(Me.GDDJ_GF1.Text, Double)
        Dim D_price_GF2 As Double = CType(Me.GDDJ_GF2.Text, Double)
        Dim D_price_F1 As Double = CType(Me.GDDJ_F1.Text, Double)
        Dim D_price_F2 As Double = CType(Me.GDDJ_F2.Text, Double)
        Dim D_price_P1 As Double = CType(Me.GDDJ_P1.Text, Double)
        Dim D_price_P2 As Double = CType(Me.GDDJ_P2.Text, Double)
        Dim D_price_G1 As Double = CType(Me.GDDJ_G1.Text, Double)
        Dim D_price_G2 As Double = CType(Me.GDDJ_G2.Text, Double)
        Dim D_price_QT1 As Double = CType(Me.GDDJ_QT1.Text, Double)
        Dim D_price_QT2 As Double = CType(Me.GDDJ_QT2.Text, Double)
        '计算模式设置为2
        Dim calculation_mode As Integer = 2
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '如果输入的为空，报错
        If FHTJJD_shuru.Text = Nothing Then
            MsgBox("必须输入负荷调节进度参数，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If GDDJ_GF1.Text = Nothing And GDDJ_GF2.Text = Nothing And GDDJ_F1.Text = Nothing And GDDJ_F2.Text = Nothing And GDDJ_P1.Text = Nothing And GDDJ_P2.Text = Nothing And GDDJ_G1.Text = Nothing And GDDJ_G2.Text = Nothing And GDDJ_QT1.Text = Nothing And GDDJ_QT2.Text = Nothing And TRQDJ.Text = Nothing Then
            MsgBox("购电单价和天然气单价必须根据实际情况正确输入，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If precision <= 0 Then
            MsgBox("必须输入正确的负荷调节进度参数，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        '如果没有输入价格，则报错
        If D_price_GF1 < 0 And D_price_GF2 < 0 And D_price_F1 < 0 And D_price_F2 < 0 And D_price_P1 < 0 And D_price_P2 < 0 And D_price_G1 < 0 And D_price_G2 < 0 And D_price_QT1 < 0 And D_price_QT2 < 0 Then
            MsgBox("必须输入正确的购电单价，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If TRQ_price < 0 Then
            MsgBox("必须输入正确的天然气单价，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If D_price_GF1 = 0 And D_price_GF2 = 0 And D_price_F1 = 0 And D_price_F2 = 0 And D_price_P1 = 0 And D_price_P2 = 0 And D_price_G1 = 0 And D_price_G2 = 0 And D_price_QT1 = 0 And D_price_QT2 = 0 And TRQ_price = 0 Then
            MsgBox("购电单价和天然气单价必须根据实际情况正确输入，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        'Excel中选择的用电时间段是否全部已经输入了电价
        '高峰1
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" And D_price_GF1 <= 0 Then
                MsgBox("没有输入正确的高峰1购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '高峰2
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" And D_price_GF2 <= 0 Then
                MsgBox("没有输入正确的高峰2购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '峰1
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" And D_price_F1 <= 0 Then
                MsgBox("没有输入正确的峰1购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '峰2
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" And D_price_F2 <= 0 Then
                MsgBox("没有输入正确的峰2购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '平1
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" And D_price_P1 <= 0 Then
                MsgBox("没有输入正确的平1购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '平2
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" And D_price_P2 <= 0 Then
                MsgBox("没有输入正确的平2购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '谷1
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" And D_price_G1 <= 0 Then
                MsgBox("没有输入正确的谷1购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '谷2
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" And D_price_G2 <= 0 Then
                MsgBox("没有输入正确的谷2购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '其它1
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And D_price_QT1 <= 0 Then
                MsgBox("没有输入正确的其它1购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '其它2
        For i = 1 To 50
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And D_price_QT2 <= 0 Then
                MsgBox("没有输入正确的其它2购电单价，请重新输入！")
                Exit Sub
            End If
        Next
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '隐藏窗体
        Me.Hide()
        '调用主程序进行计算
        mainprogram.计算主程序(ExcelApp, precision, D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode)
        '关闭窗口
        Me.Close()
    End Sub

End Class