Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO

Public Class 典型日输入参数赋值
    Private Sub 确定参数_Click(sender As Object, e As EventArgs) Handles 确定参数.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '清空已有的计算结果
        '第49到200
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B56:EK207").ClearContents
        '判断是否需要清空25到48
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 12).Value <> Nothing And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 1, 12).Value <> Nothing Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B32:EK55").ClearContents
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 21).Value <> Nothing And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 1, 21).Value <> Nothing Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B32:EK55").ClearContents
        End If
        '——————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————
        '读取输入的各种参数
        '冷负荷占比
        Dim LFH_100_num As Double = CType(Me.LFH_100.Text, Double) / 100
        Dim LFH_75_num As Double = CType(Me.LFH_75.Text, Double) / 100
        Dim LFH_50_num As Double = CType(Me.LFH_50.Text, Double) / 100
        Dim LFH_25_num As Double = CType(Me.LFH_25.Text, Double) / 100
        '热负荷占比
        Dim RFH_100_num As Double = CType(Me.RFH_100.Text, Double) / 100
        Dim RFH_75_num As Double = CType(Me.RFH_75.Text, Double) / 100
        Dim RFH_50_num As Double = CType(Me.RFH_50.Text, Double) / 100
        Dim RFH_25_num As Double = CType(Me.RFH_25.Text, Double) / 100
        '占比求和
        Dim LFH_num_all As Double = LFH_100_num + LFH_75_num + LFH_50_num + LFH_25_num
        Dim RFH_num_all As Double = RFH_100_num + RFH_75_num + RFH_50_num + RFH_25_num
        '——————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————
        '读取前24个工况输入的冷负荷和热负荷
        Dim LFH_1_24 As New List(Of Double)
        Dim RFH_1_24 As New List(Of Double)
        '读取25到48个工况输入的冷负荷和热负荷
        Dim LFH_25_48 As New List(Of Double)
        Dim RFH_25_48 As New List(Of Double)
        '读入数据1到24
        For i = 1 To 24
            '冷负荷
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value <> Nothing Then
                Dim LFH_1_24_temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value
                LFH_1_24.Add(LFH_1_24_temp)
            End If
            '热负荷
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value <> Nothing Then
                Dim RFH_1_24_temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value
                RFH_1_24.Add(RFH_1_24_temp)
            End If
        Next
        '读入数据25到48
        For i = 25 To 48
            '冷负荷
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value <> Nothing Then
                Dim LFH_25_48_temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value
                LFH_25_48.Add(LFH_25_48_temp)
            End If
            '热负荷
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value <> Nothing Then
                Dim RFH_25_48_temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value
                RFH_25_48.Add(RFH_25_48_temp)
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '判断输入的各种参数是否正确
        '不可以输入负数
        If LFH_100_num < 0 Or LFH_75_num < 0 Or LFH_50_num < 0 Or LFH_25_num < 0 Then
            MsgBox("冷负荷100典型日、75%典型日、50%典型日、25%典型日占比的数据必须全部是大于0的数，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        If RFH_100_num < 0 Or RFH_75_num < 0 Or RFH_50_num < 0 Or RFH_25_num < 0 Then
            MsgBox("热负荷100典型日、75%典型日、50%典型日、25%典型日占比的数据必须全部是大于0的数，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        '求和，必须为100%（但是允许1%的偏差）
        If LFH_num_all > 0 And (LFH_num_all > 1.01 Or LFH_num_all < 0.99) Then
            MsgBox("冷负荷100典型日、75%典型日、50%典型日、25%典型日占比输入的数据之和必须等于100%，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        If RFH_num_all > 0 And (RFH_num_all > 1.01 Or RFH_num_all < 0.99) Then
            MsgBox("热负荷100典型日、75%典型日、50%典型日、25%典型日占比输入的数据之和必须等于100%，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        '可以全部不输入（假设没有冷或者热负荷），但是如果输入了其中一个，其它三个必须也输入
        If LFH_num_all > 0 And (LFH_100_num = 0 Or LFH_75_num = 0 Or LFH_50_num = 0 Or LFH_25_num = 0) Then
            MsgBox("冷负荷100典型日、75%典型日、50%典型日、25%典型日占比的数据必须全部输入，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        If RFH_num_all > 0 And (RFH_100_num = 0 Or RFH_75_num = 0 Or RFH_50_num = 0 Or RFH_25_num = 0) Then
            MsgBox("热负荷100典型日、75%典型日、50%典型日、25%典型日占比的数据必须全部输入，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        '输入的冷负荷数量必须是24个，忽略为0的情况
        If LFH_1_24.Count + LFH_25_48.Count <> 24 And LFH_1_24.Count + LFH_25_48.Count <> 0 Then
            MsgBox("输入的初始冷负荷数量必须是24个，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        '输入的热负荷数量必须是24个，忽略为0的情况
        If RFH_1_24.Count + RFH_25_48.Count <> 24 And RFH_1_24.Count + RFH_25_48.Count <> 0 Then
            MsgBox("输入的初始热负荷数量必须是24个，请重新输入！", vbOKOnly)
            Exit Sub
        End If
        '如果输入了冷负荷，但是没输入冷负荷占比数据
        If LFH_1_24.Count + LFH_25_48.Count <> 0 And (LFH_100_num = 0 Or LFH_75_num = 0 Or LFH_50_num = 0 Or LFH_25_num = 0) Then
            MsgBox("请输入冷负荷100典型日、75%典型日、50%典型日、25%典型日的占比数据！", vbOKOnly)
            Exit Sub
        End If
        '如果输入了热负荷，但是没输入热负荷占比数据
        If RFH_1_24.Count + RFH_25_48.Count <> 0 And (RFH_100_num = 0 Or RFH_75_num = 0 Or RFH_50_num = 0 Or RFH_25_num = 0) Then
            MsgBox("请输入热负荷100典型日、75%典型日、50%典型日、25%典型日的占比数据！", vbOKOnly)
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————
        '隐藏窗体
        Me.Hide()
        '——————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————
        '如果第25行没有输入冷负荷和热负荷，都是空的
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 12).Value = Nothing And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 21).Value = Nothing Then
            '则说明本次计算只有冷负荷或者只有热负荷
            '第1到24行
            For i = 1 To 24
                '冷负荷100%典型日
                '热负荷100%典型日
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_100_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_100_num
                End If
            Next
            '第25到48行
            For i = 25 To 48
                '冷负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 12).Value
                '热负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 25).Value / 0.75
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_75_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_75_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 27).Value
                '公用计算状态选择
                For j = 77 To 83
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '日期序号选择
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 2
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 2 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 3 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 4
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 4 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 3
                End If
            Next
            '第49到72行
            For i = 49 To 72
                '冷负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24, 12).Value
                '热负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24, 25).Value / 0.5
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_50_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_50_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 27).Value
                '公用计算状态选择
                For j = 77 To 83
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '日期序号选择
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 2
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 2 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 3 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 4
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 4 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 3
                End If
            Next
            '第73到96行
            For i = 73 To 96
                '冷负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24 - 24, 12).Value
                '热负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24 - 24, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24 - 24 - 24, 25).Value / 0.25
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_25_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_25_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 27).Value
                '公用计算状态选择
                For j = 77 To 83
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, j).Value
                Next
                '日期序号选择
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 2
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 2 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 3 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 4
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 24, 84).Value = 4 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 84).Value = 3
                End If
            Next
            '所有结果保留3位小数，增加美观
            For i = 1 To 96
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value, 3)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value, 3)
                End If
            Next
            '清空冷热负荷结果为0的单元格
            For i = 25 To 96
                '冷负荷
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = Nothing
                End If
                '热负荷
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = Nothing
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————
            '如果第25行输入了冷负荷或者热负荷，只有其中一个是空的
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 12).Value <> Nothing Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + 25, 21).Value <> Nothing Then
            '则说明本次计算冷负荷和热负荷都有
            '第1到24行
            For i = 1 To 24
                '冷负荷100%典型日
                '热负荷100%典型日
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_100_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_100_num
                End If
            Next
            '第25到48行
            For i = 25 To 48
                '冷负荷100%典型日
                '热负荷100%典型日
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_100_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_100_num
                End If
            Next
            '第49到72行
            For i = 49 To 72
                '冷负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 12).Value
                '热负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 25).Value / 0.75
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_75_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_75_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '第73到96行
            For i = 73 To 96
                '冷负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 12).Value
                '热负荷75%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.75 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 25).Value / 0.75
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_75_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_75_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '第97到120行
            For i = 97 To 120
                '冷负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 12).Value
                '热负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 25).Value / 0.5
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_50_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_50_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '第121到144行
            For i = 121 To 144
                '冷负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 12).Value
                '热负荷50%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.5 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48, 25).Value / 0.5
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_50_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_50_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '第145到168行
            For i = 145 To 168
                '冷负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 12).Value
                '热负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 25).Value / 0.25
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_25_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_25_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '第169到192行
            For i = 169 To 192
                '冷负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 12).Value
                '热负荷25%典型日
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0.25 * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 21).Value
                '供冷设备选择
                For j = 6 To 11
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '供热设备选择
                For j = 15 To 20
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
                '负荷段全年天数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 24).Value
                '负荷段每天小时数
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48 - 48 - 48, 25).Value / 0.25
                '负荷段时间频数
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    '冷负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = LFH_25_num
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    '热负荷占比
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 26).Value = RFH_25_num
                End If
                '向外供电功率上限
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, 27).Value
                '公用计算状态选择
                For j = 77 To 84
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i - 48, j).Value
                Next
            Next
            '所有结果保留3位小数，增加美观
            For i = 1 To 192
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value, 3)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value, 3)
                End If
            Next
            '清空冷热负荷结果为0的单元格
            For i = 49 To 192
                '冷负荷
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = Nothing
                End If
                '热负荷
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = Nothing
                End If
            Next
        End If
        '——————————————————————————————————————————————————————————————————————————————————————
        MsgBox("100典型日、75%典型日、50%典型日、25%典型日计算边界输入完成！", vbOKOnly)
        Me.Close()
    End Sub

    Private Sub 清空数据_Click(sender As Object, e As EventArgs) Handles 清空数据.Click
        Me.LFH_100.Text = Nothing
        Me.LFH_75.Text = Nothing
        Me.LFH_50.Text = Nothing
        Me.LFH_25.Text = Nothing
        Me.RFH_100.Text = Nothing
        Me.RFH_75.Text = Nothing
        Me.RFH_50.Text = Nothing
        Me.RFH_25.Text = Nothing
    End Sub
End Class