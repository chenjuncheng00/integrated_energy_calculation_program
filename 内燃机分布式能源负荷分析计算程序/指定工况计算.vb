Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class 指定工况计算

    Private Sub 常规计算模式_Click(sender As Object, e As EventArgs) Handles 常规计算模式.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        'b表示当前正在计算的工况序号，n表示总工况数量
        Dim b As Integer
        Dim n As Integer
        '定义计算步长和负荷调节精度
        Dim JSBC As Integer
        Dim FHTJJD As Double
        '工况序号
        Dim GKXH(5) As Integer
        '计算模式设置为1
        Dim calculation_mode As Integer = 1
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '隐藏窗体
        Me.Hide()
        '————————————————————————————————————————————————————————————————————————————————————————
        If Me.TextBox1.Text <> Nothing Then
            GKXH(1) = CType(Me.TextBox1.Text, Integer)
        Else
            GKXH(1) = 0
        End If
        If Me.TextBox2.Text <> Nothing Then
            GKXH(2) = CType(Me.TextBox2.Text, Integer)
        Else
            GKXH(2) = 0
        End If
        If Me.TextBox3.Text <> Nothing Then
            GKXH(3) = CType(Me.TextBox3.Text, Integer)
        Else
            GKXH(3) = 0
        End If
        If Me.TextBox4.Text <> Nothing Then
            GKXH(4) = CType(Me.TextBox4.Text, Integer)
        Else
            GKXH(4) = 0
        End If
        If Me.TextBox5.Text <> Nothing Then
            GKXH(5) = CType(Me.TextBox5.Text, Integer)
        Else
            GKXH(5) = 0
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '计数，统计一共有多少种不同工况
        For i = 57 To 8 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(i, 24).Value > 0 Then
                n = i - 7 '工况总数
                Exit For '跳出循环
            End If
        Next
        '判断选择的制冷制热设备是否正确，有错误则报错并终止计算
        Dim ZTJC_EQ As Integer = mainprogram.判断制冷制热设备选择是否正确(ExcelApp, n)
        If ZTJC_EQ = 1 Then
            Call mainprogram.锁定工作表(ExcelApp)
            ZTJC_EQ = 0
            Exit Sub
        End If
        '读取输入的各工况冷负荷需求量、热负荷需求量、蓄冷量、蓄热量
        If n > 0 Then
            '让用户输入负荷调节精度
            FHTJJD = CType(Me.FHTJJD_shuru.Text, Double)
            '判断输入的各种负荷率是否有错误，有错误则报错并终止计算
            Dim ZTJC_SHUJU As Integer = mainprogram.读取输入的各种数据并添加报错功能(ExcelApp, FHTJJD, n)
            If ZTJC_SHUJU = 1 Then
                Call mainprogram.锁定工作表(ExcelApp)
                ZTJC_SHUJU = 0
                Exit Sub
            End If
            '针对输入的负荷调节精度，添加报错功能
            If FHTJJD = Nothing Then '输入的调节精度为空的情况
                MsgBox("输入的负荷调节精度不能为空，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            ElseIf FHTJJD = 0 Then '输入的调节精度为0的情况
                MsgBox("输入的负荷调节精度不能为0，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            ElseIf (FHTJJD <> 0 And (25 - FHTJJD * CInt(25 / FHTJJD)) <> 0) Then '输入的调节精度不能被25整除的情况
                MsgBox("输入的负荷调节精度必需能够被25整除，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            End If
            '根据用户输入的负荷调节精度，计算出最大计算步长
            JSBC = CInt(25 / FHTJJD)
            '指定工况进行计算
            '读取窗体文本框中的工况序号在“指定工况计算.确定”
            For i = 1 To 5
                If GKXH(i) > 0 Then
                    GKXH(i) = GKXH(i)
                End If
            Next
            For i = 1 To 5 '读取输入的工况序号，并添加报错功能
                If GKXH(i) > n Then
                    MsgBox("输入的工况序号不可以大于最大工况数量，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
                If GKXH(i) < 0 Then
                    MsgBox("输入的工况序号不可以为负数，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
            Next
            For i = 1 To 4
                If GKXH(i + 1) > 0 And GKXH(i) = 0 Then
                    MsgBox("输入的工况序号必需从上向下依次输入，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            For i = 1 To 5
                b = GKXH(i)
                If b > 0 Then '忽略为0的工况
                    Call mainprogram.清空指定工况输入输出数据(ExcelApp, b)
                    '进行正常的负荷分析（主要技术指标）计算
                    Call mainprogram.负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， 0, 0, calculation_mode, 0, 0)
                    '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call mainprogram.制冷和蓄冷空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    '只有全局寻优计算模式才修正
                    Call mainprogram.制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, calculation_mode)
                End If
            Next
            '计算循环体
            Call mainprogram.计算循环体(ExcelApp, n)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '梯级供热或者混水供热自动计算
            'Call 梯级或者混水供热计算()
        End If
        '判断各种计算结果是否正确，不正确则报错
        Call mainprogram.判断各种计算结果是否正确(ExcelApp, FHTJJD, n)
        '—————————————————————————————————————————————————————————————————————————————————————————
        '在窗体中显示计算已完成
        '实例化一个计算过程显示窗体
        Dim Calculate_Progress As New 计算进度显示
        Calculate_Progress.Show()
        Calculate_Progress.Label1.Text = "计算已经完成，请查看计算结果！"
        Calculate_Progress.TopMost = True
        Application.DoEvents()
        '清空输入的指定工况序号
        Me.TextBox1.Clear()
        Me.TextBox2.Clear()
        Me.TextBox3.Clear()
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        For i = 1 To 5 '清空已有的GKXH数组，防止出错
            GKXH(i) = 0
        Next
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        Me.Close()
    End Sub
    Private Sub 全局寻优计算_Click(sender As Object, e As EventArgs) Handles 全局寻优计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        'b表示当前正在计算的工况序号，n表示总工况数量
        Dim b As Integer
        Dim n As Integer
        '定义计算步长和负荷调节精度
        Dim JSBC As Integer
        Dim FHTJJD As Double
        '工况序号
        Dim GKXH(5) As Integer
        '输入的购电单价和天然气单价参数
        Dim D_price As Double = CType(Me.GDDJ.Text, Double)
        Dim TRQ_price As Double = CType(TRQDJ.Text, Double)
        '如果没有输入价格，则报错
        If D_price < 0 Then
            MsgBox("必须输入正确的购电单价，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If TRQ_price < 0 Then
            MsgBox("必须输入正确的天然气单价，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        If D_price = 0 And TRQ_price = 0 Then
            MsgBox("购电单价和天然气单价必须根据实际情况正确输入，否则无法进行全局寻优计算！")
            Exit Sub
        End If
        '计算模式设置为2
        Dim calculation_mode As Integer = 2
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '隐藏窗体
        Me.Hide()
        '————————————————————————————————————————————————————————————————————————————————————————
        If Me.TextBox1.Text <> Nothing Then
            GKXH(1) = CType(Me.TextBox1.Text, Integer)
        Else
            GKXH(1) = 0
        End If
        If Me.TextBox2.Text <> Nothing Then
            GKXH(2) = CType(Me.TextBox2.Text, Integer)
        Else
            GKXH(2) = 0
        End If
        If Me.TextBox3.Text <> Nothing Then
            GKXH(3) = CType(Me.TextBox3.Text, Integer)
        Else
            GKXH(3) = 0
        End If
        If Me.TextBox4.Text <> Nothing Then
            GKXH(4) = CType(Me.TextBox4.Text, Integer)
        Else
            GKXH(4) = 0
        End If
        If Me.TextBox5.Text <> Nothing Then
            GKXH(5) = CType(Me.TextBox5.Text, Integer)
        Else
            GKXH(5) = 0
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '计数，统计一共有多少种不同工况
        For i = 57 To 8 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(i, 24).Value > 0 Then
                n = i - 7 '工况总数
                Exit For '跳出循环
            End If
        Next
        '判断选择的制冷制热设备是否正确，有错误则报错并终止计算
        Dim ZTJC_EQ As Integer = mainprogram.判断制冷制热设备选择是否正确(ExcelApp, n)
        If ZTJC_EQ = 1 Then
            Call mainprogram.锁定工作表(ExcelApp)
            ZTJC_EQ = 0
            Exit Sub
        End If
        '读取输入的各工况冷负荷需求量、热负荷需求量、蓄冷量、蓄热量
        If n > 0 Then
            '让用户输入负荷调节精度
            FHTJJD = CType(Me.FHTJJD_shuru.Text, Double)
            '判断输入的各种负荷率是否有错误，有错误则报错并终止计算
            Dim ZTJC_SHUJU As Integer = mainprogram.读取输入的各种数据并添加报错功能(ExcelApp, FHTJJD, n)
            If ZTJC_SHUJU = 1 Then
                Call mainprogram.锁定工作表(ExcelApp)
                ZTJC_SHUJU = 0
                Exit Sub
            End If
            '针对输入的负荷调节精度，添加报错功能
            If FHTJJD = Nothing Then '输入的调节精度为空的情况
                MsgBox("输入的负荷调节精度不能为空，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            ElseIf FHTJJD = 0 Then '输入的调节精度为0的情况
                MsgBox("输入的负荷调节精度不能为0，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            ElseIf (FHTJJD <> 0 And (25 - FHTJJD * CInt(25 / FHTJJD)) <> 0) Then '输入的调节精度不能被25整除的情况
                MsgBox("输入的负荷调节精度必需能够被25整除，请重新输入！")
                Call mainprogram.锁定工作表(ExcelApp)
                Exit Sub
            End If
            '根据用户输入的负荷调节精度，计算出最大计算步长
            JSBC = CInt(25 / FHTJJD)
            '指定工况进行计算
            '读取窗体文本框中的工况序号在“指定工况计算.确定”
            For i = 1 To 5
                If GKXH(i) > 0 Then
                    GKXH(i) = GKXH(i)
                End If
            Next
            For i = 1 To 5 '读取输入的工况序号，并添加报错功能
                If GKXH(i) > n Then
                    MsgBox("输入的工况序号不可以大于最大工况数量，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
                If GKXH(i) < 0 Then
                    MsgBox("输入的工况序号不可以为负数，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
            Next
            For i = 1 To 4
                If GKXH(i + 1) > 0 And GKXH(i) = 0 Then
                    MsgBox("输入的工况序号必需从上向下依次输入，请重新输入")
                    Call mainprogram.锁定工作表(ExcelApp)
                    Exit Sub
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            For i = 1 To 5
                b = GKXH(i)
                If b > 0 Then '忽略为0的工况
                    Call mainprogram.清空指定工况输入输出数据(ExcelApp, b)
                    '进行正常的负荷分析（主要技术指标）计算
                    Call mainprogram.负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， D_price, TRQ_price, calculation_mode, 0, 0)
                    '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call mainprogram.制冷和蓄冷空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    '只有全局寻优计算模式才修正
                    Call mainprogram.制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, calculation_mode)
                End If
            Next
            '计算循环体
            Call mainprogram.计算循环体(ExcelApp, n)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '梯级供热或者混水供热自动计算
            'Call 梯级或者混水供热计算()
        End If
        '判断各种计算结果是否正确，不正确则报错
        Call mainprogram.判断各种计算结果是否正确(ExcelApp, FHTJJD, n)
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '在窗体中显示计算已完成
        '实例化一个计算过程显示窗体
        Dim Calculate_Progress As New 计算进度显示
        Calculate_Progress.Show()
        Calculate_Progress.Label1.Text = "计算已经完成，请查看计算结果！"
        Calculate_Progress.TopMost = True
        Application.DoEvents()
        '清空输入的指定工况序号
        Me.TextBox1.Clear()
        Me.TextBox2.Clear()
        Me.TextBox3.Clear()
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        For i = 1 To 5 '清空已有的GKXH数组，防止出错
            GKXH(i) = 0
        Next
        '—————————————————————————————————————————————————————————————————————————————————————————
        Me.Close()
    End Sub

    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim key As String
        key = e.KeyChar
        '检验按键是否为回车键，如果是就把焦点附给按钮1，并执行Click命令
        If key = Microsoft.VisualBasic.ChrW(13) Then
            常规计算模式.Focus()
            常规计算模式.PerformClick()
        End If
    End Sub
    Private Sub 指定工况计算_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyBase.KeyPreview = True
    End Sub
End Class