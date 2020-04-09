Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class 指定工况计算
    Public Shared GKXH(10000) As Integer '工况序号
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles 确定.Click
        '本SUB为主程序
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '隐藏窗体
        Me.Hide()
        '————————————————————————————————————————————————————————————————————————————————————————
        If Com内燃机分布式能源负荷分析计算程序.Form1.TextBox1.Text <> Nothing Then
            GKXH(1) = CType(Com内燃机分布式能源负荷分析计算程序.Form1.TextBox1.Text, Integer)
        Else
            GKXH(1) = 0
        End If
        If Com内燃机分布式能源负荷分析计算程序.Form1.TextBox2.Text <> Nothing Then
            GKXH(2) = CType(Com内燃机分布式能源负荷分析计算程序.Form1.TextBox2.Text, Integer)
        Else
            GKXH(2) = 0
        End If
        If Com内燃机分布式能源负荷分析计算程序.Form1.TextBox3.Text <> Nothing Then
            GKXH(3) = CType(Com内燃机分布式能源负荷分析计算程序.Form1.TextBox3.Text, Integer)
        Else
            GKXH(3) = 0
        End If
        If Com内燃机分布式能源负荷分析计算程序.Form1.TextBox4.Text <> Nothing Then
            GKXH(4) = CType(Com内燃机分布式能源负荷分析计算程序.Form1.TextBox4.Text, Integer)
        Else
            GKXH(4) = 0
        End If
        If Com内燃机分布式能源负荷分析计算程序.Form1.TextBox5.Text <> Nothing Then
            GKXH(5) = CType(Com内燃机分布式能源负荷分析计算程序.Form1.TextBox5.Text, Integer)
        Else
            GKXH(5) = 0
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '计数，统计一共有多少种不同工况
        For i = 57 To 8 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(i, 24).Value > 0 Then
                Com内燃机分布式能源负荷分析计算程序.n = i - 7 '工况总数
                Exit For '跳出循环
            End If
        Next
        '判断选择的制冷制热设备是否正确，有错误则报错并终止计算
        Call mainprogram.判断制冷制热设备选择是否正确()
        If Com内燃机分布式能源负荷分析计算程序.ZTJC = 1 Then
            Call mainprogram.锁定工作表()
            Com内燃机分布式能源负荷分析计算程序.ZTJC = 0
            Exit Sub
        End If
        '读取输入的各工况冷负荷需求量、热负荷需求量、蓄冷量、蓄热量
        If Com内燃机分布式能源负荷分析计算程序.n > 0 Then
            '判断输入的各种负荷率是否有错误，有错误则报错并终止计算
            Call mainprogram.读取输入的各种数据并添加报错功能()
            If Com内燃机分布式能源负荷分析计算程序.ZTJC = 1 Then
                Call mainprogram.锁定工作表()
                Com内燃机分布式能源负荷分析计算程序.ZTJC = 0
                Exit Sub
            End If
            '让用户输入负荷调节精度
            Com内燃机分布式能源负荷分析计算程序.FHTJJD = InputBox("请输入在进行计算时，各种设备的负荷调节精度（单位为%），请输入可以被25整除的数字，输入的数字越小，计算精度越高，但计算速度越慢！", "请输入设备负荷调节精度（单位为%）", 0.5)
            '针对输入的负荷调节精度，添加报错功能
            If Com内燃机分布式能源负荷分析计算程序.FHTJJD = Nothing Then '输入的调节精度为空的情况
                MsgBox("输入的负荷调节精度不能为空，请重新输入！")
                Call mainprogram.锁定工作表()
                Exit Sub
            ElseIf Com内燃机分布式能源负荷分析计算程序.FHTJJD = 0 Then '输入的调节精度为0的情况
                MsgBox("输入的负荷调节精度不能为0，请重新输入！")
                Call mainprogram.锁定工作表()
                Exit Sub
            ElseIf (Com内燃机分布式能源负荷分析计算程序.FHTJJD <> 0 And (25 - Com内燃机分布式能源负荷分析计算程序.FHTJJD * CInt(25 / Com内燃机分布式能源负荷分析计算程序.FHTJJD)) <> 0) Then '输入的调节精度不能被25整除的情况
                MsgBox("输入的负荷调节精度必需能够被25整除，请重新输入！")
                Call mainprogram.锁定工作表()
                Exit Sub
            End If
            '根据用户输入的负荷调节精度，计算出最大计算步长
            Com内燃机分布式能源负荷分析计算程序.JSBC = CInt(25 / Com内燃机分布式能源负荷分析计算程序.FHTJJD)
            '指定工况进行计算
            '读取窗体文本框中的工况序号在“指定工况计算.确定”
            For i = 1 To 5
                If GKXH(i) > 0 Then
                    Com内燃机分布式能源负荷分析计算程序.GKXH(i) = GKXH(i)
                End If
            Next
            For i = 1 To 5 '读取输入的工况序号，并添加报错功能
                If GKXH(i) > Com内燃机分布式能源负荷分析计算程序.n Then
                    MsgBox("输入的工况序号不可以大于最大工况数量，请重新输入")
                    Call mainprogram.锁定工作表()
                    Exit Sub
                End If
                If GKXH(i) < 0 Then
                    MsgBox("输入的工况序号不可以为负数，请重新输入")
                    Call mainprogram.锁定工作表()
                    Exit Sub
                End If
                If GKXH(i + 1) > 0 And GKXH(i) = 0 Then
                    MsgBox("输入的工况序号必需从上向下依次输入，请重新输入")
                    Call mainprogram.锁定工作表()
                    Exit Sub
                End If
            Next
            For i = 1 To 5
                Com内燃机分布式能源负荷分析计算程序.b = GKXH(i)
                If Com内燃机分布式能源负荷分析计算程序.b > 0 Then '忽略为0的工况
                    Call mainprogram.清空指定工况输入输出数据()
                    Call mainprogram.负荷分析计算程序()
                    Call mainprogram.存在混水供热的工况特殊处理()
                End If
            Next
            '将上面计算出的内燃机负荷率中，单台负荷率低于30%的内燃机负荷率修改为0
            Call mainprogram.将内燃机单台负荷率低于百分之30的内燃机关闭()
            '计算循环体
            Call mainprogram.计算循环体()
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '梯级供热或者混水供热自动计算
            'Call 梯级或者混水供热计算()
        End If
        '判断各种计算结果是否正确，不正确则报错
        Call mainprogram.判断各种计算结果是否正确()
        '—————————————————————————————————————————————————————————————————————————————————————————
        '在窗体中显示计算已完成
        '实例化一个计算过程显示窗体
        Dim Calculate_Progress As New 计算进度显示
        Calculate_Progress.Show()
        Calculate_Progress.Label1.Text = "计算已经完成，请查看计算结果！"
        Calculate_Progress.TopMost = True
        Application.DoEvents()
        '清空输入的指定工况序号
        Com内燃机分布式能源负荷分析计算程序.Form1.TextBox1.Clear()
        Com内燃机分布式能源负荷分析计算程序.Form1.TextBox2.Clear()
        Com内燃机分布式能源负荷分析计算程序.Form1.TextBox3.Clear()
        Com内燃机分布式能源负荷分析计算程序.Form1.TextBox4.Clear()
        Com内燃机分布式能源负荷分析计算程序.Form1.TextBox5.Clear()
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
            确定.Focus()
            确定.PerformClick()
        End If
    End Sub
    Private Sub 指定工况计算_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyBase.KeyPreview = True
    End Sub
End Class