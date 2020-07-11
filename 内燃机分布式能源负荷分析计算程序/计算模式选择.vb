Public Class 计算模式选择
    Private Sub 常规计算模式_Click(sender As Object, e As EventArgs) Handles 常规计算模式.Click
        On Error Resume Next
        '实例化一个主计算程序
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '读取输入的负荷调节精度
        Dim precision As Double = CType(Me.FHTJJD_shuru.Text, Double)
        '计算模式设置为1
        Dim calculation_mode As Integer = 1
        '调用主程序进行计算
        mainprogram.计算主程序(precision, 0, 0, calculation_mode)
        '关闭窗口
        Me.Close()
    End Sub

    Private Sub 全局寻优计算_Click(sender As Object, e As EventArgs) Handles 全局寻优计算.Click
        On Error Resume Next
        '实例化一个主计算程序
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '读取输入的负荷调节精度
        Dim precision As Double = CType(Me.FHTJJD_shuru.Text, Double)
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
        '调用主程序进行计算
        mainprogram.计算主程序(precision, D_price, TRQ_price, calculation_mode)
        '关闭窗口
        Me.Close()
    End Sub

End Class