Public Class 蓄能分配计算设置
    Private Sub 确定参数_Click(sender As Object, e As EventArgs) Handles 确定参数.Click
        On Error Resume Next
        '读取输入的蓄冷平均功率、最大功率
        Dim XLGL_PJ As Double = CType(Me.XLGL_PJ.Text, Double)
        Dim XLGL_MAX As Double = CType(Me.XLGL_MAX.Text, Double)
        '读取输入的蓄热平均功率、最大功率
        Dim XRGL_PJ As Double = CType(Me.XRGL_PJ.Text, Double)
        Dim XRGL_MAX As Double = CType(Me.XRGL_MAX.Text, Double)
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '判断选择是否正确
        '蓄冷分配
        If XL_MAX_CHECK.Checked = True And XLGL_MAX = 0 Then
            MsgBox("勾选了蓄冷分配计算存在最大值选项，但是没有输入蓄冷最大功率，请重新输入！！")
            Exit Sub
        End If
        If XL_MAX_CHECK.Checked = True And XLGL_MAX < XLGL_PJ Then
            MsgBox("输入蓄冷最大功率不可以小于蓄冷平均功率，请重新输入！！")
            Exit Sub
        End If
        If XL_MAX_CHECK.Checked = False And XLGL_MAX > 0 Then
            MsgBox("没有勾选蓄冷分配计算存在最大值选项，不可以输入蓄冷最大功率，请重新输入！！")
            Exit Sub
        End If
        '蓄热分配
        If XR_MAX_CHECK.Checked = True And XRGL_MAX = 0 Then
            MsgBox("勾选了蓄热分配计算存在最大值选项，但是没有输入蓄热最大功率，请重新输入！！")
            Exit Sub
        End If
        If XR_MAX_CHECK.Checked = True And XRGL_MAX < XRGL_PJ Then
            MsgBox("输入蓄热最大功率不可以小于蓄热平均功率，请重新输入！！")
            Exit Sub
        End If
        If XR_MAX_CHECK.Checked = False And XRGL_MAX > 0 Then
            MsgBox("没有勾选蓄热分配计算存在最大值选项，不可以输入蓄热最大功率，请重新输入！！")
            Exit Sub
        End If
        '不可以小于0
        If XLGL_PJ < 0 Or XLGL_MAX < 0 Or XRGL_PJ < 0 Or XRGL_MAX < 0 Then
            MsgBox("不可以输入小于0的数值，请重新输入！！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的情况，确定计算模式
        '蓄冷计算模式
        Dim XLJS_MS As Integer
        If XL_MAX_CHECK.Checked = True Then
            XLJS_MS = 1
        Else
            XLJS_MS = 2
        End If
        '蓄热计算模式
        Dim XRJS_MS As Integer
        If XR_MAX_CHECK.Checked = True Then
            XRJS_MS = 1
        Else
            XRJS_MS = 2
        End If
        '实例化与一个主计算程序
        Dim mainprogram As New Com内燃机分布式能源负荷分析计算程序
        '进入计算
        Call mainprogram.蓄能分配计算主程序(XLGL_PJ, XLGL_MAX, XLJS_MS, XRGL_PJ, XRGL_MAX, XRJS_MS)
        '—————————————————————————————————————————————————————————————————————————————————————————
        Me.Close()
    End Sub

    Private Sub 清空数据_Click(sender As Object, e As EventArgs) Handles 清空数据.Click
        Me.XLGL_PJ = Nothing
        Me.XLGL_MAX = Nothing
        Me.XL_MAX_CHECK.Checked = False
        Me.XRGL_PJ = Nothing
        Me.XRGL_MAX = Nothing
        Me.XR_MAX_CHECK.Checked = False
    End Sub
End Class