Module 计算前后正确性判断和系数计算
    Sub 将内燃机单台负荷率低于百分之30的内燃机关闭(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSLFH As Double, HSRFH As Double， D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '定义局部变量
        '当前的内燃机单台负荷率
        Dim ZLNRJDTFHL1 As Double '制冷内燃机1单台负荷率
        Dim ZLNRJDTFHL2 As Double '制冷内燃机2单台负荷率
        Dim ZRNRJDTFHL1 As Double '制热内燃机1单台负荷率
        Dim ZRNRJDTFHL2 As Double '制热内燃机2单台负荷率
        '内燃机装机数量
        '读取制冷季装机方案及参数
        Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
        '制冷内燃机1装机数量
        Dim ZLNRJZJSL1 As Double = ans_ZJFA_L(0)
        '制冷内燃机2装机数量
        Dim ZLNRJZJSL2 As Double = ans_ZJFA_L(1)
        '读取采暖季装机方案及参数
        Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
        '制热内燃机1装机数量
        Dim ZRNRJZJSL1 As Double = ans_ZJFA_R(0)
        '制热内燃机2装机数量
        Dim ZRNRJZJSL2 As Double = ans_ZJFA_R(1)
        '检查每个工况的内燃机负荷率，如果单台负荷率低于30%，则修改为0，关闭内燃机
        '制冷内燃机1
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2).Value > 0 Then
            ZLNRJDTFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2).Value * ZLNRJZJSL1 '制冷内燃机1单台负荷率
            If ZLNRJDTFHL1 < FHL1_min_NRJ Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2).Value = 0
                '进行制冷设备计算
                Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                'MsgBox("制冷时间段内燃机(1)负荷率低于30%，被设置为0" & Chr(10) & "工况序号为：" & b)
            End If
        End If
        '制冷内燃机2
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 3).Value > 0 Then
            ZLNRJDTFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 3).Value * ZLNRJZJSL2 '制冷内燃机2单台负荷率
            If ZLNRJDTFHL2 < FHL2_min_NRJ Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 3).Value = 0
                '进行制冷设备计算
                Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                'MsgBox("制冷时间段内燃机(2)负荷率低于30%，被设置为0" & Chr(10) & "工况序号为：" & b)
            End If
        End If
        '制热内燃机1
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value > 0 Then
            ZRNRJDTFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value * ZRNRJZJSL1 '制热内燃机1单台负荷率
            If ZRNRJDTFHL1 < FHL1_min_NRJ Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                '进行制热设备计算
                Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                'MsgBox("制热时间段内燃机(1)负荷率低于30%，被设置为0" & Chr(10) & "工况序号为：" & b)
            End If
        End If
        '制热内燃机2
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value > 0 Then
            ZRNRJDTFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value * ZRNRJZJSL2 '制热内燃机2单台负荷率
            If ZRNRJDTFHL2 < FHL1_min_NRJ Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                '进行制热设备计算
                Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                'MsgBox("制热时间段内燃机(2)负荷率低于30%，被设置为0" & Chr(10) & "工况序号为：" & b)
            End If
        End If
    End Sub
    Function 判断制冷制热设备选择是否正确(ExcelApp As Object, n As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '状态检测
        Dim ZTJC_EQ As Integer
        '定义局部变量
        Dim f3 = 0
        '判断制冷顺序一至顺序六的设备选择是否有重复，如果有重复则报错
        For f1 = 1 To n '工况序号
            For f2 = 6 To 10 '制冷设备顺序1至顺序5的列号,顺序6的设备只用来被比较
                If f2 = 6 Then
                    f3 = 5
                ElseIf f2 = 7 Then
                    f3 = 4
                ElseIf f2 = 8 Then
                    f3 = 3
                ElseIf f2 = 9 Then
                    f3 = 2
                ElseIf f2 = 10 Then
                    f3 = 1
                End If
                For f4 = 1 To f3
                    '判断制冷设备顺序1至顺序6是否有重复的，且都不为空
                    If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2 + f4).Value And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2).Value <> Nothing And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2 + f4).Value <> Nothing) Then
                        '报错
                        MsgBox("顺序一至顺序六的制冷设备类型不可以有重复，请重新选择设备类型！")
                        ZTJC_EQ = 1
                        Call 锁定工作表(ExcelApp)
                        GoTo qqq
                    End If
                Next
            Next
        Next
        '判断制热顺序一至顺序六的设备选择是否有重复，如果有重复则报错
        For f1 = 1 To n '工况序号
            For f2 = 15 To 19 '制热设备顺序1至顺序5的列号,顺序6的设备只用来被比较
                If f2 = 15 Then
                    f3 = 5
                ElseIf f2 = 16 Then
                    f3 = 4
                ElseIf f2 = 17 Then
                    f3 = 3
                ElseIf f2 = 18 Then
                    f3 = 2
                ElseIf f2 = 19 Then
                    f3 = 1
                End If
                For f4 = 1 To f3
                    '判断制热设备顺序1至顺序5是否有重复的，且都不为空
                    If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2 + f4).Value And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2).Value <> Nothing And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + f1, f2 + f4).Value <> Nothing) Then
                        '报错
                        MsgBox("顺序一至顺序六的制热设备类型不可以有重复，请重新选择设备类型！")
                        ZTJC_EQ = 1
                        Call 锁定工作表(ExcelApp)
                        GoTo qqq
                    End If
                Next
            Next
        Next
        '判断第一到第六顺序制热设备与梯级供热设备是否有重复
        For i = 1 To n '工况序号（行号）
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value <> Nothing Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Then '其中一个不为空
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value Then '第一到第六顺序中的某一个与梯级供热设备相同
                        MsgBox("顺序一至顺序六的制热设备类型不可以与梯级供热（仅计算耗电量不计算供热量）设备重复，请重新选择设备类型！")
                        ZTJC_EQ = 1
                        Call 锁定工作表(ExcelApp)
                        GoTo qqq
                    End If
                End If
            Next
        Next
        '如果有梯级供热设备，但是顺序1到6供热设备没有离心式热泵，报错
        For i = 1 To n '工况序号（行号）
            Dim TJGS_JC As Integer = 0 '梯级供热检测
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Then '存在梯级供热设备
                    '如果顺序1到6设备是空或者不是离心式热泵
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "离心式热泵" Then
                        TJGS_JC = TJGS_JC + 1
                    End If
                End If
            Next
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing And TJGS_JC = 0 Then '如果顺序1到6，都是空或者不是离心式热泵，说明没有选择离心式热泵，报错
                MsgBox("存在梯级供热设备，但是在顺序1到顺序6供热设备中没有选择离心式热泵，请检查并重新选择！")
                ZTJC_EQ = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
        Next
        '判断第一到第六顺序制热设备与混水供热设备是否有重复
        For i = 1 To n '工况序号（行号）
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value <> Nothing Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '其中一个不为空
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value Then '第一到第六顺序中的某一个与混水供热设备相同
                        MsgBox("顺序一至顺序六的制热设备类型不可以与混水供热设备重复，请重新选择设备类型！")
                        ZTJC_EQ = 1
                        Call 锁定工作表(ExcelApp)
                        GoTo qqq
                    End If
                End If
            Next
        Next
        '如果有混水供热设备，但是顺序1到6供热设备没有天然气锅炉或者直燃型溴化锂或者电采暖锅炉，报错
        For i = 1 To n '工况序号（行号）
            Dim HSGS_JC As Integer = 0 '混水供热检测
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '存在混水供热设备
                    '如果顺序1到6设备是空或者不是天然气锅炉、直燃型溴化锂、电采暖锅炉
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "天然气锅炉" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "直燃型溴化锂" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "电采暖锅炉" Then
                        HSGS_JC = HSGS_JC + 1
                    End If
                End If
            Next
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing And HSGS_JC = 0 Then '如果顺序1到6，都是空或者不是天然气锅炉、直燃型溴化锂、电采暖锅炉，说明没有选择天然气锅炉、直燃型溴化锂、电采暖锅炉，报错
                MsgBox("存在混水供热设备，但是在顺序1到顺序6供热设备中没有选择天然气锅炉或者直燃型溴化锂或者电采暖锅炉，请检查并重新选择！")
                ZTJC_EQ = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
        Next
        '判断存在混水供热设备，同时被混水设备为天然气锅炉或者直燃型溴化锂，此时是否有蓄热负荷需求，如果有，报错（天然气锅炉不用于蓄热，因为没有电价差优势）
        '判断第一到第六顺序制热设备与混水供热设备是否有重复
        For i = 1 To n '工况序号（行号）
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '存在混水供热设备
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "天然气锅炉" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "直燃型溴化锂" Then '供热设备为天然气锅炉或者直燃型溴化锂
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > 0 Then '存在蓄热负荷
                            MsgBox("存在混水供热设备，顺序1到顺序6供热设备中存在天然气锅炉或者直燃型溴化锂，同时此时有蓄热负荷，天然气锅炉或者直燃型溴化锂不可以用于蓄热计算，因为天然气没有谷电价格优势，请检查并重新选择！")
                            ZTJC_EQ = 1
                            Call 锁定工作表(ExcelApp)
                            GoTo qqq
                        End If
                    End If
                End If
            Next
        Next
        '如果存在混水供热设备，但是顺序1到顺序6同时存在了天然气锅炉、电锅炉、直燃型溴化锂中的不止一项，报错
        For i = 1 To n '工况序号（行号）
            '确定是否有某种设备
            Dim ZTJC_TRQGL As Integer = 0
            Dim ZTJC_DGL As Integer = 0
            Dim ZTJC_ZRXXHL As Integer = 0
            For j = 15 To 20 '第一到第六顺序制热设备（列号）
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '存在混水供热设备
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "天然气锅炉" Then
                        ZTJC_TRQGL = 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "电采暖锅炉" Then
                        ZTJC_DGL = 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, j).Value = "直燃型溴化锂" Then
                        ZTJC_ZRXXHL = 1
                    End If
                    If ZTJC_TRQGL + ZTJC_DGL + ZTJC_ZRXXHL > 1 Then
                        MsgBox("存在混水供热设备，顺序1到顺序6供热设备中同时存在天然气锅炉、直燃型溴化锂和电采暖锅炉中的不止一项，暂不支持这种计算，仅可以选择一种设备，请检查并重新选择！")
                        ZTJC_EQ = 1
                        Call 锁定工作表(ExcelApp)
                        GoTo qqq
                    End If
                End If
            Next
        Next
        '如果存在混水供热设备，但天然气锅炉、电锅炉、直燃型溴化锂装机选择不止一项，则报错
        For i = 1 To n '工况序号（行号）
            '确定是否有某种设备
            Dim ZTJC_TRQGL As Integer = 0
            Dim ZTJC_DGL As Integer = 0
            Dim ZTJC_ZRXXHL As Integer = 0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '存在混水供热设备
                '读取天然气锅炉、直燃型溴化锂机组、电采暖锅炉制热总功率（装机量，制热出力最大值）
                '读取采暖季装机方案及参数
                Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
                '天然气锅炉装机功率（总和）
                Dim TRQGL1ZRGL As Double = ans_ZJFA_R(12)
                Dim TRQGL2ZRGL As Double = ans_ZJFA_R(13)
                '电锅炉装机功率（总和）
                Dim DCNGL1ZRGL As Double = ans_ZJFA_R(20)
                Dim DCNGL2ZRGL As Double = ans_ZJFA_R(21)
                '直燃型溴化锂装机功率（总和）
                Dim ZRXHL1ZRGL As Double = ans_ZJFA_R(60)
                Dim ZRXHL2ZRGL As Double = ans_ZJFA_R(61)
                '——————————————————————————————————————————————————————————————————————————————————
                If TRQGL1ZRGL + TRQGL2ZRGL > 0 Then
                    ZTJC_TRQGL = 1
                End If
                If ZRXHL1ZRGL + ZRXHL2ZRGL > 0 Then
                    ZTJC_ZRXXHL = 1
                End If
                If DCNGL1ZRGL + DCNGL2ZRGL > 0 Then
                    ZTJC_DGL = 1
                End If
                If ZTJC_TRQGL + ZTJC_DGL + ZTJC_ZRXXHL > 1 Then
                    MsgBox("存在混水供热设备，同时装机方案中存在天然气锅炉、直燃型溴化锂和电采暖锅炉中的不止一项，暂不支持这种计算，仅可以选择一种设备，请检查并重新选择！")
                    ZTJC_EQ = 1
                    Call 锁定工作表(ExcelApp)
                    GoTo qqq
                End If
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
qqq:
        '返回结果
        Return ZTJC_EQ
    End Function
    Sub 判断各种计算结果是否正确(ExcelApp As Object, FHTJJD As Double, n As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '判断供冷计算结果是否都正确，如果出现不正确，则报错
        Dim GL_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 18).Value = "不正确" Then
                Dim XH_GL As String = "(" & i & ")"
                GL_ERROR = GL_ERROR & XH_GL & "  "
                'MsgBox("供冷的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)                
            End If
        Next
        '判断供热计算结果是否都正确，如果出现不正确，则报错
        Dim GR_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 35).Value = "不正确" Then
                Dim XH_GR As String = "(" & i & ")"
                GR_ERROR = GR_ERROR & XH_GR & "  "
                'MsgBox("供热的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
            End If
        Next
        '判断蓄冷计算结果是否都正确，如果出现不正确，则报错
        Dim XL_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 49).Value = "不正确" Then
                Dim XH_XL As String = "(" & i & ")"
                XL_ERROR = XL_ERROR & XH_XL & "  "
                'MsgBox("蓄冷的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
            End If
        Next
        '判断蓄热计算结果是否都正确，如果出现不正确，则报错
        Dim XR_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 61).Value = "不正确" Then
                Dim XH_XR As String = "(" & i & ")"
                XR_ERROR = XR_ERROR & XH_XR & "  "
                'MsgBox("蓄热的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
            End If
        Next
        '判断向外供电计算结果是否正确，如果出现不正确则报错
        '只有当处于不可以向外供电的计算模式，并且内燃机发电量大于总耗电量，才会报错
        '制冷季
        Dim GD_ERROR_L As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 62).Value = "不正确" Then
                Dim XH_GD_L As String = "(" & i & ")"
                GD_ERROR_L = GD_ERROR_L & XH_GD_L & "  "
                'MsgBox("供电的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
            End If
        Next
        '制热季
        Dim GD_ERROR_R As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 63).Value = "不正确" Then
                Dim XH_GD_R As String = "(" & i & ")"
                GD_ERROR_R = GD_ERROR_R & XH_GD_R & "  "
                'MsgBox("供电的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
            End If
        Next
        '检查天然气耗量计算结果中是否出现了不正确
        'For i = 17 To 29
        '    If ExcelApp.ThisWorkbook.Worksheets("成本测算").Cells(i, 2).Value = "不正确" Then
        '        MsgBox("内燃机天然气耗量计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & i)
        '        Exit For
        '    End If
        'Next
        '检查计算出来的各种设备负荷率是否小于0或者大于1，有则报错
        '内燃机负荷率
        Dim NRJ_ERROR_L_1 As String = Nothing
        Dim NRJ_ERROR_L_2 As String = Nothing
        Dim NRJ_ERROR_R_1 As String = Nothing
        Dim NRJ_ERROR_R_2 As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 2).Value > 1 Then
                Dim XH_NRJ_L_1 As String = "(" & i & ")"
                NRJ_ERROR_L_1 = NRJ_ERROR_L_1 & XH_NRJ_L_1 & "  "
                'MsgBox("内燃机负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 3).Value > 1 Then
                Dim XH_NRJ_L_2 As String = "(" & i & ")"
                NRJ_ERROR_L_2 = NRJ_ERROR_L_2 & XH_NRJ_L_2 & "  "
                'MsgBox("内燃机负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 19).Value > 1 Then
                Dim XH_NRJ_R_1 As String = "(" & i & ")"
                NRJ_ERROR_R_1 = NRJ_ERROR_R_1 & XH_NRJ_R_1 & "  "
                'MsgBox("内燃机负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 20).Value > 1 Then
                Dim XH_NRJ_R_2 As String = "(" & i & ")"
                NRJ_ERROR_R_2 = NRJ_ERROR_R_2 & XH_NRJ_R_2 & "  "
                'MsgBox("内燃机负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
        Next
        '制冷+蓄冷设备（上限设置，放大一些容错）
        Dim ZL_XL_FHL_ERROR As String = Nothing
        For i = 1 To n
            For j = 1 To 12
                If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 3 + j).Value + ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 35 + j).Value > (1 + 4 * FHTJJD / 100) Then
                    Dim XH_ZL_XL As String = "(" & i & ")"
                    ZL_XL_FHL_ERROR = ZL_XL_FHL_ERROR & XH_ZL_XL & "  "
                    'MsgBox("制冷设备和蓄冷设备负荷率之和的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
                End If
            Next
        Next
        '纯制冷设备（上限设置，放大一些容错）
        Dim ZL_FHL_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 16).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZL As String = "(" & i & ")"
                ZL_FHL_ERROR = ZL_FHL_ERROR & XH_ZL & "  "
                'MsgBox("制冷设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 17).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZL As String = "(" & i & ")"
                ZL_FHL_ERROR = ZL_FHL_ERROR & XH_ZL & "  "
                'MsgBox("制冷设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
        Next
        '制热+蓄热设备（上限设置，放大一些容错）
        Dim ZR_XR_FHL_ERROR As String = Nothing
        For i = 1 To n
            For j = 1 To 8
                If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 22 + j).Value + ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 49 + j).Value > (1 + 4 * FHTJJD / 100) Then
                    Dim XH_ZR_XR As String = "(" & i & ")"
                    ZR_XR_FHL_ERROR = ZR_XR_FHL_ERROR & XH_ZR_XR & "  "
                    'MsgBox("制热设备和蓄热设备负荷率之和的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
                End If
            Next
        Next
        For i = 1 To n
            For j = 1 To 2
                If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 32 + j).Value + ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 57 + j).Value > (1 + 4 * FHTJJD / 100) Then
                    Dim XH_ZR_XR As String = "(" & i & ")"
                    ZR_XR_FHL_ERROR = ZR_XR_FHL_ERROR & XH_ZR_XR & "  "
                    'MsgBox("制热设备和蓄热设备负荷率之和的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
                End If
            Next
        Next
        '纯制热设备（上限设置，放大一些容错）
        Dim ZR_FHL_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 21).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZR As String = "(" & i & ")"
                ZR_FHL_ERROR = ZR_FHL_ERROR & XH_ZR & "  "
                'MsgBox("制热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 22).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZR As String = "(" & i & ")"
                ZR_FHL_ERROR = ZR_FHL_ERROR & XH_ZR & "  "
                'MsgBox("制热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 31).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZR As String = "(" & i & ")"
                ZR_FHL_ERROR = ZR_FHL_ERROR & XH_ZR & "  "
                'MsgBox("制热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 32).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_ZR As String = "(" & i & ")"
                ZR_FHL_ERROR = ZR_FHL_ERROR & XH_ZR & "  "
                'MsgBox("制热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
        Next
        '梯级供热设备负荷率计算结果（上限设置，放大一些容错）
        Dim TJGR_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 81).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_TJGR As String = "(" & i & ")"
                TJGR_ERROR = TJGR_ERROR & XH_TJGR & "  "
                ' MsgBox("梯级供热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
        Next
        '混水供热设备负荷率计算结果（上限设置，放大一些容错）
        Dim HSGR_ERROR As String = Nothing
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value > (1 + 4 * FHTJJD / 100) Then
                Dim XH_HSGR As String = "(" & i & ")"
                HSGR_ERROR = HSGR_ERROR & XH_HSGR & "  "
                'MsgBox("混水供热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & i)
            End If
        Next
        '报错弹框
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        If GL_ERROR <> Nothing Then
            MsgBox("供冷的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & GL_ERROR)
        End If
        If GR_ERROR <> Nothing Then
            MsgBox("供热的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & GR_ERROR)
        End If
        If XL_ERROR <> Nothing Then
            MsgBox("蓄冷的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & XL_ERROR)
        End If
        If XR_ERROR <> Nothing Then
            MsgBox("蓄热的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & XR_ERROR)
        End If
        If GD_ERROR_L <> Nothing Then
            MsgBox("供冷季供电的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & GD_ERROR_L)
        End If
        If GD_ERROR_R <> Nothing Then
            MsgBox("采暖季供电的计算结果中出现了误差较大的不正确结果，请检查！！" & "不正确的工况序号为： " & GD_ERROR_R)
        End If
        If NRJ_ERROR_L_1 <> Nothing Then
            MsgBox("供冷季内燃机(1)负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & NRJ_ERROR_L_1)
        End If
        If NRJ_ERROR_L_2 <> Nothing Then
            MsgBox("供冷季内燃机(2)负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & NRJ_ERROR_L_2)
        End If
        If NRJ_ERROR_R_1 <> Nothing Then
            MsgBox("采暖季内燃机(1)负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & NRJ_ERROR_R_1)
        End If
        If NRJ_ERROR_R_2 <> Nothing Then
            MsgBox("采暖季内燃机(2)负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & NRJ_ERROR_R_2)
        End If
        If ZL_XL_FHL_ERROR <> Nothing Then
            MsgBox("制冷设备和蓄冷设备负荷率之和的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & ZL_XL_FHL_ERROR)
        End If
        If ZL_FHL_ERROR <> Nothing Then
            MsgBox("制冷设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & ZL_FHL_ERROR)
        End If
        If ZR_XR_FHL_ERROR <> Nothing Then
            MsgBox("制热设备和蓄热设备负荷率之和的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & ZR_XR_FHL_ERROR)
        End If
        If ZR_FHL_ERROR <> Nothing Then
            MsgBox("制热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & ZR_FHL_ERROR)
        End If
        If TJGR_ERROR <> Nothing Then
            MsgBox("梯级供热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & TJGR_ERROR)
        End If
        If HSGR_ERROR <> Nothing Then
            MsgBox("混水供热设备负荷率的计算结果出现了大于1的情况，计算结果不正确，请检查！" & "不正确的工况序号为： " & HSGR_ERROR)
        End If
    End Sub
    Function 判断冷热负荷需求量是否大于冷热负荷装机量(ExcelApp As Object, a As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '状态检测
        Dim ZTJC_LRFH As Integer = 0
        '定义局部变量
        Dim XHLZLPD As Double '溴化锂制冷功率
        Dim ZZLGL As Double '总制冷功率
        Dim XHLZRPD As Double '溴化锂制热功率
        Dim ZZRGL As Double '总制热功率
        'a表示当前正在计算的工况序号
        '————————————————————————————————————————————————————————————————————————————————————————
        '将内燃机余热利用方式带入
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value
        '计算制冷装机是否能够满足冷负荷需求量和蓄冷需求量
        '将制冷时内燃机负荷率全部设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0
        '读取溴化锂制冷功率
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value = "溴化锂" Then
            XHLZLPD = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 9).Value
        Else
            XHLZLPD = 0
        End If
        '顺序1到顺序6设备的制热功率
        Dim GLGL_1 As Double = 0
        Dim GLGL_2 As Double = 0
        Dim GLGL_3 As Double = 0
        Dim GLGL_4 As Double = 0
        Dim GLGL_5 As Double = 0
        Dim GLGL_6 As Double = 0
        '顺序1
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "离心式冷水机" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "空气源热泵" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "水冷螺杆机" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "水(地)源热泵" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "风冷螺杆机" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "离心式热泵" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 6).Value = "直燃型溴化锂" Then
            GLGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '顺序2
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "离心式冷水机" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "空气源热泵" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "水冷螺杆机" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "水(地)源热泵" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "风冷螺杆机" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "离心式热泵" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 7).Value = "直燃型溴化锂" Then
            GLGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '顺序3
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "离心式冷水机" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "空气源热泵" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "水冷螺杆机" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "水(地)源热泵" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "风冷螺杆机" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "离心式热泵" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 8).Value = "直燃型溴化锂" Then
            GLGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '顺序4
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "离心式冷水机" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "空气源热泵" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "水冷螺杆机" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "水(地)源热泵" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "风冷螺杆机" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "离心式热泵" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 9).Value = "直燃型溴化锂" Then
            GLGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '顺序5
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "离心式冷水机" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "空气源热泵" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "水冷螺杆机" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "水(地)源热泵" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "风冷螺杆机" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "离心式热泵" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 10).Value = "直燃型溴化锂" Then
            GLGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '顺序6
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "离心式冷水机" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "空气源热泵" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "水冷螺杆机" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "水(地)源热泵" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "风冷螺杆机" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "离心式热泵" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 11).Value = "直燃型溴化锂" Then
            GLGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '计算所有设备总制冷功率
        ZZLGL = XHLZLPD + XNGLGL(a) + GLGL_1 + GLGL_2 + GLGL_3 + GLGL_4 + GLGL_5 + GLGL_6
        If LFHZXQL(a) + XNXLGL(a) > ZZLGL * 1.005 Then '放大0.5%，防止极小偏差
            MsgBox("出现了冷负荷总需求量+蓄冷负荷需求量之和大于所有设备总制冷量的情况！" & Chr(10) & "工况序号为：" & a)
            ZTJC_LRFH = 1
            Call 锁定工作表(ExcelApp)
            GoTo qqqqq
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算制热装机是否能够满足热负荷需求量和蓄热需求量
        '将制热时内燃机负荷率全部设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 1
        '读取溴化锂制热功率
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value = "溴化锂" Then
            XHLZRPD = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 9).Value
        Else
            XHLZRPD = 0
        End If
        '顺序1到顺序6设备的制热功率
        Dim GRGL_1 As Double = 0
        Dim GRGL_2 As Double = 0
        Dim GRGL_3 As Double = 0
        Dim GRGL_4 As Double = 0
        Dim GRGL_5 As Double = 0
        Dim GRGL_6 As Double = 0
        '顺序1
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "空气源热泵" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "水(地)源热泵" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "风冷螺杆机" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "离心式热泵" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "天然气锅炉" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "直燃型溴化锂" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 15).Value = "电采暖锅炉" Then
            GRGL_1 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '顺序2
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "空气源热泵" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "水(地)源热泵" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "风冷螺杆机" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "离心式热泵" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "天然气锅炉" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "直燃型溴化锂" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 16).Value = "电采暖锅炉" Then
            GRGL_2 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '顺序3
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "空气源热泵" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "水(地)源热泵" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "风冷螺杆机" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "离心式热泵" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "天然气锅炉" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "直燃型溴化锂" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 17).Value = "电采暖锅炉" Then
            GRGL_3 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '顺序4
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "空气源热泵" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "水(地)源热泵" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "风冷螺杆机" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "离心式热泵" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "天然气锅炉" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "直燃型溴化锂" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 18).Value = "电采暖锅炉" Then
            GRGL_4 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '顺序5
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "空气源热泵" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "水(地)源热泵" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "风冷螺杆机" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "离心式热泵" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "天然气锅炉" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "直燃型溴化锂" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 19).Value = "电采暖锅炉" Then
            GRGL_5 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '顺序6
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "空气源热泵" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "水(地)源热泵" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "风冷螺杆机" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "离心式热泵" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "天然气锅炉" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "直燃型溴化锂" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 20).Value = "电采暖锅炉" Then
            GRGL_6 = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
        Dim HSGRGL As Double = 0
        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 82).Value = "空气源热泵" Then
            HSGRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 82).Value = "水(地)源热泵" Then
            HSGRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 82).Value = "风冷螺杆机" Then
            HSGRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        '计算所有设备总制热功率
        ZZRGL = XHLZRPD + XNGRGL(a) + GRGL_1 + GRGL_2 + GRGL_3 + GRGL_4 + GRGL_5 + GRGL_6 + HSGRGL
        If RFHZXQL(a) + XNXRGL(a) > ZZRGL * 1.005 Then '放大0.5%，防止极小偏差
            MsgBox("出现了热负荷总需求量+蓄热负荷需求量之和大于所有设备总制热量的情况！" & Chr(10) & "工况序号为：" & a)
            ZTJC_LRFH = 1
            Call 锁定工作表(ExcelApp)
            GoTo qqqqq
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
qqqqq:
        '清空数据
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B3:CO3").ClearContents
        '返回的结果列表
        Dim ans(4)
        ans(0) = ZTJC_LRFH
        ans(1) = ZZRGL
        ans(2) = XHLZRPD + GRGL_1 + GRGL_2 + GRGL_3 + GRGL_4 + GRGL_5 + GRGL_6
        ans(3) = GRGL_1 + GRGL_2 + GRGL_3 + GRGL_4 + GRGL_5 + GRGL_6
        ans(4) = HSGRGL
        '返回状态监测结果
        Return ans
    End Function

    Sub 制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp As Object, b As Integer, FHTJJD As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '只有全局寻优计算模式才进行修正
        If calculation_mode = 2 And LFHZXQL(b) + XNXLGL(b) > 0 Then
            '读取制冷季装机方案及参数
            Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
            '内燃机
            Dim NUM1_NRJ As Double = ans_ZJFA_L(0)
            Dim NUM2_NRJ As Double = ans_ZJFA_L(1)
            Dim FDGL1_ED_NRJ As Double = ans_ZJFA_L(2)
            Dim FDGL2_ED_NRJ As Double = ans_ZJFA_L(3)
            Dim YRGL1_ED_NRJ As Double = ans_ZJFA_L(4)
            Dim YRGL2_ED_NRJ As Double = ans_ZJFA_L(5)
            Dim TRQ1_ED_NRJ As Double = ans_ZJFA_L(6)
            Dim TRQ2_ED_NRJ As Double = ans_ZJFA_L(7)
            Dim FJHD1_ED_NRJ As Double = ans_ZJFA_L(8)
            Dim FJHD2_ED_NRJ As Double = ans_ZJFA_L(9)
            '离心式冷水机
            Dim NUM1_LXSLSJ As Double = ans_ZJFA_L(10)
            Dim NUM2_LXSLSJ As Double = ans_ZJFA_L(11)
            Dim ZJLGL1_LXSLSJ As Double = ans_ZJFA_L(12)
            Dim ZJLGL2_LXSLSJ As Double = ans_ZJFA_L(13)
            Dim BTHD1_ED_LXSLSJ As Double = ans_ZJFA_L(14)
            Dim BTHD2_ED_LXSLSJ As Double = ans_ZJFA_L(15)
            Dim FJHD1_ED_LXSLSJ As Double = ans_ZJFA_L(16)
            Dim FJHD2_ED_LXSLSJ As Double = ans_ZJFA_L(17)
            '水冷螺杆机
            Dim NUM1_SLLGJ As Double = ans_ZJFA_L(18)
            Dim NUM2_SLLGJ As Double = ans_ZJFA_L(19)
            Dim ZJLGL1_SLLGJ As Double = ans_ZJFA_L(20)
            Dim ZJLGL2_SLLGJ As Double = ans_ZJFA_L(21)
            Dim BTHD1_ED_SLLGJ As Double = ans_ZJFA_L(22)
            Dim BTHD2_ED_SLLGJ As Double = ans_ZJFA_L(23)
            Dim FJHD1_ED_SLLGJ As Double = ans_ZJFA_L(24)
            Dim FJHD2_ED_SLLGJ As Double = ans_ZJFA_L(25)
            '风冷螺杆机
            Dim NUM1_FLLGJ As Double = ans_ZJFA_L(26)
            Dim NUM2_FLLGJ As Double = ans_ZJFA_L(27)
            Dim ZJLGL1_FLLGJ As Double = ans_ZJFA_L(28)
            Dim ZJLGL2_FLLGJ As Double = ans_ZJFA_L(29)
            Dim BTHD1_ED_FLLGJ As Double = ans_ZJFA_L(30)
            Dim BTHD2_ED_FLLGJ As Double = ans_ZJFA_L(31)
            Dim FJHD1_ED_FLLGJ As Double = ans_ZJFA_L(32)
            Dim FJHD2_ED_FLLGJ As Double = ans_ZJFA_L(33)
            '水地源热泵
            Dim NUM1_SDYRB As Double = ans_ZJFA_L(34)
            Dim NUM2_SDYRB As Double = ans_ZJFA_L(35)
            Dim ZJLGL1_SDYRB As Double = ans_ZJFA_L(36)
            Dim ZJLGL2_SDYRB As Double = ans_ZJFA_L(37)
            Dim BTHD1_ED_SDYRB As Double = ans_ZJFA_L(38)
            Dim BTHD2_ED_SDYRB As Double = ans_ZJFA_L(39)
            Dim FJHD1_ED_SDYRB As Double = ans_ZJFA_L(40)
            Dim FJHD2_ED_SDYRB As Double = ans_ZJFA_L(41)
            '离心式热泵
            Dim NUM1_LXSRB As Double = ans_ZJFA_L(42)
            Dim NUM2_LXSRB As Double = ans_ZJFA_L(43)
            Dim ZJLGL1_LXSRB As Double = ans_ZJFA_L(44)
            Dim ZJLGL2_LXSRB As Double = ans_ZJFA_L(45)
            Dim BTHD1_ED_LXSRB As Double = ans_ZJFA_L(46)
            Dim BTHD2_ED_LXSRB As Double = ans_ZJFA_L(47)
            Dim FJHD1_ED_LXSRB As Double = ans_ZJFA_L(48)
            Dim FJHD2_ED_LXSRB As Double = ans_ZJFA_L(49)
            '空气源热泵
            Dim NUM1_KQYRB As Double = ans_ZJFA_L(50)
            Dim NUM2_KQYRB As Double = ans_ZJFA_L(51)
            Dim ZJLGL1_KQYRB As Double = ans_ZJFA_L(52)
            Dim ZJLGL2_KQYRB As Double = ans_ZJFA_L(53)
            Dim BTHD1_ED_KQYRB As Double = ans_ZJFA_L(54)
            Dim BTHD2_ED_KQYRB As Double = ans_ZJFA_L(55)
            Dim FJHD1_ED_KQYRB As Double = ans_ZJFA_L(56)
            Dim FJHD2_ED_KQYRB As Double = ans_ZJFA_L(57)
            '直燃型溴化锂
            Dim NUM1_ZRXXHL As Double = ans_ZJFA_L(58)
            Dim NUM2_ZRXXHL As Double = ans_ZJFA_L(59)
            Dim ZJLGL1_ZRXXHL As Double = ans_ZJFA_L(60)
            Dim ZJLGL2_ZRXXHL As Double = ans_ZJFA_L(61)
            Dim BTHQ1_ED_ZRXXHL As Double = ans_ZJFA_L(62)
            Dim BTHQ2_ED_ZRXXHL As Double = ans_ZJFA_L(63)
            Dim FJHD1_ED_ZRXXHL As Double = ans_ZJFA_L(64)
            Dim FJHD2_ED_ZRXXHL As Double = ans_ZJFA_L(65)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '读取各种修正系数
            Dim ans_XZXS_L = 读取制冷季输入的计算系数(ExcelApp)
            Dim BTHDXS_GL_water As Double = ans_XZXS_L(0)
            Dim BTHDXS_XL_water As Double = ans_XZXS_L(1)
            Dim BTHDXS_GL_air As Double = ans_XZXS_L(2)
            Dim BTHDXS_XL_air As Double = ans_XZXS_L(3)
            Dim FJHDXS As Double = ans_XZXS_L(4)
            Dim ZRXXHLZLHQXZ As Double = ans_XZXS_L(5)
            Dim TRQHLXZXS_QT As Double = ans_XZXS_L(6)
            Dim TRQHLXZXS_NRJ As Double = ans_XZXS_L(7)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '————————————————————————————————————————————————————————————————————————————————————————  
            '负荷调整系数（全局寻优时候能否计算到了负荷上限的倍数）
            Dim TZXS As Double = 1 + 10 * FHTJJD / 100
            '各个设备可以允许运行的最低负荷率
            '当设备不存在（数量=0时），负荷率下限设置为1，加快计算
            '内燃机
            '内燃机及其余热利用系统最低允许负荷率
            Dim NRJFHL1_min As Double
            Dim NRJFHL2_min As Double
            If NUM1_NRJ > 0 Then
                NRJFHL1_min = FHL1_min_NRJ / NUM1_NRJ
            Else
                NRJFHL1_min = 0
            End If
            If NUM2_NRJ > 0 Then
                NRJFHL2_min = FHL2_min_NRJ / NUM2_NRJ
            Else
                NRJFHL2_min = 0
            End If
            '离心式冷水机
            '设备最低允许运行的负荷率
            Dim LXSLSJFHL1_min As Double
            Dim LXSLSJFHL2_min As Double
            If NUM1_LXSLSJ > 0 Then
                LXSLSJFHL1_min = FHL1_min_LXSLSJ / NUM1_LXSLSJ
            Else
                LXSLSJFHL1_min = 0
            End If
            If NUM2_LXSLSJ > 0 Then
                LXSLSJFHL2_min = FHL2_min_LXSLSJ / NUM2_LXSLSJ
            Else
                LXSLSJFHL2_min = 0
            End If
            '水冷螺杆机
            '设备最低允许运行的负荷率
            Dim SLLGJFHL1_min As Double
            Dim SLLGJFHL2_min As Double
            If NUM1_SLLGJ > 0 Then
                SLLGJFHL1_min = FHL1_min_SLLGJ / NUM1_SLLGJ
            Else
                SLLGJFHL1_min = 0
            End If
            If NUM2_SLLGJ > 0 Then
                SLLGJFHL2_min = FHL2_min_SLLGJ / NUM2_SLLGJ
            Else
                SLLGJFHL2_min = 0
            End If
            '风冷螺杆机
            '设备最低允许运行的负荷率
            Dim FLLGJFHL1_min As Double
            Dim FLLGJFHL2_min As Double
            If NUM1_FLLGJ > 0 Then
                FLLGJFHL1_min = FHL1_min_FLLGJ / NUM1_FLLGJ
            Else
                FLLGJFHL1_min = 0
            End If
            If NUM2_FLLGJ > 0 Then
                FLLGJFHL2_min = FHL2_min_FLLGJ / NUM2_FLLGJ
            Else
                FLLGJFHL2_min = 0
            End If
            '水（地）源热泵
            '设备最低允许运行的负荷率
            Dim SDYRBFHL1_min As Double
            Dim SDYRBFHL2_min As Double
            If NUM1_SDYRB > 0 Then
                SDYRBFHL1_min = FHL1_min_SDYRB / NUM1_SDYRB
            Else
                SDYRBFHL1_min = 0
            End If
            If NUM2_SDYRB > 0 Then
                SDYRBFHL2_min = FHL2_min_SDYRB / NUM2_SDYRB
            Else
                SDYRBFHL2_min = 0
            End If
            '离心式热泵
            '设备最低允许运行的负荷率
            Dim LXSRBFHL1_min As Double
            Dim LXSRBFHL2_min As Double
            If NUM1_LXSRB > 0 Then
                LXSRBFHL1_min = FHL1_min_LXSRB / NUM1_LXSRB
            Else
                LXSRBFHL1_min = 0
            End If
            If NUM2_LXSRB > 0 Then
                LXSRBFHL2_min = FHL2_min_LXSRB / NUM2_LXSRB
            Else
                LXSRBFHL2_min = 0
            End If
            '空气源热泵
            '设备最低允许运行的负荷率
            Dim KQYRBFHL1_min As Double
            Dim KQYRBFHL2_min As Double
            If NUM1_KQYRB > 0 Then
                KQYRBFHL1_min = FHL1_min_KQYRB / NUM1_KQYRB
            Else
                KQYRBFHL1_min = 0
            End If
            If NUM2_KQYRB > 0 Then
                KQYRBFHL2_min = FHL2_min_KQYRB / NUM2_KQYRB
            Else
                KQYRBFHL2_min = 0
            End If
            '直燃型溴化锂
            '设备最低允许运行的负荷率
            Dim ZRXXHLFHL1_min As Double
            Dim ZRXXHLFHL2_min As Double
            If NUM1_ZRXXHL > 0 Then
                ZRXXHLFHL1_min = FHL1_min_ZRXXHL / NUM1_ZRXXHL
            Else
                ZRXXHLFHL1_min = 0
            End If
            If NUM2_ZRXXHL > 0 Then
                ZRXXHLFHL2_min = FHL2_min_ZRXXHL / NUM2_ZRXXHL
            Else
                ZRXXHLFHL2_min = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '各空调设备本体耗电功率在考虑修正系数前的耗电功率和考虑修正系数后的耗电系数
            '离心式冷水机
            Dim LXSLSJBTHD1_ALL_a As Double
            Dim LXSLSJBTHD1_ALL_b As Double
            Dim LXSLSJBTHD2_ALL_a As Double
            Dim LXSLSJBTHD2_ALL_b As Double
            '水冷螺杆机
            Dim SLLGJBTHD1_ALL_a As Double
            Dim SLLGJBTHD1_ALL_b As Double
            Dim SLLGJBTHD2_ALL_a As Double
            Dim SLLGJBTHD2_ALL_b As Double
            '风冷螺杆机
            Dim FLLGJBTHD1_ALL_a As Double
            Dim FLLGJBTHD1_ALL_b As Double
            Dim FLLGJBTHD2_ALL_a As Double
            Dim FLLGJBTHD2_ALL_b As Double
            '水（地）源热泵
            Dim SDYRBBTHD1_ALL_a As Double
            Dim SDYRBBTHD1_ALL_b As Double
            Dim SDYRBBTHD2_ALL_a As Double
            Dim SDYRBBTHD2_ALL_b As Double
            '离心式热泵
            Dim LXSRBBTHD1_ALL_a As Double
            Dim LXSRBBTHD1_ALL_b As Double
            Dim LXSRBBTHD2_ALL_a As Double
            Dim LXSRBBTHD2_ALL_b As Double
            '空气源热泵
            Dim KQYRBBTHD1_ALL_a As Double
            Dim KQYRBBTHD1_ALL_b As Double
            Dim KQYRBBTHD2_ALL_a As Double
            Dim KQYRBBTHD2_ALL_b As Double
            '各空调设备本体天然气耗量在考虑修正系数前的天然气耗量和考虑修正系数后的天然气耗量
            '直燃型溴化锂
            Dim ZRXXHLBTTRQ1_ALL_a As Double
            Dim ZRXXHLBTTRQ1_ALL_b As Double
            Dim ZRXXHLBTTRQ2_ALL_a As Double
            Dim ZRXXHLBTTRQ2_ALL_b As Double
            '内燃发电机
            Dim NRJTRQ1_ALL_a As Double
            Dim NRJTRQ1_ALL_b As Double
            Dim NRJTRQ2_ALL_a As Double
            Dim NRJTRQ2_ALL_b As Double
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '离心式冷水机（1）
            '设备的装机总数量，台数
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSLSJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value
            Dim LXSLSJFHL1_XL As Double
            If NUM1_LXSLSJ = 0 Then
                LXSLSJFHL1_XL = 0
            Else
                LXSLSJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value / NUM1_LXSLSJ
            End If
            Dim LXSLSJFHL1 As Double = LXSLSJFHL1_GL + LXSLSJFHL1_XL
            If NUM1_LXSLSJ > 0 And LXSLSJFHL1 >= LXSLSJFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSLSJBTHD1_ED As Double = NUM1_LXSLSJ * BTHD1_ED_LXSLSJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSLSJBTHD1_GL As Double = BTHDXS_GL_water * LXSLSJBTHD1_ED * LXSLSJFHL1_GL
                Dim LXSLSJBTHD1_XL As Double = BTHDXS_XL_water * LXSLSJBTHD1_ED * LXSLSJFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                LXSLSJBTHD1_ALL_a = LXSLSJBTHD1_GL + LXSLSJBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_LXSLSJ_List As New List(Of Double)
                Dim FHL1_LXSLSJ_single As New List(Of Double)
                Dim NUM1_LXSLSJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_LXSLSJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_LXSLSJ / NUM1_LXSLSJ < LXSLSJFHL1 * ZJLGL1_LXSLSJ Then
                        GoTo aaa
                    End If
                    For a1 = FHL1_min_LXSLSJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_LXSLSJ / NUM1_LXSLSJ >= LXSLSJFHL1 * ZJLGL1_LXSLSJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSLSJ_List.Add(离心式冷水机制冷COP曲线(a1))
                            FHL1_LXSLSJ_single.Add(a1)
                            NUM1_LXSLSJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo aaa
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSLSJ_List.Add(离心式冷水机制冷COP曲线(a1))
                            FHL1_LXSLSJ_single.Add(a1)
                            NUM1_LXSLSJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo aaa
                        End If
                    Next
aaa:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim LXSLSJZLCOPXZ1 As Double = COPXZ1_LXSLSJ_List.Max
                '经过修正后的设备本体耗电功率
                LXSLSJBTHD1_ALL_b = LXSLSJBTHD1_ALL_a / LXSLSJZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_LXSLSJ_List.IndexOf(LXSLSJZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 112).Value = NUM1_LXSLSJ_QD(COP_MAX_index)
            Else
                LXSLSJBTHD1_ALL_a = 0
                LXSLSJBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 112).Value = 0
            End If
            '离心式冷水机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSLSJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value
            Dim LXSLSJFHL2_XL As Double
            If NUM2_LXSLSJ = 0 Then
                LXSLSJFHL2_XL = 0
            Else
                LXSLSJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value / NUM2_LXSLSJ
            End If
            Dim LXSLSJFHL2 As Double = LXSLSJFHL2_GL + LXSLSJFHL2_XL
            If NUM2_LXSLSJ > 0 And LXSLSJFHL2 >= LXSLSJFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSLSJBTHD2_ED As Double = NUM2_LXSLSJ * BTHD2_ED_LXSLSJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSLSJBTHD2_GL As Double = BTHDXS_GL_water * LXSLSJBTHD2_ED * LXSLSJFHL2_GL
                Dim LXSLSJBTHD2_XL As Double = BTHDXS_XL_water * LXSLSJBTHD2_ED * LXSLSJFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                LXSLSJBTHD2_ALL_a = LXSLSJBTHD2_GL + LXSLSJBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_LXSLSJ_List As New List(Of Double)
                Dim FHL2_LXSLSJ_single As New List(Of Double)
                Dim NUM2_LXSLSJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_LXSLSJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_LXSLSJ / NUM2_LXSLSJ < LXSLSJFHL2 * ZJLGL2_LXSLSJ Then
                        GoTo bbb
                    End If
                    For a2 = FHL2_min_LXSLSJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_LXSLSJ / NUM2_LXSLSJ >= LXSLSJFHL2 * ZJLGL2_LXSLSJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSLSJ_List.Add(离心式冷水机制冷COP曲线(a2))
                            FHL2_LXSLSJ_single.Add(a2)
                            NUM2_LXSLSJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo bbb
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSLSJ_List.Add(离心式冷水机制冷COP曲线(a2))
                            FHL2_LXSLSJ_single.Add(a2)
                            NUM2_LXSLSJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo bbb
                        End If
                    Next
bbb:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim LXSLSJZLCOPXZ2 As Double = COPXZ2_LXSLSJ_List.Max
                '经过修正后的设备本体耗电功率
                LXSLSJBTHD2_ALL_b = LXSLSJBTHD2_ALL_a / LXSLSJZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_LXSLSJ_List.IndexOf(LXSLSJZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 113).Value = NUM2_LXSLSJ_QD(COP_MAX_index)
            Else
                LXSLSJBTHD2_ALL_a = 0
                LXSLSJBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 113).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '水冷螺杆机（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SLLGJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value
            Dim SLLGJFHL1_XL As Double
            If NUM1_SLLGJ = 0 Then
                SLLGJFHL1_XL = 0
            Else
                SLLGJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value / NUM1_SLLGJ
            End If
            Dim SLLGJFHL1 As Double = SLLGJFHL1_GL + SLLGJFHL1_XL
            If NUM1_SLLGJ > 0 And SLLGJFHL1 >= SLLGJFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SLLGJBTHD1_ED As Double = NUM1_SLLGJ * BTHD1_ED_SLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SLLGJBTHD1_GL As Double = BTHDXS_GL_water * SLLGJBTHD1_ED * SLLGJFHL1_GL
                Dim SLLGJBTHD1_XL As Double = BTHDXS_XL_water * SLLGJBTHD1_ED * SLLGJFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                SLLGJBTHD1_ALL_a = SLLGJBTHD1_GL + SLLGJBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_SLLGJ_List As New List(Of Double)
                Dim FHL1_SLLGJ_single As New List(Of Double)
                Dim NUM1_SLLGJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_SLLGJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_SLLGJ / NUM1_SLLGJ < SLLGJFHL1 * ZJLGL1_SLLGJ Then
                        GoTo ccc
                    End If
                    For a1 = FHL1_min_SLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_SLLGJ / NUM1_SLLGJ >= SLLGJFHL1 * ZJLGL1_SLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SLLGJ_List.Add(水冷螺杆机制冷COP曲线(a1))
                            FHL1_SLLGJ_single.Add(a1)
                            NUM1_SLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ccc
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SLLGJ_List.Add(水冷螺杆机制冷COP曲线(a1))
                            FHL1_SLLGJ_single.Add(a1)
                            NUM1_SLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ccc
                        End If
                    Next
ccc:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim SLLGJZLCOPXZ1 As Double = COPXZ1_SLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                SLLGJBTHD1_ALL_b = SLLGJBTHD1_ALL_a / SLLGJZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_SLLGJ_List.IndexOf(SLLGJZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 120).Value = NUM1_SLLGJ_QD(COP_MAX_index)
            Else
                SLLGJBTHD1_ALL_a = 0
                SLLGJBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 120).Value = 0
            End If
            '水冷螺杆机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SLLGJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value
            Dim SLLGJFHL2_XL As Double
            If NUM2_SLLGJ = 0 Then
                SLLGJFHL2_XL = 0
            Else
                SLLGJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value / NUM2_SLLGJ
            End If
            Dim SLLGJFHL2 As Double = SLLGJFHL2_GL + SLLGJFHL2_XL
            If NUM2_SLLGJ > 0 And SLLGJFHL2 >= SLLGJFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SLLGJBTHD2_ED As Double = NUM2_SLLGJ * BTHD2_ED_SLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SLLGJBTHD2_GL As Double = BTHDXS_GL_water * SLLGJBTHD2_ED * SLLGJFHL2_GL
                Dim SLLGJBTHD2_XL As Double = BTHDXS_XL_water * SLLGJBTHD2_ED * SLLGJFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                SLLGJBTHD2_ALL_a = SLLGJBTHD2_GL + SLLGJBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_SLLGJ_List As New List(Of Double)
                Dim FHL2_SLLGJ_single As New List(Of Double)
                Dim NUM2_SLLGJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_SLLGJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_SLLGJ / NUM2_SLLGJ < SLLGJFHL2 * ZJLGL2_SLLGJ Then
                        GoTo ddd
                    End If
                    For a2 = FHL2_min_SLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_SLLGJ / NUM2_SLLGJ >= SLLGJFHL2 * ZJLGL2_SLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SLLGJ_List.Add(水冷螺杆机制冷COP曲线(a2))
                            FHL2_SLLGJ_single.Add(a2)
                            NUM2_SLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ddd
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SLLGJ_List.Add(水冷螺杆机制冷COP曲线(a2))
                            FHL2_SLLGJ_single.Add(a2)
                            NUM2_SLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ddd
                        End If
                    Next
ddd:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim SLLGJZLCOPXZ2 As Double = COPXZ2_SLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                SLLGJBTHD2_ALL_b = SLLGJBTHD2_ALL_a / SLLGJZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_SLLGJ_List.IndexOf(SLLGJZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 121).Value = NUM2_SLLGJ_QD(COP_MAX_index)
            Else
                SLLGJBTHD2_ALL_a = 0
                SLLGJBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 121).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '风冷螺杆机（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim FLLGJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value
            Dim FLLGJFHL1_XL As Double
            If NUM1_FLLGJ = 0 Then
                FLLGJFHL1_XL = 0
            Else
                FLLGJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value / NUM1_FLLGJ
            End If
            Dim FLLGJFHL1 As Double = FLLGJFHL1_GL + FLLGJFHL1_XL
            If NUM1_FLLGJ > 0 And FLLGJFHL1 >= FLLGJFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim FLLGJBTHD1_ED As Double = NUM1_FLLGJ * BTHD1_ED_FLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim FLLGJBTHD1_GL As Double = BTHDXS_GL_air * FLLGJBTHD1_ED * FLLGJFHL1_GL
                Dim FLLGJBTHD1_XL As Double = BTHDXS_XL_air * FLLGJBTHD1_ED * FLLGJFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                FLLGJBTHD1_ALL_a = FLLGJBTHD1_GL + FLLGJBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_FLLGJ_List As New List(Of Double)
                Dim FHL1_FLLGJ_single As New List(Of Double)
                Dim NUM1_FLLGJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_FLLGJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_FLLGJ / NUM1_FLLGJ < FLLGJFHL1 * ZJLGL1_FLLGJ Then
                        GoTo eee
                    End If
                    For a1 = FHL1_min_FLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_FLLGJ / NUM1_FLLGJ >= FLLGJFHL1 * ZJLGL1_FLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_FLLGJ_List.Add(风冷螺杆式热泵制冷COP曲线(a1))
                            FHL1_FLLGJ_single.Add(a1)
                            NUM1_FLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo eee
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_FLLGJ_List.Add(风冷螺杆式热泵制冷COP曲线(a1))
                            FHL1_FLLGJ_single.Add(a1)
                            NUM1_FLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo eee
                        End If
                    Next
eee:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim FLLGJZLCOPXZ1 As Double = COPXZ1_FLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                FLLGJBTHD1_ALL_b = FLLGJBTHD1_ALL_a / FLLGJZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_FLLGJ_List.IndexOf(FLLGJZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 114).Value = NUM1_FLLGJ_QD(COP_MAX_index)
            Else
                FLLGJBTHD1_ALL_a = 0
                FLLGJBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 114).Value = 0
            End If
            '风冷螺杆机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim FLLGJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value
            Dim FLLGJFHL2_XL As Double
            If NUM2_FLLGJ = 0 Then
                FLLGJFHL2_XL = 0
            Else
                FLLGJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value / NUM2_FLLGJ
            End If
            Dim FLLGJFHL2 As Double = FLLGJFHL2_GL + FLLGJFHL2_XL
            If NUM2_FLLGJ > 0 And FLLGJFHL2 >= FLLGJFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim FLLGJBTHD2_ED As Double = NUM2_FLLGJ * BTHD2_ED_FLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim FLLGJBTHD2_GL As Double = BTHDXS_GL_air * FLLGJBTHD2_ED * FLLGJFHL2_GL
                Dim FLLGJBTHD2_XL As Double = BTHDXS_XL_air * FLLGJBTHD2_ED * FLLGJFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                FLLGJBTHD2_ALL_a = FLLGJBTHD2_GL + FLLGJBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_FLLGJ_List As New List(Of Double)
                Dim FHL2_FLLGJ_single As New List(Of Double)
                Dim NUM2_FLLGJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_FLLGJ Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_FLLGJ / NUM2_FLLGJ < FLLGJFHL2 * ZJLGL2_FLLGJ Then
                        GoTo fff
                    End If
                    For a2 = FHL2_min_FLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_FLLGJ / NUM2_FLLGJ >= FLLGJFHL2 * ZJLGL2_FLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_FLLGJ_List.Add(风冷螺杆式热泵制冷COP曲线(a2))
                            FHL2_FLLGJ_single.Add(a2)
                            NUM2_FLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo fff
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_FLLGJ_List.Add(风冷螺杆式热泵制冷COP曲线(a2))
                            FHL2_FLLGJ_single.Add(a2)
                            NUM2_FLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo fff
                        End If
                    Next
fff:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim FLLGJZLCOPXZ2 As Double = COPXZ2_FLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                FLLGJBTHD2_ALL_b = FLLGJBTHD2_ALL_a / FLLGJZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_FLLGJ_List.IndexOf(FLLGJZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 115).Value = NUM2_FLLGJ_QD(COP_MAX_index)
            Else
                FLLGJBTHD2_ALL_a = 0
                FLLGJBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 115).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '水（地）源热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SDYRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value
            Dim SDYRBFHL1_XL As Double
            If NUM1_SDYRB = 0 Then
                SDYRBFHL1_XL = 0
            Else
                SDYRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value / NUM1_SDYRB
            End If
            Dim SDYRBFHL1 As Double = SDYRBFHL1_GL + SDYRBFHL1_XL
            If NUM1_SDYRB > 0 And SDYRBFHL1 >= SDYRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SDYRBBTHD1_ED As Double = NUM1_SDYRB * BTHD1_ED_SDYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SDYRBBTHD1_GL As Double = BTHDXS_GL_water * SDYRBBTHD1_ED * SDYRBFHL1_GL
                Dim SDYRBBTHD1_XL As Double = BTHDXS_XL_water * SDYRBBTHD1_ED * SDYRBFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                SDYRBBTHD1_ALL_a = SDYRBBTHD1_GL + SDYRBBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_SDYRB_List As New List(Of Double)
                Dim FHL1_SDYRB_single As New List(Of Double)
                Dim NUM1_SDYRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_SDYRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_SDYRB / NUM1_SDYRB < SDYRBFHL1 * ZJLGL1_SDYRB Then
                        GoTo ggg
                    End If
                    For a1 = FHL1_min_SDYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_SDYRB / NUM1_SDYRB >= SDYRBFHL1 * ZJLGL1_SDYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SDYRB_List.Add(水_地源热泵制冷COP曲线(a1))
                            FHL1_SDYRB_single.Add(a1)
                            NUM1_SDYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ggg
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SDYRB_List.Add(水_地源热泵制冷COP曲线(a1))
                            FHL1_SDYRB_single.Add(a1)
                            NUM1_SDYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ggg
                        End If
                    Next
ggg:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim SDYRBZLCOPXZ1 As Double = COPXZ1_SDYRB_List.Max
                '经过修正后的设备本体耗电功率
                SDYRBBTHD1_ALL_b = SDYRBBTHD1_ALL_a / SDYRBZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_SDYRB_List.IndexOf(SDYRBZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 118).Value = NUM1_SDYRB_QD(COP_MAX_index)
            Else
                SDYRBBTHD1_ALL_a = 0
                SDYRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 118).Value = 0
            End If
            '水（地）源热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SDYRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value
            Dim SDYRBFHL2_XL As Double
            If NUM2_SDYRB = 0 Then
                SDYRBFHL2_XL = 0
            Else
                SDYRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value / NUM2_SDYRB
            End If
            Dim SDYRBFHL2 As Double = SDYRBFHL2_GL + SDYRBFHL2_XL
            If NUM2_SDYRB > 0 And SDYRBFHL2 >= SDYRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SDYRBBTHD2_ED As Double = NUM2_SDYRB * BTHD2_ED_SDYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SDYRBBTHD2_GL As Double = BTHDXS_GL_water * SDYRBBTHD2_ED * SDYRBFHL2_GL
                Dim SDYRBBTHD2_XL As Double = BTHDXS_XL_water * SDYRBBTHD2_ED * SDYRBFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                SDYRBBTHD2_ALL_a = SDYRBBTHD2_GL + SDYRBBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_SDYRB_List As New List(Of Double)
                Dim FHL2_SDYRB_single As New List(Of Double)
                Dim NUM2_SDYRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_SDYRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_SDYRB / NUM2_SDYRB < SDYRBFHL2 * ZJLGL2_SDYRB Then
                        GoTo hhh
                    End If
                    For a2 = FHL2_min_SDYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_SDYRB / NUM2_SDYRB >= SDYRBFHL2 * ZJLGL2_SDYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SDYRB_List.Add(水_地源热泵制冷COP曲线(a2))
                            FHL2_SDYRB_single.Add(a2)
                            NUM2_SDYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo hhh
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SDYRB_List.Add(水_地源热泵制冷COP曲线(a2))
                            FHL2_SDYRB_single.Add(a2)
                            NUM2_SDYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo hhh
                        End If
                    Next
hhh:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim SDYRBZLCOPXZ2 As Double = COPXZ2_SDYRB_List.Max
                '经过修正后的设备本体耗电功率
                SDYRBBTHD2_ALL_b = SDYRBBTHD2_ALL_a / SDYRBZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_SDYRB_List.IndexOf(SDYRBZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 119).Value = NUM2_SDYRB_QD(COP_MAX_index)
            Else
                SDYRBBTHD2_ALL_a = 0
                SDYRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 119).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '离心式热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value
            Dim LXSRBFHL1_XL As Double
            If NUM1_LXSRB = 0 Then
                LXSRBFHL1_XL = 0
            Else
                LXSRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value / NUM1_LXSRB
            End If
            Dim LXSRBFHL1 As Double = LXSRBFHL1_GL + LXSRBFHL1_XL
            If NUM1_LXSRB > 0 And LXSRBFHL1 >= LXSRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSRBBTHD1_ED As Double = NUM1_LXSRB * BTHD1_ED_LXSRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSRBBTHD1_GL As Double = BTHDXS_GL_water * LXSRBBTHD1_ED * LXSRBFHL1_GL
                Dim LXSRBBTHD1_XL As Double = BTHDXS_XL_water * LXSRBBTHD1_ED * LXSRBFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                LXSRBBTHD1_ALL_a = LXSRBBTHD1_GL + LXSRBBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_LXSRB_List As New List(Of Double)
                Dim FHL1_LXSRB_single As New List(Of Double)
                Dim NUM1_LXSRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_LXSRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_LXSRB / NUM1_LXSRB < LXSRBFHL1 * ZJLGL1_LXSRB Then
                        GoTo iii
                    End If
                    For a1 = FHL1_min_LXSRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_LXSRB / NUM1_LXSRB >= LXSRBFHL1 * ZJLGL1_LXSRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSRB_List.Add(离心式热泵制冷COP曲线(a1))
                            FHL1_LXSRB_single.Add(a1)
                            NUM1_LXSRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo iii
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSRB_List.Add(离心式热泵制冷COP曲线(a1))
                            FHL1_LXSRB_single.Add(a1)
                            NUM1_LXSRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo iii
                        End If
                    Next
iii:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim LXSRBZLCOPXZ1 As Double = COPXZ1_LXSRB_List.Max
                '经过修正后的设备本体耗电功率
                LXSRBBTHD1_ALL_b = LXSRBBTHD1_ALL_a / LXSRBZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_LXSRB_List.IndexOf(LXSRBZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 122).Value = NUM1_LXSRB_QD(COP_MAX_index)
            Else
                LXSRBBTHD1_ALL_a = 0
                LXSRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 122).Value = 0
            End If
            '离心式热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value
            Dim LXSRBFHL2_XL As Double
            If NUM2_LXSRB = 0 Then
                LXSRBFHL2_XL = 0
            Else
                LXSRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value / NUM2_LXSRB
            End If
            Dim LXSRBFHL2 As Double = LXSRBFHL2_GL + LXSRBFHL2_XL
            If NUM2_LXSRB > 0 And LXSRBFHL2 >= LXSRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSRBBTHD2_ED As Double = NUM2_LXSRB * BTHD2_ED_LXSRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSRBBTHD2_GL As Double = BTHDXS_GL_water * LXSRBBTHD2_ED * LXSRBFHL2_GL
                Dim LXSRBBTHD2_XL As Double = BTHDXS_XL_water * LXSRBBTHD2_ED * LXSRBFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                LXSRBBTHD2_ALL_a = LXSRBBTHD2_GL + LXSRBBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_LXSRB_List As New List(Of Double)
                Dim FHL2_LXSRB_single As New List(Of Double)
                Dim NUM2_LXSRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_LXSRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_LXSRB / NUM2_LXSRB < LXSRBFHL2 * ZJLGL2_LXSRB Then
                        GoTo jjj
                    End If
                    For a2 = FHL2_min_LXSRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_LXSRB / NUM2_LXSRB >= LXSRBFHL2 * ZJLGL2_LXSRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSRB_List.Add(离心式热泵制冷COP曲线(a2))
                            FHL2_LXSRB_single.Add(a2)
                            NUM2_LXSRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo jjj
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSRB_List.Add(离心式热泵制冷COP曲线(a2))
                            FHL2_LXSRB_single.Add(a2)
                            NUM2_LXSRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo jjj
                        End If
                    Next
jjj:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim LXSRBZLCOPXZ2 As Double = COPXZ2_LXSRB_List.Max
                '经过修正后的设备本体耗电功率
                LXSRBBTHD2_ALL_b = LXSRBBTHD2_ALL_a / LXSRBZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_LXSRB_List.IndexOf(LXSRBZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 123).Value = NUM2_LXSRB_QD(COP_MAX_index)
            Else
                LXSRBBTHD2_ALL_a = 0
                LXSRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 123).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '空气源热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim KQYRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value
            Dim KQYRBFHL1_XL As Double
            If NUM1_KQYRB = 0 Then
                KQYRBFHL1_XL = 0
            Else
                KQYRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value / NUM1_KQYRB
            End If
            Dim KQYRBFHL1 As Double = KQYRBFHL1_GL + KQYRBFHL1_XL
            If NUM1_KQYRB > 0 And KQYRBFHL1 >= KQYRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim KQYRBBTHD1_ED As Double = NUM1_KQYRB * BTHD1_ED_KQYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim KQYRBBTHD1_GL As Double = BTHDXS_GL_air * KQYRBBTHD1_ED * KQYRBFHL1_GL
                Dim KQYRBBTHD1_XL As Double = BTHDXS_XL_air * KQYRBBTHD1_ED * KQYRBFHL1_XL
                '在没有考虑修正系数前的，总耗电功率
                KQYRBBTHD1_ALL_a = KQYRBBTHD1_GL + KQYRBBTHD1_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_KQYRB_List As New List(Of Double)
                Dim FHL1_KQYRB_single As New List(Of Double)
                Dim NUM1_KQYRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_KQYRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_KQYRB / NUM1_KQYRB < KQYRBFHL1 * ZJLGL1_KQYRB Then
                        GoTo kkk
                    End If
                    For a1 = FHL1_min_KQYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_KQYRB / NUM1_KQYRB >= KQYRBFHL1 * ZJLGL1_KQYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_KQYRB_List.Add(空气源热泵制冷COP曲线(a1))
                            FHL1_KQYRB_single.Add(a1)
                            NUM1_KQYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo kkk
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_KQYRB_List.Add(空气源热泵制冷COP曲线(a1))
                            FHL1_KQYRB_single.Add(a1)
                            NUM1_KQYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo kkk
                        End If
                    Next
kkk:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim KQYRBZLCOPXZ1 As Double = COPXZ1_KQYRB_List.Max
                '经过修正后的设备本体耗电功率
                KQYRBBTHD1_ALL_b = KQYRBBTHD1_ALL_a / KQYRBZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_KQYRB_List.IndexOf(KQYRBZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 116).Value = NUM1_KQYRB_QD(COP_MAX_index)
            Else
                KQYRBBTHD1_ALL_a = 0
                KQYRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 116).Value = 0
            End If
            '空气源热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim KQYRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value
            Dim KQYRBFHL2_XL As Double
            If NUM2_KQYRB = 0 Then
                KQYRBFHL2_XL = 0
            Else
                KQYRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value / NUM2_KQYRB
            End If
            Dim KQYRBFHL2 As Double = KQYRBFHL2_GL + KQYRBFHL2_XL
            If NUM2_KQYRB > 0 And KQYRBFHL2 >= KQYRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim KQYRBBTHD2_ED As Double = NUM2_KQYRB * BTHD2_ED_KQYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim KQYRBBTHD2_GL As Double = BTHDXS_GL_air * KQYRBBTHD2_ED * KQYRBFHL2_GL
                Dim KQYRBBTHD2_XL As Double = BTHDXS_XL_air * KQYRBBTHD2_ED * KQYRBFHL2_XL
                '在没有考虑修正系数前的，总耗电功率
                KQYRBBTHD2_ALL_a = KQYRBBTHD2_GL + KQYRBBTHD2_XL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_KQYRB_List As New List(Of Double)
                Dim FHL2_KQYRB_single As New List(Of Double)
                Dim NUM2_KQYRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_KQYRB Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_KQYRB / NUM2_KQYRB < KQYRBFHL2 * ZJLGL2_KQYRB Then
                        GoTo lll
                    End If
                    For a2 = FHL2_min_KQYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_KQYRB / NUM2_KQYRB >= KQYRBFHL2 * ZJLGL2_KQYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_KQYRB_List.Add(空气源热泵制冷COP曲线(a2))
                            FHL2_KQYRB_single.Add(a2)
                            NUM2_KQYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo lll
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_KQYRB_List.Add(空气源热泵制冷COP曲线(a2))
                            FHL2_KQYRB_single.Add(a2)
                            NUM2_KQYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo lll
                        End If
                    Next
lll:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim KQYRBZLCOPXZ2 As Double = COPXZ2_KQYRB_List.Max
                '经过修正后的设备本体耗电功率
                KQYRBBTHD2_ALL_b = KQYRBBTHD2_ALL_a / KQYRBZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_KQYRB_List.IndexOf(KQYRBZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 117).Value = NUM2_KQYRB_QD(COP_MAX_index)
            Else
                KQYRBBTHD2_ALL_a = 0
                KQYRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 117).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '直燃型溴化锂（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value
            Dim ZRXXHLFHL1 As Double = ZRXXHLFHL1_GL
            If NUM1_ZRXXHL > 0 And ZRXXHLFHL1 >= ZRXXHLFHL1_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim ZRXXHLBTTRQ1_ED As Double = NUM1_ZRXXHL * BTHQ1_ED_ZRXXHL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim ZRXXHLBTTRQ1_GL As Double = TRQHLXZXS_QT * ZRXXHLBTTRQ1_ED * ZRXXHLFHL1 * ZRXXHLZLHQXZ
                '在没有考虑修正系数前的，总天然气耗量
                ZRXXHLBTTRQ1_ALL_a = ZRXXHLBTTRQ1_GL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_ZRXXHL_List As New List(Of Double)
                Dim FHL1_ZRXXHL_single As New List(Of Double)
                Dim NUM1_ZRXXHL_QD As New List(Of Double)
                For n1 = 0 To NUM1_ZRXXHL Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJLGL1_ZRXXHL / NUM1_ZRXXHL < ZRXXHLFHL1 * ZJLGL1_ZRXXHL Then
                        GoTo mmm
                    End If
                    '直燃型溴化锂可以超发到1.2
                    For a1 = FHL1_min_ZRXXHL To (1.2 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * ZJLGL1_ZRXXHL / NUM1_ZRXXHL >= ZRXXHLFHL1 * ZJLGL1_ZRXXHL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_ZRXXHL_List.Add(直燃型溴化锂制冷COP曲线(a1))
                            FHL1_ZRXXHL_single.Add(a1)
                            NUM1_ZRXXHL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo mmm
                        ElseIf a1 >= 1.2 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_ZRXXHL_List.Add(直燃型溴化锂制冷COP曲线(a1))
                            FHL1_ZRXXHL_single.Add(a1)
                            NUM1_ZRXXHL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo mmm
                        End If
                    Next
mmm:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim ZRXXHLZLCOPXZ1 As Double = COPXZ1_ZRXXHL_List.Max
                '经过修正后的设备本体天然气耗量
                ZRXXHLBTTRQ1_ALL_b = ZRXXHLBTTRQ1_ALL_a / ZRXXHLZLCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_ZRXXHL_List.IndexOf(ZRXXHLZLCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 124).Value = NUM1_ZRXXHL_QD(COP_MAX_index)
            Else
                ZRXXHLBTTRQ1_ALL_a = 0
                ZRXXHLBTTRQ1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 124).Value = 0
            End If
            '直燃型溴化锂（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value
            Dim ZRXXHLFHL2 As Double = ZRXXHLFHL2_GL
            If NUM2_ZRXXHL > 0 And ZRXXHLFHL2 >= ZRXXHLFHL2_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim ZRXXHLBTTRQ2_ED As Double = NUM2_ZRXXHL * BTHQ2_ED_ZRXXHL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim ZRXXHLBTTRQ2_GL As Double = TRQHLXZXS_QT * ZRXXHLBTTRQ2_ED * ZRXXHLFHL2 * ZRXXHLZLHQXZ
                '在没有考虑修正系数前的，总天然气耗量
                ZRXXHLBTTRQ2_ALL_a = ZRXXHLBTTRQ2_GL
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_ZRXXHL_List As New List(Of Double)
                Dim FHL2_ZRXXHL_single As New List(Of Double)
                Dim NUM2_ZRXXHL_QD As New List(Of Double)
                For n2 = 0 To NUM2_ZRXXHL Step 1
                    '如果设备启动100%，制冷量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJLGL2_ZRXXHL / NUM2_ZRXXHL < ZRXXHLFHL2 * ZJLGL2_ZRXXHL Then
                        GoTo nnn
                    End If
                    '直燃型溴化锂可以超发到1.2
                    For a2 = FHL2_min_ZRXXHL To (1.2 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * ZJLGL2_ZRXXHL / NUM2_ZRXXHL >= ZRXXHLFHL2 * ZJLGL2_ZRXXHL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_ZRXXHL_List.Add(直燃型溴化锂制冷COP曲线(a2))
                            FHL2_ZRXXHL_single.Add(a2)
                            NUM2_ZRXXHL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo nnn
                        ElseIf a2 >= 1.2 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_ZRXXHL_List.Add(直燃型溴化锂制冷COP曲线(a2))
                            FHL2_ZRXXHL_single.Add(a2)
                            NUM2_ZRXXHL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo nnn
                        End If
                    Next
nnn:
                Next
                '计算此时的制冷COP修正系数所有结果中得到最大值
                Dim ZRXXHLZLCOPXZ2 As Double = COPXZ2_ZRXXHL_List.Max
                '经过修正后的设备本体天然气耗量
                ZRXXHLBTTRQ2_ALL_b = ZRXXHLBTTRQ2_ALL_a / ZRXXHLZLCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_ZRXXHL_List.IndexOf(ZRXXHLZLCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 125).Value = NUM2_ZRXXHL_QD(COP_MAX_index)
            Else
                ZRXXHLBTTRQ2_ALL_a = 0
                ZRXXHLBTTRQ2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 125).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '根据内燃机负荷率，计算内燃机的发电效率
            '内燃机（1）
            '内燃发电机（1）
            '此时的设备负荷率
            Dim NRJFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2).Value
            If NUM1_NRJ > 0 And NRJFHL1 >= NRJFHL1_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim NRJTRQ1_ED As Double = NUM1_NRJ * TRQ1_ED_NRJ
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim NRJTRQ1 As Double = TRQHLXZXS_NRJ * NRJTRQ1_ED * NRJFHL1
                '在没有考虑修正系数前的，总天然气耗量
                NRJTRQ1_ALL_a = NRJTRQ1
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的效率修正系数、单台设备负荷率、设备启动数量
                Dim NRJZLXLXZ1_List As New List(Of Double)
                Dim FHL1_NRJ_single As New List(Of Double)
                Dim NUM1_NRJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_NRJ Step 1
                    '如果设备启动100%，发电量都不够，则跳入下一个循环，提高速度
                    If n1 * FDGL1_ED_NRJ < NRJFHL1 * FDGL1_ED_NRJ * NUM1_NRJ Then
                        GoTo ooo
                    End If
                    For a1 = FHL1_min_NRJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * FDGL1_ED_NRJ >= NRJFHL1 * FDGL1_ED_NRJ * NUM1_NRJ Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ1_List.Add(内燃机发电效率曲线(a1))
                            FHL1_NRJ_single.Add(a1)
                            NUM1_NRJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ooo
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ1_List.Add(内燃机发电效率曲线(a1))
                            FHL1_NRJ_single.Add(a1)
                            NUM1_NRJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ooo
                        End If
                    Next
ooo:
                Next
                '计算此时的效率修正系数所有结果中得到最大值
                Dim NRJZLXLXZ1 As Double = NRJZLXLXZ1_List.Max
                '经过修正后的设备本体天然气耗量
                NRJTRQ1_ALL_b = NRJTRQ1_ALL_a / NRJZLXLXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = NRJZLXLXZ1_List.IndexOf(NRJZLXLXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 110).Value = NUM1_NRJ_QD(COP_MAX_index)
            Else
                NRJTRQ1_ALL_a = 0
                NRJTRQ1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 110).Value = 0
            End If
            '内燃发电机（2）
            '此时的设备负荷率
            Dim NRJFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3).Value
            If NUM2_NRJ > 0 And NRJFHL2 >= NRJFHL2_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim NRJTRQ2_ED As Double = NUM2_NRJ * TRQ2_ED_NRJ
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim NRJTRQ2 As Double = TRQHLXZXS_NRJ * NRJTRQ2_ED * NRJFHL2
                '在没有考虑修正系数前的，总天然气耗量
                NRJTRQ2_ALL_a = NRJTRQ2
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的效率修正系数、单台设备负荷率、设备启动数量
                Dim NRJZLXLXZ2_List As New List(Of Double)
                Dim FHL2_NRJ_single As New List(Of Double)
                Dim NUM2_NRJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_NRJ Step 1
                    '如果设备启动100%，发电量都不够，则跳入下一个循环，提高速度
                    If n2 * FDGL2_ED_NRJ < NRJFHL2 * FDGL2_ED_NRJ * NUM2_NRJ Then
                        GoTo ppp
                    End If
                    For a2 = FHL2_min_NRJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * FDGL2_ED_NRJ >= NRJFHL2 * FDGL2_ED_NRJ * NUM2_NRJ Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ2_List.Add(内燃机发电效率曲线(a2))
                            FHL2_NRJ_single.Add(a2)
                            NUM2_NRJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ppp
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ2_List.Add(内燃机发电效率曲线(a2))
                            FHL2_NRJ_single.Add(a2)
                            NUM2_NRJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo ppp
                        End If
                    Next
ppp:
                Next
                '计算此时的效率修正系数所有结果中得到最大值
                Dim NRJZLXLXZ2 As Double = NRJZLXLXZ2_List.Max
                '经过修正后的设备本体天然气耗量
                NRJTRQ2_ALL_b = NRJTRQ2_ALL_a / NRJZLXLXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = NRJZLXLXZ2_List.IndexOf(NRJZLXLXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 111).Value = NUM2_NRJ_QD(COP_MAX_index)
            Else
                NRJTRQ2_ALL_a = 0
                NRJTRQ2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 111).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算空调设备本体耗电综合修正系数
            '空调设备在考虑修正系数前的本体总耗电功率
            Dim ZHD_ALL_a As Double = LXSLSJBTHD1_ALL_a + LXSLSJBTHD2_ALL_a + SLLGJBTHD1_ALL_a + SLLGJBTHD2_ALL_a + FLLGJBTHD1_ALL_a + FLLGJBTHD2_ALL_a + SDYRBBTHD1_ALL_a + SDYRBBTHD2_ALL_a + KQYRBBTHD1_ALL_a + KQYRBBTHD2_ALL_a + LXSRBBTHD1_ALL_a + LXSRBBTHD2_ALL_a
            '空调设备在考虑修正系数之后的本体总耗电功率
            Dim ZHD_ALL_b As Double = LXSLSJBTHD1_ALL_b + LXSLSJBTHD2_ALL_b + SLLGJBTHD1_ALL_b + SLLGJBTHD2_ALL_b + FLLGJBTHD1_ALL_b + FLLGJBTHD2_ALL_b + SDYRBBTHD1_ALL_b + SDYRBBTHD2_ALL_b + KQYRBBTHD1_ALL_b + KQYRBBTHD2_ALL_b + LXSRBBTHD1_ALL_b + LXSRBBTHD2_ALL_b
            '设备本体耗电功率综合修正系数
            Dim BTHD_XZ As Double = 1
            If ZHD_ALL_a > 0 And ZHD_ALL_b > 0 Then
                BTHD_XZ = ZHD_ALL_b / ZHD_ALL_a
            Else
                '默认值等于1
                BTHD_XZ = 1
            End If
            '天然气耗电综合修正系数
            Dim ZHTRQ_ALL_a As Double = ZRXXHLBTTRQ1_ALL_a + ZRXXHLBTTRQ2_ALL_a + NRJTRQ1_ALL_a + NRJTRQ2_ALL_a
            Dim ZHTRQ_ALL_b As Double = ZRXXHLBTTRQ1_ALL_b + ZRXXHLBTTRQ2_ALL_b + NRJTRQ1_ALL_b + NRJTRQ2_ALL_b
            Dim TRQ_XZ As Double = 1
            If ZHTRQ_ALL_a > 0 And ZHTRQ_ALL_b > 0 Then
                TRQ_XZ = ZHTRQ_ALL_b / ZHTRQ_ALL_a
            Else
                '默认值等于1
                TRQ_XZ = 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '将综合修正系数写入Excel
            '制冷季系数
            '制冷设备本体耗电修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 106).Value = BTHD_XZ
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 106).Value = BTHD_XZ
            '制冷天然气消耗修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 107).Value = TRQ_XZ
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 107).Value = TRQ_XZ
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算出的结果保留3位小数，增加美观度
            For i = 106 To 107
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '如果本工况没有选择内燃机，则将内燃机及其余热利用的各种系数设置为1
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 98).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 98).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 99).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 99).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 100).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 100).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 101).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 101).Value = 1
            End If
        Else
            '常规计算模式，修正系数全部设置为1
            '制冷季系数
            '制冷设备本体耗电修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 106).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 106).Value = 1
            '制冷天然气消耗修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 107).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 107).Value = 1
            '内燃机及其余热利用系数
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 101)).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 101)).Value = 1
        End If
    End Sub
    Sub 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp As Object, b As Integer, FHTJJD As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '只有全局寻优计算模式才进行修正
        If calculation_mode = 2 And RFHZXQL(b) + XNXRGL(b) > 0 Then
            '读取采暖季装机方案及参数
            Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
            '内燃机
            Dim NUM1_NRJ As Double = ans_ZJFA_R(0)
            Dim NUM2_NRJ As Double = ans_ZJFA_R(1)
            Dim FDGL1_ED_NRJ As Double = ans_ZJFA_R(2)
            Dim FDGL2_ED_NRJ As Double = ans_ZJFA_R(3)
            Dim YRGL1_ED_NRJ As Double = ans_ZJFA_R(4)
            Dim YRGL2_ED_NRJ As Double = ans_ZJFA_R(5)
            Dim TRQ1_ED_NRJ As Double = ans_ZJFA_R(6)
            Dim TRQ2_ED_NRJ As Double = ans_ZJFA_R(7)
            Dim FJHD1_ED_NRJ As Double = ans_ZJFA_R(8)
            Dim FJHD2_ED_NRJ As Double = ans_ZJFA_R(9)
            '天然气锅炉
            Dim NUM1_TRQGL As Double = ans_ZJFA_R(10)
            Dim NUM2_TRQGL As Double = ans_ZJFA_R(11)
            Dim ZJRGL1_TRQGL As Double = ans_ZJFA_R(12)
            Dim ZJRGL2_TRQGL As Double = ans_ZJFA_R(13)
            Dim BTHQ1_ED_TRQGL As Double = ans_ZJFA_R(14)
            Dim BTHQ2_ED_TRQGL As Double = ans_ZJFA_R(15)
            Dim FJHD1_ED_TRQGL As Double = ans_ZJFA_R(16)
            Dim FJHD2_ED_TRQGL As Double = ans_ZJFA_R(17)
            '电锅炉
            Dim NUM1_DGL As Double = ans_ZJFA_R(18)
            Dim NUM2_DGL As Double = ans_ZJFA_R(19)
            Dim ZJRGL1_DGL As Double = ans_ZJFA_R(20)
            Dim ZJRGL2_DGL As Double = ans_ZJFA_R(21)
            Dim BTHD1_ED_DGL As Double = ans_ZJFA_R(22)
            Dim BTHD2_ED_DGL As Double = ans_ZJFA_R(23)
            Dim FJHD1_ED_DGL As Double = ans_ZJFA_R(24)
            Dim FJHD2_ED_DGL As Double = ans_ZJFA_R(25)
            '风冷螺杆机
            Dim NUM1_FLLGJ As Double = ans_ZJFA_R(26)
            Dim NUM2_FLLGJ As Double = ans_ZJFA_R(27)
            Dim ZJRGL1_FLLGJ As Double = ans_ZJFA_R(28)
            Dim ZJRGL2_FLLGJ As Double = ans_ZJFA_R(29)
            Dim BTHD1_ED_FLLGJ As Double = ans_ZJFA_R(30)
            Dim BTHD2_ED_FLLGJ As Double = ans_ZJFA_R(31)
            Dim FJHD1_ED_FLLGJ As Double = ans_ZJFA_R(32)
            Dim FJHD2_ED_FLLGJ As Double = ans_ZJFA_R(33)
            '水地源热泵
            Dim NUM1_SDYRB As Double = ans_ZJFA_R(34)
            Dim NUM2_SDYRB As Double = ans_ZJFA_R(35)
            Dim ZJRGL1_SDYRB As Double = ans_ZJFA_R(36)
            Dim ZJRGL2_SDYRB As Double = ans_ZJFA_R(37)
            Dim BTHD1_ED_SDYRB As Double = ans_ZJFA_R(38)
            Dim BTHD2_ED_SDYRB As Double = ans_ZJFA_R(39)
            Dim FJHD1_ED_SDYRB As Double = ans_ZJFA_R(40)
            Dim FJHD2_ED_SDYRB As Double = ans_ZJFA_R(41)
            '离心式热泵
            Dim NUM1_LXSRB As Double = ans_ZJFA_R(42)
            Dim NUM2_LXSRB As Double = ans_ZJFA_R(43)
            Dim ZJRGL1_LXSRB As Double = ans_ZJFA_R(44)
            Dim ZJRGL2_LXSRB As Double = ans_ZJFA_R(45)
            Dim BTHD1_ED_LXSRB As Double = ans_ZJFA_R(46)
            Dim BTHD2_ED_LXSRB As Double = ans_ZJFA_R(47)
            Dim FJHD1_ED_LXSRB As Double = ans_ZJFA_R(48)
            Dim FJHD2_ED_LXSRB As Double = ans_ZJFA_R(49)
            '空气源热泵
            Dim NUM1_KQYRB As Double = ans_ZJFA_R(50)
            Dim NUM2_KQYRB As Double = ans_ZJFA_R(51)
            Dim ZJRGL1_KQYRB As Double = ans_ZJFA_R(52)
            Dim ZJRGL2_KQYRB As Double = ans_ZJFA_R(53)
            Dim BTHD1_ED_KQYRB As Double = ans_ZJFA_R(54)
            Dim BTHD2_ED_KQYRB As Double = ans_ZJFA_R(55)
            Dim FJHD1_ED_KQYRB As Double = ans_ZJFA_R(56)
            Dim FJHD2_ED_KQYRB As Double = ans_ZJFA_R(57)
            '直燃型溴化锂
            Dim NUM1_ZRXXHL As Double = ans_ZJFA_R(58)
            Dim NUM2_ZRXXHL As Double = ans_ZJFA_R(59)
            Dim ZJRGL1_ZRXXHL As Double = ans_ZJFA_R(60)
            Dim ZJRGL2_ZRXXHL As Double = ans_ZJFA_R(61)
            Dim BTHQ1_ED_ZRXXHL As Double = ans_ZJFA_R(62)
            Dim BTHQ2_ED_ZRXXHL As Double = ans_ZJFA_R(63)
            Dim FJHD1_ED_ZRXXHL As Double = ans_ZJFA_R(64)
            Dim FJHD2_ED_ZRXXHL As Double = ans_ZJFA_R(65)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '读取各种修正系数
            Dim ans_XZXS_R = 读取采暖季输入的计算系数(ExcelApp)
            Dim BTHDXS_GR_water As Double = ans_XZXS_R(0)
            Dim BTHDXS_XR_water As Double = ans_XZXS_R(1)
            Dim BTHDXS_GR_air As Double = ans_XZXS_R(2)
            Dim BTHDXS_XR_air As Double = ans_XZXS_R(3)
            Dim FJHDXS As Double = ans_XZXS_R(4)
            Dim TRQHLXZXS_QT As Double = ans_XZXS_R(5)
            Dim TRQHLXZXS_NRJ As Double = ans_XZXS_R(6)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '各个设备可以允许运行的最低负荷率
            '内燃机
            Dim NRJFHL1_min As Double
            Dim NRJFHL2_min As Double
            If NUM1_NRJ > 0 Then
                NRJFHL1_min = FHL1_min_NRJ / NUM1_NRJ
            Else
                NRJFHL1_min = 0
            End If
            If NUM2_NRJ > 0 Then
                NRJFHL2_min = FHL2_min_NRJ / NUM2_NRJ
            Else
                NRJFHL2_min = 0
            End If
            '天然气采暖锅炉
            '设备最低运行负荷率
            Dim TRQGLFHL1_min As Double
            Dim TRQGLFHL2_min As Double
            If NUM1_TRQGL > 0 Then
                TRQGLFHL1_min = FHL1_min_TRQGL / NUM1_TRQGL
            Else
                TRQGLFHL1_min = 0
            End If
            If NUM2_TRQGL > 0 Then
                TRQGLFHL2_min = FHL2_min_TRQGL / NUM2_TRQGL
            Else
                TRQGLFHL2_min = 0
            End If
            '风冷螺杆机
            '设备最低运行负荷率
            Dim FLLGJFHL1_min As Double
            Dim FLLGJFHL2_min As Double
            If NUM1_FLLGJ > 0 Then
                FLLGJFHL1_min = FHL1_min_FLLGJ / NUM1_FLLGJ
            Else
                FLLGJFHL1_min = 0
            End If
            If NUM2_FLLGJ > 0 Then
                FLLGJFHL2_min = FHL2_min_FLLGJ / NUM2_FLLGJ
            Else
                FLLGJFHL2_min = 0
            End If
            '水（地）源热泵
            '设备最低运行负荷率
            Dim SDYRBFHL1_min As Double
            Dim SDYRBFHL2_min As Double
            If NUM1_SDYRB > 0 Then
                SDYRBFHL1_min = FHL1_min_SDYRB / NUM1_SDYRB
            Else
                SDYRBFHL1_min = 0
            End If
            If NUM2_SDYRB > 0 Then
                SDYRBFHL2_min = FHL2_min_SDYRB / NUM2_SDYRB
            Else
                SDYRBFHL2_min = 0
            End If
            '空气源热泵
            '设备最低运行负荷率
            Dim KQYRBFHL1_min As Double
            Dim KQYRBFHL2_min As Double
            If NUM1_KQYRB > 0 Then
                KQYRBFHL1_min = FHL1_min_KQYRB / NUM1_KQYRB
            Else
                KQYRBFHL1_min = 0
            End If
            If NUM2_KQYRB > 0 Then
                KQYRBFHL2_min = FHL2_min_KQYRB / NUM2_KQYRB
            Else
                KQYRBFHL2_min = 0
            End If
            '离心式热泵
            '设备最低运行负荷率
            Dim LXSRBFHL1_min As Double
            Dim LXSRBFHL2_min As Double
            If NUM1_LXSRB > 0 Then
                LXSRBFHL1_min = FHL1_min_LXSRB / NUM1_LXSRB
            Else
                LXSRBFHL1_min = 0
            End If
            If NUM2_LXSRB > 0 Then
                LXSRBFHL2_min = FHL2_min_LXSRB / NUM2_LXSRB
            Else
                LXSRBFHL2_min = 0
            End If
            '直燃型溴化锂
            '设备最低运行负荷率
            Dim ZRXXHLFHL1_min As Double
            Dim ZRXXHLFHL2_min As Double
            If NUM1_ZRXXHL > 0 Then
                ZRXXHLFHL1_min = FHL1_min_ZRXXHL / NUM1_ZRXXHL
            Else
                ZRXXHLFHL1_min = 0
            End If
            If NUM2_ZRXXHL > 0 Then
                ZRXXHLFHL2_min = FHL2_min_ZRXXHL / NUM2_ZRXXHL
            Else
                ZRXXHLFHL2_min = 0
            End If
            '电采暖锅炉
            Dim DGLFHL1_min As Double
            Dim DGLFHL2_min As Double
            If NUM1_DGL > 0 Then
                DGLFHL1_min = FHL1_min_DGL / NUM1_DGL
            Else
                DGLFHL1_min = 0
            End If
            If NUM2_DGL > 0 Then
                DGLFHL2_min = FHL2_min_DGL / NUM2_DGL
            Else
                DGLFHL2_min = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '各空调设备本体耗电功率在考虑修正系数前的耗电功率和考虑修正系数后的耗电系数
            '风冷螺杆机
            Dim FLLGJBTHD1_ALL_a As Double
            Dim FLLGJBTHD1_ALL_b As Double
            Dim FLLGJBTHD2_ALL_a As Double
            Dim FLLGJBTHD2_ALL_b As Double
            '水（地）源热泵
            Dim SDYRBBTHD1_ALL_a As Double
            Dim SDYRBBTHD1_ALL_b As Double
            Dim SDYRBBTHD2_ALL_a As Double
            Dim SDYRBBTHD2_ALL_b As Double
            '离心式热泵
            Dim LXSRBBTHD1_ALL_a As Double
            Dim LXSRBBTHD1_ALL_b As Double
            Dim LXSRBBTHD2_ALL_a As Double
            Dim LXSRBBTHD2_ALL_b As Double
            '空气源热泵
            Dim KQYRBBTHD1_ALL_a As Double
            Dim KQYRBBTHD1_ALL_b As Double
            Dim KQYRBBTHD2_ALL_a As Double
            Dim KQYRBBTHD2_ALL_b As Double
            '电采暖锅炉
            Dim DGLBTHD1_ALL_a As Double
            Dim DGLBTHD1_ALL_b As Double
            Dim DGLBTHD2_ALL_a As Double
            Dim DGLBTHD2_ALL_b As Double
            '各空调设备本体天然气耗量在考虑修正系数前的天然气耗量和考虑修正系数后的天然气耗量
            '天然气锅炉
            Dim TRQGLBTTRQ1_ALL_a As Double
            Dim TRQGLBTTRQ1_ALL_b As Double
            Dim TRQGLBTTRQ2_ALL_a As Double
            Dim TRQGLBTTRQ2_ALL_b As Double
            '直燃型溴化锂
            Dim ZRXXHLBTTRQ1_ALL_a As Double
            Dim ZRXXHLBTTRQ1_ALL_b As Double
            Dim ZRXXHLBTTRQ2_ALL_a As Double
            Dim ZRXXHLBTTRQ2_ALL_b As Double
            '内燃发电机
            Dim NRJTRQ1_ALL_a As Double
            Dim NRJTRQ1_ALL_b As Double
            Dim NRJTRQ2_ALL_a As Double
            Dim NRJTRQ2_ALL_b As Double
            '混水设备
            Dim BTHD_ALL_HS_a As Double
            Dim BTHD_ALL_HS_b As Double
            '梯级供热设备
            Dim BTHD_ALL_TJ_a As Double
            Dim BTHD_ALL_TJ_b As Double
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————                      
            '天然气锅炉（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim TRQGLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value
            Dim TRQGLFHL1 As Double = TRQGLFHL1_GR
            If NUM1_TRQGL > 0 And TRQGLFHL1 >= TRQGLFHL1_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim TRQGLBTTRQ1_ED As Double = NUM1_TRQGL * BTHQ1_ED_TRQGL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim TRQGLBTTRQ1_GR As Double = TRQHLXZXS_QT * TRQGLBTTRQ1_ED * TRQGLFHL1_GR
                '在没有考虑修正系数前的，总天然气耗量
                TRQGLBTTRQ1_ALL_a = TRQGLBTTRQ1_GR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_TRQGL_List As New List(Of Double)
                Dim FHL1_TRQGL_single As New List(Of Double)
                Dim NUM1_TRQGL_QD As New List(Of Double)
                For n1 = 0 To NUM1_TRQGL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_TRQGL / NUM1_TRQGL < TRQGLFHL1 * ZJRGL1_TRQGL Then
                        GoTo aaa
                    End If
                    For a1 = FHL1_min_TRQGL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_TRQGL / NUM1_TRQGL >= TRQGLFHL1 * ZJRGL1_TRQGL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_TRQGL_List.Add(天然气采暖锅炉制热效率曲线(a1))
                            FHL1_TRQGL_single.Add(a1)
                            NUM1_TRQGL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo aaa
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_TRQGL_List.Add(天然气采暖锅炉制热效率曲线(a1))
                            FHL1_TRQGL_single.Add(a1)
                            NUM1_TRQGL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo aaa
                        End If
                    Next
aaa:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim TRQGLZRCOPXZ1 As Double = COPXZ1_TRQGL_List.Max
                '经过修正后的设备本体天然气耗量
                TRQGLBTTRQ1_ALL_b = TRQGLBTTRQ1_ALL_a / TRQGLZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_TRQGL_List.IndexOf(TRQGLZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 128).Value = NUM1_TRQGL_QD(COP_MAX_index)
            Else
                TRQGLBTTRQ1_ALL_a = 0
                TRQGLBTTRQ1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 128).Value = 0
            End If
            '天然气锅炉（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim TRQGLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value
            Dim TRQGLFHL2 As Double = TRQGLFHL2_GR
            If NUM2_TRQGL > 0 And TRQGLFHL2 >= TRQGLFHL2_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim TRQGLBTTRQ2_ED As Double = NUM2_TRQGL * BTHQ2_ED_TRQGL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim TRQGLBTTRQ2_GR As Double = TRQHLXZXS_QT * TRQGLBTTRQ2_ED * TRQGLFHL2_GR
                '在没有考虑修正系数前的，总天然气耗量
                TRQGLBTTRQ2_ALL_a = TRQGLBTTRQ2_GR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_TRQGL_List As New List(Of Double)
                Dim FHL2_TRQGL_single As New List(Of Double)
                Dim NUM2_TRQGL_QD As New List(Of Double)
                For n2 = 0 To NUM2_TRQGL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_TRQGL / NUM2_TRQGL < TRQGLFHL2 * ZJRGL2_TRQGL Then
                        GoTo bbb
                    End If
                    For a2 = FHL2_min_TRQGL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_TRQGL / NUM2_TRQGL >= TRQGLFHL2 * ZJRGL2_TRQGL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_TRQGL_List.Add(天然气采暖锅炉制热效率曲线(a2))
                            FHL2_TRQGL_single.Add(a2)
                            NUM2_TRQGL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo bbb
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_TRQGL_List.Add(天然气采暖锅炉制热效率曲线(a2))
                            FHL2_TRQGL_single.Add(a2)
                            NUM2_TRQGL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo bbb
                        End If
                    Next
bbb:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim TRQGLZRCOPXZ2 As Double = COPXZ2_TRQGL_List.Max
                '经过修正后的设备本体天然气耗量
                TRQGLBTTRQ2_ALL_b = TRQGLBTTRQ2_ALL_a / TRQGLZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_TRQGL_List.IndexOf(TRQGLZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 129).Value = NUM2_TRQGL_QD(COP_MAX_index)
            Else
                TRQGLBTTRQ2_ALL_a = 0
                TRQGLBTTRQ2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 129).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '直燃型溴化锂（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value
            Dim ZRXXHLFHL1 As Double = ZRXXHLFHL1_GR
            If NUM1_ZRXXHL > 0 And ZRXXHLFHL1 >= ZRXXHLFHL1_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim ZRXXHLBTTRQ1_ED As Double = NUM1_ZRXXHL * BTHQ1_ED_ZRXXHL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim ZRXXHLBTTRQ1_GR As Double = TRQHLXZXS_QT * ZRXXHLBTTRQ1_ED * ZRXXHLFHL1_GR
                '在没有考虑修正系数前的，总天然气耗量
                ZRXXHLBTTRQ1_ALL_a = ZRXXHLBTTRQ1_GR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_ZRXXHL_List As New List(Of Double)
                Dim FHL1_ZRXXHL_single As New List(Of Double)
                Dim NUM1_ZRXXHL_QD As New List(Of Double)
                For n1 = 0 To NUM1_ZRXXHL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_ZRXXHL / NUM1_ZRXXHL < ZRXXHLFHL1 * ZJRGL1_ZRXXHL Then
                        GoTo ccc
                    End If
                    For a1 = FHL1_min_ZRXXHL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_ZRXXHL / NUM1_ZRXXHL >= ZRXXHLFHL1 * ZJRGL1_ZRXXHL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_ZRXXHL_List.Add(直燃型溴化锂制热COP曲线(a1))
                            FHL1_ZRXXHL_single.Add(a1)
                            NUM1_ZRXXHL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ccc
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_ZRXXHL_List.Add(直燃型溴化锂制热COP曲线(a1))
                            FHL1_ZRXXHL_single.Add(a1)
                            NUM1_ZRXXHL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ccc
                        End If
                    Next
ccc:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim ZRXXHLZRCOPXZ1 As Double = COPXZ1_ZRXXHL_List.Max
                '经过修正后的设备本体天然气耗量
                ZRXXHLBTTRQ1_ALL_b = ZRXXHLBTTRQ1_ALL_a / ZRXXHLZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_ZRXXHL_List.IndexOf(ZRXXHLZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 138).Value = NUM1_ZRXXHL_QD(COP_MAX_index)
            Else
                ZRXXHLBTTRQ1_ALL_a = 0
                ZRXXHLBTTRQ1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 138).Value = 0
            End If
            '直燃型溴化锂（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value
            Dim ZRXXHLFHL2 As Double = ZRXXHLFHL2_GR
            If NUM2_ZRXXHL > 0 And ZRXXHLFHL2 >= ZRXXHLFHL2_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim ZRXXHLBTTRQ2_ED As Double = NUM2_ZRXXHL * BTHQ2_ED_ZRXXHL
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim ZRXXHLBTTRQ2_GR As Double = TRQHLXZXS_QT * ZRXXHLBTTRQ2_ED * ZRXXHLFHL2_GR
                '在没有考虑修正系数前的，总天然气耗量
                ZRXXHLBTTRQ2_ALL_a = ZRXXHLBTTRQ2_GR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_ZRXXHL_List As New List(Of Double)
                Dim FHL2_ZRXXHL_single As New List(Of Double)
                Dim NUM2_ZRXXHL_QD As New List(Of Double)
                For n2 = 0 To NUM2_ZRXXHL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_ZRXXHL / NUM2_ZRXXHL < ZRXXHLFHL2 * ZJRGL2_ZRXXHL Then
                        GoTo ddd
                    End If
                    For a2 = FHL2_min_ZRXXHL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_ZRXXHL / NUM2_ZRXXHL >= ZRXXHLFHL2 * ZJRGL2_ZRXXHL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_ZRXXHL_List.Add(直燃型溴化锂制热COP曲线(a2))
                            FHL2_ZRXXHL_single.Add(a2)
                            NUM2_ZRXXHL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ddd
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_ZRXXHL_List.Add(直燃型溴化锂制热COP曲线(a2))
                            FHL2_ZRXXHL_single.Add(a2)
                            NUM2_ZRXXHL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ddd
                        End If
                    Next
ddd:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim ZRXXHLZRCOPXZ2 As Double = COPXZ2_ZRXXHL_List.Max
                '经过修正后的设备本体天然气耗量
                ZRXXHLBTTRQ2_ALL_b = ZRXXHLBTTRQ2_ALL_a / ZRXXHLZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_ZRXXHL_List.IndexOf(ZRXXHLZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 139).Value = NUM2_ZRXXHL_QD(COP_MAX_index)
            Else
                ZRXXHLBTTRQ2_ALL_a = 0
                ZRXXHLBTTRQ2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 139).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '根据内燃机负荷率，计算内燃机的发电效率
            '内燃机（1）
            '内燃发电机（1）
            '此时的设备负荷率
            Dim NRJFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            If NUM1_NRJ > 0 And NRJFHL1 >= NRJFHL1_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim NRJTRQ1_ED As Double = NUM1_NRJ * TRQ1_ED_NRJ
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim NRJTRQ1 As Double = TRQHLXZXS_NRJ * NRJTRQ1_ED * NRJFHL1
                '在没有考虑修正系数前的，总天然气耗量
                NRJTRQ1_ALL_a = NRJTRQ1
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的效率修正系数、单台设备负荷率、设备启动数量
                Dim NRJZLXLXZ1_List As New List(Of Double)
                Dim FHL1_NRJ_single As New List(Of Double)
                Dim NUM1_NRJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_NRJ Step 1
                    '如果设备启动100%，发电量都不够，则跳入下一个循环，提高速度
                    If n1 * FDGL1_ED_NRJ < NRJFHL1 * FDGL1_ED_NRJ * NUM1_NRJ Then
                        GoTo eee
                    End If
                    For a1 = FHL1_min_NRJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a1 * n1 * FDGL1_ED_NRJ >= NRJFHL1 * FDGL1_ED_NRJ * NUM1_NRJ Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ1_List.Add(内燃机发电效率曲线(a1))
                            FHL1_NRJ_single.Add(a1)
                            NUM1_NRJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo eee
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ1_List.Add(内燃机发电效率曲线(a1))
                            FHL1_NRJ_single.Add(a1)
                            NUM1_NRJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo eee
                        End If
                    Next
eee:
                Next
                '计算此时的效率修正系数所有结果中得到最大值
                Dim NRJZLXLXZ1 As Double = NRJZLXLXZ1_List.Max
                '经过修正后的设备本体天然气耗量
                NRJTRQ1_ALL_b = NRJTRQ1_ALL_a / NRJZLXLXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = NRJZLXLXZ1_List.IndexOf(NRJZLXLXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 126).Value = NUM1_NRJ_QD(COP_MAX_index)
            Else
                NRJTRQ1_ALL_a = 0
                NRJTRQ1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 126).Value = 0
            End If
            '内燃发电机（2）
            '此时的设备负荷率
            Dim NRJFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            If NUM2_NRJ > 0 And NRJFHL2 >= NRJFHL2_min Then
                '设备100%负荷时的额定本体天然气耗量
                Dim NRJTRQ2_ED As Double = NUM2_NRJ * TRQ2_ED_NRJ
                '在没有考虑修正系数前的，设备本体天然气耗量
                Dim NRJTRQ2 As Double = TRQHLXZXS_NRJ * NRJTRQ2_ED * NRJFHL2
                '在没有考虑修正系数前的，总天然气耗量
                NRJTRQ2_ALL_a = NRJTRQ2
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的效率修正系数、单台设备负荷率、设备启动数量
                Dim NRJZLXLXZ2_List As New List(Of Double)
                Dim FHL2_NRJ_single As New List(Of Double)
                Dim NUM2_NRJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_NRJ Step 1
                    '如果设备启动100%，发电量都不够，则跳入下一个循环，提高速度
                    If n2 * FDGL2_ED_NRJ < NRJFHL2 * FDGL2_ED_NRJ * NUM2_NRJ Then
                        GoTo fff
                    End If
                    For a2 = FHL2_min_NRJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100
                        '制冷功率满足需求
                        If a2 * n2 * FDGL2_ED_NRJ >= NRJFHL2 * FDGL2_ED_NRJ * NUM2_NRJ Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ2_List.Add(内燃机发电效率曲线(a2))
                            FHL2_NRJ_single.Add(a2)
                            NUM2_NRJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo fff
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的效率修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            NRJZLXLXZ2_List.Add(内燃机发电效率曲线(a2))
                            FHL2_NRJ_single.Add(a2)
                            NUM2_NRJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制冷满足要求的情况
                            GoTo fff
                        End If
                    Next
fff:
                Next
                '计算此时的效率修正系数所有结果中得到最大值
                Dim NRJZLXLXZ2 As Double = NRJZLXLXZ2_List.Max
                '经过修正后的设备本体天然气耗量
                NRJTRQ2_ALL_b = NRJTRQ2_ALL_a / NRJZLXLXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = NRJZLXLXZ2_List.IndexOf(NRJZLXLXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 127).Value = NUM2_NRJ_QD(COP_MAX_index)
            Else
                NRJTRQ2_ALL_a = 0
                NRJTRQ2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 127).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————               
            '风冷螺杆机（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim FLLGJFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value
            Dim FLLGJFHL1_XR As Double
            If NUM1_FLLGJ = 0 Then
                FLLGJFHL1_XR = 0
            Else
                FLLGJFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value / NUM1_FLLGJ
            End If
            Dim FLLGJFHL1 As Double = FLLGJFHL1_GR + FLLGJFHL1_XR
            If NUM1_FLLGJ > 0 And FLLGJFHL1 >= FLLGJFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim FLLGJBTHD1_ED As Double = NUM1_FLLGJ * BTHD1_ED_FLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim FLLGJBTHD1_GR As Double = BTHDXS_GR_air * FLLGJBTHD1_ED * FLLGJFHL1_GR
                Dim FLLGJBTHD1_XR As Double = BTHDXS_XR_air * FLLGJBTHD1_ED * FLLGJFHL1_XR
                '在没有考虑修正系数前的，总耗电功率
                FLLGJBTHD1_ALL_a = FLLGJBTHD1_GR + FLLGJBTHD1_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_FLLGJ_List As New List(Of Double)
                Dim FHL1_FLLGJ_single As New List(Of Double)
                Dim NUM1_FLLGJ_QD As New List(Of Double)
                For n1 = 0 To NUM1_FLLGJ Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_FLLGJ / NUM1_FLLGJ < FLLGJFHL1 * ZJRGL1_FLLGJ Then
                        GoTo ggg
                    End If
                    For a1 = FHL1_min_FLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_FLLGJ / NUM1_FLLGJ >= FLLGJFHL1 * ZJRGL1_FLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_FLLGJ_List.Add(风冷螺杆式热泵制热COP曲线(a1))
                            FHL1_FLLGJ_single.Add(a1)
                            NUM1_FLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ggg
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_FLLGJ_List.Add(风冷螺杆式热泵制热COP曲线(a1))
                            FHL1_FLLGJ_single.Add(a1)
                            NUM1_FLLGJ_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ggg
                        End If
                    Next
ggg:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim FLLGJZRCOPXZ1 As Double = COPXZ1_FLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                FLLGJBTHD1_ALL_b = FLLGJBTHD1_ALL_a / FLLGJZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_FLLGJ_List.IndexOf(FLLGJZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 130).Value = NUM1_FLLGJ_QD(COP_MAX_index)
            Else
                FLLGJBTHD1_ALL_a = 0
                FLLGJBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 130).Value = 0
            End If
            '风冷螺杆机（2）
            '设备的装机总数量，台数
            'Dim FLLGJZJSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 12).Value
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim FLLGJFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value
            Dim FLLGJFHL2_XR As Double
            If NUM2_FLLGJ = 0 Then
                FLLGJFHL2_XR = 0
            Else
                FLLGJFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value / NUM2_FLLGJ
            End If
            Dim FLLGJFHL2 As Double = FLLGJFHL2_GR + FLLGJFHL2_XR
            If NUM2_FLLGJ > 0 And FLLGJFHL2 >= FLLGJFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim FLLGJBTHD2_ED As Double = NUM2_FLLGJ * BTHD2_ED_FLLGJ
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim FLLGJBTHD2_GR As Double = BTHDXS_GR_air * FLLGJBTHD2_ED * FLLGJFHL2_GR
                Dim FLLGJBTHD2_XR As Double = BTHDXS_XR_air * FLLGJBTHD2_ED * FLLGJFHL2_XR
                '在没有考虑修正系数前的，总耗电功率
                FLLGJBTHD2_ALL_a = FLLGJBTHD2_GR + FLLGJBTHD2_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_FLLGJ_List As New List(Of Double)
                Dim FHL2_FLLGJ_single As New List(Of Double)
                Dim NUM2_FLLGJ_QD As New List(Of Double)
                For n2 = 0 To NUM2_FLLGJ Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_FLLGJ / NUM2_FLLGJ < FLLGJFHL2 * ZJRGL2_FLLGJ Then
                        GoTo hhh
                    End If
                    For a2 = FHL2_min_FLLGJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_FLLGJ / NUM2_FLLGJ >= FLLGJFHL2 * ZJRGL2_FLLGJ Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_FLLGJ_List.Add(风冷螺杆式热泵制热COP曲线(a2))
                            FHL2_FLLGJ_single.Add(a2)
                            NUM2_FLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo hhh
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_FLLGJ_List.Add(风冷螺杆式热泵制热COP曲线(a2))
                            FHL2_FLLGJ_single.Add(a2)
                            NUM2_FLLGJ_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo hhh
                        End If
                    Next
hhh:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim FLLGJZRCOPXZ2 As Double = COPXZ2_FLLGJ_List.Max
                '经过修正后的设备本体耗电功率
                FLLGJBTHD2_ALL_b = FLLGJBTHD2_ALL_a / FLLGJZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_FLLGJ_List.IndexOf(FLLGJZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 131).Value = NUM2_FLLGJ_QD(COP_MAX_index)
            Else
                FLLGJBTHD2_ALL_a = 0
                FLLGJBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 131).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————            
            '水（地）源热泵（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim SDYRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value
            Dim SDYRBFHL1_XR As Double
            If NUM1_SDYRB = 0 Then
                SDYRBFHL1_XR = 0
            Else
                SDYRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value / NUM1_SDYRB
            End If
            Dim SDYRBFHL1 As Double = SDYRBFHL1_GR + SDYRBFHL1_XR
            If NUM1_SDYRB > 0 And SDYRBFHL1 >= SDYRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SDYRBBTHD1_ED As Double = NUM1_SDYRB * BTHD1_ED_SDYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SDYRBBTHD1_GR As Double = BTHDXS_GR_water * SDYRBBTHD1_ED * SDYRBFHL1_GR
                Dim SDYRBBTHD1_XR As Double = BTHDXS_XR_water * SDYRBBTHD1_ED * SDYRBFHL1_XR
                '在没有考虑修正系数前的，总耗电功率
                SDYRBBTHD1_ALL_a = SDYRBBTHD1_GR + SDYRBBTHD1_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_SDYRB_List As New List(Of Double)
                Dim FHL1_SDYRB_single As New List(Of Double)
                Dim NUM1_SDYRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_SDYRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_SDYRB / NUM1_SDYRB < SDYRBFHL1 * ZJRGL1_SDYRB Then
                        GoTo iii
                    End If
                    For a1 = FHL1_min_SDYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_SDYRB / NUM1_SDYRB >= SDYRBFHL1 * ZJRGL1_SDYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SDYRB_List.Add(水_地源热泵制热COP曲线(a1))
                            FHL1_SDYRB_single.Add(a1)
                            NUM1_SDYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo iii
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_SDYRB_List.Add(水_地源热泵制热COP曲线(a1))
                            FHL1_SDYRB_single.Add(a1)
                            NUM1_SDYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo iii
                        End If
                    Next
iii:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim SDYRBZRCOPXZ1 As Double = COPXZ1_SDYRB_List.Max
                '经过修正后的设备本体耗电功率
                SDYRBBTHD1_ALL_b = SDYRBBTHD1_ALL_a / SDYRBZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_SDYRB_List.IndexOf(SDYRBZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 134).Value = NUM1_SDYRB_QD(COP_MAX_index)
            Else
                SDYRBBTHD1_ALL_a = 0
                SDYRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 134).Value = 0
            End If
            '水（地）源热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim SDYRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value
            Dim SDYRBFHL2_XR As Double
            If NUM2_SDYRB = 0 Then
                SDYRBFHL2_XR = 0
            Else
                SDYRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value / NUM2_SDYRB
            End If
            Dim SDYRBFHL2 As Double = SDYRBFHL2_GR + SDYRBFHL2_XR
            If NUM2_SDYRB > 0 And SDYRBFHL2 >= SDYRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim SDYRBBTHD2_ED As Double = NUM2_SDYRB * BTHD2_ED_SDYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim SDYRBBTHD2_GR As Double = BTHDXS_GR_water * SDYRBBTHD2_ED * SDYRBFHL2_GR
                Dim SDYRBBTHD2_XR As Double = BTHDXS_XR_water * SDYRBBTHD2_ED * SDYRBFHL2_XR
                '在没有考虑修正系数前的，总耗电功率
                SDYRBBTHD2_ALL_a = SDYRBBTHD2_GR + SDYRBBTHD2_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_SDYRB_List As New List(Of Double)
                Dim FHL2_SDYRB_single As New List(Of Double)
                Dim NUM2_SDYRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_SDYRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_SDYRB / NUM2_SDYRB < SDYRBFHL2 * ZJRGL2_SDYRB Then
                        GoTo jjj
                    End If
                    For a2 = FHL2_min_SDYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_SDYRB / NUM2_SDYRB >= SDYRBFHL2 * ZJRGL2_SDYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SDYRB_List.Add(水_地源热泵制热COP曲线(a2))
                            FHL2_SDYRB_single.Add(a2)
                            NUM2_SDYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo jjj
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_SDYRB_List.Add(水_地源热泵制热COP曲线(a2))
                            FHL2_SDYRB_single.Add(a2)
                            NUM2_SDYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo jjj
                        End If
                    Next
jjj:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim SDYRBZRCOPXZ2 As Double = COPXZ2_SDYRB_List.Max
                '经过修正后的设备本体耗电功率
                SDYRBBTHD2_ALL_b = SDYRBBTHD2_ALL_a / SDYRBZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_SDYRB_List.IndexOf(SDYRBZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 135).Value = NUM2_SDYRB_QD(COP_MAX_index)
            Else
                SDYRBBTHD2_ALL_a = 0
                SDYRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 135).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————            
            '空气源热泵（1）
            '设备的装机总数量，台数
            'Dim KQYRBZJSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 16).Value
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim KQYRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value
            Dim KQYRBFHL1_XR As Double
            If NUM1_KQYRB = 0 Then
                KQYRBFHL1_XR = 0
            Else
                KQYRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value / NUM1_KQYRB
            End If
            Dim KQYRBFHL1 As Double = KQYRBFHL1_GR + KQYRBFHL1_XR
            If NUM1_KQYRB > 0 And KQYRBFHL1 >= KQYRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim KQYRBBTHD1_ED As Double = NUM1_KQYRB * BTHD1_ED_KQYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim KQYRBBTHD1_GR As Double = BTHDXS_GR_air * KQYRBBTHD1_ED * KQYRBFHL1_GR
                Dim KQYRBBTHD1_XR As Double = BTHDXS_XR_air * KQYRBBTHD1_ED * KQYRBFHL1_XR
                '在没有考虑修正系数前的，总耗电功率
                KQYRBBTHD1_ALL_a = KQYRBBTHD1_GR + KQYRBBTHD1_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_KQYRB_List As New List(Of Double)
                Dim FHL1_KQYRB_single As New List(Of Double)
                Dim NUM1_KQYRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_KQYRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_KQYRB / NUM1_KQYRB < KQYRBFHL1 * ZJRGL1_KQYRB Then
                        GoTo kkk
                    End If
                    For a1 = FHL1_min_KQYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_KQYRB / NUM1_KQYRB >= KQYRBFHL1 * ZJRGL1_KQYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_KQYRB_List.Add(空气源热泵制热COP曲线(a1))
                            FHL1_KQYRB_single.Add(a1)
                            NUM1_KQYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo kkk
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_KQYRB_List.Add(空气源热泵制热COP曲线(a1))
                            FHL1_KQYRB_single.Add(a1)
                            NUM1_KQYRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo kkk
                        End If
                    Next
kkk:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim KQYRBZRCOPXZ1 As Double = COPXZ1_KQYRB_List.Max
                '经过修正后的设备本体耗电功率
                KQYRBBTHD1_ALL_b = KQYRBBTHD1_ALL_a / KQYRBZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_KQYRB_List.IndexOf(KQYRBZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 132).Value = NUM1_KQYRB_QD(COP_MAX_index)
            Else
                KQYRBBTHD1_ALL_a = 0
                KQYRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 132).Value = 0
            End If
            '空气源热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim KQYRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value
            Dim KQYRBFHL2_XR As Double
            If NUM2_KQYRB = 0 Then
                KQYRBFHL2_XR = 0
            Else
                KQYRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value / NUM2_KQYRB
            End If
            Dim KQYRBFHL2 As Double = KQYRBFHL2_GR + KQYRBFHL2_XR
            If NUM2_KQYRB > 0 And KQYRBFHL2 >= KQYRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim KQYRBBTHD2_ED As Double = NUM2_KQYRB * BTHD2_ED_KQYRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim KQYRBBTHD2_GR As Double = BTHDXS_GR_air * KQYRBBTHD2_ED * KQYRBFHL2_GR
                Dim KQYRBBTHD2_XR As Double = BTHDXS_XR_air * KQYRBBTHD2_ED * KQYRBFHL2_XR
                '在没有考虑修正系数前的，总耗电功率
                KQYRBBTHD2_ALL_a = KQYRBBTHD2_GR + KQYRBBTHD2_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_KQYRB_List As New List(Of Double)
                Dim FHL2_KQYRB_single As New List(Of Double)
                Dim NUM2_KQYRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_KQYRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_KQYRB / NUM2_KQYRB < KQYRBFHL2 * ZJRGL2_KQYRB Then
                        GoTo lll
                    End If
                    For a2 = FHL2_min_KQYRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_KQYRB / NUM2_KQYRB >= KQYRBFHL2 * ZJRGL2_KQYRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_KQYRB_List.Add(空气源热泵制热COP曲线(a2))
                            FHL2_KQYRB_single.Add(a2)
                            NUM2_KQYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo lll
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_KQYRB_List.Add(空气源热泵制热COP曲线(a2))
                            FHL2_KQYRB_single.Add(a2)
                            NUM2_KQYRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo lll
                        End If
                    Next
lll:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim KQYRBZRCOPXZ2 As Double = COPXZ2_KQYRB_List.Max
                '经过修正后的设备本体耗电功率
                KQYRBBTHD2_ALL_b = KQYRBBTHD2_ALL_a / KQYRBZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_KQYRB_List.IndexOf(KQYRBZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 133).Value = NUM2_KQYRB_QD(COP_MAX_index)
            Else
                KQYRBBTHD2_ALL_a = 0
                KQYRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 133).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '离心式热泵（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim LXSRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value
            Dim LXSRBFHL1_XR As Double
            If NUM1_LXSRB = 0 Then
                LXSRBFHL1_XR = 0
            Else
                LXSRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value / NUM1_LXSRB
            End If
            Dim LXSRBFHL1 As Double = LXSRBFHL1_GR + LXSRBFHL1_XR
            If NUM1_LXSRB > 0 And LXSRBFHL1 >= LXSRBFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSRBBTHD1_ED As Double = NUM1_LXSRB * BTHD1_ED_LXSRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSRBBTHD1_GR As Double = BTHDXS_GR_water * LXSRBBTHD1_ED * LXSRBFHL1_GR
                Dim LXSRBBTHD1_XR As Double = BTHDXS_XR_water * LXSRBBTHD1_ED * LXSRBFHL1_XR
                '在没有考虑修正系数前的，总耗电功率
                LXSRBBTHD1_ALL_a = LXSRBBTHD1_GR + LXSRBBTHD1_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_LXSRB_List As New List(Of Double)
                Dim FHL1_LXSRB_single As New List(Of Double)
                Dim NUM1_LXSRB_QD As New List(Of Double)
                For n1 = 0 To NUM1_LXSRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_LXSRB / NUM1_LXSRB < LXSRBFHL1 * ZJRGL1_LXSRB Then
                        GoTo mmm
                    End If
                    For a1 = FHL1_min_LXSRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_LXSRB / NUM1_LXSRB >= LXSRBFHL1 * ZJRGL1_LXSRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSRB_List.Add(离心式热泵制热COP曲线(a1))
                            FHL1_LXSRB_single.Add(a1)
                            NUM1_LXSRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo mmm
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_LXSRB_List.Add(离心式热泵制热COP曲线(a1))
                            FHL1_LXSRB_single.Add(a1)
                            NUM1_LXSRB_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo mmm
                        End If
                    Next
mmm:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim LXSRBZRCOPXZ1 As Double = COPXZ1_LXSRB_List.Max
                '经过修正后的设备本体耗电功率
                LXSRBBTHD1_ALL_b = LXSRBBTHD1_ALL_a / LXSRBZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_LXSRB_List.IndexOf(LXSRBZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 136).Value = NUM1_LXSRB_QD(COP_MAX_index)
            Else
                LXSRBBTHD1_ALL_a = 0
                LXSRBBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 136).Value = 0
            End If
            '离心式热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim LXSRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value
            Dim LXSRBFHL2_XR As Double
            If NUM2_LXSRB = 0 Then
                LXSRBFHL2_XR = 0
            Else
                LXSRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value / NUM2_LXSRB
            End If
            Dim LXSRBFHL2 As Double = LXSRBFHL2_GR + LXSRBFHL2_XR
            If NUM2_LXSRB > 0 And LXSRBFHL2 >= LXSRBFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                Dim LXSRBBTHD2_ED As Double = NUM2_LXSRB * BTHD2_ED_LXSRB
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim LXSRBBTHD2_GR As Double = BTHDXS_GR_water * LXSRBBTHD2_ED * LXSRBFHL2_GR
                Dim LXSRBBTHD2_XR As Double = BTHDXS_XR_water * LXSRBBTHD2_ED * LXSRBFHL2_XR
                '在没有考虑修正系数前的，总耗电功率
                LXSRBBTHD2_ALL_a = LXSRBBTHD2_GR + LXSRBBTHD2_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_LXSRB_List As New List(Of Double)
                Dim FHL2_LXSRB_single As New List(Of Double)
                Dim NUM2_LXSRB_QD As New List(Of Double)
                For n2 = 0 To NUM2_LXSRB Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_LXSRB / NUM2_LXSRB < LXSRBFHL2 * ZJRGL2_LXSRB Then
                        GoTo nnn
                    End If
                    For a2 = FHL2_min_LXSRB To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_LXSRB / NUM2_LXSRB >= LXSRBFHL2 * ZJRGL2_LXSRB Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSRB_List.Add(离心式热泵制热COP曲线(a2))
                            FHL2_LXSRB_single.Add(a2)
                            NUM2_LXSRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo nnn
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_LXSRB_List.Add(离心式热泵制热COP曲线(a2))
                            FHL2_LXSRB_single.Add(a2)
                            NUM2_LXSRB_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo nnn
                        End If
                    Next
nnn:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim LXSRBZRCOPXZ2 As Double = COPXZ2_LXSRB_List.Max
                '经过修正后的设备本体耗电功率
                LXSRBBTHD2_ALL_b = LXSRBBTHD2_ALL_a / LXSRBZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_LXSRB_List.IndexOf(LXSRBZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 137).Value = NUM2_LXSRB_QD(COP_MAX_index)
            Else
                LXSRBBTHD2_ALL_a = 0
                LXSRBBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 137).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '电采暖锅炉（1）
            '设备的装机总功率
            'Dim DGLZJZGL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 20).Value
            '此时的设备负荷率=供热负荷率+蓄热负荷率
            Dim DGLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value
            Dim DGLFHL1_XR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value
            Dim DGLFHL1 As Double = DGLFHL1_GR + DGLFHL1_XR
            If ZJRGL1_DGL > 0 And DGLFHL1 >= DGLFHL1_min Then
                '设备100%负荷时的额定本体耗电功率
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim DGLBTHD1_GR As Double = BTHD1_ED_DGL * DGLFHL1_GR
                Dim DGLBTHD1_XR As Double = BTHD1_ED_DGL * DGLFHL1_XR
                '在没有考虑修正系数前的，总耗电功率
                DGLBTHD1_ALL_a = DGLBTHD1_GR + DGLBTHD1_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ1_DGL_List As New List(Of Double)
                Dim FHL1_DGL_single As New List(Of Double)
                Dim NUM1_DGL_QD As New List(Of Double)
                For n1 = 0 To NUM1_DGL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n1 * ZJRGL1_DGL / NUM1_DGL < DGLFHL1 * ZJRGL1_DGL Then
                        GoTo ooo
                    End If
                    For a1 = FHL1_min_DGL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a1 * n1 * ZJRGL1_DGL / NUM1_DGL >= DGLFHL1 * ZJRGL1_DGL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_DGL_List.Add(电采暖锅炉制热效率曲线(a1))
                            FHL1_DGL_single.Add(a1)
                            NUM1_DGL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ooo
                        ElseIf a1 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a1）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ1_DGL_List.Add(电采暖锅炉制热效率曲线(a1))
                            FHL1_DGL_single.Add(a1)
                            NUM1_DGL_QD.Add(n1)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ooo
                        End If
                    Next
ooo:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim DGLZRCOPXZ1 As Double = COPXZ1_DGL_List.Max
                '经过修正后的设备本体耗电功率
                DGLBTHD1_ALL_b = DGLBTHD1_ALL_a / DGLZRCOPXZ1
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ1_DGL_List.IndexOf(DGLZRCOPXZ1)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 140).Value = NUM1_DGL_QD(COP_MAX_index)
            Else
                DGLBTHD1_ALL_a = 0
                DGLBTHD1_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 140).Value = 0
            End If
            '电采暖锅炉（2）
            '设备的装机总功率
            'Dim DGLZJZGL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 20).Value
            '此时的设备负荷率=供热负荷率+蓄热负荷率
            Dim DGLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value
            Dim DGLFHL2_XR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value
            Dim DGLFHL2 As Double = DGLFHL2_GR + DGLFHL2_XR
            If ZJRGL2_DGL > 0 And DGLFHL2 >= DGLFHL2_min Then
                '设备100%负荷时的额定本体耗电功率
                '在没有考虑修正系数前的，设备本体耗电功率
                Dim DGLBTHD2_GR As Double = BTHD2_ED_DGL * DGLFHL2_GR
                Dim DGLBTHD2_XR As Double = BTHD2_ED_DGL * DGLFHL2_XR
                '在没有考虑修正系数前的，总耗电功率
                DGLBTHD2_ALL_a = DGLBTHD2_GR + DGLBTHD2_XR
                '用循环反算出单台设备负荷率最优解和此时的设备启动数量
                '列表，储存计算出的COP修正系数、单台设备负荷率、设备启动数量
                Dim COPXZ2_DGL_List As New List(Of Double)
                Dim FHL2_DGL_single As New List(Of Double)
                Dim NUM2_DGL_QD As New List(Of Double)
                For n2 = 0 To NUM2_DGL Step 1
                    '如果设备启动100%，制热量都不够，则跳入下一个循环，提高速度
                    If n2 * ZJRGL2_DGL / NUM2_DGL < DGLFHL2 * ZJRGL2_DGL Then
                        GoTo ppp
                    End If
                    For a2 = FHL2_min_DGL To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 '单台负荷率
                        '制热功率满足需求
                        If a2 * n2 * ZJRGL2_DGL / NUM2_DGL >= DGLFHL2 * ZJRGL2_DGL Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_DGL_List.Add(电采暖锅炉制热效率曲线(a2))
                            FHL2_DGL_single.Add(a2)
                            NUM2_DGL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ppp
                        ElseIf a2 >= 1 Then
                            '计算此时单台设备的COP修正系数（此时负荷率是a2）、单台设备负荷率、设备启动数量，并加入列表
                            COPXZ2_DGL_List.Add(电采暖锅炉制热效率曲线(a2))
                            FHL2_DGL_single.Add(a2)
                            NUM2_DGL_QD.Add(n2)
                            '结束这一次循环，目的是仅计算第一个制热满足要求的情况
                            GoTo ppp
                        End If
                    Next
ppp:
                Next
                '计算此时的制热COP修正系数所有结果中得到最大值
                Dim DGLZRCOPXZ2 As Double = COPXZ2_DGL_List.Max
                '经过修正后的设备本体耗电功率
                DGLBTHD2_ALL_b = DGLBTHD2_ALL_a / DGLZRCOPXZ2
                '找到所在标签
                Dim COP_MAX_index As Integer = COPXZ2_DGL_List.IndexOf(DGLZRCOPXZ2)
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 141).Value = NUM2_DGL_QD(COP_MAX_index)
            Else
                DGLBTHD2_ALL_a = 0
                DGLBTHD2_ALL_b = 0
                '设备启动数量写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 141).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '参与混水供热的设备
            '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
            Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
            '混水设备装机数量
            Dim NUM1_HS As Double = 0
            Dim NUM2_HS As Double = 0
            '混水设备（1）（2）装机功率
            Dim ZJRGL1_HS As Double = 0
            Dim ZJRGL2_HS As Double = 0
            '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
            Dim ZJRGL_HS_ALL As Double = 0
            '参与混水供热的设备的本体耗电功率和
            Dim BTHD1_ED_HS As Double = 0
            Dim BTHD2_ED_HS As Double = 0
            '本体耗电修正系数
            Dim BTHD_XZXS_HS As Double = 1
            '混水设备负荷率下限（单台）
            Dim FHL1_min_HS As Double
            Dim FHL2_min_HS As Double
            '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                '装机制热功率
                ZJRGL_HS_ALL = ZJRGL1_KQYRB + ZJRGL2_KQYRB + RCXS_a
                ZJRGL1_HS = ZJRGL1_KQYRB + RCXS_a
                ZJRGL2_HS = ZJRGL2_KQYRB + RCXS_a
                '设备（1）装机数量
                NUM1_HS = NUM1_KQYRB
                '设备（2）装机数量
                NUM2_HS = NUM2_KQYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_KQYRB
                BTHD2_ED_HS = BTHD2_ED_KQYRB
                '混水设备负荷率下限（单台）
                FHL1_min_HS = FHL1_min_KQYRB
                FHL2_min_HS = FHL2_min_KQYRB
                '本体耗电修正系数
                BTHD_XZXS_HS = BTHDXS_GR_air
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                '制热功率
                ZJRGL_HS_ALL = ZJRGL1_SDYRB + ZJRGL2_SDYRB + RCXS_a
                ZJRGL1_HS = ZJRGL1_SDYRB + RCXS_a
                ZJRGL2_HS = ZJRGL2_SDYRB + RCXS_a
                '设备（1）装机数量
                NUM1_HS = NUM1_SDYRB
                '设备（2）装机数量
                NUM2_HS = NUM2_SDYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_SDYRB
                BTHD2_ED_HS = BTHD2_ED_SDYRB
                '混水设备负荷率下限（单台）
                FHL1_min_HS = FHL1_min_SDYRB
                FHL2_min_HS = FHL2_min_SDYRB
                '本体耗电修正系数
                BTHD_XZXS_HS = BTHDXS_GR_water
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                '制热功率
                ZJRGL_HS_ALL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ + RCXS_a
                ZJRGL1_HS = ZJRGL1_FLLGJ + RCXS_a
                ZJRGL2_HS = ZJRGL2_FLLGJ + RCXS_a
                '设备（1）装机数量
                NUM1_HS = NUM1_FLLGJ
                '设备（2）装机数量
                NUM2_HS = NUM2_FLLGJ
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_FLLGJ
                BTHD2_ED_HS = BTHD2_ED_FLLGJ
                '混水设备负荷率下限（单台）
                FHL1_min_HS = FHL1_min_FLLGJ
                FHL2_min_HS = FHL2_min_FLLGJ
                '本体耗电修正系数
                BTHD_XZXS_HS = BTHDXS_GR_air
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '混水供热此时实际的设备总耗电
            Dim HD_ALL_HS_now As Double = 0
            '混水设备负荷率
            Dim FHL_HS_now As Double = 0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing Then
                '混水设备当前计算出的负荷率
                FHL_HS_now = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value
                '混水供热的设备计算
                Dim GRGL_HS_now As Double = FHL_HS_now * ZJRGL_HS_ALL
                '设备本体耗电修正系数
                Dim XZXS_HS As Double = 1
                '没有考虑COP修正时的本体耗电
                BTHD_ALL_HS_a = FHL_HS_now * (NUM1_HS * BTHD1_ED_HS + NUM2_HS * BTHD2_ED_HS) * BTHD_XZXS_HS
                '定义列表，储存混水寻优计算结果
                '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                Dim HD_ALL_HS As New List(Of Double）
                Dim FHL1_HS_single As New List(Of Double)
                Dim FHL2_HS_single As New List(Of Double)
                Dim NUM1_HS_QD As New List(Of Double)
                Dim NUM2_HS_QD As New List(Of Double)
                '参与混水的设备也需要进行寻优计算
                '穷举设备数量
                For n1_hs = 0 To NUM1_HS Step 1
                    For n2_hs = 0 To NUM2_HS Step 1
                        '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                        Dim RFH1_ALL_HS_temp As Double
                        If NUM1_HS = 0 Then
                            RFH1_ALL_HS_temp = 0
                        Else
                            RFH1_ALL_HS_temp = n1_hs * ZJRGL1_HS / NUM1_HS
                        End If
                        Dim RFH2_ALL_HS_temp As Double
                        If NUM2_HS = 0 Then
                            RFH2_ALL_HS_temp = 0
                        Else
                            RFH2_ALL_HS_temp = n2_hs * ZJRGL2_HS / NUM2_HS
                        End If
                        If (RFH1_ALL_HS_temp + RFH2_ALL_HS_temp) < GRGL_HS_now Then
                            GoTo qqq
                        End If
                        '穷举设备负荷率
                        For a1_hs = FHL1_min_HS To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1_hs表示设备(1)负荷率（单台）
                            For a2_hs = FHL2_min_HS To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a2_hs表示设备(2)负荷率（单台）
                                '计算设备（1）和设备（2）本体的效率修正系数
                                Dim XZXS1_HS As Double = 1
                                Dim XZXS2_HS As Double = 1
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                                    XZXS1_HS = 空气源热泵制热COP曲线(a1_hs)
                                    XZXS2_HS = 空气源热泵制热COP曲线(a2_hs)
                                End If
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                                    XZXS1_HS = 水_地源热泵制热COP曲线(a1_hs)
                                    XZXS2_HS = 水_地源热泵制热COP曲线(a2_hs)
                                End If
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                                    XZXS1_HS = 风冷螺杆式热泵制热COP曲线(a1_hs)
                                    XZXS2_HS = 风冷螺杆式热泵制热COP曲线(a2_hs)
                                End If
                                '混水设备的总和本体耗电系数
                                Dim BTHDXS_ZH_HS As Double = BTHD_XZXS_HS
                                '计算设备（1）和设备（2）本体的耗电功率
                                Dim BTHD1_HS_now As Double = BTHDXS_ZH_HS * n1_hs * a1_hs * BTHD1_ED_HS / XZXS1_HS
                                Dim BTHD2_HS_now As Double = BTHDXS_ZH_HS * n2_hs * a2_hs * BTHD2_ED_HS / XZXS2_HS
                                '计算此时的总耗电功率
                                Dim BTZHD_HS As Double = BTHD1_HS_now + BTHD2_HS_now
                                '计算此时的混水总出力
                                Dim RGL1_out_now_HS As Double
                                If ZJRGL1_HS <= RCXS_a Then
                                    RGL1_out_now_HS = 0
                                Else
                                    If NUM1_HS = 0 Then
                                        RGL1_out_now_HS = 0
                                    Else
                                        RGL1_out_now_HS = n1_hs * a1_hs * ZJRGL1_HS / NUM1_HS
                                    End If
                                End If
                                Dim RGL2_out_now_HS As Double
                                If ZJRGL2_HS <= RCXS_a Then
                                    RGL2_out_now_HS = 0
                                Else
                                    If NUM2_HS = 0 Then
                                        RGL2_out_now_HS = 0
                                    Else
                                        RGL2_out_now_HS = n2_hs * a2_hs * ZJRGL2_HS / NUM2_HS
                                    End If
                                End If
                                Dim RGL_out_all_HS As Double = RGL1_out_now_HS + RGL2_out_now_HS
                                '如果达到了热负荷需求，则跳出内层循环
                                If RGL_out_all_HS >= GRGL_HS_now Then
                                    '计算结果加入列表
                                    HD_ALL_HS.Add(BTZHD_HS)
                                    FHL1_HS_single.Add(a1_hs)
                                    FHL2_HS_single.Add(a2_hs)
                                    NUM1_HS_QD.Add(n1_hs)
                                    NUM2_HS_QD.Add(n2_hs)
                                    '跳出循环
                                    Exit For
                                ElseIf a1_hs >= 1 And a2_hs >= 1 Then
                                    '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                    '计算结果加入列表
                                    HD_ALL_HS.Add(BTZHD_HS)
                                    FHL1_HS_single.Add(a1_hs)
                                    FHL2_HS_single.Add(a2_hs)
                                    NUM1_HS_QD.Add(n1_hs)
                                    NUM2_HS_QD.Add(n2_hs)
                                    '跳出循环
                                    Exit For
                                End If
                            Next
                        Next
qqq:
                    Next
                Next
                '考虑修正后的总耗电
                BTHD_ALL_HS_b = HD_ALL_HS.Min
                '找到所在标签
                Dim COP_MAX_index As Integer = HD_ALL_HS.IndexOf(BTHD_ALL_HS_b)
                '设备启动数量写入Excel
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 132).Value = NUM1_HS_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 133).Value = NUM2_HS_QD(COP_MAX_index)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 134).Value = NUM1_HS_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 135).Value = NUM2_HS_QD(COP_MAX_index)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 130).Value = NUM1_HS_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 131).Value = NUM2_HS_QD(COP_MAX_index)
                End If
            Else
                BTHD_ALL_HS_a = 0
                BTHD_ALL_HS_b = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '参与梯级供热的设备
            '梯级设备装机数量
            Dim NUM1_TJ As Double = 0
            Dim NUM2_TJ As Double = 0
            '梯级设备（1）（2）装机功率
            Dim ZJRGL1_TJ As Double = 0
            Dim ZJRGL2_TJ As Double = 0
            '参与梯级的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
            Dim ZJRGL_TJ_ALL As Double = 0
            '参与梯级供热的设备的本体耗电功率和
            Dim BTHD1_ED_TJ As Double = 0
            Dim BTHD2_ED_TJ As Double = 0
            '本体耗电修正系数
            Dim BTHD_XZXS_TJ As Double = 1
            '梯级设备负荷率下限（单台）
            Dim FHL1_min_TJ As Double
            Dim FHL2_min_TJ As Double
            '梯级设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为梯级设备的装机总功率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "空气源热泵" Then
                '装机制热功率
                ZJRGL_TJ_ALL = ZJRGL1_KQYRB + ZJRGL2_KQYRB + RCXS_a
                ZJRGL1_TJ = ZJRGL1_KQYRB + RCXS_a
                ZJRGL2_TJ = ZJRGL2_KQYRB + RCXS_a
                '设备（1）装机数量
                NUM1_TJ = NUM1_KQYRB
                '设备（2）装机数量
                NUM2_TJ = NUM2_KQYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_KQYRB
                BTHD2_ED_TJ = BTHD2_ED_KQYRB
                '梯级设备负荷率下限（单台）
                FHL1_min_TJ = FHL1_min_KQYRB
                FHL2_min_TJ = FHL2_min_KQYRB
                '本体耗电修正系数
                BTHD_XZXS_TJ = BTHDXS_GR_air
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "水(地)源热泵" Then
                '制热功率
                ZJRGL_TJ_ALL = ZJRGL1_SDYRB + ZJRGL2_SDYRB + RCXS_a
                ZJRGL1_TJ = ZJRGL1_SDYRB + RCXS_a
                ZJRGL2_TJ = ZJRGL2_SDYRB + RCXS_a
                '设备（1）装机数量GRGL_TJ_now
                NUM1_TJ = NUM1_SDYRB
                '设备（2）装机数量
                NUM2_TJ = NUM2_SDYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_SDYRB
                BTHD2_ED_TJ = BTHD2_ED_SDYRB
                '梯级设备负荷率下限（单台）
                FHL1_min_TJ = FHL1_min_SDYRB
                FHL2_min_TJ = FHL2_min_SDYRB
                '本体耗电修正系数
                BTHD_XZXS_TJ = BTHDXS_GR_water
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "风冷螺杆机" Then
                '制热功率
                ZJRGL_TJ_ALL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ + RCXS_a
                ZJRGL1_TJ = ZJRGL1_FLLGJ + RCXS_a
                ZJRGL2_TJ = ZJRGL2_FLLGJ + RCXS_a
                '设备（1）装机数量
                NUM1_TJ = NUM1_FLLGJ
                '设备（2）装机数量
                NUM2_TJ = NUM2_FLLGJ
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_FLLGJ
                BTHD2_ED_TJ = BTHD2_ED_FLLGJ
                '梯级设备负荷率下限（单台）
                FHL1_min_TJ = FHL1_min_FLLGJ
                FHL2_min_TJ = FHL2_min_FLLGJ
                '本体耗电修正系数
                BTHD_XZXS_TJ = BTHDXS_GR_air
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '梯级供热此时实际的设备总耗电
            Dim HD_ALL_TJ_now As Double = 0
            '梯级设备负荷率
            Dim FHL_TJ_now As Double = 0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing Then
                '梯级设备当前计算出的负荷率
                FHL_TJ_now = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value
                '设备本体耗电修正系数
                Dim XZXS_TJ As Double = 1
                '梯级供热的设备计算
                Dim GRGL_TJ_now As Double = FHL_TJ_now * ZJRGL_TJ_ALL
                '没有考虑COP修正时的本体耗电
                BTHD_ALL_TJ_a = FHL_TJ_now * (NUM1_TJ * BTHD1_ED_TJ + NUM2_TJ * BTHD2_ED_TJ) * BTHD_XZXS_TJ
                '定义列表，储存梯级寻优计算结果
                '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                Dim HD_ALL_TJ As New List(Of Double）
                Dim FHL1_TJ_single As New List(Of Double)
                Dim FHL2_TJ_single As New List(Of Double)
                Dim NUM1_TJ_QD As New List(Of Double)
                Dim NUM2_TJ_QD As New List(Of Double)
                '参与梯级的设备也需要进行寻优计算
                '穷举设备数量
                For n1_TJ = 0 To NUM1_TJ Step 1
                    For n2_TJ = 0 To NUM2_TJ Step 1
                        '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                        Dim RFH1_ALL_TJ_temp As Double
                        If NUM1_TJ = 0 Then
                            RFH1_ALL_TJ_temp = 0
                        Else
                            RFH1_ALL_TJ_temp = n1_TJ * ZJRGL1_TJ / NUM1_TJ
                        End If
                        Dim RFH2_ALL_TJ_temp As Double
                        If NUM2_TJ = 0 Then
                            RFH2_ALL_TJ_temp = 0
                        Else
                            RFH2_ALL_TJ_temp = n2_TJ * ZJRGL2_TJ / NUM2_TJ
                        End If
                        If (RFH1_ALL_TJ_temp + RFH2_ALL_TJ_temp) < GRGL_TJ_now Then
                            GoTo rrr
                        End If
                        '穷举设备负荷率
                        For a1_TJ = FHL1_min_TJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1_TJ表示设备(1)负荷率（单台）
                            For a2_TJ = FHL2_min_TJ To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a2_TJ表示设备(2)负荷率（单台）
                                '计算设备（1）和设备（2）本体的效率修正系数
                                Dim XZXS1_TJ As Double = 1
                                Dim XZXS2_TJ As Double = 1
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "空气源热泵" Then
                                    XZXS1_TJ = 空气源热泵制热COP曲线(a1_TJ)
                                    XZXS2_TJ = 空气源热泵制热COP曲线(a2_TJ)
                                End If
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "水(地)源热泵" Then
                                    XZXS1_TJ = 水_地源热泵制热COP曲线(a1_TJ)
                                    XZXS2_TJ = 水_地源热泵制热COP曲线(a2_TJ)
                                End If
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "风冷螺杆机" Then
                                    XZXS1_TJ = 风冷螺杆式热泵制热COP曲线(a1_TJ)
                                    XZXS2_TJ = 风冷螺杆式热泵制热COP曲线(a2_TJ)
                                End If
                                '梯级设备的总和本体耗电系数
                                Dim BTHDXS_ZH_TJ As Double = BTHD_XZXS_TJ
                                '计算设备（1）和设备（2）本体的耗电功率
                                Dim BTHD1_TJ_now As Double = BTHDXS_ZH_TJ * n1_TJ * a1_TJ * BTHD1_ED_TJ / XZXS1_TJ
                                Dim BTHD2_TJ_now As Double = BTHDXS_ZH_TJ * n2_TJ * a2_TJ * BTHD2_ED_TJ / XZXS2_TJ
                                '计算此时的总耗电功率
                                Dim BTZHD_TJ As Double = BTHD1_TJ_now + BTHD2_TJ_now
                                '计算此时的梯级总出力
                                Dim RGL1_out_now_TJ As Double
                                If ZJRGL1_TJ <= RCXS_a Then
                                    RGL1_out_now_TJ = 0
                                Else
                                    If NUM1_TJ = 0 Then
                                        RGL1_out_now_TJ = 0
                                    Else
                                        RGL1_out_now_TJ = n1_TJ * a1_TJ * ZJRGL1_TJ / NUM1_TJ
                                    End If
                                End If
                                Dim RGL2_out_now_TJ As Double
                                If ZJRGL2_TJ <= RCXS_a Then
                                    RGL2_out_now_TJ = 0
                                Else
                                    If NUM2_TJ = 0 Then
                                        RGL2_out_now_TJ = 0
                                    Else
                                        RGL2_out_now_TJ = n2_TJ * a2_TJ * ZJRGL2_TJ / NUM2_TJ
                                    End If
                                End If
                                Dim RGL_out_all_TJ As Double = RGL1_out_now_TJ + RGL2_out_now_TJ
                                '如果达到了热负荷需求，则跳出内层循环
                                If RGL_out_all_TJ >= GRGL_TJ_now Then
                                    '计算结果加入列表
                                    HD_ALL_TJ.Add(BTZHD_TJ)
                                    FHL1_TJ_single.Add(a1_TJ)
                                    FHL2_TJ_single.Add(a2_TJ)
                                    NUM1_TJ_QD.Add(n1_TJ)
                                    NUM2_TJ_QD.Add(n2_TJ)
                                    '跳出循环
                                    Exit For
                                ElseIf a1_TJ >= 1 And a2_TJ >= 1 Then
                                    '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                    '计算结果加入列表
                                    HD_ALL_TJ.Add(BTZHD_TJ)
                                    FHL1_TJ_single.Add(a1_TJ)
                                    FHL2_TJ_single.Add(a2_TJ)
                                    NUM1_TJ_QD.Add(n1_TJ)
                                    NUM2_TJ_QD.Add(n2_TJ)
                                    '跳出循环
                                    Exit For
                                End If
                            Next
                        Next
rrr:
                    Next
                Next
                '考虑修正后的总耗电
                BTHD_ALL_TJ_b = HD_ALL_TJ.Min
                '找到所在标签
                Dim COP_MAX_index As Integer = HD_ALL_TJ.IndexOf(BTHD_ALL_HS_b)
                '设备启动数量写入Excel
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "空气源热泵" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 132).Value = NUM1_TJ_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 133).Value = NUM2_TJ_QD(COP_MAX_index)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "水(地)源热泵" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 134).Value = NUM1_TJ_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 135).Value = NUM2_TJ_QD(COP_MAX_index)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "风冷螺杆机" Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 130).Value = NUM1_TJ_QD(COP_MAX_index)
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 131).Value = NUM2_TJ_QD(COP_MAX_index)
                End If
            Else
                BTHD_ALL_TJ_a = 0
                BTHD_ALL_TJ_b = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算空调设备本体耗电综合修正系数
            '空调设备在考虑修正系数前的本体总耗电功率
            Dim ZHD_ALL_a As Double = FLLGJBTHD1_ALL_a + FLLGJBTHD2_ALL_a + SDYRBBTHD1_ALL_a + SDYRBBTHD2_ALL_a + KQYRBBTHD1_ALL_a + KQYRBBTHD2_ALL_a + LXSRBBTHD1_ALL_a + LXSRBBTHD2_ALL_a + DGLBTHD1_ALL_a + DGLBTHD2_ALL_a + BTHD_ALL_HS_a + BTHD_ALL_TJ_a
            '空调设备在考虑修正系数之后的本体总耗电功率
            Dim ZHD_ALL_b As Double = FLLGJBTHD1_ALL_b + FLLGJBTHD2_ALL_b + SDYRBBTHD1_ALL_b + SDYRBBTHD2_ALL_b + KQYRBBTHD1_ALL_b + KQYRBBTHD2_ALL_b + LXSRBBTHD1_ALL_b + LXSRBBTHD2_ALL_b + DGLBTHD1_ALL_b + DGLBTHD2_ALL_b + BTHD_ALL_HS_b + BTHD_ALL_TJ_b
            '设备本体耗电功率综合修正系数
            Dim BTHD_XZ As Double = 1
            If ZHD_ALL_a > 0 And ZHD_ALL_b > 0 Then
                BTHD_XZ = ZHD_ALL_b / ZHD_ALL_a
            Else
                '默认值等于1
                BTHD_XZ = 1
            End If
            '天然气耗电综合修正系数
            Dim ZHTRQ_ALL_a As Double = ZRXXHLBTTRQ1_ALL_a + ZRXXHLBTTRQ2_ALL_a + TRQGLBTTRQ1_ALL_a + TRQGLBTTRQ2_ALL_a + NRJTRQ1_ALL_a + NRJTRQ2_ALL_a
            Dim ZHTRQ_ALL_b As Double = ZRXXHLBTTRQ1_ALL_b + ZRXXHLBTTRQ2_ALL_b + TRQGLBTTRQ1_ALL_b + TRQGLBTTRQ2_ALL_b + NRJTRQ1_ALL_b + NRJTRQ2_ALL_b
            Dim TRQ_XZ As Double = 1
            If ZHTRQ_ALL_a > 0 And ZHTRQ_ALL_b > 0 Then
                TRQ_XZ = ZHTRQ_ALL_b / ZHTRQ_ALL_a
            Else
                '默认值等于1
                TRQ_XZ = 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '将综合修正系数写入Excel
            '制热季系数
            '制热设备本体耗电修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 108).Value = BTHD_XZ
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 108).Value = BTHD_XZ
            '制热天然气消耗修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 109).Value = TRQ_XZ
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 109).Value = TRQ_XZ
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算出的结果保留3位小数，增加美观度
            For i = 108 To 109
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '如果本工况没有选择内燃机，则将内燃机及其余热利用的各种系数设置为1
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 102).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 102).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 103).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 103).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = 1
            End If
        Else
            '常规计算模式，修正系数全部设置为1
            '制热设备本体耗电修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 108).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 108).Value = 1
            '制热天然气消耗修正
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 109).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 109).Value = 1
            '内燃机及其余热利用系数
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 102), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105)).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 102), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105)).Value = 1
        End If
    End Sub
    Sub 制冷和蓄冷空调设备负荷率修正(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '常规计算模式就不进行修正了
        If calculation_mode = 2 Then
            '读取制冷季装机方案及参数
            Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
            '内燃机
            Dim NUM1_NRJ As Double = ans_ZJFA_L(0)
            Dim NUM2_NRJ As Double = ans_ZJFA_L(1)
            '离心式冷水机
            Dim NUM1_LXSLSJ As Double = ans_ZJFA_L(10)
            Dim NUM2_LXSLSJ As Double = ans_ZJFA_L(11)
            '水冷螺杆机
            Dim NUM1_SLLGJ As Double = ans_ZJFA_L(18)
            Dim NUM2_SLLGJ As Double = ans_ZJFA_L(19)
            '风冷螺杆机
            Dim NUM1_FLLGJ As Double = ans_ZJFA_L(26)
            Dim NUM2_FLLGJ As Double = ans_ZJFA_L(27)
            '水地源热泵
            Dim NUM1_SDYRB As Double = ans_ZJFA_L(34)
            Dim NUM2_SDYRB As Double = ans_ZJFA_L(35)
            '离心式热泵
            Dim NUM1_LXSRB As Double = ans_ZJFA_L(42)
            Dim NUM2_LXSRB As Double = ans_ZJFA_L(43)
            '空气源热泵
            Dim NUM1_KQYRB As Double = ans_ZJFA_L(50)
            Dim NUM2_KQYRB As Double = ans_ZJFA_L(51)
            '直燃型溴化锂
            Dim NUM1_ZRXXHL As Double = ans_ZJFA_L(58)
            Dim NUM2_ZRXXHL As Double = ans_ZJFA_L(59)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '————————————————————————————————————————————————————————————————————————————————————————  
            '参与制冷和蓄冷的空调设备，根据设置的设备允许最低负荷率，对计算出的负荷率进行修正
            '进行修正的负荷率=供冷负荷率+蓄冷负荷率
            '各设备允许运行的最低负荷率
            '离心式冷水机
            '设备最低允许运行的负荷率
            Dim LXSLSJFHL1_min As Double
            Dim LXSLSJFHL2_min As Double
            If NUM1_LXSLSJ > 0 Then
                LXSLSJFHL1_min = FHL1_min_LXSLSJ / NUM1_LXSLSJ
            Else
                LXSLSJFHL1_min = 0
            End If
            If NUM2_LXSLSJ > 0 Then
                LXSLSJFHL2_min = FHL2_min_LXSLSJ / NUM2_LXSLSJ
            Else
                LXSLSJFHL2_min = 0
            End If
            '水冷螺杆机
            '设备最低允许运行的负荷率
            Dim SLLGJFHL1_min As Double
            Dim SLLGJFHL2_min As Double
            If NUM1_SLLGJ > 0 Then
                SLLGJFHL1_min = FHL1_min_SLLGJ / NUM1_SLLGJ
            Else
                SLLGJFHL1_min = 0
            End If
            If NUM2_SLLGJ > 0 Then
                SLLGJFHL2_min = FHL2_min_SLLGJ / NUM2_SLLGJ
            Else
                SLLGJFHL2_min = 0
            End If
            '风冷螺杆机
            '设备最低允许运行的负荷率
            Dim FLLGJFHL1_min As Double
            Dim FLLGJFHL2_min As Double
            If NUM1_FLLGJ > 0 Then
                FLLGJFHL1_min = FHL1_min_FLLGJ / NUM1_FLLGJ
            Else
                FLLGJFHL1_min = 0
            End If
            If NUM2_FLLGJ > 0 Then
                FLLGJFHL2_min = FHL2_min_FLLGJ / NUM2_FLLGJ
            Else
                FLLGJFHL2_min = 0
            End If
            '水（地）源热泵
            '设备最低允许运行的负荷率
            Dim SDYRBFHL1_min As Double
            Dim SDYRBFHL2_min As Double
            If NUM1_SDYRB > 0 Then
                SDYRBFHL1_min = FHL1_min_SDYRB / NUM1_SDYRB
            Else
                SDYRBFHL1_min = 0
            End If
            If NUM2_SDYRB > 0 Then
                SDYRBFHL2_min = FHL2_min_SDYRB / NUM2_SDYRB
            Else
                SDYRBFHL2_min = 0
            End If
            '离心式热泵
            '设备最低允许运行的负荷率
            Dim LXSRBFHL1_min As Double
            Dim LXSRBFHL2_min As Double
            If NUM1_LXSRB > 0 Then
                LXSRBFHL1_min = FHL1_min_LXSRB / NUM1_LXSRB
            Else
                LXSRBFHL1_min = 0
            End If
            If NUM2_LXSRB > 0 Then
                LXSRBFHL2_min = FHL2_min_LXSRB / NUM2_LXSRB
            Else
                LXSRBFHL2_min = 0
            End If
            '空气源热泵
            '设备最低允许运行的负荷率
            Dim KQYRBFHL1_min As Double
            Dim KQYRBFHL2_min As Double
            If NUM1_KQYRB > 0 Then
                KQYRBFHL1_min = FHL1_min_KQYRB / NUM1_KQYRB
            Else
                KQYRBFHL1_min = 0
            End If
            If NUM2_KQYRB > 0 Then
                KQYRBFHL2_min = FHL2_min_KQYRB / NUM2_KQYRB
            Else
                KQYRBFHL2_min = 0
            End If
            '直燃型溴化锂
            '设备最低允许运行的负荷率
            Dim ZRXXHLFHL1_min As Double
            Dim ZRXXHLFHL2_min As Double
            If NUM1_ZRXXHL > 0 Then
                ZRXXHLFHL1_min = FHL1_min_ZRXXHL / NUM1_ZRXXHL
            Else
                ZRXXHLFHL1_min = 0
            End If
            If NUM2_ZRXXHL > 0 Then
                ZRXXHLFHL2_min = FHL2_min_ZRXXHL / NUM2_ZRXXHL
            Else
                ZRXXHLFHL2_min = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '如果计算出的设备总负荷率大于0，但是小于设备允许的最低负荷率，则重新修正设备负荷率
            '离心式冷水机（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSLSJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value
            Dim LXSLSJFHL1_XL As Double
            If NUM1_LXSLSJ = 0 Then
                LXSLSJFHL1_XL = 0
            Else
                LXSLSJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value / NUM1_LXSLSJ
            End If
            Dim LXSLSJFHL1 As Double = LXSLSJFHL1_GL + LXSLSJFHL1_XL
            If NUM1_LXSLSJ > 0 And LXSLSJFHL1 < LXSLSJFHL1_min And LXSLSJFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = LXSLSJFHL1_min * (LXSLSJFHL1_GL / LXSLSJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28).Value = LXSLSJFHL1_min * (LXSLSJFHL1_GL / LXSLSJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value = LXSLSJFHL1_min * (LXSLSJFHL1_XL / LXSLSJFHL1) * NUM1_LXSLSJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40).Value = LXSLSJFHL1_min * (LXSLSJFHL1_XL / LXSLSJFHL1) * NUM1_LXSLSJ
            ElseIf NUM1_LXSLSJ > 0 And LXSLSJFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = 1 * (LXSLSJFHL1_GL / LXSLSJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28).Value = 1 * (LXSLSJFHL1_GL / LXSLSJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value = 1 * (LXSLSJFHL1_XL / LXSLSJFHL1) * NUM1_LXSLSJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40).Value = 1 * (LXSLSJFHL1_XL / LXSLSJFHL1) * NUM1_LXSLSJ
            End If
            '离心式冷水机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSLSJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value
            Dim LXSLSJFHL2_XL As Double
            If NUM2_LXSLSJ = 0 Then
                LXSLSJFHL2_XL = 0
            Else
                LXSLSJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value / NUM2_LXSLSJ
            End If
            Dim LXSLSJFHL2 As Double = LXSLSJFHL2_GL + LXSLSJFHL2_XL
            If NUM2_LXSLSJ > 0 And LXSLSJFHL2 < LXSLSJFHL2_min And LXSLSJFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = LXSLSJFHL2_min * (LXSLSJFHL2_GL / LXSLSJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 29).Value = LXSLSJFHL2_min * (LXSLSJFHL2_GL / LXSLSJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value = LXSLSJFHL2_min * (LXSLSJFHL2_XL / LXSLSJFHL2) * NUM2_LXSLSJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 41).Value = LXSLSJFHL2_min * (LXSLSJFHL2_XL / LXSLSJFHL2) * NUM2_LXSLSJ
            ElseIf NUM2_LXSLSJ > 0 And LXSLSJFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = 1 * (LXSLSJFHL2_GL / LXSLSJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 29).Value = 1 * (LXSLSJFHL2_GL / LXSLSJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value = 1 * (LXSLSJFHL2_XL / LXSLSJFHL2) * NUM2_LXSLSJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 41).Value = 1 * (LXSLSJFHL2_XL / LXSLSJFHL2) * NUM2_LXSLSJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '水冷螺杆机（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SLLGJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value
            Dim SLLGJFHL1_XL As Double
            If NUM1_SLLGJ = 0 Then
                SLLGJFHL1_XL = 0
            Else
                SLLGJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value / NUM1_SLLGJ
            End If
            Dim SLLGJFHL1 As Double = SLLGJFHL1_GL + SLLGJFHL1_XL
            If NUM1_SLLGJ > 0 And SLLGJFHL1 < SLLGJFHL1_min And SLLGJFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = SLLGJFHL1_min * (SLLGJFHL1_GL / SLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 36).Value = SLLGJFHL1_min * (SLLGJFHL1_GL / SLLGJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value = SLLGJFHL1_min * (SLLGJFHL1_XL / SLLGJFHL1) * NUM1_SLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 48).Value = SLLGJFHL1_min * (SLLGJFHL1_XL / SLLGJFHL1) * NUM1_SLLGJ
            ElseIf NUM1_SLLGJ > 0 And SLLGJFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = 1 * (SLLGJFHL1_GL / SLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 36).Value = 1 * (SLLGJFHL1_GL / SLLGJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value = 1 * (SLLGJFHL1_XL / SLLGJFHL1) * NUM1_SLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 48).Value = 1 * (SLLGJFHL1_XL / SLLGJFHL1) * NUM1_SLLGJ
            End If
            '水冷螺杆机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SLLGJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value
            Dim SLLGJFHL2_XL As Double
            If NUM2_SLLGJ = 0 Then
                SLLGJFHL2_XL = 0
            Else
                SLLGJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value / NUM2_SLLGJ
            End If
            Dim SLLGJFHL2 As Double = SLLGJFHL2_GL + SLLGJFHL2_XL
            If NUM2_SLLGJ > 0 And SLLGJFHL2 < SLLGJFHL2_min And SLLGJFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = SLLGJFHL2_min * (SLLGJFHL2_GL / SLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 37).Value = SLLGJFHL2_min * (SLLGJFHL2_GL / SLLGJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value = SLLGJFHL2_min * (SLLGJFHL2_XL / SLLGJFHL2) * NUM2_SLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 49).Value = SLLGJFHL2_min * (SLLGJFHL2_XL / SLLGJFHL2) * NUM2_SLLGJ
            ElseIf NUM2_SLLGJ > 0 And SLLGJFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = 1 * (SLLGJFHL2_GL / SLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 37).Value = 1 * (SLLGJFHL2_GL / SLLGJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value = 1 * (SLLGJFHL2_XL / SLLGJFHL2) * NUM2_SLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 49).Value = 1 * (SLLGJFHL2_XL / SLLGJFHL2) * NUM2_SLLGJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '风冷螺杆机（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim FLLGJFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value
            Dim FLLGJFHL1_XL As Double
            If NUM1_FLLGJ = 0 Then
                FLLGJFHL1_XL = 0
            Else
                FLLGJFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value / NUM1_FLLGJ
            End If
            Dim FLLGJFHL1 As Double = FLLGJFHL1_GL + FLLGJFHL1_XL
            If NUM1_FLLGJ > 0 And FLLGJFHL1 < FLLGJFHL1_min And FLLGJFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = FLLGJFHL1_min * (FLLGJFHL1_GL / FLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 30).Value = FLLGJFHL1_min * (FLLGJFHL1_GL / FLLGJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value = FLLGJFHL1_min * (FLLGJFHL1_XL / FLLGJFHL1) * NUM1_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 42).Value = FLLGJFHL1_min * (FLLGJFHL1_XL / FLLGJFHL1) * NUM1_FLLGJ
            ElseIf NUM1_FLLGJ > 0 And FLLGJFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = 1 * (FLLGJFHL1_GL / FLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 30).Value = 1 * (FLLGJFHL1_GL / FLLGJFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value = 1 * (FLLGJFHL1_XL / FLLGJFHL1) * NUM1_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 42).Value = 1 * (FLLGJFHL1_XL / FLLGJFHL1) * NUM1_FLLGJ
            End If
            '风冷螺杆机（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim FLLGJFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value
            Dim FLLGJFHL2_XL As Double
            If NUM2_FLLGJ = 0 Then
                FLLGJFHL2_XL = 0
            Else
                FLLGJFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value / NUM2_FLLGJ
            End If
            Dim FLLGJFHL2 As Double = FLLGJFHL2_GL + FLLGJFHL2_XL
            If NUM2_FLLGJ > 0 And FLLGJFHL2 < FLLGJFHL2_min And FLLGJFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = FLLGJFHL2_min * (FLLGJFHL2_GL / FLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 31).Value = FLLGJFHL2_min * (FLLGJFHL2_GL / FLLGJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value = FLLGJFHL2_min * (FLLGJFHL2_XL / FLLGJFHL2) * NUM2_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 43).Value = FLLGJFHL2_min * (FLLGJFHL2_XL / FLLGJFHL2) * NUM2_FLLGJ
            ElseIf NUM2_FLLGJ > 0 And FLLGJFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = 1 * (FLLGJFHL2_GL / FLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 31).Value = 1 * (FLLGJFHL2_GL / FLLGJFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value = 1 * (FLLGJFHL2_XL / FLLGJFHL2) * NUM2_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 43).Value = 1 * (FLLGJFHL2_XL / FLLGJFHL2) * NUM2_FLLGJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '水（地）源热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SDYRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value
            Dim SDYRBFHL1_XL As Double
            If NUM1_SDYRB = 0 Then
                SDYRBFHL1_XL = 0
            Else
                SDYRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value / NUM1_SDYRB
            End If
            Dim SDYRBFHL1 As Double = SDYRBFHL1_GL + SDYRBFHL1_XL
            If NUM1_SDYRB > 0 And SDYRBFHL1 < SDYRBFHL1_min And SDYRBFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = SDYRBFHL1_min * (SDYRBFHL1_GL / SDYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 34).Value = SDYRBFHL1_min * (SDYRBFHL1_GL / SDYRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value = SDYRBFHL1_min * (SDYRBFHL1_XL / SDYRBFHL1) * NUM1_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 46).Value = SDYRBFHL1_min * (SDYRBFHL1_XL / SDYRBFHL1) * NUM1_SDYRB
            ElseIf NUM1_SDYRB > 0 And SDYRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = 1 * (SDYRBFHL1_GL / SDYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 34).Value = 1 * (SDYRBFHL1_GL / SDYRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value = 1 * (SDYRBFHL1_XL / SDYRBFHL1) * NUM1_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 46).Value = 1 * (SDYRBFHL1_XL / SDYRBFHL1) * NUM1_SDYRB
            End If
            '水（地）源热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim SDYRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value
            Dim SDYRBFHL2_XL As Double
            If NUM2_SDYRB = 0 Then
                SDYRBFHL2_XL = 0
            Else
                SDYRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value / NUM2_SDYRB
            End If
            Dim SDYRBFHL2 As Double = SDYRBFHL2_GL + SDYRBFHL2_XL
            If NUM2_SDYRB > 0 And SDYRBFHL2 < SDYRBFHL2_min And SDYRBFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = SDYRBFHL2_min * (SDYRBFHL2_GL / SDYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 35).Value = SDYRBFHL2_min * (SDYRBFHL2_GL / SDYRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value = SDYRBFHL2_min * (SDYRBFHL2_XL / SDYRBFHL2) * NUM2_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 47).Value = SDYRBFHL2_min * (SDYRBFHL2_XL / SDYRBFHL2) * NUM2_SDYRB
            ElseIf NUM2_SDYRB > 0 And SDYRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = 1 * (SDYRBFHL2_GL / SDYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 35).Value = 1 * (SDYRBFHL2_GL / SDYRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value = 1 * (SDYRBFHL2_XL / SDYRBFHL2) * NUM2_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 47).Value = 1 * (SDYRBFHL2_XL / SDYRBFHL2) * NUM2_SDYRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '离心式热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value
            Dim LXSRBFHL1_XL As Double
            If NUM1_LXSRB = 0 Then
                LXSRBFHL1_XL = 0
            Else
                LXSRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value / NUM1_LXSRB
            End If
            Dim LXSRBFHL1 As Double = LXSRBFHL1_GL + LXSRBFHL1_XL
            If NUM1_LXSRB > 0 And LXSRBFHL1 < LXSRBFHL1_min And LXSRBFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = LXSRBFHL1_min * (LXSRBFHL1_GL / LXSRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 38).Value = LXSRBFHL1_min * (LXSRBFHL1_GL / LXSRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value = LXSRBFHL1_min * (LXSRBFHL1_XL / LXSRBFHL1) * NUM1_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 50).Value = LXSRBFHL1_min * (LXSRBFHL1_XL / LXSRBFHL1) * NUM1_LXSRB
            ElseIf NUM1_LXSRB > 0 And LXSRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = 1 * (LXSRBFHL1_GL / LXSRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 38).Value = 1 * (LXSRBFHL1_GL / LXSRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value = 1 * (LXSRBFHL1_XL / LXSRBFHL1) * NUM1_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 50).Value = 1 * (LXSRBFHL1_XL / LXSRBFHL1) * NUM1_LXSRB
            End If
            '离心式热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim LXSRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value
            Dim LXSRBFHL2_XL As Double
            If NUM2_LXSRB = 0 Then
                LXSRBFHL2_XL = 0
            Else
                LXSRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value / NUM2_LXSRB
            End If
            Dim LXSRBFHL2 As Double = LXSRBFHL2_GL + LXSRBFHL2_XL
            If NUM2_LXSRB > 0 And LXSRBFHL2 < LXSRBFHL2_min And LXSRBFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = LXSRBFHL2_min * (LXSRBFHL2_GL / LXSRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 39).Value = LXSRBFHL2_min * (LXSRBFHL2_GL / LXSRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value = LXSRBFHL2_min * (LXSRBFHL2_XL / LXSRBFHL2) * NUM2_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51).Value = LXSRBFHL2_min * (LXSRBFHL2_XL / LXSRBFHL2) * NUM2_LXSRB
            ElseIf NUM2_LXSRB > 0 And LXSRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = 1 * (LXSRBFHL2_GL / LXSRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 39).Value = 1 * (LXSRBFHL2_GL / LXSRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value = 1 * (LXSRBFHL2_XL / LXSRBFHL2) * NUM2_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51).Value = 1 * (LXSRBFHL2_XL / LXSRBFHL2) * NUM2_LXSRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '空气源热泵（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim KQYRBFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value
            Dim KQYRBFHL1_XL As Double
            If NUM1_KQYRB = 0 Then
                KQYRBFHL1_XL = 0
            Else
                KQYRBFHL1_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value / NUM1_KQYRB
            End If
            Dim KQYRBFHL1 As Double = KQYRBFHL1_GL + KQYRBFHL1_XL
            If NUM1_KQYRB > 0 And KQYRBFHL1 < KQYRBFHL1_min And KQYRBFHL1 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = KQYRBFHL1_min * (KQYRBFHL1_GL / KQYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 32).Value = KQYRBFHL1_min * (KQYRBFHL1_GL / KQYRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value = KQYRBFHL1_min * (KQYRBFHL1_XL / KQYRBFHL1) * NUM1_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 44).Value = KQYRBFHL1_min * (KQYRBFHL1_XL / KQYRBFHL1) * NUM1_KQYRB
            ElseIf NUM1_KQYRB > 0 And KQYRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = 1 * (KQYRBFHL1_GL / KQYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 32).Value = 1 * (KQYRBFHL1_GL / KQYRBFHL1)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value = 1 * (KQYRBFHL1_XL / KQYRBFHL1) * NUM1_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 44).Value = 1 * (KQYRBFHL1_XL / KQYRBFHL1) * NUM1_KQYRB
            End If
            '空气源热泵（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim KQYRBFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value
            Dim KQYRBFHL2_XL As Double
            If NUM2_KQYRB = 0 Then
                KQYRBFHL2_XL = 0
            Else
                KQYRBFHL2_XL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value / NUM2_KQYRB
            End If
            Dim KQYRBFHL2 As Double = KQYRBFHL2_GL + KQYRBFHL2_XL
            If NUM2_KQYRB > 0 And KQYRBFHL2 < KQYRBFHL2_min And KQYRBFHL2 > 0 Then
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = KQYRBFHL2_min * (KQYRBFHL2_GL / KQYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 33).Value = KQYRBFHL2_min * (KQYRBFHL2_GL / KQYRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value = KQYRBFHL2_min * (KQYRBFHL2_XL / KQYRBFHL2) * NUM2_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 45).Value = KQYRBFHL2_min * (KQYRBFHL2_XL / KQYRBFHL2) * NUM2_KQYRB
            ElseIf NUM2_KQYRB > 0 And KQYRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供冷和蓄冷负荷率按比例重新分配
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = 1 * (KQYRBFHL2_GL / KQYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 33).Value = 1 * (KQYRBFHL2_GL / KQYRBFHL2)
                '蓄冷（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value = 1 * (KQYRBFHL2_XL / KQYRBFHL2) * NUM2_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 45).Value = 1 * (KQYRBFHL2_XL / KQYRBFHL2) * NUM2_KQYRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '直燃型溴化锂（1）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL1_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value
            Dim ZRXXHLFHL1 As Double = ZRXXHLFHL1_GL
            If NUM1_ZRXXHL > 0 And ZRXXHLFHL1 < ZRXXHLFHL1_min And ZRXXHLFHL1 > 0 Then
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = ZRXXHLFHL1_min * (ZRXXHLFHL1_GL / ZRXXHLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90).Value = ZRXXHLFHL1_min * (ZRXXHLFHL1_GL / ZRXXHLFHL1)
            ElseIf NUM1_ZRXXHL > 0 And ZRXXHLFHL1 > 1.2 Then
                '如果算出来的总负荷率大于1.2（直燃型溴化锂制冷可以超发到1.2）
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = 1.2 * (ZRXXHLFHL1_GL / ZRXXHLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90).Value = 1.2 * (ZRXXHLFHL1_GL / ZRXXHLFHL1)
            End If
            '直燃型溴化锂（2）
            '此时的设备负荷率=供冷负荷率+蓄冷负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL2_GL As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value
            Dim ZRXXHLFHL2 As Double = ZRXXHLFHL2_GL
            If NUM2_ZRXXHL > 0 And ZRXXHLFHL2 < ZRXXHLFHL2_min And ZRXXHLFHL2 > 0 Then
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = ZRXXHLFHL2_min * (ZRXXHLFHL2_GL / ZRXXHLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 91).Value = ZRXXHLFHL2_min * (ZRXXHLFHL2_GL / ZRXXHLFHL2)
            ElseIf NUM2_ZRXXHL > 0 And ZRXXHLFHL2 > 1.2 Then
                '如果算出来的总负荷率大于1.2（直燃型溴化锂制冷可以超发到1.2）
                '供冷
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = 1.2 * (ZRXXHLFHL2_GL / ZRXXHLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 91).Value = 1.2 * (ZRXXHLFHL2_GL / ZRXXHLFHL2)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算出的负荷率结果均保留3位小数，增加美观度
            '列号
            For i = 28 To 51
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
            '直燃型溴化锂
            '列号
            For i = 90 To 91
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
        End If
    End Sub
    Sub 制热和蓄热空调设备负荷率修正(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '常规计算模式就不进行修正了
        If calculation_mode = 2 Then
            '读取采暖季装机方案及参数
            Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
            '内燃机
            Dim NUM1_NRJ As Double = ans_ZJFA_R(0)
            Dim NUM2_NRJ As Double = ans_ZJFA_R(1)
            '天然气锅炉
            Dim NUM1_TRQGL As Double = ans_ZJFA_R(10)
            Dim NUM2_TRQGL As Double = ans_ZJFA_R(11)
            '电锅炉
            Dim NUM1_DGL As Double = ans_ZJFA_R(18)
            Dim NUM2_DGL As Double = ans_ZJFA_R(19)
            Dim ZJRGL1_DGL As Double = ans_ZJFA_R(20)
            Dim ZJRGL2_DGL As Double = ans_ZJFA_R(21)
            '风冷螺杆机
            Dim NUM1_FLLGJ As Double = ans_ZJFA_R(26)
            Dim NUM2_FLLGJ As Double = ans_ZJFA_R(27)
            '水地源热泵
            Dim NUM1_SDYRB As Double = ans_ZJFA_R(34)
            Dim NUM2_SDYRB As Double = ans_ZJFA_R(35)
            '离心式热泵
            Dim NUM1_LXSRB As Double = ans_ZJFA_R(42)
            Dim NUM2_LXSRB As Double = ans_ZJFA_R(43)
            '空气源热泵
            Dim NUM1_KQYRB As Double = ans_ZJFA_R(50)
            Dim NUM2_KQYRB As Double = ans_ZJFA_R(51)
            '直燃型溴化锂
            Dim NUM1_ZRXXHL As Double = ans_ZJFA_R(58)
            Dim NUM2_ZRXXHL As Double = ans_ZJFA_R(59)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '参与制热和蓄热的空调设备，根据设置的设备允许最低负荷率，对计算出的负荷率进行修正
            '进行修正的负荷率=供热负荷率+蓄热负荷率
            '各设备允许运行的最低负荷率
            '设备最低运行负荷率
            Dim TRQGLFHL1_min As Double
            Dim TRQGLFHL2_min As Double
            If NUM1_TRQGL > 0 Then
                TRQGLFHL1_min = FHL1_min_TRQGL / NUM1_TRQGL
            Else
                TRQGLFHL1_min = 0
            End If
            If NUM2_TRQGL > 0 Then
                TRQGLFHL2_min = FHL2_min_TRQGL / NUM2_TRQGL
            Else
                TRQGLFHL2_min = 0
            End If
            '风冷螺杆机
            '设备最低运行负荷率
            Dim FLLGJFHL1_min As Double
            Dim FLLGJFHL2_min As Double
            If NUM1_FLLGJ > 0 Then
                FLLGJFHL1_min = FHL1_min_FLLGJ / NUM1_FLLGJ
            Else
                FLLGJFHL1_min = 0
            End If
            If NUM2_FLLGJ > 0 Then
                FLLGJFHL2_min = FHL2_min_FLLGJ / NUM2_FLLGJ
            Else
                FLLGJFHL2_min = 0
            End If
            '水（地）源热泵
            '设备最低运行负荷率
            Dim SDYRBFHL1_min As Double
            Dim SDYRBFHL2_min As Double
            If NUM1_SDYRB > 0 Then
                SDYRBFHL1_min = FHL1_min_SDYRB / NUM1_SDYRB
            Else
                SDYRBFHL1_min = 0
            End If
            If NUM2_SDYRB > 0 Then
                SDYRBFHL2_min = FHL2_min_SDYRB / NUM2_SDYRB
            Else
                SDYRBFHL2_min = 0
            End If
            '空气源热泵
            '设备最低运行负荷率
            Dim KQYRBFHL1_min As Double
            Dim KQYRBFHL2_min As Double
            If NUM1_KQYRB > 0 Then
                KQYRBFHL1_min = FHL1_min_KQYRB / NUM1_KQYRB
            Else
                KQYRBFHL1_min = 0
            End If
            If NUM2_KQYRB > 0 Then
                KQYRBFHL2_min = FHL2_min_KQYRB / NUM2_KQYRB
            Else
                KQYRBFHL2_min = 0
            End If
            '离心式热泵
            '设备最低运行负荷率
            Dim LXSRBFHL1_min As Double
            Dim LXSRBFHL2_min As Double
            If NUM1_LXSRB > 0 Then
                LXSRBFHL1_min = FHL1_min_LXSRB / NUM1_LXSRB
            Else
                LXSRBFHL1_min = 0
            End If
            If NUM2_LXSRB > 0 Then
                LXSRBFHL2_min = FHL2_min_LXSRB / NUM2_LXSRB
            Else
                LXSRBFHL2_min = 0
            End If
            '直燃型溴化锂
            '设备最低运行负荷率
            Dim ZRXXHLFHL1_min As Double
            Dim ZRXXHLFHL2_min As Double
            If NUM1_ZRXXHL > 0 Then
                ZRXXHLFHL1_min = FHL1_min_ZRXXHL / NUM1_ZRXXHL
            Else
                ZRXXHLFHL1_min = 0
            End If
            If NUM2_ZRXXHL > 0 Then
                ZRXXHLFHL2_min = FHL2_min_ZRXXHL / NUM2_ZRXXHL
            Else
                ZRXXHLFHL2_min = 0
            End If
            '电采暖锅炉
            Dim DGLFHL1_min As Double = FHL1_min_DGL
            Dim DGLFHL2_min As Double = FHL2_min_DGL
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '供热的综合负荷率计算要考虑混水供热设备和梯级供热设备
            '梯级供热设备
            Dim FHL_TJGR_FLLGJ As Double = 0
            Dim FHL_TJGR_KQYRB As Double = 0
            Dim FHL_TJGR_SDYRB As Double = 0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "风冷螺杆机" Then
                FHL_TJGR_FLLGJ = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "空气源热泵" Then
                FHL_TJGR_KQYRB = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "水(地)源热泵" Then
                FHL_TJGR_SDYRB = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value
            End If
            '3种设备不可能同时存在
            If FHL_TJGR_FLLGJ + FHL_TJGR_KQYRB + FHL_TJGR_SDYRB < 0.1 And FHL_TJGR_FLLGJ + FHL_TJGR_KQYRB + FHL_TJGR_SDYRB > 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = 0.1
            ElseIf FHL_TJGR_FLLGJ + FHL_TJGR_KQYRB + FHL_TJGR_SDYRB > 1 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = 1
            End If
            '混水供热
            Dim FHL_HSGR_FLLGJ As Double = 0
            Dim FHL_HSGR_KQYRB As Double = 0
            Dim FHL_HSGR_SDYRB As Double = 0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                FHL_HSGR_FLLGJ = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                FHL_HSGR_KQYRB = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                FHL_HSGR_SDYRB = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value
            End If
            '3种设备不可能同时存在
            If FHL_HSGR_FLLGJ + FHL_HSGR_KQYRB + FHL_HSGR_SDYRB < 0.1 And FHL_HSGR_FLLGJ + FHL_HSGR_KQYRB + FHL_HSGR_SDYRB <> 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = 0.1
            ElseIf FHL_HSGR_FLLGJ + FHL_HSGR_KQYRB + FHL_HSGR_SDYRB > 1 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————  
            '如果计算出的设备总负荷率大于0，但是小于设备允许的最低负荷率，则重新修正设备负荷率
            '天然气锅炉（1）
            '此时的设备负荷率=供热负荷率
            Dim TRQGLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value
            Dim TRQGLFHL1 As Double = TRQGLFHL1_GR
            If NUM1_TRQGL > 0 And TRQGLFHL1 < TRQGLFHL1_min And TRQGLFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = TRQGLFHL1_min * (TRQGLFHL1_GR / TRQGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = TRQGLFHL1_min * (TRQGLFHL1_GR / TRQGLFHL1)
            ElseIf NUM1_TRQGL > 0 And TRQGLFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = 1 * (TRQGLFHL1_GR / TRQGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = 1 * (TRQGLFHL1_GR / TRQGLFHL1)
            End If
            '天然气锅炉（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim TRQGLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value
            Dim TRQGLFHL2 As Double = TRQGLFHL2_GR
            If NUM2_TRQGL > 0 And TRQGLFHL2 < TRQGLFHL2_min And TRQGLFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = TRQGLFHL2_min * (TRQGLFHL2_GR / TRQGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = TRQGLFHL2_min * (TRQGLFHL2_GR / TRQGLFHL2)
            ElseIf NUM2_TRQGL > 0 And TRQGLFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = 1 * (TRQGLFHL2_GR / TRQGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = 1 * (TRQGLFHL2_GR / TRQGLFHL2)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '电锅炉（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率
            Dim DGLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value
            Dim DGLFHL1_XR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value
            Dim DGLFHL1 As Double = DGLFHL1_GR + DGLFHL1_XR
            If ZJRGL1_DGL > 0 And DGLFHL1 < DGLFHL1_min And DGLFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = DGLFHL1_min * (DGLFHL1_GR / DGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = DGLFHL1_min * (DGLFHL1_GR / DGLFHL1)
                '蓄热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = DGLFHL1_min * (DGLFHL1_XR / DGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = DGLFHL1_min * (DGLFHL1_XR / DGLFHL1)
            ElseIf ZJRGL1_DGL > 0 And DGLFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = 1 * (DGLFHL1_GR / DGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = 1 * (DGLFHL1_GR / DGLFHL1)
                '蓄热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = 1 * (DGLFHL1_XR / DGLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = 1 * (DGLFHL1_XR / DGLFHL1)
            End If
            '电锅炉（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim DGLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value
            Dim DGLFHL2_XR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value
            Dim DGLFHL2 As Double = DGLFHL2_GR + DGLFHL2_XR
            If ZJRGL2_DGL > 0 And DGLFHL2 < DGLFHL2_min And DGLFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = DGLFHL2_min * (DGLFHL2_GR / DGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = DGLFHL2_min * (DGLFHL2_GR / DGLFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = DGLFHL2_min * (DGLFHL2_XR / DGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = DGLFHL2_min * (DGLFHL2_XR / DGLFHL2)
            ElseIf ZJRGL2_DGL > 0 And DGLFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = 1 * (DGLFHL2_GR / DGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = 1 * (DGLFHL2_GR / DGLFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = 1 * (DGLFHL2_XR / DGLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = 1 * (DGLFHL2_XR / DGLFHL2)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————              
            '风冷螺杆机（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim FLLGJFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value
            Dim FLLGJFHL1_XR As Double
            If NUM1_FLLGJ = 0 Then
                FLLGJFHL1_XR = 0
            Else
                FLLGJFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value / NUM1_FLLGJ
            End If
            Dim FLLGJFHL1 As Double = FLLGJFHL1_GR + FLLGJFHL1_XR
            If NUM1_FLLGJ > 0 And FLLGJFHL1 < FLLGJFHL1_min And FLLGJFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = FLLGJFHL1_min * (FLLGJFHL1_GR / FLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 54).Value = FLLGJFHL1_min * (FLLGJFHL1_GR / FLLGJFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value = FLLGJFHL1_min * (FLLGJFHL1_XR / FLLGJFHL1) * NUM1_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62).Value = FLLGJFHL1_min * (FLLGJFHL1_XR / FLLGJFHL1) * NUM1_FLLGJ
            ElseIf NUM1_FLLGJ > 0 And FLLGJFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = 1 * (FLLGJFHL1_GR / FLLGJFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 54).Value = 1 * (FLLGJFHL1_GR / FLLGJFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value = 1 * (FLLGJFHL1_XR / FLLGJFHL1) * NUM1_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62).Value = 1 * (FLLGJFHL1_XR / FLLGJFHL1) * NUM1_FLLGJ
            End If
            '风冷螺杆机（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim FLLGJFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value
            Dim FLLGJFHL2_XR As Double
            If NUM2_FLLGJ = 0 Then
                FLLGJFHL2_XR = 0
            Else
                FLLGJFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value / NUM2_FLLGJ
            End If
            Dim FLLGJFHL2 As Double = FLLGJFHL2_GR + FLLGJFHL2_XR
            If NUM2_FLLGJ > 0 And FLLGJFHL2 < FLLGJFHL2_min And FLLGJFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = FLLGJFHL2_min * (FLLGJFHL2_GR / FLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 55).Value = FLLGJFHL2_min * (FLLGJFHL2_GR / FLLGJFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value = FLLGJFHL2_min * (FLLGJFHL2_XR / FLLGJFHL2) * NUM2_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 63).Value = FLLGJFHL2_min * (FLLGJFHL2_XR / FLLGJFHL2) * NUM2_FLLGJ
            ElseIf NUM2_FLLGJ > 0 And FLLGJFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = 1 * (FLLGJFHL2_GR / FLLGJFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 55).Value = 1 * (FLLGJFHL2_GR / FLLGJFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value = 1 * (FLLGJFHL2_XR / FLLGJFHL2) * NUM2_FLLGJ
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 63).Value = 1 * (FLLGJFHL2_XR / FLLGJFHL2) * NUM2_FLLGJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————                   
            '水（地）源热泵（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim SDYRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value
            Dim SDYRBFHL1_XR As Double
            If NUM1_SDYRB = 0 Then
                SDYRBFHL1_XR = 0
            Else
                SDYRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value / NUM1_SDYRB
            End If
            Dim SDYRBFHL1 As Double = SDYRBFHL1_GR + SDYRBFHL1_XR
            If NUM1_SDYRB > 0 And SDYRBFHL1 < SDYRBFHL1_min And SDYRBFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = SDYRBFHL1_min * (SDYRBFHL1_GR / SDYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 58).Value = SDYRBFHL1_min * (SDYRBFHL1_GR / SDYRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value = SDYRBFHL1_min * (SDYRBFHL1_XR / SDYRBFHL1) * NUM1_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 66).Value = SDYRBFHL1_min * (SDYRBFHL1_XR / SDYRBFHL1) * NUM1_SDYRB
            ElseIf NUM1_SDYRB > 0 And SDYRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = 1 * (SDYRBFHL1_GR / SDYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 58).Value = 1 * (SDYRBFHL1_GR / SDYRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value = 1 * (SDYRBFHL1_XR / SDYRBFHL1) * NUM1_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 66).Value = 1 * (SDYRBFHL1_XR / SDYRBFHL1) * NUM1_SDYRB
            End If
            '水（地）源热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim SDYRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value
            Dim SDYRBFHL2_XR As Double
            If NUM2_SDYRB = 0 Then
                SDYRBFHL2_XR = 0
            Else
                SDYRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value / NUM2_SDYRB
            End If
            Dim SDYRBFHL2 As Double = SDYRBFHL2_GR + SDYRBFHL2_XR
            If NUM2_SDYRB > 0 And SDYRBFHL2 < SDYRBFHL2_min And SDYRBFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = SDYRBFHL2_min * (SDYRBFHL2_GR / SDYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 59).Value = SDYRBFHL2_min * (SDYRBFHL2_GR / SDYRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value = SDYRBFHL2_min * (SDYRBFHL2_XR / SDYRBFHL2) * NUM2_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 67).Value = SDYRBFHL2_min * (SDYRBFHL2_XR / SDYRBFHL2) * NUM2_SDYRB
            ElseIf NUM2_SDYRB > 0 And SDYRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = 1 * (SDYRBFHL2_GR / SDYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 59).Value = 1 * (SDYRBFHL2_GR / SDYRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value = 1 * (SDYRBFHL2_XR / SDYRBFHL2) * NUM2_SDYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 67).Value = 1 * (SDYRBFHL2_XR / SDYRBFHL2) * NUM2_SDYRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '离心式热泵（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim LXSRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value
            Dim LXSRBFHL1_XR As Double
            If NUM1_LXSRB = 0 Then
                LXSRBFHL1_XR = 0
            Else
                LXSRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value / NUM1_LXSRB
            End If
            Dim LXSRBFHL1 As Double = LXSRBFHL1_GR + LXSRBFHL1_XR
            If NUM1_LXSRB > 0 And LXSRBFHL1 < LXSRBFHL1_min And LXSRBFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = LXSRBFHL1_min * (LXSRBFHL1_GR / LXSRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value = LXSRBFHL1_min * (LXSRBFHL1_GR / LXSRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value = LXSRBFHL1_min * (LXSRBFHL1_XR / LXSRBFHL1) * NUM1_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value = LXSRBFHL1_min * (LXSRBFHL1_XR / LXSRBFHL1) * NUM1_LXSRB
            ElseIf NUM1_LXSRB > 0 And LXSRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = 1 * (LXSRBFHL1_GR / LXSRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value = 1 * (LXSRBFHL1_GR / LXSRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value = 1 * (LXSRBFHL1_XR / LXSRBFHL1) * NUM1_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value = 1 * (LXSRBFHL1_XR / LXSRBFHL1) * NUM1_LXSRB
            End If
            '离心式热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim LXSRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value
            Dim LXSRBFHL2_XR As Double
            If NUM2_LXSRB = 0 Then
                LXSRBFHL2_XR = 0
            Else
                LXSRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value / NUM2_LXSRB
            End If
            Dim LXSRBFHL2 As Double = LXSRBFHL2_GR + LXSRBFHL2_XR
            If NUM2_LXSRB > 0 And LXSRBFHL2 < LXSRBFHL2_min And LXSRBFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = LXSRBFHL2_min * (LXSRBFHL2_GR / LXSRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value = LXSRBFHL2_min * (LXSRBFHL2_GR / LXSRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value = LXSRBFHL2_min * (LXSRBFHL2_XR / LXSRBFHL2) * NUM2_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value = LXSRBFHL2_min * (LXSRBFHL2_XR / LXSRBFHL2) * NUM2_LXSRB
            ElseIf NUM2_LXSRB > 0 And LXSRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = 1 * (LXSRBFHL2_GR / LXSRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value = 1 * (LXSRBFHL2_GR / LXSRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value = 1 * (LXSRBFHL2_XR / LXSRBFHL2) * NUM2_LXSRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value = 1 * (LXSRBFHL2_XR / LXSRBFHL2) * NUM2_LXSRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————                    
            '空气源热泵（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim KQYRBFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value
            Dim KQYRBFHL1_XR As Double
            If NUM1_KQYRB = 0 Then
                KQYRBFHL1_XR = 0
            Else
                KQYRBFHL1_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value / NUM1_KQYRB
            End If
            Dim KQYRBFHL1 As Double = KQYRBFHL1_GR + KQYRBFHL1_XR
            If NUM1_KQYRB > 0 And KQYRBFHL1 < KQYRBFHL1_min And KQYRBFHL1 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = KQYRBFHL1_min * (KQYRBFHL1_GR / KQYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 56).Value = KQYRBFHL1_min * (KQYRBFHL1_GR / KQYRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value = KQYRBFHL1_min * (KQYRBFHL1_XR / KQYRBFHL1) * NUM1_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 64).Value = KQYRBFHL1_min * (KQYRBFHL1_XR / KQYRBFHL1) * NUM1_KQYRB
            ElseIf NUM1_KQYRB > 0 And KQYRBFHL1 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = 1 * (KQYRBFHL1_GR / KQYRBFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 56).Value = 1 * (KQYRBFHL1_GR / KQYRBFHL1)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value = 1 * (KQYRBFHL1_XR / KQYRBFHL1) * NUM1_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 64).Value = 1 * (KQYRBFHL1_XR / KQYRBFHL1) * NUM1_KQYRB
            End If
            '空气源热泵（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim KQYRBFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value
            Dim KQYRBFHL2_XR As Double
            If NUM2_KQYRB = 0 Then
                KQYRBFHL2_XR = 0
            Else
                KQYRBFHL2_XR = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value / NUM2_KQYRB
            End If
            Dim KQYRBFHL2 As Double = KQYRBFHL2_GR + KQYRBFHL2_XR
            If NUM2_KQYRB > 0 And KQYRBFHL2 < KQYRBFHL2_min And KQYRBFHL2 > 0 Then
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = KQYRBFHL2_min * (KQYRBFHL2_GR / KQYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 57).Value = KQYRBFHL2_min * (KQYRBFHL2_GR / KQYRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value = KQYRBFHL2_min * (KQYRBFHL2_XR / KQYRBFHL2) * NUM2_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 65).Value = KQYRBFHL2_min * (KQYRBFHL2_XR / KQYRBFHL2) * NUM2_KQYRB
            ElseIf NUM2_KQYRB > 0 And KQYRBFHL2 > 1 Then
                '如果算出来的总负荷率大于1
                '根据之前计算出的供热和蓄热负荷率按比例重新分配
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = 1 * (KQYRBFHL2_GR / KQYRBFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 57).Value = 1 * (KQYRBFHL2_GR / KQYRBFHL2)
                '蓄热（需要转换成台数）
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value = 1 * (KQYRBFHL2_XR / KQYRBFHL2) * NUM2_KQYRB
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 65).Value = 1 * (KQYRBFHL2_XR / KQYRBFHL2) * NUM2_KQYRB
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '直燃型溴化锂（1）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL1_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value
            Dim ZRXXHLFHL1 As Double = ZRXXHLFHL1_GR
            If NUM1_ZRXXHL > 0 And ZRXXHLFHL1 < ZRXXHLFHL1_min And ZRXXHLFHL1 > 0 Then
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = ZRXXHLFHL1_min * (ZRXXHLFHL1_GR / ZRXXHLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = ZRXXHLFHL1_min * (ZRXXHLFHL1_GR / ZRXXHLFHL1)
            ElseIf NUM1_ZRXXHL > 0 And ZRXXHLFHL1 > 1.2 Then
                '如果算出来的总负荷率大于1
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = 1 * (ZRXXHLFHL1_GR / ZRXXHLFHL1)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = 1 * (ZRXXHLFHL1_GR / ZRXXHLFHL1)
            End If
            '直燃型溴化锂（2）
            '此时的设备负荷率=供热负荷率+蓄热负荷率（除以装机数量换算一下）
            Dim ZRXXHLFHL2_GR As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value
            Dim ZRXXHLFHL2 As Double = ZRXXHLFHL2_GR
            If NUM2_ZRXXHL > 0 And ZRXXHLFHL2 < ZRXXHLFHL2_min And ZRXXHLFHL2 > 0 Then
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = ZRXXHLFHL2_min * (ZRXXHLFHL2_GR / ZRXXHLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = ZRXXHLFHL2_min * (ZRXXHLFHL2_GR / ZRXXHLFHL2)
            ElseIf NUM2_ZRXXHL > 0 And ZRXXHLFHL2 > 1.2 Then
                '如果算出来的总负荷率大于1
                '供热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = 1 * (ZRXXHLFHL2_GR / ZRXXHLFHL2)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = 1 * (ZRXXHLFHL2_GR / ZRXXHLFHL2)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算出的负荷率结果均保留3位小数，增加美观度
            '列号
            For i = 52 To 70
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
            '电锅炉和直燃型溴化锂
            '列号
            For i = 92 To 97
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, i).Value, 3)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value, 3)
            Next
        End If
    End Sub
End Module
