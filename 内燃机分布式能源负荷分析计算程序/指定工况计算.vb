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
        Dim GKXH(6) As Integer
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
            '判断是否需要混水供热或者梯级供热，并输入相关计算比例
            Dim TJGRZTJC = 0 '梯级供热状态监测
            Dim HSGRZTJC = 0 '混水供热状态监测
            For i = 1 To 5
                b = GKXH(i)
                '忽略为0的工况
                If b > 0 Then
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing Then '如果有仅计算耗电量，不计算供热量的设备
                        TJGRZTJC = TJGRZTJC + 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing Then '如果有混水供热的设备
                        HSGRZTJC = HSGRZTJC + 1
                    End If
                End If
            Next
            '输入体积供热比例
            Dim TJGRFHBL As Double = 0
            If TJGRZTJC > 0 Then
                '让用户手动输入负荷率比例（一级热泵负荷率/二级热泵负荷率），一级热泵跟随二级热泵负荷率的变化而变化，在二级热泵负荷率基础上乘以系数，即为一级热泵负荷率
                TJGRFHBL = InputBox("请输入梯级供热一级热泵（仅计算耗电量，不计算供热量设备）与二级热泵负荷比例系数，（一级热泵负荷比例÷二级热泵负荷比例）", "请输入梯级供热设备负荷比例系数", 1)
            End If
            '存在混水计算，输入混水计算比例
            Dim HSGRGLBL As Double = 0
            If HSGRZTJC > 0 Then
                '让用户手动输入混水设备制热功率，占总热功率（总热功率指的是，几种混水向外供热的设备总供热功率的合计）的比例
                HSGRGLBL = InputBox("请输入混水供热功率比例系数（选择的混水供热设备的供热功率，占两种设备总供热功率的比例，例如(风冷热泵的制热功率/（天然气锅炉的制热功率+风冷热泵的制热功率）)）", "请输入混水供热功率比例系数", 0.5)
                '读取天然气锅炉、直燃型溴化锂机组、电采暖锅炉制热总功率（装机量，制热出力最大值）
                '读取采暖季装机方案及参数
                Dim ans_ZJFA_R = mainprogram.读取采暖季装机方案参数(ExcelApp)
                '天然气锅炉装机功率（总和）
                Dim TRQGL1ZRGL As Double = ans_ZJFA_R(12)
                Dim TRQGL2ZRGL As Double = ans_ZJFA_R(13)
                '电锅炉装机功率（总和）
                Dim DCNGL1ZRGL As Double = ans_ZJFA_R(20)
                Dim DCNGL2ZRGL As Double = ans_ZJFA_R(21)
                '风冷螺杆机（总和）
                Dim ZJRGL1_FLLGJ As Double = ans_ZJFA_R(28)
                Dim ZJRGL2_FLLGJ As Double = ans_ZJFA_R(29)
                '水地源热泵（总和）
                Dim ZJRGL1_SDYRB As Double = ans_ZJFA_R(36)
                Dim ZJRGL2_SDYRB As Double = ans_ZJFA_R(37)
                '空气源热泵（总和）
                Dim ZJRGL1_KQYRB As Double = ans_ZJFA_R(52)
                Dim ZJRGL2_KQYRB As Double = ans_ZJFA_R(53)
                '直燃型溴化锂装机功率（总和）
                Dim ZRXHL1ZRGL As Double = ans_ZJFA_R(60)
                Dim ZRXHL2ZRGL As Double = ans_ZJFA_R(61)
                '将不合理的工况序号显示出了
                Dim XianShi As String = Nothing
                For i = 1 To n
                    b = GKXH(i)
                    '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
                    Dim HSSBGL As Double = 0
                    '忽略为0的工况
                    If b > 0 Then
                        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "空气源热泵" Then
                            HSSBGL = ZJRGL1_KQYRB + ZJRGL2_KQYRB
                        End If
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "水(地)源热泵" Then
                            HSSBGL = ZJRGL1_SDYRB + ZJRGL2_SDYRB
                        End If
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "风冷螺杆机" Then
                            HSSBGL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ
                        End If
                        '检查混水供热的两种设备的装机功率比例和输入的混水设备比例的大小关系，如果输入的比例大于实际装机比例，报错
                        If HSGRGLBL > HSSBGL / (HSSBGL + TRQGL1ZRGL + TRQGL2ZRGL + ZRXHL1ZRGL + ZRXHL2ZRGL + DCNGL1ZRGL + DCNGL2ZRGL) Then
                            Dim XXX As String = "(" & b & ")、"
                            XianShi = XianShi & XXX.ToString & "  "
                        End If
                    End If
                Next
                If XianShi <> Nothing Then
                    MsgBox("输入的<混水供热供冷比例系数>大于装机选择的两种设备的实际比例，计算结果可能会不正确！工况序号为： " & XianShi)
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            For i = 1 To 5
                b = GKXH(i)
                If b > 0 Then '忽略为0的工况
                    Call mainprogram.清空指定工况输入输出数据(ExcelApp, b)
                    '进行正常的负荷分析（主要技术指标）计算
                    Call mainprogram.负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， 0, 0， 0, 0， 0, 0， 0, 0， 0, 0， 0, calculation_mode, TJGRFHBL, HSGRGLBL)
                    '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call mainprogram.制冷和蓄冷空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    '只有全局寻优计算模式才修正
                    Call mainprogram.制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
                    Call mainprogram.制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
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
        Dim GKXH(6) As Integer
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
            '判断是否需要混水供热或者梯级供热，并输入相关计算比例
            Dim TJGRZTJC = 0 '梯级供热状态监测
            Dim HSGRZTJC = 0 '混水供热状态监测
            For i = 1 To 5
                b = GKXH(i)
                '忽略为0的工况
                If b > 0 Then
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing Then '如果有仅计算耗电量，不计算供热量的设备
                        TJGRZTJC = TJGRZTJC + 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing Then '如果有混水供热的设备
                        HSGRZTJC = HSGRZTJC + 1
                    End If
                End If
            Next
            '输入体积供热比例
            Dim TJGRFHBL As Double = 0
            If TJGRZTJC > 0 Then
                '让用户手动输入负荷率比例（一级热泵负荷率/二级热泵负荷率），一级热泵跟随二级热泵负荷率的变化而变化，在二级热泵负荷率基础上乘以系数，即为一级热泵负荷率
                TJGRFHBL = InputBox("请输入梯级供热一级热泵（仅计算耗电量，不计算供热量设备）与二级热泵负荷比例系数，（一级热泵负荷比例÷二级热泵负荷比例）", "请输入梯级供热设备负荷比例系数", 1)
            End If
            '存在混水计算，输入混水计算比例
            Dim HSGRGLBL As Double = 0
            If HSGRZTJC > 0 Then
                '让用户手动输入混水设备制热功率，占总热功率（总热功率指的是，几种混水向外供热的设备总供热功率的合计）的比例
                HSGRGLBL = InputBox("请输入混水供热功率比例系数（选择的混水供热设备的供热功率，占两种设备总供热功率的比例，例如(风冷热泵的制热功率/（天然气锅炉的制热功率+风冷热泵的制热功率）)）", "请输入混水供热功率比例系数", 0.5)
                '读取天然气锅炉、直燃型溴化锂机组、电采暖锅炉制热总功率（装机量，制热出力最大值）
                '读取采暖季装机方案及参数
                Dim ans_ZJFA_R = mainprogram.读取采暖季装机方案参数(ExcelApp)
                '天然气锅炉装机功率（总和）
                Dim TRQGL1ZRGL As Double = ans_ZJFA_R(12)
                Dim TRQGL2ZRGL As Double = ans_ZJFA_R(13)
                '电锅炉装机功率（总和）
                Dim DCNGL1ZRGL As Double = ans_ZJFA_R(20)
                Dim DCNGL2ZRGL As Double = ans_ZJFA_R(21)
                '风冷螺杆机（总和）
                Dim ZJRGL1_FLLGJ As Double = ans_ZJFA_R(28)
                Dim ZJRGL2_FLLGJ As Double = ans_ZJFA_R(29)
                '水地源热泵（总和）
                Dim ZJRGL1_SDYRB As Double = ans_ZJFA_R(36)
                Dim ZJRGL2_SDYRB As Double = ans_ZJFA_R(37)
                '空气源热泵（总和）
                Dim ZJRGL1_KQYRB As Double = ans_ZJFA_R(52)
                Dim ZJRGL2_KQYRB As Double = ans_ZJFA_R(53)
                '直燃型溴化锂装机功率（总和）
                Dim ZRXHL1ZRGL As Double = ans_ZJFA_R(60)
                Dim ZRXHL2ZRGL As Double = ans_ZJFA_R(61)
                '将不合理的工况序号显示出了
                Dim XianShi As String = Nothing
                For i = 1 To n
                    b = GKXH(i)
                    '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
                    Dim HSSBGL As Double = 0
                    '忽略为0的工况
                    If b > 0 Then
                        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "空气源热泵" Then
                            HSSBGL = ZJRGL1_KQYRB + ZJRGL2_KQYRB
                        End If
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "水(地)源热泵" Then
                            HSSBGL = ZJRGL1_SDYRB + ZJRGL2_SDYRB
                        End If
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value = "风冷螺杆机" Then
                            HSSBGL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ
                        End If
                        '检查混水供热的两种设备的装机功率比例和输入的混水设备比例的大小关系，如果输入的比例大于实际装机比例，报错
                        If HSGRGLBL > HSSBGL / (HSSBGL + TRQGL1ZRGL + TRQGL2ZRGL + ZRXHL1ZRGL + ZRXHL2ZRGL + DCNGL1ZRGL + DCNGL2ZRGL) Then
                            Dim XXX As String = "(" & b & ")、"
                            XianShi = XianShi & XXX.ToString & "  "
                        End If
                    End If
                Next
                If XianShi <> Nothing Then
                    MsgBox("输入的<混水供热供冷比例系数>大于装机选择的两种设备的实际比例，计算结果可能会不正确！工况序号为： " & XianShi)
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            For i = 1 To 5
                b = GKXH(i)
                If b > 0 Then '忽略为0的工况
                    Call mainprogram.清空指定工况输入输出数据(ExcelApp, b)
                    '进行正常的负荷分析（主要技术指标）计算
                    Call mainprogram.负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， D_price， D_price， D_price， D_price， D_price， D_price， D_price， D_price， D_price， D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call mainprogram.制冷和蓄冷空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    Call mainprogram.制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                    '只有全局寻优计算模式才修正
                    Call mainprogram.制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
                    Call mainprogram.制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
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