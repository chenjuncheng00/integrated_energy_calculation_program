Module 主计算程序
    Sub 计算主程序(ExcelApp As Object, precision As Double, GDDJ_GF1 As Double, GDDJ_GF2 As Double, GDDJ_F1 As Double, GDDJ_F2 As Double, GDDJ_P1 As Double, GDDJ_P2 As Double, GDDJ_G1 As Double, GDDJ_G2 As Double, GDDJ_QT1 As Double, GDDJ_QT2 As Double, TRQDJ As Double, JSMS As Integer, TS_XZ As Boolean)
        '本SUB为主程序
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        'precision：是用户在（计算模式选择）的窗体内输入负荷调节精度
        'calculation_mode ：用户选择的计算模式，1表示常规计算模式，2表示全局寻优计算模式
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Call 解锁工作表(ExcelApp)
        Call 清空输入输出数据(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计数，统计一共有多少种不同工况
        Dim n As Integer = 0
        For i = 207 To 8 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(i, 24).Value > 0 Then
                n = i - 7 '工况总数
                Exit For '跳出循环
            End If
        Next
        If n = 0 Then
            MsgBox("负荷段全年天数不能为空，请重新输入！")
            Call 锁定工作表(ExcelApp)
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '将输入的冷热负荷仅保留最大3位小数，增加美观性
        For i = 1 To n
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value > 0 Then
                Dim Temp = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value = Math.Round(Temp, 3)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                Dim Temp = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value = Math.Round(Temp, 3)
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '判断选择的制冷制热设备是否正确，有错误则报错并终止计算
        Dim ZTJC_EQ As Integer = 判断制冷制热设备选择是否正确(ExcelApp, n)
        If ZTJC_EQ = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EQ = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '将已有的梯级供热设备和混水供热设备负荷率清空
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 81), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(57, 81)).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 83), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(57, 83)).Value = Nothing
        '读取输入的各工况冷负荷需求量、热负荷需求量、蓄冷量、蓄热量
        If n > 0 Then
            '让用户输入负荷调节精度
            Dim FHTJJD As Double = precision
            '计算模式
            Dim calculation_mode As Integer = JSMS
            '购电单价
            Dim D_price_GF1 As Double = GDDJ_GF1
            Dim D_price_GF2 As Double = GDDJ_GF2
            Dim D_price_F1 As Double = GDDJ_F1
            Dim D_price_F2 As Double = GDDJ_F2
            Dim D_price_P1 As Double = GDDJ_P1
            Dim D_price_P2 As Double = GDDJ_P2
            Dim D_price_G1 As Double = GDDJ_G1
            Dim D_price_G2 As Double = GDDJ_G2
            Dim D_price_QT1 As Double = GDDJ_QT1
            Dim D_price_QT2 As Double = GDDJ_QT2
            '天然气单价
            Dim TRQ_price As Double = TRQDJ
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '判断输入的各种负荷率是否有错误，有错误则报错并终止计算
            Dim ZTJC_SHUJU As Integer = 读取输入的各种数据并添加报错功能(ExcelApp, FHTJJD, n)
            If ZTJC_SHUJU = 1 Then
                Call 锁定工作表(ExcelApp)
                ZTJC_SHUJU = 0
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '针对输入的负荷调节精度，添加报错功能
            If FHTJJD = Nothing Then '输入的调节精度为空的情况
                MsgBox("输入的负荷调节精度不能为空，请重新输入！")
                Call 锁定工作表(ExcelApp)
                Exit Sub
            ElseIf FHTJJD = 0 Then '输入的调节精度为0的情况
                MsgBox("输入的负荷调节精度不能为0，请重新输入！")
                Call 锁定工作表(ExcelApp)
                Exit Sub
            ElseIf (FHTJJD <> 0 And (25 - FHTJJD * CInt(25 / FHTJJD)) <> 0) Then '输入的调节精度不能被25整除的情况
                MsgBox("输入的负荷调节精度必需能够被25整除，请重新输入！")
                Call 锁定工作表(ExcelApp)
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '根据用户输入的负荷调节精度，计算出最大计算步长
            Dim JSBC As Integer = CInt(25 / FHTJJD)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '判断梯级供热设备（仅计算耗电量，不计算供热量的设备）和混水供热设备是否存在
            Dim TJGRZTJC = 0 '梯级供热状态监测
            Dim HSGRZTJC = 0 '混水供热状态监测
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Then '如果有仅计算耗电量，不计算供热量的设备
                    TJGRZTJC = TJGRZTJC + 1
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '如果有混水供热的设备
                    HSGRZTJC = HSGRZTJC + 1
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
                '读取天然气锅炉、直燃型溴化锂机组、电采暖锅炉制热总功率（装机量，制热出力最大值）
                '读取采暖季装机方案及参数
                Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
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
                '如果选择的天然气锅炉、电采暖锅炉、直燃型溴化锂中，存在两种或者三种，则禁止计算（暂时不支持这么复杂的功能）
                Dim JC_NUM As Integer = 0
                If TRQGL1ZRGL + TRQGL2ZRGL > 0 Then
                    JC_NUM = JC_NUM + 1
                End If
                If DCNGL1ZRGL + DCNGL2ZRGL > 0 Then
                    JC_NUM = JC_NUM + 1
                End If
                If ZRXHL1ZRGL + ZRXHL2ZRGL > 0 Then
                    JC_NUM = JC_NUM + 1
                End If
                If JC_NUM = 0 Then
                    MsgBox("混水供热计算必须选择<天然气采暖锅炉>、<电采暖锅炉>、<直燃型溴化锂>中的一种设备。")
                    Exit Sub
                End If
                If JC_NUM > 1 Then
                    MsgBox("混水供热计算只可以选择<天然气采暖锅炉>、<电采暖锅炉>、<直燃型溴化锂>中的一种设备，当前选择的设备种类大于一种，暂不支持这种计算，请联系开发者或重新选择。")
                    Exit Sub
                End If
                '让用户手动输入混水设备制热功率，占总热功率（总热功率指的是，几种混水向外供热的设备总供热功率的合计）的比例
                HSGRGLBL = InputBox("请输入混水供热功率比例系数（选择的混水供热设备的供热功率，占两种设备总供热功率的比例，例如(风冷热泵的制热功率/（天然气锅炉的制热功率+风冷热泵的制热功率）)）", "请输入混水供热功率比例系数", 0.5)
                '将不合理的工况序号显示出了
                Dim XianShi As String = Nothing
                For i = 1 To n
                    '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
                    Dim HSSBGL As Double = 0
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
                    If HSGRGLBL < （HSSBGL / (HSSBGL + TRQGL1ZRGL + TRQGL2ZRGL + ZRXHL1ZRGL + ZRXHL2ZRGL + DCNGL1ZRGL + DCNGL2ZRGL)） * 0.9 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then
                        Dim XXX As String = "(" & i & ")、"
                        XianShi = XianShi & XXX.ToString & "  "
                    ElseIf HSGRGLBL > (HSSBGL / (HSSBGL + TRQGL1ZRGL + TRQGL2ZRGL + ZRXHL1ZRGL + ZRXHL2ZRGL + DCNGL1ZRGL + DCNGL2ZRGL)) * 1.1 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then
                        Dim XXX As String = "(" & i & ")、"
                        XianShi = XianShi & XXX.ToString & "  "
                    End If
                Next
                If XianShi <> Nothing Then
                    MsgBox("输入的<混水供热供冷比例系数>与装机选择的两种设备的实际比例存在一定的偏差，计算结果可能会不正确！工况序号为： " & XianShi)
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '先进行一遍正常的负荷分析计算
            '此时的计算循环起点=1，终点=n
            'b表示当前正在计算的工况序号
            For b = 1 To n
                '进行正常的负荷分析（主要技术指标）计算
                Call 负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                Call 制冷和蓄冷空调设备负荷率修正(ExcelApp, b, calculation_mode)
                Call 制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                '计算制冷季和制热季天然气耗量和耗电量综合修正系数
                Call 制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
                Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '对于全局寻优计算模式中，如果制冷设备或者制热设备会被全部利用，则根据记录下的工况序号，直接进行常规模式计算，不报错
            Dim QJXY_CGMS_BH As String = Nothing '用于在Msgbox中显示
            '新的列表，用来储存筛选后的，没有重复的工况序号
            Dim XH_QJXY_CGMS As New List(Of Integer)
            If QJXY_CGMS.Count > 0 And calculation_mode = 2 Then
                For i = 0 To QJXY_CGMS.Count - 1
                    If QJXY_CGMS(i) > QJXY_CGMS(i - 1) Then
                        Dim XH As String = "(" & QJXY_CGMS(i) & ")" & "  "
                        QJXY_CGMS_BH = QJXY_CGMS_BH & XH
                        '加入新的列表
                        XH_QJXY_CGMS.Add(QJXY_CGMS(i))
                    End If
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '全局寻优计算出现错误的工况序号列表长度
            '针对计算出错的工况，采用常规模式重新进行计算
            Dim QJXY_ERROR_BH As String = Nothing '用于在Msgbox中显示
            '新的列表，用来储存筛选后的，没有重复的全局寻优计算模式出现错误的工况序号
            Dim XH_QJXY_ERROR As New List(Of Integer)
            If QJXYJS_ERROR.Count > 0 And calculation_mode = 2 Then
                For i = 0 To QJXYJS_ERROR.Count - 1
                    If QJXYJS_ERROR(i) > QJXYJS_ERROR(i - 1) Then
                        Dim XH As String = "(" & QJXYJS_ERROR(i) & ")" & "  "
                        QJXY_ERROR_BH = QJXY_ERROR_BH & XH
                        '加入新的列表
                        XH_QJXY_ERROR.Add(QJXYJS_ERROR(i))
                    End If
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '将两个列表合并（全局寻优模式转为常规模式的工况）
            Dim QJXY_TO_CGMS As New List(Of Integer)
            QJXY_TO_CGMS.AddRange(XH_QJXY_CGMS)
            QJXY_TO_CGMS.AddRange(XH_QJXY_ERROR)
            Dim QJXY_TO_CGMS_BH As String = Nothing '用于在Msgbox中显示
            '新的列表，用来储存筛选后的，没有重复的工况序号
            Dim XH_QJXY_TO_CGMS As New List(Of Integer)
            If QJXY_TO_CGMS.Count > 0 And calculation_mode = 2 Then
                For i = 0 To QJXY_TO_CGMS.Count - 1
                    If QJXY_TO_CGMS(i) > QJXY_TO_CGMS(i - 1) Then
                        Dim XH As String = "(" & QJXY_TO_CGMS(i) & ")" & "  "
                        QJXY_TO_CGMS_BH = QJXY_TO_CGMS_BH & XH
                        '加入新的列表
                        XH_QJXY_TO_CGMS.Add(QJXY_TO_CGMS(i))
                    End If
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            If XH_QJXY_TO_CGMS.Count > 0 And calculation_mode = 2 Then
                '提醒
                If TS_XZ = True Then
                    If XH_QJXY_CGMS.Count > 0 Then
                        MsgBox("全局寻优计算中部分工况制冷制热设备会被全部使用，程序会自动采用常规模型重算错误工况！" & "工况序号为： " & QJXY_CGMS_BH)
                    End If
                    If XH_QJXY_ERROR.Count > 0 Then
                        MsgBox("全局寻优计算存在计算错误的工况，程序会自动采用常规模型重算错误工况！" & "工况序号为： " & QJXY_ERROR_BH)
                    End If
                End If
                '采用模式1进行计算
                calculation_mode = 1
                '重新开始计算
                For i = 0 To XH_QJXY_TO_CGMS.Count - 1
                    '错误的工况序号
                    Dim a As Integer = XH_QJXY_TO_CGMS(i)
                    '调用程序进行计算
                    Call 清空指定工况输入输出数据(ExcelApp, a)
                    '进行正常的负荷分析（主要技术指标）计算
                    Call 负荷分析计算程序(ExcelApp, a, FHTJJD, JSBC， D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call 制冷和蓄冷空调设备负荷率修正(ExcelApp, a, calculation_mode)
                    Call 制热和蓄热空调设备负荷率修正(ExcelApp, a, calculation_mode)
                    '计算制冷季和制热季天然气耗量和耗电量综合修正系数
                    Call 制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, a, FHTJJD, calculation_mode)
                    Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, a, FHTJJD, calculation_mode)
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '常规计算模式，出现计算错误的工况序号，加入列表
            Dim CGJSMS_ERROR As New List(Of Integer)
            '计算错误，调到这里重算
cgjsms_again:
            'FHTJJD参数最低是0.1
            If calculation_mode = 1 And FHTJJD > 0.1 And CGJSMS_ERROR.Count > 0 Then
                '针对寻找出来的计算错误的工况序号，减小FHTJJD参数，重新进行计算
                Dim CGJS_ERROR_BH As String = Nothing '用于在Msgbox中显示
                '新的列表，用来储存筛选后的，没有重复的常规计算模式出现错误的工况序号
                Dim XH_CGJS_ERROR As New List(Of Integer)
                If CGJSMS_ERROR.Count > 0 Then
                    For i = 0 To CGJSMS_ERROR.Count - 1
                        If CGJSMS_ERROR(i) > CGJSMS_ERROR(i - 1) Then
                            Dim XH As String = "(" & CGJSMS_ERROR(i) & ")" & "  "
                            CGJS_ERROR_BH = CGJS_ERROR_BH & XH
                            '加入新的列表
                            XH_CGJS_ERROR.Add(CGJSMS_ERROR(i))
                        End If
                    Next
                    If TS_XZ = True Then
                        MsgBox("常规计算模式存在计算错误的工况，程序会自动修改<负荷调节精度>参数重新计算！" & "工况序号为： " & CGJS_ERROR_BH)
                    End If
                    '修改FHTJJD和JSBC参数
                    If FHTJJD > 0.5 Then
                        FHTJJD = 0.5
                    ElseIf FHTJJD <= 0.5 And FHTJJD > 0.25 Then
                        FHTJJD = 0.25
                    Else
                        FHTJJD = 0.1
                    End If
                    JSBC = CInt(25 / FHTJJD)
                    '重新开始计算
                    For i = 0 To XH_CGJS_ERROR.Count - 1
                        '错误的工况序号
                        Dim a As Integer = XH_CGJS_ERROR(i)
                        '调用程序进行计算
                        Call 清空指定工况输入输出数据(ExcelApp, a)
                        '进行正常的负荷分析（主要技术指标）计算
                        Call 负荷分析计算程序(ExcelApp, a, FHTJJD, JSBC， D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '对计算出的制冷和制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                        Call 制冷和蓄冷空调设备负荷率修正(ExcelApp, a, calculation_mode)
                        Call 制热和蓄热空调设备负荷率修正(ExcelApp, a, calculation_mode)
                        '计算制冷季和制热季天然气耗量和耗电量综合修正系数
                        Call 制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, a, FHTJJD, calculation_mode)
                        Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, a, FHTJJD, calculation_mode)
                    Next
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算循环体
            Call 计算循环体(ExcelApp, n)
            '————————————————————————————————————————————————————————————————————————————————————————
            '寻找计算结果出现错误的工况序号，加入列表
            If calculation_mode = 1 And FHTJJD > 0.1 Then
                '清空列表
                CGJSMS_ERROR.Clear()
                '寻找计算结果出现错误的工况序号，加入列表
                For i = 1 To n
                    If ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 18).Value = "不正确" Or ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 35).Value = "不正确" Or ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 49).Value = "不正确" Or ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 61).Value = "不正确" Or ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 62).Value = "不正确" Or ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 63).Value = "不正确" Then
                        CGJSMS_ERROR.Add(i)
                    End If
                Next
                If CGJSMS_ERROR.Count > 0 Then
                    '跳至标签处重算
                    GoTo cgjsms_again
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '判断各种计算结果是否正确，不正确则报错
            Call 判断各种计算结果是否正确(ExcelApp, FHTJJD, n)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '—————————————————————————————————————————————————————————————————————————————————————————
        '实例化一个计算过程显示窗体
        Dim Calculate_Progress As New 计算进度显示
        '在窗体中显示计算已完成
        Calculate_Progress.Show()
        Calculate_Progress.Label1.Text = "计算已经完成，请查看计算结果！"
        Calculate_Progress.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '重新锁定工作表
        Call 锁定工作表(ExcelApp)
    End Sub

    Sub 负荷分析计算程序(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer， D_price_GF1 As Double, D_price_GF2 As Double, D_price_F1 As Double, D_price_F2 As Double, D_price_P1 As Double, D_price_P2 As Double, D_price_G1 As Double, D_price_G2 As Double, D_price_QT1 As Double, D_price_QT2 As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '实例化一个计算过程显示窗体
        Dim Calculate_Progress As New 计算进度显示
        '————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————
        '确定当前的电价
        Dim D_price As Double
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "高峰1" Then
            D_price = D_price_GF1
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "高峰2" Then
            D_price = D_price_GF2
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "峰1" Then
            D_price = D_price_F1
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "峰2" Then
            D_price = D_price_F2
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "平1" Then
            D_price = D_price_P1
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "平2" Then
            D_price = D_price_P2
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "谷1" Then
            D_price = D_price_G1
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "谷2" Then
            D_price = D_price_G2
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "其它1" Then
            D_price = D_price_QT1
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = "其它2" Then
            D_price = D_price_QT2
        Else
            D_price = 1
        End If
        '混水冷负荷热负荷默认值（常规计算模式需要）
        Dim HSLFH As Double = 0
        Dim HSRFH As Double = 0
        '判断内燃机是否可以向外供电
        '先计算内燃机可以向外供电的情况
        '当处于内燃机可以向外供电的情况下，需要手动输入各个内燃机的负荷率，程序会按照这个负荷率进行计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "Y" Then
            '在窗体显示目前的计算进度，显示目前在计算第几个工况
            Calculate_Progress.Show()
            Calculate_Progress.TopMost = True
            Calculate_Progress.Label1.Text = "计算进行中，目前正在计算的工况序号为：" & b
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            ' 将已有的所有设备计算过程量清空
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97)).Value = Nothing
            '各种修正系数默认值均为1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 109)).Value = 1
            '全局寻优计算模式设备启动数量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = Nothing
            '混水冷负荷热负荷计算（常规计算模式需要）
            Dim ans_HSJS = 混水供冷供热量计算(ExcelApp, b, calculation_mode)
            HSLFH = ans_HSJS(0)
            HSRFH = ans_HSJS(1)
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '当内燃机的向外供电量没有任何限制时，采用以下计算
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(28, 9).Value = "N" Then
                '如果内燃机余热不可以被浪费，则进行下列计算
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 78).Value = "N" Then
                    '此时内燃机负荷率先设置为1，如果蓄能装置供冷供热量和溴化锂供冷供热量大于了冷热负荷总需求量，则内燃机降负荷，直到蓄能装置供冷供热量和溴化锂供冷供热量不大于冷热负荷总需求量
                    '如果蓄冷量为0，冷负荷需求量大于0，则溴化锂的制冷量只用于供冷
                    If (XNXLGL(b) = 0 And LFHZXQL(b) > 0) Then
                        '调用子程序进行计算——制冷，不同设备负荷率计算
                        '第一步，判断内燃机负荷率是否需要被修正
                        Call 内燃机可以向外供电且内燃机余热不可以浪费时供冷蓄冷内燃机负荷率调节(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                        '第二步，进行制冷设备计算
                        Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果蓄热量为0，热负荷需求量大于0，则溴化锂的制热量只用于供热
                    If (XNXRGL(b) = 0 And RFHZXQL(b) > 0) Then
                        '第一步，判断内燃机负荷率是否需要被修正
                        Call 内燃机可以向外供电且内燃机余热不可以浪费时供热蓄热内燃机负荷率调节(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '第二步，进行制热设备计算
                        Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                    '如果蓄冷量大于0，冷负荷需求量大于等于0，则溴化锂制冷量不仅用于供冷，也可以用于蓄冷
                    If (XNXLGL(b) > 0 And LFHZXQL(b) >= 0) Then
                        '判断内燃机负荷率是否需要被修正
                        Call 内燃机可以向外供电且内燃机余热不可以浪费时供冷蓄冷内燃机负荷率调节(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                        '进行计算
                        Call 内燃机可以向外供电时制冷和蓄冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果蓄热量大于0，热负荷需求量大于等于0，则溴化锂制热量不仅用于供热，也可以用于蓄热
                    If (XNXRGL(b) > 0 And RFHZXQL(b) >= 0) Then
                        '判断内燃机负荷率是否需要被修正
                        Call 内燃机可以向外供电且内燃机余热不可以浪费时供热蓄热内燃机负荷率调节(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '进行计算
                        Call 内燃机可以向外供电时制热和蓄热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '如果内燃机余热可以被浪费，则进行系列计算
                ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 78).Value = "Y" Then
                    '此时内燃机负荷率按照用户输入的负荷率进行计算，如果溴化锂供冷供热量和蓄能装置供冷供热量大于冷热负荷总需求量，多余的溴化锂制冷供热量浪费掉
                    '如果蓄冷量为0，冷负荷需求量大于0，则溴化锂的制冷量只用于供冷
                    If (XNXLGL(b) >= 0 And LFHZXQL(b) >= 0) Then
                        Call 内燃机可以向外供电且内燃机余热可以浪费时供冷和蓄冷计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果蓄热量为0，热负荷需求量大于0，则溴化锂的制热量只用于供热
                    If (XNXRGL(b) >= 0 And RFHZXQL(b) >= 0) Then
                        Call 内燃机可以向外供电且内燃机余热可以浪费时供热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                    '————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '————————————————————————————————————————————————————————————————————————————————————————————————————————
                Else '如果内燃机余热是否浪费的选项留空，则进行一下计算
                    '按照用户输入的内燃机负荷率直接计算，不修正负荷率
                    '如果蓄冷量为0，冷负荷需求量大于0，则溴化锂的制冷量只用于供冷
                    If (XNXLGL(b) = 0 And LFHZXQL(b) > 0) Then
                        '调用子程序进行计算——制冷，不同设备负荷率计算
                        '进行制冷设备计算
                        Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果蓄热量为0，热负荷需求量大于0，则溴化锂的制热量只用于供热
                    If (XNXRGL(b) = 0 And RFHZXQL(b) > 0) Then
                        '进行制热设备计算
                        Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                    '如果蓄冷量大于0，冷负荷需求量大于等于0，则溴化锂制冷量不仅用于供冷，也可以用于蓄冷
                    If (XNXLGL(b) > 0 And LFHZXQL(b) >= 0) Then
                        '进行计算
                        Call 内燃机可以向外供电时制冷和蓄冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果蓄热量大于0，热负荷需求量大于等于0，则溴化锂制热量不仅用于供热，也可以用于蓄热
                    If (XNXRGL(b) > 0 And RFHZXQL(b) >= 0) Then
                        '进行计算
                        Call 内燃机可以向外供电时制热和蓄热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '当内燃机的向外供电量存在最高上限时，采用以下计算
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(28, 9).Value = "Y" Then
                '采用内燃机不可以向外供电的代码进行计算，采用新的供电上限
                '将本工况的冷热负荷需求量和向外供电上限带入，冷热负荷需求量的带入是为了判断本工况是制冷还是制热
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 12).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 12).Value '冷负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 21).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 21).Value '热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 27).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 27).Value '向外供电上限
                '如果此时处于内燃机不启动的时间段，则关闭所有内燃机，内燃机负荷率为0
                If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(24, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(25, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(26, 9).Value) Then
                    '将已有的制冷制热设备运行负荷率，蓄冷蓄热工况设备启动情况清空
                    Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                    '内燃机负荷率赋值为0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0
                    '使用内燃机可以向外供电情况下的代码进行计算
                    '第一步，进行制冷设备计算
                    Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    '第二步，进行制热设备计算
                    Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    '—————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '—————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '如果此时不处于内燃机不启动的时间段，则会开启内燃机，自动计算内燃机负荷率
                Else
                    '如果本工况蓄冷功率为0同时冷负荷总需求量大于0，则进行下列计算（蓄冷功率为0时，内燃机发电量只需要平衡掉制冷空调设备的耗电即可）
                    If (XNXLGL(b) = 0 And LFHZXQL(b) > 0) Then
                        '将已有的制冷和蓄冷设备运行负荷率重置为0，每一次重新计算必需重置一次0
                        Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                        '调用子程序进行计算——制冷不同设备负荷率计算
                        '进行制冷设备计算
                        Call 内燃机不可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '如果本工况蓄热功率为0同时热负荷总需求量大于0，则进行下列计算（蓄热功率为0时，内燃机发电量只需要平衡掉制热空调设备的耗电即可）
                    If (XNXRGL(b) = 0 And RFHZXQL(b) > 0) Then
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '调用子程序进行计算——制热不同设备负荷率计算
                        '进行制热设备计算
                        Call 内燃机不可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                    '当本工况蓄冷功率大于0时同时冷负荷总需求量大于0，制冷内燃机的发电功率需要能够抵消掉制冷设备和蓄冷设备的总耗电量
                    If (XNXLGL(b) > 0 And LFHZXQL(b) >= 0) Then
                        '将已有的制冷和蓄冷设备运行负荷率重置为0，每一次重新计算必需重置一次0
                        Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                        '进行计算
                        Call 内燃机不可以向外供电时制冷和蓄冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                    End If
                    '当本工况蓄热功率大于0时同时热负荷总需求量大于0，制热内燃机的发电功率需要能够抵消掉制热设备和蓄热设备的总耗电量
                    If (XNXRGL(b) > 0 And RFHZXQL(b) >= 0) Then
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '进行计算
                        Call 内燃机不可以向外供电时制热和蓄热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    End If
                End If
            End If
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '内燃机不可以向外供电的情况计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value <> "Y" Then
            '在窗体显示目前的计算进度，显示目前在计算第几个工况
            Calculate_Progress.Show()
            Calculate_Progress.TopMost = True
            Calculate_Progress.Label1.Text = "计算进行中，目前正在计算的工况序号为：" & b
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '清空需要计算的这个工况已经输入的内燃机负荷率数据
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = Nothing
            ' 将已有的所有设备计算过程量清空
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97)).Value = Nothing
            '各种修正系数默认值均为1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 109)).Value = 1
            '全局寻优计算模式设备启动数量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = Nothing
            '混水冷负荷热负荷计算（常规计算模式需要）
            Dim ans_HSJS = 混水供冷供热量计算(ExcelApp, b, calculation_mode)
            HSLFH = ans_HSJS(0)
            HSRFH = ans_HSJS(1)
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '如果此时处于内燃机不启动的时间段，则关闭所有内燃机，内燃机负荷率为0
            If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(24, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(25, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(26, 9).Value) Then
                '将已有的制冷制热设备运行负荷率，蓄冷蓄热工况设备启动情况清空
                Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '内燃机负荷率赋值为0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0
                '使用内燃机可以向外供电情况下的代码进行计算
                '第一步，进行制冷设备计算
                Call 内燃机可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                '第二步，进行制热设备计算
                Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                '—————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '—————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '如果此时不处于内燃机不启动的时间段，则会开启内燃机，自动计算内燃机负荷率
            Else
                '如果本工况蓄冷功率为0同时冷负荷总需求量大于0，则进行下列计算（蓄冷功率为0时，内燃机发电量只需要平衡掉制冷空调设备的耗电即可）
                If (XNXLGL(b) = 0 And LFHZXQL(b) > 0) Then
                    '将已有的制冷和蓄冷设备运行负荷率重置为0，每一次重新计算必需重置一次0
                    Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                    '调用子程序进行计算——制冷不同设备负荷率计算
                    '第一步，进行制冷设备计算
                    Call 内燃机不可以向外供电时制冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                End If
                '如果本工况蓄热功率为0同时热负荷总需求量大于0，则进行下列计算（蓄热功率为0时，内燃机发电量只需要平衡掉制热空调设备的耗电即可）
                If (XNXRGL(b) = 0 And RFHZXQL(b) > 0) Then
                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                    '调用子程序进行计算——制热不同设备负荷率计算
                    '进行制热设备计算
                    Call 内燃机不可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                End If
                '当本工况蓄冷功率大于0时同时冷负荷总需求量大于0，制冷内燃机的发电功率需要能够抵消掉制冷设备和蓄冷设备的总耗电量
                If (XNXLGL(b) > 0 And LFHZXQL(b) >= 0) Then
                    '将已有的制冷和蓄冷设备运行负荷率重置为0，每一次重新计算必需重置一次0
                    Call 清空制冷和蓄冷设备负荷率计算结果(ExcelApp, b)
                    '进行计算
                    Call 内燃机不可以向外供电时制冷和蓄冷设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSLFH， D_price, TRQ_price, calculation_mode)
                End If
                '当本工况蓄热功率大于0时同时热负荷总需求量大于0，制热内燃机的发电功率需要能够抵消掉制热设备和蓄热设备的总耗电量
                If (XNXRGL(b) > 0 And RFHZXQL(b) >= 0) Then
                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                    '进行计算
                    Call 内燃机不可以向外供电时制热和蓄热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                End If
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————       
        '将上面计算出的内燃机负荷率中，单台负荷率低于30%的内燃机负荷率修改为0
        Call 将内燃机单台负荷率低于百分之30的内燃机关闭(ExcelApp, b, FHTJJD, JSBC, HSLFH, HSRFH， D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        '————————————————————————————————————————————————————————————————————————————————————————
        Calculate_Progress.Close()
    End Sub
    Sub 蓄能分配计算主程序(ExcelApp As Object, XLGL_PJ As Double, XLGL_MAX As Double, XLJS_MS As Integer, XRGL_PJ As Double, XRGL_MAX As Double, XRJS_MS As Integer, QT1_SJD As Boolean, QT2_SJD As Boolean)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        Call 解锁工作表(ExcelApp)
        Call 清空输入输出数据(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计数，统计一共有多少种不同工况
        Dim n As Integer = 0
        For i = 207 To 8 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(i, 24).Value > 0 Then
                n = i - 7 '工况总数
                Exit For '跳出循环
            End If
        Next
        If n > 0 Then
            '清空已经输入的蓄冷、蓄热数据
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 13), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(207, 14)).Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 22), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(207, 23)).Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 13), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 14)).Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 23)).Value = Nothing
            '申明局部变量和数组
            Dim RQXH(10000) As Integer '日期序号
            Dim GKXSS(10000) As Double '每个工况的小时数
            Dim RQXHJS(10000) As Integer '定义数组，用于日期序号计数,每次日期序号发生变化时，最后一个不变的日期序号的工况序号储存在数组中
            Dim c As Integer 'RQXHJS数组的元素数量，有多少种不同的日期序号
            '读取各种计算系数
            Dim ans_XZXS_L = 读取制冷季输入的计算系数(ExcelApp)
            '溴化锂制冷COP
            Dim XHL_COP_L As Double = ans_XZXS_L(10)
            '读取各种修正系数
            Dim ans_XZXS_R = 读取采暖季输入的计算系数(ExcelApp)
            '溴化锂制热COP
            Dim XHL_COP_R As Double = ans_XZXS_R(9)
            '将第一个工况的内燃机余热利用方式带入计算，内燃机负荷率设置为1
            '制冷时烟气热水型溴化锂功率计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 77).Value
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3).Value = 1
            Dim XHLZLa As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(172, 4).Value * XHL_COP_L
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2).Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3).Value = Nothing
            '制热时烟气热水型溴化锂功率计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 1
            Dim XHLZRa As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(173, 4).Value * XHL_COP_R
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = Nothing
            '制冷装机总功率（不包括蓄冷装置）
            Dim ZLZJGL As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(58, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(57, 7).Value + XHLZLa
            '制热装机总功率（不包括蓄热装置）
            Dim ZRZJGL As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(69, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(68, 7).Value + XHLZRa
            '蓄冷装机总功率（不包括蓄冷装置、烟气热水型溴化锂、直燃性溴化锂）
            Dim XLZJGL As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(58, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(57, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
            '蓄热装机总功率（不包括蓄热装置、烟气热水型溴化锂、直燃性溴化锂、天然气锅炉）
            Dim XRZJGL As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(69, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(68, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
            '如果存在梯级供热的情况，则修正制热装机功率（梯级供热设置只计算耗电量，不计算供热量）
            Dim TJGRZTJC As Double = 0 '梯级供热状态监测
            Dim FLLGJJC_TJGR As Double = 0 '梯级供热使用风冷螺杆机状态监测
            Dim KQYRBJC_TJGR As Double = 0 '梯级供热使用空气源热泵状态监测
            Dim SYRBJC_TJGR As Double = 0 '梯级供热使用水源热泵状态监测
            '统计一个有多少个制热负荷段
            Dim JS_RFH As Integer = 0
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value > 0 Then
                    JS_RFH = JS_RFH + 1
                End If
            Next
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Then '如果有仅计算耗电量，不计算供热量的设备
                    TJGRZTJC = TJGRZTJC + 1
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value = "风冷螺杆机" Then
                    FLLGJJC_TJGR = FLLGJJC_TJGR + 1
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value = "空气源热泵" Then
                    KQYRBJC_TJGR = KQYRBJC_TJGR + 1
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value = "水(地)源热泵" Then
                    SYRBJC_TJGR = SYRBJC_TJGR + 1
                End If
            Next
            If TJGRZTJC = JS_RFH Then '如果存在梯级供热，将制热装机功率减去风冷热泵和空气源热泵装机功率(只有所有热负荷的都梯级供热，才减去相关装机)
                If FLLGJJC_TJGR = JS_RFH Then
                    ZRZJGL = ZRZJGL - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
                End If
                If KQYRBJC_TJGR = JS_RFH Then
                    ZRZJGL = ZRZJGL - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
                End If
                If SYRBJC_TJGR = JS_RFH Then
                    ZRZJGL = ZRZJGL - ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
                End If
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '读取输入的各种数据，添加部分报错功能（并不是所有的报错功能都需要，因此单独提出来重写，没有直接调用模块）
            For a = 1 To n
                LFHZXQL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 12).Value
                RFHZXQL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 21).Value
                '添加一些报错功能
                '如果某工况的用电时间段没有选择，为空，则报错
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 79).Value = Nothing Then
                    MsgBox("本工况用电时间段不能为空，请重新选择！")
                    'ZTJC = 1
                    Call 锁定工作表(ExcelApp)
                    Exit Sub
                End If
                '每个计算工况不可以同时计算供冷和供热，如果一个工况不仅有冷负荷需求，还有热负荷需求，报错
                If (LFHZXQL(a) > 0 And RFHZXQL(a) > 0) Then
                    MsgBox("同一个工况不可以同时计算供冷和供热，必需分开计算！")
                    'ZTJC = 1
                    Call 锁定工作表(ExcelApp)
                    Exit Sub
                End If
                '
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            For a = 1 To n
                '将输入的每一个日期编号计入数组
                RQXH(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 84).Value
                '每个工况的小时数
                GKXSS(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 24).Value * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 25).Value * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 26).Value
            Next
            '针对工况小时数添加报错功能
            For a = 1 To n
                If GKXSS(a) = 0 Then
                    MsgBox("每个负荷段的全年运行天数*每天运行小时数*负荷时间频数不可以等于0！，请重新出入！")
                    'ZTJC = 1
                    Call 锁定工作表(ExcelApp)
                    Exit Sub
                End If
            Next
            '找出每次日期序号发生变化时的工况序号
            RQXHJS(0) = 0
            For a = 1 To n
                If RQXH(a + 1) <> RQXH(a) Then '如果第a+1个工况的日期序号与第a个不相同
                    c = c + 1 'RQXHJS数组的元素数量，有多少种不同的日期序号
                    RQXHJS(c) = a
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '让用户输入蓄冷和蓄热工况的最大功率（kW）
            Dim PJXLGL As Double = XLGL_PJ '平均蓄冷功率            
            Dim PJXRGL As Double = XRGL_PJ '平均蓄热功率
            '如果输入的都是0，则退出计算
            If PJXLGL = 0 And PJXRGL = 0 Then
                Exit Sub
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '确定每个不同的日期序号的个数
            For a = 1 To c 'RQXHJS数组最多只有c个元素
                '每个日期序号相同的一组，对应的工况序号上限和下限
                Dim GKXHmin As Integer = RQXHJS(a - 1) + 1
                Dim GKXHmax As Integer = RQXHJS(a)
                '判断这一组日期序号代表制冷还是制热
                If RQXH(RQXHJS(a)) = 1 Or RQXH(RQXHJS(a)) = 2 Then '制冷工况；RXHCJS(a):日期序号发生变化前的最后一个工况序号；RQXH()：对应的日期序号
                    '求可以蓄冷的最大功率
                    '如果勾选了蓄冷功率存在最大值（模式=1）
                    Dim ZDXLGL As Double = 0 '最大蓄冷功率
                    Dim LFH As New List(Of Double) '输入的每一条冷负荷
                    If XLJS_MS = 1 Then
                        '等于输入的最大蓄冷功率
                        ZDXLGL = XLGL_MAX '最大蓄冷功率
                    Else
                        '遍历所有的冷负荷总需求量
                        For i = GKXHmin To GKXHmax
                            LFH.Add(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 12).Value)
                        Next
                        '等于制冷装机功率与冷负荷的大值
                        ZDXLGL = Math.Max(LFH.Max, ZLZJGL - XHLZLa) + 10 '最大值加10，放大一点，防止出错
                    End If
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    Dim SYXLGL As Double = 0 '剩余蓄冷功率
                    Dim XLZL As Double = 0 '蓄冷总量
                    Dim GLZL As Double = 0 '供冷总量
                    '先计算所有的（谷1、谷2、其它1、其它2）最多可以蓄冷多少kWh
                    Dim GD1_GD2_QT1_QT2_XLZL As Double = 0 '（谷1、谷2、其它1、其它2）蓄冷总量
                    Dim GD1_GD2_QT1_QT2_XSS_L As Double = 0 '冷工况（谷1、谷2、其它1、其它2）小时数
                    For i = GKXHmin To GKXHmax '遍历所有的冷负荷段
                        '蓄冷时间段为：谷1、谷2、其它1、其它2
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            GD1_GD2_QT1_QT2_XLZL = GD1_GD2_QT1_QT2_XLZL + GKXSS(i) * PJXLGL '所有为（谷电、时间段1、时间段2）的小时数乘以输入的最大蓄冷功率
                            GD1_GD2_QT1_QT2_XSS_L = GD1_GD2_QT1_QT2_XSS_L + GKXSS(i) '冷负荷段的（谷电、时间段1、时间段2）总小时数
                        End If
                    Next
                    SYXLGL = SYXLGL + GD1_GD2_QT1_QT2_XLZL '剩余蓄冷功率
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄冷和蓄热
                        '将用电负荷段为（谷1、谷2、其它1、其它2）的工况，蓄冷功率设置为输入的PJXLGL，同时也要满足装机需求（蓄冷功率+制冷功率不可以大于设备总制冷功率）
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            '先寻找蓄冷功率+制冷功率大于设备总制冷功率的情况（谷1、谷2、其它1、其它2时内燃机关闭，需要减去溴化锂制冷量）
                            If LFHZXQL(i) + PJXLGL > ZLZJGL - XHLZLa And SYXLGL > 0 Then
                                '（谷1、谷2、其它1、其它2）时内燃机关闭，需要减去溴化锂制冷量，同时再缩小一点，以免出错
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = Math.Round((ZLZJGL - XHLZLa - LFHZXQL(i)), 2)
                                '统计蓄冷量
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > 0 Then
                                    XLZL = XLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                End If
                            ElseIf LFHZXQL(i) + PJXLGL <= ZLZJGL - XHLZLa And SYXLGL > 0 Then
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = PJXLGL '等于平均蓄冷功率
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > 0 Then
                                    XLZL = XLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                End If
                            End If
                        End If
                    Next
                    '如果还有剩余的蓄冷功率，统计现在有多少个工况点设备装机量没有被完全利用
                    Dim GD1_GD2_QT1_QT2_XSS_L_2 As Double = 0
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄冷和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XLGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value '目前已经有的蓄冷功率
                            If LFHZXQL(i) + XLGL_now < ZLZJGL - XHLZLa And SYXLGL > 0 And XLGL_now = PJXLGL Then '如果目前的蓄冷功率不等于平均蓄冷功率
                                GD1_GD2_QT1_QT2_XSS_L_2 = GD1_GD2_QT1_QT2_XSS_L_2 + GKXSS(i)
                            End If
                        End If
                    Next
                    Dim GD1_GD2_QT1_QT2_XLGL_L_2 As Double = SYXLGL / GD1_GD2_QT1_QT2_XSS_L_2 '平均蓄冷功率2
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄冷和蓄热
                        '将用电负荷段为（谷1、谷2、其它1、其它2）的工况，蓄冷功率设置为输入的GD1_GD2_QT1_QT2_XLGL_L_2，同时也要满足装机需求（蓄冷功率+制冷功率不可以大于设备总制冷功率）
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XLGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value '目前已经有的蓄冷功率
                            If SYXLGL > 0 And LFHZXQL(i) + XLGL_now < ZLZJGL - XHLZLa And XLGL_now = PJXLGL Then '如果剩余蓄冷功率大于0，且装机功率没有被完全利用
                                If ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now)) >= GD1_GD2_QT1_QT2_XLGL_L_2 Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = XLGL_now + GD1_GD2_QT1_QT2_XLGL_L_2
                                    XLZL = XLZL + GD1_GD2_QT1_QT2_XLGL_L_2 * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = XLGL_now + ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now))
                                    XLZL = XLZL + ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now)) * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                End If
                            End If
                        End If
                    Next
                    '如果此时还有剩余没使用的蓄冷功率，则继续蓄冷
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄冷和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XLGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value '目前已经有的蓄冷功率
                            If SYXLGL > 0 And LFHZXQL(i) + XLGL_now <= ZLZJGL - XHLZLa Then '如果剩余蓄冷功率大于0，且装机功率没有被完全利用
                                If ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now)) >= SYXLGL / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = XLGL_now + SYXLGL / GKXSS(i)
                                    XLZL = XLZL + SYXLGL / GKXSS(i) * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = XLGL_now + ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now))
                                    XLZL = XLZL + ((ZLZJGL - XHLZLa) - (LFHZXQL(i) + XLGL_now)) * GKXSS(i) '统计全部蓄冷量
                                    SYXLGL = GD1_GD2_QT1_QT2_XLZL - XLZL '剩下还需要满足的蓄冷量
                                End If
                            End If
                        End If
                    Next
                    '判断计算出的蓄冷功率是否大于输入的最大蓄冷功率
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄冷和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XLGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value '目前已经有的蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > ZDXLGL Then
                                '蓄冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = ZDXLGL
                                '重新计算蓄冷总量
                                XLZL = XLZL + (ZDXLGL - XLGL_now) * GKXSS(i) '统计全部蓄冷量
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    For i = GKXHmin To GKXHmax '第二步，计算进行削峰需要的供冷功率                     
                        '即使是（谷1、谷2、其它1、其它2），如果冷负荷大于了装机，也要进行削峰
                        If LFHZXQL(i) > ZLZJGL Then '冷负荷总需求量大于制冷装机量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = (LFHZXQL(i) - ZLZJGL)
                            GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i) '统计全部供冷量
                        End If
                        '判断当前削峰功率是否大于输入的最大冷负荷，如果是，则报错，但是不修改数值，使得计算可以继续进行
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                            MsgBox("蓄冷装置供冷功率大于输入的最大蓄冷功率，装机方案选择不合理，程序不会自动修改计算出的数值，但请检查并重新选择装机方案！！" & "装机方案不合理的工况序号为： " & i)
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '统计高峰段有多少个小时
                    Dim ZLSL_GF As Double = 0
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If LFHZXQL(i) > 0 Then
                                ZLSL_GF = ZLSL_GF + GKXSS(i) '高峰段总小时数
                            End If
                        End If
                    Next
                    Dim PJGLGL_GF As Double '高峰段每小时平均供冷功率
                    PJGLGL_GF = (XLZL - GLZL) / ZLSL_GF '剩余的几个小时，平均每个小时可以供冷的功率
                    For i = GKXHmin To GKXHmax '第三步，计算高峰段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If LFHZXQL(i) > 0 And PJGLGL_GF > 0 Then '冷负荷大于0才进行计算,平均功率大于0才计算
                                Dim XFGLGL_GF As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value '削峰供冷功率
                                If PJGLGL_GF >= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = LFHZXQL(i)
                                    GLZL = GLZL + LFHZXQL(i) * GKXSS(i) '统计全部供冷量
                                ElseIf PJGLGL_GF + XFGLGL_GF <= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = Math.Round(PJGLGL_GF + XFGLGL_GF, 2)
                                    GLZL = GLZL + PJGLGL_GF * GKXSS(i) '统计全部供冷量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else '冷负荷需求量介于两者之间
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XFGLGL_GF + LFHZXQL(i) - XFGLGL_GF
                                    GLZL = GLZL + (LFHZXQL(i) - XFGLGL_GF) * GKXSS(i) '统计全部供冷量
                                End If
                                ZLSL_GF = ZLSL_GF - GKXSS(i) '剩余的高峰段小时数
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄冷量还有剩余，则继续用于高峰段供冷
                    For i = GKXHmin To GKXHmax '第三步，计算高峰段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If XLZL - GLZL > 0 And LFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value Then
                                '记录目前已经有的蓄冷供冷量
                                Dim XLGLGL_now_GF As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value
                                If (LFHZXQL(i) - XLGLGL_now_GF) >= (XLZL - GLZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_GF + (XLZL - GLZL) / GKXSS(i)
                                    GLZL = GLZL + (XLZL - GLZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_GF + LFHZXQL(i) - XLGLGL_now_GF
                                    GLZL = GLZL + (LFHZXQL(i) - XLGLGL_now_GF) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '统计峰段有多少个小时
                    Dim ZLSL_F As Double = 0
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If LFHZXQL(i) > 0 Then
                                ZLSL_F = ZLSL_F + GKXSS(i) '峰段总小时数
                            End If
                        End If
                    Next
                    Dim PJGLGL_F As Double '峰段每小时平均供冷功率
                    PJGLGL_F = (XLZL - GLZL) / ZLSL_F '剩余的几个小时，平均每个小时可以供冷的功率
                    For i = GKXHmin To GKXHmax '第四步，计算峰段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If LFHZXQL(i) > 0 And PJGLGL_F > 0 Then '冷负荷大于0才进行计算,平均功率大于0才计算
                                Dim XFGLGL_F As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value '削峰供冷功率
                                If PJGLGL_F >= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = LFHZXQL(i)
                                    GLZL = GLZL + LFHZXQL(i) * GKXSS(i) '统计全部供冷量
                                ElseIf PJGLGL_F + XFGLGL_F <= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = Math.Round(PJGLGL_F + XFGLGL_F, 2)
                                    GLZL = GLZL + PJGLGL_F * GKXSS(i) '统计全部供冷量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else '冷负荷需求量介于两者之间
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XFGLGL_F + LFHZXQL(i) - XFGLGL_F
                                    GLZL = GLZL + (LFHZXQL(i) - XFGLGL_F) * GKXSS(i) '统计全部供冷量
                                End If
                                ZLSL_F = ZLSL_F - GKXSS(i) '剩余的峰段小时数
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄冷量还有剩余，则继续用于峰段供冷
                    For i = GKXHmin To GKXHmax '第四步，计算峰段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If XLZL - GLZL > 0 And LFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value Then
                                '记录目前已经有的蓄冷供冷量
                                Dim XLGLGL_now_F As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value
                                If (LFHZXQL(i) - XLGLGL_now_F) >= (XLZL - GLZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_F + (XLZL - GLZL) / GKXSS(i)
                                    GLZL = GLZL + (XLZL - GLZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_F + LFHZXQL(i) - XLGLGL_now_F
                                    GLZL = GLZL + (LFHZXQL(i) - XLGLGL_now_F) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '如果蓄冷量还有剩余，用于平段供冷
                    Dim ZLSL_P As Double = 0
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If LFHZXQL(i) > 0 Then
                                ZLSL_P = ZLSL_P + GKXSS(i) '平段总小时数
                            End If
                        End If
                    Next
                    Dim PJGLGL_P As Double '平段每小时平均供冷功率
                    PJGLGL_P = (XLZL - GLZL) / ZLSL_P '剩余的几个小时，平均每个小时可以供冷的功率
                    For i = GKXHmin To GKXHmax '第五步，计算平段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If LFHZXQL(i) > 0 And PJGLGL_P > 0 Then '冷负荷大于0才进行计算,平均功率大于0时，才进行计算
                                Dim XFGLGL_P As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value '削峰供冷功率
                                If PJGLGL_P >= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = LFHZXQL(i)
                                    GLZL = GLZL + LFHZXQL(i) * GKXSS(i) '统计全部供冷量
                                ElseIf PJGLGL_P + XFGLGL_P <= LFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = Math.Round(PJGLGL_P + XFGLGL_P, 2)
                                    GLZL = GLZL + PJGLGL_P * GKXSS(i) '统计全部供冷量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else '冷负荷需求量介于两者之间
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XFGLGL_P + LFHZXQL(i) - XFGLGL_P
                                    GLZL = GLZL + (LFHZXQL(i) - XFGLGL_P) * GKXSS(i) '统计全部供冷量
                                End If
                                ZLSL_P = ZLSL_P - GKXSS(i) '剩余的峰段小时数
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄冷量还有剩余，则继续用于平段供冷
                    For i = GKXHmin To GKXHmax '第四步，计算平段的供冷功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If XLZL - GLZL > 0 And LFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value Then
                                '记录目前已经有的蓄冷供冷量
                                Dim XLGLGL_now_P As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value
                                If (LFHZXQL(i) - XLGLGL_now_P) >= (XLZL - GLZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_P + (XLZL - GLZL) / GKXSS(i)
                                    GLZL = GLZL + (XLZL - GLZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = XLGLGL_now_P + LFHZXQL(i) - XLGLGL_now_P
                                    GLZL = GLZL + (LFHZXQL(i) - XLGLGL_now_P) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄冷装置供冷功率是否大于输入的最大蓄冷功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > ZDXLGL Then
                                '将这一条工况已经供冷的量从总供冷量中减去（减去已经写入的所有负荷）
                                GLZL = GLZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                                '供冷功率设置成最大蓄冷功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = ZDXLGL
                                '将新的供冷功率加到供冷总量中（加上已经写入的所有负荷）
                                GLZL = GLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '第三步，根据蓄能装置供冷总能量，反向修正蓄能装置蓄冷功率
                    '如果蓄冷和供冷能量存在较大偏差（大于4%）
                    If Math.Abs((XLZL - GLZL) / XLZL) > 0.04 Then
                        '只会出现蓄冷比供冷大的情况，所以修改蓄冷的值
                        '根据供冷总量，求蓄冷平均功率
                        Dim XLGL_PJ_a As Double = GLZL / GD1_GD2_QT1_QT2_XSS_L
                        For i = GKXHmin To GKXHmax
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > XLGL_PJ_a Then
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = XLGL_PJ_a
                            End If
                        Next
                    End If
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '第四步，验算
                    '重新计算蓄能装置的蓄冷和供冷总能量
                    Dim XNGLZL As Double = 0 '蓄能装置供冷总量
                    Dim XNXLZL As Double = 0 '蓄能装置蓄冷总量
                    For i = GKXHmin To GKXHmax
                        XNGLZL = XNGLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value * GKXSS(i)
                        XNXLZL = XNXLZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value * GKXSS(i)
                    Next
                    If Math.Abs((XNXLZL - XNGLZL) / XNXLZL) > 0.05 Then '如果蓄冷总量和供冷总量误差超过5%，报错
                        MsgBox("蓄冷装置蓄冷总量与蓄冷装置供冷总量之间的误差超过了5%，请检查！！")
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————————————————
                '判断这一组日期序号代表制热还是制热
                If RQXH(RQXHJS(a)) = 3 Or RQXH(RQXHJS(a)) = 4 Then '制热工况；RXHCJS(a):日期序号发生变化前的最后一个工况序号；RQXH()：对应的日期序号
                    '求可以蓄热的最大功率
                    '如果勾选了蓄热功率存在最大值（模式=1）
                    Dim ZDXRGL As Double = 0 '最大蓄热功率
                    Dim RFH As New List(Of Double) '输入的每一条热负荷
                    '如果勾选了蓄热功率存在最大值（模式=1）
                    If XRJS_MS = 1 Then
                        ZDXRGL = XRGL_MAX '最大蓄热功率
                    Else
                        '遍历所有的热负荷总需求量
                        For i = GKXHmin To GKXHmax
                            RFH.Add(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 21).Value)
                        Next
                        '等于制热装机功率和热负荷最大值中的大值
                        ZDXRGL = Math.Max(ZRZJGL - XHLZRa, RFH.Max) + 10 '最大值加10，放大一点，防止出错
                    End If
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    Dim SYXRGL As Double = 0 '剩余蓄热功率
                    Dim XRZL As Double = 0 '蓄热总量
                    Dim GRZL As Double = 0 '供热总量
                    '先计算所有的（谷1、谷2、其它1、其它2）时间段最多可以蓄热多少kWh
                    Dim GD1_GD2_QT1_QT2_XRZL As Double = 0 '（谷1、谷2、其它1、其它2）蓄热总量
                    Dim GD1_GD2_QT1_QT2_XSS_R As Double = 0 '热工况（谷1、谷2、其它1、其它2）小时数
                    For i = GKXHmin To GKXHmax '遍历所有的热负荷段
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            GD1_GD2_QT1_QT2_XRZL = GD1_GD2_QT1_QT2_XRZL + GKXSS(i) * PJXRGL '所有为（谷1、谷2、其它1、其它2）时间段的小时数乘以输入的最大蓄热功率
                            GD1_GD2_QT1_QT2_XSS_R = GD1_GD2_QT1_QT2_XSS_R + GKXSS(i) '热负荷段的（谷电、时间段1、时间段2）总小时数
                        End If
                    Next
                    SYXRGL = SYXRGL + GD1_GD2_QT1_QT2_XRZL '剩余蓄热功率
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄热和蓄热
                        '将用电负荷段为（谷1、谷2、其它1、其它2）的工况，蓄热功率设置为输入的PJXRGL
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            If RFHZXQL(i) + PJXRGL > ZRZJGL - XHLZRa And SYXRGL > 0 Then '寻找蓄热功率+制热功率大于设备总制热功率的情况（谷1、谷2、其它1、其它2时内燃机关闭，需要减去溴化锂制热量）
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = Math.Round((ZRZJGL - XHLZRa - RFHZXQL(i)), 2) '（谷1、谷2、其它1、其它2）时内燃机关闭，需要减去溴化锂制热量，同时再缩小一点，以免出错
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > 0 Then
                                    XRZL = XRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                End If
                            ElseIf RFHZXQL(i) + PJXRGL <= ZRZJGL - XHLZRa And SYXRGL > 0 Then
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = PJXRGL '等于平均蓄热功率
                                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > 0 Then
                                    XRZL = XRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                End If
                            End If
                        End If
                    Next
                    '如果还有剩余的蓄热功率，统计现在有多少个工况点设备装机量没有被完全利用
                    Dim GD1_GD2_QT1_QT2_XSS_R_2 As Double = 0
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄热和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XRGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value '目前已经有的蓄热功率
                            If RFHZXQL(i) + XRGL_now < ZRZJGL - XHLZRa And SYXRGL > 0 And XRGL_now = PJXRGL Then '如果目前的蓄热功率不等于平均蓄热功率
                                GD1_GD2_QT1_QT2_XSS_R_2 = GD1_GD2_QT1_QT2_XSS_R_2 + GKXSS(i)
                            End If
                        End If
                    Next
                    Dim GD1_GD2_QT1_QT2_XRGL_R_2 As Double = SYXRGL / GD1_GD2_QT1_QT2_XSS_R_2 '平均蓄热功率2
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄热和蓄热
                        '将用电负荷段为（谷1、谷2、其它1、其它2）的工况，蓄热功率设置为输入的GD1_GD2_QT1_QT2_XRGL_R_2，同时也要满足装机需求（蓄热功率+制热功率不可以大于设备总制热功率）
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XRGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value '目前已经有的蓄热功率
                            If SYXRGL > 0 And RFHZXQL(i) + XRGL_now < ZRZJGL - XHLZRa And XRGL_now = PJXRGL Then '如果剩余蓄热功率大于0，且装机功率没有被完全利用
                                If ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now)) >= GD1_GD2_QT1_QT2_XRGL_R_2 Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = XRGL_now + GD1_GD2_QT1_QT2_XRGL_R_2
                                    XRZL = XRZL + GD1_GD2_QT1_QT2_XRGL_R_2 * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = XRGL_now + ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now))
                                    XRZL = XRZL + ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now)) * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                End If
                            End If
                        End If
                    Next
                    '如果此时还有剩余没使用的蓄热功率，则继续蓄热
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄热和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XRGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value '目前已经有的蓄热功率
                            If SYXRGL > 0 And RFHZXQL(i) + XRGL_now <= ZRZJGL - XHLZRa Then
                                If ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now)) >= SYXRGL / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = XRGL_now + SYXRGL / GKXSS(i)
                                    XRZL = XRZL + SYXRGL / GKXSS(i) * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = XRGL_now + ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now))
                                    XRZL = XRZL + ((ZRZJGL - XHLZRa) - (RFHZXQL(i) + XRGL_now)) * GKXSS(i) '统计全部蓄热量
                                    SYXRGL = GD1_GD2_QT1_QT2_XRZL - XRZL '剩下还需要满足的蓄热量
                                End If
                            End If
                        End If
                    Next
                    '判断计算出的蓄热功率是否大于输入的最大蓄热功率
                    For i = GKXHmin To GKXHmax '第一步，先计算蓄热和蓄热
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "谷2" Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它1" And QT1_SJD = True) Or (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "其它2" And QT2_SJD = True) Then
                            Dim XRGL_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value '目前已经有的蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > ZDXRGL Then
                                '蓄热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = ZDXRGL
                                '重新计算蓄热总量
                                XRZL = XRZL + (ZDXRGL - XRGL_now) * GKXSS(i) '统计全部蓄热量
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    For i = GKXHmin To GKXHmax '第二步，计算进行削峰需要的供热功率
                        '即使是（谷1、谷2、其它1、其它2），如果冷负荷大于了装机，也要进行削峰
                        If RFHZXQL(i) > ZRZJGL Then '热负荷总需求量大于制热装机量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = RFHZXQL(i) - ZRZJGL
                            GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i) '统计全部供热量
                        End If
                        '判断当前削峰功率是否大于输入的最大热负荷，如果是，则报错，但是不修改数值，使得计算可以继续进行
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                            MsgBox("蓄热装置供热功率大于输入的最大蓄热功率，装机方案选择不合理，程序不会自动修改计算出的数值，但请检查并重新选择装机方案！！" & "装机方案不合理的工况序号为： " & i)
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '统计有多少个高峰段
                    Dim ZRSL_GF As Double = 0 '制热时间段高峰段小时数
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If RFHZXQL(i) > 0 Then
                                ZRSL_GF = ZRSL_GF + GKXSS(i)
                            End If
                        End If
                    Next
                    Dim PJGRGL_GF As Double '平均供热功率
                    PJGRGL_GF = (XRZL - GRZL) / ZRSL_GF '剩余的几个工况，平均每个工况可以供热的功率
                    For i = GKXHmin To GKXHmax '第三步，计算高峰段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If RFHZXQL(i) > 0 And PJGRGL_GF > 0 Then
                                Dim XFGRGL_GF As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value '削峰供热功率
                                If PJGRGL_GF >= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = RFHZXQL(i)
                                    GRZL = GRZL + RFHZXQL(i) * GKXSS(i) '统计全部供热量
                                ElseIf PJGRGL_GF + XFGRGL_GF <= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = Math.Round(PJGRGL_GF + XFGRGL_GF, 2)
                                    GRZL = GRZL + PJGRGL_GF * GKXSS(i) '统计全部供热量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XFGRGL_GF + RFHZXQL(i) - XFGRGL_GF
                                    GRZL = GRZL + (RFHZXQL(i) - XFGRGL_GF) * GKXSS(i) '统计全部供热量
                                End If
                                ZRSL_GF = ZRSL_GF - GKXSS(i)
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄热量还有剩余，则继续用于高峰段供热
                    For i = GKXHmin To GKXHmax '第三步，计算平段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "高峰2" Then
                            If XRZL - GRZL > 0 And RFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value Then
                                '记录目前已经有的蓄热供热量
                                Dim XRGRGL_now_GF As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value
                                If (RFHZXQL(i) - XRGRGL_now_GF) >= (XRZL - GRZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_GF + (XRZL - GRZL) / GKXSS(i)
                                    GRZL = GRZL + (XRZL - GRZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_GF + RFHZXQL(i) - XRGRGL_now_GF
                                    GRZL = GRZL + (RFHZXQL(i) - XRGRGL_now_GF) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '统计有多少个峰段
                    Dim ZRSL_F As Double = 0 '制热时间段峰段小时数
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If RFHZXQL(i) > 0 Then
                                ZRSL_F = ZRSL_F + GKXSS(i)
                            End If
                        End If
                    Next
                    Dim PJGRGL_F As Double '平均供热功率
                    PJGRGL_F = (XRZL - GRZL) / ZRSL_F '剩余的几个工况，平均每个工况可以供热的功率
                    For i = GKXHmin To GKXHmax '第四步，计算峰段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If RFHZXQL(i) > 0 And PJGRGL_F > 0 Then
                                Dim XFGRGL_F As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value '削峰供热功率
                                If PJGRGL_F >= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = RFHZXQL(i)
                                    GRZL = GRZL + RFHZXQL(i) * GKXSS(i) '统计全部供热量
                                ElseIf PJGRGL_F + XFGRGL_F <= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = Math.Round(PJGRGL_F + XFGRGL_F, 2)
                                    GRZL = GRZL + PJGRGL_F * GKXSS(i) '统计全部供热量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XFGRGL_F + RFHZXQL(i) - XFGRGL_F
                                    GRZL = GRZL + (RFHZXQL(i) - XFGRGL_F) * GKXSS(i) '统计全部供热量
                                End If
                                ZRSL_F = ZRSL_F - GKXSS(i)
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄热量还有剩余，则继续用于峰段供热
                    For i = GKXHmin To GKXHmax '第四步，计算平段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "峰2" Then
                            If XRZL - GRZL > 0 And RFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value Then
                                '记录目前已经有的蓄热供热量
                                Dim XRGRGL_now_F As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value
                                If (RFHZXQL(i) - XRGRGL_now_F) >= (XRZL - GRZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_F + (XRZL - GRZL) / GKXSS(i)
                                    GRZL = GRZL + (XRZL - GRZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_F + RFHZXQL(i) - XRGRGL_now_F
                                    GRZL = GRZL + (RFHZXQL(i) - XRGRGL_now_F) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '如果蓄热量还有剩余，用于平段蓄热
                    Dim ZRSL_P As Double = 0 '制热时间段平段小时数
                    For i = GKXHmin To GKXHmax
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If RFHZXQL(i) > 0 Then
                                ZRSL_P = ZRSL_P + GKXSS(i)
                            End If
                        End If
                    Next
                    Dim PJGRGL_P As Double '平均供热功率
                    PJGRGL_P = (XRZL - GRZL) / ZRSL_P '剩余的几个工况，平均每个工况可以供热的功率
                    For i = GKXHmin To GKXHmax '第四步，计算平段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If RFHZXQL(i) > 0 And PJGRGL_P > 0 Then
                                Dim XFGRGL_P As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value '削平供热功率
                                If PJGRGL_P >= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = RFHZXQL(i)
                                    GRZL = GRZL + RFHZXQL(i) * GKXSS(i) '统计全部供热量
                                ElseIf PJGRGL_P + XFGRGL_P <= RFHZXQL(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = Math.Round(PJGRGL_P + XFGRGL_P, 2)
                                    GRZL = GRZL + PJGRGL_P * GKXSS(i) '统计全部供热量，仅仅加上新增的部分（不含第二步计算的削峰量）
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XFGRGL_P + RFHZXQL(i) - XFGRGL_P
                                    GRZL = GRZL + (RFHZXQL(i) - XFGRGL_P) * GKXSS(i) '统计全部供热量
                                End If
                                ZRSL_P = ZRSL_P - GKXSS(i)
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '如果此时蓄热量还有剩余，则继续用于平段供热
                    For i = GKXHmin To GKXHmax '第四步，计算平段的供热功率
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平1" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 79).Value = "平2" Then
                            If XRZL - GRZL > 0 And RFHZXQL(i) > ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value Then
                                '记录目前已经有的蓄热供热量
                                Dim XRGRGL_now_P As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value
                                If (RFHZXQL(i) - XRGRGL_now_P) >= (XRZL - GRZL) / GKXSS(i) Then
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_P + (XRZL - GRZL) / GKXSS(i)
                                    GRZL = GRZL + (XRZL - GRZL) / GKXSS(i) * GKXSS(i)
                                Else
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = XRGRGL_now_P + RFHZXQL(i) - XRGRGL_now_P
                                    GRZL = GRZL + (RFHZXQL(i) - XRGRGL_now_P) * GKXSS(i)
                                End If
                            End If
                            '判断计算出的蓄热装置供热功率是否大于输入的最大蓄热功率
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > ZDXRGL Then
                                '将这一条工况已经供热的量从总供热量中减去（减去已经写入的所有负荷）
                                GRZL = GRZL - ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                                '供热功率设置成最大蓄热功率
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = ZDXRGL
                                '将新的供热功率加到供热总量中（加上已经写入的所有负荷）
                                GRZL = GRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                            End If
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '第三步，根据蓄能装置供热总能量，反向修正蓄能装置蓄热功率
                    '如果蓄热和供热能量存在较大偏差（大于4%）
                    If Math.Abs((XRZL - GRZL) / XRZL) > 0.04 Then
                        '只会出现蓄热比供热大的情况，所以修改蓄热的值
                        '根据供热总量，求蓄热平均功率
                        Dim XRGL_PJ_a As Double = GRZL / GD1_GD2_QT1_QT2_XSS_R
                        For i = GKXHmin To GKXHmax
                            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > XRGL_PJ_a Then
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = XRGL_PJ_a
                            End If
                        Next
                    End If
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    '第四步，验算
                    '重新计算蓄能装置的蓄热和供热总能量
                    Dim XNGRZL As Double = 0 '蓄能装置供热总量
                    Dim XNXRZL As Double = 0 '蓄能装置蓄热总量
                    For i = GKXHmin To GKXHmax
                        XNGRZL = XNGRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value * GKXSS(i)
                        XNXRZL = XNXRZL + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value * GKXSS(i)
                    Next
                    If Math.Abs((XNXRZL - XNGRZL) / XNXRZL) > 0.05 Then '如果蓄热总量和供热总量误差超过5%，报错
                        MsgBox("蓄热装置蓄热总量与蓄热装置供热总量之间的误差超过了5%，请检查！！")
                    End If
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '清除计算出的明显很小的不合理的值
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value < 1 Then '供冷
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = Nothing
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value < 1 Then '蓄冷
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = Nothing
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value < 1 Then '供热
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = Nothing
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value < 1 Then '蓄热
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = Nothing
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '计算出来的值只保留最大3位小数
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value > 0 Then
                    Dim Temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 13).Value = Math.Round(Temp, 3)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > 0 Then
                    Dim Temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value = Math.Round(Temp, 3)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value > 0 Then
                    Dim Temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 22).Value = Math.Round(Temp, 3)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > 0 Then
                    Dim Temp As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value = Math.Round(Temp, 3)
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '检查计算结果，（谷1、谷2、其它1、其它2）时间段蓄冷功率不可以大于蓄冷装机总功率（蓄冷只用电设备）
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 14).Value > XLZJGL Then
                    MsgBox("计算出的逐工况蓄冷功率中，存在大于蓄冷装机总功率的情况（蓄冷时只是用耗电设备），请检查！")
                    Exit For
                End If
            Next
            '检查计算结果，（谷1、谷2、其它1、其它2）时间段蓄热功率不可以大于蓄冷装机总功率（蓄热只用电设备）
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 23).Value > XRZJGL Then
                    MsgBox("计算出的逐工况蓄热功率中，存在大于蓄热装机总功率的情况（蓄热时只是用耗电设备），请检查！")
                    Exit For
                End If
            Next
        End If
        Call 锁定工作表(ExcelApp)
    End Sub

End Module
