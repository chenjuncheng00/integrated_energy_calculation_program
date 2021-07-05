Module 采暖季全局寻优计算
    Sub 制热季设备制热和蓄热全局寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, XHLZR As Double, XHLXR As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        '热负荷总量=供热+蓄热
        Dim RFH_GR_now As Double = RFHZXQL(b) - (XNGRGL(b) + XHLZR)
        Dim RFH_XR_now As Double = XNXRGL(b) - XHLXR
        '如果当前热负荷总需求量=0，则将RFH_GR ,强制设置为0，防止出错
        If RFHZXQL(b) = 0 Then
            RFH_GR_now = 0
        End If
        '如果当前蓄热负荷总需求量=0，则将RFH_XR ,强制设置为0，防止出错
        If XNXRGL(b) = 0 Then
            RFH_XR_now = 0
        End If
        Dim RFH_ALL As Double = RFH_GR_now + RFH_XR_now
        '负荷调整系数（全局寻优时候能否计算到了负荷上限的倍数）
        Dim TZXS As Double = 1 + 10 * FHTJJD / 100
        '误差系数，允许误差的最大比例
        Dim WCXS As Double = 1 + 4 * FHTJJD / 100
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '根据选择的设备类型，获取被选择的设备类型，由于采用全局寻优计算方法，不需要考虑设备选择顺序
        '设置变量判断某一种设备是否被启用，0表示没有启用，1表示启用
        '天然气锅炉
        Dim ZJJC_TRQGL As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "天然气锅炉" Then
                ZJJC_TRQGL = 1
                Exit For
            End If
        Next
        '电采暖锅炉
        Dim ZJJC_DGL As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "电采暖锅炉" Then
                ZJJC_DGL = 1
                Exit For
            End If
        Next
        '风冷螺杆机
        Dim ZJJC_FLLGJ As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "风冷螺杆机" Then
                ZJJC_FLLGJ = 1
                Exit For
            End If
        Next
        '水（地）源热泵
        Dim ZJJC_SDYRB As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "水(地)源热泵" Then
                ZJJC_SDYRB = 1
                Exit For
            End If
        Next
        '离心式热泵
        Dim ZJJC_LXSRB As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "离心式热泵" Then
                ZJJC_LXSRB = 1
                Exit For
            End If
        Next
        '空气源热泵
        Dim ZJJC_KQYRB As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "空气源热泵" Then
                ZJJC_KQYRB = 1
                Exit For
            End If
        Next
        '直燃型溴化锂
        Dim ZJJC_ZRXXHL As Integer = 0
        '遍历顺序1到6
        For i = 15 To 20
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "直燃型溴化锂" Then
                ZJJC_ZRXXHL = 1
                Exit For
            End If
        Next
        '装机监测情况汇总
        Dim ZJJC_ALL As Integer = ZJJC_FLLGJ + ZJJC_KQYRB + ZJJC_TRQGL + ZJJC_LXSRB + ZJJC_SDYRB + ZJJC_DGL + ZJJC_ZRXXHL
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '根据选择的计算模式类型，只有是模式2的时候才会计算
        If calculation_mode = 2 And RFH_GR_now + RFH_XR_now > 0 And ZJJC_ALL > 0 Then
            '读取采暖季装机方案及参数
            Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
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
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '混水供热设备装机数量
            Dim NUM1_HS As Double = 0
            Dim NUM2_HS As Double = 0
            '混水设备（1）（2）装机功率
            Dim ZJRGL1_HS As Double = 0
            Dim ZJRGL2_HS As Double = 0
            '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
            Dim ZJRGL_HS_ALL As Double = 0
            '参与混水供热的设备的本体耗电功率和辅机设备耗电功率
            Dim BTHD1_ED_HS As Double = 0
            Dim BTHD2_ED_HS As Double = 0
            Dim FJHD1_ED_HS As Double = 0
            Dim FJHD2_ED_HS As Double = 0
            '混水设备负荷率下限（单台）
            Dim FHL1_min_HS As Double
            Dim FHL2_min_HS As Double
            '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                '装机制热功率
                ZJRGL_HS_ALL = ZJRGL1_KQYRB + ZJRGL2_KQYRB
                ZJRGL1_HS = ZJRGL1_KQYRB
                ZJRGL2_HS = ZJRGL2_KQYRB
                '设备（1）装机数量
                NUM1_HS = NUM1_KQYRB
                '设备（2）装机数量
                NUM2_HS = NUM2_KQYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_KQYRB
                BTHD2_ED_HS = BTHD2_ED_KQYRB
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_HS = FJHD1_ED_KQYRB
                FJHD2_ED_HS = FJHD2_ED_KQYRB
                '混水设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_HS = FHL1_min_KQYRB
                FHL2_min_HS = FHL2_min_KQYRB
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                '制热功率
                ZJRGL_HS_ALL = ZJRGL1_SDYRB + ZJRGL2_SDYRB
                ZJRGL1_HS = ZJRGL1_SDYRB
                ZJRGL2_HS = ZJRGL2_SDYRB
                '设备（1）装机数量
                NUM1_HS = NUM1_SDYRB
                '设备（2）装机数量
                NUM2_HS = NUM2_SDYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_SDYRB
                BTHD2_ED_HS = BTHD2_ED_SDYRB
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_HS = FJHD1_ED_SDYRB
                FJHD2_ED_HS = FJHD2_ED_SDYRB
                '混水设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_HS = FHL1_min_SDYRB
                FHL2_min_HS = FHL2_min_SDYRB
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                '制热功率
                ZJRGL_HS_ALL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ
                ZJRGL1_HS = ZJRGL1_FLLGJ
                ZJRGL2_HS = ZJRGL2_FLLGJ
                '设备（1）装机数量
                NUM1_HS = NUM1_FLLGJ
                '设备（2）装机数量
                NUM2_HS = NUM2_FLLGJ
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_HS = BTHD1_ED_FLLGJ
                BTHD2_ED_HS = BTHD2_ED_FLLGJ
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_HS = FJHD1_ED_FLLGJ
                FJHD2_ED_HS = FJHD2_ED_FLLGJ
                '混水设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_HS = FHL1_min_FLLGJ
                FHL2_min_HS = FHL2_min_FLLGJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '梯级供热设备装机数量
            Dim NUM1_TJ As Double = 0
            Dim NUM2_TJ As Double = 0
            '梯级设备（1）（2）装机功率
            Dim ZJRGL1_TJ As Double = 0
            Dim ZJRGL2_TJ As Double = 0
            '参与梯级的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
            Dim ZJRGL_TJ_ALL As Double = 0
            '参与梯级供热的设备的本体耗电功率和辅机设备耗电功率
            Dim BTHD1_ED_TJ As Double = 0
            Dim BTHD2_ED_TJ As Double = 0
            Dim FJHD1_ED_TJ As Double = 0
            Dim FJHD2_ED_TJ As Double = 0
            '梯级设备负荷率下限（单台）
            Dim FHL1_min_TJ As Double
            Dim FHL2_min_TJ As Double
            '梯级设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为梯级设备的装机总功率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "空气源热泵" Then
                '装机制热功率
                ZJRGL_TJ_ALL = ZJRGL1_KQYRB + ZJRGL2_KQYRB
                ZJRGL1_TJ = ZJRGL1_KQYRB
                ZJRGL2_TJ = ZJRGL2_KQYRB
                '设备（1）装机数量
                NUM1_TJ = NUM1_KQYRB
                '设备（2）装机数量
                NUM2_TJ = NUM2_KQYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_KQYRB
                BTHD2_ED_TJ = BTHD2_ED_KQYRB
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_TJ = FJHD1_ED_KQYRB
                FJHD2_ED_TJ = FJHD2_ED_KQYRB
                '梯级设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_TJ = FHL1_min_KQYRB
                FHL2_min_TJ = FHL2_min_KQYRB
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "水(地)源热泵" Then
                '制热功率
                ZJRGL_TJ_ALL = ZJRGL1_SDYRB + ZJRGL2_SDYRB
                ZJRGL1_TJ = ZJRGL1_SDYRB
                ZJRGL2_TJ = ZJRGL2_SDYRB
                '设备（1）装机数量
                NUM1_TJ = NUM1_SDYRB
                '设备（2）装机数量
                NUM2_TJ = NUM2_SDYRB
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_SDYRB
                BTHD2_ED_TJ = BTHD2_ED_SDYRB
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_TJ = FJHD1_ED_SDYRB
                FJHD2_ED_TJ = FJHD2_ED_SDYRB
                '梯级设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_TJ = FHL1_min_SDYRB
                FHL2_min_TJ = FHL2_min_SDYRB
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value = "风冷螺杆机" Then
                '制热功率
                ZJRGL_TJ_ALL = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ
                ZJRGL1_TJ = ZJRGL1_FLLGJ
                ZJRGL2_TJ = ZJRGL2_FLLGJ
                '设备（1）装机数量
                NUM1_TJ = NUM1_FLLGJ
                '设备（2）装机数量
                NUM2_TJ = NUM2_FLLGJ
                '100%负荷时本体耗电功率（单台）
                BTHD1_ED_TJ = BTHD1_ED_FLLGJ
                BTHD2_ED_TJ = BTHD2_ED_FLLGJ
                '100%负荷时辅机耗电功率（单台）
                FJHD1_ED_TJ = FJHD1_ED_FLLGJ
                FJHD2_ED_TJ = FJHD2_ED_FLLGJ
                '梯级设备负荷率下限（单台，没有除以设备数量）
                FHL1_min_TJ = FHL1_min_FLLGJ
                FHL2_min_TJ = FHL2_min_FLLGJ
            End If
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '各个设备可以装机功率（总和）
            '天然气锅炉
            '装机总功率（总和）
            Dim ZJZGL_TRQGL As Double = 0
            If ZJJC_TRQGL = 1 Then
                ZJZGL_TRQGL = ZJRGL1_TRQGL + ZJRGL2_TRQGL
            Else
                ZJZGL_TRQGL = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_TRQGL As Double = 0
            Dim FH1_min_TRQGL As Double = 0
            Dim FH2_min_TRQGL As Double = 0
            If ZJJC_TRQGL = 1 Then
                If NUM1_TRQGL > 0 Then
                    FH1_min_TRQGL = ZJRGL1_TRQGL * FHL1_min_TRQGL / NUM1_TRQGL
                Else
                    FH1_min_TRQGL = 0
                End If
                If NUM2_TRQGL > 0 Then
                    FH2_min_TRQGL = ZJRGL2_TRQGL * FHL2_min_TRQGL / NUM2_TRQGL
                Else
                    FH2_min_TRQGL = 0
                End If
                FH_min_TRQGL = FH1_min_TRQGL + FH2_min_TRQGL
            Else
                FH_min_TRQGL = 0
            End If
            '电采暖锅炉
            '装机总功率（总和）
            Dim ZJZGL_DGL As Double = 0
            If ZJJC_DGL = 1 Then
                ZJZGL_DGL = ZJRGL1_DGL + ZJRGL2_DGL
            Else
                ZJZGL_DGL = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_DGL As Double = 0
            Dim FH1_min_DGL As Double = 0
            Dim FH2_min_DGL As Double = 0
            If ZJJC_DGL = 1 Then
                If NUM1_DGL > 0 Then
                    FH1_min_DGL = ZJRGL1_DGL * FHL1_min_DGL / NUM1_DGL
                Else
                    FH1_min_DGL = 0
                End If
                If NUM2_DGL > 0 Then
                    FH2_min_DGL = ZJRGL2_DGL * FHL2_min_DGL / NUM2_DGL
                Else
                    FH2_min_DGL = 0
                End If
                FH_min_DGL = FH1_min_DGL + FH2_min_DGL
            Else
                FH_min_DGL = 0
            End If
            '水（地）源热泵
            '装机总功率（总和）
            Dim ZJZGL_SDYRB As Double = 0
            If ZJJC_SDYRB = 1 Then
                ZJZGL_SDYRB = ZJRGL1_SDYRB + ZJRGL2_SDYRB
            Else
                ZJZGL_SDYRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_SDYRB As Double = 0
            Dim FH1_min_SDYRB As Double = 0
            Dim FH2_min_SDYRB As Double = 0
            If ZJJC_SDYRB = 1 Then
                If NUM1_SDYRB > 0 Then
                    FH1_min_SDYRB = ZJRGL1_SDYRB * FHL1_min_SDYRB / NUM1_SDYRB
                Else
                    FH1_min_SDYRB = 0
                End If
                If NUM2_SDYRB > 0 Then
                    FH2_min_SDYRB = ZJRGL2_SDYRB * FHL2_min_SDYRB / NUM2_SDYRB
                Else
                    FH2_min_SDYRB = 0
                End If
                FH_min_SDYRB = FH1_min_SDYRB + FH2_min_SDYRB
            Else
                FH_min_SDYRB = 0
            End If
            '离心式热泵
            '装机总功率（总和）
            Dim ZJZGL_LXSRB As Double = 0
            If ZJJC_LXSRB = 1 Then
                ZJZGL_LXSRB = ZJRGL1_LXSRB + ZJRGL2_LXSRB
            Else
                ZJZGL_LXSRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_LXSRB As Double = 0
            Dim FH1_min_LXSRB As Double = 0
            Dim FH2_min_LXSRB As Double = 0
            If ZJJC_LXSRB = 1 Then
                If NUM1_LXSRB > 0 Then
                    FH1_min_LXSRB = ZJRGL1_LXSRB * FHL1_min_LXSRB / NUM1_LXSRB
                Else
                    FH1_min_LXSRB = 0
                End If
                If NUM2_LXSRB > 0 Then
                    FH2_min_LXSRB = ZJRGL2_LXSRB * FHL2_min_LXSRB / NUM2_LXSRB
                Else
                    FH2_min_LXSRB = 0
                End If
                FH_min_LXSRB = FH1_min_LXSRB + FH2_min_LXSRB
            Else
                FH_min_LXSRB = 0
            End If
            '风冷螺杆机
            '装机总功率（总和）
            Dim ZJZGL_FLLGJ As Double = 0
            If ZJJC_FLLGJ = 1 Then
                ZJZGL_FLLGJ = ZJRGL1_FLLGJ + ZJRGL2_FLLGJ
            Else
                ZJZGL_FLLGJ = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_FLLGJ As Double = 0
            Dim FH1_min_FLLGJ As Double = 0
            Dim FH2_min_FLLGJ As Double = 0
            If ZJJC_FLLGJ = 1 Then
                If NUM1_FLLGJ > 0 Then
                    FH1_min_FLLGJ = ZJRGL1_FLLGJ * FHL1_min_FLLGJ / NUM1_FLLGJ
                Else
                    FH1_min_FLLGJ = 0
                End If
                If NUM2_FLLGJ > 0 Then
                    FH2_min_FLLGJ = ZJRGL2_FLLGJ * FHL2_min_FLLGJ / NUM2_FLLGJ
                Else
                    FH2_min_FLLGJ = 0
                End If
                FH_min_FLLGJ = FH1_min_FLLGJ + FH2_min_FLLGJ
            Else
                FH_min_FLLGJ = 0
            End If
            '空气源热泵
            '装机总功率（总和）
            Dim ZJZGL_KQYRB As Double = 0
            If ZJJC_KQYRB = 1 Then
                ZJZGL_KQYRB = ZJRGL1_KQYRB + ZJRGL2_KQYRB
            Else
                ZJZGL_KQYRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_KQYRB As Double = 0
            Dim FH1_min_KQYRB As Double = 0
            Dim FH2_min_KQYRB As Double = 0
            If ZJJC_KQYRB = 1 Then
                If NUM1_KQYRB > 0 Then
                    FH1_min_KQYRB = ZJRGL1_KQYRB * FHL1_min_KQYRB / NUM1_KQYRB
                Else
                    FH1_min_KQYRB = 0
                End If
                If NUM2_KQYRB > 0 Then
                    FH2_min_KQYRB = ZJRGL2_KQYRB * FHL2_min_KQYRB / NUM2_KQYRB
                Else
                    FH2_min_KQYRB = 0
                End If
                FH_min_KQYRB = FH1_min_KQYRB + FH2_min_KQYRB
            Else
                FH_min_KQYRB = 0
            End If
            '直燃型溴化锂
            '装机总功率（总和）
            Dim ZJZGL_ZRXXHL As Double = 0
            If ZJJC_ZRXXHL = 1 Then
                ZJZGL_ZRXXHL = ZJRGL1_ZRXXHL + ZJRGL2_ZRXXHL
            Else
                ZJZGL_ZRXXHL = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_ZRXXHL As Double = 0
            Dim FH1_min_ZRXXHL As Double = 0
            Dim FH2_min_ZRXXHL As Double = 0
            If ZJJC_ZRXXHL = 1 Then
                If NUM1_ZRXXHL > 0 Then
                    FH1_min_ZRXXHL = ZJRGL1_ZRXXHL * FHL1_min_ZRXXHL / NUM1_ZRXXHL
                Else
                    FH1_min_ZRXXHL = 0
                End If
                If NUM2_ZRXXHL > 0 Then
                    FH2_min_ZRXXHL = ZJRGL2_ZRXXHL * FHL2_min_ZRXXHL / NUM2_ZRXXHL
                Else
                    FH2_min_ZRXXHL = 0
                End If
                FH_min_ZRXXHL = FH1_min_ZRXXHL + FH2_min_ZRXXHL
            Else
                FH_min_ZRXXHL = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————   
            '如果此时制热装机总功率=供热功率+蓄热功率，装机功率会被全部用掉，则需要增加此时的计算容错性
            '溴化锂装机数量
            Dim NUM1_XHL As Double = ans_ZJFA_R(0)
            Dim NUM2_XHL As Double = ans_ZJFA_R(1)
            '内燃机余热功率
            Dim YRGL1_ED_NRJ As Double = ans_ZJFA_R(4)
            Dim YRGL2_ED_NRJ As Double = ans_ZJFA_R(5)
            '溴化锂制热COP
            Dim XHL_COP_R As Double = ans_XZXS_R(9)
            '计算溴化锂制热功率
            Dim ZJZGL_XHL As Double = (NUM1_XHL * YRGL1_ED_NRJ + NUM2_XHL * YRGL2_ED_NRJ) * XHL_COP_R
            '制热装机功率求和
            Dim ZJRGL_ALL As Double = 0
            '如果处于不启动内燃机的时间段
            If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(24, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(25, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(26, 9).Value) Then
                ZJRGL_ALL = ZJZGL_TRQGL + ZJZGL_DGL + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_ALL
            Else
                ZJRGL_ALL = ZJZGL_TRQGL + ZJZGL_DGL + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_ALL + ZJZGL_XHL
            End If
            '如果此时的热装机会被全部用掉，将工况序号记录下来，直接采用常规计算模式进行计算，不报错
            If Math.Abs(RFHZXQL(b) - XNGRGL(b) + XNXRGL(b) - ZJRGL_ALL) <= RCXS Then
                '将当前工况的序号加入列表
                QJXY_CGMS.Add(b)
                '直接结束本SUB
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————   
            '将总热负荷（供热+蓄热）分配给6个设备
            '负荷分配的次数，次数越多计算的越精细，但计算速度越慢
            Dim FHFPCS_start As Integer = 20 '初始值
            Dim FHFPCS_max As Integer = 100 '负荷分配次数参数的最大允许值
            Dim JS_start As Integer = 0 '已经计算过的次数计数，用于改变FHFPCS的值
zzzzz：
            Dim FHFPCS As Integer = FHFPCS_start + JS_start * 5
            '————————————————————————————————————————————————————————————————————————————————————————
            '如果FHFPCS太大，直接结束计算并报错
            If FHFPCS > FHFPCS_max Then
                '将当前工况的序号加入列表
                QJXYJS_ERROR.Add(b)
                'MsgBox("供热和蓄热全局寻优计算出错，计算结束！！" & “当前正在计算的工况序号为：   ” & b)
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '定义列表，储存各个设备计算出的结果
            '天然气锅炉
            Dim HQ_ALL_min_TRQGL As New List(Of Double)
            Dim HD_ALL_TRQGL As New List(Of Double)
            Dim GR_ALL_TRQGL As New List(Of Double)
            Dim XR_ALL_TRQGL As New List(Of Double)
            Dim FHL1_GR_TRQGL As New List(Of Double)
            Dim FHL1_XR_TRQGL As New List(Of Double)
            Dim FHL2_GR_TRQGL As New List(Of Double)
            Dim FHL2_XR_TRQGL As New List(Of Double)
            '电采暖锅炉计算出的各种结果存入列表
            Dim HD_ALL_min_DGL As New List(Of Double)
            Dim GR_ALL_DGL As New List(Of Double)
            Dim XR_ALL_DGL As New List(Of Double)
            Dim FHL1_GR_DGL As New List(Of Double)
            Dim FHL1_XR_DGL As New List(Of Double)
            Dim FHL2_GR_DGL As New List(Of Double)
            Dim FHL2_XR_DGL As New List(Of Double)
            '水（地）源热泵
            Dim HD_ALL_min_SDYRB As New List(Of Double)
            Dim GR_ALL_SDYRB As New List(Of Double)
            Dim XR_ALL_SDYRB As New List(Of Double)
            Dim FHL1_GR_SDYRB As New List(Of Double)
            Dim FHL1_XR_SDYRB As New List(Of Double)
            Dim FHL2_GR_SDYRB As New List(Of Double)
            Dim FHL2_XR_SDYRB As New List(Of Double)
            '离心式热泵
            Dim HD_ALL_min_LXSRB As New List(Of Double)
            Dim GR_ALL_LXSRB As New List(Of Double)
            Dim XR_ALL_LXSRB As New List(Of Double)
            Dim FHL1_GR_LXSRB As New List(Of Double)
            Dim FHL1_XR_LXSRB As New List(Of Double)
            Dim FHL2_GR_LXSRB As New List(Of Double)
            Dim FHL2_XR_LXSRB As New List(Of Double)
            '风冷螺杆机
            Dim HD_ALL_min_FLLGJ As New List(Of Double)
            Dim GR_ALL_FLLGJ As New List(Of Double)
            Dim XR_ALL_FLLGJ As New List(Of Double)
            Dim FHL1_GR_FLLGJ As New List(Of Double)
            Dim FHL1_XR_FLLGJ As New List(Of Double)
            Dim FHL2_GR_FLLGJ As New List(Of Double)
            Dim FHL2_XR_FLLGJ As New List(Of Double)
            '空气源热泵
            Dim HD_ALL_min_KQYRB As New List(Of Double)
            Dim GR_ALL_KQYRB As New List(Of Double)
            Dim XR_ALL_KQYRB As New List(Of Double)
            Dim FHL1_GR_KQYRB As New List(Of Double)
            Dim FHL1_XR_KQYRB As New List(Of Double)
            Dim FHL2_GR_KQYRB As New List(Of Double)
            Dim FHL2_XR_KQYRB As New List(Of Double)
            '直燃型溴化锂
            Dim HQ_ALL_min_ZRXXHL As New List(Of Double)
            Dim HD_ALL_ZRXXHL As New List(Of Double)
            Dim GR_ALL_ZRXXHL As New List(Of Double)
            Dim XR_ALL_ZRXXHL As New List(Of Double)
            Dim FHL1_GR_ZRXXHL As New List(Of Double)
            Dim FHL1_XR_ZRXXHL As New List(Of Double)
            Dim FHL2_GR_ZRXXHL As New List(Of Double)
            Dim FHL2_XR_ZRXXHL As New List(Of Double)
            '混水设备负荷率
            Dim FHL_HS_DGL As New List(Of Double)
            Dim FHL_HS_TRQGL As New List(Of Double)
            Dim FHL_HS_ZRXXHL As New List(Of Double)
            '梯级供热设备负荷率
            Dim FHL_TJGR_LXSRB As New List(Of Double)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '供热和蓄热一起寻优
            '混水设备供热
            Dim ZJRGL_HS_a_1 As Double
            Dim ZJRGL_HS_a_2 As Double
            'Dim ZJRGL_HS_a_7 As Double
            '寻优先后顺序：天然气锅炉、电锅炉、水（地）源热泵、离心式热泵、风冷螺杆机、空气源热泵、直燃型溴化锂
            '穷举计算
            '各种设备已经计算的次数计数（如果设备不存在或者装机量为0，才参与计算）
            '天然气锅炉计算结束监测
            Dim JS_end_TRQGL As Integer = 0
            '将总负荷分为供热功率和蓄热功率,进行计算
            For a_1 = 0 To ZJZGL_TRQGL Step ZJZGL_TRQGL / FHFPCS
                '直接结束计算
                If JS_end_TRQGL >= 1 Then
                    Exit For
                End If
                '如果不存在装机或者装机功率等于0
                If ZJJC_TRQGL = 0 Or ZJZGL_TRQGL = 0 Then
                    JS_end_TRQGL = JS_end_TRQGL + 1
                End If
                '混水供热设备（目的是可以正常进行内部循环进行计算）
                '用天然气锅炉的功率计算混水设备功率
                If ZJJC_TRQGL = 0 Or ZJZGL_TRQGL = 0 Then
                    ZJRGL_HS_a_1 = ZJRGL_HS_ALL
                Else
                    If ZJRGL_HS_ALL > 0 Then
                        ZJRGL_HS_a_1 = a_1 * HSGRGLBL / (1 - HSGRGLBL)
                    Else
                        ZJRGL_HS_a_1 = 0
                    End If
                End If
                '根据装机量判断是否直接进入下一次循环
                If a_1 + ZJZGL_DGL + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_a_1 < RFH_ALL Then
                    GoTo aaa
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                Dim JS_end_DGL As Integer = 0
                For a_2 = 0 To ZJZGL_DGL Step ZJZGL_DGL / FHFPCS
                    '直接结束计算
                    If JS_end_DGL >= 1 Then
                        Exit For
                    End If
                    '如果不存在装机或者装机功率等于0
                    If ZJJC_DGL = 0 Or ZJZGL_DGL = 0 Then
                        JS_end_DGL = JS_end_DGL + 1
                    End If
                    '混水供热设备（目的是可以正常进行内部循环进行计算）
                    '如果电锅炉不存在，此时仍然用天然气锅炉的功率计算混水设备功率
                    If ZJJC_DGL = 0 Or ZJZGL_DGL = 0 Then
                        '如果天然气锅炉也不存在，则混水设备会是直燃性溴化锂，直接等于全部的混水设备功率，进入下一个循环
                        If ZJJC_TRQGL = 0 Or ZJZGL_TRQGL = 0 Then
                            ZJRGL_HS_a_2 = ZJRGL_HS_ALL
                        Else
                            '天然气锅炉存在，此时用天然气锅炉功率计算混水设备功率
                            If ZJRGL_HS_ALL > 0 Then
                                ZJRGL_HS_a_2 = a_1 * HSGRGLBL / (1 - HSGRGLBL)
                            Else
                                ZJRGL_HS_a_2 = 0
                            End If
                        End If
                    Else
                        '如果电锅炉存在，则用电锅炉计算混水设备的功率
                        If ZJRGL_HS_ALL > 0 Then
                            ZJRGL_HS_a_2 = a_2 * HSGRGLBL / (1 - HSGRGLBL)
                        Else
                            ZJRGL_HS_a_2 = 0
                        End If
                    End If
                    '根据装机量判断是否直接进入下一次循环
                    If a_1 + a_2 + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_a_2 < RFH_ALL Then
                        GoTo bbb
                    End If
                    '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                    Dim JS_end_SDYRB As Integer = 0
                    For a_3 = 0 To ZJZGL_SDYRB Step ZJZGL_SDYRB / FHFPCS
                        '直接结束计算
                        If JS_end_SDYRB >= 1 Then
                            Exit For
                        End If
                        '如果不存在装机或者装机功率等于0
                        If ZJJC_SDYRB = 0 Or ZJZGL_SDYRB = 0 Then
                            JS_end_SDYRB = JS_end_SDYRB + 1
                        End If
                        '根据装机量判断是否直接进入下一次循环
                        If a_1 + a_2 + a_3 + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_a_2 < RFH_ALL Then
                            GoTo ccc
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                        Dim JS_end_LXSRB As Integer = 0
                        For a_4 = 0 To ZJZGL_LXSRB Step ZJZGL_LXSRB / FHFPCS
                            '直接结束计算
                            If JS_end_LXSRB >= 1 Then
                                Exit For
                            End If
                            '如果不存在装机或者装机功率等于0
                            If ZJJC_LXSRB = 0 Or ZJZGL_LXSRB = 0 Then
                                JS_end_LXSRB = JS_end_LXSRB + 1
                            End If
                            '根据装机量判断是否直接进入下一次循环
                            If a_1 + a_2 + a_3 + a_4 + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_a_2 < RFH_ALL Then
                                GoTo ddd
                            End If
                            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                            Dim JS_end_FLLGJ As Integer = 0
                            For a_5 = 0 To ZJZGL_FLLGJ Step ZJZGL_FLLGJ / FHFPCS
                                '直接结束计算
                                If JS_end_FLLGJ >= 1 Then
                                    Exit For
                                End If
                                '如果不存在装机或者装机功率等于0
                                If ZJJC_FLLGJ = 0 Or ZJZGL_FLLGJ = 0 Then
                                    JS_end_FLLGJ = JS_end_FLLGJ + 1
                                End If
                                '根据装机量判断是否直接进入下一次循环
                                If a_1 + a_2 + a_3 + a_4 + a_5 + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJRGL_HS_a_2 < RFH_ALL Then
                                    GoTo eee
                                End If
                                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                Dim JS_end_KQYRB As Integer = 0
                                For a_6 = 0 To ZJZGL_KQYRB Step ZJZGL_KQYRB / FHFPCS
                                    '直接结束计算
                                    If JS_end_KQYRB >= 1 Then
                                        Exit For
                                    End If
                                    '如果不存在装机或者装机功率等于0
                                    If ZJJC_KQYRB = 0 Or ZJZGL_KQYRB = 0 Then
                                        JS_end_KQYRB = JS_end_KQYRB + 1
                                    End If
                                    '根据装机量判断是否直接进入下一次循环
                                    If a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + ZJZGL_ZRXXHL + ZJRGL_HS_a_2 < RFH_ALL Then
                                        GoTo fff
                                    End If
                                    '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                    Dim JS_end_ZRXXHL As Integer = 0
                                    For a_7 = 0 To ZJZGL_ZRXXHL Step ZJZGL_ZRXXHL / FHFPCS
                                        '直接结束计算
                                        If JS_end_ZRXXHL >= 1 Then
                                            Exit For
                                        End If
                                        '如果不存在装机或者装机功率等于0
                                        If ZJJC_ZRXXHL = 0 Or ZJZGL_ZRXXHL = 0 Then
                                            JS_end_ZRXXHL = JS_end_ZRXXHL + 1
                                        End If
                                        '当前混水设备热功率
                                        Dim GRGL_HS_now As Double
                                        If ZJRGL_HS_ALL > 0 Then
                                            '默认天然气锅炉、电锅炉、直燃型溴化锂不会同时存在
                                            GRGL_HS_now = (a_1 + a_2 + a_7) * HSGRGLBL / (1 - HSGRGLBL)
                                        Else
                                            GRGL_HS_now = 0
                                        End If
                                        '根据装机量判断是否直接进入下一次循环
                                        If a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + a_7 + GRGL_HS_now < RFH_ALL Then
                                            GoTo ggg
                                        ElseIf a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + a_7 + GRGL_HS_now > RFH_ALL * TZXS Then
                                            GoTo ggg
                                        End If
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                        '天然气锅炉计算
                                        If a_1 >= FH_min_TRQGL And ZJJC_TRQGL = 1 And ZJZGL_TRQGL > 0 Then
                                            Dim ans_TRQGL = 天然气锅炉供热寻优计算(ExcelApp, b, FHTJJD, a_1 + GRGL_HS_now, FHFPCS, RFH_GR_now, RFH_XR_now, D_price, TRQ_price, HSGRGLBL, NUM1_TRQGL, NUM2_TRQGL, ZJRGL1_TRQGL, ZJRGL2_TRQGL, BTHQ1_ED_TRQGL, BTHQ2_ED_TRQGL, FJHD1_ED_TRQGL, FJHD2_ED_TRQGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                                            HD_ALL_TRQGL.AddRange(ans_TRQGL(0))
                                            GR_ALL_TRQGL.AddRange(ans_TRQGL(1))
                                            XR_ALL_TRQGL.AddRange(ans_TRQGL(2))
                                            FHL1_GR_TRQGL.AddRange(ans_TRQGL(4))
                                            FHL1_XR_TRQGL.AddRange(ans_TRQGL(5))
                                            FHL2_GR_TRQGL.AddRange(ans_TRQGL(6))
                                            FHL2_XR_TRQGL.AddRange(ans_TRQGL(7))
                                            HQ_ALL_min_TRQGL.AddRange(ans_TRQGL(8))
                                            FHL_HS_TRQGL.AddRange(ans_TRQGL(9))
                                        End If
                                        '电锅炉计算
                                        If a_2 >= FH_min_DGL And ZJJC_DGL = 1 And ZJZGL_DGL > 0 Then
                                            Dim ans_DGL = 电锅炉供热和蓄热分配寻优计算(ExcelApp, b, FHTJJD, a_2 + GRGL_HS_now, FHFPCS, RFH_GR_now, RFH_XR_now, HSGRGLBL, NUM1_DGL, NUM2_DGL, ZJRGL1_DGL, ZJRGL2_DGL, BTHD1_ED_DGL, BTHD2_ED_DGL, FJHD1_ED_DGL, FJHD2_ED_DGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                                            HD_ALL_min_DGL.AddRange(ans_DGL(0))
                                            GR_ALL_DGL.AddRange(ans_DGL(1))
                                            XR_ALL_DGL.AddRange(ans_DGL(2))
                                            FHL1_GR_DGL.AddRange(ans_DGL(4))
                                            FHL1_XR_DGL.AddRange(ans_DGL(5))
                                            FHL2_GR_DGL.AddRange(ans_DGL(6))
                                            FHL2_XR_DGL.AddRange(ans_DGL(7))
                                            FHL_HS_DGL.AddRange(ans_DGL(8))
                                        End If
                                        '水（地）源热泵计算
                                        If a_3 >= FH_min_SDYRB And ZJJC_SDYRB = 1 And ZJZGL_SDYRB > 0 Then
                                            Dim ans_SDYRB = 水_地源热泵供热和蓄热分配寻优计算(ExcelApp, b, FHTJJD, a_3, FHFPCS, RFH_GR_now, RFH_XR_now, NUM1_SDYRB, NUM2_SDYRB, ZJRGL1_SDYRB, ZJRGL2_SDYRB, BTHD1_ED_SDYRB, BTHD2_ED_SDYRB, FJHD1_ED_SDYRB, FJHD2_ED_SDYRB, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                                            HD_ALL_min_SDYRB.AddRange(ans_SDYRB(0))
                                            GR_ALL_SDYRB.AddRange(ans_SDYRB(1))
                                            XR_ALL_SDYRB.AddRange(ans_SDYRB(2))
                                            FHL1_GR_SDYRB.AddRange(ans_SDYRB(4))
                                            FHL1_XR_SDYRB.AddRange(ans_SDYRB(5))
                                            FHL2_GR_SDYRB.AddRange(ans_SDYRB(6))
                                            FHL2_XR_SDYRB.AddRange(ans_SDYRB(7))
                                        End If
                                        '离心式热泵计算
                                        If a_4 >= FH_min_LXSRB And ZJJC_LXSRB = 1 And ZJZGL_LXSRB > 0 Then
                                            Dim ans_LXSRB = 离心式热泵供热和蓄热分配寻优计算(ExcelApp, b, FHTJJD, a_4, FHFPCS, RFH_GR_now, RFH_XR_now, TJGRFHBL, NUM1_LXSRB, NUM2_LXSRB, ZJRGL1_LXSRB, ZJRGL2_LXSRB, BTHD1_ED_LXSRB, BTHD2_ED_LXSRB, FJHD1_ED_LXSRB, FJHD2_ED_LXSRB, NUM1_TJ, NUM2_TJ, FHL1_min_TJ, FHL2_min_TJ, ZJRGL1_TJ, ZJRGL2_TJ, BTHD1_ED_TJ, BTHD2_ED_TJ, FJHD1_ED_TJ, FJHD2_ED_TJ, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                                            HD_ALL_min_LXSRB.AddRange(ans_LXSRB(0))
                                            GR_ALL_LXSRB.AddRange(ans_LXSRB(1))
                                            XR_ALL_LXSRB.AddRange(ans_LXSRB(2))
                                            FHL1_GR_LXSRB.AddRange(ans_LXSRB(4))
                                            FHL1_XR_LXSRB.AddRange(ans_LXSRB(5))
                                            FHL2_GR_LXSRB.AddRange(ans_LXSRB(6))
                                            FHL2_XR_LXSRB.AddRange(ans_LXSRB(7))
                                            FHL_TJGR_LXSRB.AddRange(ans_LXSRB(8))
                                        End If
                                        '风冷螺杆机计算
                                        If a_5 >= FH_min_FLLGJ And ZJJC_FLLGJ = 1 And ZJZGL_FLLGJ > 0 Then
                                            Dim ans_FLLGJ = 风冷螺杆机供热和蓄热分配寻优计算(ExcelApp, b, FHTJJD, a_5, FHFPCS, RFH_GR_now, RFH_XR_now, NUM1_FLLGJ, NUM2_FLLGJ, ZJRGL1_FLLGJ, ZJRGL2_FLLGJ, BTHD1_ED_FLLGJ, BTHD2_ED_FLLGJ, FJHD1_ED_FLLGJ, FJHD2_ED_FLLGJ, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
                                            HD_ALL_min_FLLGJ.AddRange(ans_FLLGJ(0))
                                            GR_ALL_FLLGJ.AddRange(ans_FLLGJ(1))
                                            XR_ALL_FLLGJ.AddRange(ans_FLLGJ(2))
                                            FHL1_GR_FLLGJ.AddRange(ans_FLLGJ(4))
                                            FHL1_XR_FLLGJ.AddRange(ans_FLLGJ(5))
                                            FHL2_GR_FLLGJ.AddRange(ans_FLLGJ(6))
                                            FHL2_XR_FLLGJ.AddRange(ans_FLLGJ(7))
                                        End If
                                        '空气源热泵计算
                                        If a_6 >= FH_min_KQYRB And ZJJC_KQYRB = 1 And ZJZGL_KQYRB > 0 Then
                                            Dim ans_KQYRB = 空气源热泵供热和蓄热分配寻优计算(ExcelApp, b, FHTJJD, a_6, FHFPCS, RFH_GR_now, RFH_XR_now, NUM1_KQYRB, NUM2_KQYRB, ZJRGL1_KQYRB, ZJRGL2_KQYRB, BTHD1_ED_KQYRB, BTHD2_ED_KQYRB, FJHD1_ED_KQYRB, FJHD2_ED_KQYRB, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
                                            HD_ALL_min_KQYRB.AddRange(ans_KQYRB(0))
                                            GR_ALL_KQYRB.AddRange(ans_KQYRB(1))
                                            XR_ALL_KQYRB.AddRange(ans_KQYRB(2))
                                            FHL1_GR_KQYRB.AddRange(ans_KQYRB(4))
                                            FHL1_XR_KQYRB.AddRange(ans_KQYRB(5))
                                            FHL2_GR_KQYRB.AddRange(ans_KQYRB(6))
                                            FHL2_XR_KQYRB.AddRange(ans_KQYRB(7))
                                        End If
                                        '直燃型溴化锂计算
                                        If a_7 >= FH_min_ZRXXHL And ZJJC_ZRXXHL = 1 And ZJZGL_ZRXXHL > 0 Then
                                            Dim ans_ZRXXHL = 直燃型溴化锂供热寻优计算(ExcelApp, b, FHTJJD, a_7 + GRGL_HS_now, FHFPCS, RFH_GR_now, RFH_XR_now, D_price, TRQ_price, HSGRGLBL, NUM1_ZRXXHL, NUM2_ZRXXHL, ZJRGL1_ZRXXHL, ZJRGL2_ZRXXHL, BTHQ1_ED_ZRXXHL, BTHQ2_ED_ZRXXHL, FJHD1_ED_ZRXXHL, FJHD2_ED_ZRXXHL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                                            HD_ALL_ZRXXHL.AddRange(ans_ZRXXHL(0))
                                            GR_ALL_ZRXXHL.AddRange(ans_ZRXXHL(1))
                                            XR_ALL_ZRXXHL.AddRange(ans_ZRXXHL(2))
                                            FHL1_GR_ZRXXHL.AddRange(ans_ZRXXHL(4))
                                            FHL1_XR_ZRXXHL.AddRange(ans_ZRXXHL(5))
                                            FHL2_GR_ZRXXHL.AddRange(ans_ZRXXHL(6))
                                            FHL2_XR_ZRXXHL.AddRange(ans_ZRXXHL(7))
                                            HQ_ALL_min_ZRXXHL.AddRange(ans_ZRXXHL(8))
                                            FHL_HS_ZRXXHL.AddRange(ans_ZRXXHL(9))
                                        End If
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                        '求各个计算结果列表的长度，并求出最长的列表长度
                                        Dim n_1 As Integer = HQ_ALL_min_TRQGL.Count
                                        Dim n_2 As Integer = HD_ALL_min_DGL.Count
                                        Dim n_3 As Integer = HD_ALL_min_SDYRB.Count
                                        Dim n_4 As Integer = HD_ALL_min_LXSRB.Count
                                        Dim n_5 As Integer = HD_ALL_min_FLLGJ.Count
                                        Dim n_6 As Integer = HD_ALL_min_KQYRB.Count
                                        Dim n_7 As Integer = HQ_ALL_min_ZRXXHL.Count
                                        '用列表，查找所有结果长度的最大值
                                        Dim n_max As New List(Of Integer)
                                        n_max.Add(n_1)
                                        n_max.Add(n_2)
                                        n_max.Add(n_3)
                                        n_max.Add(n_4)
                                        n_max.Add(n_5)
                                        n_max.Add(n_6)
                                        n_max.Add(n_7)
                                        '求最大值
                                        Dim n_max_result As Integer = n_max.Max
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                        '天然气锅炉计算结果的列表长度修正
                                        If a_1 < FH_min_TRQGL Or ZJJC_TRQGL = 0 Or ZJZGL_TRQGL = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_1
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_TRQGL.AddRange(ans_temp)
                                                GR_ALL_TRQGL.AddRange(ans_temp)
                                                XR_ALL_TRQGL.AddRange(ans_temp)
                                                FHL1_GR_TRQGL.AddRange(ans_temp)
                                                FHL1_XR_TRQGL.AddRange(ans_temp)
                                                FHL2_GR_TRQGL.AddRange(ans_temp)
                                                FHL2_XR_TRQGL.AddRange(ans_temp)
                                                HQ_ALL_min_TRQGL.AddRange(ans_temp)
                                                FHL_HS_TRQGL.AddRange(ans_temp)
                                            End If
                                        End If
                                        '电锅炉计算结果的列表长度修正
                                        If a_2 < FH_min_DGL Or ZJJC_DGL = 0 Or ZJZGL_DGL = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_2
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_DGL.AddRange(ans_temp)
                                                GR_ALL_DGL.AddRange(ans_temp)
                                                XR_ALL_DGL.AddRange(ans_temp)
                                                FHL1_GR_DGL.AddRange(ans_temp)
                                                FHL1_XR_DGL.AddRange(ans_temp)
                                                FHL2_GR_DGL.AddRange(ans_temp)
                                                FHL2_XR_DGL.AddRange(ans_temp)
                                                FHL_HS_DGL.AddRange(ans_temp)
                                            End If
                                        End If
                                        '水（地）源热泵计算结果的列表长度修正
                                        If a_3 < FH_min_SDYRB Or ZJJC_SDYRB = 0 Or ZJZGL_SDYRB = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_3
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_SDYRB.AddRange(ans_temp)
                                                GR_ALL_SDYRB.AddRange(ans_temp)
                                                XR_ALL_SDYRB.AddRange(ans_temp)
                                                FHL1_GR_SDYRB.AddRange(ans_temp)
                                                FHL1_XR_SDYRB.AddRange(ans_temp)
                                                FHL2_GR_SDYRB.AddRange(ans_temp)
                                                FHL2_XR_SDYRB.AddRange(ans_temp)
                                            End If
                                        End If
                                        '离心式热泵计算结果的列表长度修正
                                        If a_4 < FH_min_LXSRB Or ZJJC_LXSRB = 0 Or ZJZGL_LXSRB = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_4
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_LXSRB.AddRange(ans_temp)
                                                GR_ALL_LXSRB.AddRange(ans_temp)
                                                XR_ALL_LXSRB.AddRange(ans_temp)
                                                FHL1_GR_LXSRB.AddRange(ans_temp)
                                                FHL1_XR_LXSRB.AddRange(ans_temp)
                                                FHL2_GR_LXSRB.AddRange(ans_temp)
                                                FHL2_XR_LXSRB.AddRange(ans_temp)
                                                FHL_TJGR_LXSRB.AddRange(ans_temp)
                                            End If
                                        End If
                                        '风冷螺杆机计算结果的列表长度修正
                                        If a_5 < FH_min_FLLGJ Or ZJJC_FLLGJ = 0 Or ZJZGL_FLLGJ = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_5
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_FLLGJ.AddRange(ans_temp)
                                                GR_ALL_FLLGJ.AddRange(ans_temp)
                                                XR_ALL_FLLGJ.AddRange(ans_temp)
                                                FHL1_GR_FLLGJ.AddRange(ans_temp)
                                                FHL1_XR_FLLGJ.AddRange(ans_temp)
                                                FHL2_GR_FLLGJ.AddRange(ans_temp)
                                                FHL2_XR_FLLGJ.AddRange(ans_temp)
                                            End If
                                        End If
                                        '空气源热泵计算结果的列表长度修正
                                        If a_6 < FH_min_KQYRB Or ZJJC_KQYRB = 0 Or ZJZGL_KQYRB = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_6
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_KQYRB.AddRange(ans_temp)
                                                GR_ALL_KQYRB.AddRange(ans_temp)
                                                XR_ALL_KQYRB.AddRange(ans_temp)
                                                FHL1_GR_KQYRB.AddRange(ans_temp)
                                                FHL1_XR_KQYRB.AddRange(ans_temp)
                                                FHL2_GR_KQYRB.AddRange(ans_temp)
                                                FHL2_XR_KQYRB.AddRange(ans_temp)
                                            End If
                                        End If
                                        '直燃型溴化锂计算结果的列表长度修正
                                        If a_7 < FH_min_ZRXXHL Or ZJJC_ZRXXHL = 0 Or ZJZGL_ZRXXHL = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_7
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_ZRXXHL.AddRange(ans_temp)
                                                GR_ALL_ZRXXHL.AddRange(ans_temp)
                                                XR_ALL_ZRXXHL.AddRange(ans_temp)
                                                FHL1_GR_ZRXXHL.AddRange(ans_temp)
                                                FHL1_XR_ZRXXHL.AddRange(ans_temp)
                                                FHL2_GR_ZRXXHL.AddRange(ans_temp)
                                                FHL2_XR_ZRXXHL.AddRange(ans_temp)
                                                HQ_ALL_min_ZRXXHL.AddRange(ans_temp)
                                                FHL_HS_ZRXXHL.AddRange(ans_temp)
                                            End If
                                        End If
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
ggg:
                                    Next
fff:
                                Next
eee:
                            Next
ddd:
                        Next
ccc:
                    Next
bbb:
                Next
aaa:
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '再次检查各个计算结果的列表长度，找到最短的列表
            Dim len_1 As Integer = HQ_ALL_min_TRQGL.Count
            Dim len_2 As Integer = HD_ALL_min_DGL.Count
            Dim len_3 As Integer = HD_ALL_min_SDYRB.Count
            Dim len_4 As Integer = HD_ALL_min_LXSRB.Count
            Dim len_5 As Integer = HD_ALL_min_FLLGJ.Count
            Dim len_6 As Integer = HD_ALL_min_KQYRB.Count
            Dim len_7 As Integer = HQ_ALL_min_ZRXXHL.Count
            '用列表，查找所有结果长度的最大值
            Dim len_min As New List(Of Integer)
            len_min.Add(len_1)
            len_min.Add(len_2)
            len_min.Add(len_3)
            len_min.Add(len_4)
            len_min.Add(len_5)
            len_min.Add(len_6)
            len_min.Add(len_7)
            '求最小值
            Dim len_min_result As Integer = len_min.Min
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '遍历结果，求和
            '耗电耗气
            Dim HQ_ALL_r As New List(Of Double)
            Dim HD_ALL_r As New List(Of Double)
            '总成本
            Dim COST_ALL As New List(Of Double)
            '供热和蓄热
            Dim GR_ALL_r As New List(Of Double)
            Dim XR_ALL_r As New List(Of Double)
            '遍历结果，找出符合要求的各种设备运行负荷率
            '天然气锅炉
            Dim FHL1_GR_TRQGL_r As New List(Of Double)
            Dim FHL1_XR_TRQGL_r As New List(Of Double)
            Dim FHL2_GR_TRQGL_r As New List(Of Double)
            Dim FHL2_XR_TRQGL_r As New List(Of Double)
            '电锅炉
            Dim FHL1_GR_DGL_r As New List(Of Double)
            Dim FHL1_XR_DGL_r As New List(Of Double)
            Dim FHL2_GR_DGL_r As New List(Of Double)
            Dim FHL2_XR_DGL_r As New List(Of Double)
            '水（地）源热泵
            Dim FHL1_GR_SDYRB_r As New List(Of Double)
            Dim FHL1_XR_SDYRB_r As New List(Of Double)
            Dim FHL2_GR_SDYRB_r As New List(Of Double)
            Dim FHL2_XR_SDYRB_r As New List(Of Double)
            '离心式热泵
            Dim FHL1_GR_LXSRB_r As New List(Of Double)
            Dim FHL1_XR_LXSRB_r As New List(Of Double)
            Dim FHL2_GR_LXSRB_r As New List(Of Double)
            Dim FHL2_XR_LXSRB_r As New List(Of Double)
            '风冷螺杆机
            Dim FHL1_GR_FLLGJ_r As New List(Of Double)
            Dim FHL1_XR_FLLGJ_r As New List(Of Double)
            Dim FHL2_GR_FLLGJ_r As New List(Of Double)
            Dim FHL2_XR_FLLGJ_r As New List(Of Double)
            '空气源热泵
            Dim FHL1_GR_KQYRB_r As New List(Of Double)
            Dim FHL1_XR_KQYRB_r As New List(Of Double)
            Dim FHL2_GR_KQYRB_r As New List(Of Double)
            Dim FHL2_XR_KQYRB_r As New List(Of Double)
            '直燃型溴化锂
            Dim FHL1_GR_ZRXXHL_r As New List(Of Double)
            Dim FHL1_XR_ZRXXHL_r As New List(Of Double)
            Dim FHL2_GR_ZRXXHL_r As New List(Of Double)
            Dim FHL2_XR_ZRXXHL_r As New List(Of Double)
            '混水设备负荷率
            Dim FHL_HS_DGL_r As New List(Of Double)
            Dim FHL_HS_TRQGL_r As New List(Of Double)
            Dim FHL_HS_ZRXXHL_r As New List(Of Double)
            '梯级供热负荷率
            Dim FHL_TJGR_LSXRB_r As New List(Of Double)
            '遍历所有结果
            For i = 0 To len_min_result - 1
                Dim GR_ALL_i As Double = GR_ALL_TRQGL(i) + GR_ALL_DGL(i) + GR_ALL_SDYRB(i) + GR_ALL_LXSRB(i) + GR_ALL_FLLGJ(i) + GR_ALL_KQYRB(i) + GR_ALL_ZRXXHL(i)
                Dim XR_ALL_i As Double = XR_ALL_TRQGL(i) + XR_ALL_DGL(i) + XR_ALL_SDYRB(i) + XR_ALL_LXSRB(i) + XR_ALL_FLLGJ(i) + XR_ALL_KQYRB(i) + XR_ALL_ZRXXHL(i)
                '如果满足条件
                If GR_ALL_i >= RFH_GR_now And XR_ALL_i >= RFH_XR_now Then
                    '耗电、耗气
                    Dim HQ_ALL_i As Double = HQ_ALL_min_ZRXXHL(i) + HQ_ALL_min_TRQGL(i)
                    Dim HD_ALL_i As Double = HD_ALL_TRQGL(i) + HD_ALL_min_DGL(i) + HD_ALL_min_SDYRB(i) + HD_ALL_min_LXSRB(i) + HD_ALL_min_FLLGJ(i) + HD_ALL_min_KQYRB(i) + HD_ALL_ZRXXHL(i)
                    HQ_ALL_r.Add(HQ_ALL_i)
                    HD_ALL_r.Add(HD_ALL_i)
                    '总成本
                    COST_ALL.Add(HD_ALL_i * D_price + HQ_ALL_i * TRQ_price)
                    '供热、蓄热
                    GR_ALL_r.Add(GR_ALL_i)
                    XR_ALL_r.Add(XR_ALL_i)
                    '各设备负荷率计算结果
                    '天然气锅炉
                    FHL1_GR_TRQGL_r.Add(FHL1_GR_TRQGL(i))
                    FHL1_XR_TRQGL_r.Add(FHL1_XR_TRQGL(i))
                    FHL2_GR_TRQGL_r.Add(FHL2_GR_TRQGL(i))
                    FHL2_XR_TRQGL_r.Add(FHL2_XR_TRQGL(i))
                    '电锅炉
                    FHL1_GR_DGL_r.Add(FHL1_GR_DGL(i))
                    FHL1_XR_DGL_r.Add(FHL1_XR_DGL(i))
                    FHL2_GR_DGL_r.Add(FHL2_GR_DGL(i))
                    FHL2_XR_DGL_r.Add(FHL2_XR_DGL(i))
                    '水地源热泵
                    FHL1_GR_SDYRB_r.Add(FHL1_GR_SDYRB(i))
                    FHL1_XR_SDYRB_r.Add(FHL1_XR_SDYRB(i))
                    FHL2_GR_SDYRB_r.Add(FHL2_GR_SDYRB(i))
                    FHL2_XR_SDYRB_r.Add(FHL2_XR_SDYRB(i))
                    '离心式热泵
                    FHL1_GR_LXSRB_r.Add(FHL1_GR_LXSRB(i))
                    FHL1_XR_LXSRB_r.Add(FHL1_XR_LXSRB(i))
                    FHL2_GR_LXSRB_r.Add(FHL2_GR_LXSRB(i))
                    FHL2_XR_LXSRB_r.Add(FHL2_XR_LXSRB(i))
                    '风冷螺杆机
                    FHL1_GR_FLLGJ_r.Add(FHL1_GR_FLLGJ(i))
                    FHL1_XR_FLLGJ_r.Add(FHL1_XR_FLLGJ(i))
                    FHL2_GR_FLLGJ_r.Add(FHL2_GR_FLLGJ(i))
                    FHL2_XR_FLLGJ_r.Add(FHL2_XR_FLLGJ(i))
                    '空气源热泵
                    FHL1_GR_KQYRB_r.Add(FHL1_GR_KQYRB(i))
                    FHL1_XR_KQYRB_r.Add(FHL1_XR_KQYRB(i))
                    FHL2_GR_KQYRB_r.Add(FHL2_GR_KQYRB(i))
                    FHL2_XR_KQYRB_r.Add(FHL2_XR_KQYRB(i))
                    '直燃型溴化锂
                    FHL1_GR_ZRXXHL_r.Add(FHL1_GR_ZRXXHL(i))
                    FHL1_XR_ZRXXHL_r.Add(FHL1_XR_ZRXXHL(i))
                    FHL2_GR_ZRXXHL_r.Add(FHL2_GR_ZRXXHL(i))
                    FHL2_XR_ZRXXHL_r.Add(FHL2_XR_ZRXXHL(i))
                    '混水设备负荷率
                    FHL_HS_DGL_r.Add(FHL_HS_DGL(i))
                    FHL_HS_TRQGL_r.Add(FHL_HS_TRQGL(i))
                    FHL_HS_ZRXXHL_r.Add(FHL_HS_ZRXXHL(i))
                    '梯级供热负荷率
                    FHL_TJGR_LSXRB_r.Add(FHL_TJGR_LXSRB(i))
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '找出总（电、气）成本最低的运行模式
            Dim COST_min As Double
            Try
                COST_min = COST_ALL.Min
            Catch ex As Exception
                JS_start = JS_start + 1
                '返回zzzzz处重算
                GoTo zzzzz
            End Try
            '找到所在标签
            Dim COST_min_index As Integer = COST_ALL.IndexOf(COST_min)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '找到总成本最低对应的各种结果
            '耗电耗气
            Dim HQ_ALL_result As Double = HQ_ALL_r(COST_min_index)
            Dim HD_ALL_result As Double = HD_ALL_r(COST_min_index)
            '总成本
            Dim COST_ALL_result As Double = COST_ALL(COST_min_index)
            '供热和蓄热
            Dim GR_ALL_result As Double = GR_ALL_r(COST_min_index)
            Dim XR_ALL_result As Double = XR_ALL_r(COST_min_index)
            '天然气锅炉
            Dim FHL1_GR_TRQGL_result As Double = FHL1_GR_TRQGL_r(COST_min_index)
            Dim FHL1_XR_TRQGL_result As Double = FHL1_XR_TRQGL_r(COST_min_index)
            Dim FHL2_GR_TRQGL_result As Double = FHL2_GR_TRQGL_r(COST_min_index)
            Dim FHL2_XR_TRQGL_result As Double = FHL2_XR_TRQGL_r(COST_min_index)
            '电锅炉
            Dim FHL1_GR_DGL_result As Double = FHL1_GR_DGL_r(COST_min_index)
            Dim FHL1_XR_DGL_result As Double = FHL1_XR_DGL_r(COST_min_index)
            Dim FHL2_GR_DGL_result As Double = FHL2_GR_DGL_r(COST_min_index)
            Dim FHL2_XR_DGL_result As Double = FHL2_XR_DGL_r(COST_min_index)
            '水（地）源热泵
            Dim FHL1_GR_SDYRB_result As Double = FHL1_GR_SDYRB_r(COST_min_index)
            Dim FHL1_XR_SDYRB_result As Double = FHL1_XR_SDYRB_r(COST_min_index)
            Dim FHL2_GR_SDYRB_result As Double = FHL2_GR_SDYRB_r(COST_min_index)
            Dim FHL2_XR_SDYRB_result As Double = FHL2_XR_SDYRB_r(COST_min_index)
            '离心式热泵
            Dim FHL1_GR_LXSRB_result As Double = FHL1_GR_LXSRB_r(COST_min_index)
            Dim FHL1_XR_LXSRB_result As Double = FHL1_XR_LXSRB_r(COST_min_index)
            Dim FHL2_GR_LXSRB_result As Double = FHL2_GR_LXSRB_r(COST_min_index)
            Dim FHL2_XR_LXSRB_result As Double = FHL2_XR_LXSRB_r(COST_min_index)
            '风冷螺杆机
            Dim FHL1_GR_FLLGJ_result As Double = FHL1_GR_FLLGJ_r(COST_min_index)
            Dim FHL1_XR_FLLGJ_result As Double = FHL1_XR_FLLGJ_r(COST_min_index)
            Dim FHL2_GR_FLLGJ_result As Double = FHL2_GR_FLLGJ_r(COST_min_index)
            Dim FHL2_XR_FLLGJ_result As Double = FHL2_XR_FLLGJ_r(COST_min_index)
            '空气源热泵
            Dim FHL1_GR_KQYRB_result As Double = FHL1_GR_KQYRB_r(COST_min_index)
            Dim FHL1_XR_KQYRB_result As Double = FHL1_XR_KQYRB_r(COST_min_index)
            Dim FHL2_GR_KQYRB_result As Double = FHL2_GR_KQYRB_r(COST_min_index)
            Dim FHL2_XR_KQYRB_result As Double = FHL2_XR_KQYRB_r(COST_min_index)
            '直燃型溴化锂
            Dim FHL1_GR_ZRXXHL_result As Double = FHL1_GR_ZRXXHL_r(COST_min_index)
            Dim FHL1_XR_ZRXXHL_result As Double = FHL1_XR_ZRXXHL_r(COST_min_index)
            Dim FHL2_GR_ZRXXHL_result As Double = FHL2_GR_ZRXXHL_r(COST_min_index)
            Dim FHL2_XR_ZRXXHL_result As Double = FHL2_XR_ZRXXHL_r(COST_min_index)
            '混水设备负荷率
            Dim FHL_HS_DGL_result As Double = FHL_HS_DGL_r(COST_min_index)
            Dim FHL_HS_TRQGL_result As Double = FHL_HS_TRQGL_r(COST_min_index)
            Dim FHL_HS_ZRXXHL_result As Double = FHL_HS_ZRXXHL_r(COST_min_index)
            '梯级供热负荷率
            Dim FHL_TJGR_LXSRB_result As Double = FHL_TJGR_LSXRB_r(COST_min_index)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '针对计算出的结果进行放缩，防止出现大的误差
            '供热、蓄热负荷出力和需求量的比例
            Dim GR_bl As Double = GR_ALL_result / RFH_GR_now
            Dim XR_bl As Double = XR_ALL_result / RFH_XR_now
            '供热
            If GR_bl > WCXS Then
                '天然气锅炉
                FHL1_GR_TRQGL_result = FHL1_GR_TRQGL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_TRQGL_result = FHL2_GR_TRQGL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '电锅炉
                FHL1_GR_DGL_result = FHL1_GR_DGL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_DGL_result = FHL2_GR_DGL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '水（地）源热泵
                FHL1_GR_SDYRB_result = FHL1_GR_SDYRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_SDYRB_result = FHL2_GR_SDYRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '离心式热泵
                FHL1_GR_LXSRB_result = FHL1_GR_LXSRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_LXSRB_result = FHL2_GR_LXSRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '风冷螺杆机
                FHL1_GR_FLLGJ_result = FHL1_GR_FLLGJ_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_FLLGJ_result = FHL2_GR_FLLGJ_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '空气源热泵
                FHL1_GR_KQYRB_result = FHL1_GR_KQYRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_KQYRB_result = FHL2_GR_KQYRB_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '直燃型溴化锂
                FHL1_GR_ZRXXHL_result = FHL1_GR_ZRXXHL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GR_ZRXXHL_result = FHL2_GR_ZRXXHL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                '混水设备负荷率（天然气锅炉、直燃型溴化锂）
                FHL_HS_TRQGL_result = FHL_HS_TRQGL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
                FHL_HS_ZRXXHL_result = FHL_HS_ZRXXHL_r(COST_min_index) / GR_bl * (1 + 2 * FHTJJD / 100)
            End If
            '蓄热
            If XR_bl > WCXS Then
                '天然气锅炉 
                FHL1_XR_TRQGL_result = FHL1_XR_TRQGL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_TRQGL_result = FHL2_XR_TRQGL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '电锅炉
                FHL1_XR_DGL_result = FHL1_XR_DGL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_DGL_result = FHL2_XR_DGL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '水（地）源热泵
                FHL1_XR_SDYRB_result = FHL1_XR_SDYRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_SDYRB_result = FHL2_XR_SDYRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '离心式热泵
                FHL1_XR_LXSRB_result = FHL1_XR_LXSRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_LXSRB_result = FHL2_XR_LXSRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '风冷螺杆机
                FHL1_XR_FLLGJ_result = FHL1_XR_FLLGJ_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_FLLGJ_result = FHL2_XR_FLLGJ_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '空气源热泵
                FHL1_XR_KQYRB_result = FHL1_XR_KQYRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_KQYRB_result = FHL2_XR_KQYRB_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                '直燃型溴化锂
                FHL1_XR_ZRXXHL_result = FHL1_XR_ZRXXHL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XR_ZRXXHL_result = FHL2_XR_ZRXXHL_r(COST_min_index) / XR_bl * (1 + 2 * FHTJJD / 100)
            End If
            '修正混水设备负荷率（电锅炉）
            Dim GR_XR_bl As Double = (GR_ALL_result + XR_ALL_result) / (RFH_GR_now + RFH_XR_now)
            If GR_XR_bl > WCXS Then
                FHL_HS_DGL_result = FHL_HS_DGL_r(COST_min_index) / GR_XR_bl * (1 + 2 * FHTJJD / 100)
            End If
            '修正梯级供热设备负荷率
            If GR_XR_bl > WCXS Then
                FHL_TJGR_LXSRB_result = FHL_TJGR_LSXRB_r(COST_min_index) / GR_XR_bl * (1 + 2 * FHTJJD / 100)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '将各个设备的负荷率计算结果写入Excel
            '天然气锅炉（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = FHL1_GR_TRQGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = FHL1_GR_TRQGL_result
            '天然气锅炉（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = FHL2_GR_TRQGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = FHL2_GR_TRQGL_result
            '风冷螺杆机（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = FHL1_GR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value = FHL1_XR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 54).Value = FHL1_GR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62).Value = FHL1_XR_FLLGJ_result
            '风冷螺杆机（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = FHL2_GR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value = FHL2_XR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 55).Value = FHL2_GR_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 63).Value = FHL2_XR_FLLGJ_result
            '空气源热泵(1)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = FHL1_GR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value = FHL1_XR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 56).Value = FHL1_GR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 64).Value = FHL1_XR_KQYRB_result
            '空气源热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = FHL2_GR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value = FHL2_XR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 57).Value = FHL2_GR_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 65).Value = FHL2_XR_KQYRB_result
            '水（地）源热泵（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = FHL1_GR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value = FHL1_XR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 58).Value = FHL1_GR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 66).Value = FHL1_XR_SDYRB_result
            '水（地）源热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = FHL2_GR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value = FHL2_XR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 59).Value = FHL2_GR_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 67).Value = FHL2_XR_SDYRB_result
            '电锅炉（1） 
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = FHL1_GR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = FHL1_XR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = FHL1_GR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = FHL1_XR_DGL_result
            '电锅炉(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = FHL2_GR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = FHL2_XR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = FHL2_GR_DGL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = FHL2_XR_DGL_result
            '离心式热泵（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = FHL1_GR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value = FHL1_XR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value = FHL1_GR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value = FHL1_XR_LXSRB_result
            '离心式热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = FHL2_GR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value = FHL2_XR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value = FHL2_GR_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value = FHL2_XR_LXSRB_result
            '直燃型溴化锂（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = FHL1_GR_ZRXXHL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = FHL1_GR_ZRXXHL_result
            '直燃型溴化锂（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = FHL2_GR_ZRXXHL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = FHL2_GR_ZRXXHL_result
            '混水设备负荷率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = FHL_HS_DGL_result + FHL_HS_TRQGL_result + FHL_HS_ZRXXHL_result
            End If
            '梯级供热设备负荷率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = FHL_TJGR_LXSRB_result
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算制热季天然气耗量和耗电量综合修正系数
            Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
        End If
    End Sub
    Function 水_地源热泵供热和蓄热分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_SDYRB / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_SDYRB / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJRGL1 + ZJRGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1 + FHL2_min * ZJRGL2
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                Dim ans_temp = 水_地源热泵供热和蓄热计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GR.Add(ans_temp(1))
                FHL1_XR.Add(ans_temp(2))
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GR_out.Add(b_1)
                XR_out.Add(c_1)
                ZGR_out.Add(RFH_max)
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水_地源热泵供热和蓄热计算(ExcelApp, FHTJJD, RFH_max, 0, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水_地源热泵供热和蓄热计算(ExcelApp, FHTJJD, 0, RFH_max, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(0)
            XR_out.Add(RFH_max)
            ZGR_out.Add(RFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄热设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 离心式热泵供热和蓄热分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, TJGRFHBL As Double, NUM1_LXSRB As Double, NUM2_LXSRB As Double, ZJRGL1_LXSRB As Double, ZJRGL2_LXSRB As Double, BTHD1_ED_LXSRB As Double, BTHD2_ED_LXSRB As Double, FJHD1_ED_LXSRB As Double, FJHD2_ED_LXSRB As Double, NUM1_TJ As Double, NUM2_TJ As Double, FHL1_min_TJ As Double, FHL2_min_TJ As Double, ZJRGL1_TJ As Double, ZJRGL2_TJ As Double, BTHD1_ED_TJ As Double, BTHD2_ED_TJ As Double, FJHD1_ED_TJ As Double, FJHD2_ED_TJ As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1_LXSRB > 0 Then
            FHL1_min = FHL1_min_LXSRB / NUM1_LXSRB
        Else
            FHL1_min = 0
        End If
        If NUM2_LXSRB > 0 Then
            FHL2_min = FHL2_min_LXSRB / NUM2_LXSRB
        Else
            FHL2_min = 0
        End If
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1_LXSRB + ZJRGL2_LXSRB
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        Dim FHL_TJGR As New List(Of Double) '梯级供热负荷率（综合）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1_LXSRB + FHL2_min * ZJRGL2_LXSRB
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                Dim ans_temp = 离心式热泵供热和蓄热计算(ExcelApp, b, FHTJJD, b_1, c_1, TJGRFHBL, NUM1_LXSRB, NUM2_LXSRB, ZJRGL1_LXSRB, ZJRGL2_LXSRB, BTHD1_ED_LXSRB, BTHD2_ED_LXSRB, FJHD1_ED_LXSRB, FJHD2_ED_LXSRB, NUM1_TJ, NUM2_TJ, FHL1_min_TJ, FHL2_min_TJ, ZJRGL1_TJ, ZJRGL2_TJ, BTHD1_ED_TJ, BTHD2_ED_TJ, FJHD1_ED_TJ, FJHD2_ED_TJ, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GR.Add(ans_temp(1))
                FHL1_XR.Add(ans_temp(2))
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(ans_temp(4))
                FHL_TJGR.Add(ans_temp(5))
                NUM1_List.Add(ans_temp(6))
                NUM2_List.Add(ans_temp(7))
                GR_out.Add(b_1)
                XR_out.Add(c_1)
                ZGR_out.Add(RFH_max)
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式热泵供热和蓄热计算(ExcelApp, b, FHTJJD, RFH_max, 0, TJGRFHBL, NUM1_LXSRB, NUM2_LXSRB, ZJRGL1_LXSRB, ZJRGL2_LXSRB, BTHD1_ED_LXSRB, BTHD2_ED_LXSRB, FJHD1_ED_LXSRB, FJHD2_ED_LXSRB, NUM1_TJ, NUM2_TJ, FHL1_min_TJ, FHL2_min_TJ, ZJRGL1_TJ, ZJRGL2_TJ, BTHD1_ED_TJ, BTHD2_ED_TJ, FJHD1_ED_TJ, FJHD2_ED_TJ, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            FHL_TJGR.Add(ans_temp(5))
            NUM1_List.Add(ans_temp(6))
            NUM2_List.Add(ans_temp(7))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式热泵供热和蓄热计算(ExcelApp, b, FHTJJD, 0, RFH_max, TJGRFHBL, NUM1_LXSRB, NUM2_LXSRB, ZJRGL1_LXSRB, ZJRGL2_LXSRB, BTHD1_ED_LXSRB, BTHD2_ED_LXSRB, FJHD1_ED_LXSRB, FJHD2_ED_LXSRB, NUM1_TJ, NUM2_TJ, FHL1_min_TJ, FHL2_min_TJ, ZJRGL1_TJ, ZJRGL2_TJ, BTHD1_ED_TJ, BTHD2_ED_TJ, FJHD1_ED_TJ, FJHD2_ED_TJ, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            FHL_TJGR.Add(ans_temp(5))
            NUM1_List.Add(ans_temp(6))
            NUM2_List.Add(ans_temp(7))
            GR_out.Add(0)
            XR_out.Add(RFH_max)
            ZGR_out.Add(RFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_TJGR.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄热设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(11)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = FHL_TJGR
        ans(9) = NUM1_List
        ans(10) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 风冷螺杆机供热和蓄热分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_FLLGJ / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_FLLGJ / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1 + ZJRGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1 + FHL2_min * ZJRGL2
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                Dim ans_temp = 风冷螺杆机供热和蓄热计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GR.Add(ans_temp(1))
                FHL1_XR.Add(ans_temp(2))
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GR_out.Add(b_1)
                XR_out.Add(c_1)
                ZGR_out.Add(RFH_max)
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 风冷螺杆机供热和蓄热计算(ExcelApp, FHTJJD, RFH_max, 0, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 风冷螺杆机供热和蓄热计算(ExcelApp, FHTJJD, 0, RFH_max, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(0)
            XR_out.Add(RFH_max)
            ZGR_out.Add(RFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄热设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 空气源热泵供热和蓄热分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_KQYRB / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_KQYRB / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1 + ZJRGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1 + FHL2_min * ZJRGL2
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                Dim ans_temp = 空气源热泵供热和蓄热计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GR.Add(ans_temp(1))
                FHL1_XR.Add(ans_temp(2))
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GR_out.Add(b_1)
                XR_out.Add(c_1)
                ZGR_out.Add(RFH_max)
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 空气源热泵供热和蓄热计算(ExcelApp, FHTJJD, RFH_max, 0, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 空气源热泵供热和蓄热计算(ExcelApp, FHTJJD, 0, RFH_max, NUM1, NUM2, ZJRGL1, ZJRGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GR_air, BTHDXS_XR_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(0)
            XR_out.Add(RFH_max)
            ZGR_out.Add(RFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄热设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 电锅炉供热和蓄热分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, HSGRGLBL As Double, NUM1_DGL As Double, NUM2_DGL As Double, ZJRGL1_DGL As Double, ZJRGL2_DGL As Double, BTHD1_ED_DGL As Double, BTHD2_ED_DGL As Double, FJHD1_ED_DGL As Double, FJHD2_ED_DGL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '对于天然气采暖锅炉来说，只能供热，不能蓄热，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1_DGL > 0 Then
            FHL1_min = FHL1_min_DGL / NUM1_DGL
        Else
            FHL1_min = 0
        End If
        If NUM2_DGL > 0 Then
            FHL2_min = FHL2_min_DGL / NUM2_DGL
        Else
            FHL2_min = 0
        End If
        '混水设备负荷率下限（除以设备数量）
        Dim FHL1_min_HS_a As Double
        Dim FHL2_min_HS_a As Double
        If NUM1_HS > 0 Then
            FHL1_min_HS_a = FHL1_min_HS / NUM1_HS
        Else
            FHL1_min_HS_a = 0
        End If
        If NUM2_HS > 0 Then
            FHL2_min_HS_a = FHL2_min_HS / NUM2_HS
        Else
            FHL2_min_HS_a = 0
        End If
        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1_DGL + ZJRGL2_DGL + ZJRGL_HS_ALL
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        Dim FHL_HS As New List(Of Double) '混水供热负荷率（综合）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1_DGL + FHL2_min * ZJRGL2_DGL + FHL1_min_HS_a * ZJRGL1_HS + FHL2_min_HS_a * ZJRGL2_HS
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                Dim ans_temp = 电锅炉供热和蓄热计算(ExcelApp, b, FHTJJD, b_1, c_1, HSGRGLBL, NUM1_DGL, NUM2_DGL, ZJRGL1_DGL, ZJRGL2_DGL, BTHD1_ED_DGL, BTHD2_ED_DGL, FJHD1_ED_DGL, FJHD2_ED_DGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GR.Add(ans_temp(1))
                FHL1_XR.Add(ans_temp(2))
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(ans_temp(4))
                FHL_HS.Add(ans_temp(5))
                NUM1_List.Add(ans_temp(6))
                NUM2_List.Add(ans_temp(7))
                GR_out.Add(b_1)
                XR_out.Add(c_1)
                ZGR_out.Add(RFH_max)
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 电锅炉供热和蓄热计算(ExcelApp, b, FHTJJD, RFH_max, 0, HSGRGLBL, NUM1_DGL, NUM2_DGL, ZJRGL1_DGL, ZJRGL2_DGL, BTHD1_ED_DGL, BTHD2_ED_DGL, FJHD1_ED_DGL, FJHD2_ED_DGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            FHL_HS.Add(ans_temp(5))
            NUM1_List.Add(ans_temp(6))
            NUM2_List.Add(ans_temp(7))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 电锅炉供热和蓄热计算(ExcelApp, b, FHTJJD, 0, RFH_max, HSGRGLBL, NUM1_DGL, NUM2_DGL, ZJRGL1_DGL, ZJRGL2_DGL, BTHD1_ED_DGL, BTHD2_ED_DGL, FJHD1_ED_DGL, FJHD2_ED_DGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GR.Add(ans_temp(1))
            FHL1_XR.Add(ans_temp(2))
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(ans_temp(4))
            FHL_HS.Add(ans_temp(5))
            NUM1_List.Add(ans_temp(6))
            NUM2_List.Add(ans_temp(7))
            GR_out.Add(0)
            XR_out.Add(RFH_max)
            ZGR_out.Add(RFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_HS.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
        End If
        '将计算出的所有列表返回(电锅炉蓄热设备负荷率不需要转换成了台数)
        Dim ans(11)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = FHL_HS
        ans(9) = NUM1_List
        ans(10) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 天然气锅炉供热寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, D_price As Double, TRQ_price As Double, HSGRGLBL As Double, NUM1_TRQGL As Double, NUM2_TRQGL As Double, ZJRGL1_TRQGL As Double, ZJRGL2_TRQGL As Double, BTHQ1_ED_TRQGL As Double, BTHQ2_ED_TRQGL As Double, FJHD1_ED_TRQGL As Double, FJHD2_ED_TRQGL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '对于天然气采暖锅炉来说，只能供热，不能蓄热，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1_TRQGL > 0 Then
            FHL1_min = FHL1_min_TRQGL / NUM1_TRQGL
        Else
            FHL1_min = 0
        End If
        If NUM2_TRQGL > 0 Then
            FHL2_min = FHL2_min_TRQGL / NUM2_TRQGL
        Else
            FHL2_min = 0
        End If
        '混水设备负荷率下限（除以设备数量）
        Dim FHL1_min_HS_a As Double
        Dim FHL2_min_HS_a As Double
        If NUM1_HS > 0 Then
            FHL1_min_HS_a = FHL1_min_HS / NUM1_HS
        Else
            FHL1_min_HS_a = 0
        End If
        If NUM2_HS > 0 Then
            FHL2_min_HS_a = FHL2_min_HS / NUM2_HS
        Else
            FHL2_min_HS_a = 0
        End If
        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1_TRQGL + ZJRGL2_TRQGL + ZJRGL_HS_ALL
        '计算出的各种结果存入列表
        Dim HQ_ALL As New List(Of Double) '天然气总耗量
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        Dim FHL_HS As New List(Of Double) '混水供热负荷率（综合）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1_TRQGL + FHL2_min * ZJRGL2_TRQGL + FHL1_min_HS_a * ZJRGL1_HS + FHL2_min_HS_a * ZJRGL2_HS
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                'Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                '对于直燃型溴化锂来说，只能供热，不能蓄热，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
                Dim ans_temp = 天然气锅炉供热计算_方法二(ExcelApp, b, FHTJJD, RFH_max, D_price, TRQ_price, HSGRGLBL, NUM1_TRQGL, NUM2_TRQGL, ZJRGL1_TRQGL, ZJRGL2_TRQGL, BTHQ1_ED_TRQGL, BTHQ2_ED_TRQGL, FJHD1_ED_TRQGL, FJHD2_ED_TRQGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(1))
                FHL1_GR.Add(ans_temp(2))
                FHL1_XR.Add(0)
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(0)
                FHL_HS.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GR_out.Add(RFH_max)
                XR_out.Add(0)
                ZGR_out.Add(RFH_max)
                HQ_ALL.Add(ans_temp(0))
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 天然气锅炉供热计算_方法二(ExcelApp, b, FHTJJD, RFH_max, D_price, TRQ_price, HSGRGLBL, NUM1_TRQGL, NUM2_TRQGL, ZJRGL1_TRQGL, ZJRGL2_TRQGL, BTHQ1_ED_TRQGL, BTHQ2_ED_TRQGL, FJHD1_ED_TRQGL, FJHD2_ED_TRQGL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(1))
            FHL1_GR.Add(ans_temp(2))
            FHL1_XR.Add(0)
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(0)
            FHL_HS.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
            HQ_ALL.Add(ans_temp(0))
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '计算结果加入列表，直燃型溴化锂不能蓄热，结果全部是0
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_HS.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
            HQ_ALL.Add(0)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_HS.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
            HQ_ALL.Add(0)
        End If
        '将计算出的所有列表返回
        Dim ans(12)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = HQ_ALL
        ans(9) = FHL_HS
        ans(10) = NUM1_List
        ans(11) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂供热寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, RFH_GR_now As Double, RFH_XR_now As Double, D_price As Double, TRQ_price As Double, HSGRGLBL As Double, NUM1_ZRXXHL As Double, NUM2_ZRXXHL As Double, ZJRGL1_ZRXXHL As Double, ZJRGL2_ZRXXHL As Double, BTHQ1_ED_ZRXXHL As Double, BTHQ2_ED_ZRXXHL As Double, FJHD1_ED_ZRXXHL As Double, FJHD2_ED_ZRXXHL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '对于直燃型溴化锂来说，只能供热，不能蓄热，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
        'ZFH：供热和蓄热负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1_ZRXXHL > 0 Then
            FHL1_min = FHL1_min_ZRXXHL / NUM1_ZRXXHL
        Else
            FHL1_min = 0
        End If
        If NUM2_ZRXXHL > 0 Then
            FHL2_min = FHL2_min_ZRXXHL / NUM2_ZRXXHL
        Else
            FHL2_min = 0
        End If
        '混水设备负荷率下限（除以设备数量）
        Dim FHL1_min_HS_a As Double
        Dim FHL2_min_HS_a As Double
        If NUM1_HS > 0 Then
            FHL1_min_HS_a = FHL1_min_HS / NUM1_HS
        Else
            FHL1_min_HS_a = 0
        End If
        If NUM2_HS > 0 Then
            FHL2_min_HS_a = FHL2_min_HS / NUM2_HS
        Else
            FHL2_min_HS_a = 0
        End If
        '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '装机总功率
        Dim ZJZGL As Double = ZJRGL1_ZRXXHL + ZJRGL2_ZRXXHL + ZJRGL_HS_ALL
        '计算出的各种结果存入列表
        Dim HQ_ALL As New List(Of Double) '天然气总耗量
        Dim HD_ALL As New List(Of Double) '耗电总量
        Dim FHL1_GR As New List(Of Double) '综合
        Dim FHL1_XR As New List(Of Double) '综合
        Dim FHL2_GR As New List(Of Double) '综合
        Dim FHL2_XR As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GR_out As New List(Of Double) '供热总输出（总和）
        Dim XR_out As New List(Of Double) '蓄热总输出（总和）
        Dim ZGR_out As New List(Of Double) '总的供热（总和）+蓄热出力（总和）
        Dim FHL_HS As New List(Of Double) '混水供热负荷率（综合）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim RFH_min As Double = FHL1_min * ZJRGL1_ZRXXHL + FHL2_min * ZJRGL2_ZRXXHL + FHL1_min_HS_a * ZJRGL1_HS + FHL2_min_HS_a * ZJRGL2_HS
        '设备可以供热的上限（冷负荷需求量和设备装机量中较小的值）
        Dim RFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄热负荷，再将负荷分成供热负荷和蓄热负荷
        If RFH_GR_now > 0 And RFH_XR_now > 0 Then
            '如果供热和蓄热均大于0
            'b_1表示当前分配给供热的负荷
            For b_1 = 0 To RFH_max Step RFH_max / FHFPCS
                'c_1表示当前分配蓄热的负荷
                'Dim c_1 As Double = RFH_max - b_1
                '求计算结果
                '对于直燃型溴化锂来说，只能供热，不能蓄热，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
                Dim ans_temp = 直燃型溴化锂供热计算_方法二(ExcelApp, b, FHTJJD, RFH_max, D_price, TRQ_price, HSGRGLBL, NUM1_ZRXXHL, NUM2_ZRXXHL, ZJRGL1_ZRXXHL, ZJRGL2_ZRXXHL, BTHQ1_ED_ZRXXHL, BTHQ2_ED_ZRXXHL, FJHD1_ED_ZRXXHL, FJHD2_ED_ZRXXHL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(1))
                FHL1_GR.Add(ans_temp(2))
                FHL1_XR.Add(0)
                FHL2_GR.Add(ans_temp(3))
                FHL2_XR.Add(0)
                FHL_HS.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GR_out.Add(RFH_max)
                XR_out.Add(0)
                ZGR_out.Add(RFH_max)
                HQ_ALL.Add(ans_temp(0))
            Next
        ElseIf RFH_GR_now > 0 And RFH_XR_now = 0 Then
            '如果没有蓄热负荷，则直接计算
            '求计算结果
            Dim ans_temp = 直燃型溴化锂供热计算_方法二(ExcelApp, b, FHTJJD, RFH_max, D_price, TRQ_price, HSGRGLBL, NUM1_ZRXXHL, NUM2_ZRXXHL, ZJRGL1_ZRXXHL, ZJRGL2_ZRXXHL, BTHQ1_ED_ZRXXHL, BTHQ2_ED_ZRXXHL, FJHD1_ED_ZRXXHL, FJHD2_ED_ZRXXHL, NUM1_HS, NUM2_HS, FHL1_min_HS, FHL2_min_HS, ZJRGL1_HS, ZJRGL2_HS, BTHD1_ED_HS, BTHD2_ED_HS, FJHD1_ED_HS, FJHD2_ED_HS, TRQHLXZXS_QT, BTHDXS_GR_air, BTHDXS_XR_air, BTHDXS_GR_water, BTHDXS_XR_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(1))
            FHL1_GR.Add(ans_temp(2))
            FHL1_XR.Add(0)
            FHL2_GR.Add(ans_temp(3))
            FHL2_XR.Add(0)
            FHL_HS.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GR_out.Add(RFH_max)
            XR_out.Add(0)
            ZGR_out.Add(RFH_max)
            HQ_ALL.Add(ans_temp(0))
        ElseIf RFH_GR_now = 0 And RFH_XR_now > 0 Then
            '如果没有供热负荷，则直接计算
            '计算结果加入列表，直燃型溴化锂不能蓄热，结果全部是0
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_HS.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
            HQ_ALL.Add(0)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GR.Add(0)
            FHL1_XR.Add(0)
            FHL2_GR.Add(0)
            FHL2_XR.Add(0)
            FHL_HS.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GR_out.Add(0)
            XR_out.Add(0)
            ZGR_out.Add(0)
            HQ_ALL.Add(0)
        End If
        '将计算出的所有列表返回
        Dim ans(12)
        ans(0) = HD_ALL
        ans(1) = GR_out
        ans(2) = XR_out
        ans(3) = ZGR_out
        ans(4) = FHL1_GR
        ans(5) = FHL1_XR
        ans(6) = FHL2_GR
        ans(7) = FHL2_XR
        ans(8) = HQ_ALL
        ans(9) = FHL_HS
        ans(10) = NUM1_List
        ans(11) = NUM2_List
        '返回结果
        Return ans
    End Function

    Function 水_地源热泵供热和蓄热计算(ExcelApp As Object, FHTJJD As Double, RFH_GR As Double, RFH_XR As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_SDYRB
        Dim FHL2_min As Double = FHL2_min_SDYRB
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供热功率和蓄热功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GR_water * RFH_GR / (RFH_GR + RFH_XR) + BTHDXS_XR_water * RFH_XR / (RFH_GR + RFH_XR)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1 = ZJRGL1 + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2 = ZJRGL2 + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HD_ALL As New List(Of Double）
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '穷举法求各种可能的组合的总耗电功率
        '穷举设备数量
        For n1 = 0 To NUM1 Step 1 '设备1数量
            For n2 = 0 To NUM2 Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1 / NUM1
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2 / NUM2
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp) < (RFH_GR + RFH_XR) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 水_地源热泵制热COP曲线(a1)
                        Dim XZXS2 As Double = 水_地源热泵制热COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1 <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                RGL1_out_now = 0
                            Else
                                RGL1_out_now = n1 * a1 * ZJRGL1 / NUM1
                            End If
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2 <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                RGL2_out_now = 0
                            Else
                                RGL2_out_now = n2 * a2 * ZJRGL2 / NUM2
                            End If
                        End If
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR + RFH_XR Then
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
                    Next
                Next
zzzz:
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总耗电功率最低的运行工况并写入Excel
        Dim HD_ALL_min As Double
        Dim HD_index_min As Integer
        '求此时设备（1）和设备（2）的负荷率（单台）
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数（单台）
        Dim FHL1_GR_result As Double
        Dim FHL1_XR_result As Double
        Dim FHL2_GR_result As Double
        Dim FHL2_XR_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJRGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1 / (ZJRGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2 / (ZJRGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GR_result = 0
            FHL1_XR_result = 0
        Else
            FHL1_GR_result = (FHL1_result * NUM1_result / NUM1) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_1
            FHL1_XR_result = (FHL1_result * NUM1_result / NUM1) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GR_result = 0
            FHL2_XR_result = 0
        Else
            FHL2_GR_result = (FHL2_result * NUM2_result / NUM2) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_2
            FHL2_XR_result = (FHL2_result * NUM2_result / NUM2) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GR_result
        ans(2) = FHL1_XR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL2_XR_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 离心式热泵供热和蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, RFH_GR As Double, RFH_XR As Double, TJGRFHBL As Double, NUM1_LXSRB As Double, NUM2_LXSRB As Double, ZJRGL1_LXSRB As Double, ZJRGL2_LXSRB As Double, BTHD1_ED_LXSRB As Double, BTHD2_ED_LXSRB As Double, FJHD1_ED_LXSRB As Double, FJHD2_ED_LXSRB As Double, NUM1_TJ As Double, NUM2_TJ As Double, FHL1_min_TJ As Double, FHL2_min_TJ As Double, ZJRGL1_TJ As Double, ZJRGL2_TJ As Double, BTHD1_ED_TJ As Double, BTHD2_ED_TJ As Double, FJHD1_ED_TJ As Double, FJHD2_ED_TJ As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_LXSRB
        Dim FHL2_min As Double = FHL2_min_LXSRB
        '梯级设备负荷率下限（不除以设备数量）
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供热功率和蓄热功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GR_water * RFH_GR / (RFH_GR + RFH_XR) + BTHDXS_XR_water * RFH_XR / (RFH_GR + RFH_XR)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1_LXSRB = ZJRGL1_LXSRB + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2_LXSRB = ZJRGL2_LXSRB + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '全局寻优计算模式时的梯级供热计算
        '装机制热功率
        ZJRGL1_TJ = ZJRGL1_TJ + RCXS_a
        ZJRGL2_TJ = ZJRGL2_TJ + RCXS_a
        Dim ZJRGL_TJ_ALL As Double = ZJRGL1_TJ + ZJRGL2_TJ
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HD_ALL As New List(Of Double）
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '列表，储存梯级供热设备负荷率
        Dim FHL_TJGR As New List(Of Double)
        '穷举法求各种可能的组合的总耗电功率
        '穷举设备数量
        For n1 = 0 To NUM1_LXSRB Step 1 '设备1数量
            For n2 = 0 To NUM2_LXSRB Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1_LXSRB = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1_LXSRB / NUM1_LXSRB
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2_LXSRB = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2_LXSRB / NUM2_LXSRB
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp) < (RFH_GR + RFH_XR) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1_LXSRB <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            If NUM1_LXSRB = 0 Then
                                RGL1_out_now = 0
                            Else
                                RGL1_out_now = n1 * a1 * ZJRGL1_LXSRB / NUM1_LXSRB
                            End If
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2_LXSRB <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            If NUM2_LXSRB = 0 Then
                                RGL2_out_now = 0
                            Else
                                RGL2_out_now = n2 * a2 * ZJRGL2_LXSRB / NUM2_LXSRB
                            End If
                        End If
                        '加快计算
                        If (RGL1_out_now + RGL2_out_now) < (RFH_GR + RFH_XR) Then
                            GoTo qqqq
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 离心式热泵制热COP曲线(a1)
                        Dim XZXS2 As Double = 离心式热泵制热COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED_LXSRB / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED_LXSRB / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED_LXSRB
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED_LXSRB
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '此时实际的梯级供热设备总耗电
                        Dim ZHD_TJ_now As Double = 0
                        '梯级设备负荷率
                        Dim FHL_TJ_now As Double = 0
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing And TJGRFHBL > 0 Then
                            '梯级设备负荷率
                            FHL_TJ_now = TJGRFHBL * (n1 * a1 * (ZJRGL1_LXSRB - RCXS_a) / NUM1_LXSRB + n2 * a2 * (ZJRGL2_LXSRB - RCXS_a) / NUM2_LXSRB) / ((ZJRGL1_LXSRB - RCXS_a) + (ZJRGL2_LXSRB - RCXS_a))
                            '设备本体耗电修正系数
                            Dim XZXS_TJ As Double = 1
                            '定义列表，储存梯级寻优计算结果
                            '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                            Dim HD_ALL_TJ As New List(Of Double）
                            Dim FHL1_TJ As New List(Of Double）
                            Dim FHL2_TJ As New List(Of Double）
                            '列表，储存设备（1）和设备（2）的启动数量
                            Dim NUM1_TJ_List As New List(Of Double)
                            Dim NUM2_TJ_List As New List(Of Double)
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
                                    If (RFH1_ALL_TJ_temp + RFH2_ALL_TJ_temp) < ZJRGL_TJ_ALL * FHL_TJ_now Then
                                        GoTo kkk
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
                                            Dim BTHDXS_ZH_TJ As Double = (BTHDXS_GR_air * RFH_GR + BTHDXS_XR_air * RFH_XR) / (RFH_GR + RFH_XR)
                                            '计算设备（1）和设备（2）本体的耗电功率
                                            Dim BTHD1_TJ_now As Double = BTHDXS_ZH_TJ * n1_TJ * a1_TJ * BTHD1_ED_TJ / XZXS1_TJ
                                            Dim BTHD2_TJ_now As Double = BTHDXS_ZH_TJ * n2_TJ * a2_TJ * BTHD2_ED_TJ / XZXS2_TJ
                                            '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                                            Dim FJHD1_TJ_now As Double = FJHDXS * n1_TJ * a1_TJ * FJHD1_ED_TJ
                                            Dim FJHD2_TJ_now As Double = FJHDXS * n2_TJ * a2_TJ * FJHD2_ED_TJ
                                            '计算此时的总耗电功率
                                            Dim ZHD_TJ As Double = BTHD1_TJ_now + BTHD2_TJ_now + FJHD1_TJ_now + FJHD2_TJ_now
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
                                            If RGL_out_all_TJ >= ZJRGL_TJ_ALL * FHL_TJ_now Then
                                                '计算结果加入列表
                                                HD_ALL_TJ.Add(ZHD_TJ)
                                                FHL1_TJ.Add(a1_TJ)
                                                FHL2_TJ.Add(a2_TJ)
                                                NUM1_TJ_List.Add(n1_TJ)
                                                NUM2_TJ_List.Add(n2_TJ)
                                                '跳出循环
                                                Exit For
                                            ElseIf a1_TJ >= 1 And a2_TJ >= 1 Then
                                                '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                                '计算结果加入列表
                                                HD_ALL_TJ.Add(ZHD_TJ)
                                                FHL1_TJ.Add(a1_TJ)
                                                FHL2_TJ.Add(a2_TJ)
                                                NUM1_TJ_List.Add(n1_TJ)
                                                NUM2_TJ_List.Add(n2_TJ)
                                                '跳出循环
                                                Exit For
                                            End If
                                        Next
                                    Next
kkk:
                                Next
                            Next
                            '找出总耗电功率最低的运行工况并写入Excel
                            Dim HD_ALL_min_TJ As Double
                            Dim HD_index_min_TJ As Integer
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            Dim FHL1_result_TJ As Double
                            Dim FHL2_result_TJ As Double
                            '得到结果
                            HD_ALL_min_TJ = HD_ALL_TJ.Min
                            HD_index_min_TJ = HD_ALL_TJ.IndexOf(HD_ALL_min_TJ)
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            FHL1_result_TJ = FHL1_TJ(HD_index_min_TJ)
                            FHL2_result_TJ = FHL2_TJ(HD_index_min_TJ)
                            '此时实际的梯级供热设备总耗电
                            ZHD_TJ_now = HD_ALL_min_TJ
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '计算此时的总出力
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now
                        '计算此时的总耗电功率
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now + ZHD_TJ_now
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR + RFH_XR Then
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_TJGR.Add(FHL_TJ_now)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_TJGR.Add(FHL_TJ_now)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
qqqq:
                    Next
                Next
zzzz:
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总耗电功率最低的运行工况并写入Excel
        Dim HD_ALL_min As Double
        Dim HD_index_min As Integer
        '求此时设备（1）和设备（2）的负荷率（单台）
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '求此时梯级供热设备负荷率
        Dim FHL_TJGR_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数（单台）
        Dim FHL1_GR_result As Double
        Dim FHL1_XR_result As Double
        Dim FHL2_GR_result As Double
        Dim FHL2_XR_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '求此时梯级供热设备负荷率
        FHL_TJGR_result = FHL_TJGR(HD_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        '负荷率的换算还要考虑RXCS的因素
        Dim RCZHXS_1 As Double
        If ZJRGL1_LXSRB <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1_LXSRB / (ZJRGL1_LXSRB - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2_LXSRB <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2_LXSRB / (ZJRGL2_LXSRB - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1_LXSRB = 0 Then
            FHL1_GR_result = 0
            FHL1_XR_result = 0
        Else
            FHL1_GR_result = (FHL1_result * NUM1_result / NUM1_LXSRB) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_1
            FHL1_XR_result = (FHL1_result * NUM1_result / NUM1_LXSRB) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_1 * NUM1_LXSRB
        End If
        If NUM2_LXSRB = 0 Then
            FHL2_GR_result = 0
            FHL2_XR_result = 0
        Else
            FHL2_GR_result = (FHL2_result * NUM2_result / NUM2_LXSRB) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_2
            FHL2_XR_result = (FHL2_result * NUM2_result / NUM2_LXSRB) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_2 * NUM2_LXSRB
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(8) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GR_result
        ans(2) = FHL1_XR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL2_XR_result
        ans(5) = FHL_TJGR_result
        ans(6) = NUM1_result
        ans(7) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 风冷螺杆机供热和蓄热计算(ExcelApp As Object, FHTJJD As Double, RFH_GR As Double, RFH_XR As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_FLLGJ
        Dim FHL2_min As Double = FHL2_min_FLLGJ
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供热功率和蓄热功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GR_air * RFH_GR / (RFH_GR + RFH_XR) + BTHDXS_XR_air * RFH_XR / (RFH_GR + RFH_XR)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1 = ZJRGL1 + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2 = ZJRGL2 + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HD_ALL As New List(Of Double）
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '穷举法求各种可能的组合的总耗电功率
        '穷举设备数量
        For n1 = 0 To NUM1 Step 1 '设备1数量
            For n2 = 0 To NUM2 Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1 / NUM1
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2 / NUM2
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp) < (RFH_GR + RFH_XR) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 风冷螺杆式热泵制热COP曲线(a1)
                        Dim XZXS2 As Double = 风冷螺杆式热泵制热COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1 <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                RGL1_out_now = 0
                            Else
                                RGL1_out_now = n1 * a1 * ZJRGL1 / NUM1
                            End If
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2 <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                RGL2_out_now = 0
                            Else
                                RGL2_out_now = n2 * a2 * ZJRGL2 / NUM2
                            End If
                        End If
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR + RFH_XR Then
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
                    Next
                Next
zzzz:
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总耗电功率最低的运行工况并写入Excel
        Dim HD_ALL_min As Double
        Dim HD_index_min As Integer
        '求此时设备（1）和设备（2）的负荷率（单台）
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数（单台）
        Dim FHL1_GR_result As Double
        Dim FHL1_XR_result As Double
        Dim FHL2_GR_result As Double
        Dim FHL2_XR_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJRGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1 / (ZJRGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2 / (ZJRGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GR_result = 0
            FHL1_XR_result = 0
        Else
            FHL1_GR_result = (FHL1_result * NUM1_result / NUM1) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_1
            FHL1_XR_result = (FHL1_result * NUM1_result / NUM1) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GR_result = 0
            FHL2_XR_result = 0
        Else
            FHL2_GR_result = (FHL2_result * NUM2_result / NUM2) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_2
            FHL2_XR_result = (FHL2_result * NUM2_result / NUM2) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GR_result
        ans(2) = FHL1_XR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL2_XR_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 空气源热泵供热和蓄热计算(ExcelApp As Object, FHTJJD As Double, RFH_GR As Double, RFH_XR As Double, NUM1 As Double, NUM2 As Double, ZJRGL1 As Double, ZJRGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_KQYRB
        Dim FHL2_min As Double = FHL2_min_KQYRB
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供热功率和蓄热功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GR_air * RFH_GR / (RFH_GR + RFH_XR) + BTHDXS_XR_air * RFH_XR / (RFH_GR + RFH_XR)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1 = ZJRGL1 + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2 = ZJRGL2 + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HD_ALL As New List(Of Double）
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '穷举法求各种可能的组合的总耗电功率
        '穷举设备数量
        For n1 = 0 To NUM1 Step 1 '设备1数量
            For n2 = 0 To NUM2 Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1 / NUM1
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2 / NUM2
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp) < (RFH_GR + RFH_XR) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 空气源热泵制热COP曲线(a1)
                        Dim XZXS2 As Double = 空气源热泵制热COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1 <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                RGL1_out_now = 0
                            Else
                                RGL1_out_now = n1 * a1 * ZJRGL1 / NUM1
                            End If
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2 <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                RGL2_out_now = 0
                            Else
                                RGL2_out_now = n2 * a2 * ZJRGL2 / NUM2
                            End If
                        End If
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR + RFH_XR Then
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
                    Next
                Next
zzzz:
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总耗电功率最低的运行工况并写入Excel
        Dim HD_ALL_min As Double
        Dim HD_index_min As Integer
        '求此时设备（1）和设备（2）的负荷率（单台）
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数（单台）
        Dim FHL1_GR_result As Double
        Dim FHL1_XR_result As Double
        Dim FHL2_GR_result As Double
        Dim FHL2_XR_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJRGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1 / (ZJRGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2 / (ZJRGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GR_result = 0
            FHL1_XR_result = 0
        Else
            FHL1_GR_result = (FHL1_result * NUM1_result / NUM1) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_1
            FHL1_XR_result = (FHL1_result * NUM1_result / NUM1) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GR_result = 0
            FHL2_XR_result = 0
        Else
            FHL2_GR_result = (FHL2_result * NUM2_result / NUM2) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_2
            FHL2_XR_result = (FHL2_result * NUM2_result / NUM2) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GR_result
        ans(2) = FHL1_XR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL2_XR_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 电锅炉供热和蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, RFH_GR As Double, RFH_XR As Double, HSGRGLBL As Double, NUM1_DGL As Double, NUM2_DGL As Double, ZJRGL1_DGL As Double, ZJRGL2_DGL As Double, BTHD1_ED_DGL As Double, BTHD2_ED_DGL As Double, FJHD1_ED_DGL As Double, FJHD2_ED_DGL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_DGL
        Dim FHL2_min As Double = FHL2_min_DGL
        '混水设备负荷率下限（不除以设备数量）
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供热功率和蓄热功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = RFH_GR / (RFH_GR + RFH_XR) + RFH_XR / (RFH_GR + RFH_XR)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1_DGL = ZJRGL1_DGL + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2_DGL = ZJRGL2_DGL + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '全局寻优计算模式时的混水供热计算
        '装机制热功率
        ZJRGL1_HS = ZJRGL1_HS + RCXS_a
        ZJRGL2_HS = ZJRGL2_HS + RCXS_a
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HD_ALL As New List(Of Double）
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '混水设备负荷率
        Dim FHL_HS As New List(Of Double)
        '穷举设备数量
        For n1 = 0 To NUM1_DGL Step 1 '设备1数量
            For n2 = 0 To NUM2_DGL Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1_DGL = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1_DGL / NUM1_DGL
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2_DGL = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2_DGL / NUM2_DGL
                End If
                '混水供热的设备计算
                Dim HSRFH_temp As Double = 0
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                    HSRFH_temp = HSGRGLBL * (RFH1_ALL_temp + RFH2_ALL_temp) / (1 - HSGRGLBL)
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp + HSRFH_temp) < (RFH_GR + RFH_XR) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1_DGL <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            If NUM1_DGL = 0 Then
                                RGL1_out_now = 0
                            Else
                                RGL1_out_now = n1 * a1 * ZJRGL1_DGL / NUM1_DGL
                            End If
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2_DGL <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            If NUM2_DGL = 0 Then
                                RGL2_out_now = 0
                            Else
                                RGL2_out_now = n2 * a2 * ZJRGL2_DGL / NUM2_DGL
                            End If
                        End If
                        '计算此时混水设备可以提供的热负荷
                        '混水供热的设备计算
                        Dim HSGRGL_temp As Double = 0
                        HSGRGL_temp = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                        If (RGL1_out_now + RGL2_out_now + HSGRGL_temp) < (RFH_GR + RFH_XR) Then
                            GoTo qqqq
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 电采暖锅炉制热效率曲线(a1)
                        Dim XZXS2 As Double = 电采暖锅炉制热效率曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED_DGL / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED_DGL / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED_DGL
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED_DGL
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '混水供热的设备计算
                        Dim HSGRGL_now As Double = 0
                        '此时实际的设备总耗电
                        Dim HD_ALL_HS_now As Double = 0
                        '混水设备负荷率
                        Dim FHL_HS_now As Double = 0
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                            '根据HSGRGLBL，计算出此时混水设备的负荷率
                            HSGRGL_now = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                            '混水设备负荷率
                            FHL_HS_now = HSGRGL_now / ZJRGL_HS_ALL
                            '设备本体耗电修正系数
                            Dim XZXS_HS As Double = 1
                            '定义列表，储存混水寻优计算结果
                            '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                            Dim HD_ALL_HS As New List(Of Double）
                            Dim FHL1_HS As New List(Of Double）
                            Dim FHL2_HS As New List(Of Double）
                            '列表，储存设备（1）和设备（2）的启动数量
                            Dim NUM1_HS_List As New List(Of Double)
                            Dim NUM2_HS_List As New List(Of Double)
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
                                    If (RFH1_ALL_HS_temp + RFH2_ALL_HS_temp) < HSGRGL_now Then
                                        GoTo kkk
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
                                            Dim BTHDXS_ZH_HS As Double = (BTHDXS_GR_air * RFH_GR + BTHDXS_XR_air * RFH_XR) / (RFH_GR + RFH_XR)
                                            '计算设备（1）和设备（2）本体的耗电功率
                                            Dim BTHD1_HS_now As Double = BTHDXS_ZH_HS * n1_hs * a1_hs * BTHD1_ED_HS / XZXS1_HS
                                            Dim BTHD2_HS_now As Double = BTHDXS_ZH_HS * n2_hs * a2_hs * BTHD2_ED_HS / XZXS2_HS
                                            '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                                            Dim FJHD1_HS_now As Double = FJHDXS * n1_hs * a1_hs * FJHD1_ED_HS
                                            Dim FJHD2_HS_now As Double = FJHDXS * n2_hs * a2_hs * FJHD2_ED_HS
                                            '计算此时的总耗电功率
                                            Dim ZHD_HS As Double = BTHD1_HS_now + BTHD2_HS_now + FJHD1_HS_now + FJHD2_HS_now
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
                                            If RGL_out_all_HS >= HSGRGL_now Then
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            ElseIf a1_hs >= 1 And a2_hs >= 1 Then
                                                '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            End If
                                        Next
                                    Next
kkk:
                                Next
                            Next
                            '找出总耗电功率最低的运行工况并写入Excel
                            Dim HD_ALL_min_HS As Double
                            Dim HD_index_min_HS As Integer
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            Dim FHL1_result_HS As Double
                            Dim FHL2_result_HS As Double
                            '得到结果
                            HD_ALL_min_HS = HD_ALL_HS.Min
                            HD_index_min_HS = HD_ALL_HS.IndexOf(HD_ALL_min_HS)
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            FHL1_result_HS = FHL1_HS(HD_index_min_HS)
                            FHL2_result_HS = FHL2_HS(HD_index_min_HS)
                            '总耗电
                            HD_ALL_HS_now = HD_ALL_min_HS
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '总制热出力
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now + HSGRGL_now
                        '计算此时的总耗电功率
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now + HD_ALL_HS_now
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR + RFH_XR Then
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
qqqq:
                    Next
                Next
zzzz:
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总耗电功率最低的运行工况并写入Excel
        Dim HD_ALL_min As Double
        Dim HD_index_min As Integer
        '求此时设备（1）和设备（2）的负荷率（单台）
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数（单台）
        Dim FHL1_GR_result As Double
        Dim FHL1_XR_result As Double
        Dim FHL2_GR_result As Double
        Dim FHL2_XR_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，电锅炉蓄热负荷率不需要转为台数
        '负荷率的换算还要考虑RXCS的因素
        Dim RCZHXS_1 As Double
        If ZJRGL1_DGL <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1_DGL / (ZJRGL1_DGL - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2_DGL <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2_DGL / (ZJRGL2_DGL - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率不需要转换为台数（综合）
        If NUM1_DGL = 0 Then
            FHL1_GR_result = 0
            FHL1_XR_result = 0
        Else
            FHL1_GR_result = (FHL1_result * NUM1_result / NUM1_DGL) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_1
            FHL1_XR_result = (FHL1_result * NUM1_result / NUM1_DGL) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_1
        End If
        If NUM2_DGL = 0 Then
            FHL2_GR_result = 0
            FHL2_XR_result = 0
        Else
            FHL2_GR_result = (FHL2_result * NUM2_result / NUM2_DGL) * RFH_GR / (RFH_GR + RFH_XR) * RCZHXS_2
            FHL2_XR_result = (FHL2_result * NUM2_result / NUM2_DGL) * RFH_XR / (RFH_GR + RFH_XR) * RCZHXS_2
        End If
        '混水设备负荷率计算结果
        Dim FHL_HS_result As Double = FHL_HS(HD_index_min)
        '返回结果存到数组
        Dim ans(8) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GR_result
        ans(2) = FHL1_XR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL2_XR_result
        ans(5) = FHL_HS_result
        ans(6) = NUM1_result
        ans(7) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 天然气锅炉供热计算_方法二(ExcelApp As Object, b As Integer, FHTJJD As Double, RFH_GR As Double, D_price As Double, TRQ_price As Double, HSGRGLBL As Double, NUM1_TRQGL As Double, NUM2_TRQGL As Double, ZJRGL1_TRQGL As Double, ZJRGL2_TRQGL As Double, BTHQ1_ED_TRQGL As Double, BTHQ2_ED_TRQGL As Double, FJHD1_ED_TRQGL As Double, FJHD2_ED_TRQGL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_TRQGL
        Dim FHL2_min As Double = FHL2_min_TRQGL
        '混水设备负荷率下限（不除以设备数量）
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1_TRQGL = ZJRGL1_TRQGL + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2_TRQGL = ZJRGL2_TRQGL + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '全局寻优计算模式时的混水供热计算
        '装机制热功率
        ZJRGL1_HS = ZJRGL1_HS + RCXS_a
        ZJRGL2_HS = ZJRGL2_HS + RCXS_a
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '总成本
        Dim COST_ALL As New List(Of Double)
        '列表，储存总耗气量，耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HQ_ALL As New List(Of Double）
        Dim HD_ALL As New List(Of Double)
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '混水设备负荷率
        Dim FHL_HS As New List(Of Double)
        '穷举设备数量
        For n1 = 0 To NUM1_TRQGL Step 1 '设备1数量
            For n2 = 0 To NUM2_TRQGL Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1_TRQGL = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1_TRQGL / NUM1_TRQGL
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2_TRQGL = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2_TRQGL / NUM2_TRQGL
                End If
                '混水供热的设备计算
                Dim HSRFH_temp As Double = 0
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                    HSRFH_temp = HSGRGLBL * (RFH1_ALL_temp + RFH2_ALL_temp) / (1 - HSGRGLBL)
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp + HSRFH_temp) < (RFH_GR + 0) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                '穷举法求各种可能的组合的总耗电功率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1_TRQGL <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            RGL1_out_now = a1 * ZJRGL1_TRQGL
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2_TRQGL <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            RGL2_out_now = a2 * ZJRGL2_TRQGL
                        End If
                        '计算此时混水设备可以提供的热负荷
                        '混水供热的设备计算
                        Dim HSGRGL_temp As Double = 0
                        HSGRGL_temp = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                        If (RGL1_out_now + RGL2_out_now + HSGRGL_temp) < (RFH_GR + 0) Then
                            GoTo qqqq
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 天然气采暖锅炉制热效率曲线(a1)
                        Dim XZXS2 As Double = 天然气采暖锅炉制热效率曲线(a2)
                        '计算设备（1）和设备（2）本体的耗气量
                        Dim BTHQ1_now As Double = TRQHLXZXS_QT * a1 * BTHQ1_ED_TRQGL / XZXS1
                        Dim BTHQ2_now As Double = TRQHLXZXS_QT * a2 * BTHQ2_ED_TRQGL / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * a1 * FJHD1_ED_TRQGL
                        Dim FJHD2_now As Double = FJHDXS * a2 * FJHD2_ED_TRQGL
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '混水供热的设备计算
                        Dim HSGRGL_now As Double = 0
                        '此时实际的设备总耗电
                        Dim HD_ALL_HS_now As Double = 0
                        '混水设备负荷率
                        Dim FHL_HS_now As Double = 0
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                            '根据HSGRGLBL，计算出此时混水设备的负荷率
                            HSGRGL_now = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                            '混水设备负荷率
                            FHL_HS_now = HSGRGL_now / ZJRGL_HS_ALL
                            '设备本体耗电修正系数
                            Dim XZXS_HS As Double = 1
                            '定义列表，储存混水寻优计算结果
                            '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                            Dim HD_ALL_HS As New List(Of Double）
                            Dim FHL1_HS As New List(Of Double）
                            Dim FHL2_HS As New List(Of Double）
                            '列表，储存设备（1）和设备（2）的启动数量
                            Dim NUM1_HS_List As New List(Of Double)
                            Dim NUM2_HS_List As New List(Of Double)
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
                                    If (RFH1_ALL_HS_temp + RFH2_ALL_HS_temp) < HSGRGL_now Then
                                        GoTo kkk
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
                                            Dim BTHDXS_ZH_HS As Double = BTHDXS_GR_air
                                            '计算设备（1）和设备（2）本体的耗电功率
                                            Dim BTHD1_HS_now As Double = BTHDXS_ZH_HS * n1_hs * a1_hs * BTHD1_ED_HS / XZXS1_HS
                                            Dim BTHD2_HS_now As Double = BTHDXS_ZH_HS * n2_hs * a2_hs * BTHD2_ED_HS / XZXS2_HS
                                            '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                                            Dim FJHD1_HS_now As Double = FJHDXS * n1_hs * a1_hs * FJHD1_ED_HS
                                            Dim FJHD2_HS_now As Double = FJHDXS * n2_hs * a2_hs * FJHD2_ED_HS
                                            '计算此时的总耗电功率
                                            Dim ZHD_HS As Double = BTHD1_HS_now + BTHD2_HS_now + FJHD1_HS_now + FJHD2_HS_now
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
                                            If RGL_out_all_HS >= HSGRGL_now Then
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            ElseIf a1_hs >= 1 And a2_hs >= 1 Then
                                                '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            End If
                                        Next
                                    Next
kkk:
                                Next
                            Next
                            '找出总耗电功率最低的运行工况并写入Excel
                            Dim HD_ALL_min_HS As Double
                            Dim HD_index_min_HS As Integer
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            Dim FHL1_result_HS As Double
                            Dim FHL2_result_HS As Double
                            '得到结果
                            HD_ALL_min_HS = HD_ALL_HS.Min
                            HD_index_min_HS = HD_ALL_HS.IndexOf(HD_ALL_min_HS)
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            FHL1_result_HS = FHL1_HS(HD_index_min_HS)
                            FHL2_result_HS = FHL2_HS(HD_index_min_HS)
                            '总耗电
                            HD_ALL_HS_now = HD_ALL_min_HS
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '总制热出力
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now + HSGRGL_now
                        '计算此时的总耗气量
                        Dim ZHQ As Double = BTHQ1_now + BTHQ2_now
                        '计算此时的总耗电量
                        Dim ZHD As Double = FJHD1_now + FJHD2_now + FJHD1_now + FJHD2_now + HD_ALL_HS_now
                        '计算总成本
                        Dim COST_ALL_now As Double = D_price * ZHD + TRQ_price * ZHQ
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR Then
                            '计算结果加入列表
                            HQ_ALL.Add(ZHQ)
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            COST_ALL.Add(COST_ALL_now)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            COST_ALL.Add(COST_ALL_now)
                            '跳出循环
                            Exit For
                        End If
qqqq:
                    Next
                Next
zzzz：
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总成本量最低的运行工况
        Dim COST_ALL_min As Double
        Dim COST_index_min As Integer
        '求此时的天然气耗量
        Dim HQ_result As Double
        '求此时辅机耗电总功率
        Dim HD_result As Double
        '求此时设备（1）和设备（2）的负荷率
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        Dim FHL1_GR_result As Double
        Dim FHL2_GR_result As Double
        '得到结果
        COST_ALL_min = COST_ALL.Min
        COST_index_min = COST_ALL.IndexOf(COST_ALL_min)
        '求此时的天然气耗量
        HQ_result = HQ_ALL(COST_index_min)
        '求此时辅机耗电总功率
        HD_result = HD_ALL(COST_index_min)
        '求此时设备（1）和设备（2）的负荷率
        FHL1_result = FHL1(COST_index_min)
        FHL2_result = FHL2(COST_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(COST_index_min)
        NUM2_result = NUM2_List(COST_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配
        '负荷率的换算还要考虑RXCS的因素
        Dim RCZHXS_1 As Double
        If ZJRGL1_TRQGL <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1_TRQGL / (ZJRGL1_TRQGL - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2_TRQGL <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2_TRQGL / (ZJRGL2_TRQGL - RCXS_a)）
        End If
        FHL1_GR_result = FHL1_result * RCZHXS_1
        FHL2_GR_result = FHL2_result * RCZHXS_2
        '混水设备负荷率计算结果
        Dim FHL_HS_result As Double = FHL_HS(COST_index_min)
        '返回结果存到数组
        Dim ans(5) As Double
        ans(0) = HQ_result
        ans(1) = HD_result
        ans(2) = FHL1_GR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL_HS_result
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂供热计算_方法二(ExcelApp As Object, b As Integer, FHTJJD As Double, RFH_GR As Double, D_price As Double, TRQ_price As Double, HSGRGLBL As Double, NUM1_ZRXXHL As Double, NUM2_ZRXXHL As Double, ZJRGL1_ZRXXHL As Double, ZJRGL2_ZRXXHL As Double, BTHQ1_ED_ZRXXHL As Double, BTHQ2_ED_ZRXXHL As Double, FJHD1_ED_ZRXXHL As Double, FJHD2_ED_ZRXXHL As Double, NUM1_HS As Double, NUM2_HS As Double, FHL1_min_HS As Double, FHL2_min_HS As Double, ZJRGL1_HS As Double, ZJRGL2_HS As Double, BTHD1_ED_HS As Double, BTHD2_ED_HS As Double, FJHD1_ED_HS As Double, FJHD2_ED_HS As Double, TRQHLXZXS_QT As Double, BTHDXS_GR_air As Double, BTHDXS_XR_air As Double, BTHDXS_GR_water As Double, BTHDXS_XR_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_ZRXXHL
        Dim FHL2_min As Double = FHL2_min_ZRXXHL
        '混水设备负荷率下限（不除以设备数量）
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL1_ZRXXHL = ZJRGL1_ZRXXHL + RCXS_a
        '设备（2）装机功率（总和），在读取的值的基础上放大一些
        ZJRGL2_ZRXXHL = ZJRGL2_ZRXXHL + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '全局寻优计算模式时的混水供热计算
        '装机制热功率
        ZJRGL1_HS = ZJRGL1_HS + RCXS_a
        ZJRGL2_HS = ZJRGL2_HS + RCXS_a
        Dim ZJRGL_HS_ALL As Double = ZJRGL1_HS + ZJRGL2_HS
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '总成本
        Dim COST_ALL As New List(Of Double)
        '列表，储存总耗气量，耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HQ_ALL As New List(Of Double）
        Dim HD_ALL As New List(Of Double)
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '混水设备负荷率
        Dim FHL_HS As New List(Of Double)
        '穷举设备数量
        For n1 = 0 To NUM1_ZRXXHL Step 1 '设备1数量
            For n2 = 0 To NUM2_ZRXXHL Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1_ZRXXHL = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZJRGL1_ZRXXHL / NUM1_ZRXXHL
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2_ZRXXHL = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZJRGL2_ZRXXHL / NUM2_ZRXXHL
                End If
                '混水供热的设备计算
                Dim HSRFH_temp As Double = 0
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                    HSRFH_temp = HSGRGLBL * (RFH1_ALL_temp + RFH2_ALL_temp) / (1 - HSGRGLBL)
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp + HSRFH_temp) < (RFH_GR + 0) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                '穷举法求各种可能的组合的总耗电功率
                '直燃型溴化锂制热负荷率上限为1
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率
                        '计算此时的总出力
                        Dim RGL1_out_now As Double
                        If ZJRGL1_ZRXXHL <= RCXS_a Then
                            RGL1_out_now = 0
                        Else
                            RGL1_out_now = a1 * ZJRGL1_ZRXXHL
                        End If
                        Dim RGL2_out_now As Double
                        If ZJRGL2_ZRXXHL <= RCXS_a Then
                            RGL2_out_now = 0
                        Else
                            RGL2_out_now = a2 * ZJRGL2_ZRXXHL
                        End If
                        '计算此时混水设备可以提供的热负荷
                        '混水供热的设备计算
                        Dim HSGRGL_temp As Double = 0
                        HSGRGL_temp = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                        If (RGL1_out_now + RGL2_out_now + HSGRGL_temp) < (RFH_GR + 0) Then
                            GoTo qqqq
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '计算设备（1）和设备（2）本体的效率修正系数
                        Dim XZXS1 As Double = 直燃型溴化锂制热COP曲线(a1)
                        Dim XZXS2 As Double = 直燃型溴化锂制热COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗气量
                        Dim BTHQ1_now As Double = TRQHLXZXS_QT * a1 * BTHQ1_ED_ZRXXHL / XZXS1
                        Dim BTHQ2_now As Double = TRQHLXZXS_QT * a2 * BTHQ2_ED_ZRXXHL / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                        Dim FJHD1_now As Double = FJHDXS * a1 * FJHD1_ED_ZRXXHL
                        Dim FJHD2_now As Double = FJHDXS * a2 * FJHD2_ED_ZRXXHL
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '混水供热的设备计算
                        Dim HSGRGL_now As Double = 0
                        '此时实际的设备总耗电
                        Dim HD_ALL_HS_now As Double = 0
                        '混水设备负荷率
                        Dim FHL_HS_now As Double = 0
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And HSGRGLBL > 0 Then
                            '根据HSGRGLBL，计算出此时混水设备的负荷率
                            HSGRGL_now = HSGRGLBL * (RGL1_out_now + RGL2_out_now) / (1 - HSGRGLBL)
                            '混水设备负荷率
                            FHL_HS_now = HSGRGL_now / ZJRGL_HS_ALL
                            '设备本体耗电修正系数
                            Dim XZXS_HS As Double = 1
                            '定义列表，储存混水寻优计算结果
                            '列表，储存总耗电量，设备（1）负荷率和设备（2）负荷率
                            Dim HD_ALL_HS As New List(Of Double）
                            Dim FHL1_HS As New List(Of Double）
                            Dim FHL2_HS As New List(Of Double）
                            '列表，储存设备（1）和设备（2）的启动数量
                            Dim NUM1_HS_List As New List(Of Double)
                            Dim NUM2_HS_List As New List(Of Double)
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
                                    If (RFH1_ALL_HS_temp + RFH2_ALL_HS_temp) < HSGRGL_now Then
                                        GoTo kkk
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
                                            Dim BTHDXS_ZH_HS As Double = BTHDXS_GR_air
                                            '计算设备（1）和设备（2）本体的耗电功率
                                            Dim BTHD1_HS_now As Double = BTHDXS_ZH_HS * n1_hs * a1_hs * BTHD1_ED_HS / XZXS1_HS
                                            Dim BTHD2_HS_now As Double = BTHDXS_ZH_HS * n2_hs * a2_hs * BTHD2_ED_HS / XZXS2_HS
                                            '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）
                                            Dim FJHD1_HS_now As Double = FJHDXS * n1_hs * a1_hs * FJHD1_ED_HS
                                            Dim FJHD2_HS_now As Double = FJHDXS * n2_hs * a2_hs * FJHD2_ED_HS
                                            '计算此时的总耗电功率
                                            Dim ZHD_HS As Double = BTHD1_HS_now + BTHD2_HS_now + FJHD1_HS_now + FJHD2_HS_now
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
                                            If RGL_out_all_HS >= HSGRGL_now Then
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            ElseIf a1_hs >= 1 And a2_hs >= 1 Then
                                                '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                                                '计算结果加入列表
                                                HD_ALL_HS.Add(ZHD_HS)
                                                FHL1_HS.Add(a1_hs)
                                                FHL2_HS.Add(a2_hs)
                                                NUM1_HS_List.Add(n1_hs)
                                                NUM2_HS_List.Add(n2_hs)
                                                '跳出循环
                                                Exit For
                                            End If
                                        Next
                                    Next
kkk:
                                Next
                            Next
                            '找出总耗电功率最低的运行工况并写入Excel
                            Dim HD_ALL_min_HS As Double
                            Dim HD_index_min_HS As Integer
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            Dim FHL1_result_HS As Double
                            Dim FHL2_result_HS As Double
                            '得到结果
                            HD_ALL_min_HS = HD_ALL_HS.Min
                            HD_index_min_HS = HD_ALL_HS.IndexOf(HD_ALL_min_HS)
                            '求此时设备（1）和设备（2）的负荷率（单台）
                            FHL1_result_HS = FHL1_HS(HD_index_min_HS)
                            FHL2_result_HS = FHL2_HS(HD_index_min_HS)
                            '总耗电
                            HD_ALL_HS_now = HD_ALL_min_HS
                        End If
                        '————————————————————————————————————————————————————————————————————————————————————————
                        '总制热出力
                        Dim RGL_out_all As Double = RGL1_out_now + RGL2_out_now + HSGRGL_now
                        '计算此时的总耗气量
                        Dim ZHQ As Double = BTHQ1_now + BTHQ2_now
                        '计算此时的总耗电量
                        Dim ZHD As Double = FJHD1_now + FJHD2_now + FJHD1_now + FJHD2_now + HD_ALL_HS_now
                        '计算总成本
                        Dim COST_ALL_now As Double = D_price * ZHD + TRQ_price * ZHQ
                        '如果达到了热负荷需求，则跳出内层循环
                        If RGL_out_all >= RFH_GR Then
                            '计算结果加入列表
                            HQ_ALL.Add(ZHQ)
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            COST_ALL.Add(COST_ALL_now)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            FHL_HS.Add(FHL_HS_now)
                            COST_ALL.Add(COST_ALL_now)
                            '跳出循环
                            Exit For
                        End If
qqqq:
                    Next
                Next
zzzz：
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出总成本量最低的运行工况
        Dim COST_ALL_min As Double
        Dim COST_index_min As Integer
        '求此时的天然气耗量
        Dim HQ_result As Double
        '求此时辅机耗电总功率
        Dim HD_result As Double
        '求此时设备（1）和设备（2）的负荷率
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配，蓄热负荷率转换为台数
        Dim FHL1_GR_result As Double
        Dim FHL2_GR_result As Double
        '得到结果
        COST_ALL_min = COST_ALL.Min
        COST_index_min = COST_ALL.IndexOf(COST_ALL_min)
        '求此时的天然气耗量
        HQ_result = HQ_ALL(COST_index_min)
        '求此时辅机耗电总功率
        HD_result = HD_ALL(COST_index_min)
        '求此时设备（1）和设备（2）的负荷率
        FHL1_result = FHL1(COST_index_min)
        FHL2_result = FHL2(COST_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(COST_index_min)
        NUM2_result = NUM2_List(COST_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供热和蓄热功率比例进行分配
        '负荷率的换算还要考虑RXCS的因素
        Dim RCZHXS_1 As Double
        If ZJRGL1_ZRXXHL <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJRGL1_ZRXXHL / (ZJRGL1_ZRXXHL - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJRGL2_ZRXXHL <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJRGL2_ZRXXHL / (ZJRGL2_ZRXXHL - RCXS_a)）
        End If
        FHL1_GR_result = FHL1_result * RCZHXS_1
        FHL2_GR_result = FHL2_result * RCZHXS_2
        '混水设备负荷率计算结果
        Dim FHL_HS_result As Double = FHL_HS(COST_index_min)
        '返回结果存到数组
        Dim ans(5) As Double
        ans(0) = HQ_result
        ans(1) = HD_result
        ans(2) = FHL1_GR_result
        ans(3) = FHL2_GR_result
        ans(4) = FHL_HS_result
        '返回结果
        Return ans
    End Function

End Module
