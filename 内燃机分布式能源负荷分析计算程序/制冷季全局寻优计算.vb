Module 制冷季全局寻优计算
    Sub 制冷季设备制冷和蓄冷全局寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, XHLZL As Double, XHLXL As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer)
        '冷负荷总量=供冷+蓄冷
        Dim LFH_GL_now As Double = LFHZXQL(b) - (XNGLGL(b) + XHLZL)
        Dim LFH_XL_now As Double = XNXLGL(b) - XHLXL
        '如果当前冷负荷总需求量=0，则将LFH_GL ,强制设置为0，防止出错
        If LFHZXQL(b) = 0 Then
            LFH_GL_now = 0
        End If
        '如果当前蓄冷负荷总需求量=0，则将LFH_XL ,强制设置为0，防止出错
        If XNXLGL(b) = 0 Then
            LFH_XL_now = 0
        End If
        Dim LFH_ALL As Double = LFH_GL_now + LFH_XL_now
        '负荷调整系数（全局寻优时候能否计算到了负荷上限的倍数）
        Dim TZXS As Double = 1 + 10 * FHTJJD / 100
        '误差系数，允许误差的最大比例
        Dim WCXS As Double = 1 + 4 * FHTJJD / 100
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '根据选择的设备类型，获取被选择的设备类型，由于采用全局寻优计算方法，不需要考虑设备选择顺序
        '设置变量判断某一种设备是否被启用，0表示没有启用，1表示启用
        '离心式冷水机
        Dim ZJJC_LXSLSJ As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "离心式冷水机" Then
                ZJJC_LXSLSJ = 1
                Exit For
            End If
        Next
        '水冷螺杆机
        Dim ZJJC_SLLGJ As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "水冷螺杆机" Then
                ZJJC_SLLGJ = 1
                Exit For
            End If
        Next
        '风冷螺杆机
        Dim ZJJC_FLLGJ As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "风冷螺杆机" Then
                ZJJC_FLLGJ = 1
                Exit For
            End If
        Next
        '水（地）源热泵
        Dim ZJJC_SDYRB As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "水(地)源热泵" Then
                ZJJC_SDYRB = 1
                Exit For
            End If
        Next
        '离心式热泵
        Dim ZJJC_LXSRB As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "离心式热泵" Then
                ZJJC_LXSRB = 1
                Exit For
            End If
        Next
        '空气源热泵
        Dim ZJJC_KQYRB As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "空气源热泵" Then
                ZJJC_KQYRB = 1
                Exit For
            End If
        Next
        '直燃型溴化锂
        Dim ZJJC_ZRXXHL As Integer = 0
        '遍历顺序1到6
        For i = 6 To 11
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "直燃型溴化锂" Then
                ZJJC_ZRXXHL = 1
                Exit For
            End If
        Next
        '装机监测情况汇总
        Dim ZJJC_ALL As Integer = ZJJC_FLLGJ + ZJJC_KQYRB + ZJJC_LXSLSJ + ZJJC_LXSRB + ZJJC_SDYRB + ZJJC_SLLGJ + ZJJC_ZRXXHL
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '根据选择的计算模式类型，只有是模式2的时候才会计算
        If calculation_mode = 2 And LFH_GL_now + LFH_XL_now > 0 And ZJJC_ALL > 0 Then
            '读取制冷季装机方案及参数
            Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
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
            Dim TRQZLHQXZ As Double = ans_XZXS_L(5)
            Dim TRQHLXZXS_QT As Double = ans_XZXS_L(6)
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '各个设备可以装机功率（总和）
            '离心式冷水机            
            '装机总功率
            Dim ZJZGL_LXSLSJ = 0
            If ZJJC_LXSLSJ = 1 Then
                ZJZGL_LXSLSJ = ZJLGL1_LXSLSJ + ZJLGL2_LXSLSJ
            Else
                ZJZGL_LXSLSJ = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_LXSLSJ As Double = 0
            Dim FH1_min_LXSLSJ As Double = 0
            Dim FH2_min_LXSLSJ As Double = 0
            If ZJJC_LXSLSJ = 1 Then
                If NUM1_LXSLSJ > 0 Then
                    FH1_min_LXSLSJ = ZJLGL1_LXSLSJ * FHL1_min_LXSLSJ / NUM1_LXSLSJ
                Else
                    FH1_min_LXSLSJ = 0
                End If
                If NUM2_LXSLSJ > 0 Then
                    FH2_min_LXSLSJ = ZJLGL2_LXSLSJ * FHL2_min_LXSLSJ / NUM2_LXSLSJ
                Else
                    FH2_min_LXSLSJ = 0
                End If
                FH_min_LXSLSJ = FH1_min_LXSLSJ + FH2_min_LXSLSJ
            Else
                FH_min_LXSLSJ = 0
            End If
            '水冷螺杆机
            '装机总功率
            Dim ZJZGL_SLLGJ As Double = 0
            If ZJJC_SLLGJ = 1 Then
                ZJZGL_SLLGJ = ZJLGL1_SLLGJ + ZJLGL2_SLLGJ
            Else
                ZJZGL_SLLGJ = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_SLLGJ As Double = 0
            Dim FH1_min_SLLGJ As Double = 0
            Dim FH2_min_SLLGJ As Double = 0
            If ZJJC_SLLGJ = 1 Then
                If NUM1_SLLGJ > 0 Then
                    FH1_min_SLLGJ = ZJLGL1_SLLGJ * FHL1_min_SLLGJ / NUM1_SLLGJ
                Else
                    FH1_min_SLLGJ = 0
                End If
                If NUM2_SLLGJ > 0 Then
                    FH2_min_SLLGJ = ZJLGL2_SLLGJ * FHL2_min_SLLGJ / NUM2_SLLGJ
                Else
                    FH2_min_SLLGJ = 0
                End If
                FH_min_SLLGJ = FH1_min_SLLGJ + FH2_min_SLLGJ
            Else
                FH_min_SLLGJ = 0
            End If
            '水（地）源热泵
            '装机总功率（总和）
            Dim ZJZGL_SDYRB As Double = 0
            If ZJJC_SDYRB = 1 Then
                ZJZGL_SDYRB = ZJLGL1_SDYRB + ZJLGL2_SDYRB
            Else
                ZJZGL_SDYRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_SDYRB As Double = 0
            Dim FH1_min_SDYRB As Double = 0
            Dim FH2_min_SDYRB As Double = 0
            If ZJJC_SDYRB = 1 Then
                If NUM1_SDYRB > 0 Then
                    FH1_min_SDYRB = ZJLGL1_SDYRB * FHL1_min_SDYRB / NUM1_SDYRB
                Else
                    FH1_min_SDYRB = 0
                End If
                If NUM2_SDYRB > 0 Then
                    FH2_min_SDYRB = ZJLGL2_SDYRB * FHL2_min_SDYRB / NUM2_SDYRB
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
                ZJZGL_LXSRB = ZJLGL1_LXSRB + ZJLGL2_LXSRB
            Else
                ZJZGL_LXSRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_LXSRB As Double = 0
            Dim FH1_min_LXSRB As Double = 0
            Dim FH2_min_LXSRB As Double = 0
            If ZJJC_LXSRB = 1 Then
                If NUM1_LXSRB > 0 Then
                    FH1_min_LXSRB = ZJLGL1_LXSRB * FHL1_min_LXSRB / NUM1_LXSRB
                Else
                    FH1_min_LXSRB = 0
                End If
                If NUM2_LXSRB > 0 Then
                    FH2_min_LXSRB = ZJLGL2_LXSRB * FHL2_min_LXSRB / NUM2_LXSRB
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
                ZJZGL_FLLGJ = ZJLGL1_FLLGJ + ZJLGL2_FLLGJ
            Else
                ZJZGL_FLLGJ = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_FLLGJ As Double = 0
            Dim FH1_min_FLLGJ As Double = 0
            Dim FH2_min_FLLGJ As Double = 0
            If ZJJC_FLLGJ = 1 Then
                If NUM1_FLLGJ > 0 Then
                    FH1_min_FLLGJ = ZJLGL1_FLLGJ * FHL1_min_FLLGJ / NUM1_FLLGJ
                Else
                    FH1_min_FLLGJ = 0
                End If
                If NUM2_FLLGJ > 0 Then
                    FH2_min_FLLGJ = ZJLGL2_FLLGJ * FHL2_min_FLLGJ / NUM2_FLLGJ
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
                ZJZGL_KQYRB = ZJLGL1_KQYRB + ZJLGL2_KQYRB
            Else
                ZJZGL_KQYRB = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_KQYRB As Double = 0
            Dim FH1_min_KQYRB As Double = 0
            Dim FH2_min_KQYRB As Double = 0
            If ZJJC_KQYRB = 1 Then
                If NUM1_KQYRB > 0 Then
                    FH1_min_KQYRB = ZJLGL1_KQYRB * FHL1_min_KQYRB / NUM1_KQYRB
                Else
                    FH1_min_KQYRB = 0
                End If
                If NUM2_KQYRB > 0 Then
                    FH2_min_KQYRB = ZJLGL2_KQYRB * FHL2_min_KQYRB / NUM2_KQYRB
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
                '直燃型溴化锂制冷允许超发到120%
                ZJZGL_ZRXXHL = (ZJLGL1_ZRXXHL + ZJLGL2_ZRXXHL) * 1.2
            Else
                ZJZGL_ZRXXHL = 0
            End If
            '允许运行最低负荷（要除以设备装机数量）
            Dim FH_min_ZRXXHL As Double = 0
            Dim FH1_min_ZRXXHL As Double = 0
            Dim FH2_min_ZRXXHL As Double = 0
            If ZJJC_ZRXXHL = 1 Then
                If NUM1_ZRXXHL > 0 Then
                    FH1_min_ZRXXHL = ZJLGL1_ZRXXHL * FHL1_min_ZRXXHL / NUM1_ZRXXHL
                Else
                    FH1_min_ZRXXHL = 0
                End If
                If NUM2_ZRXXHL > 0 Then
                    FH2_min_ZRXXHL = ZJLGL2_ZRXXHL * FHL2_min_ZRXXHL / NUM2_ZRXXHL
                Else
                    FH2_min_ZRXXHL = 0
                End If
                FH_min_ZRXXHL = FH1_min_ZRXXHL + FH2_min_ZRXXHL
            Else
                FH_min_ZRXXHL = 0
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '如果此时制冷装机总功率=供冷功率+蓄冷功率，装机功率会被全部用掉，则需要增加此时的计算容错性
            '溴化锂装机数量
            Dim NUM1_XHL As Double = ans_ZJFA_L(0)
            Dim NUM2_XHL As Double = ans_ZJFA_L(1)
            '内燃机余热功率
            Dim YRGL1_ED_NRJ As Double = ans_ZJFA_L(4)
            Dim YRGL2_ED_NRJ As Double = ans_ZJFA_L(5)
            '溴化锂制冷COP
            Dim XHL_COP_L As Double = ans_XZXS_L(10)
            '计算溴化锂制热功率
            Dim ZJZGL_XHL As Double = (NUM1_XHL * YRGL1_ED_NRJ + NUM2_XHL * YRGL2_ED_NRJ) * XHL_COP_L
            '制冷装机功率求和
            Dim ZJLGL_ALL As Double = 0
            '如果处于不启动内燃机的时间段
            If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(24, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(25, 9).Value Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 79).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(26, 9).Value) Then
                ZJLGL_ALL = ZJZGL_LXSLSJ + ZJZGL_SLLGJ + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL
            Else
                ZJLGL_ALL = ZJZGL_LXSLSJ + ZJZGL_SLLGJ + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL + ZJZGL_XHL
            End If
            '如果此时的冷装机会被全部用掉，将工况序号记录下来，直接采用常规计算模式进行计算，不报错
            If Math.Abs(LFHZXQL(b) - XNGLGL(b) + XNXLGL(b) - ZJLGL_ALL) <= RCXS Then
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
                'MsgBox("供冷和蓄冷全局寻优计算出错，计算结束！！" & “当前正在计算的工况序号为：   ” & b)
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '定义列表，储存各个设备计算出的结果
            '离心式冷水机计算出的各种结果存入列表
            Dim HD_ALL_min_LXSLSJ As New List(Of Double)
            Dim GL_ALL_LXSLSJ As New List(Of Double)
            Dim XL_ALL_LXSLSJ As New List(Of Double)
            Dim FHL1_GL_LXSLSJ As New List(Of Double)
            Dim FHL1_XL_LXSLSJ As New List(Of Double)
            Dim FHL2_GL_LXSLSJ As New List(Of Double)
            Dim FHL2_XL_LXSLSJ As New List(Of Double)
            '水冷螺杆机
            Dim HD_ALL_min_SLLGJ As New List(Of Double)
            Dim GL_ALL_SLLGJ As New List(Of Double)
            Dim XL_ALL_SLLGJ As New List(Of Double)
            Dim FHL1_GL_SLLGJ As New List(Of Double)
            Dim FHL1_XL_SLLGJ As New List(Of Double)
            Dim FHL2_GL_SLLGJ As New List(Of Double)
            Dim FHL2_XL_SLLGJ As New List(Of Double)
            '水（地）源热泵
            Dim HD_ALL_min_SDYRB As New List(Of Double)
            Dim GL_ALL_SDYRB As New List(Of Double)
            Dim XL_ALL_SDYRB As New List(Of Double)
            Dim FHL1_GL_SDYRB As New List(Of Double)
            Dim FHL1_XL_SDYRB As New List(Of Double)
            Dim FHL2_GL_SDYRB As New List(Of Double)
            Dim FHL2_XL_SDYRB As New List(Of Double)
            '离心式热泵
            Dim HD_ALL_min_LXSRB As New List(Of Double)
            Dim GL_ALL_LXSRB As New List(Of Double)
            Dim XL_ALL_LXSRB As New List(Of Double)
            Dim FHL1_GL_LXSRB As New List(Of Double)
            Dim FHL1_XL_LXSRB As New List(Of Double)
            Dim FHL2_GL_LXSRB As New List(Of Double)
            Dim FHL2_XL_LXSRB As New List(Of Double)
            '风冷螺杆机
            Dim HD_ALL_min_FLLGJ As New List(Of Double)
            Dim GL_ALL_FLLGJ As New List(Of Double)
            Dim XL_ALL_FLLGJ As New List(Of Double)
            Dim FHL1_GL_FLLGJ As New List(Of Double)
            Dim FHL1_XL_FLLGJ As New List(Of Double)
            Dim FHL2_GL_FLLGJ As New List(Of Double)
            Dim FHL2_XL_FLLGJ As New List(Of Double)
            '空气源热泵
            Dim HD_ALL_min_KQYRB As New List(Of Double)
            Dim GL_ALL_KQYRB As New List(Of Double)
            Dim XL_ALL_KQYRB As New List(Of Double)
            Dim FHL1_GL_KQYRB As New List(Of Double)
            Dim FHL1_XL_KQYRB As New List(Of Double)
            Dim FHL2_GL_KQYRB As New List(Of Double)
            Dim FHL2_XL_KQYRB As New List(Of Double)
            '直燃型溴化锂
            Dim HQ_ALL_min_ZRXXHL As New List(Of Double)
            Dim HD_ALL_ZRXXHL As New List(Of Double)
            Dim GL_ALL_ZRXXHL As New List(Of Double)
            Dim XL_ALL_ZRXXHL As New List(Of Double)
            Dim FHL1_GL_ZRXXHL As New List(Of Double)
            Dim FHL1_XL_ZRXXHL As New List(Of Double)
            Dim FHL2_GL_ZRXXHL As New List(Of Double)
            Dim FHL2_XL_ZRXXHL As New List(Of Double)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '供冷和蓄冷一起寻优
            '寻优先后顺序：离心式冷水机、水冷螺杆机、水（地）源热泵、离心式热泵、风冷螺杆机、空气源热泵、直燃型溴化锂
            '穷举计算
            '各种设备已经计算的次数计数（如果设备不存在或者装机量为0，才参与计算）
            Dim JS_end_LXSLSJ As Integer = 0
            '将总负荷分为供冷功率和蓄冷功率,进行计算
            For a_1 = 0 To ZJZGL_LXSLSJ Step ZJZGL_LXSLSJ / FHFPCS
                '直接结束计算
                If JS_end_LXSLSJ >= 1 Then
                    Exit For
                End If
                '如果不存在装机或者装机功率等于0
                If ZJJC_LXSLSJ = 0 Or ZJZGL_LXSLSJ = 0 Then
                    JS_end_LXSLSJ = JS_end_LXSLSJ + 1
                End If
                '根据装机量判断是否直接进入下一次循环
                If a_1 + ZJZGL_SLLGJ + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL < LFH_ALL Then
                    GoTo aaa
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                Dim JS_end_SLLGJ As Integer = 0
                For a_2 = 0 To ZJZGL_SLLGJ Step ZJZGL_SLLGJ / FHFPCS
                    '直接结束计算
                    If JS_end_SLLGJ >= 1 Then
                        Exit For
                    End If
                    '如果不存在装机或者装机功率等于0
                    If ZJJC_SLLGJ = 0 Or ZJZGL_SLLGJ = 0 Then
                        JS_end_SLLGJ = JS_end_SLLGJ + 1
                    End If
                    '根据装机量判断是否直接进入下一次循环
                    If a_1 + a_2 + ZJZGL_SDYRB + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL < LFH_ALL Then
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
                        If a_1 + a_2 + a_3 + ZJZGL_LXSRB + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL < LFH_ALL Then
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
                            If a_1 + a_2 + a_3 + a_4 + ZJZGL_FLLGJ + ZJZGL_KQYRB + ZJZGL_ZRXXHL < LFH_ALL Then
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
                                If a_1 + a_2 + a_3 + a_4 + a_5 + ZJZGL_KQYRB + ZJZGL_ZRXXHL < LFH_ALL Then
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
                                    If a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + ZJZGL_ZRXXHL < LFH_ALL Then
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
                                        '根据装机量判断是否直接进入下一次循环
                                        If a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + a_7 < LFH_ALL Then
                                            GoTo ggg
                                        ElseIf a_1 + a_2 + a_3 + a_4 + a_5 + a_6 + a_7 > LFH_ALL * TZXS Then
                                            GoTo ggg
                                        End If
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                        '离心式冷水机计算
                                        If a_1 >= FH_min_LXSLSJ And ZJJC_LXSLSJ = 1 And ZJZGL_LXSLSJ > 0 Then
                                            Dim ans_LXSLSJ = 离心式冷水机供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_1, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_LXSLSJ, NUM2_LXSLSJ, ZJLGL1_LXSLSJ, ZJLGL2_LXSLSJ, BTHD1_ED_LXSLSJ, BTHD2_ED_LXSLSJ, FJHD1_ED_LXSLSJ, FJHD2_ED_LXSLSJ, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                                            HD_ALL_min_LXSLSJ.AddRange(ans_LXSLSJ(0))
                                            GL_ALL_LXSLSJ.AddRange(ans_LXSLSJ(1))
                                            XL_ALL_LXSLSJ.AddRange(ans_LXSLSJ(2))
                                            FHL1_GL_LXSLSJ.AddRange(ans_LXSLSJ(4))
                                            FHL1_XL_LXSLSJ.AddRange(ans_LXSLSJ(5))
                                            FHL2_GL_LXSLSJ.AddRange(ans_LXSLSJ(6))
                                            FHL2_XL_LXSLSJ.AddRange(ans_LXSLSJ(7))
                                        End If
                                        '水冷螺杆机计算
                                        If a_2 >= FH_min_SLLGJ And ZJJC_SLLGJ = 1 And ZJZGL_SLLGJ > 0 Then
                                            Dim ans_SLLGJ = 水冷螺杆机供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_2, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_SLLGJ, NUM2_SLLGJ, ZJLGL1_SLLGJ, ZJLGL2_SLLGJ, BTHD1_ED_SLLGJ, BTHD2_ED_SLLGJ, FJHD1_ED_SLLGJ, FJHD2_ED_SLLGJ, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                                            HD_ALL_min_SLLGJ.AddRange(ans_SLLGJ(0))
                                            GL_ALL_SLLGJ.AddRange(ans_SLLGJ(1))
                                            XL_ALL_SLLGJ.AddRange(ans_SLLGJ(2))
                                            FHL1_GL_SLLGJ.AddRange(ans_SLLGJ(4))
                                            FHL1_XL_SLLGJ.AddRange(ans_SLLGJ(5))
                                            FHL2_GL_SLLGJ.AddRange(ans_SLLGJ(6))
                                            FHL2_XL_SLLGJ.AddRange(ans_SLLGJ(7))
                                        End If
                                        '水（地）源热泵计算
                                        If a_3 >= FH_min_SDYRB And ZJJC_SDYRB = 1 And ZJZGL_SDYRB > 0 Then
                                            Dim ans_SDYRB = 水_地源热泵供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_3, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_SDYRB, NUM2_SDYRB, ZJLGL1_SDYRB, ZJLGL2_SDYRB, BTHD1_ED_SDYRB, BTHD2_ED_SDYRB, FJHD1_ED_SDYRB, FJHD2_ED_SDYRB, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                                            HD_ALL_min_SDYRB.AddRange(ans_SDYRB(0))
                                            GL_ALL_SDYRB.AddRange(ans_SDYRB(1))
                                            XL_ALL_SDYRB.AddRange(ans_SDYRB(2))
                                            FHL1_GL_SDYRB.AddRange(ans_SDYRB(4))
                                            FHL1_XL_SDYRB.AddRange(ans_SDYRB(5))
                                            FHL2_GL_SDYRB.AddRange(ans_SDYRB(6))
                                            FHL2_XL_SDYRB.AddRange(ans_SDYRB(7))
                                        End If
                                        '离心式热泵计算
                                        If a_4 >= FH_min_LXSRB And ZJJC_LXSRB = 1 And ZJZGL_LXSRB > 0 Then
                                            Dim ans_LXSRB = 离心式热泵供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_4, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_LXSRB, NUM2_LXSRB, ZJLGL1_LXSRB, ZJLGL2_LXSRB, BTHD1_ED_LXSRB, BTHD2_ED_LXSRB, FJHD1_ED_LXSRB, FJHD2_ED_LXSRB, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                                            HD_ALL_min_LXSRB.AddRange(ans_LXSRB(0))
                                            GL_ALL_LXSRB.AddRange(ans_LXSRB(1))
                                            XL_ALL_LXSRB.AddRange(ans_LXSRB(2))
                                            FHL1_GL_LXSRB.AddRange(ans_LXSRB(4))
                                            FHL1_XL_LXSRB.AddRange(ans_LXSRB(5))
                                            FHL2_GL_LXSRB.AddRange(ans_LXSRB(6))
                                            FHL2_XL_LXSRB.AddRange(ans_LXSRB(7))
                                        End If
                                        '风冷螺杆机计算
                                        If a_5 >= FH_min_FLLGJ And ZJJC_FLLGJ = 1 And ZJZGL_FLLGJ > 0 Then
                                            Dim ans_FLLGJ = 风冷螺杆机供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_5, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_FLLGJ, NUM2_FLLGJ, ZJLGL1_FLLGJ, ZJLGL2_FLLGJ, BTHD1_ED_FLLGJ, BTHD2_ED_FLLGJ, FJHD1_ED_FLLGJ, FJHD2_ED_FLLGJ, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
                                            HD_ALL_min_FLLGJ.AddRange(ans_FLLGJ(0))
                                            GL_ALL_FLLGJ.AddRange(ans_FLLGJ(1))
                                            XL_ALL_FLLGJ.AddRange(ans_FLLGJ(2))
                                            FHL1_GL_FLLGJ.AddRange(ans_FLLGJ(4))
                                            FHL1_XL_FLLGJ.AddRange(ans_FLLGJ(5))
                                            FHL2_GL_FLLGJ.AddRange(ans_FLLGJ(6))
                                            FHL2_XL_FLLGJ.AddRange(ans_FLLGJ(7))
                                        End If
                                        '空气源热泵计算
                                        If a_6 >= FH_min_KQYRB And ZJJC_KQYRB = 1 And ZJZGL_KQYRB > 0 Then
                                            Dim ans_KQYRB = 空气源热泵供冷和蓄冷分配寻优计算(ExcelApp, b, FHTJJD, a_6, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_KQYRB, NUM2_KQYRB, ZJLGL1_KQYRB, ZJLGL2_KQYRB, BTHD1_ED_KQYRB, BTHD2_ED_KQYRB, FJHD1_ED_KQYRB, FJHD2_ED_KQYRB, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
                                            HD_ALL_min_KQYRB.AddRange(ans_KQYRB(0))
                                            GL_ALL_KQYRB.AddRange(ans_KQYRB(1))
                                            XL_ALL_KQYRB.AddRange(ans_KQYRB(2))
                                            FHL1_GL_KQYRB.AddRange(ans_KQYRB(4))
                                            FHL1_XL_KQYRB.AddRange(ans_KQYRB(5))
                                            FHL2_GL_KQYRB.AddRange(ans_KQYRB(6))
                                            FHL2_XL_KQYRB.AddRange(ans_KQYRB(7))
                                        End If
                                        '直燃型溴化锂计算
                                        If a_7 >= FH_min_ZRXXHL And ZJJC_ZRXXHL = 1 And ZJZGL_ZRXXHL > 0 Then
                                            Dim ans_ZRXXHL = 直燃型溴化锂供冷寻优计算(ExcelApp, b, FHTJJD, a_7, FHFPCS, LFH_GL_now, LFH_XL_now, NUM1_ZRXXHL, NUM2_ZRXXHL, ZJLGL1_ZRXXHL, ZJLGL2_ZRXXHL, BTHQ1_ED_ZRXXHL, BTHQ2_ED_ZRXXHL, FJHD1_ED_ZRXXHL, FJHD2_ED_ZRXXHL, TRQZLHQXZ, TRQHLXZXS_QT, FJHDXS)
                                            HD_ALL_ZRXXHL.AddRange(ans_ZRXXHL(0))
                                            GL_ALL_ZRXXHL.AddRange(ans_ZRXXHL(1))
                                            XL_ALL_ZRXXHL.AddRange(ans_ZRXXHL(2))
                                            FHL1_GL_ZRXXHL.AddRange(ans_ZRXXHL(4))
                                            FHL1_XL_ZRXXHL.AddRange(ans_ZRXXHL(5))
                                            FHL2_GL_ZRXXHL.AddRange(ans_ZRXXHL(6))
                                            FHL2_XL_ZRXXHL.AddRange(ans_ZRXXHL(7))
                                            HQ_ALL_min_ZRXXHL.AddRange(ans_ZRXXHL(8))
                                        End If
                                        '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                        '求各个计算结果列表的长度，并求出最长的列表长度
                                        Dim n_1 As Integer = HD_ALL_min_LXSLSJ.Count
                                        Dim n_2 As Integer = HD_ALL_min_SLLGJ.Count
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
                                        '离心式冷水机计算结果的列表长度修正
                                        If a_1 < FH_min_LXSLSJ Or ZJJC_LXSLSJ = 0 Or ZJZGL_LXSLSJ = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_1
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_LXSLSJ.AddRange(ans_temp)
                                                GL_ALL_LXSLSJ.AddRange(ans_temp)
                                                XL_ALL_LXSLSJ.AddRange(ans_temp)
                                                FHL1_GL_LXSLSJ.AddRange(ans_temp)
                                                FHL1_XL_LXSLSJ.AddRange(ans_temp)
                                                FHL2_GL_LXSLSJ.AddRange(ans_temp)
                                                FHL2_XL_LXSLSJ.AddRange(ans_temp)
                                            End If
                                        End If
                                        '水冷螺杆机计算结果的列表长度修正
                                        If a_2 < FH_min_SLLGJ Or ZJJC_SLLGJ = 0 Or ZJZGL_SLLGJ = 0 Then
                                            '长度补充
                                            Dim n_x As Integer = n_max_result - n_2
                                            If n_x > 0 Then
                                                '生成全0列表
                                                Dim ans_temp As New List(Of Double)
                                                For i = 0 To n_x - 1
                                                    ans_temp.Add(0)
                                                Next
                                                HD_ALL_min_SLLGJ.AddRange(ans_temp)
                                                GL_ALL_SLLGJ.AddRange(ans_temp)
                                                XL_ALL_SLLGJ.AddRange(ans_temp)
                                                FHL1_GL_SLLGJ.AddRange(ans_temp)
                                                FHL1_XL_SLLGJ.AddRange(ans_temp)
                                                FHL2_GL_SLLGJ.AddRange(ans_temp)
                                                FHL2_XL_SLLGJ.AddRange(ans_temp)
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
                                                GL_ALL_SDYRB.AddRange(ans_temp)
                                                XL_ALL_SDYRB.AddRange(ans_temp)
                                                FHL1_GL_SDYRB.AddRange(ans_temp)
                                                FHL1_XL_SDYRB.AddRange(ans_temp)
                                                FHL2_GL_SDYRB.AddRange(ans_temp)
                                                FHL2_XL_SDYRB.AddRange(ans_temp)
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
                                                GL_ALL_LXSRB.AddRange(ans_temp)
                                                XL_ALL_LXSRB.AddRange(ans_temp)
                                                FHL1_GL_LXSRB.AddRange(ans_temp)
                                                FHL1_XL_LXSRB.AddRange(ans_temp)
                                                FHL2_GL_LXSRB.AddRange(ans_temp)
                                                FHL2_XL_LXSRB.AddRange(ans_temp)
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
                                                GL_ALL_FLLGJ.AddRange(ans_temp)
                                                XL_ALL_FLLGJ.AddRange(ans_temp)
                                                FHL1_GL_FLLGJ.AddRange(ans_temp)
                                                FHL1_XL_FLLGJ.AddRange(ans_temp)
                                                FHL2_GL_FLLGJ.AddRange(ans_temp)
                                                FHL2_XL_FLLGJ.AddRange(ans_temp)
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
                                                GL_ALL_KQYRB.AddRange(ans_temp)
                                                XL_ALL_KQYRB.AddRange(ans_temp)
                                                FHL1_GL_KQYRB.AddRange(ans_temp)
                                                FHL1_XL_KQYRB.AddRange(ans_temp)
                                                FHL2_GL_KQYRB.AddRange(ans_temp)
                                                FHL2_XL_KQYRB.AddRange(ans_temp)
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
                                                GL_ALL_ZRXXHL.AddRange(ans_temp)
                                                XL_ALL_ZRXXHL.AddRange(ans_temp)
                                                FHL1_GL_ZRXXHL.AddRange(ans_temp)
                                                FHL1_XL_ZRXXHL.AddRange(ans_temp)
                                                FHL2_GL_ZRXXHL.AddRange(ans_temp)
                                                FHL2_XL_ZRXXHL.AddRange(ans_temp)
                                                HQ_ALL_min_ZRXXHL.AddRange(ans_temp)
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
            Dim len_1 As Integer = HD_ALL_min_LXSLSJ.Count
            Dim len_2 As Integer = HD_ALL_min_SLLGJ.Count
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
            '供冷和蓄冷
            Dim GL_ALL_r As New List(Of Double)
            Dim XL_ALL_r As New List(Of Double)
            '遍历结果，找出符合要求的各种设备运行负荷率
            '离心式冷水机
            Dim FHL1_GL_LXSLSJ_r As New List(Of Double)
            Dim FHL1_XL_LXSLSJ_r As New List(Of Double)
            Dim FHL2_GL_LXSLSJ_r As New List(Of Double)
            Dim FHL2_XL_LXSLSJ_r As New List(Of Double)
            '水冷螺杆机
            Dim FHL1_GL_SLLGJ_r As New List(Of Double)
            Dim FHL1_XL_SLLGJ_r As New List(Of Double)
            Dim FHL2_GL_SLLGJ_r As New List(Of Double)
            Dim FHL2_XL_SLLGJ_r As New List(Of Double)
            '水（地）源热泵
            Dim FHL1_GL_SDYRB_r As New List(Of Double)
            Dim FHL1_XL_SDYRB_r As New List(Of Double)
            Dim FHL2_GL_SDYRB_r As New List(Of Double)
            Dim FHL2_XL_SDYRB_r As New List(Of Double)
            '离心式热泵
            Dim FHL1_GL_LXSRB_r As New List(Of Double)
            Dim FHL1_XL_LXSRB_r As New List(Of Double)
            Dim FHL2_GL_LXSRB_r As New List(Of Double)
            Dim FHL2_XL_LXSRB_r As New List(Of Double)
            '风冷螺杆机
            Dim FHL1_GL_FLLGJ_r As New List(Of Double)
            Dim FHL1_XL_FLLGJ_r As New List(Of Double)
            Dim FHL2_GL_FLLGJ_r As New List(Of Double)
            Dim FHL2_XL_FLLGJ_r As New List(Of Double)
            '空气源热泵
            Dim FHL1_GL_KQYRB_r As New List(Of Double)
            Dim FHL1_XL_KQYRB_r As New List(Of Double)
            Dim FHL2_GL_KQYRB_r As New List(Of Double)
            Dim FHL2_XL_KQYRB_r As New List(Of Double)
            '直燃型溴化锂
            Dim FHL1_GL_ZRXXHL_r As New List(Of Double)
            Dim FHL1_XL_ZRXXHL_r As New List(Of Double)
            Dim FHL2_GL_ZRXXHL_r As New List(Of Double)
            Dim FHL2_XL_ZRXXHL_r As New List(Of Double)
            '遍历所有结果
            For i = 0 To len_min_result - 1
                Dim GL_ALL_i As Double = GL_ALL_LXSLSJ(i) + GL_ALL_SLLGJ(i) + GL_ALL_SDYRB(i) + GL_ALL_LXSRB(i) + GL_ALL_FLLGJ(i) + GL_ALL_KQYRB(i) + GL_ALL_ZRXXHL(i)
                Dim XL_ALL_i As Double = XL_ALL_LXSLSJ(i) + XL_ALL_SLLGJ(i) + XL_ALL_SDYRB(i) + XL_ALL_LXSRB(i) + XL_ALL_FLLGJ(i) + XL_ALL_KQYRB(i) + XL_ALL_ZRXXHL(i)
                '如果满足条件
                If GL_ALL_i >= LFH_GL_now And XL_ALL_i >= LFH_XL_now Then
                    '耗电、耗气
                    Dim HQ_ALL_i As Double = HQ_ALL_min_ZRXXHL(i)
                    Dim HD_ALL_i As Double = HD_ALL_min_LXSLSJ(i) + HD_ALL_min_SLLGJ(i) + HD_ALL_min_SDYRB(i) + HD_ALL_min_LXSRB(i) + HD_ALL_min_FLLGJ(i) + HD_ALL_min_KQYRB(i) + HD_ALL_ZRXXHL(i)
                    HQ_ALL_r.Add(HQ_ALL_i)
                    HD_ALL_r.Add(HD_ALL_i)
                    '总成本
                    COST_ALL.Add(HD_ALL_i * D_price + HQ_ALL_i * TRQ_price)
                    '供冷、蓄冷
                    GL_ALL_r.Add(GL_ALL_i)
                    XL_ALL_r.Add(XL_ALL_i)
                    '各设备负荷率计算结果
                    '离心式冷水机
                    FHL1_GL_LXSLSJ_r.Add(FHL1_GL_LXSLSJ(i))
                    FHL1_XL_LXSLSJ_r.Add(FHL1_XL_LXSLSJ(i))
                    FHL2_GL_LXSLSJ_r.Add(FHL2_GL_LXSLSJ(i))
                    FHL2_XL_LXSLSJ_r.Add(FHL2_XL_LXSLSJ(i))
                    '水冷螺杆机
                    FHL1_GL_SLLGJ_r.Add(FHL1_GL_SLLGJ(i))
                    FHL1_XL_SLLGJ_r.Add(FHL1_XL_SLLGJ(i))
                    FHL2_GL_SLLGJ_r.Add(FHL2_GL_SLLGJ(i))
                    FHL2_XL_SLLGJ_r.Add(FHL2_XL_SLLGJ(i))
                    '水地源热泵
                    FHL1_GL_SDYRB_r.Add(FHL1_GL_SDYRB(i))
                    FHL1_XL_SDYRB_r.Add(FHL1_XL_SDYRB(i))
                    FHL2_GL_SDYRB_r.Add(FHL2_GL_SDYRB(i))
                    FHL2_XL_SDYRB_r.Add(FHL2_XL_SDYRB(i))
                    '离心式热泵
                    FHL1_GL_LXSRB_r.Add(FHL1_GL_LXSRB(i))
                    FHL1_XL_LXSRB_r.Add(FHL1_XL_LXSRB(i))
                    FHL2_GL_LXSRB_r.Add(FHL2_GL_LXSRB(i))
                    FHL2_XL_LXSRB_r.Add(FHL2_XL_LXSRB(i))
                    '风冷螺杆机
                    FHL1_GL_FLLGJ_r.Add(FHL1_GL_FLLGJ(i))
                    FHL1_XL_FLLGJ_r.Add(FHL1_XL_FLLGJ(i))
                    FHL2_GL_FLLGJ_r.Add(FHL2_GL_FLLGJ(i))
                    FHL2_XL_FLLGJ_r.Add(FHL2_XL_FLLGJ(i))
                    '空气源热泵
                    FHL1_GL_KQYRB_r.Add(FHL1_GL_KQYRB(i))
                    FHL1_XL_KQYRB_r.Add(FHL1_XL_KQYRB(i))
                    FHL2_GL_KQYRB_r.Add(FHL2_GL_KQYRB(i))
                    FHL2_XL_KQYRB_r.Add(FHL2_XL_KQYRB(i))
                    '直燃型溴化锂
                    FHL1_GL_ZRXXHL_r.Add(FHL1_GL_ZRXXHL(i))
                    FHL1_XL_ZRXXHL_r.Add(FHL1_XL_ZRXXHL(i))
                    FHL2_GL_ZRXXHL_r.Add(FHL2_GL_ZRXXHL(i))
                    FHL2_XL_ZRXXHL_r.Add(FHL2_XL_ZRXXHL(i))
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
                '回头重算
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
            '供冷和蓄冷
            Dim GL_ALL_result As Double = GL_ALL_r(COST_min_index)
            Dim XL_ALL_result As Double = XL_ALL_r(COST_min_index)
            '离心式冷水机
            Dim FHL1_GL_LXSLSJ_result As Double = FHL1_GL_LXSLSJ_r(COST_min_index)
            Dim FHL1_XL_LXSLSJ_result As Double = FHL1_XL_LXSLSJ_r(COST_min_index)
            Dim FHL2_GL_LXSLSJ_result As Double = FHL2_GL_LXSLSJ_r(COST_min_index)
            Dim FHL2_XL_LXSLSJ_result As Double = FHL2_XL_LXSLSJ_r(COST_min_index)
            '水冷螺杆机
            Dim FHL1_GL_SLLGJ_result As Double = FHL1_GL_SLLGJ_r(COST_min_index)
            Dim FHL1_XL_SLLGJ_result As Double = FHL1_XL_SLLGJ_r(COST_min_index)
            Dim FHL2_GL_SLLGJ_result As Double = FHL2_GL_SLLGJ_r(COST_min_index)
            Dim FHL2_XL_SLLGJ_result As Double = FHL2_XL_SLLGJ_r(COST_min_index)
            '水（地）源热泵
            Dim FHL1_GL_SDYRB_result As Double = FHL1_GL_SDYRB_r(COST_min_index)
            Dim FHL1_XL_SDYRB_result As Double = FHL1_XL_SDYRB_r(COST_min_index)
            Dim FHL2_GL_SDYRB_result As Double = FHL2_GL_SDYRB_r(COST_min_index)
            Dim FHL2_XL_SDYRB_result As Double = FHL2_XL_SDYRB_r(COST_min_index)
            '离心式热泵
            Dim FHL1_GL_LXSRB_result As Double = FHL1_GL_LXSRB_r(COST_min_index)
            Dim FHL1_XL_LXSRB_result As Double = FHL1_XL_LXSRB_r(COST_min_index)
            Dim FHL2_GL_LXSRB_result As Double = FHL2_GL_LXSRB_r(COST_min_index)
            Dim FHL2_XL_LXSRB_result As Double = FHL2_XL_LXSRB_r(COST_min_index)
            '风冷螺杆机
            Dim FHL1_GL_FLLGJ_result As Double = FHL1_GL_FLLGJ_r(COST_min_index)
            Dim FHL1_XL_FLLGJ_result As Double = FHL1_XL_FLLGJ_r(COST_min_index)
            Dim FHL2_GL_FLLGJ_result As Double = FHL2_GL_FLLGJ_r(COST_min_index)
            Dim FHL2_XL_FLLGJ_result As Double = FHL2_XL_FLLGJ_r(COST_min_index)
            '空气源热泵
            Dim FHL1_GL_KQYRB_result As Double = FHL1_GL_KQYRB_r(COST_min_index)
            Dim FHL1_XL_KQYRB_result As Double = FHL1_XL_KQYRB_r(COST_min_index)
            Dim FHL2_GL_KQYRB_result As Double = FHL2_GL_KQYRB_r(COST_min_index)
            Dim FHL2_XL_KQYRB_result As Double = FHL2_XL_KQYRB_r(COST_min_index)
            '直燃型溴化锂
            Dim FHL1_GL_ZRXXHL_result As Double = FHL1_GL_ZRXXHL_r(COST_min_index)
            Dim FHL1_XL_ZRXXHL_result As Double = FHL1_XL_ZRXXHL_r(COST_min_index)
            Dim FHL2_GL_ZRXXHL_result As Double = FHL2_GL_ZRXXHL_r(COST_min_index)
            Dim FHL2_XL_ZRXXHL_result As Double = FHL2_XL_ZRXXHL_r(COST_min_index)
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '针对计算出的结果进行放缩，防止出现大的误差
            '供冷、蓄冷负荷出力和需求量的比例
            Dim GL_bl As Double = GL_ALL_result / LFH_GL_now
            Dim XL_bl As Double = XL_ALL_result / LFH_XL_now
            '供冷
            If GL_bl > WCXS Then
                '离心式冷水机
                FHL1_GL_LXSLSJ_result = FHL1_GL_LXSLSJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_LXSLSJ_result = FHL2_GL_LXSLSJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '水冷螺杆机
                FHL1_GL_SLLGJ_result = FHL1_GL_SLLGJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_SLLGJ_result = FHL2_GL_SLLGJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '水（地）源热泵
                FHL1_GL_SDYRB_result = FHL1_GL_SDYRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_SDYRB_result = FHL2_GL_SDYRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '离心式热泵
                FHL1_GL_LXSRB_result = FHL1_GL_LXSRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_LXSRB_result = FHL2_GL_LXSRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '风冷螺杆机
                FHL1_GL_FLLGJ_result = FHL1_GL_FLLGJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_FLLGJ_result = FHL2_GL_FLLGJ_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '空气源热泵
                FHL1_GL_KQYRB_result = FHL1_GL_KQYRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_KQYRB_result = FHL2_GL_KQYRB_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                '直燃型溴化锂
                FHL1_GL_ZRXXHL_result = FHL1_GL_ZRXXHL_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_GL_ZRXXHL_result = FHL2_GL_ZRXXHL_r(COST_min_index) / GL_bl * (1 + 2 * FHTJJD / 100)
            End If
            '蓄冷
            If XL_bl > WCXS Then
                '离心式冷水机
                FHL1_XL_LXSLSJ_result = FHL1_XL_LXSLSJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_LXSLSJ_result = FHL2_XL_LXSLSJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '水冷螺杆机
                FHL1_XL_SLLGJ_result = FHL1_XL_SLLGJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_SLLGJ_result = FHL2_XL_SLLGJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '水（地）源热泵
                FHL1_XL_SDYRB_result = FHL1_XL_SDYRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_SDYRB_result = FHL2_XL_SDYRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '离心式热泵
                FHL1_XL_LXSRB_result = FHL1_XL_LXSRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_LXSRB_result = FHL2_XL_LXSRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '风冷螺杆机
                FHL1_XL_FLLGJ_result = FHL1_XL_FLLGJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_FLLGJ_result = FHL2_XL_FLLGJ_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '空气源热泵
                FHL1_XL_KQYRB_result = FHL1_XL_KQYRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_KQYRB_result = FHL2_XL_KQYRB_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                '直燃型溴化锂
                FHL1_XL_ZRXXHL_result = FHL1_XL_ZRXXHL_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
                FHL2_XL_ZRXXHL_result = FHL2_XL_ZRXXHL_r(COST_min_index) / XL_bl * (1 + 2 * FHTJJD / 100)
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            '将各个设备的负荷率计算结果写入Excel
            '负荷率为结合设备数量和单台设备负荷率，折算出的综合负荷率
            '离心式冷水机（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = FHL1_GL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value = FHL1_XL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28).Value = FHL1_GL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40).Value = FHL1_XL_LXSLSJ_result
            '离心式冷水机（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = FHL2_GL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value = FHL2_XL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 29).Value = FHL2_GL_LXSLSJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 41).Value = FHL2_XL_LXSLSJ_result
            '风冷螺杆机（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = FHL1_GL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value = FHL1_XL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 30).Value = FHL1_GL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 42).Value = FHL1_XL_FLLGJ_result
            '风冷螺杆机（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = FHL2_GL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value = FHL2_XL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 31).Value = FHL2_GL_FLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 43).Value = FHL2_XL_FLLGJ_result
            '空气源热泵(1)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = FHL1_GL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value = FHL1_XL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 32).Value = FHL1_GL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 44).Value = FHL1_XL_KQYRB_result
            '空气源热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = FHL2_GL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value = FHL2_XL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 33).Value = FHL2_GL_KQYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 45).Value = FHL2_XL_KQYRB_result
            '水（地）源热泵（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = FHL1_GL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value = FHL1_XL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 34).Value = FHL1_GL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 46).Value = FHL1_XL_SDYRB_result
            '水（地）源热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = FHL2_GL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value = FHL2_XL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 35).Value = FHL2_GL_SDYRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 47).Value = FHL2_XL_SDYRB_result
            '水冷螺杆机（1） 
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = FHL1_GL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value = FHL1_XL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 36).Value = FHL1_GL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 48).Value = FHL1_XL_SLLGJ_result
            '水冷螺杆机(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = FHL2_GL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value = FHL2_XL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 37).Value = FHL2_GL_SLLGJ_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 49).Value = FHL2_XL_SLLGJ_result
            '离心式热泵（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = FHL1_GL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value = FHL1_XL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 38).Value = FHL1_GL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 50).Value = FHL1_XL_LXSRB_result
            '离心式热泵(2)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = FHL2_GL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value = FHL2_XL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 39).Value = FHL2_GL_LXSRB_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51).Value = FHL2_XL_LXSRB_result
            '直燃型溴化锂（1）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = FHL1_GL_ZRXXHL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90).Value = FHL1_GL_ZRXXHL_result
            '直燃型溴化锂（2）
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = FHL2_GL_ZRXXHL_result
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 91).Value = FHL2_GL_ZRXXHL_result
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算制冷季天然气耗量和耗电量综合修正系数
            Call 制冷季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD， calculation_mode)
        End If
    End Sub
    Function 离心式冷水机供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_LXSLSJ / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_LXSLSJ / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 离心式冷水机供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式冷水机供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式冷水机供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 水冷螺杆机供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_SLLGJ / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_SLLGJ / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 水冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 水_地源热泵供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
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
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 水_地源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水_地源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 水_地源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 离心式热泵供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_LXSRB / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_LXSRB / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 离心式热泵供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式热泵供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 离心式热泵供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_water, BTHDXS_XL_water, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 风冷螺杆机供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_air As Double, BTHDXS_XL_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
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
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 风冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 风冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 风冷螺杆机供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 空气源热泵供冷和蓄冷分配寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_air As Double, BTHDXS_XL_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        'ZFH：供冷和蓄冷负荷之和
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
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HD_ALL As New List(Of Double)
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                Dim ans_temp = 空气源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, b_1, c_1, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(0))
                FHL1_GL.Add(ans_temp(1))
                FHL1_XL.Add(ans_temp(2))
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(ans_temp(4))
                NUM1_List.Add(ans_temp(5))
                NUM2_List.Add(ans_temp(6))
                GL_out.Add(b_1)
                XL_out.Add(c_1)
                ZGL_out.Add(LFH_max)
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 空气源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, LFH_max, 0, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 空气源热泵供冷和蓄冷计算(ExcelApp, FHTJJD, 0, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHD1_ED, BTHD2_ED, FJHD1_ED, FJHD2_ED, BTHDXS_GL_air, BTHDXS_XL_air, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(0))
            FHL1_GL.Add(ans_temp(1))
            FHL1_XL.Add(ans_temp(2))
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(ans_temp(4))
            NUM1_List.Add(ans_temp(5))
            NUM2_List.Add(ans_temp(6))
            GL_out.Add(0)
            XL_out.Add(LFH_max)
            ZGL_out.Add(LFH_max)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
        End If
        '将计算出的所有列表返回(蓄冷设备负荷率已经转换成了台数，不需要再次转换)
        Dim ans(10)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = NUM1_List
        ans(9) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂供冷寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double, ZFH As Double, FHFPCS As Integer, LFH_GL_now As Double, LFH_XL_now As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHQ1_ED As Double, BTHQ2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, ZLHQXZ As Double, TRQHLXZXS_QT As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '对于直燃型溴化锂来说，只能供冷，不能蓄冷，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
        'ZFH：供冷和蓄冷负荷之和
        '设备负荷率下限（除以设备数量）
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_ZRXXHL / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_ZRXXHL / NUM2
        Else
            FHL2_min = 0
        End If
        '装机总功率（总和）
        Dim ZJZGL As Double = ZJLGL1 + ZJLGL2
        '计算出的各种结果存入列表
        Dim HQ_ALL As New List(Of Double) '天然气总耗量
        Dim HD_ALL As New List(Of Double) '辅机耗电总量
        Dim FHL1_GL As New List(Of Double) '综合
        Dim FHL1_XL As New List(Of Double) '综合
        Dim FHL2_GL As New List(Of Double) '综合
        Dim FHL2_XL As New List(Of Double) '综合
        Dim NUM1_List As New List(Of Double) '设备（1）数量
        Dim NUM2_List As New List(Of Double) '设备（2）数量
        Dim GL_out As New List(Of Double) '供冷总输出（总和）
        Dim XL_out As New List(Of Double) '蓄冷总输出（总和）
        Dim ZGL_out As New List(Of Double) '总的供冷（总和）+蓄冷出力（总和）
        '设备可以被分配的负荷，从最低允许运行负荷开始
        Dim LFH_min As Double = FHL1_min * ZJLGL1 + FHL2_min * ZJLGL2
        '设备可以供冷的上限（冷负荷需求量和设备装机量中较小的值）
        Dim LFH_max As Double = Math.Min(ZJZGL, ZFH)
        '如果此时存在蓄冷负荷，再将负荷分成供冷负荷和蓄冷负荷
        If LFH_GL_now > 0 And LFH_XL_now > 0 Then
            '如果供冷和蓄冷均大于0
            'b_1表示当前分配给供冷的负荷
            For b_1 = 0 To LFH_max Step LFH_max / FHFPCS
                'c_1表示当前分配蓄冷的负荷
                'Dim c_1 As Double = LFH_max - b_1
                '求计算结果
                '对于直燃型溴化锂来说，只能供冷，不能蓄冷，但是为了保持在计算过程中，结果的列表长度一致，也进行同样的寻优，但是结果不变
                Dim ans_temp = 直燃型溴化锂供冷计算_方法二(ExcelApp, FHTJJD, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHQ1_ED, BTHQ2_ED, FJHD1_ED, FJHD2_ED, ZLHQXZ, TRQHLXZXS_QT, FJHDXS)
                '计算结果加入列表
                HD_ALL.Add(ans_temp(1))
                FHL1_GL.Add(ans_temp(2))
                FHL1_XL.Add(0)
                FHL2_GL.Add(ans_temp(3))
                FHL2_XL.Add(0)
                NUM1_List.Add(ans_temp(4))
                NUM2_List.Add(ans_temp(5))
                GL_out.Add(LFH_max)
                XL_out.Add(0)
                ZGL_out.Add(LFH_max)
                HQ_ALL.Add(ans_temp(0))
            Next
        ElseIf LFH_GL_now > 0 And LFH_XL_now = 0 Then
            '如果没有蓄冷负荷，则直接计算
            '求计算结果
            Dim ans_temp = 直燃型溴化锂供冷计算_方法二(ExcelApp, FHTJJD, LFH_max, NUM1, NUM2, ZJLGL1, ZJLGL2, BTHQ1_ED, BTHQ2_ED, FJHD1_ED, FJHD2_ED, ZLHQXZ, TRQHLXZXS_QT, FJHDXS)
            '计算结果加入列表
            HD_ALL.Add(ans_temp(1))
            FHL1_GL.Add(ans_temp(2))
            FHL1_XL.Add(0)
            FHL2_GL.Add(ans_temp(3))
            FHL2_XL.Add(0)
            NUM1_List.Add(ans_temp(4))
            NUM2_List.Add(ans_temp(5))
            GL_out.Add(LFH_max)
            XL_out.Add(0)
            ZGL_out.Add(LFH_max)
            HQ_ALL.Add(ans_temp(0))
        ElseIf LFH_GL_now = 0 And LFH_XL_now > 0 Then
            '如果没有供冷负荷，则直接计算
            '计算结果加入列表，直燃型溴化锂不能蓄冷，结果全部是0
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
            HQ_ALL.Add(0)
        Else
            '计算结果加入列表（此时结果全为0）
            HD_ALL.Add(0)
            FHL1_GL.Add(0)
            FHL1_XL.Add(0)
            FHL2_GL.Add(0)
            FHL2_XL.Add(0)
            NUM1_List.Add(0)
            NUM2_List.Add(0)
            GL_out.Add(0)
            XL_out.Add(0)
            ZGL_out.Add(0)
            HQ_ALL.Add(0)
        End If
        '将计算出的所有列表返回
        Dim ans(11)
        ans(0) = HD_ALL
        ans(1) = GL_out
        ans(2) = XL_out
        ans(3) = ZGL_out
        ans(4) = FHL1_GL
        ans(5) = FHL1_XL
        ans(6) = FHL2_GL
        ans(7) = FHL2_XL
        ans(8) = HQ_ALL
        ans(9) = NUM1_List
        ans(10) = NUM2_List
        '返回结果
        Return ans
    End Function
    Function 离心式冷水机供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_LXSLSJ
        Dim FHL2_min As Double = FHL2_min_LXSLSJ
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_water * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_water * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a2表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 离心式冷水机制冷COP曲线(a1)
                        Dim XZXS2 As Double = 离心式冷水机制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 水冷螺杆机供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_SLLGJ
        Dim FHL2_min As Double = FHL2_min_SLLGJ
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_water * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_water * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 水冷螺杆机制冷COP曲线(a1)
                        Dim XZXS2 As Double = 水冷螺杆机制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 水_地源热泵供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_SDYRB
        Dim FHL2_min As Double = FHL2_min_SDYRB
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_water * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_water * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 水_地源热泵制冷COP曲线(a1)
                        Dim XZXS2 As Double = 水_地源热泵制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 离心式热泵供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_water As Double, BTHDXS_XL_water As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_LXSRB
        Dim FHL2_min As Double = FHL2_min_LXSRB
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_water * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_water * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 离心式热泵制冷COP曲线(a1)
                        Dim XZXS2 As Double = 离心式热泵制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 风冷螺杆机供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_air As Double, BTHDXS_XL_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_FLLGJ
        Dim FHL2_min As Double = FHL2_min_FLLGJ
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_air * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_air * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 风冷螺杆式热泵制冷COP曲线(a1)
                        Dim XZXS2 As Double = 风冷螺杆式热泵制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 空气源热泵供冷和蓄冷计算(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, LFH_XL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHD1_ED As Double, BTHD2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, BTHDXS_GL_air As Double, BTHDXS_XL_air As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_KQYRB
        Dim FHL2_min As Double = FHL2_min_KQYRB
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '根据供冷功率和蓄冷功率的比例，计算出本体耗电的综合修正系数
        Dim BTHDXS_ZH As Double = BTHDXS_GL_air * LFH_GL / (LFH_GL + LFH_XL) + BTHDXS_XL_air * LFH_XL / (LFH_GL + LFH_XL)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL + LFH_XL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 空气源热泵制冷COP曲线(a1)
                        Dim XZXS2 As Double = 空气源热泵制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗电功率（总和）
                        Dim BTHD1_now As Double = BTHDXS_ZH * n1 * a1 * BTHD1_ED / XZXS1
                        Dim BTHD2_now As Double = BTHDXS_ZH * n2 * a2 * BTHD2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗电功率（总和）
                        Dim ZHD As Double = BTHD1_now + BTHD2_now + FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL + LFH_XL Then
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表（存入单台负荷率，设备启动数量）
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
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数（单台）
        Dim FHL1_GL_result As Double
        Dim FHL1_XL_result As Double
        Dim FHL2_GL_result As Double
        Dim FHL2_XL_result As Double
        '得到结果
        HD_ALL_min = HD_ALL.Min
        HD_index_min = HD_ALL.IndexOf(HD_ALL_min)
        '求此时设备（1）和设备（2）的负荷率（单台）
        FHL1_result = FHL1(HD_index_min)
        FHL2_result = FHL2(HD_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HD_index_min)
        NUM2_result = NUM2_List(HD_index_min)
        '负荷率的换算还要考虑RXCS的因素（单台）
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        '负荷率根据供冷和蓄冷功率比例进行分配， 蓄冷负荷率转换为台数（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
            FHL1_XL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_1
            FHL1_XL_result = (FHL1_result * NUM1_result / NUM1) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_1 * NUM1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
            FHL2_XL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * LFH_GL / (LFH_GL + LFH_XL) * RCZHXS_2
            FHL2_XL_result = (FHL2_result * NUM2_result / NUM2) * LFH_XL / (LFH_GL + LFH_XL) * RCZHXS_2 * NUM2
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果存到数组
        Dim ans(7) As Double
        ans(0) = HD_ALL_min
        ans(1) = FHL1_GL_result
        ans(2) = FHL1_XL_result
        ans(3) = FHL2_GL_result
        ans(4) = FHL2_XL_result
        ans(5) = NUM1_result
        ans(6) = NUM2_result
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂供冷计算_方法二(ExcelApp As Object, FHTJJD As Double, LFH_GL As Double, NUM1 As Double, NUM2 As Double, ZJLGL1 As Double, ZJLGL2 As Double, BTHQ1_ED As Double, BTHQ2_ED As Double, FJHD1_ED As Double, FJHD2_ED As Double, ZLHQXZ As Double, TRQHLXZXS_QT As Double, FJHDXS As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备可以允许运行的负荷率下限（不除以设备数量，表示的是单台的负荷率下限）
        Dim FHL1_min As Double = FHL1_min_ZRXXHL
        Dim FHL2_min As Double = FHL2_min_ZRXXHL
        '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
        Dim RCXS_a As Double = RCXS * (FHTJJD / 0.5)
        '不严格按照设置好的六种设备启动顺序进行计算，而是六种设备综合在一起进行全局寻优计算，寻找总成本最低的运行模式
        '设备（1）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL1 = ZJLGL1 + RCXS_a
        '设备（2）装机功率（总和），在从Excel读取的值的基础上放大一些
        ZJLGL2 = ZJLGL2 + RCXS_a
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '列表，储存总耗气量，耗电量，设备（1）负荷率和设备（2）负荷率
        Dim HQ_ALL As New List(Of Double）
        Dim HD_ALL As New List(Of Double)
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
                Dim LFH1_ALL_temp As Double
                If NUM1 = 0 Then
                    LFH1_ALL_temp = 0
                Else
                    LFH1_ALL_temp = n1 * ZJLGL1 / NUM1
                End If
                Dim LFH2_ALL_temp As Double
                If NUM2 = 0 Then
                    LFH2_ALL_temp = 0
                Else
                    LFH2_ALL_temp = n2 * ZJLGL2 / NUM2
                End If
                If (LFH1_ALL_temp + LFH2_ALL_temp) < (LFH_GL) Then
                    GoTo zzzz
                End If
                '穷举设备负荷率
                '直燃型溴化锂制冷负荷率可以超发到1.2
                For a1 = FHL1_min To (1.2 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率（单台）
                    For a2 = FHL2_min To (1.2 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率（单台）
                        '计算设备（1）和设备（2）本体的效率修正系数（单台）
                        Dim XZXS1 As Double = 直燃型溴化锂制冷COP曲线(a1)
                        Dim XZXS2 As Double = 直燃型溴化锂制冷COP曲线(a2)
                        '计算设备（1）和设备（2）本体的耗气量（总和）
                        Dim BTHQ1_now As Double = TRQHLXZXS_QT * ZLHQXZ * n1 * a1 * BTHQ1_ED / XZXS1
                        Dim BTHQ2_now As Double = TRQHLXZXS_QT * ZLHQXZ * n2 * a2 * BTHQ2_ED / XZXS2
                        '计算设备（1）和设备（2）辅助设备的耗电功率（直接按照负荷率打折）（总和）
                        Dim FJHD1_now As Double = FJHDXS * n1 * a1 * FJHD1_ED
                        Dim FJHD2_now As Double = FJHDXS * n2 * a2 * FJHD2_ED
                        '计算此时的总耗气量（总和）
                        Dim ZHQ As Double = BTHQ1_now + BTHQ2_now
                        '计算此时的总耗电量（总和）
                        Dim ZHD As Double = FJHD1_now + FJHD2_now
                        '计算此时的总出力（总和）
                        Dim LGL1_out_now As Double
                        If ZJLGL1 <= RCXS_a Then
                            LGL1_out_now = 0
                        Else
                            If NUM1 = 0 Then
                                LGL1_out_now = 0
                            Else
                                LGL1_out_now = n1 * a1 * ZJLGL1 / NUM1
                            End If
                        End If
                        Dim LGL2_out_now As Double
                        If ZJLGL2 <= RCXS_a Then
                            LGL2_out_now = 0
                        Else
                            If NUM2 = 0 Then
                                LGL2_out_now = 0
                            Else
                                LGL2_out_now = n2 * a2 * ZJLGL2 / NUM2
                            End If
                        End If
                        Dim LGL_out_all As Double = LGL1_out_now + LGL2_out_now
                        '如果达到了冷负荷需求，则跳出内层循环（总和）
                        If LGL_out_all >= LFH_GL Then
                            '计算结果加入列表
                            HQ_ALL.Add(ZHQ)
                            HD_ALL.Add(ZHD)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1.2 And a2 >= 1.2 Then
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
        '找出总耗气量最低的运行工况
        Dim HQ_ALL_min As Double
        Dim HQ_index_min As Integer
        '求此时辅机耗电总功率
        Dim HD_result As Double
        '求此时设备（1）和设备（2）的负荷率
        Dim FHL1_result As Double
        Dim FHL2_result As Double
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double
        Dim NUM2_result As Double
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配，蓄冷负荷率转换为台数
        Dim FHL1_GL_result As Double
        Dim FHL2_GL_result As Double
        '得到结果
        HQ_ALL_min = HQ_ALL.Min
        HQ_index_min = HQ_ALL.IndexOf(HQ_ALL_min)
        '求此时辅机耗电总功率
        HD_result = HD_ALL(HQ_index_min)
        '求此时设备（1）和设备（2）的负荷率
        FHL1_result = FHL1(HQ_index_min)
        FHL2_result = FHL2(HQ_index_min)
        '求此时设备（1）和设备（2）启动数量
        NUM1_result = NUM1_List(HQ_index_min)
        NUM2_result = NUM2_List(HQ_index_min)
        '返回计算出的设备（1）和设备（2）负荷率，负荷率根据供冷和蓄冷功率比例进行分配
        '负荷率的换算还要考虑RXCS的因素
        Dim RCZHXS_1 As Double
        If ZJLGL1 <= RCXS_a Then
            RCZHXS_1 = 0
        Else
            RCZHXS_1 = （ZJLGL1 / (ZJLGL1 - RCXS_a)）
        End If
        Dim RCZHXS_2 As Double
        If ZJLGL2 <= RCXS_a Then
            RCZHXS_2 = 0
        Else
            RCZHXS_2 = （ZJLGL2 / (ZJLGL2 - RCXS_a)）
        End If
        '根据单台设备负荷率和设备启动数量，计算此时的综合负荷率
        '返回计算出的设备（1）和设备（2）负荷率（综合）
        If NUM1 = 0 Then
            FHL1_GL_result = 0
        Else
            FHL1_GL_result = (FHL1_result * NUM1_result / NUM1) * RCZHXS_1
        End If
        If NUM2 = 0 Then
            FHL2_GL_result = 0
        Else
            FHL2_GL_result = (FHL2_result * NUM2_result / NUM2) * RCZHXS_2
        End If
        '返回结果存到数组
        Dim ans(6) As Double
        ans(0) = HQ_ALL_min
        ans(1) = HD_result
        ans(2) = FHL1_GL_result
        ans(3) = FHL2_GL_result
        ans(4) = NUM1_result
        ans(5) = NUM2_result
        '返回结果
        Return ans
    End Function

End Module
