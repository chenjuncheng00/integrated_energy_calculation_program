Module 梯级供热和混水供热计算
    Sub 存在混水供热的工况特殊处理(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, calculation_mode As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '如果存在混水供热设备，如果此时工况的总制热功率和蓄热功率不满足热负荷的需求（因为存在混水供热设备，因此常规的顺序1至6设备的总装机功率会低于设计值，因此在最大负荷情况下可能会存在不满足的情况）
        '此时可以将参与混水的天然气锅炉、直燃型溴化锂、电采暖锅炉的负荷率突破100%，暂时满足供热和蓄热功率的需求
        '仅适用于计算模式1
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing And calculation_mode = 1 Then
            '检测存在的设备类型
            Dim GLXHLJC As Integer = 0 '天然气锅炉或者直燃型溴化锂检测，如果存在，则为1，否则为0
            Dim DGLJC As Integer = 0 '电锅炉检测，如果存在，则为1，否则为0
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
            '记录此时天然气锅炉制热负荷率
            Dim TRQGLFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value
            Dim TRQGLFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value
            '记录此时直燃型溴化锂制热负荷率
            Dim ZRXHLFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value
            Dim ZRXHLFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value
            '记录此时电采暖锅炉制热负荷率和蓄热负荷率
            Dim DGLZRFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value
            Dim DGLZRFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value
            Dim DGLXRFHL1 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value
            Dim DGLXRFHL2 As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '如果设备为天然气锅炉或者直燃型溴化锂（只有供热，没有蓄热）
            For i = 15 To 20 '列号
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "天然气锅炉" Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "直燃型溴化锂" Then
                    GLXHLJC = 1
                    Exit For
                End If
            Next
            '如果设备为电采暖锅炉（不仅供热，而且蓄热）
            For i = 15 To 20 '列号
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, i).Value = "电采暖锅炉" Then
                    DGLJC = 1
                    Exit For
                End If
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '存在天然气锅炉或者直燃型溴化锂
            If GLXHLJC = 1 Then
                '检测此时实际制热功率是否满足热负荷总需求量
                '天然气锅炉
                '（1）（2）同时有，则同时往上加
                If TRQGL1ZRGL > 0 And TRQGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = TRQGLFHL1 + FHTJJD * i / 100
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = TRQGLFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（1）
                If TRQGL1ZRGL > 0 And TRQGL2ZRGL = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在天然气锅炉1，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = TRQGLFHL1 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（2）
                If TRQGL1ZRGL = 0 And TRQGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在天然气锅炉2，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = TRQGLFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value
                            Exit For
                        End If
                    Next
                End If
                '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '直燃型溴化锂
                '（1）（2）同时有，则同时往上加
                If ZRXHL1ZRGL > 0 And ZRXHL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在直燃型溴化锂1，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = ZRXHLFHL1 + FHTJJD * i / 100
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = ZRXHLFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（1）
                If ZRXHL1ZRGL > 0 And ZRXHL2ZRGL = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在直燃型溴化锂1，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = ZRXHLFHL1 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（2）
                If ZRXHL1ZRGL = 0 And ZRXHL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在直燃型溴化锂2，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = ZRXHLFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value
                            Exit For
                        End If
                    Next
                End If
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '存在电采暖锅炉
            If DGLJC = 1 Then
                '电采暖锅炉供热
                '（1）（2）同时有，则同时往上加
                If DCNGL1ZRGL > 0 And DCNGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在电采暖锅炉1，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = DGLZRFHL1 + FHTJJD * i / 100
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = DGLZRFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（1）
                If DCNGL1ZRGL > 0 And DCNGL2ZRGL = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在电采暖锅炉1，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = DGLZRFHL1 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（2）
                If DCNGL1ZRGL = 0 And DCNGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b) Then ' 存在电采暖锅炉2，且制热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = DGLZRFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value
                            Exit For
                        End If
                    Next
                End If
                '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '电采暖锅炉蓄热（有可能存在电锅炉都用于供热了，没有去蓄热，因此只要检测到有电锅炉并且蓄热不满足，就计算）
                '（1）（2）同时有，则同时往上加
                If DCNGL1ZRGL > 0 And DCNGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Then ' 存在电采暖锅炉1，且蓄热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = DGLXRFHL1 + FHTJJD * i / 100
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = DGLXRFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（1）
                If DCNGL1ZRGL > 0 And DCNGL2ZRGL = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Then ' 存在电采暖锅炉1，且蓄热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = DGLXRFHL1 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value
                            Exit For
                        End If
                    Next
                End If
                '只有（2）
                If DCNGL1ZRGL = 0 And DCNGL2ZRGL > 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Then ' 存在电采暖锅炉2，且蓄热负荷不满足需求
                    For i = 1 To 1000
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = DGLXRFHL2 + FHTJJD * i / 100
                        '跳出循环条件
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 Then
                            '记录下此时的设备负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Function 混水供冷供热量计算(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '常规计算模式需要
        '混水冷负荷、热负荷重置为0
        Dim HSLFH As Double = 0
        Dim HSRFH As Double = 0
        '先将内燃机余热利用方式带入
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 77).Value
        '将本工况的冷热负荷需求量带入（用户区分仅计算耗电量不计算供冷供热量的设备类型、混水向外供冷供热的设备类型是供冷用还是供热用）
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 12).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 12).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 21).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 21).Value
        '将本工况的蓄冷和蓄热功率带入计算
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 14).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 14).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 23).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 23).Value
        '将仅计算耗电量不计算供冷供热量的设备类型和负荷率以及混水向外供冷供热的设备类型和设备负荷带入
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 80), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 83)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83)).Value
        '读取此时的供冷总负荷和供热总负荷（混水向外供冷供热的设备的供冷负荷和供热负荷）
        HSLFH = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value '混水冷负荷
        HSRFH = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value '混水热负荷
        '计算结果
        Dim ans(2)
        ans(0) = HSLFH
        ans(1) = HSRFH
        '返回结果
        Return ans
    End Function
    Sub 常规计算模式混水供热计算(ExcelApp As Object, FHTJJD As Double, JSBC As Integer, b As Integer, calculation_mode As Integer, XHLZR As Double, HSGRGLBL As Double)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '仅适用于计算模式1（常规计算模式）
        If calculation_mode = 1 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value <> Nothing Then
            '混水计算特殊处理
            Call 存在混水供热的工况特殊处理(ExcelApp, b, FHTJJD, JSBC, calculation_mode)
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '混水设备供热功率
            '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热总功率（装机量，制热出力最大值）
            Dim ZJRGL_HS As Double = 0
            '混水设备功率=风冷热泵+水（地）源热泵+空气源热泵（一般情况下，一个项目只会有这3种设备中的一种）,此处为混水设备的装机总功率
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "空气源热泵" Then
                ZJRGL_HS = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "水(地)源热泵" Then
                ZJRGL_HS = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value = "风冷螺杆机" Then
                ZJRGL_HS = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
            End If
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
            '天然气锅炉供热负荷率
            Dim FHL_TRQGL1_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value
            Dim FHL_TRQGL2_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value
            '直燃型溴化锂供热负荷率
            Dim FHL_ZRXXHL1_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value
            Dim FHL_ZRXXHL2_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value
            '电锅炉供热+蓄热负荷率
            Dim FHL_DGL1_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value
            Dim FHL_DGL2_now As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value
            '天然气锅炉+电锅炉+直燃型溴化锂装机总功率（默认这三种设备只会存在一种）= 参与混水的两种设备的功率之和
            Dim ZRGL_HS_ALL_now As Double = TRQGL1ZRGL * FHL_TRQGL1_now + TRQGL2ZRGL * FHL_TRQGL2_now + ZRXHL1ZRGL * FHL_ZRXXHL1_now + ZRXHL2ZRGL * FHL_ZRXXHL2_now + DCNGL1ZRGL * FHL_DGL1_now + DCNGL2ZRGL * FHL_DGL2_now
            '当前混水设备制热功率 = 风冷热泵 + 水（地）源热泵 + 空气源热泵
            Dim ZJRGL_HS_ALL_now As Double
            '当前混水设备制热负荷率 = 风冷热泵 + 水（地）源热泵 + 空气源热泵
            Dim FHL_HS_now As Double
            If ZJRGL_HS > 0 Then
                '默认天然气锅炉、电锅炉、直燃型溴化锂不会同时存在
                '求出的结果 ZJRGL_HS_ALL_now = 风冷热泵+空气源热泵+水(地)源热泵制热总功率
                ZJRGL_HS_ALL_now = ZRGL_HS_ALL_now * HSGRGLBL
                '混水设备负荷率
                FHL_HS_now = ZJRGL_HS_ALL_now / ZJRGL_HS
            Else
                ZJRGL_HS_ALL_now = 0
                FHL_HS_now = 0
            End If
            If FHL_HS_now > 0 Then
                '将计算结果写入Excel
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 82).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 82).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 83).Value = Math.Round(FHL_HS_now, 4)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 83).Value = Math.Round(FHL_HS_now, 4)
                '将已有的制热设备和蓄热设备运行负荷率重置为0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 70)).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 70)).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97)).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97)).Value = 0
                '根据供热负荷和蓄热负荷的比例，分摊混水供热设备的功率
                Dim HS_GRGL_now As Double = RFHZXQL(b) / (RFHZXQL(b) + XNXRGL(b)) * ZJRGL_HS_ALL_now
                '制热计算，常规计算模式
                Call 设备六种顺序制热常规模式计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HS_GRGL_now, calculation_mode)
                '蓄热计算，常规计算模式
                Call 蓄能装置蓄热工况常规模式计算(ExcelApp, b, FHTJJD, calculation_mode)
            End If
        End If
    End Sub
    Sub 常规计算模式梯级供热计算(ExcelApp As Object, b As Integer, calculation_mode As Integer, TJGRFHBL As Double)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '仅适用于计算模式1（常规计算模式）
        '常规计算模式和全局寻优计算模式都采用这个梯级供热计算方法
        '判断是否需要进行梯级供热计算
        '自动计算梯级供热设备（仅计算耗电量，不计算供热量的设备）的负荷率
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing And calculation_mode = 1 Then
            '让用户手动输入负荷率比例（一级热泵负荷率/二级热泵负荷率），一级热泵跟随二级热泵负荷率的变化而变化，在二级热泵负荷率基础上乘以系数，即为一级热泵负荷率
            'Dim TJGRFHBL = InputBox("请输入梯级供热一级热泵（仅计算耗电量，不计算供热量设备）与二级热泵负荷比例系数，（一级热泵负荷比例÷二级热泵负荷比例）", "请输入梯级供热设备负荷比例系数", 1)
            '逐个工况计算离心式热泵机组（供热+蓄热）的总计负荷比例系数
            '读取采暖季装机方案及参数
            Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
            '离心式热泵
            '装机数量
            Dim LXSRB1_num As Double = ans_ZJFA_R(42)
            Dim LXSRB2_num As Double = ans_ZJFA_R(43)
            '装机功率（总和）
            Dim LXSRB1ZRGL As Double = ans_ZJFA_R(44)
            Dim LXSRB2ZRGL As Double = ans_ZJFA_R(45)
            '离心式热泵装机总功率（总和）
            Dim LXSRBGRZGL As Double = LXSRB1ZRGL + LXSRB2ZRGL
            '目前离心式热泵1和2的（供热+蓄热）总负荷率
            Dim LXSRB1_FHL As Double = 0
            Dim LXSRB2_FHL As Double = 0
            If LXSRB1_num > 0 Then
                LXSRB1_FHL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value / LXSRB1_num
            Else
                LXSRB1_FHL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value + 0
            End If
            If LXSRB2_num > 0 Then
                LXSRB2_FHL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value / LXSRB2_num
            Else
                LXSRB2_FHL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value + 0
            End If
            '计算离心式热泵供热+蓄热的实际功率
            Dim LXSRBSJGL As Double = LXSRB1ZRGL * LXSRB1_FHL + LXSRB2ZRGL * LXSRB2_FHL
            '一级热泵总负荷率
            Dim YJRBZFHL As Double = （LXSRBSJGL / LXSRBGRZGL） * TJGRFHBL
            '将结果写入Excel
            If YJRBZFHL > 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value <> Nothing Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 80).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 80).Value
                '将一级热泵负荷率写入表格
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 81).Value = Math.Round(YJRBZFHL, 4)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 81).Value = Math.Round(YJRBZFHL, 4)
            End If
        End If
    End Sub
    Sub 梯级或者混水供热计算(ExcelApp As Object, FHTJJD As Double, JSBC As Integer, n As Integer， D_price_GF1 As Double, D_price_GF2 As Double, D_price_F1 As Double, D_price_F2 As Double, D_price_P1 As Double, D_price_P2 As Double, D_price_G1 As Double, D_price_G2 As Double, D_price_QT1 As Double, D_price_QT2 As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '本SUB暂时弃用，本SUB暂时弃用，本SUB暂时弃用，本SUB暂时弃用，本SUB暂时弃用，本SUB暂时弃用，本SUB暂时弃用
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '定义局部变量
        Dim XH As Integer = 0 '有梯级供热或者有混水供热的工况序号
        '定义数组，记录需要计算的全部工况序号
        Dim GKXH(10000) As Integer '工况序号
        '判断是否需要进行梯级供热或者混水供热计算
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
        '寻找有梯级供热或者混水供热的工况的序号
        If TJGRZTJC > 0 Or HSGRZTJC > 0 Then
            For i = 1 To n
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then '如果有仅计算耗电量，不计算供热量的设备
                    XH = XH + 1
                    GKXH(XH) = i
                End If
            Next
        End If
        '自动计算梯级供热设备（仅计算耗电量，不计算供热量的设备）的负荷率
        If TJGRZTJC > 0 Then
            '计算循环体
            Call 计算循环体(ExcelApp, n)
            '让用户手动输入负荷率比例（一级热泵负荷率/二级热泵负荷率），一级热泵跟随二级热泵负荷率的变化而变化，在二级热泵负荷率基础上乘以系数，即为一级热泵负荷率
            'Dim TJGRFHBL = InputBox("请输入梯级供热一级热泵（仅计算耗电量，不计算供热量设备）与二级热泵负荷比例系数，（一级热泵负荷比例÷二级热泵负荷比例）", "请输入梯级供热设备负荷比例系数", 1)
            '逐个工况计算离心式热泵机组（供热+蓄热）的总计负荷比例系数
            Dim LXSRBGRZGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value '离心式热泵制热总功率
            Dim LXSRB1ZRGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 18).Value '离心式热泵1制热功率
            Dim LXSRB2ZRGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 18).Value '离心式热泵2制热功率
            Dim YJRBZFHL(10000) '数组，用于储存一级热泵总负荷率
            For i = 1 To n '遍历所有工况，计算离心式热泵总负荷率
                '计算离心式热泵供热+蓄热的实际功率
                Dim LXSRBSJGL = LXSRB1ZRGL * (ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 29).Value + ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 56).Value) + LXSRB2ZRGL * (ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 30).Value + ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 57).Value)
                YJRBZFHL(i) = （LXSRBSJGL / LXSRBGRZGL） * TJGRFHBL
                If YJRBZFHL(i) > 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 80).Value <> Nothing Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 81).Value = Math.Round(YJRBZFHL(i), 4) '将一级热泵负荷率写入表格
                End If
            Next
            '如果存在梯级供热，在进行一遍负荷分析计算
            For i = 1 To XH
                Dim b As Integer = GKXH(i)
                Call 清空指定工况输入输出数据(ExcelApp, b)
                '进行正常的负荷分析（主要技术指标）计算
                Call 负荷分析计算程序(ExcelApp, b, FHTJJD, JSBC， D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                '对计算出的制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                Call 制热和蓄热空调设备负荷率修正(ExcelApp, b, calculation_mode)
                '计算制热季天然气耗量和耗电量综合修正系数
                Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, b, FHTJJD, calculation_mode)
            Next
            '计算循环体
            Call 计算循环体(ExcelApp, n)
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '混水供热需要反复迭代计算
        '自动计算混水供热设备的负荷率
        '仅仅适用于计算模式1
        If HSGRZTJC > 0 And calculation_mode = 1 Then
            '让用户手动输入混水设备制热功率，占总热功率（总热功率指的是，几种混水向外供热的设备总供热功率的合计）的比例
            '计算模式1采用的混水设备比例
            'Dim HSGRGLBL_a As Double = InputBox("请输入混水供热功率比例系数（选择的混水供热设备的供热功率，占两种设备总供热功率的比例，例如(风冷热泵的制热功率/（天然气锅炉的制热功率+风冷热泵的制热功率）)）", "请输入混水供热功率比例系数", 0.5)
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
            '参与混水的风冷热泵+空气源热泵+水(地)源热泵制热实际运行功率
            Dim HSSBSJGL As Double '混水设备实际运行的功率
            '逐个工况计算混水设备负荷率
            Dim HSSBFHL As Double '用于储存混水设备负荷率
            Dim HSSBSJBL As Double '储存计算过程中的混水设备的实际比例
            Dim TRQGLSJGL As Double '天然气锅炉实际功率
            Dim ZRXHLSJGL As Double '直燃型溴化锂实际功率
            Dim DCNGLSJGL As Double '电采暖锅炉实际功率
            Dim HSBL(n - 1) As Double '数组，混水比例，用于在计算完成后显示各个工况计算出的实际混水比例
            'Dim JS As Integer = 0 '计数，统计有多少个混水工况
            For i = 1 To n '工况序号
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 82).Value <> Nothing Then
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
                    If HSGRGLBL > HSSBGL / (HSSBGL + TRQGL1ZRGL + TRQGL2ZRGL + ZRXHL1ZRGL + ZRXHL2ZRGL + DCNGL1ZRGL + DCNGL2ZRGL) Then
                        MsgBox("输入的<混水供热供冷比例系数>大于装机选择的两种设备的实际比例，计算结果可能会不正确！" & "工况序号为：  " & i)
                    End If
                    '计算天然气锅炉、直燃型溴化锂、电采暖锅炉实际总供热功率+蓄热功率(仅电采暖锅炉有蓄热功率)初始值（实际总供热功率也就是两种设备混水供热的总功率，然后按照输入的负荷比例进行分配）（此处读取仅用作初试计算，计算一个初试量）
                    Dim TRQGLSJGL_CSZ As Double = TRQGL1ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 52).Value + TRQGL2ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 53).Value
                    Dim ZRXHLSJGL_CSZ As Double = ZRXHL1ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 92).Value + ZRXHL2ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 93).Value
                    Dim DCNGLSJGL_CSZ As Double = DCNGL1ZRGL * (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 94).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 96).Value) + DCNGL2ZRGL * (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 95).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 97).Value)
                    '计算混水设备负荷率计算初始值（如果此时有内燃机和溴化锂，则这个初始值会偏大，后续计算会往下减；如果没有内燃机和溴化锂，这个比例基本正确，但是为了防止出错，初始值放大10%，然后往下减）
                    '如果此时TRQGLSJGL_CSZ + ZRXHLSJGL_CSZ + DCNGLSJGL_CSZ=0,直接结束计算，进入下一个工况
                    If TRQGLSJGL_CSZ + ZRXHLSJGL_CSZ + DCNGLSJGL_CSZ = 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = 0
                        HSSBSJBL = 0
                        GoTo aaaaa
                    End If
                    '(TRQGLSJGL_CSZ + ZRXHLSJGL_CSZ + DCNGLSJGL_CSZ）即为两种混水设备的总输出功率，例如风冷螺杆式热泵+天然气锅炉混水总功率
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = (TRQGLSJGL_CSZ + ZRXHLSJGL_CSZ + DCNGLSJGL_CSZ） * HSGRGLBL / HSSBGL
                    '定义混水设备负荷率初始值
                    Dim HSSBFHL_CSZ As Double = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value
                    'MsgBox("计算序号：  " & n & "比例初始值：  " & HSSBFHL_CSZ)
                    '参与混水的设备的专用FHTJJD参数，目的是提高计算速度（风冷螺杆机、空气源热泵、水地源热泵）
                    Dim FHTJJD_HS As Double = 2 * FHTJJD
                    '循环的次数最大值
                    Dim JSCS_max As Integer = CType(HSSBFHL_CSZ * 100 / FHTJJD_HS, Integer)
                    'JS = JS + 1
                    For j = 0 To JSCS_max
                        '读取混水设备的制热负荷率
                        HSSBFHL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value
                        '进行一次负荷计算
                        'Dim b As Integer = i '工况序号
                        Call 清空指定工况输入输出数据(ExcelApp, i)
                        '进行正常的负荷分析（主要技术指标）计算
                        Call 负荷分析计算程序(ExcelApp, i, FHTJJD, JSBC， D_price_GF1, D_price_GF2, D_price_F1, D_price_F2, D_price_P1, D_price_P2, D_price_G1, D_price_G2, D_price_QT1, D_price_QT2, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '计算循环体，计算一次当前工况
                        '读取计算输入量
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 109)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 109)).Value
                        '返回计算结果
                        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(7 + i, 54)).Value = ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 54)).Value
                        '返回内燃机及其余热利用计算结果
                        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 15), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 18)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 15), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 18)).Value
                        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 23), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 30)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 23), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 30)).Value
                        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 54), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 57)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 54), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 57)).Value
                        ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 62), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(8 + i, 69)).Value = ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Range(ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 62), ExcelApp.ThisWorkbook.Worksheets("内燃机及其余热利用计算结果").Cells(3, 69)).Value
                        '返回设备运行信息结果
                        ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range(ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 2), ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(7 + i, 63)).Value = ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Range(ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("设备运行信息汇总").Cells(3, 63)).Value
                        '再次计算天然气锅炉、直燃型溴化锂、电采暖锅炉实际总供热功率+蓄热功率(仅电采暖锅炉有蓄热功率)（实际总供热功率也就是两种设备混水供热的总功率，然后按照输入的负荷比例进行分配）（此处读取仅用作初试计算，计算一个初试量）
                        TRQGLSJGL = TRQGL1ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 52).Value + TRQGL2ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 53).Value
                        ZRXHLSJGL = ZRXHL1ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 92).Value + ZRXHL2ZRGL * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 93).Value
                        DCNGLSJGL = DCNGL1ZRGL * (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 94).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 96).Value) + DCNGL2ZRGL * (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 95).Value + ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 97).Value)
                        '计算此时混水设备实际运行的功率
                        HSSBSJGL = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value * HSSBGL
                        '此时混水设备的实际功率比例
                        HSSBSJBL = HSSBSJGL / (TRQGLSJGL + ZRXHLSJGL + DCNGLSJGL + HSSBSJGL)
                        'MsgBox("计算序号：  " & i & "   风冷螺杆负荷率：  " & HSSBFHL(i) & "   风冷螺杆制热功率：  " & HSSBSJGL & "   电锅炉制热功率：   " & (TRQGLSJGL + ZRXHLSJGL + DCNGLSJGL) & "   当前混水比例：  " & HSSBSJBL(i))
                        '比较此时混水设备比例是否合格，设置跳出循环的条件
                        If Math.Abs(HSSBSJBL - HSGRGLBL) / HSGRGLBL < 0.02 Then
                            Exit For
                        End If
                        '排除极端情况
                        If HSSBSJBL <= 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = 0
                            'Exit For
                        End If
                        '根据计算出的实际混水比例，进行计算
                        If HSSBSJBL = HSGRGLBL Then
                            Exit For
                        End If
                        '调整负荷率
                        If HSSBSJBL < HSGRGLBL And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value >= 0 Then
                            '如果实际混水比例比设计的小，则混水设备制热负荷率往上加
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = HSSBFHL_CSZ + FHTJJD_HS * j / 100
                            '将混水设备负荷率增加并写入表格（目前的负荷率大于等于FHTJJD的情况下才继续增加）
                        ElseIf HSSBSJBL > HSGRGLBL And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value > 0 Then
                            '如果实际混水比例比设计的大，则混水设备制热负荷率往下减
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = HSSBFHL_CSZ - FHTJJD_HS * j / 100
                            '将混水设备负荷率减小并写入表格（目前的负荷率大于等于FHTJJD的情况下才继续减小）
                        End If
                        '如果出现负荷率小于0
                        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value < 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + i, 83).Value = 0
                        End If
                    Next
aaaaa:
                    '记录下算出来的混水负荷比例，并显示出来，供用户判断计算是否正确
                    HSBL（i - 1） = Math.Round(HSSBSJBL, 4)
                    '修正供热和蓄热时的耗电量和天然气耗量修正系数
                    '对计算出的制热设备负荷率进行修正，限制设备可以计算出的最低负荷率和最高负荷率
                    Call 制热和蓄热空调设备负荷率修正(ExcelApp, i, calculation_mode)
                    '计算制热季天然气耗量和耗电量综合修正系数
                    Call 制热季天然气消耗修正系数和设备本体耗电综合修正系数计算(ExcelApp, i, FHTJJD, calculation_mode)
                End If
            Next
            '计算循环体
            Call 计算循环体(ExcelApp, n)
            '将各工况混水比例显示出来
            Dim XianShi As String = Nothing
            For Each XXX In HSBL
                XianShi = XianShi & XXX.ToString & "  "
            Next
            MsgBox("用户输入的混水功率比例系数为：" & HSGRGLBL & vbCrLf & "各工况计算出来的混水功率比例系数为：" & XianShi)
        End If
    End Sub

End Module
