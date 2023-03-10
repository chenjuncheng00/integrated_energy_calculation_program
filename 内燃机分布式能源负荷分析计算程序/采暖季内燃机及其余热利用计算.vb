Module 采暖季内燃机及其余热利用计算
    Sub 内燃机可以向外供电时制热设备运行计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '供热，供热，供热，供热，供热
        Dim XHLZR As Double = 0 '溴化锂制热量
        '当热负荷总需求量大于0时进行以下计算
        If RFHZXQL(b) > 0 Then
            '蓄能装置供热量和蓄能时间，带入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '制热时内燃机负荷率代入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value
            '计算内燃机及其余热利用
            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '此时没有溴化锂蓄热
            Dim XHLXR As Double = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        ElseIf RFHZXQL(b) = 0 Then
            '此时没有溴化锂蓄热
            Dim XHLXR As Double = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式（此时只有蓄热负荷）
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        ElseIf RFHZXQL(b) < 0 Then
            Dim XZ1
            XZ1 = MsgBox("热负荷总需求量不能小于0", vbOKCancel)
        End If
    End Sub

    Sub 内燃机不可以向外供电时制热设备运行计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '定义局部变量
        Dim ZRNRJFHL As Double '制热内燃机负荷率
        Dim ZRNRJFHL1 As Double '制热内燃机(1)负荷率
        Dim ZRNRJFHL2 As Double '制热内燃机(2)负荷率
        Dim XHLZR As Double = 0 '溴化锂制热量
        '供热，供热，供热，供热，供热
        '当热负荷总需求量大于0时进行以下计算
        If RFHZXQL(b) > 0 Then
            '蓄能装置供热量和蓄能时间，带入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '判断本工况总热负荷需求量（kW）是否小于等于蓄热装置总供热量（kW）
            '如果本工况总热负荷需求量（kW）是否大于蓄热装置总供热量（kW），则进行下列计算
            If RFHZXQL(b) > XNGRGL(b) Then
                '内燃机负荷率，从1递减
                '判断内燃机(1)和(2)是否都启用，如果两种都启用，则优先降低内燃机(2)的负荷率，当内燃机(2)的负荷率降至0时，再降低内燃机(1)的负荷率
                '如果只启用了一种内燃机，则可以使用下列代码进行计算
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                    '如果两台内燃机均没有被选择，则直接仅计算供热和蓄热设备（防止全局寻优计算模式时候速度过慢）
                    If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                        '如果两个内燃机均没有选择，则采用常规计算，防止计算速度太慢
                        '蓄能装置供热量和蓄能时间，带入计算
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
                        '制热时内燃机负荷率代入计算（负荷率都是0）
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                        '溴化锂(1)+(2)的总制热量=0
                        XHLZR = 0
                        '此时没有溴化锂蓄热
                        Dim XHLXR As Double = 0
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '制热和蓄热计算，常规计算模式
                        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '制热和蓄热计算，全局寻优计算模式
                        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    Else
                        '二分法提高计算速度
                        Dim NRJZRFHLSX = 内燃机不可以向外供电时制热设备运行计算仅启动一种内燃机二分法提高计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        For d = 0 To JSBC
                            '改变内燃发电机组负荷率
                            '两种内燃机的负荷率同时降低，因为有一种内燃机没有开启，所以修改这种内燃机的负荷率对计算并不产生影响，在计算结束后将没有启动的那个内燃机负荷率设置为0即可，从而简化代码
                            ZRNRJFHL = NRJZRFHLSX - (FHTJJD / 100) * d '制热内燃机负荷率
                            '制热时内燃机负荷率代入计算
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ZRNRJFHL
                            '计算内燃机及其余热利用
                            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '此时没有溴化锂蓄热
                            Dim XHLXR As Double = 0
                            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                            '制热和蓄热计算，常规计算模式
                            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                            '制热和蓄热计算，全局寻优计算模式
                            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            '设置跳出循环的条件，当向外供电量小于等于0，同时满足制热负荷需求时，跳出循环
                            If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                '将此时的内燃机负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                Exit For
                                '当向外供电量满足条件，但是制热负荷需求不满足条件时
                            ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                '将此时的内燃机负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                '修正计算
                                Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '当向外供电量不满足条件，同时不满足制热负荷需求时
                            ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                '将此时的内燃机负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                '修正计算
                                Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            End If
                        Next
                    End If
                End If
                '如果启用了两种内燃机，则可以使用下列代码进行计算
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value <> "J000GS") Then
                    '这种情况下优先降低内燃机(2)的负荷率，当内燃机(2)负荷率降至0后，再降低内燃机(1)负荷率
                    '先将内燃机(1)的负荷率设置为1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 1
                    '二分法提高内燃机（2）负荷率计算速度,代码在模块16
                    Dim NRJZRFHLSX2 = 内燃机不可以向外供电时制热设备运行计算启动两种内燃机二分法提高内燃机2负荷率计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    For d2 = 0 To JSBC
                        ZRNRJFHL2 = NRJZRFHLSX2 - (FHTJJD / 100) * d2 '先降低内燃机(2)负荷率
                        '将内燃机(2)的负荷率带入计算
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = ZRNRJFHL2
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                        '计算内燃机及其余热利用
                        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                        '记录下溴化锂(1)+(2)的总制热量
                        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                        '此时没有溴化锂蓄热
                        Dim XHLXR As Double = 0
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '制热和蓄热计算，常规计算模式
                        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '制热和蓄热计算，全局寻优计算模式
                        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '设置跳出循环的条件，当向外供电量小于等于0，同时满足制热负荷需求时，跳出循环
                        If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                            '将此时的内燃机(2)负荷率记录下来
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                            Exit For
                            '当向外供电量满足条件，但是制热负荷需求不满足条件时
                        ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                            '将此时的内燃机(2)负荷率记录下来
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                            '修正计算
                            Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            '当向外供电量不满足条件，同时不满足制热负荷需求时
                        ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                            '将此时的内燃机(2)负荷率记录下来
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                            '修正计算
                            Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        End If
                    Next
                    '判断是否需要修改内燃机(1)的负荷率
                    '当内燃机(2)负荷率为0，同时供电量不满足条件时，修改内燃机(1)的负荷率
                    If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value) Then
                        '二分法提高内燃机（1）负荷率计算速度,代码在模块16
                        Dim NRJZRFHLSX1 = 内燃机不可以向外供电时制热设备运行计算启动两种内燃机二分法提高内燃机1负荷率计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        For d1 = 0 To JSBC
                            ZRNRJFHL1 = NRJZRFHLSX1 - (FHTJJD / 100) * d1 '降低内燃机(1)负荷率
                            '将内燃机(1)和(2)的负荷率带入计算
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = ZRNRJFHL1
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value
                            '计算内燃机及其余热利用
                            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '此时没有溴化锂蓄热
                            Dim XHLXR As Double = 0
                            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                            '制热和蓄热计算，常规计算模式
                            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                            '制热和蓄热计算，全局寻优计算模式
                            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            '设置跳出循环的条件，当向外供电量小于等于0，同时满足制热负荷需求时，跳出循环
                            If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                '将此时的内燃机(1)负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                Exit For
                                '当向外供电量满足条件，但是制热负荷需求不满足条件时
                            ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                '将此时的内燃机(1)负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                '修正计算
                                Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '当向外供电量不满足条件，同时不满足制热负荷需求时
                            ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                '将此时的内燃机(1)负荷率记录下来
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                '修正计算
                                Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            End If
                        Next
                    End If
                End If
                '如果本工况总热负荷需求量（kW）是否小于等于蓄热装置总供热量（kW），则将本工况内燃机负荷率设置为0
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                '然后调用内燃机可以向外供电是的代码进行计算
                Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            End If
            '检查哪一种内燃机没有开启，将负荷率设置为0
            '内燃机（1）
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
            End If
            '内燃机（2）
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS" Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
            End If
        ElseIf RFHZXQL(b) = 0 Then
            '此时没有溴化锂蓄热
            Dim XHLXR As Double = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式（此时只有蓄热负荷）
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        ElseIf RFHZXQL(b) < 0 Then
            Dim XZ1
            XZ1 = MsgBox("热负荷总需求量不能小于0", vbOKCancel)
        End If
    End Sub
    Sub 内燃机不可以向外供电时制热计算修正(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim XHLZR As Double '溴化锂制热量
        '将此时的内燃机负荷率代入计算
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '此时没有溴化锂蓄热
        Dim XHLXR As Double = 0
        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
        '重新进行制热计算
        '制热和蓄热计算，常规计算模式
        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
        '制热和蓄热计算，全局寻优计算模式
        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
    End Sub
    Sub 内燃机不可以向外供电时制热和蓄热设备运行计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '定义局部变量
        Dim ZRNRJFHL As Double '制热内燃机负荷率
        Dim ZRNRJFHL1 As Double '制热内燃机(1)负荷率
        Dim ZRNRJFHL2 As Double '制热内燃机(2)负荷率
        Dim JSBC1 As Double '计算步长1
        Dim NRJZDFHL As Double '内燃机最大负荷率
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热
        '判断本工况热负荷需求量是否为0，不为0 则进行下列计算
        If RFHZXQL(b) >= 0 Then
            '蓄能装置供热量和蓄能时间，带入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '如果本工况总热负荷需求量（kW）是否等于蓄热装置总供热量（kW），且两者都不为0，则将本工况内燃机负荷率设置为0
            If (RFHZXQL(b) = XNGRGL(b) And RFHZXQL(b) <> 0 And XNGRGL(b) <> 0) Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                '然后调用内燃机可以向外供电是的代码进行计算
                Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                '判断本工况总热负荷需求量（kW）是否小于等于蓄热装置总供热量（kW）
                '如果本工况总热负荷需求量（kW）是否大于蓄热装置总供热量（kW），则进行下列计算
            Else
                '内燃机负荷率，从1递减
                '判断内燃机(1)和(2)是否都启用，如果两种都启用，则优先降低内燃机(2)的负荷率，当内燃机(2)的负荷率降至0时，再降低内燃机(1)的负荷率
                '如果只启用了一种内燃机，则可以使用下列代码进行计算
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                    '如果两台内燃机均没有被选择，则直接仅计算供热和蓄热设备（防止全局寻优计算模式时候速度过慢）
                    If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                        '如果两个内燃机都没有启动，则采用常规计算，加快计算速度
                        '蓄能装置供热量和蓄能时间，带入计算
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
                        '制热时内燃机负荷率代入计算（负荷率都是0）
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                        '溴化锂(1)+(2)的总制热量=0
                        XHLZR = 0
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '制热和蓄热计算，常规计算模式
                        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '制热和蓄热计算，全局寻优计算模式
                        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                    Else
                        '二分法提高计算速度
                        Dim NRJZRFHLSX = 内燃机不可以向外供电时制热和蓄热设备运行计算仅启动一种内燃机二分法提高计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        For d = 0 To JSBC
                            '改变内燃发电机组负荷率
                            '两种内燃机的负荷率同时降低，因为有一种内燃机没有开启，所以修改这种内燃机的负荷率对计算并不产生影响，在计算结束后将没有启动的那个内燃机负荷率设置为0即可，从而简化代码
                            ZRNRJFHL = NRJZRFHLSX - (FHTJJD / 100) * d '制热内燃机负荷率
                            '制热时内燃机负荷率代入计算
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ZRNRJFHL
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                            '计算内燃机及其余热利用
                            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
                            '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
                            If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
                                '将溴化锂蓄热的数据清空
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                '制热和蓄热计算，常规计算模式
                                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '制热和蓄热计算，全局寻优计算模式
                                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '设置循环跳出条件
                                '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                '如果有蓄热装置，则跳出循环条件为
                                If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                    '如果没有蓄热装置，则跳出循环条件为
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                    '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '当向外供电量不满足条件，同时不满足制热负荷需求时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                End If
                                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                                '如果溴化锂制热量大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
                            ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
                                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                                '溴化锂供热量等于本工况热负荷需求量
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                                If XNXRGL(b) > 0 Then
                                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                                End If
                                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                '制热和蓄热计算，常规计算模式
                                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '制热和蓄热计算，全局寻优计算模式
                                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '设置循环跳出条件
                                '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                '如果有蓄热装置，则跳出循环条件为
                                If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                    '如果没有蓄热装置，则跳出循环条件为
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                    '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '当向外供电量不满足条件，同时不满足制热负荷需求时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '将此时的内燃机负荷率记录下来
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                End If
                            End If
                        Next
                    End If
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
                '如果启用了两种内燃机，则可以使用下列代码进行计算
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value <> "J000GS") Then
                    '先将内燃机(1)(2)的负荷率设置为1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 1
                    JSBC1 = CInt(100 / FHTJJD)
                    NRJZDFHL = 1
                    For d1 = 0 To JSBC1 '内燃机(1)负荷率，作为外圈循环控制量
                        ZRNRJFHL1 = NRJZDFHL - (FHTJJD / 100) * d1 '制热内燃机负荷率
                        '将内燃机(1)负荷率带入计算，并记录此时的溴化锂总制热量
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = ZRNRJFHL1
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                        '判断是否需要改变内燃机(2)负荷率
                        '当内燃机(2)的负荷率大于0时，进入下面循环
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(78, 4).Value > 0 Then
                            '二分法提高计算速度，代码在模块18
                            Dim NRJZRFHLSX2 = 内燃机不可以向外供电时制热和蓄热设备运行计算启动两种内燃机二分法提高内燃机2负荷率计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                            For d2 = 0 To JSBC '内燃机(2)负荷率，作为内圈循环控制量
                                ZRNRJFHL2 = NRJZRFHLSX2 - (FHTJJD / 100) * d2 '制热内燃机负荷率
                                '将内燃机(2)负荷率带入计算，并记录此时的溴化锂总制热量
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = ZRNRJFHL2
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                '计算内燃机及其余热利用
                                Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                                Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                                '记录下溴化锂(1)+(2)的总制热量
                                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                                '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
                                '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
                                If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
                                    '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热
                                    '将溴化锂蓄热的数据清空
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '制热和蓄热计算，常规计算模式
                                    Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '制热和蓄热计算，全局寻优计算模式
                                    Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '设置跳出循环的条件
                                    '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                    '如果有蓄热装置，则跳出循环条件为
                                    If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                        '记录下此时的内燃机(2)负荷率
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        GoTo 2222
                                        '如果没有蓄热装置，则跳出循环条件为
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                        '记录下此时的内燃机(2)负荷率
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        GoTo 2222
                                        '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                        '将此时的内燃机(2)负荷率记录下来
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                        '修正计算
                                        Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                        '当向外供电量不满足条件，同时不满足制热负荷需求时
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                        '将此时的内燃机(2)负荷率记录下来
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                        '修正计算
                                        Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    End If
                                    '————————————————————————————————————————————————————————————————————————————————————————————————
                                    '如果溴化锂制热量大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
                                ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
                                    '如果溴化锂制热量大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
                                    '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                                    '溴化锂供热量等于本工况热负荷需求量
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                    '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                                    If XNXRGL(b) > 0 Then
                                        XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                                    End If
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '制热和蓄热计算，常规计算模式
                                    Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '制热和蓄热计算，全局寻优计算模式
                                    Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '设置跳出循环的条件
                                    '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                    '如果有蓄热装置，则跳出循环条件为
                                    If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                        '记录下此时的内燃机(2)负荷率
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        GoTo 2222
                                        '如果没有蓄热装置，则跳出循环条件为
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                        '记录下此时的内燃机(2)负荷率
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        GoTo 2222
                                        '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                        '将此时的内燃机(2)负荷率记录下来
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                        '修正计算
                                        Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                        '当向外供电量不满足条件，同时不满足制热负荷需求时
                                    ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                        '将此时的内燃机(2)负荷率记录下来
                                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                        '修正计算
                                        Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    End If
                                End If
                                '当内燃机（2）负荷率已经为0，且内燃机（1）负荷率为1时，使用二分法加快计算内燃机1负荷率，代码在模块19
                                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(78, 4).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(59, 4).Value = 1) Then
                                    Dim NRJZRFHLSX1 = 内燃机不可以向外供电时制热和蓄热设备运行计算启动两种内燃机二分法提高内燃机1负荷率计算速度(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    JSBC1 = JSBC
                                    NRJZDFHL = NRJZRFHLSX1
                                End If
                            Next
                            '————————————————————————————————————————————————————————————————————————————————————————————
                            '————————————————————————————————————————————————————————————————————————————————————————————
                        Else  '如果内燃机(2)负荷率已经变成了0，则执行下列计算
                            '将内燃机(1)负荷率带入计算，并记录此时的溴化锂总热冷量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = ZRNRJFHL1
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                            '计算内燃机及其余热利用
                            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
                            '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
                            If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
                                '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热
                                '将溴化锂蓄热的数据清空
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                '制热和蓄热计算，常规计算模式
                                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '制热和蓄热计算，全局寻优计算模式
                                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '设置跳出循环的条件
                                '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                '如果有蓄热装置，则跳出循环条件为
                                If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    GoTo 2222
                                    '如果没有蓄热装置，则跳出循环条件为
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    GoTo 2222
                                    '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '当向外供电量不满足条件，同时不满足制热负荷需求时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                End If
                                '————————————————————————————————————————————————————————————————————————————————————————————————
                                '如果溴化锂制热量大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
                            ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
                                '如果溴化锂制热量大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
                                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                                '溴化锂供热量等于本工况热负荷需求量
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                                If XNXRGL(b) > 0 Then
                                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                                End If
                                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                '制热和蓄热计算，常规计算模式
                                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '制热和蓄热计算，全局寻优计算模式
                                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                '设置跳出循环的条件
                                '当向外供电量小于等于0，同时满足制热负荷和蓄热需求时，跳出循环
                                '如果有蓄热装置，则跳出循环条件为
                                If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value = "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b) And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    GoTo 2222
                                    '如果没有蓄热装置，则跳出循环条件为
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value >= RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    GoTo 2222
                                    '当向外供电量满足条件，但是制热负荷需求不满足条件时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                    '当向外供电量不满足条件，同时不满足制热负荷需求时
                                ElseIf (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value > ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value < RFHZXQL(b)) Then
                                    '记录下此时的内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                                    '修正计算
                                    Call 内燃机不可以向外供电时制热计算修正(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                                End If
                            End If
                        End If
                    Next
2222:
                End If
            End If
            '检查哪一种内燃机没有开启，将负荷率设置为0
            '内燃机（1）
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
            End If
            '内燃机（2）
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS" Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
            End If
            'ElseIf RFHZXQL(b) = 0 Then
            '
            '    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61)).Value = 0
            '
        ElseIf RFHZXQL(b) < 0 Then
            Dim XZ1
            XZ1 = MsgBox("热负荷总需求量不能小于0", vbOKCancel)
            Call 锁定工作表(ExcelApp)
            Exit Sub
        End If
    End Sub
    Sub 内燃机可以向外供电时制热和蓄热设备运行计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '判断本工况热负荷需求量是否为0，不为0 则进行下列计算
        If RFHZXQL(b) >= 0 Then
            '蓄能装置供热量和蓄能时间，带入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '将用户输入的制热时内燃机负荷率代入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value
            '计算内燃机及其余热利用
            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '比较溴化锂制热量和蓄能装置供热量之和与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
            '如果溴化锂的制热量和蓄能装置供热量之和小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
            If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                '如果溴化锂制热量和蓄能装置供热量之和大于热负荷总需求量，则溴化锂不仅用于供热，还用于蓄热,此时所有的供热负荷均由溴化锂设备提供
            ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量减去蓄热装置供热量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            End If
        ElseIf RFHZXQL(b) < 0 Then
            Dim XZ1
            XZ1 = MsgBox("热负荷总需求量不能小于0", vbOKCancel)
            Call 锁定工作表(ExcelApp)
            Exit Sub
        End If
    End Sub
    Function 内燃机不可以向外供电时制热设备运行计算仅启动一种内燃机二分法提高计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热
        Dim NRJZRFHLSX As Double '内燃机制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        '内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时
        '二分法提高计算速度
        '同时改变改变内燃机（1）（2）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.5 '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '此时没有溴化锂蓄热
        Dim XHLXR As Double = 0
        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
        '制热和蓄热计算，常规计算模式
        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
        '制热和蓄热计算，全局寻优计算模式
        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.25 '内燃机负荷率为0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                NRJZRFHLSX = 0.25 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX = 0.5  '内燃机制热负荷率上限
            End If
        Else
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.75 '内燃机负荷率为0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                NRJZRFHLSX = 0.75 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX = 1    '内燃机制热负荷率上限
            End If
        End If
        '返回结果
        Return NRJZRFHLSX
    End Function
    Function 内燃机不可以向外供电时制热设备运行计算启动两种内燃机二分法提高内燃机2负荷率计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热
        Dim NRJZRFHLSX2 As Double '内燃机2制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        '内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2）
        '二分法提高计算速度
        '改变内燃机（2）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.5      '内燃机负荷率为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.5  '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '此时没有溴化锂蓄热
        Dim XHLXR As Double = 0
        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
        '制热和蓄热计算，常规计算模式
        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
        '制热和蓄热计算，全局寻优计算模式
        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.25      '内燃机负荷率为0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.25  '内燃机负荷率为0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                NRJZRFHLSX2 = 0.25 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX2 = 0.5  '内燃机制热负荷率上限
            End If
        Else
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.75      '内燃机负荷率为0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.75  '内燃机负荷率为0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                NRJZRFHLSX2 = 0.75 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX2 = 1    '内燃机制热负荷率上限
            End If
        End If
        '返回计算结果
        Return NRJZRFHLSX2
    End Function
    Function 内燃机不可以向外供电时制热设备运行计算启动两种内燃机二分法提高内燃机1负荷率计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热，制热
        Dim NRJZRFHLSX1 As Double '内燃机1制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        '内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1）
        '二分法提高计算速度
        '将计算好的内燃机（2）负荷率带入计算
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value
        '改变内燃机（1）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.5      '内燃机负荷率为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.5  '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '此时没有溴化锂蓄热
        Dim XHLXR As Double = 0
        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
        '制热和蓄热计算，常规计算模式
        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
        '制热和蓄热计算，全局寻优计算模式
        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.25      '内燃机负荷率为0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.25  '内燃机负荷率为0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                NRJZRFHLSX1 = 0.25 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX1 = 0.5  '内燃机制热负荷率上限
            End If
        Else
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.75      '内燃机负荷率为0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.75  '内燃机负荷率为0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                NRJZRFHLSX1 = 0.75 '内燃机制热负荷率上限
            Else
                NRJZRFHLSX1 = 1    '内燃机制热负荷率上限
            End If
        End If
        '返回计算结果
        Return NRJZRFHLSX1
    End Function
    Function 内燃机不可以向外供电时制热和蓄热设备运行计算仅启动一种内燃机二分法提高计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热
        Dim NRJZRFHLSX As Double '内燃机制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时，内燃机（1）+（2）同时
        '二分法提高计算速度
        '同时改变改变内燃机（1）（2）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.5           '内燃机负荷率为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.5   '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
        '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
        If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
            '将溴化锂蓄热的数据清空
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.25           '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.25   '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.75           '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.75   '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                    NRJZRFHLSX = 0.75 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX = 1    '内燃机制热负荷率上限
                End If
            End If
            '如果溴化锂制热量大于冷负荷总需求量，则溴化锂不仅用于供冷，还用于蓄冷,此时所有的供冷负荷均由溴化锂设备提供
        ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
            '溴化锂供热量等于本工况热负荷需求量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
            If XNXRGL(b) > 0 Then
                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
            End If
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.25           '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.25   '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.75           '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.75   '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                    NRJZRFHLSX = 0.75 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX = 1    '内燃机制热负荷率上限
                End If
            End If
        End If
        '返回计算结果
        Return NRJZRFHLSX
    End Function
    Function 内燃机不可以向外供电时制热和蓄热设备运行计算启动两种内燃机二分法提高内燃机2负荷率计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热
        Dim NRJZRFHLSX2 As Double '内燃机制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2），内燃机（2）
        '二分法提高计算速度
        '此时内燃机（1）负荷率为1
        '改变内燃机（2）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.5             '内燃机负荷率为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.5         '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
        '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
        If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
            '将溴化锂蓄热的数据清空
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.25             '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.25         '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX2 = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX2 = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.75             '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.75         '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                    NRJZRFHLSX2 = 0.75 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX2 = 1    '内燃机制热负荷率上限
                End If
            End If
            '如果溴化锂制热量大于冷负荷总需求量，则溴化锂不仅用于供冷，还用于蓄冷,此时所有的供冷负荷均由溴化锂设备提供
        ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
            '溴化锂供热量等于本工况热负荷需求量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
            If XNXRGL(b) > 0 Then
                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
            End If
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.25             '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.25         '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX2 = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX2 = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.75             '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.75         '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                    If XNXRGL(b) > 0 Then
                        '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                        Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                        '制热和蓄热计算，常规计算模式
                        Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                        '制热和蓄热计算，全局寻优计算模式
                        Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                            NRJZRFHLSX2 = 0.75 '内燃机制热负荷率上限
                        Else
                            NRJZRFHLSX2 = 1    '内燃机制热负荷率上限
                        End If
                    End If
                End If
            End If
        End If
        '返回计算结果
        Return NRJZRFHLSX2
    End Function
    Function 内燃机不可以向外供电时制热和蓄热设备运行计算启动两种内燃机二分法提高内燃机1负荷率计算速度(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim NRJZRFHLSX1 As Double '内燃机制热负荷率上限
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1），内燃机（1）
        '二分法提高计算速度
        '此时内燃机（2）负荷率为0
        '改变内燃机（1）负荷率
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.5             '内燃机负荷率为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.5         '内燃机负荷率为0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '比较溴化锂制热量与热负荷需求量的大小，根据比较结果，按照不同的方式进行计算
        '如果溴化锂的制热量小于等于热负荷总需求量，则溴化锂不参与蓄热只进行供热，进行下列计算
        If (XHLZR + XNGRGL(b)) <= RFHZXQL(b) Then
            '将溴化锂蓄热的数据清空
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.25             '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.25         '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX1 = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX1 = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.75             '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.75         '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '如果溴化锂的制热量小于等于冷负荷总需求量，则溴化锂不参与蓄冷只进行供冷
                '将溴化锂蓄热的数据清空
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                    NRJZRFHLSX1 = 0.75 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX1 = 1    '内燃机制热负荷率上限
                End If
            End If
            '如果溴化锂制热量大于冷负荷总需求量，则溴化锂不仅用于供冷，还用于蓄冷,此时所有的供冷负荷均由溴化锂设备提供
        ElseIf (XHLZR + XNGRGL(b)) > RFHZXQL(b) Then
            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
            '溴化锂供热量等于本工况热负荷需求量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
            If XNXRGL(b) > 0 Then
                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
            End If
            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            '制热和蓄热计算，常规计算模式
            Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
            '制热和蓄热计算，全局寻优计算模式
            Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
            If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.25             '内燃机负荷率为0.25
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.25         '内燃机负荷率为0.25
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '因为所有的供热负荷均有溴化锂提供，则将所有的制热空调负荷率设置为0
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率大了，或者向外供电量正好等于0
                    NRJZRFHLSX1 = 0.25 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX1 = 0.5  '内燃机制热负荷率上限
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.75             '内燃机负荷率为0.75
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.75         '内燃机负荷率为0.75
                '计算内燃机及其余热利用
                NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                '记录下溴化锂(1)+(2)的总制热量
                XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                '溴化锂供热量等于本工况热负荷需求量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                '制热和蓄热计算，常规计算模式
                Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                '制热和蓄热计算，全局寻优计算模式
                Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(85, 8).Value >= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(87, 8).Value Then '向外供电量大于0，说明内燃机负荷率小了，或者向外供电量正好等于0
                    NRJZRFHLSX1 = 0.75 '内燃机制热负荷率上限
                Else
                    NRJZRFHLSX1 = 1    '内燃机制热负荷率上限
                End If
            End If
        End If
        '返回计算结果
        Return NRJZRFHLSX1
    End Function
    Sub 内燃机可以向外供电且内燃机余热不可以浪费时供热蓄热内燃机负荷率调节(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '定义局部变量
        Dim ZRNRJFHL As Double '制热内燃机负荷率
        Dim ZRNRJFHL1 As Double '制热内燃机(1)负荷率
        Dim ZRNRJFHL2 As Double '制热内燃机(2)负荷率
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '将本工况蓄能量和蓄能时间带入计算一次
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 23).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 23).Value '蓄热量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
        '先将内燃机负荷率设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 1
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
        '如果大于，则自动调整内燃机负荷率
        If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
            '如果只启用了一种内燃机，则可以使用下列代码进行计算
            If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                '如果两台内燃机均没有被选择，则直接仅计算供热和蓄热设备（防止全局寻优计算模式时候速度过慢）
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS") Then
                    '如果两个内燃机均没有选择，则采用常规计算，防止计算速度太慢
                    '蓄能装置供热量和蓄能时间，带入计算
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
                    '制热时内燃机负荷率代入计算（负荷率都是0）
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                    '溴化锂(1)+(2)的总制热量=0
                    XHLZR = 0
                    '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                    Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                    '制热和蓄热计算，常规计算模式
                    Call 常规计算模式制热和蓄热计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode, TJGRFHBL, HSGRGLBL)
                    '制热和蓄热计算，全局寻优计算模式
                    Call 制热季设备制热和蓄热全局寻优计算(ExcelApp, b, FHTJJD, XHLZR, XHLXR, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                Else
                    '如果只有一种内燃机被选择
                    '如果计算模式为1，采用常规模式进行计算
                    If calculation_mode = 1 Then
                        '采用二分法提高计算速度
                        Dim ZRNRJFHLXX = 只有一种内燃机时二分法提高供热蓄热内燃机负荷率调节计算速度(ExcelApp, b, calculation_mode)
                        For g = 0 To JSBC
                            '两种内燃机的负荷率同时降低，因为有一种内燃机没有开启，所以修改这种内燃机的负荷率对计算并不产生影响，在计算结束后将没有启动的那个内燃机负荷率设置为0即可，从而简化代码
                            ZRNRJFHL = ZRNRJFHLXX + (FHTJJD / 100) * g '制热内燃机负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ZRNRJFHL
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                            '计算内燃机及其余热利用
                            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                            '溴化锂供热量等于本工况热负荷需求量减去蓄热装置供热量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                            If XNXRGL(b) > 0 Then
                                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                            End If
                            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                            '设置跳出条件
                            '如果有蓄热
                            If XNXRGL(b) > 0 Then
                                If ((XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0)) Then
                                    '记录下此时内燃机负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                End If
                            Else '如果没有蓄热
                                If (XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) Then
                                    '记录下此时内燃机负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = ZRNRJFHL
                                    Exit For
                                End If
                            End If
                        Next
                    ElseIf calculation_mode = 2 Then
                        '本模块，当存在两种内燃机，且计算模式=2时，采用全局寻优计算内燃机的负荷率
                        '计算内燃机负荷率
                        Dim FHL_result = 内燃机可以向外供电且余热不可以被浪费是采暖季寻优计算(ExcelApp, b, FHTJJD)
                        Dim FHL1_result As Double = FHL_result(0)
                        Dim FHL2_result As Double = FHL_result(1)
                        '如果某个内燃机不存在，负荷率改为0
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Then
                            FHL1_result = 0
                        End If
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS" Then
                            FHL2_result = 0
                        End If
                        '将负荷率结果写入Excel
                        '内燃机（1）
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = FHL1_result
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = FHL1_result
                        '内燃机（2）
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = FHL2_result
                        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = FHL2_result
                    End If
                End If
            End If
            '如果启用了两种内燃机，则可以使用下列代码进行计算
            If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value <> "J000GS") Then
                '如果计算模式为1，采用常规模式进行计算
                If calculation_mode = 1 Then
                    '优先改变内燃机1负荷率
                    '先将内燃机（1）负荷率设置为1，内燃机（2）负荷率设置为0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 1
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
                    '计算内燃机及其余热利用
                    NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                    NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                    '记录下溴化锂(1)+(2)的总制热量
                    XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                    '判断，负荷率大了，则内燃机（2）保持0，调整内燃机（1）负荷率
                    If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                        Dim ZRNRJFHLXX1 = 启动两种内燃机时二分法提高供热蓄热内燃机1负荷率调节计算速度(ExcelApp, b, calculation_mode)
                        For g1 = 0 To JSBC
                            ZRNRJFHL1 = ZRNRJFHLXX1 + (FHTJJD / 100) * g1 '制热内燃机1负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = ZRNRJFHL1
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                            '计算内燃机及其余热利用
                            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                            '溴化锂供热量等于本工况热负荷需求量减去蓄热装置供热量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                            If XNXRGL(b) > 0 Then
                                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                            End If
                            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                            '设置跳出条件
                            '如果有蓄热
                            If XNXRGL(b) > 0 Then
                                If ((XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0)) Then
                                    '记录下此时内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    Exit For
                                End If
                            Else '如果没有蓄热
                                If (XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) Then
                                    '记录下此时内燃机(1)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = ZRNRJFHL1
                                    Exit For
                                End If
                            End If
                        Next
                    Else '负荷率小了，则内燃机（1）负荷率保持（1），调整内燃机（2）负荷率
                        Dim ZRNRJFHLXX2 = 启动两种内燃机时二分法提高供热蓄热内燃机2负荷率调节计算速度(ExcelApp, b, calculation_mode)
                        For g2 = 0 To JSBC
                            ZRNRJFHL2 = ZRNRJFHLXX2 + (FHTJJD / 100) * g2 '制热内燃机2负荷率
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = ZRNRJFHL2
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                            '计算内燃机及其余热利用
                            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
                            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
                            '记录下溴化锂(1)+(2)的总制热量
                            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
                            '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
                            '溴化锂供热量等于本工况热负荷需求量减去蓄热装置供热量
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                            '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                            If XNXRGL(b) > 0 Then
                                XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b))
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                            End If
                            '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                            Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
                            '设置跳出条件
                            '如果有蓄热
                            If XNXRGL(b) > 0 Then
                                If ((XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0)) Then
                                    '记录下此时内燃机(2)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                    Exit For
                                End If
                            Else '如果没有蓄热
                                If (XNGRGL(b) + XHLZR >= RFHZXQL(b) + XNXRGL(b)) Then
                                    '记录下此时内燃机(2)负荷率
                                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = ZRNRJFHL2
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                ElseIf calculation_mode = 2 Then
                    '本模块，当存在两种内燃机，且计算模式=2时，采用全局寻优计算内燃机的负荷率
                    '计算内燃机负荷率
                    Dim FHL_result = 内燃机可以向外供电且余热不可以被浪费是采暖季寻优计算(ExcelApp, b, FHTJJD)
                    Dim FHL1_result As Double = FHL_result(0)
                    Dim FHL2_result As Double = FHL_result(1)
                    '将负荷率结果写入Excel
                    '内燃机（1）
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = FHL1_result
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = FHL1_result
                    '内燃机（2）
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = FHL2_result
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = FHL2_result
                End If
            End If
        End If
        '检查哪一种内燃机没有开启，将负荷率设置为0
        '内燃机（1）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value = "J000GS" Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0
        End If
        '内燃机（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value = "J000GS" Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
        End If
    End Sub
    Function 只有一种内燃机时二分法提高供热蓄热内燃机负荷率调节计算速度(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ZRNRJFHLXX As Double '制热内燃机负荷率下限
        Dim XHLZR As Double '溴化锂制热量
        '先将内燃机负荷率设置为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
        '大于，说明负荷率大了
        If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
            '内燃机负荷率设置为0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX = 0 '制热内燃机负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX = 0.25 '制热内燃机负荷率下限
            End If
        Else '小于，说明负荷率小了
            '内燃机负荷率设置为0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = 0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value = 0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX = 0.5 '制热内燃机负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX = 0.75 '制热内燃机负荷率下限
            End If
        End If
        '返回计算结果
        Return ZRNRJFHLXX
    End Function
    Function 启动两种内燃机时二分法提高供热蓄热内燃机1负荷率调节计算速度(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ZRNRJFHLXX1 As Double '制热内燃机1负荷率下限
        Dim XHLZR As Double '溴化锂制热量
        '内燃机（2）负荷率设置为0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0
        '内燃机(1)负荷率设置为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
        '大于，说明负荷率大了
        If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
            '内燃机(1)负荷率设置为0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX1 = 0 '制热内燃机1负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX1 = 0.25 '制热内燃机1负荷率下限
            End If
        Else '小于，说明负荷率小了
            '内燃机负荷率设置为0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX1 = 0.5 '制热内燃机1负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX1 = 0.75 '制热内燃机1负荷率下限
            End If
        End If
        '返回计算结果
        Return ZRNRJFHLXX1
    End Function
    Function 启动两种内燃机时二分法提高供热蓄热内燃机2负荷率调节计算速度(ExcelApp As Object, b As Integer, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ZRNRJFHLXX2 As Double '制热内燃机2负荷率下限
        Dim XHLZR As Double '溴化锂制热量
        '内燃机（1）负荷率设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4).Value = 1
        '内燃机(2)负荷率设置为0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.5
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.5
        '计算内燃机及其余热利用
        Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
        Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
        '记录下溴化锂(1)+(2)的总制热量
        XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
        '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
        '大于，说明负荷率大了
        If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
            '内燃机(2)负荷率设置为0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.25
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.25
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX2 = 0 '制热内燃机2负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX2 = 0.25 '制热内燃机2负荷率下限
            End If
        Else '小于，说明负荷率小了
            '内燃机负荷率设置为0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 0.75
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5).Value = 0.75
            '计算内燃机及其余热利用
            NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '判断此时蓄能装置供热量和溴化锂总供热量是否大于热负荷总需求量和蓄热总量之和
            '大于，说明负荷率大了
            If (XNGRGL(b) + XHLZR > RFHZXQL(b) + XNXRGL(b)) Then
                ZRNRJFHLXX2 = 0.5 '制热内燃机2负荷率下限
            Else '小于，说明负荷率小了
                ZRNRJFHLXX2 = 0.75 '制热内燃机2负荷率下限
            End If
        End If
        '返回计算结果
        Return ZRNRJFHLXX2
    End Function
    Sub 内燃机可以向外供电且内燃机余热可以浪费时供热和蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, HSRFH As Double, D_price As Double, TRQ_price As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '定义局部变量
        Dim XHLZRLFL As Double '溴化锂制热被浪费的量
        Dim XHLZR As Double '溴化锂制热量
        Dim XHLXR As Double '溴化锂蓄热量
        '余热可以浪费
        '制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热，制热和蓄热
        '判断本工况热负荷需求量是否为0，不为0 则进行下列计算
        If RFHZXQL(b) >= 0 Then
            '蓄能装置供热量和蓄能时间，带入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 22).Value '蓄能装置供热
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '将本工况蓄能量和蓄能时间带入计算一次
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 23).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 23).Value '蓄热量
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value '蓄热时间
            '用户输入的制热时内燃机负荷率代入计算
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 4), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 5)).Value
            '计算内燃机及其余热利用
            Dim NRJFHL1 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value
            Dim NRJFHL2 = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value
            '记录下溴化锂(1)+(2)的总制热量
            XHLZR = 制热季内燃机及其余热利用系统计算(ExcelApp, b, NRJFHL1, NRJFHL2, calculation_mode)(0)
            '如果此时溴化锂制热量与蓄能装置供热功率之和小于热负荷总需求量和蓄能装置蓄热量之和，则内燃机余热不会被浪费，全部用于供热和蓄热
            If (XHLZR + XNGRGL(b)) <= (RFHZXQL(b) + XNXRGL(b)) Then
                '如果蓄热量为0，热负荷需求量大于0，则溴化锂的制热量只用于供热
                If (XNXRGL(b) = 0 And RFHZXQL(b) > 0) Then
                    '进行制热设备计算
                    Call 内燃机可以向外供电时制热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                End If
                '如果蓄热量大于0，热负荷需求量大于等于0，则溴化锂制热量不仅用于供热，也可以用于蓄热
                If (XNXRGL(b) > 0 And RFHZXQL(b) >= 0) Then
                    '进行计算
                    Call 内燃机可以向外供电时制热和蓄热设备运行计算(ExcelApp, b, FHTJJD, JSBC, HSRFH, D_price, TRQ_price, calculation_mode, TJGRFHBL, HSGRGLBL)
                End If
            Else '如果此时溴化锂制热量与蓄能装置供热功率之和大于热负荷总需求量和蓄能装置蓄热量之和，则内燃机余热会被浪费
                '此时所有的供热负荷和蓄热负荷均有溴化锂提供，溴化锂多余的能量浪费掉
                '溴化锂制热量分成三部分，一部分浪费掉，一部分用于供热，一部分用于蓄热
                '溴化锂制热量被浪费的量等于溴化锂总制热量减去热负荷需求量减去蓄能蓄热量
                XHLZRLFL = XHLZR - (RFHZXQL(b) - XNGRGL(b)) - 1.03 * XNXRGL(b) '溴化锂制热浪费量，蓄能装置蓄能量放大百分之3，因为考虑到百分之2.5的蓄能散热损失修正
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 76).Value = XHLZRLFL
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 76).Value = XHLZRLFL
                '溴化锂供热量等于本工况热负荷总需求量减去蓄热装置供热量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
                '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
                If XNXRGL(b) > 0 Then
                    XHLXR = XHLZR - (RFHZXQL(b) - XNGRGL(b)) - XHLZRLFL
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
                End If
                '将已有的制热设备运行负荷率，蓄热工况设备启动情况清空
                Call 清空制热和蓄热设备负荷率计算结果(ExcelApp, b)
            End If
        ElseIf RFHZXQL(b) < 0 Then
            Dim XZ1
            XZ1 = MsgBox("热负荷总需求量不能小于0", vbOKCancel)
        End If
    End Sub
    Function 内燃机可以向外供电且余热不可以被浪费是采暖季寻优计算(ExcelApp As Object, b As Integer, FHTJJD As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '设备可以允许运行的负荷率下限（内燃机和溴化锂是同一个负荷率）（不除以设备数量）
        Dim FHL1_min As Double = FHL1_min_NRJ
        Dim FHL2_min As Double = FHL2_min_NRJ
        '读取各种修正系数
        Dim ans_XZXS_R = 读取采暖季输入的计算系数(ExcelApp)
        '溴化锂制热COP
        Dim XHLZR_COP As Double = ans_XZXS_R(9)
        '读取采暖季装机方案及参数
        Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
        '内燃机
        '内燃机（1）数量
        Dim NUM1_NRJ As Double = ans_ZJFA_R(0)
        '内燃机（2）数量
        Dim NUM2_NRJ As Double = ans_ZJFA_R(1)
        '内燃机（1）最大余热功率（单台）
        Dim YRGL1_MAX_NRJ As Double = ans_ZJFA_R(4)
        '内燃机（2）最大余热功率（单台）
        Dim YRGL2_MAX_NRJ As Double = ans_ZJFA_R(5)
        '内燃机（1）天然气耗量最大值（单台）
        Dim TRQ1_MAX_NRJ As Double = ans_ZJFA_R(6)
        '内燃机（2）天然气耗量最大值（单台）
        Dim TRQ2_MAX_NRJ As Double = ans_ZJFA_R(7)
        '溴化锂（1）最大制热功率（单台）
        Dim ZRGL1_MAX_XHL As Double = YRGL1_MAX_NRJ * XHLZR_COP
        '溴化锂（2）最大制热功率（单台）
        Dim ZRGL2_MAX_XHL As Double = YRGL2_MAX_NRJ * XHLZR_COP
        '列表，储存总耗天然气量，溴化锂总制热功率，设备（1）负荷率和设备（2）负荷率
        '内燃机寻优仅考虑天然气耗量，不考虑辅机耗电量
        Dim HQ_ALL As New List(Of Double）
        Dim XHLZR_ALL As New List(Of Double)
        Dim FHL1 As New List(Of Double）
        Dim FHL2 As New List(Of Double）
        '列表，储存设备（1）和设备（2）的启动数量
        Dim NUM1_List As New List(Of Double)
        Dim NUM2_List As New List(Of Double)
        '穷举设备数量
        For n1 = 0 To NUM1_NRJ Step 1 '设备1数量
            For n2 = 0 To NUM2_NRJ Step 1 '设备2数量
                '判断启动的设备100%负荷率够不够，不够跳出循环，提高速度
                Dim RFH1_ALL_temp As Double
                If NUM1_NRJ = 0 Then
                    RFH1_ALL_temp = 0
                Else
                    RFH1_ALL_temp = n1 * ZRGL1_MAX_XHL
                End If
                Dim RFH2_ALL_temp As Double
                If NUM2_NRJ = 0 Then
                    RFH2_ALL_temp = 0
                Else
                    RFH2_ALL_temp = n2 * ZRGL2_MAX_XHL
                End If
                If (RFH1_ALL_temp + RFH2_ALL_temp) < (RFHZXQL(b) + XNXRGL(b)) Then
                    GoTo zzzz
                End If
                '穷举法求各种可能的组合的总耗电功率
                For a1 = FHL1_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(1)负荷率(单台)
                    For a2 = FHL2_min To (1 + 2 * FHTJJD / 100) Step FHTJJD / 100 'a1表示设备(2)负荷率(单台)
                        '内燃机发电效率修正系数
                        Dim FDXLXZ1_NRJ As Double = 内燃机发电效率曲线(a1)
                        Dim FDXLXZ2_NRJ As Double = 内燃机发电效率曲线(a2)
                        '内燃机余热效率修正系数
                        Dim YRXLXZ1_NRJ As Double = 内燃机余热效率曲线(a1)
                        Dim YRXLXZ2_NRJ As Double = 内燃机余热效率曲线(a2)
                        '溴化锂制热COP修正系数
                        Dim XHLZRXZ1_XHL As Double = 烟气热水型溴化锂制热COP曲线(a1)
                        Dim XHLZRXZ2_XHL As Double = 烟气热水型溴化锂制热COP曲线(a2)
                        '此时的天然气总耗量
                        Dim TRQ_ALL_now As Double = n1 * a1 * TRQ1_MAX_NRJ / FDXLXZ1_NRJ + n2 * a2 * TRQ2_MAX_NRJ / FDXLXZ2_NRJ
                        '此时的溴化锂制热总量
                        Dim XHLZR_ALL_now As Double = n1 * a1 * ZRGL1_MAX_XHL * YRXLXZ1_NRJ * XHLZRXZ1_XHL + n2 * a2 * ZRGL2_MAX_XHL * YRXLXZ2_NRJ * XHLZRXZ2_XHL
                        '设置跳出条件
                        '溴化锂制热功率大于等于热负荷总需求量+蓄能装置蓄热负荷
                        If XHLZR_ALL_now >= RFHZXQL(b) + XNXRGL(b) And XHLZR_ALL_now < (RFHZXQL(b) + XNXRGL(b)) * (1 + 4 * FHTJJD / 100) Then
                            '计算结果加入列表
                            HQ_ALL.Add(TRQ_ALL_now)
                            XHLZR_ALL.Add(XHLZR_ALL_now)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        ElseIf a1 >= 1 And a2 >= 1 Then
                            '如果设备(1)和设备(2)的负荷率都到了1，也跳出循环
                            '计算结果加入列表
                            HQ_ALL.Add(TRQ_ALL_now)
                            XHLZR_ALL.Add(XHLZR_ALL_now)
                            FHL1.Add(a1)
                            FHL2.Add(a2)
                            NUM1_List.Add(n1)
                            NUM2_List.Add(n2)
                            '跳出循环
                            Exit For
                        End If
                    Next
                Next
zzzz：
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '找出天然气耗量最低的工况作为计算结果
        Dim HQ_ALL_min As Double = HQ_ALL.Min
        Dim HQ_index_min As Integer = HQ_ALL.IndexOf(HQ_ALL_min)
        '溴化锂制热量
        Dim XHLZR_result As Double = XHLZR_ALL(HQ_index_min)
        '求此时设备（1）和设备（2）的负荷率
        Dim FHL1_result As Double = FHL1(HQ_index_min)
        Dim FHL2_result As Double = FHL2(HQ_index_min)
        '求此时设备（1）和设备（2）启动数量
        Dim NUM1_result As Double = NUM1_List(HQ_index_min)
        Dim NUM2_result As Double = NUM2_List(HQ_index_min)
        '溴化锂蓄热量计算
        '将此时的溴化锂制热量分成两部分，一部分用于向外供热，一部分用于蓄热
        '溴化锂供热量等于本工况热负荷需求量减去蓄热装置供热量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 72).Value = RFHZXQL(b) - XNGRGL(b)
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 72).Value = RFHZXQL(b) - XNGRGL(b)
        '溴化锂蓄热量等于溴化锂总制热量减去本工况热负荷需求量
        If XNXRGL(b) > 0 Then
            Dim XHLXR As Double = XHLZR_result - (RFHZXQL(b) - XNGRGL(b))
            If XHLXR > 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = XHLXR
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = XHLXR
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 74).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 74).Value = 0
            End If
        End If
        '返回负荷率计算结果
        Dim ans(4)
        ans(0) = FHL1_result
        ans(1) = FHL2_result
        ans(2) = NUM1_result
        ans(3) = NUM2_result
        Return ans
    End Function
    Function 制热季内燃机及其余热利用系统计算(ExcelApp As Object, b As Integer, ZRNRJFHL1 As Double, ZRNRJFHL2 As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim XHLZR1 As Double '溴化锂制热量1
        Dim XHLZR2 As Double '溴化锂制热量2
        Dim XHLZR As Double '溴化锂制热量
        '读取采暖季装机方案及参数
        Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
        '内燃机（1）装机数量
        Dim NUM1 As Double = ans_ZJFA_R(0)
        '内燃机（2）装机数量
        Dim NUM2 As Double = ans_ZJFA_R(1)
        '内燃机及其余热利用系统最低允许负荷率
        Dim FHL1_min As Double
        Dim FHL2_min As Double
        If NUM1 > 0 Then
            FHL1_min = FHL1_min_NRJ / NUM1
        Else
            FHL1_min = 0
        End If
        If NUM2 > 0 Then
            FHL2_min = FHL2_min_NRJ / NUM2
        Else
            FHL2_min = 0
        End If
        '函数返回的结果
        Dim ans(2) As Double
        '根据内燃机1和内燃机2的负荷率， 计算内燃机1、2的发电效率和余热效率相对100%负荷时的修正系数
        '发电效率修正系数
        Dim FDXLXZ1 As Double
        Dim FDXLXZ2 As Double
        '余热效率修正系数
        Dim YRXLXZ1 As Double
        Dim YRXLXZ2 As Double
        '计算烟气热水型溴化锂制冷COP修正系数
        Dim XHLZRCOPXZ1 As Double
        Dim XHLZRCOPXZ2 As Double
        '内燃机余热锅炉效率修正系数
        '生活热水余热锅炉 
        Dim SHRSYRGLXLXZ1 As Double
        Dim SHRSYRGLXLXZ2 As Double
        '工业蒸汽余热锅炉
        Dim GYZQYRGLXLXZ1 As Double
        Dim GYZQYRGLXLXZ2 As Double
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '只有全局寻优计算模式才进行计算
        If calculation_mode = 2 Then
            '根据内燃机1和内燃机2的负荷率，计算内燃机1、2的发电效率和余热效率相对100%负荷时的修正系数
            '内燃机1
            If NUM1 > 0 Then
                '发电效率修正系数
                FDXLXZ1 = 内燃机发电效率曲线(ZRNRJFHL1)
                '余热效率修正系数
                YRXLXZ1 = 内燃机余热效率曲线(ZRNRJFHL1)
                '计算烟气热水型溴化锂制冷COP修正系数
                XHLZRCOPXZ1 = 烟气热水型溴化锂制热COP曲线(ZRNRJFHL1)
                '内燃机余热锅炉效率修正系数
                '生活热水余热锅炉
                SHRSYRGLXLXZ1 = 生活热水余热锅炉效率曲线(ZRNRJFHL1)
                '工业蒸汽余热锅炉
                GYZQYRGLXLXZ1 = 工业蒸汽余热锅炉效率曲线(ZRNRJFHL1)
                '消除特殊情况
                If FDXLXZ1 > 1 And FDXLXZ1 < 1.001 Then
                    FDXLXZ1 = 1
                End If
                If YRXLXZ1 > 1 And YRXLXZ1 < 1.001 Then
                    YRXLXZ1 = 1
                End If
                If XHLZRCOPXZ1 > 1 And XHLZRCOPXZ1 < 1.001 Then
                    XHLZRCOPXZ1 = 1
                End If
                If SHRSYRGLXLXZ1 > 1 And SHRSYRGLXLXZ1 < 1.001 Then
                    SHRSYRGLXLXZ1 = 1
                End If
                If GYZQYRGLXLXZ1 > 1 And GYZQYRGLXLXZ1 < 1.001 Then
                    GYZQYRGLXLXZ1 = 1
                End If
            Else
                '全部等于1
                '发电效率修正系数
                FDXLXZ1 = 1
                '余热效率修正系数
                YRXLXZ1 = 1
                '计算烟气热水型溴化锂制冷COP修正系数
                XHLZRCOPXZ1 = 1
                '内燃机余热锅炉效率修正系数
                '生活热水余热锅炉
                SHRSYRGLXLXZ1 = 1
                '工业蒸汽余热锅炉
                GYZQYRGLXLXZ1 = 1
            End If
            '内燃机2
            If NUM2 > 0 Then
                '发电效率修正系数
                FDXLXZ2 = 内燃机发电效率曲线(ZRNRJFHL2)
                '余热效率修正系数
                YRXLXZ2 = 内燃机余热效率曲线(ZRNRJFHL2)
                '计算烟气热水型溴化锂制冷COP修正系数
                XHLZRCOPXZ2 = 烟气热水型溴化锂制热COP曲线(ZRNRJFHL2)
                '内燃机余热锅炉效率修正系数
                '生活热水余热锅炉
                SHRSYRGLXLXZ2 = 生活热水余热锅炉效率曲线(ZRNRJFHL2)
                '工业蒸汽余热锅炉
                GYZQYRGLXLXZ2 = 工业蒸汽余热锅炉效率曲线(ZRNRJFHL2)
                '消除特殊情况
                If FDXLXZ2 > 1 And FDXLXZ2 < 1.001 Then
                    FDXLXZ2 = 1
                End If
                If YRXLXZ2 > 1 And YRXLXZ2 < 1.001 Then
                    YRXLXZ2 = 1
                End If
                If XHLZRCOPXZ2 > 1 And XHLZRCOPXZ2 < 1.001 Then
                    XHLZRCOPXZ2 = 1
                End If
                If SHRSYRGLXLXZ2 > 1 And SHRSYRGLXLXZ2 < 1.001 Then
                    SHRSYRGLXLXZ2 = 1
                End If
                If GYZQYRGLXLXZ2 > 1 And GYZQYRGLXLXZ2 < 1.001 Then
                    GYZQYRGLXLXZ2 = 1
                End If
            Else
                '全部等于1
                '发电效率修正系数
                FDXLXZ2 = 1
                '余热效率修正系数
                YRXLXZ2 = 1
                '计算烟气热水型溴化锂制冷COP修正系数
                XHLZRCOPXZ2 = 1
                '内燃机余热锅炉效率修正系数
                '生活热水余热锅炉
                SHRSYRGLXLXZ2 = 1
                '工业蒸汽余热锅炉
                GYZQYRGLXLXZ2 = 1
            End If
        Else
            '全部等于1
            '发电效率修正系数
            FDXLXZ1 = 1
            FDXLXZ2 = 1
            '余热效率修正系数
            YRXLXZ1 = 1
            YRXLXZ2 = 1
            '计算烟气热水型溴化锂制冷COP修正系数
            XHLZRCOPXZ1 = 1
            XHLZRCOPXZ2 = 1
            '内燃机余热锅炉效率修正系数
            '生活热水余热锅炉
            SHRSYRGLXLXZ1 = 1
            SHRSYRGLXLXZ2 = 1
            '工业蒸汽余热锅炉
            GYZQYRGLXLXZ1 = 1
            GYZQYRGLXLXZ2 = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '根据计算出的内燃机余热效率修正系数、溴化锂COP修正系数、生活热水余热锅炉效率修正系数、工业蒸汽余热锅炉修正系数，计算实际的溴化锂制冷出力、生活热水出力和工业蒸汽出力修正系数
        '溴化锂制热出力修正系数
        Dim XHLZRCLXZ1 As Double = XHLZRCOPXZ1 * YRXLXZ1
        Dim XHLZRCLXZ2 As Double = XHLZRCOPXZ2 * YRXLXZ2
        '生活热水余热锅炉出力修正系数
        Dim SHRSCLXZ1 As Double = SHRSYRGLXLXZ1 * YRXLXZ1
        Dim SHRSCLXZ2 As Double = SHRSYRGLXLXZ2 * YRXLXZ2
        '工业蒸汽余热锅炉出力修正系数
        Dim GYZQCLXZ1 As Double = GYZQYRGLXLXZ1 * YRXLXZ1
        Dim GYZQCLXZ2 As Double = GYZQYRGLXLXZ2 * YRXLXZ2
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '将内燃机发电效率修正系数写入Excel
        '内燃机允许最低负荷率为0.3
        '内燃机（1）
        If ZRNRJFHL1 >= FHL1_min Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 102).Value = FDXLXZ1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 102).Value = FDXLXZ1
        Else
            '默认值为1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 102).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 102).Value = 1
        End If
        '内燃机（2）
        If ZRNRJFHL2 >= FHL2_min Then
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 103).Value = FDXLXZ2
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 103).Value = FDXLXZ2
        Else
            '默认值为1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 103).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 103).Value = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '将内燃机余热效率、溴化锂制冷COP、生活热余热锅炉效率、工业蒸汽效率修正系数写入Excel
        '读取当前选择的内燃机余热利用方式
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 77).Value = "溴化锂" Then
            '溴化锂（1）
            If ZRNRJFHL1 >= FHL1_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = XHLZRCLXZ1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = XHLZRCLXZ1
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = 1
            End If
            '溴化锂（2）
            If ZRNRJFHL2 >= FHL2_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = XHLZRCLXZ2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = XHLZRCLXZ2
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = 1
            End If
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 77).Value = "生活热水" Then
            '生活热水余热锅炉（1）
            If ZRNRJFHL1 >= FHL1_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = SHRSCLXZ1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = SHRSCLXZ1
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = 1
            End If
            '生活热水余热锅炉（2）
            If ZRNRJFHL2 >= FHL2_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = SHRSCLXZ2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = SHRSCLXZ2
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = 1
            End If
        ElseIf ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 77).Value = "工业蒸汽" Then
            '工业蒸汽余热锅炉（1）
            If ZRNRJFHL1 >= FHL1_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = GYZQCLXZ1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = GYZQCLXZ1
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = 1
            End If
            '工业蒸汽余热锅炉（2）
            If ZRNRJFHL2 >= FHL2_min Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = GYZQCLXZ2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = GYZQCLXZ2
            Else
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = 1
            End If
        Else
            '如果什么都没有选择，全部系数改为1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 102).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 102).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 103).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 103).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 104).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 104).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 105).Value = 1
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 105).Value = 1
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '记录下溴化锂制热量
        XHLZR1 = NUM1 * ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 4).Value
        XHLZR2 = NUM2 * ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 4).Value
        XHLZR = XHLZR1 + XHLZR2
        '返回计算结果
        ans(0) = XHLZR
        ans(1) = FDXLXZ1
        ans(2) = FDXLXZ2
        Return ans
    End Function

End Module
