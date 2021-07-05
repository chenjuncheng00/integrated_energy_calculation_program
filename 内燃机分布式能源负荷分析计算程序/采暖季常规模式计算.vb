Module 采暖季常规模式计算
    Sub 常规计算模式制热和蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, XHLZR As Double, HSRFH As Double, calculation_mode As Integer, TJGRFHBL As Double, HSGRGLBL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '制热计算，常规计算模式
        Call 设备六种顺序制热常规模式计算(ExcelApp, b, FHTJJD, JSBC, XHLZR, HSRFH, calculation_mode)
        '蓄热计算，常规计算模式
        Call 蓄能装置蓄热工况常规模式计算(ExcelApp, b, FHTJJD, calculation_mode)
        '梯级供热计算
        Call 常规计算模式梯级供热计算(ExcelApp, b, calculation_mode, TJGRFHBL)
        '混水供热计算
        Call 常规计算模式混水供热计算(ExcelApp, FHTJJD, JSBC, b, calculation_mode, XHLZR, HSGRGLBL)
    End Sub

    Sub 设备六种顺序制热常规模式计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, XHLZR As Double, HSRFH As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '根据输入的计算模式类型，只有为类型1时，才进行计算
        If calculation_mode = 1 And RFHZXQL(b) > 0 Then
            '阶段1的总制热功率=混水供热+蓄能供热+烟气热水型溴化锂供热
            Dim JD1ZZRL As Double = HSRFH + XNGRGL(b) + XHLZR
            '制热顺序一至顺序三计算
            '进行供热负荷段计算
            '进行制热第一顺序机器负荷率计算
            '判断第一顺序使用的供热设备类型
            '优先使用蓄热装置可以供热的部分和溴化锂供热的部分，其次优先使用第一顺序制热设备
            '第一顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '第一顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZRL)
            End If
            '蓄热+溴化锂+混水热负荷+第一顺序机器总制热量--阶段2总制热量
            Dim JD2ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '进行制热第二顺序机器负荷率计算
            '判断第二顺序使用的供热设备类型
            '优先使用蓄热装置可以供热的部分和溴化锂供热的部分，其次优先使用第二顺序制热设备
            '第二顺序使用的供热设备类型为天然气采暖锅炉时
            '第二顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '第二顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZRL)
            End If
            '蓄热+溴化锂+第二顺序机器总制热量+第二顺序机器总制热量--阶段3总制热量
            Dim JD3ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————————————————————
            '进行制热第三顺序机器负荷率计算
            '判断第三顺序使用的供热设备类型
            '优先使用蓄热装置可以供热的部分和溴化锂供热的部分，其次优先使用第三顺序制热设备
            '第三顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '第三顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZRL)
            End If
            '蓄热+溴化锂+第一顺序机器总制热量+第二顺序机器总制热量+第三顺序机器总制热量--阶段4总制热量
            Dim JD4ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————
            '第四顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '第四顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZRL)
            End If
            '蓄热+溴化锂+第一顺序机器总制热量+第二顺序机器总制热量+第三顺序机器总制热量+第四顺序机器总制热量--阶段5总制热量
            Dim JD5ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————
            '第五顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '第五顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZRL)
            End If
            '蓄热+溴化锂+第一顺序机器总制热量+第二顺序机器总制热量+第三顺序机器总制热量+第四顺序机器总制热量+第五顺序机器总制热量--阶段6总制热量
            Dim JD6ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————
            '第六顺序使用的供热设备类型为天然气采暖锅炉时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "天然气锅炉" Then
                Call 天然气锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '第六顺序使用的供热设备类型为电采暖锅炉机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "电采暖锅炉" Then
                Call 电采暖锅炉常规模式供热计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZRL)
            End If
            '蓄热+溴化锂+第一顺序机器总制热量+第二顺序机器总制热量+第三顺序机器总制热量+第四顺序机器总制热量+第五顺序机器总制热量+第六顺序机器总制热量--阶段7总制热量
            Dim JD7ZZRL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(84, 4).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————————————
            '如果算出来的负荷率为1或者0，赋值到表格中
            For c2 = 52 To 61
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value = 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, c2).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value
                End If
            Next
            For c2 = 92 To 97
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value = 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, c2).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c2).Value
                End If
            Next
        End If
    End Sub

    Sub 天然气锅炉常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断天然气采暖锅炉(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 10).Value <> "ZLJ000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置天然气采暖锅炉(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 52).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-上一阶段提供的热负荷）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 52).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断天然气采暖锅炉(2)是否开启
        '天然气采暖锅炉（1）全部最大负荷运行后，使用天然气采暖锅炉（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 10).Value <> "ZLJ000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 10).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 10).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 10).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置天然气采暖锅炉(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 53).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-天然气采暖锅炉（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 10).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 53).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 风冷螺杆式热泵常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断风冷螺杆机(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 12).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置风冷螺杆机(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '风冷螺杆机(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 54).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 54).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断风冷螺杆机(2)是否开启
        '风冷螺杆机（1）全部最大负荷运行后，使用风冷螺杆机（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 12).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 12).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 12).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 12).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置风冷螺杆机(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 55).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-风冷螺杆机（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 12).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 55).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub
    Sub 空气源热泵常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断空气源热泵(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 16).Value <> "R410a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置空气源热泵(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '空气源热泵(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 56).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 56).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断空气源热泵(2)是否开启
        '空气源热泵（1）全部最大负荷运行后，使用空气源热泵（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 16).Value <> "R410a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 16).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 16).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 16).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置空气源热泵(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 57).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-空气源热泵（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 16).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 57).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 水_地源热泵常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断水(地)源热泵机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 14).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置水(地)源热泵机组(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '水(地)源热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 58).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 58).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断水(地)源热泵机组(2)是否开启
        '水(地)源热泵机组（1）全部最大负荷运行后，使用水(地)源热泵机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 14).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 14).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 14).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 14).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置水(地)源热泵机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 59).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-水(地)源热泵机组（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 14).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 59).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub
    Sub 离心式热泵常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断离心式热泵机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 18).Value <> "Lx000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式热泵机组(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '离心式热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 60).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 60).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断离心式热泵机组(2)是否开启
        '离心式热泵机组（1）全部最大负荷运行后，使用离心式热泵机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 18).Value <> "Lx000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 18).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 18).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 18).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式热泵机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 61).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-离心式热泵机组（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 18).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 61).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 直燃型溴化锂常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断直燃型溴化锂机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 20).Value <> "LDF-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置直燃型溴化锂机组(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '离心式热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 92).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 92).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断直燃型溴化锂机组(2)是否开启
        '直燃型溴化锂机组（1）全部最大负荷运行后，使用直燃型溴化锂机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 20).Value <> "LDF-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 20).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 20).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 20).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置直燃型溴化锂机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 93).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-离心式热泵机组（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 20).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 93).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 电采暖锅炉常规模式供热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZRL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim RFHLXX1, RFHLXX2 As Double '设备负荷率下限
        Dim RFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim RFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZRL：上一个阶段已经提供的制热量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断电采暖锅炉机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 22).Value > 0 Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置电采暖锅炉机组(1)负荷率，并判断是否超过最大热负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                RFHL1 = RFHLXX1 + (FHTJJD / 100) * b1 '电采暖锅炉机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 94).Value = RFHL1
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热）
                If (RFHZXQL(b) - JDZRL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 94).Value = RFHL1
                    Exit For
                End If
            Next
        End If
        '判断电采暖锅炉机组(2)是否开启
        '电采暖锅炉机组（1）全部最大负荷运行后，使用电采暖锅炉机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 22).Value > 0 Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = 0.5 '负荷率为0.5
            If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 22).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = 0.25 '负荷率等于0.25
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 22).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    RFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    RFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = 0.75 '负荷率等于0.75
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 22).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    RFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    RFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置电采暖锅炉机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                RFHL2 = RFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 95).Value = RFHL2
                '用于判断的负荷是（热负荷总需求量-蓄热供热量-溴化锂制热-电采暖锅炉机组（1）供热量）
                If (RFHZXQL(b) - JDZRL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 22).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 95).Value = RFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 蓄能装置蓄热工况常规模式计算(ExcelApp As Object, b As Integer, FHTJJD As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的计算模式类型，只有为类型1时，才进行计算
        If calculation_mode = 1 Then
            '蓄热，蓄热，蓄热，蓄热，蓄热
            '蓄热装置蓄热状态计算
            '蓄热时不同空调机的启动顺序和制热时一样
            If XNXRGL(b) > 0 Then
                '将本工况蓄能量和蓄能时间带入计算一次
                '蓄热量
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 23).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 23).Value
                '蓄热时间
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 24).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 26).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 26).Value
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取采暖季装机方案及参数
                Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
                '天然气锅炉
                Dim NUM1_TRQGL As Double = ans_ZJFA_R(10)
                Dim NUM2_TRQGL As Double = ans_ZJFA_R(11)
                '电锅炉
                Dim NUM1_DGL As Double = ans_ZJFA_R(18)
                Dim NUM2_DGL As Double = ans_ZJFA_R(19)
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
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第一顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 15).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第二顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 16).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第三顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 17).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第四顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 18).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第五顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 19).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第六顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 20).Value = "电采暖锅炉" Then
                    Call 电采暖锅炉常规模式蓄热计算(ExcelApp, b, FHTJJD, NUM1_DGL, NUM2_DGL)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '    '测算是否需要补充电制热
                '    If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0) Then
                '         For DZR = 1 To 10000000
                '              ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 70).Value = DZR
                '              If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                '                   ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 70).Value = DZR
                '                   Exit For
                '              End If
                '         Next
                '    End If
            ElseIf XNXRGL(b) = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 70)).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97)).Value = 0
            ElseIf XNXRGL(b) < 0 Then
                Dim XZ1
                XZ1 = MsgBox("蓄热负荷总需求量不能小于0", vbOKCancel)
            End If
        End If
    End Sub

    Sub 风冷螺杆式热泵常规模式蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄热启动数量
        Dim XRQDSL1, XRQDSL2 As Double
        '制热启动数量
        Dim ZRQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(58, 12).Value
        Dim ZRQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(77, 12).Value
        '装机总数量
        Dim ZRZJSL1 As Double = NUM1
        Dim ZRZJSL2 As Double = NUM2
        If (ZRZJSL1 - ZRQDSL1 > 0) Then
            For e1 = 0 To (ZRZJSL1 - ZRQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL1 = (FHTJJD / 100) * e1 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 62).Value = XRQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62).Value = XRQDSL1
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 62).Value = XRQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZRZJSL2 - ZRQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0)) Then
            For e2 = 0 To (ZRZJSL2 - ZRQDSL2) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL2 = (FHTJJD / 100) * e2 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 63).Value = XRQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 63).Value = XRQDSL2
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 63).Value = XRQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 空气源热泵常规模式蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄热启动数量
        Dim XRQDSL1, XRQDSL2 As Double
        '制热启动数量
        Dim ZRQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(58, 16).Value
        Dim ZRQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(77, 16).Value
        '装机总数量
        Dim ZRZJSL1 As Double = NUM1
        Dim ZRZJSL2 As Double = NUM2
        If (ZRZJSL1 - ZRQDSL1 > 0) Then
            For e1 = 0 To (ZRZJSL1 - ZRQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL1 = (FHTJJD / 100) * e1 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 64).Value = XRQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 64).Value = XRQDSL1
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 64).Value = XRQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZRZJSL2 - ZRQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0)) Then
            For e2 = 0 To (ZRZJSL2 - ZRQDSL2) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL2 = (FHTJJD / 100) * e2 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 65).Value = XRQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 65).Value = XRQDSL2
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 65).Value = XRQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 水_地源热泵常规模式蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄热启动数量
        Dim XRQDSL1, XRQDSL2 As Double
        '制热启动数量
        Dim ZRQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(58, 14).Value
        Dim ZRQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(77, 14).Value
        '装机总数量
        Dim ZRZJSL1 As Double = NUM1
        Dim ZRZJSL2 As Double = NUM2
        If (ZRZJSL1 - ZRQDSL1 > 0) Then
            For e1 = 0 To (ZRZJSL1 - ZRQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL1 = (FHTJJD / 100) * e1 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 66).Value = XRQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 66).Value = XRQDSL1
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 66).Value = XRQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZRZJSL2 - ZRQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0)) Then
            For e2 = 0 To (ZRZJSL2 - ZRQDSL2) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL2 = (FHTJJD / 100) * e2 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 67).Value = XRQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 67).Value = XRQDSL2
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 67).Value = XRQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 离心式热泵常规模式蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄热启动数量
        Dim XRQDSL1, XRQDSL2 As Double
        '制热启动数量
        Dim ZRQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(58, 18).Value
        Dim ZRQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(77, 18).Value
        '装机总数量
        Dim ZRZJSL1 As Double = NUM1
        Dim ZRZJSL2 As Double = NUM2
        If (ZRZJSL1 - ZRQDSL1 > 0) Then
            For e1 = 0 To (ZRZJSL1 - ZRQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL1 = (FHTJJD / 100) * e1 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 68).Value = XRQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value = XRQDSL1
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 68).Value = XRQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZRZJSL2 - ZRQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0)) Then
            For e2 = 0 To (ZRZJSL2 - ZRQDSL2) / (FHTJJD / 100) '确定最大计算数量
                XRQDSL2 = (FHTJJD / 100) * e2 '蓄热启动数量
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 69).Value = XRQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value = XRQDSL2
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 69).Value = XRQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 电采暖锅炉常规模式蓄热计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄热启动数量
        Dim XRQDSL1, XRQDSL2 As Double
        '装机总数量,蓄热启动数量
        Dim ZRQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(59, 22).Value '电锅炉（1）已经使用的负荷率
        Dim ZRQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(78, 22).Value '电锅炉（2）已经使用的负荷率
        Dim ZRZJSL1 = 1 '电锅炉（1）最大的负荷率为1
        Dim ZRZJSL2 = 1 '电锅炉（2）最大的负荷率为1
        If (ZRZJSL1 - ZRQDSL1 > 0) Then
            For e1 = 0 To (ZRZJSL1 - ZRQDSL1) Step (FHTJJD / 100)  '确定最大计算数量
                XRQDSL1 = e1 '电锅炉（1）蓄热的负荷率
                '将蓄热启动负荷率带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 96).Value = XRQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = XRQDSL1
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 96).Value = XRQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZRZJSL2 - ZRQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value = 0)) Then
            For e2 = 0 To (ZRZJSL2 - ZRQDSL2) Step (FHTJJD / 100)  '确定最大计算数量
                XRQDSL2 = e2 '电锅炉（2）蓄热的负荷率
                '将蓄热启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 97).Value = XRQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = XRQDSL2
                '设置跳出循环的条件
                '当机组蓄热负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(82, 24).Value > 0) Then
                    '记录下此时的蓄热设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 97).Value = XRQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

End Module
