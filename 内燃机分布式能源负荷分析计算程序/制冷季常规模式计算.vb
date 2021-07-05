Module 制冷季常规模式计算
    Sub 常规计算模式制冷和蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, XHLZL As Double, HSLFH As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '制冷计算,常规模式
        Call 设备六种顺序制冷常规模式计算(ExcelApp, b, FHTJJD, JSBC, XHLZL, HSLFH, calculation_mode)
        '蓄冷计算，常规模式
        Call 蓄能装置蓄冷工况常规模式计算(ExcelApp, b, FHTJJD, calculation_mode)
    End Sub
    Sub 设备六种顺序制冷常规模式计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, XHLZL As Double, HSLFH As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的计算模式类型，只有为类型1时，才进行计算
        If calculation_mode = 1 And LFHZXQL(b) > 0 Then
            '阶段1的总制冷功率=混水供冷+蓄能供冷+烟气热水型溴化锂供冷
            Dim JD1ZZLL As Double = HSLFH + XNGLGL(b) + XHLZL
            '制冷顺序一至顺序三计算
            '判断第一顺序使用的供冷设备类型
            '优先使用蓄冷装置可以供冷的部分和溴化锂供冷的部分，其次优先使用第一顺序制冷设备
            '第一顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '第一顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD1ZZLL)
            End If
            '蓄冷+溴化锂+混水冷负荷+第一顺序机器总制冷量计算——阶段2总制冷量
            Dim JD2ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————————————————————————
            '进行制冷第二顺序机器负荷率计算
            '判断第二顺序使用的供冷设备类型
            '优先使用蓄冷装置可以供冷的部分和溴化锂供冷的部分，其次优先使用第二顺序制冷设备
            '第二顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '第二顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD2ZZLL)
            End If
            '蓄冷+溴化锂+第一顺序机器总制冷量+第二顺序机器总制冷量计算--阶段3总制冷量
            Dim JD3ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '进行制冷第三顺序机器负荷率计算
            '判断第三顺序使用的供冷设备类型
            '优先使用蓄冷装置可以供冷的部分和溴化锂供冷的部分，其次优先使用第三顺序制冷设备
            '第三顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '第三顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD3ZZLL)
            End If
            '蓄冷+溴化锂+第一顺序机器总制冷量+第二顺序机器总制冷量+第三顺序机器总制冷量计算--阶段3总制冷量
            Dim JD4ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————
            '第四顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '第四顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD4ZZLL)
            End If
            '蓄冷+溴化锂+第一顺序机器总制冷量+第二顺序机器总制冷量+第三顺序机器总制冷量+第四顺序机器总制冷量计算--阶段5总制冷量
            Dim JD5ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————
            '第五顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '第五顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD5ZZLL)
            End If
            '蓄冷+溴化锂+第一顺序机器总制冷量+第二顺序机器总制冷量+第三顺序机器总制冷量+第四顺序机器总制冷量+第五顺序机器总制冷量计算--阶段6总制冷量
            Dim JD6ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————
            '第六顺序使用的供冷设备类型为离心式冷水机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "离心式冷水机" Then
                Call 离心式冷水机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为风冷螺杆式冷水(热泵)机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "风冷螺杆机" Then
                Call 风冷螺杆式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为空气源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "空气源热泵" Then
                Call 空气源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为水(地)源热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "水(地)源热泵" Then
                Call 水_地源热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为水冷螺杆机时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "水冷螺杆机" Then
                Call 水冷螺杆机常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为离心式热泵机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "离心式热泵" Then
                Call 离心式热泵常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '第六顺序使用的供冷设备类型为直燃型溴化锂机组时
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "直燃型溴化锂" Then
                Call 直燃型溴化锂常规模式供冷计算(ExcelApp, b, FHTJJD, JSBC, JD6ZZLL)
            End If
            '蓄冷+溴化锂+第一顺序机器总制冷量+第二顺序机器总制冷量+第三顺序机器总制冷量+第四顺序机器总制冷量+第五顺序机器总制冷量计算+第六顺序机器总制冷量计算--阶段7总制冷量
            Dim JD7ZZLL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 4).Value
            '————————————————————————————————————————————————————————————————————————————————————
            '如果算出来的负荷率为1或者0，赋值到表格中
            For c1 = 28 To 39
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value = 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, c1).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value
                End If
            Next
            For c1 = 90 To 91
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value = 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, c1).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, c1).Value
                End If
            Next
        End If
    End Sub

    Sub 离心式冷水机常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断离心式冷水机(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 10).Value <> "0RT" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value Then '负荷率等于0.25大了，或者恰好等于0.5
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式冷水机(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 28).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段已经提供的冷负荷）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 28).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断离心式冷水机(2)是否开启
        '离心式冷水机（1）全部最大负荷运行后，使用离心式冷水机（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 10).Value <> "0RT" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 10).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 10).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 10).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式冷水机(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 29).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段已经提供的冷负荷-离心式冷水机（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 10).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 10).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 29).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 风冷螺杆式热泵常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断风冷螺杆机(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 14).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置风冷螺杆机(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '风冷螺杆机(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 30).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段已经提供的冷负荷）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 30).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断风冷螺杆机(2)是否开启
        '风冷螺杆机（1）全部最大负荷运行后，使用风冷螺杆机（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 14).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 14).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 14).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 14).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置风冷螺杆机(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 31).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段已经提供的冷负荷-风冷螺杆机（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 14).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 14).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 31).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 空气源热泵常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断空气源热泵(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 20).Value <> "R410a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置空气源热泵(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '空气源热泵(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 32).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 32).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断空气源热泵(2)是否开启
        '空气源热泵（1）全部最大负荷运行后，使用空气源热泵（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 20).Value <> "R410a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 20).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 20).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 20).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置空气源热泵(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 33).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率-空气源热泵（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 20).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 20).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 33).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub
    Sub 水_地源热泵常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断水(地)源热泵机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 16).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置水(地)源热泵机组(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '水(地)源热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 34).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 34).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断水(地)源热泵机组(2)是否开启
        '水(地)源热泵机组（1）全部最大负荷运行后，使用水(地)源热泵机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 16).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 16).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 16).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 16).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置水(地)源热泵机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 35).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率-水(地)源热泵机组（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 16).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 16).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 35).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 水冷螺杆机常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断水冷螺杆机(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 12).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置水冷螺杆机(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '水冷螺杆机(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 36).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 36).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断水冷螺杆机(2)是否开启
        '水冷螺杆机（1）全部最大负荷运行后，使用水冷螺杆机（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 12).Value <> "R134a-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 12).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 12).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 12).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置水冷螺杆机2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 37).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率-水冷螺杆机（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 12).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 12).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 37).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 离心式热泵常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断离心式热泵机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 18).Value <> "Lx000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value Then  '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式热泵机组(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '离心式热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 38).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 38).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断离心式热泵机组(2)是否开启
        '离心式热泵机组（1）全部最大负荷运行后，使用离心式热泵机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 18).Value <> "Lx000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 18).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 18).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 18).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置离心式热泵机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 39).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率-离心式热泵机组（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 18).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 18).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 39).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 直燃型溴化锂常规模式供冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, JSBC As Integer, JDZLL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim LFHLXX1, LFHLXX2 As Double '设备负荷率下限
        Dim LFHL1 As Double = 0 '设备(1)负荷率计算结果
        Dim LFHL2 As Double = 0 '设备(2)负荷率计算结果
        'JDZLL：上一个阶段已经提供的制冷量
        '————————————————————————————————————————————————————————————————————————————————————————————
        '优先使用设备（1）
        '判断直燃型溴化锂机组(1)是否开启
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 22).Value <> "LDF-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value Then  '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX1 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX1 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX1 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX1 = 0.75 '负荷率下限1
                End If
            End If
            '设置直燃型溴化锂机组(1)负荷率，并判断是否超过最大冷负荷需求量
            For b1 = 0 To JSBC 'b1表示设备(1)
                LFHL1 = LFHLXX1 + (FHTJJD / 100) * b1 '离心式热泵机组(1)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 90).Value = LFHL1
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率）
                If (LFHZXQL(b) - JDZLL) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 90).Value = LFHL1
                    Exit For
                End If
            Next
        End If
        '判断直燃型溴化锂机组(2)是否开启
        '直燃型溴化锂机组（1）全部最大负荷运行后，使用直燃型溴化锂机组（2）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 22).Value <> "LDF-000" Then
            '二分法提高计算速度
            ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = 0.5 '负荷率为0.5
            If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 22).Value Then '负荷率等于0.5大了，或者恰好等于0.5
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = 0.25 '负荷率等于0.25
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 22).Value Then '负荷率等于0.25大了，或者恰好等于0.25
                    LFHLXX2 = 0    '负荷率下限1
                Else '如果负荷率等于0.25小了
                    LFHLXX2 = 0.25 '负荷率下限1
                End If
            Else '负荷率等于0.5小了
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = 0.75 '负荷率等于0.75
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 22).Value Then '负荷率等于0.75大了，或者恰好等于0.75
                    LFHLXX2 = 0.5  '负荷率下限1
                Else '如果负荷率等于0.75小了
                    LFHLXX2 = 0.75 '负荷率下限1
                End If
            End If
            '设置直燃型溴化锂机组(2)负荷率
            For b2 = 0 To JSBC 'b2表示设备(2)
                LFHL2 = LFHLXX2 + (FHTJJD / 100) * b2 '(2)负荷率
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 91).Value = LFHL2
                '用于判断的负荷是（冷负荷总需求量-上阶段供冷总功率-直燃型溴化锂机组（1）供冷量）
                If (LFHZXQL(b) - JDZLL - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 22).Value) <= ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 22).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 91).Value = LFHL2
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub 蓄能装置蓄冷工况常规模式计算(ExcelApp As Object, b As Integer, FHTJJD As Double, calculation_mode As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的计算模式类型，只有为类型1时，才进行计算
        If calculation_mode = 1 Then
            '蓄冷，蓄冷，蓄冷，蓄冷，蓄冷
            '蓄冷装置蓄冷状态计算
            '蓄冷时不同空调机的启动顺序和制冷时一样
            If XNXLGL(b) > 0 Then
                '将本工况蓄能量和蓄能时间带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 14).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 14).Value
                '蓄冷时间
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 24).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 24).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 25).Value
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 26).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 26).Value
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取制冷季装机方案及参数
                Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
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
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第一顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 6).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第二顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 7).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第三顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 8).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第四顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 9).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第五顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 10).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
                '读取第六顺序
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "离心式冷水机" Then
                    Call 离心式冷水机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSLSJ, NUM2_LXSLSJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "风冷螺杆机" Then
                    Call 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_FLLGJ, NUM2_FLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "空气源热泵" Then
                    Call 空气源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_KQYRB, NUM2_KQYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "水(地)源热泵" Then
                    Call 水_地源热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SDYRB, NUM2_SDYRB)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "水冷螺杆机" Then
                    Call 水冷螺杆机常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_SLLGJ, NUM2_SLLGJ)
                End If
                If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 11).Value = "离心式热泵" Then
                    Call 离心式热泵常规模式蓄冷计算(ExcelApp, b, FHTJJD, NUM1_LXSRB, NUM2_LXSRB)
                End If
                '————————————————————————————————————————————————————————————————————————————————————————————
                '————————————————————————————————————————————————————————————————————————————————————————————
            ElseIf XNXLGL(b) = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51)).Value = 0
            ElseIf XNXLGL(b) < 0 Then
                Dim XZ1
                XZ1 = MsgBox("蓄冷负荷总需求量不能小于0", vbOKCancel)
            End If
        End If
    End Sub

    Sub 离心式冷水机常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷已经启动的数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 10).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 10).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 40).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 40).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 41).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 41).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 41).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 风冷螺杆式热泵常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷启动数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 14).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 14).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 42).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 42).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 42).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 43).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 43).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 43).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 空气源热泵常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷启动数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 20).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 20).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 44).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 44).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 44).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 45).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 45).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 45).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 水_地源热泵常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷启动数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 16).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 16).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 46).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 46).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 46).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 47).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 47).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 47).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 水冷螺杆机常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷启动数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 12).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 12).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 48).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 48).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 48).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 49).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 49).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 49).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

    Sub 离心式热泵常规模式蓄冷计算(ExcelApp As Object, b As Integer, FHTJJD As Double, NUM1 As Double, NUM2 As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————————
        '蓄冷启动数量
        Dim XLQDSL1, XLQDSL2 As Double
        '制冷启动数量
        Dim ZLQDSL1 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 18).Value
        Dim ZLQDSL2 As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 18).Value
        '装机总数量
        Dim ZLZJSL1 As Double = NUM1
        Dim ZLZJSL2 As Double = NUM2
        If (ZLZJSL1 - ZLQDSL1 > 0) Then
            For e1 = 0 To (ZLZJSL1 - ZLQDSL1) / (FHTJJD / 100) '确定最大计算数量
                XLQDSL1 = (FHTJJD / 100) * e1 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 50).Value = XLQDSL1
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 50).Value = XLQDSL1
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 50).Value = XLQDSL1
                    Exit For '跳出循环
                End If
            Next
        End If
        If ((ZLZJSL2 - ZLQDSL2 > 0) And (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value = 0)) Then
            For e2 = 0 To (ZLZJSL2 - ZLQDSL2) / (FHTJJD / 100)
                XLQDSL2 = (FHTJJD / 100) * e2 '蓄冷启动数量
                '将蓄冷启动数量带入计算一次
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 51).Value = XLQDSL2
                ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51).Value = XLQDSL2
                '设置跳出循环的条件
                '当机组蓄冷负荷率小于等于1并且大于0时，跳出循环
                If (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value <= 1 And ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(39, 24).Value > 0) Then
                    '记录下此时的蓄冷设备启动数量
                    ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + b, 51).Value = XLQDSL2
                    Exit For '跳出循环
                End If
            Next
        End If
    End Sub

End Module
