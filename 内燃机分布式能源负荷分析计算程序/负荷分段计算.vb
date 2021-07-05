Module 负荷分段计算
    Sub 制冷设备负荷分段(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '清空已有的计算结果
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Range("D81:D89").ClearContents
        '添加部分报错功能
        '如果第一个工况的第一顺序制冷设备为空，报错
        Dim ZLZJ As Double = 0 '除了溴化锂和蓄冷以外的装机功率合计
        For i = 1 To 7
            ZLZJ = ZLZJ + ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(49 + i, 7).Value
        Next
        If ZLZJ > 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = Nothing Then
            MsgBox("第1个工况的第1到第6顺序制冷设备类型选择不能为空！")
            Exit Sub
        End If
        '如果第一个工况的内燃机余热利用方式选择为空，报错
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value > 0 Then '有内燃机装机量
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 77).Value = Nothing Then
                MsgBox("第1个工况的内燃机余热利用方式选择不能为空！")
                Exit Sub
            End If
        End If
        '如果蓄冷供冷功率>0，但是小时数=0，报错
        If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 13).Value <> 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 25).Value = 0) Then
            MsgBox("第1个工况的蓄冷供冷功率不为0时，工况小时数必需也不为0！")
            Exit Sub
        End If
        '定义局部变量
        '制冷设备负荷分段计算中用到的变量
        '第一到第六顺序制冷设备COP（初始值为0）
        Dim D1SXZLCOP As Double = 0
        Dim D2SXZLCOP As Double = 0
        Dim D3SXZLCOP As Double = 0
        Dim D4SXZLCOP As Double = 0
        Dim D5SXZLCOP As Double = 0
        Dim D6SXZLCOP As Double = 0
        '第一到第六顺序制冷设备制冷功率（初始值为0）
        Dim D1SXZLGL As Double = 0
        Dim D2SXZLGL As Double = 0
        Dim D3SXZLGL As Double = 0
        Dim D4SXZLGL As Double = 0
        Dim D5SXZLGL As Double = 0
        Dim D6SXZLGL As Double = 0
        Dim NRJGDGL, NRJGDGL30p '内燃机供电功率，最小功率单台内燃机供电功率30%，扣除内燃机和溴化锂的自用电
        Dim XHLZLa, XHLZLa30p '溴化锂制冷功率，最小功率单台溴化锂制冷功率30%
        Dim NRJ30pGDGLSY = 0 '单台内燃机30%负荷发电功率还剩余的量
        Dim NRJGDGLSY = 0 '全部内燃机100%负荷发电功率还剩余的量
        Dim D1SXZLGLSY = 0, D2SXZLGLSY = 0, D3SXZLGLSY = 0, D4SXZLGLSY = 0, D5SXZLGLSY = 0, D6SXZLGLSY = 0 '第一到第六顺序制冷设备制冷功率剩余量
        '解锁工作表，读取计算参数
        Call 解锁工作表(ExcelApp)
        '第一顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 6).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D1SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D1SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '第二顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 7).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D2SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D2SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '第三顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 8).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D3SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D3SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '第四顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 9).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D4SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D4SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '第五顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 10).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D5SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D5SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '第六顺序制冷电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "离心式冷水机" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(50, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(51, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "水冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(52, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(53, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(54, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(55, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 11).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D6SXZLCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 12).Value
            D6SXZLGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(56, 7).Value
        End If
        '将第一个工况的内燃机余热利用方式带入计算，内燃机负荷率设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 77).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3).Value = 1
        '将蓄冷供冷功率和小时数带入计算
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 13).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 13).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 25).Value
        Dim XLGLGLa = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 13).Value '记录蓄冷供冷功率
        Dim XLGLFJHDGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(8, 24).Value '记录蓄冷供冷辅机耗电功率
        '按照选择的制冷设备启动顺序，计算启动不同设备时的冷负荷是多少kW。
        '如果有内燃机装机量，采用以下计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value > 0 Then
            '读取各种计算系数
            Dim ans_XZXS_L = 读取制冷季输入的计算系数(ExcelApp)
            '溴化锂制冷COP
            Dim XHL_COP_L As Double = ans_XZXS_L(10)
            '全部内燃机100%负荷情况下，内燃机扣除自用电后向外供电功率以及溴化锂制冷功率
            NRJGDGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(170, 4).Value
            XHLZLa = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(172, 4).Value * XHL_COP_L
            '两种型号内燃机中，单机功率最小的一台
            Dim NRJDJFDGLmin, NRJDJYRGLmin As Double '内燃机单机发电功率min，内燃机单机余热功率min
            '读取制冷季装机方案及参数
            Dim ans_ZJFA_L = 读取制冷季装机方案参数(ExcelApp)
            '内燃机
            '内燃机1单机发电功率
            Dim NRJ1DJFDGL As Double = ans_ZJFA_L(2)
            '内燃机2单机发电功率
            Dim NRJ2DJFDGL As Double = ans_ZJFA_L(3)
            '内燃机1单机余热功率
            Dim NRJ1DJYRGL As Double = ans_ZJFA_L(4)
            '内燃机2单机余热功率
            Dim NRJ2DJYRGL As Double = ans_ZJFA_L(5)
            '内燃机1辅机耗电功率
            Dim FJHD1_ED_NRJ As Double = ans_ZJFA_L(8)
            '内燃机2辅机耗电功率
            Dim FJHD2_ED_NRJ As Double = ans_ZJFA_L(9)
            '内燃机1+溴化锂1辅机耗电
            Dim FJHD1 = FJHD1_ED_NRJ + ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 4).Value
            '内燃机2+溴化锂2辅机耗电
            Dim FJHD2 = FJHD2_ED_NRJ + ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 4).Value
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL = 0 Then
                NRJDJFDGLmin = NRJ1DJFDGL - FJHD1
                NRJDJYRGLmin = NRJ1DJYRGL
            End If
            If NRJ2DJFDGL <> 0 And NRJ1DJFDGL = 0 Then
                NRJDJFDGLmin = NRJ2DJFDGL - FJHD2
                NRJDJYRGLmin = NRJ2DJYRGL
            End If
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL <> 0 And NRJ1DJFDGL >= NRJ2DJFDGL Then
                NRJDJFDGLmin = NRJ2DJFDGL - FJHD2
                NRJDJYRGLmin = NRJ2DJYRGL
            End If
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL <> 0 And NRJ1DJFDGL <= NRJ2DJFDGL Then
                NRJDJFDGLmin = NRJ1DJFDGL - FJHD1
                NRJDJYRGLmin = NRJ1DJYRGL
            End If
            '功率最小的单台内燃机在30%负荷情况下，内燃机扣除自用电后向外供电功率以及溴化锂制冷功率
            NRJGDGL30p = 0.3 * NRJDJFDGLmin
            XHLZLa30p = 0.3 * NRJDJYRGLmin * XHL_COP_L
            '如果选择的内燃机余热利用方式不是溴化锂，将溴化锂制冷功率重置为0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value <> "溴化锂" Then
                XHLZLa = 0
                XHLZLa30p = 0
            End If
            '按照内燃机可不可以向外供电，采用两种不同的计算方式
            '内燃机不可以向外供电时，采用以下计算方式：
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "N" Then
                '判断此时的单台内燃机30%负荷供电全部用于蓄能装置供冷，是否能完全消纳掉
                '单台内燃机30%负荷供电量如果不能被蓄能装置供冷完全消纳
                If NRJGDGL30p > XLGLFJHDGL Then
                    '内燃机供电功率剩余1（蓄冷设备全部启动后，内燃机发电功率剩余量）
                    NRJ30pGDGLSY = NRJGDGL30p - XLGLFJHDGL
                    '制冷设备负荷第一段=蓄冷设备100%制冷量，内燃机达不到单台30%负荷，不能开启，无溴化锂制冷
                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(81, 4).Value = Math.Round(XLGLGLa, 2)
                    '判断此时的单台内燃机30%负荷供电全部用于蓄冷设备和第一顺序制冷，是否能完全消纳掉
                    '单台内燃机30%负荷供电量如果不能完全消纳
                    If NRJGDGL30p > XLGLFJHDGL + D1SXZLGL / D1SXZLCOP Then
                        '蓄冷设备和第一顺序空调设备全部启动，计算此时的耗电量
                        '内燃机供电功率剩余1（蓄冷设备和第一顺序空调设备全部启动后，内燃机发电功率剩余量），给第二顺序设备使用
                        NRJ30pGDGLSY = NRJGDGL30p - XLGLFJHDGL - D1SXZLGL / D1SXZLCOP
                        '制冷设备负荷第二段=蓄冷100%供冷功率+第一顺序空调设备100%制冷量，内燃机达不到单台30%负荷，不能开启，无溴化锂制冷
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + D1SXZLGL, 2)
                        '第一顺序制冷设备制冷功率剩余量=0
                        D1SXZLGLSY = 0
                        '为了计算快捷，假定单台内燃机30%负荷供电量的剩余量能够被第二顺序空调设备完全消纳
                        '制冷设备负荷第三段=单台溴化锂30%制冷量+蓄冷100%供冷功率+第一顺序空调设备100%制冷量+剩余的发电量*第二顺序COP
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XHLZLa30p + XLGLGLa + D1SXZLGL + NRJ30pGDGLSY * D2SXZLCOP, 2)
                        D2SXZLGLSY = D2SXZLGL - NRJ30pGDGLSY * D2SXZLCOP
                        NRJ30pGDGLSY = 0
                        '判断蓄冷设备100%+第一顺序制冷设备100%+第二顺序制冷设备100%运行时的耗电量，与内燃机100%运行的供电量的大小
                        '如果蓄冷设备100%+第一顺序制冷设备100%+第二顺序制冷设备100%运行时能够完全消纳内燃机供电量
                        If NRJGDGL <= XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） Then
                            '制冷设备负荷第四段=全部溴化锂100%供冷量+蓄冷100%供冷功率+第一顺序空调设备100%供冷+第二顺序空调设备100%可以用掉内燃机发电量时的制冷量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + (NRJGDGL - XLGLFJHDGL - D1SXZLGL / D1SXZLCOP) * D2SXZLCOP, 2)
                            '制冷设备负荷第五段=全部溴化锂100%供冷量+蓄冷100%供冷功率+第一和第二顺序空调设备100%制冷量，后面的以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                        End If
                        '如果蓄冷设备100%+第一顺序制冷设备100%+第二顺序制冷设备100%运行时不能完全消纳内燃机供电量
                        If NRJGDGL > XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） Then
                            '制冷设备负荷第四段=蓄冷100%供冷功率+第一和第二顺序空调设备100%制冷量+第一和第二顺序设备恰好100%负荷运行时的溴化锂制冷量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + (（XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP）） / (NRJGDGL)) * XHLZLa, 2)
                            NRJGDGLSY = NRJGDGL - XLGLFJHDGL - （（D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP）） '内燃机供电功率剩余量
                            D1SXZLGLSY = 0
                            D2SXZLGLSY = 0
                            '判断内燃机100%供电量能否被第一和第二和第三顺序设备全部消纳
                            '如果蓄冷设备100%+第一顺序制冷设备100%+第二顺序制冷设备100%+第三顺序制冷设备100%运行时能完全消纳内燃机供电量
                            If NRJGDGL <= XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP） Then
                                '制冷设备负荷第五段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一和第二顺序空调设备100%制冷量+内燃机发电量恰好100%时第三顺序空调制冷量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + NRJGDGLSY * D3SXZLCOP, 2)
                                '制冷设备负荷第六段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一、第二和第三顺序空调设备100%制冷量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                            End If
                            '如果蓄冷设备100%+第一顺序制冷设备100%+第二顺序制冷设备100%+第三顺序制冷设备100%运行时不能完全消纳内燃机供电量
                            If NRJGDGL > XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP） Then
                                '制冷设备负荷第五段=蓄冷100%供冷功率+第一和第二和第三顺序空调设备100%制冷量+第一和第二和第三顺序设备恰好100%负荷运行时的溴化锂制冷量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + (（XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP）） / (NRJGDGL)) * XHLZLa, 2)
                                NRJGDGLSY = NRJGDGL - XLGLFJHDGL - （（D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP）） '剩余的内燃机供电功率
                                '为了计算简便，假定内燃机100%全部供电量都能够被第四顺序制冷设备消纳掉
                                '制冷设备负荷第六段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一和第二和第三顺序空调设备100%制冷量+内燃机发电量恰好100%时第四顺序空调制冷量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + NRJGDGLSY * D4SXZLCOP, 2)
                                '制冷设备负荷第七段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一、第二、第三和第四顺序空调设备100%制冷量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XHLZLa + XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                            End If
                        End If
                    End If
                    '单台内燃机30%负荷供电量如果能够被蓄冷装置和第一顺序空调设备完全消纳
                    If NRJGDGL30p <= XLGLFJHDGL + D1SXZLGL / D1SXZLCOP Then
                        '制冷设备负荷第二段=蓄冷100%供冷功率+单台溴化锂30%制冷量+单台内燃机30%发电量*第一顺序制冷设备COP
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + XHLZLa30p + NRJGDGL30p * D1SXZLCOP, 2)
                        '第一顺序制冷设备制冷功率剩余量
                        D1SXZLGLSY = D1SXZLGL - NRJGDGL30p * D1SXZLCOP
                        '单台内燃机30%负荷供电功率剩余1=0
                        NRJ30pGDGLSY = 0
                        '判断第一顺序全部制冷功率制冷耗电量，与全部内燃机全部100%负荷供电量的大小
                        '全部内燃机100%负荷供电功率能够被第一顺序制冷设备和蓄冷设备100%供冷全部消纳
                        If NRJGDGL <= XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) Then
                            '制冷设备负荷第三段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一顺序空调设备100%可以用掉内燃机发电量时的制冷量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + XHLZLa + NRJGDGL * D1SXZLCOP, 2)
                            '制冷设备负荷第四段=蓄冷100%供冷功率+全部溴化锂100%供冷量+第一顺序空调设备100%制冷量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL, 2)
                            D1SXZLGLSY = 0
                            NRJGDGLSY = 0
                            '制冷设备负荷第五段=全部溴化锂100%供冷量+第一顺序空调设备100%制冷量+第二顺序空调设备100%制冷量，以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                        End If
                        '全部内燃机100%负荷供电功率不能被第一顺序制冷设备全部消纳
                        If NRJGDGL > XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) Then
                            '制冷设备负荷第三段=蓄冷100%供冷功率+第一顺序空调设备100%制冷量+第一顺序设备恰好100%负荷运行时的溴化锂制冷量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + (XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) / NRJGDGL) * XHLZLa, 2)
                            NRJGDGLSY = NRJGDGL - (D1SXZLGL / D1SXZLCOP)
                            D1SXZLGLSY = 0
                            '全部内燃机100%负荷供电功率不能被蓄冷100%供冷功率+第一和第二顺序制冷设备全部消纳
                            If NRJGDGL > XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） Then
                                '制冷设备负荷第四段=蓄冷100%供冷功率+第一和第二顺序空调设备100%制冷量+第一和第二顺序设备恰好100%负荷运行时的溴化锂制冷量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + (（XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP）） / (NRJGDGL)) * XHLZLa, 2)
                                NRJGDGLSY = NRJGDGL - XLGLFJHDGL - （（D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP））
                                D1SXZLGLSY = 0
                                D2SXZLGLSY = 0
                                '判断内燃机100%供电量能否被蓄冷100%供冷功率+第一和第二和第三顺序设备全部消纳
                                '如果内燃机100%供电量能够被蓄冷100%供冷功率+第一和第二和第三顺序设备全部消纳
                                If NRJGDGL <= XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP） Then
                                    '制冷设备负荷第五段=全部溴化锂100%供冷量+第一和第二顺序空调设备100%制冷量+内燃机发电量恰好100%时第三顺序空调制冷量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + NRJGDGLSY * D3SXZLCOP, 2)
                                    '制冷设备负荷第六段=全部溴化锂100%供冷量+第一、第二和第三顺序空调设备100%制冷量，后面的以此类推
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                                End If
                                '如果内燃机100%供电量不能够被第一和第二和第三顺序设备全部消纳
                                If NRJGDGL > XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP） Then
                                    '制冷设备负荷第五段=第一和第二和第三顺序空调设备100%制冷量+第一和第二和第三顺序设备恰好100%负荷运行时的溴化锂制冷量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + (（XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP）） / (NRJGDGL)) * XHLZLa, 2)
                                    NRJGDGLSY = NRJGDGL - XLGLFJHDGL - （（D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） + （D3SXZLGL / D3SXZLCOP）） '剩余的内燃机供电功率
                                    '为了计算简便，假定内燃机100%全部供电量都能够被第四顺序制冷设备消纳掉
                                    '制冷设备负荷第六段=全部溴化锂100%供冷量+第一和第二和第三顺序空调设备100%制冷量+内燃机发电量恰好100%时第四顺序空调制冷量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + NRJGDGLSY * D4SXZLCOP, 2)
                                    '制冷设备负荷第七段=全部溴化锂100%供冷量+第一、第二、第三和第四顺序空调设备100%制冷量，后面的以此类推
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                                End If
                            End If
                            '全部内燃机100%负荷供电功率能被第一和第二顺序制冷设备全部消纳
                            If NRJGDGL <= XLGLFJHDGL + （D1SXZLGL / D1SXZLCOP） + （D2SXZLGL / D2SXZLCOP） Then
                                '制冷设备负荷第四段=全部溴化锂100%供冷量+第一顺序空调设备100%制冷量+内燃机发电量恰好100%时第二顺序空调制冷量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + NRJGDGLSY * D2SXZLCOP, 2)
                                '制冷设备负荷第五段=全部溴化锂100%供冷量+第一和第二顺序空调设备100%制冷量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                            End If
                        End If
                    End If
                End If
                '单台内燃机30%负荷供电量如果能被蓄能装置供冷完全消纳
                If NRJGDGL30p <= XLGLFJHDGL Then
                    '制冷设备负荷第一段=蓄冷设备100%制冷量，内燃机恰好为单台30%负荷时溴化锂制冷功率
                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(81, 4).Value = Math.Round(XLGLGLa + XHLZLa30p, 2)
                    '判断内燃机100%发电功率能否被蓄冷和第一顺序制冷设备消纳
                    '如果内燃机100%发电功率能被蓄冷和第一顺序制冷设备消纳
                    If NRJGDGL <= XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) Then
                        '制冷设备负荷第二段=蓄冷设备100%制冷量+溴化锂100%制冷功率+内燃机恰好为100%负荷时第一顺序制冷功率
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + XHLZLa + NRJGDGL * D1SXZLCOP, 2)
                        D1SXZLGLSY = D1SXZLGL - NRJGDGL * D1SXZLCOP '第一顺序制冷功率剩余
                        NRJGDGLSY = 0 '内燃机功率剩余
                        '制冷设备负荷第三段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷功率，后面的以此类推
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                    End If
                    '如果内燃机100%发电功率不能被蓄冷和第一顺序制冷设备消纳
                    If NRJGDGL > XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) Then
                        '制冷设备负荷第二段=蓄冷设备100%制冷量+第一顺序制冷功率+此时溴化锂供冷功率
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + XHLZLa * ((D1SXZLGL / D1SXZLCOP) + XLGLFJHDGL) / NRJGDGL, 2)
                        NRJGDGLSY = NRJGDGL - XLGLFJHDGL - (D1SXZLGL / D1SXZLCOP) '内燃机供电功率剩余量
                        '如果内燃机100%发电功率能够被蓄冷+第一顺序+第二顺序设备消纳
                        If NRJGDGL <= XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) Then
                            '制冷设备负荷第三段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷+内燃机恰好为100%负荷时第二顺序制冷功率
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + NRJGDGLSY * D2SXZLCOP, 2)
                            '制冷设备负荷第四段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷功率+第二顺序，后面的以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                        End If
                        '如果内燃机100%发电功率不能够被蓄冷+第一顺序+第二顺序设备消纳
                        If NRJGDGL > XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) Then
                            '制冷设备负荷第三段=蓄冷设备100%制冷量+第一顺序制冷+第二顺序制冷+此时溴化锂供冷功率
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + XHLZLa * ((D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) + XLGLFJHDGL) / NRJGDGL, 2)
                            NRJGDGLSY = NRJGDGL - XLGLFJHDGL - (D1SXZLGL / D1SXZLCOP) - (D2SXZLGL / D2SXZLCOP) '内燃机供电功率剩余量
                            '如果内燃机100%发电功率能够被蓄冷+第一顺序+第二顺序+第三顺序设备消纳
                            If NRJGDGL <= XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) + (D3SXZLGL / D3SXZLCOP) Then
                                '制冷设备负荷第四段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷+第二顺序制冷+内燃机恰好为100%负荷时第三顺序制冷功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + NRJGDGLSY * D3SXZLCOP, 2)
                                '制冷设备负荷第五段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷功率+第二顺序+第三顺序，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                            End If
                            '如果内燃机100%发电功率不能够被蓄冷+第一顺序+第二顺序+第三顺序设备消纳
                            If NRJGDGL > XLGLFJHDGL + (D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) + (D3SXZLGL / D3SXZLCOP) Then
                                '制冷设备负荷第四段=蓄冷设备100%制冷量+第一顺序制冷+第二顺序制冷+第三顺序+此时溴化锂供冷功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + XHLZLa * ((D1SXZLGL / D1SXZLCOP) + (D2SXZLGL / D2SXZLCOP) + (D3SXZLGL / D3SXZLCOP) + XLGLFJHDGL) / NRJGDGL, 2)
                                NRJGDGLSY = NRJGDGL - XLGLFJHDGL - (D1SXZLGL / D1SXZLCOP) - (D2SXZLGL / D2SXZLCOP) '内燃机供电功率剩余量
                                '为了计算方便，假定内燃机100%供电量能够被第四顺序设备设备消纳
                                '制冷设备负荷第五段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷+第二顺序制冷+第三顺序制冷+内燃机恰好为100%负荷时第四顺序制冷功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + NRJGDGLSY * D4SXZLCOP, 2)
                                '制冷设备负荷第六段=蓄冷设备100%制冷量+溴化锂100%制冷功率+第一顺序制冷功率+第二顺序+第三顺序+第四顺序，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
                            End If
                        End If
                    End If
                End If
            End If
            '如果内燃机可以向外供电
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "Y" Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(81, 4).Value = Math.Round(XLGLGLa, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + XHLZLa30p, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + XHLZLa, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(88, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(89, 4).Value = Math.Round(XLGLGLa + XHLZLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
            End If
        End If
        '如果没有内燃机装机量，采用以下计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(81, 4).Value = Math.Round(XLGLGLa, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(82, 4).Value = Math.Round(XLGLGLa + D1SXZLGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(83, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(84, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(85, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(86, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(87, 4).Value = Math.Round(XLGLGLa + D1SXZLGL + D2SXZLGL + D3SXZLGL + D4SXZLGL + D5SXZLGL + D6SXZLGL, 2)
        End If
        '将重复的值删除
        For i = 8 To 1 Step -1 '从下向上查找，把紧挨着的重复的值去掉
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value = Nothing
            End If
        Next
        For i = 1 To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value <= ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value = Nothing
            End If
        Next
        '如果有为0的结果，将后面的值前移
        Dim KSXH As Integer = 1 '开始检索的序号，默认值为1
        Dim js As Integer = 0 '计数
        For i = 2 To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value <> Nothing Then
                js = js + 1
            End If
        Next
        If js > 1 Then '如果有超过两段，从第一个开始检索，否则第二个
            KSXH = 1
        Else
            KSXH = 2
        End If
        For i = KSXH To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value < ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(79 + i, 4).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(80 + i, 4).Value = Nothing
            End If
        Next
        '清理部分数据
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 3).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 13).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = Nothing
    End Sub
    Sub 制热设备负荷分段(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '清空已有的计算结果
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Range("D92:D100").ClearContents
        '添加部分报错功能
        '如果第一个工况的第一顺序制热设备为空，报错
        Dim ZRZJ As Double = 0 '除了溴化锂和蓄热以外的装机功率合计
        For i = 1 To 6
            ZRZJ = ZRZJ + ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(60 + i, 7).Value
        Next
        If ZRZJ > 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = Nothing Then
            MsgBox("第1个工况的第1到第6顺序制热设备类型选择不能为空！")
            Exit Sub
        End If
        '如果第一个工况的内燃机余热利用方式选择为空，报错
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value > 0 Then '有内燃机装机量
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 77).Value = Nothing Then
                MsgBox("第1个工况的内燃机余热利用方式选择不能为空！")
                Exit Sub
            End If
        End If
        '蓄热供热功率和工况小时数必需同时为0或者同时大于0
        If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 22).Value <> 0 And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 25).Value = 0) Then
            MsgBox("第1个工况的蓄热供热功率不为0时，工况小时数必需也不为0！")
            Exit Sub
        End If
        '定义局部变量
        '制热设备负荷分段计算中用到的变量
        '第一到第六顺序制热设备COP（初始值=0）
        Dim D1SXZRCOP As Double = 0
        Dim D2SXZRCOP As Double = 0
        Dim D3SXZRCOP As Double = 0
        Dim D4SXZRCOP As Double = 0
        Dim D5SXZRCOP As Double = 0
        Dim D6SXZRCOP As Double = 0
        '第一到第六顺序制热设备制热功率（初始值=0）
        Dim D1SXZRGL As Double = 0
        Dim D2SXZRGL As Double = 0
        Dim D3SXZRGL As Double = 0
        Dim D4SXZRGL As Double = 0
        Dim D5SXZRGL As Double = 0
        Dim D6SXZRGL As Double = 0
        Dim NRJGDGL, NRJGDGL30p '内燃机供电功率，最小功率单台内燃机供电功率30%，扣除内燃机和溴化锂的自用电
        Dim XHLZRa, XHLZRa30p '溴化锂制热功率，最小功率单台溴化锂制热功率30%
        Dim NRJ30pGDGLSY = 0 '单台内燃机30%负荷发电功率还剩余的量
        Dim NRJGDGLSY = 0 '全部内燃机1000%负荷发电功率还剩余的量
        Dim D1SXZRGLSY = 0, D2SXZRGLSY = 0, D3SXZRGLSY = 0, D4SXZRGLSY = 0, D5SXZRGLSY = 0, D6SXZRGLSY = 0 '第一到第六顺序制热设备制热功率剩余量
        '解锁工作表，读取计算参数
        Call 解锁工作表(ExcelApp)
        '第一顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 15).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D1SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D1SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '第二顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 16).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D2SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D2SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '第三顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 17).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D3SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D3SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '第四顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 18).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D4SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D4SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '第五顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 19).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D5SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D5SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '第六顺序制热电空调设备COP（电）
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "空气源热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(61, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "水(地)源热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(62, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "风冷螺杆机" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(63, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "离心式热泵" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(64, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "天然气锅炉" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(65, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "直燃型溴化锂" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(66, 7).Value
        End If
        If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 20).Value = "电采暖锅炉" Then '全部按照第1工况的顺序进行计算
            D6SXZRCOP = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 12).Value
            D6SXZRGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(67, 7).Value
        End If
        '将第一个工况的内燃机余热利用方式带入计算，内燃机负荷率设置为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 77).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = 1
        '将蓄热供热功率和小时数带入计算
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 22).Value
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 25).Value
        Dim XRGRGLa = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 22).Value '记录蓄热供热功率
        Dim XRGRFJHDGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(53, 24).Value '记录蓄热供热辅机耗电功率
        '按照选择的制热设备启动顺序，计算启动不同设备时的热负荷是多少kW。
        '如果有内燃机装机量，采用以下计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value > 0 Then
            '读取各种修正系数
            Dim ans_XZXS_R = 读取采暖季输入的计算系数(ExcelApp)
            '溴化锂制热COP
            Dim XHL_COP_R As Double = ans_XZXS_R(9)
            '全部内燃机100%负荷情况下，内燃机扣除自用电后向外供电功率以及溴化锂制热功率
            NRJGDGL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value - ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(171, 4).Value
            XHLZRa = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(173, 4).Value * XHL_COP_R
            '功率最小的单台内燃机在30%负荷情况下，内燃机扣除自用电后向外供电功率以及溴化锂制热功率
            '两种型号内燃机中，单机功率最小的一台
            '内燃机单机发电功率min
            Dim NRJDJFDGLmin As Double = 0
            '内燃机单机余热功率min
            Dim NRJDJYRGLmin As Double = 0
            '读取采暖季装机方案及参数
            Dim ans_ZJFA_R = 读取采暖季装机方案参数(ExcelApp)
            '内燃机
            '内燃机1单机发电功率
            Dim NRJ1DJFDGL As Double = ans_ZJFA_R(2)
            '内燃机1单机发电功率
            Dim NRJ2DJFDGL As Double = ans_ZJFA_R(3)
            '内燃机1单机余热功率
            Dim NRJ1DJYRGL As Double = ans_ZJFA_R(4)
            '内燃机2单机余热功率
            Dim NRJ2DJYRGL As Double = ans_ZJFA_R(5)
            '内燃机1辅机耗电功率
            Dim FJHD1_ED_NRJ As Double = ans_ZJFA_R(8)
            '内燃机2辅机耗电功率
            Dim FJHD2_ED_NRJ As Double = ans_ZJFA_R(9)
            '内燃机1+溴化锂1辅机耗电
            Dim FJHD1 = FJHD1_ED_NRJ + ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 4).Value
            '内燃机2+溴化锂2辅机耗电
            Dim FJHD2 = FJHD2_ED_NRJ + ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 4).Value
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL = 0 Then
                NRJDJFDGLmin = NRJ1DJFDGL - FJHD1
                NRJDJYRGLmin = NRJ1DJYRGL
            End If
            If NRJ2DJFDGL <> 0 And NRJ1DJFDGL = 0 Then
                NRJDJFDGLmin = NRJ2DJFDGL - FJHD2
                NRJDJYRGLmin = NRJ2DJYRGL
            End If
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL <> 0 And NRJ1DJFDGL >= NRJ2DJFDGL Then
                NRJDJFDGLmin = NRJ2DJFDGL - FJHD2
                NRJDJYRGLmin = NRJ2DJYRGL
            End If
            If NRJ1DJFDGL <> 0 And NRJ2DJFDGL <> 0 And NRJ1DJFDGL <= NRJ2DJFDGL Then
                NRJDJFDGLmin = NRJ1DJFDGL - FJHD1
                NRJDJYRGLmin = NRJ1DJYRGL
            End If
            '功率最小的单台内燃机在30%负荷情况下，内燃机扣除自用电后向外供电功率以及溴化锂制冷功率
            NRJGDGL30p = 0.3 * NRJDJFDGLmin
            XHLZRa30p = 0.3 * NRJDJYRGLmin * XHL_COP_R
            '如果选择的内燃机余热利用方式不是溴化锂，将溴化锂制热功率重置为0
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value <> "溴化锂" Then
                XHLZRa = 0
                XHLZRa30p = 0
            End If
            '按照内燃机可不可以向外供电，采用两种不同的计算方式
            '内燃机不可以向外供电时，采用以下计算方式：
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "N" Then
                '判断此时的单台内燃机30%负荷供电全部用于蓄能装置供热，是否能完全消纳掉
                '单台内燃机30%负荷供电量如果不能被蓄能装置供热完全消纳
                If NRJGDGL30p > XRGRFJHDGL Then
                    '内燃机供电功率剩余1（蓄热设备全部启动后，内燃机发电功率剩余量）
                    NRJ30pGDGLSY = NRJGDGL30p - XRGRFJHDGL
                    '制热设备负荷第一段=蓄热设备100%制热量，内燃机达不到单台30%负荷，不能开启，无溴化锂制热
                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(92, 4).Value = Math.Round(XRGRGLa, 2)
                    '判断此时的单台内燃机30%负荷供电全部用于蓄热设备和第一顺序制热，是否能完全消纳掉
                    '单台内燃机30%负荷供电量如果不能完全消纳
                    If NRJGDGL30p > XRGRFJHDGL + D1SXZRGL / D1SXZRCOP Then
                        '蓄热设备和第一顺序空调设备全部启动，计算此时的耗电量
                        '内燃机供电功率剩余1（蓄热设备和第一顺序空调设备全部启动后，内燃机发电功率剩余量），给第二顺序设备使用
                        NRJ30pGDGLSY = NRJGDGL30p - XRGRFJHDGL - D1SXZRGL / D1SXZRCOP
                        '制热设备负荷第二段=蓄热100%供热功率+第一顺序空调设备100%制热量，内燃机达不到单台30%负荷，不能开启，无溴化锂制热
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + D1SXZRGL, 2)
                        '第一顺序制热设备制热功率剩余量=0
                        D1SXZRGLSY = 0
                        '为了计算快捷，假定单台内燃机30%负荷供电量的剩余量能够被第二顺序空调设备完全消纳
                        '制热设备负荷第三段=单台溴化锂30%制热量+蓄热100%供热功率+第一顺序空调设备100%制热量+剩余的发电量*第二顺序COP
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XHLZRa30p + XRGRGLa + D1SXZRGL + NRJ30pGDGLSY * D2SXZRCOP, 2)
                        D2SXZRGLSY = D2SXZRGL - NRJ30pGDGLSY * D2SXZRCOP
                        NRJ30pGDGLSY = 0
                        '判断蓄热设备100%+第一顺序制热设备100%+第二顺序制热设备100%运行时的耗电量，与内燃机100%运行的供电量的大小
                        '如果蓄热设备100%+第一顺序制热设备100%+第二顺序制热设备100%运行时能够完全消纳内燃机供电量
                        If NRJGDGL <= XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） Then
                            '制热设备负荷第四段=全部溴化锂100%供热量+蓄热100%供热功率+第一顺序空调设备100%供热+第二顺序空调设备100%可以用掉内燃机发电量时的制热量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + (NRJGDGL - XRGRFJHDGL - D1SXZRGL / D1SXZRCOP) * D2SXZRCOP, 2)
                            '制热设备负荷第五段=全部溴化锂100%供热量+蓄热100%供热功率+第一和第二顺序空调设备100%制热量，后面的以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                        End If
                        '如果蓄热设备100%+第一顺序制热设备100%+第二顺序制热设备100%运行时不能完全消纳内燃机供电量
                        If NRJGDGL > XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） Then
                            '制热设备负荷第四段=蓄热100%供热功率+第一和第二顺序空调设备100%制热量+第一和第二顺序设备恰好100%负荷运行时的溴化锂制热量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + (（XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP）） / (NRJGDGL)) * XHLZRa, 2)
                            NRJGDGLSY = NRJGDGL - XRGRFJHDGL - （（D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP）） '内燃机供电功率剩余量
                            D1SXZRGLSY = 0
                            D2SXZRGLSY = 0
                            '判断内燃机100%供电量能否被第一和第二和第三顺序设备全部消纳
                            '如果蓄热设备100%+第一顺序制热设备100%+第二顺序制热设备100%+第三顺序制热设备100%运行时能完全消纳内燃机供电量
                            If NRJGDGL <= XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP） Then
                                '制热设备负荷第五段=蓄热100%供热功率+全部溴化锂100%供热量+第一和第二顺序空调设备100%制热量+内燃机发电量恰好100%时第三顺序空调制热量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + NRJGDGLSY * D3SXZRCOP, 2)
                                '制热设备负荷第六段=蓄热100%供热功率+全部溴化锂100%供热量+第一、第二和第三顺序空调设备100%制热量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                            End If
                            '如果蓄热设备100%+第一顺序制热设备100%+第二顺序制热设备100%+第三顺序制热设备100%运行时不能完全消纳内燃机供电量
                            If NRJGDGL > XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP） Then
                                '制热设备负荷第五段=蓄热100%供热功率+第一和第二和第三顺序空调设备100%制热量+第一和第二和第三顺序设备恰好100%负荷运行时的溴化锂制热量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + (（XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP）） / (NRJGDGL)) * XHLZRa, 2)
                                NRJGDGLSY = NRJGDGL - XRGRFJHDGL - （（D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP）） '剩余的内燃机供电功率
                                '为了计算简便，假定内燃机100%全部供电量都能够被第四顺序制热设备消纳掉
                                '制热设备负荷第六段=蓄热100%供热功率+全部溴化锂100%供热量+第一和第二和第三顺序空调设备100%制热量+内燃机发电量恰好100%时第四顺序空调制热量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + NRJGDGLSY * D4SXZRCOP, 2)
                                '制热设备负荷第七段=蓄热100%供热功率+全部溴化锂100%供热量+第一、第二、第三和第四顺序空调设备100%制热量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XHLZRa + XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                            End If
                        End If
                    End If
                    '单台内燃机30%负荷供电量如果能够被蓄热装置和第一顺序空调设备完全消纳
                    If NRJGDGL30p <= XRGRFJHDGL + D1SXZRGL / D1SXZRCOP Then
                        '制热设备负荷第二段=蓄热100%供热功率+单台溴化锂30%制热量+单台内燃机30%发电量*第一顺序制热设备COP
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + XHLZRa30p + NRJGDGL30p * D1SXZRCOP, 2)
                        '第一顺序制热设备制热功率剩余量
                        D1SXZRGLSY = D1SXZRGL - NRJGDGL30p * D1SXZRCOP
                        '单台内燃机30%负荷供电功率剩余1=0
                        NRJ30pGDGLSY = 0
                        '判断第一顺序全部制热功率制热耗电量，与全部内燃机全部100%负荷供电量的大小
                        '全部内燃机100%负荷供电功率能够被第一顺序制热设备和蓄热设备100%供热全部消纳
                        If NRJGDGL <= XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) Then
                            '制热设备负荷第三段=蓄热100%供热功率+全部溴化锂100%供热量+第一顺序空调设备100%可以用掉内燃机发电量时的制热量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + XHLZRa + NRJGDGL * D1SXZRCOP, 2)
                            '制热设备负荷第四段=蓄热100%供热功率+全部溴化锂100%供热量+第一顺序空调设备100%制热量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL, 2)
                            D1SXZRGLSY = 0
                            NRJGDGLSY = 0
                            '制热设备负荷第五段=全部溴化锂100%供热量+第一顺序空调设备100%制热量+第二顺序空调设备100%制热量，以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                        End If
                        '全部内燃机100%负荷供电功率不能被第一顺序制热设备全部消纳
                        If NRJGDGL > XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) Then
                            '制热设备负荷第三段=蓄热100%供热功率+第一顺序空调设备100%制热量+第一顺序设备恰好100%负荷运行时的溴化锂制热量
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + (XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) / NRJGDGL) * XHLZRa, 2)
                            NRJGDGLSY = NRJGDGL - (D1SXZRGL / D1SXZRCOP)
                            D1SXZRGLSY = 0
                            '全部内燃机100%负荷供电功率不能被蓄热100%供热功率+第一和第二顺序制热设备全部消纳
                            If NRJGDGL > XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） Then
                                '制热设备负荷第四段=蓄热100%供热功率+第一和第二顺序空调设备100%制热量+第一和第二顺序设备恰好100%负荷运行时的溴化锂制热量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + (（XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP）） / (NRJGDGL)) * XHLZRa, 2)
                                NRJGDGLSY = NRJGDGL - XRGRFJHDGL - （（D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP））
                                D1SXZRGLSY = 0
                                D2SXZRGLSY = 0
                                '判断内燃机100%供电量能否被蓄热100%供热功率+第一和第二和第三顺序设备全部消纳
                                '如果内燃机100%供电量能够被蓄热100%供热功率+第一和第二和第三顺序设备全部消纳
                                If NRJGDGL <= XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP） Then
                                    '制热设备负荷第五段=全部溴化锂100%供热量+第一和第二顺序空调设备100%制热量+内燃机发电量恰好100%时第三顺序空调制热量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + NRJGDGLSY * D3SXZRCOP, 2)
                                    '制热设备负荷第六段=全部溴化锂100%供热量+第一、第二和第三顺序空调设备100%制热量，后面的以此类推
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                                End If
                                '如果内燃机100%供电量不能够被第一和第二和第三顺序设备全部消纳
                                If NRJGDGL > XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP） Then
                                    '制热设备负荷第五段=第一和第二和第三顺序空调设备100%制热量+第一和第二和第三顺序设备恰好100%负荷运行时的溴化锂制热量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + (（XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP）） / (NRJGDGL)) * XHLZRa, 2)
                                    NRJGDGLSY = NRJGDGL - XRGRFJHDGL - （（D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） + （D3SXZRGL / D3SXZRCOP）） '剩余的内燃机供电功率
                                    '为了计算简便，假定内燃机100%全部供电量都能够被第四顺序制热设备消纳掉
                                    '制热设备负荷第六段=全部溴化锂100%供热量+第一和第二和第三顺序空调设备100%制热量+内燃机发电量恰好100%时第四顺序空调制热量
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + NRJGDGLSY * D4SXZRCOP, 2)
                                    '制热设备负荷第七段=全部溴化锂100%供热量+第一、第二、第三和第四顺序空调设备100%制热量，后面的以此类推
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                                End If
                            End If
                            '全部内燃机100%负荷供电功率能被第一和第二顺序制热设备全部消纳
                            If NRJGDGL <= XRGRFJHDGL + （D1SXZRGL / D1SXZRCOP） + （D2SXZRGL / D2SXZRCOP） Then
                                '制热设备负荷第四段=全部溴化锂100%供热量+第一顺序空调设备100%制热量+内燃机发电量恰好100%时第二顺序空调制热量
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + NRJGDGLSY * D2SXZRCOP, 2)
                                '制热设备负荷第五段=全部溴化锂100%供热量+第一和第二顺序空调设备100%制热量，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                            End If
                        End If
                    End If
                End If
                '单台内燃机30%负荷供电量如果能被蓄能装置供热完全消纳
                If NRJGDGL30p <= XRGRFJHDGL Then
                    '制热设备负荷第一段=蓄热设备100%制热量，内燃机恰好为单台30%负荷时溴化锂制热功率
                    ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(92, 4).Value = Math.Round(XRGRGLa + XHLZRa30p, 2)
                    '判断内燃机100%发电功率能否被蓄热和第一顺序制热设备消纳
                    '如果内燃机100%发电功率能被蓄热和第一顺序制热设备消纳
                    If NRJGDGL <= XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) Then
                        '制热设备负荷第二段=蓄热设备100%制热量+溴化锂100%制热功率+内燃机恰好为100%负荷时第一顺序制热功率
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + XHLZRa + NRJGDGL * D1SXZRCOP, 2)
                        D1SXZRGLSY = D1SXZRGL - NRJGDGL * D1SXZRCOP '第一顺序制热功率剩余
                        NRJGDGLSY = 0 '内燃机功率剩余
                        '制热设备负荷第三段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热功率，后面的以此类推
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                    End If
                    '如果内燃机100%发电功率不能被蓄热和第一顺序制热设备消纳
                    If NRJGDGL > XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) Then
                        '制热设备负荷第二段=蓄热设备100%制热量+第一顺序制热功率+此时溴化锂供热功率
                        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + XHLZRa * ((D1SXZRGL / D1SXZRCOP) + XRGRFJHDGL) / NRJGDGL, 2)
                        NRJGDGLSY = NRJGDGL - XRGRFJHDGL - (D1SXZRGL / D1SXZRCOP) '内燃机供电功率剩余量
                        '如果内燃机100%发电功率能够被蓄热+第一顺序+第二顺序设备消纳
                        If NRJGDGL <= XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) Then
                            '制热设备负荷第三段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热+内燃机恰好为100%负荷时第二顺序制热功率
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + NRJGDGLSY * D2SXZRCOP, 2)
                            '制热设备负荷第四段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热功率+第二顺序，后面的以此类推
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                        End If
                        '如果内燃机100%发电功率不能够被蓄热+第一顺序+第二顺序设备消纳
                        If NRJGDGL > XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) Then
                            '制热设备负荷第三段=蓄热设备100%制热量+第一顺序制热+第二顺序制热+此时溴化锂供热功率
                            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + XHLZRa * ((D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) + XRGRFJHDGL) / NRJGDGL, 2)
                            NRJGDGLSY = NRJGDGL - XRGRFJHDGL - (D1SXZRGL / D1SXZRCOP) - (D2SXZRGL / D2SXZRCOP) '内燃机供电功率剩余量
                            '如果内燃机100%发电功率能够被蓄热+第一顺序+第二顺序+第三顺序设备消纳
                            If NRJGDGL <= XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) + (D3SXZRGL / D3SXZRCOP) Then
                                '制热设备负荷第四段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热+第二顺序制热+内燃机恰好为100%负荷时第三顺序制热功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + NRJGDGLSY * D3SXZRCOP, 2)
                                '制热设备负荷第五段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热功率+第二顺序+第三顺序，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                            End If
                            '如果内燃机100%发电功率不能够被蓄热+第一顺序+第二顺序+第三顺序设备消纳
                            If NRJGDGL > XRGRFJHDGL + (D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) + (D3SXZRGL / D3SXZRCOP) Then
                                '制热设备负荷第四段=蓄热设备100%制热量+第一顺序制热+第二顺序制热+第三顺序+此时溴化锂供热功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + XHLZRa * ((D1SXZRGL / D1SXZRCOP) + (D2SXZRGL / D2SXZRCOP) + (D3SXZRGL / D3SXZRCOP) + XRGRFJHDGL) / NRJGDGL, 2)
                                NRJGDGLSY = NRJGDGL - XRGRFJHDGL - (D1SXZRGL / D1SXZRCOP) - (D2SXZRGL / D2SXZRCOP) '内燃机供电功率剩余量
                                '为了计算方便，假定内燃机100%供电量能够被第四顺序设备设备消纳
                                '制热设备负荷第五段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热+第二顺序制热+第三顺序制热+内燃机恰好为100%负荷时第四顺序制热功率
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + NRJGDGLSY * D4SXZRCOP, 2)
                                '制热设备负荷第六段=蓄热设备100%制热量+溴化锂100%制热功率+第一顺序制热功率+第二顺序+第三顺序+第四顺序，后面的以此类推
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
                            End If
                        End If
                    End If
                End If
            End If
            '如果内燃机可以向外供电
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "Y" Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(92, 4).Value = Math.Round(XRGRGLa, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + XHLZRa30p, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + XHLZRa, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(99, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(100, 4).Value = Math.Round(XRGRGLa + XHLZRa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
            End If
        End If
        '如果没有内燃机装机量，采用以下计算
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(47, 7).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(92, 4).Value = Math.Round(XRGRGLa, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(93, 4).Value = Math.Round(XRGRGLa + D1SXZRGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(94, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(95, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(96, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(97, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL, 2)
            ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(98, 4).Value = Math.Round(XRGRGLa + D1SXZRGL + D2SXZRGL + D3SXZRGL + D4SXZRGL + D5SXZRGL + D6SXZRGL, 2)
        End If
        '将重复的值删除
        For i = 8 To 1 Step -1 '从下向上查找，把紧挨着的重复的值去掉
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value = Nothing
            End If
        Next
        For i = 1 To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value <= ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value = Nothing
            End If
        Next
        '如果有为0的结果，将后面的值前移
        Dim KSXH As Integer = 1 '开始检索的序号，默认值为1
        Dim js As Integer = 0 '计数
        For i = 2 To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value <> Nothing Then
                js = js + 1
            End If
        Next
        If js > 1 Then '如果有超过两段，从第一个开始检索，否则第二个
            KSXH = 1
        Else
            KSXH = 2
        End If
        For i = KSXH To 9
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value < ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value Then
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(90 + i, 4).Value = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value
                ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(91 + i, 4).Value = Nothing
            End If
        Next
        '清理部分数据
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 4).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 5).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 77).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 22).Value = Nothing
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 25).Value = Nothing
    End Sub

End Module
