Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

<ComClass(Com内燃机分布式能源负荷分析计算程序.ClassId, Com内燃机分布式能源负荷分析计算程序.InterfaceId, Com内燃机分布式能源负荷分析计算程序.EventsId)>
Public Class Com内燃机分布式能源负荷分析计算程序
#Region "COM GUID"
    ' 这些 GUID 提供此类的 COM 标识 
    ' 及其 COM 接口。若更改它们，则现有的
    ' 客户端将不再能访问此类。
    Public Const ClassId As String = "9cdda2a4-549b-45c8-8f3b-bf37b5dc0793"
    Public Const InterfaceId As String = "21cd8fe7-d572-45d4-829e-bed25960763d"
    Public Const EventsId As String = "65b11446-71af-4e54-a69d-21b93ad5871c"
#End Region
    ' 可创建的 COM 类必须具有一个不带参数的 Sub New() 
    ' 否则， 将不会在 
    ' COM 注册表中注册此类，且无法通过
    ' CreateObject 创建此类。
    Sub New()
        MyBase.New()
    End Sub
    Sub MainCaculationProgram()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '验证版本号
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        Call 解锁工作表(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '选择计算模式
        '实例化窗体并显示
        Dim Form_JSMSXZ As New 计算模式选择
        Form_JSMSXZ.ShowDialog() '窗口显示
        Form_JSMSXZ.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Call 锁定工作表(ExcelApp)
    End Sub

    Sub 指定工况计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        Call 解锁工作表(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim Form1 As New 指定工况计算
        Form1.ShowDialog() '窗口显示
        Form1.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '重新锁定工作表
        Call 锁定工作表(ExcelApp)
    End Sub
    Sub 冷热负荷分段计算()
        On Error Resume Next
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        Call 解锁工作表(ExcelApp)
        Call 清空输入输出数据(ExcelApp)
        Call 制冷设备负荷分段(ExcelApp)
        Call 制热设备负荷分段(ExcelApp)
        Call 锁定工作表(ExcelApp)
    End Sub
    Sub 蓄能分配计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
        Dim Form_XNFP As New 蓄能分配计算设置
        Form_XNFP.ShowDialog() '窗口显示
        Form_XNFP.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 锁定工作表(ExcelApp)
    End Sub
    Sub 典型日输入参数赋值()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
        Dim Form_DXRSR As New 典型日输入参数赋值
        Form_DXRSR.ShowDialog() '窗口显示
        Form_DXRSR.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 锁定工作表(ExcelApp)
    End Sub
    Sub 打开表格自动运行()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '打开自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件触发
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(3, 20).Value < 20200711 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
        End If
        Call 锁定工作表(ExcelApp)
        Call 工作表设置为彻底隐藏(ExcelApp)
        Call 自保护程序(ExcelApp)
    End Sub
    Sub 清空全部数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————
        Call 解锁工作表(ExcelApp)
        '清空计算输入输出数据
        Call 清空输入输出数据(ExcelApp)
        '清空计算输入的用户输入量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B8:EK207").ClearContents
        '将所有数据重置为0
        '将所有数据变成0，计算一次，以清空全部数据
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 97)).Value = 0
        '各种修正系数默认值均为1
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 98), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 109)).Value = 1
        '全局寻优计算模式设备启动数量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 110), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = Nothing
        '读取计算输入量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(3, 141)).Value = ExcelApp.ThisWorkbook.Worksheets("计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 2), ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(8, 141)).Value
        '返回计算结果
        ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(8, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(8, 54)).Value = ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("计算结果输出").Cells(3, 54)).Value
        '再清空一次全部数据
        Call 清空输入输出数据(ExcelApp)
        '情况计算输入的用户输入量
        ExcelApp.ThisWorkbook.Worksheets("计算输入").Range("B8:EK207").ClearContents
        '清空冷热负荷分段计算结果
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Range("D81:D89").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Range("D92:D100").ClearContents
        '重新锁定工作表
        Call 锁定工作表(ExcelApp)
    End Sub
    Sub 进入维护模式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim ZTJC_EXCEL As Integer = Excel版本号验证(ExcelApp)
        If ZTJC_EXCEL = 1 Then
            Call 锁定工作表(ExcelApp)
            ZTJC_EXCEL = 0
            Exit Sub
        End If
        Dim Form2 As New 进入维护模式
        '读取计数
        If (ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value < 3 And ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value >= 0) Then '最多只可以连续错3次。单元格S1
            Form2.ShowDialog() '窗口显示
            Form2.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            Form2.TextBox1.Text = Nothing '清空已有的内容
            '剩下的内容转入“进入维护模式.确定”
        Else
            MsgBox("已超过最大尝试次数，不可以再尝试输入密码！")
            Call 工作表设置为彻底隐藏(ExcelApp)
            Call 锁定工作表(ExcelApp)
        End If
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

End Class
