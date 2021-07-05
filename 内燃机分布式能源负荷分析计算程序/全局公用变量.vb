Module 全局公用变量
    '申明公用变量和数组
    '定义和申明各种全局变量和数组
    '全局寻优计算出现错误的工况序号，加入列表
    Public QJXYJS_ERROR As New List(Of Integer)
    '全局寻优计算模式时，设备会被全部用掉的工况，记录工况序号，加入列表，直接采用常规计算模式进行计算，不在报错
    Public QJXY_CGMS As New List(Of Integer)
    '定义数组，用于储存输入的各个工况冷热负荷需求量(kW)
    Public LFHZXQL(10000) As Double '冷负荷总需求量
    Public RFHZXQL(10000) As Double '热负荷总需求量
    '定义数组用于储存蓄能装置的供冷功率，蓄冷功率，供热功率，蓄热功率
    Public XNGLGL(10000) As Double '供冷功率
    Public XNXLGL(10000) As Double '蓄冷功率
    Public XNGRGL(10000) As Double '供热功率
    Public XNXRGL(10000) As Double '蓄热功率
    '各个设备可以允许运行的最低负荷率
    '内燃机
    Public FHL1_min_NRJ As Double = 0.3
    Public FHL2_min_NRJ As Double = 0.3
    '离心式冷水机
    Public FHL1_min_LXSLSJ As Double = 0.1
    Public FHL2_min_LXSLSJ As Double = 0.1
    '水冷螺杆机
    Public FHL1_min_SLLGJ As Double = 0.1
    Public FHL2_min_SLLGJ As Double = 0.1
    '风冷螺杆机
    Public FHL1_min_FLLGJ As Double = 0.1
    Public FHL2_min_FLLGJ As Double = 0.1
    '水（地）源热泵
    Public FHL1_min_SDYRB As Double = 0.1
    Public FHL2_min_SDYRB As Double = 0.1
    '离心式热泵
    Public FHL1_min_LXSRB As Double = 0.1
    Public FHL2_min_LXSRB As Double = 0.1
    '空气源热泵
    Public FHL1_min_KQYRB As Double = 0.1
    Public FHL2_min_KQYRB As Double = 0.1
    '直燃型溴化锂
    Public FHL1_min_ZRXXHL As Double = 0.5
    Public FHL2_min_ZRXXHL As Double = 0.5
    '天然气采暖锅炉
    Public FHL1_min_TRQGL As Double = 0.1
    Public FHL2_min_TRQGL As Double = 0.1
    '电采暖锅炉
    Public FHL1_min_DGL As Double = 0.1
    Public FHL2_min_DGL As Double = 0.1
    '装机功率容错系数（程序在计算时候采用的设备（1）（2）装机功率需要比Excel读入的数据略大，防止程序出错）
    'RCXS=50为初始值，此时FHTJJD=0.5，之后按比例增加
    Public RCXS As Double = 50
    '用于显示计算进度的窗体
    '指定工况计算窗体
    'Public Shared Form1 As New 指定工况计算
    'Public Shared Form2 As New 进入维护模式
End Module
