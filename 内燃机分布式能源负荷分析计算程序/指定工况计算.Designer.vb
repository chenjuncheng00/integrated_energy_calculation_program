<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 指定工况计算
    Inherits System.Windows.Forms.Form
    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer
    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.工况序号tmp = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.常规计算模式 = New System.Windows.Forms.Button()
        Me.FHTJJD_shuru = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.全局寻优计算 = New System.Windows.Forms.Button()
        Me.TRQDJ = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.工况序号列表 = New System.Windows.Forms.ListBox()
        Me.清空输入 = New System.Windows.Forms.Button()
        Me.添加输入 = New System.Windows.Forms.Button()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.GDDJ_G2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_G1 = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.GDDJ_QT2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_QT1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_P2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_P1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_F2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_F1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_GF2 = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GDDJ_GF1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        '工况序号tmp
        '
        Me.工况序号tmp.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.工况序号tmp.Location = New System.Drawing.Point(234, 456)
        Me.工况序号tmp.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.工况序号tmp.Name = "工况序号tmp"
        Me.工况序号tmp.Size = New System.Drawing.Size(160, 35)
        Me.工况序号tmp.TabIndex = 0
        Me.工况序号tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(42, 22)
        Me.Label3.Margin = New System.Windows.Forms.Padding(7, 0, 7, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(540, 56)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "输入需要计算的工况序号"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '常规计算模式
        '
        Me.常规计算模式.BackColor = System.Drawing.SystemColors.ControlLight
        Me.常规计算模式.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.常规计算模式.Location = New System.Drawing.Point(326, 610)
        Me.常规计算模式.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.常规计算模式.Name = "常规计算模式"
        Me.常规计算模式.Size = New System.Drawing.Size(256, 82)
        Me.常规计算模式.TabIndex = 16
        Me.常规计算模式.Text = "常规计算模式"
        Me.常规计算模式.UseVisualStyleBackColor = False
        '
        'FHTJJD_shuru
        '
        Me.FHTJJD_shuru.Location = New System.Drawing.Point(919, 96)
        Me.FHTJJD_shuru.Name = "FHTJJD_shuru"
        Me.FHTJJD_shuru.Size = New System.Drawing.Size(145, 35)
        Me.FHTJJD_shuru.TabIndex = 22
        Me.FHTJJD_shuru.Text = "0.5"
        Me.FHTJJD_shuru.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(670, 100)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(235, 24)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "设备负荷调节精度："
        '
        '全局寻优计算
        '
        Me.全局寻优计算.BackColor = System.Drawing.SystemColors.ControlLight
        Me.全局寻优计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.全局寻优计算.Location = New System.Drawing.Point(42, 610)
        Me.全局寻优计算.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.全局寻优计算.Name = "全局寻优计算"
        Me.全局寻优计算.Size = New System.Drawing.Size(256, 82)
        Me.全局寻优计算.TabIndex = 23
        Me.全局寻优计算.Text = "全局寻优计算"
        Me.全局寻优计算.UseVisualStyleBackColor = False
        '
        'TRQDJ
        '
        Me.TRQDJ.Location = New System.Drawing.Point(919, 657)
        Me.TRQDJ.Name = "TRQDJ"
        Me.TRQDJ.Size = New System.Drawing.Size(145, 35)
        Me.TRQDJ.TabIndex = 28
        Me.TRQDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(670, 660)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(160, 24)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "天然气单价："
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(674, 19)
        Me.Label2.Margin = New System.Windows.Forms.Padding(7, 0, 7, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(499, 62)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "计算参数输入"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '工况序号列表
        '
        Me.工况序号列表.FormattingEnabled = True
        Me.工况序号列表.ItemHeight = 24
        Me.工况序号列表.Location = New System.Drawing.Point(234, 96)
        Me.工况序号列表.Name = "工况序号列表"
        Me.工况序号列表.Size = New System.Drawing.Size(160, 340)
        Me.工况序号列表.TabIndex = 30
        '
        '清空输入
        '
        Me.清空输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空输入.Location = New System.Drawing.Point(326, 509)
        Me.清空输入.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.清空输入.Name = "清空输入"
        Me.清空输入.Size = New System.Drawing.Size(256, 82)
        Me.清空输入.TabIndex = 75
        Me.清空输入.Text = "清空输入"
        Me.清空输入.UseVisualStyleBackColor = False
        '
        '添加输入
        '
        Me.添加输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.添加输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.添加输入.Location = New System.Drawing.Point(42, 509)
        Me.添加输入.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.添加输入.Name = "添加输入"
        Me.添加输入.Size = New System.Drawing.Size(256, 82)
        Me.添加输入.TabIndex = 74
        Me.添加输入.Text = "添加输入"
        Me.添加输入.UseVisualStyleBackColor = False
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label31.Location = New System.Drawing.Point(1086, 504)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(87, 24)
        Me.Label31.TabIndex = 105
        Me.Label31.Text = "元/kWh"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label30.Location = New System.Drawing.Point(1082, 453)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(87, 24)
        Me.Label30.TabIndex = 104
        Me.Label30.Text = "元/kWh"
        '
        'GDDJ_G2
        '
        Me.GDDJ_G2.Location = New System.Drawing.Point(919, 497)
        Me.GDDJ_G2.Name = "GDDJ_G2"
        Me.GDDJ_G2.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_G2.TabIndex = 103
        Me.GDDJ_G2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_G1
        '
        Me.GDDJ_G1.Location = New System.Drawing.Point(919, 448)
        Me.GDDJ_G1.Name = "GDDJ_G1"
        Me.GDDJ_G1.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_G1.TabIndex = 102
        Me.GDDJ_G1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label29.Location = New System.Drawing.Point(670, 501)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(199, 24)
        Me.Label29.TabIndex = 101
        Me.Label29.Text = "购电单价(谷2)："
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label28.Location = New System.Drawing.Point(670, 451)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(199, 24)
        Me.Label28.TabIndex = 100
        Me.Label28.Text = "购电单价(谷1)："
        '
        'GDDJ_QT2
        '
        Me.GDDJ_QT2.Location = New System.Drawing.Point(919, 602)
        Me.GDDJ_QT2.Name = "GDDJ_QT2"
        Me.GDDJ_QT2.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_QT2.TabIndex = 99
        Me.GDDJ_QT2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_QT1
        '
        Me.GDDJ_QT1.Location = New System.Drawing.Point(919, 551)
        Me.GDDJ_QT1.Name = "GDDJ_QT1"
        Me.GDDJ_QT1.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_QT1.TabIndex = 98
        Me.GDDJ_QT1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_P2
        '
        Me.GDDJ_P2.Location = New System.Drawing.Point(919, 395)
        Me.GDDJ_P2.Name = "GDDJ_P2"
        Me.GDDJ_P2.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_P2.TabIndex = 97
        Me.GDDJ_P2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_P1
        '
        Me.GDDJ_P1.Location = New System.Drawing.Point(919, 344)
        Me.GDDJ_P1.Name = "GDDJ_P1"
        Me.GDDJ_P1.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_P1.TabIndex = 96
        Me.GDDJ_P1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_F2
        '
        Me.GDDJ_F2.Location = New System.Drawing.Point(919, 295)
        Me.GDDJ_F2.Name = "GDDJ_F2"
        Me.GDDJ_F2.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_F2.TabIndex = 95
        Me.GDDJ_F2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_F1
        '
        Me.GDDJ_F1.Location = New System.Drawing.Point(919, 244)
        Me.GDDJ_F1.Name = "GDDJ_F1"
        Me.GDDJ_F1.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_F1.TabIndex = 94
        Me.GDDJ_F1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_GF2
        '
        Me.GDDJ_GF2.Location = New System.Drawing.Point(919, 194)
        Me.GDDJ_GF2.Name = "GDDJ_GF2"
        Me.GDDJ_GF2.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_GF2.TabIndex = 93
        Me.GDDJ_GF2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label27.Location = New System.Drawing.Point(1082, 609)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(87, 24)
        Me.Label27.TabIndex = 92
        Me.Label27.Text = "元/kWh"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label26.Location = New System.Drawing.Point(1082, 557)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(87, 24)
        Me.Label26.TabIndex = 91
        Me.Label26.Text = "元/kWh"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label25.Location = New System.Drawing.Point(1082, 402)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(87, 24)
        Me.Label25.TabIndex = 90
        Me.Label25.Text = "元/kWh"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label24.Location = New System.Drawing.Point(1082, 354)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(87, 24)
        Me.Label24.TabIndex = 89
        Me.Label24.Text = "元/kWh"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label23.Location = New System.Drawing.Point(1082, 302)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(87, 24)
        Me.Label23.TabIndex = 88
        Me.Label23.Text = "元/kWh"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label22.Location = New System.Drawing.Point(1082, 252)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(87, 24)
        Me.Label22.TabIndex = 87
        Me.Label22.Text = "元/kWh"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label21.Location = New System.Drawing.Point(1082, 200)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(87, 24)
        Me.Label21.TabIndex = 86
        Me.Label21.Text = "元/kWh"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label20.Location = New System.Drawing.Point(670, 605)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(224, 24)
        Me.Label20.TabIndex = 85
        Me.Label20.Text = "购电单价(其它2)："
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label19.Location = New System.Drawing.Point(670, 555)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(224, 24)
        Me.Label19.TabIndex = 84
        Me.Label19.Text = "购电单价(其它1)："
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.Location = New System.Drawing.Point(670, 399)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(199, 24)
        Me.Label17.TabIndex = 83
        Me.Label17.Text = "购电单价(平2)："
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label18.Location = New System.Drawing.Point(670, 349)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(199, 24)
        Me.Label18.TabIndex = 82
        Me.Label18.Text = "购电单价(平1)："
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label16.Location = New System.Drawing.Point(670, 299)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(199, 24)
        Me.Label16.TabIndex = 81
        Me.Label16.Text = "购电单价(峰2)："
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label15.Location = New System.Drawing.Point(670, 247)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(199, 24)
        Me.Label15.TabIndex = 80
        Me.Label15.Text = "购电单价(峰1)："
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label14.Location = New System.Drawing.Point(670, 198)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(224, 24)
        Me.Label14.TabIndex = 79
        Me.Label14.Text = "购电单价(高峰2)："
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label10.Location = New System.Drawing.Point(1082, 150)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(87, 24)
        Me.Label10.TabIndex = 78
        Me.Label10.Text = "元/kWh"
        '
        'GDDJ_GF1
        '
        Me.GDDJ_GF1.Location = New System.Drawing.Point(919, 145)
        Me.GDDJ_GF1.Name = "GDDJ_GF1"
        Me.GDDJ_GF1.Size = New System.Drawing.Size(145, 35)
        Me.GDDJ_GF1.TabIndex = 77
        Me.GDDJ_GF1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(670, 148)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 24)
        Me.Label1.TabIndex = 76
        Me.Label1.Text = "购电单价(高峰1)："
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(1086, 662)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 24)
        Me.Label4.TabIndex = 106
        Me.Label4.Text = "元/Nm3"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(1085, 100)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 26)
        Me.Label5.TabIndex = 107
        Me.Label5.Text = "%"
        '
        '指定工况计算
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1222, 724)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.GDDJ_G2)
        Me.Controls.Add(Me.GDDJ_G1)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.GDDJ_QT2)
        Me.Controls.Add(Me.GDDJ_QT1)
        Me.Controls.Add(Me.GDDJ_P2)
        Me.Controls.Add(Me.GDDJ_P1)
        Me.Controls.Add(Me.GDDJ_F2)
        Me.Controls.Add(Me.GDDJ_F1)
        Me.Controls.Add(Me.GDDJ_GF2)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.GDDJ_GF1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.清空输入)
        Me.Controls.Add(Me.添加输入)
        Me.Controls.Add(Me.工况序号列表)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TRQDJ)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.全局寻优计算)
        Me.Controls.Add(Me.FHTJJD_shuru)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.常规计算模式)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.工况序号tmp)
        Me.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.Name = "指定工况计算"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "指定工况计算"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents 工况序号tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents 常规计算模式 As System.Windows.Forms.Button
    Friend WithEvents FHTJJD_shuru As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents 全局寻优计算 As System.Windows.Forms.Button
    Friend WithEvents TRQDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents 工况序号列表 As System.Windows.Forms.ListBox
    Friend WithEvents 清空输入 As System.Windows.Forms.Button
    Friend WithEvents 添加输入 As System.Windows.Forms.Button
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_G2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_G1 As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_QT2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_QT1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_P2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_P1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_F2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_F1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_GF2 As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_GF1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
