<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 计算模式选择
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.常规计算模式 = New System.Windows.Forms.Button()
        Me.全局寻优计算 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.FHTJJD_shuru = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GDDJ_GF1 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TRQDJ = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.GDDJ_GF2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_F1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_F2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_P1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_P2 = New System.Windows.Forms.TextBox()
        Me.GDDJ_QT1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_QT2 = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.GDDJ_G1 = New System.Windows.Forms.TextBox()
        Me.GDDJ_G2 = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.TS_XZ = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(223, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(302, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "方法一：常规计算模式"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(223, 143)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(302, 28)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "方法二：全局寻优计算"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(31, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(658, 24)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "不能充分体现各设备变负荷运行时的性能特性，但计算速度快"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(31, 199)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(682, 24)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "可以充分体现各设备变负荷运行时的性能特性，但计算速度缓慢"
        '
        '常规计算模式
        '
        Me.常规计算模式.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.常规计算模式.Location = New System.Drawing.Point(35, 536)
        Me.常规计算模式.Name = "常规计算模式"
        Me.常规计算模式.Size = New System.Drawing.Size(284, 100)
        Me.常规计算模式.TabIndex = 4
        Me.常规计算模式.Text = "常规计算模式"
        Me.常规计算模式.UseVisualStyleBackColor = True
        '
        '全局寻优计算
        '
        Me.全局寻优计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.全局寻优计算.Location = New System.Drawing.Point(429, 536)
        Me.全局寻优计算.Name = "全局寻优计算"
        Me.全局寻优计算.Size = New System.Drawing.Size(284, 100)
        Me.全局寻优计算.TabIndex = 5
        Me.全局寻优计算.Text = "全局寻优计算"
        Me.全局寻优计算.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(89, 268)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(331, 28)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "输入设备负荷调节精度："
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(31, 392)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(682, 24)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "输入的负荷调节精度数值越小，计算精度越高，但计算速度稍慢"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(73, 337)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(598, 24)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "输入的负荷调节精度数值必须可以被25整除（单位为%）" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'FHTJJD_shuru
        '
        Me.FHTJJD_shuru.Location = New System.Drawing.Point(431, 264)
        Me.FHTJJD_shuru.Name = "FHTJJD_shuru"
        Me.FHTJJD_shuru.Size = New System.Drawing.Size(170, 35)
        Me.FHTJJD_shuru.TabIndex = 9
        Me.FHTJJD_shuru.Text = "0.5"
        Me.FHTJJD_shuru.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(616, 265)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 32)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "%"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(803, 93)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(260, 28)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "购电单价(高峰1)："
        '
        'GDDJ_GF1
        '
        Me.GDDJ_GF1.Location = New System.Drawing.Point(1078, 90)
        Me.GDDJ_GF1.Name = "GDDJ_GF1"
        Me.GDDJ_GF1.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_GF1.TabIndex = 12
        Me.GDDJ_GF1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(1266, 92)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(111, 32)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "元/kWh"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(823, 604)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(186, 28)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "天然气单价："
        '
        'TRQDJ
        '
        Me.TRQDJ.Location = New System.Drawing.Point(1078, 601)
        Me.TRQDJ.Name = "TRQDJ"
        Me.TRQDJ.Size = New System.Drawing.Size(170, 35)
        Me.TRQDJ.TabIndex = 15
        Me.TRQDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(1269, 602)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(108, 32)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "元/Nm3"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label13.Location = New System.Drawing.Point(764, 31)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(650, 28)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "全局寻优计算模式需要输入购电单价和天然气单价"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label14.Location = New System.Drawing.Point(803, 143)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(260, 28)
        Me.Label14.TabIndex = 18
        Me.Label14.Text = "购电单价(高峰2)："
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label15.Location = New System.Drawing.Point(803, 192)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(231, 28)
        Me.Label15.TabIndex = 19
        Me.Label15.Text = "购电单价(峰1)："
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label16.Location = New System.Drawing.Point(803, 244)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(231, 28)
        Me.Label16.TabIndex = 20
        Me.Label16.Text = "购电单价(峰2)："
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.Location = New System.Drawing.Point(803, 344)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(231, 28)
        Me.Label17.TabIndex = 22
        Me.Label17.Text = "购电单价(平2)："
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label18.Location = New System.Drawing.Point(803, 294)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(231, 28)
        Me.Label18.TabIndex = 21
        Me.Label18.Text = "购电单价(平1)："
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label19.Location = New System.Drawing.Point(803, 500)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(260, 28)
        Me.Label19.TabIndex = 23
        Me.Label19.Text = "购电单价(其它1)："
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label20.Location = New System.Drawing.Point(803, 550)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(260, 28)
        Me.Label20.TabIndex = 24
        Me.Label20.Text = "购电单价(其它2)："
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(1266, 139)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(111, 32)
        Me.Label21.TabIndex = 25
        Me.Label21.Text = "元/kWh"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(1266, 190)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(111, 32)
        Me.Label22.TabIndex = 26
        Me.Label22.Text = "元/kWh"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(1266, 240)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(111, 32)
        Me.Label23.TabIndex = 27
        Me.Label23.Text = "元/kWh"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(1266, 292)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(111, 32)
        Me.Label24.TabIndex = 28
        Me.Label24.Text = "元/kWh"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(1266, 340)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(111, 32)
        Me.Label25.TabIndex = 29
        Me.Label25.Text = "元/kWh"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(1266, 495)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(111, 32)
        Me.Label26.TabIndex = 30
        Me.Label26.Text = "元/kWh"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(1266, 547)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(111, 32)
        Me.Label27.TabIndex = 31
        Me.Label27.Text = "元/kWh"
        '
        'GDDJ_GF2
        '
        Me.GDDJ_GF2.Location = New System.Drawing.Point(1078, 139)
        Me.GDDJ_GF2.Name = "GDDJ_GF2"
        Me.GDDJ_GF2.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_GF2.TabIndex = 32
        Me.GDDJ_GF2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_F1
        '
        Me.GDDJ_F1.Location = New System.Drawing.Point(1078, 189)
        Me.GDDJ_F1.Name = "GDDJ_F1"
        Me.GDDJ_F1.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_F1.TabIndex = 33
        Me.GDDJ_F1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_F2
        '
        Me.GDDJ_F2.Location = New System.Drawing.Point(1078, 240)
        Me.GDDJ_F2.Name = "GDDJ_F2"
        Me.GDDJ_F2.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_F2.TabIndex = 34
        Me.GDDJ_F2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_P1
        '
        Me.GDDJ_P1.Location = New System.Drawing.Point(1078, 289)
        Me.GDDJ_P1.Name = "GDDJ_P1"
        Me.GDDJ_P1.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_P1.TabIndex = 35
        Me.GDDJ_P1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_P2
        '
        Me.GDDJ_P2.Location = New System.Drawing.Point(1078, 340)
        Me.GDDJ_P2.Name = "GDDJ_P2"
        Me.GDDJ_P2.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_P2.TabIndex = 36
        Me.GDDJ_P2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_QT1
        '
        Me.GDDJ_QT1.Location = New System.Drawing.Point(1078, 496)
        Me.GDDJ_QT1.Name = "GDDJ_QT1"
        Me.GDDJ_QT1.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_QT1.TabIndex = 37
        Me.GDDJ_QT1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_QT2
        '
        Me.GDDJ_QT2.Location = New System.Drawing.Point(1078, 547)
        Me.GDDJ_QT2.Name = "GDDJ_QT2"
        Me.GDDJ_QT2.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_QT2.TabIndex = 38
        Me.GDDJ_QT2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label28.Location = New System.Drawing.Point(803, 396)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(231, 28)
        Me.Label28.TabIndex = 39
        Me.Label28.Text = "购电单价(谷1)："
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label29.Location = New System.Drawing.Point(803, 446)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(231, 28)
        Me.Label29.TabIndex = 40
        Me.Label29.Text = "购电单价(谷2)："
        '
        'GDDJ_G1
        '
        Me.GDDJ_G1.Location = New System.Drawing.Point(1078, 393)
        Me.GDDJ_G1.Name = "GDDJ_G1"
        Me.GDDJ_G1.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_G1.TabIndex = 41
        Me.GDDJ_G1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GDDJ_G2
        '
        Me.GDDJ_G2.Location = New System.Drawing.Point(1078, 442)
        Me.GDDJ_G2.Name = "GDDJ_G2"
        Me.GDDJ_G2.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ_G2.TabIndex = 42
        Me.GDDJ_G2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(1266, 391)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(111, 32)
        Me.Label30.TabIndex = 43
        Me.Label30.Text = "元/kWh"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.Location = New System.Drawing.Point(1269, 442)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(111, 32)
        Me.Label31.TabIndex = 44
        Me.Label31.Text = "元/kWh"
        '
        'TS_XZ
        '
        Me.TS_XZ.AutoSize = True
        Me.TS_XZ.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TS_XZ.Location = New System.Drawing.Point(94, 462)
        Me.TS_XZ.Name = "TS_XZ"
        Me.TS_XZ.Size = New System.Drawing.Size(595, 32)
        Me.TS_XZ.TabIndex = 45
        Me.TS_XZ.Text = "当计算过程中出现异常时，是否弹出提示？"
        Me.TS_XZ.UseVisualStyleBackColor = True
        '
        '计算模式选择
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1438, 666)
        Me.Controls.Add(Me.TS_XZ)
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
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.TRQDJ)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.GDDJ_GF1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.FHTJJD_shuru)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.全局寻优计算)
        Me.Controls.Add(Me.常规计算模式)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "计算模式选择"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "计算模式选择"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents 常规计算模式 As System.Windows.Forms.Button
    Friend WithEvents 全局寻优计算 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents FHTJJD_shuru As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_GF1 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TRQDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_GF2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_F1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_F2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_P1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_P2 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_QT1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_QT2 As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents GDDJ_G1 As System.Windows.Forms.TextBox
    Friend WithEvents GDDJ_G2 As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TS_XZ As System.Windows.Forms.CheckBox
End Class
