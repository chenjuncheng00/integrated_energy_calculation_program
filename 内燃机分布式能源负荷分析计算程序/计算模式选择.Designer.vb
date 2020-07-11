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
        Me.GDDJ = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TRQDJ = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
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
        Me.常规计算模式.Location = New System.Drawing.Point(35, 653)
        Me.常规计算模式.Name = "常规计算模式"
        Me.常规计算模式.Size = New System.Drawing.Size(284, 100)
        Me.常规计算模式.TabIndex = 4
        Me.常规计算模式.Text = "常规计算模式"
        Me.常规计算模式.UseVisualStyleBackColor = True
        '
        '全局寻优计算
        '
        Me.全局寻优计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.全局寻优计算.Location = New System.Drawing.Point(429, 653)
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
        Me.Label6.Text = "输入的负荷调节精度数值越小，计算精度越高，但计算速度越慢"
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
        Me.Label9.Location = New System.Drawing.Point(137, 515)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(157, 28)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "购电单价："
        '
        'GDDJ
        '
        Me.GDDJ.Location = New System.Drawing.Point(305, 512)
        Me.GDDJ.Name = "GDDJ"
        Me.GDDJ.Size = New System.Drawing.Size(170, 35)
        Me.GDDJ.TabIndex = 12
        Me.GDDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(490, 514)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(111, 32)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "元/kWh"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(108, 578)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(186, 28)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "天然气单价："
        '
        'TRQDJ
        '
        Me.TRQDJ.Location = New System.Drawing.Point(305, 578)
        Me.TRQDJ.Name = "TRQDJ"
        Me.TRQDJ.Size = New System.Drawing.Size(170, 35)
        Me.TRQDJ.TabIndex = 15
        Me.TRQDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(497, 578)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(108, 32)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "元/Nm3"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label13.Location = New System.Drawing.Point(50, 455)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(650, 28)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "全局寻优计算模式需要输入购电单价和天然气单价"
        '
        '计算模式选择
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(748, 794)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.TRQDJ)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.GDDJ)
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
    Friend WithEvents GDDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TRQDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
End Class
