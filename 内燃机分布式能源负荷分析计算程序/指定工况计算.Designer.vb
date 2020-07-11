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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.常规计算模式 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.FHTJJD_shuru = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.全局寻优计算 = New System.Windows.Forms.Button()
        Me.TRQDJ = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.GDDJ = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(390, 159)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(108, 35)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(32, 28)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(488, 62)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "请输入需要进行指定计算的工况序号"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(32, 80)
        Me.Label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(488, 62)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "一次最多可以指定计算五个不同的工况"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(56, 148)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(322, 56)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "第一个指定计算工况序号："
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(390, 220)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(6)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(108, 35)
        Me.TextBox2.TabIndex = 8
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox3
        '
        Me.TextBox3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(390, 281)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(6)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(108, 35)
        Me.TextBox3.TabIndex = 10
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox4
        '
        Me.TextBox4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(390, 344)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(6)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(108, 35)
        Me.TextBox4.TabIndex = 12
        Me.TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox5
        '
        Me.TextBox5.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox5.Location = New System.Drawing.Point(390, 404)
        Me.TextBox5.Margin = New System.Windows.Forms.Padding(6)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(108, 35)
        Me.TextBox5.TabIndex = 14
        Me.TextBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '常规计算模式
        '
        Me.常规计算模式.BackColor = System.Drawing.SystemColors.ControlLight
        Me.常规计算模式.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.常规计算模式.Location = New System.Drawing.Point(36, 663)
        Me.常规计算模式.Margin = New System.Windows.Forms.Padding(6)
        Me.常规计算模式.Name = "常规计算模式"
        Me.常规计算模式.Size = New System.Drawing.Size(236, 82)
        Me.常规计算模式.TabIndex = 16
        Me.常规计算模式.Text = "常规计算模式"
        Me.常规计算模式.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(56, 209)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(322, 56)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "第二个指定计算工况序号："
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(56, 270)
        Me.Label5.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(322, 56)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "第三个指定计算工况序号："
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(56, 333)
        Me.Label6.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(322, 56)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "第四个指定计算工况序号："
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(56, 394)
        Me.Label7.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(322, 56)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "第五个指定计算工况序号："
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FHTJJD_shuru
        '
        Me.FHTJJD_shuru.Location = New System.Drawing.Point(390, 466)
        Me.FHTJJD_shuru.Name = "FHTJJD_shuru"
        Me.FHTJJD_shuru.Size = New System.Drawing.Size(108, 35)
        Me.FHTJJD_shuru.TabIndex = 22
        Me.FHTJJD_shuru.Text = "0.5"
        Me.FHTJJD_shuru.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(60, 472)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(324, 24)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "输入设备负荷调节精度(%)："
        '
        '全局寻优计算
        '
        Me.全局寻优计算.BackColor = System.Drawing.SystemColors.ControlLight
        Me.全局寻优计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.全局寻优计算.Location = New System.Drawing.Point(302, 663)
        Me.全局寻优计算.Margin = New System.Windows.Forms.Padding(6)
        Me.全局寻优计算.Name = "全局寻优计算"
        Me.全局寻优计算.Size = New System.Drawing.Size(236, 82)
        Me.全局寻优计算.TabIndex = 23
        Me.全局寻优计算.Text = "全局寻优计算"
        Me.全局寻优计算.UseVisualStyleBackColor = False
        '
        'TRQDJ
        '
        Me.TRQDJ.Location = New System.Drawing.Point(390, 589)
        Me.TRQDJ.Name = "TRQDJ"
        Me.TRQDJ.Size = New System.Drawing.Size(108, 35)
        Me.TRQDJ.TabIndex = 28
        Me.TRQDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(60, 596)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(287, 24)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "天然气单价（元/Nm3）："
        '
        'GDDJ
        '
        Me.GDDJ.Location = New System.Drawing.Point(390, 527)
        Me.GDDJ.Name = "GDDJ"
        Me.GDDJ.Size = New System.Drawing.Size(108, 35)
        Me.GDDJ.TabIndex = 25
        Me.GDDJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(60, 536)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(262, 24)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "购电单价（元/kWh）："
        '
        '指定工况计算
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(568, 780)
        Me.Controls.Add(Me.TRQDJ)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.GDDJ)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.全局寻优计算)
        Me.Controls.Add(Me.FHTJJD_shuru)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.常规计算模式)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.Name = "指定工况计算"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "指定工况计算"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents 常规计算模式 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents FHTJJD_shuru As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents 全局寻优计算 As System.Windows.Forms.Button
    Friend WithEvents TRQDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GDDJ As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
