<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 蓄能分配计算设置
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.XLGL_PJ = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.XLGL_MAX = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.XRGL_MAX = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.XRGL_PJ = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.确定参数 = New System.Windows.Forms.Button()
        Me.清空数据 = New System.Windows.Forms.Button()
        Me.XL_MAX_CHECK = New System.Windows.Forms.CheckBox()
        Me.XR_MAX_CHECK = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(118, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(302, 28)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "蓄冷分配计算参数设置"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(118, 271)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(302, 28)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "蓄热分配计算参数设置"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(445, 96)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(58, 32)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "kW"
        '
        'XLGL_PJ
        '
        Me.XLGL_PJ.Location = New System.Drawing.Point(259, 94)
        Me.XLGL_PJ.Name = "XLGL_PJ"
        Me.XLGL_PJ.Size = New System.Drawing.Size(170, 35)
        Me.XLGL_PJ.TabIndex = 15
        Me.XLGL_PJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(48, 98)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(215, 28)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "蓄冷平均功率："
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(48, 209)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(215, 28)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "蓄冷最大功率："
        '
        'XLGL_MAX
        '
        Me.XLGL_MAX.Location = New System.Drawing.Point(259, 205)
        Me.XLGL_MAX.Name = "XLGL_MAX"
        Me.XLGL_MAX.Size = New System.Drawing.Size(170, 35)
        Me.XLGL_MAX.TabIndex = 18
        Me.XLGL_MAX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(445, 207)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 32)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "kW"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(445, 449)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 32)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "kW"
        '
        'XRGL_MAX
        '
        Me.XRGL_MAX.Location = New System.Drawing.Point(259, 447)
        Me.XRGL_MAX.Name = "XRGL_MAX"
        Me.XRGL_MAX.Size = New System.Drawing.Size(170, 35)
        Me.XRGL_MAX.TabIndex = 24
        Me.XRGL_MAX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(48, 450)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(215, 28)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "蓄热最大功率："
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(445, 329)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 32)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "kW"
        '
        'XRGL_PJ
        '
        Me.XRGL_PJ.Location = New System.Drawing.Point(259, 327)
        Me.XRGL_PJ.Name = "XRGL_PJ"
        Me.XRGL_PJ.Size = New System.Drawing.Size(170, 35)
        Me.XRGL_PJ.TabIndex = 21
        Me.XRGL_PJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(48, 331)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(215, 28)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "蓄热平均功率："
        '
        '确定参数
        '
        Me.确定参数.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.确定参数.Location = New System.Drawing.Point(53, 520)
        Me.确定参数.Name = "确定参数"
        Me.确定参数.Size = New System.Drawing.Size(190, 84)
        Me.确定参数.TabIndex = 26
        Me.确定参数.Text = "确定参数"
        Me.确定参数.UseVisualStyleBackColor = True
        '
        '清空数据
        '
        Me.清空数据.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空数据.Location = New System.Drawing.Point(313, 520)
        Me.清空数据.Name = "清空数据"
        Me.清空数据.Size = New System.Drawing.Size(190, 84)
        Me.清空数据.TabIndex = 27
        Me.清空数据.Text = "清空数据"
        Me.清空数据.UseVisualStyleBackColor = True
        '
        'XL_MAX_CHECK
        '
        Me.XL_MAX_CHECK.AutoSize = True
        Me.XL_MAX_CHECK.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.XL_MAX_CHECK.Location = New System.Drawing.Point(53, 153)
        Me.XL_MAX_CHECK.Name = "XL_MAX_CHECK"
        Me.XL_MAX_CHECK.Size = New System.Drawing.Size(421, 32)
        Me.XL_MAX_CHECK.TabIndex = 28
        Me.XL_MAX_CHECK.Text = "蓄冷分配计算存在最大值限制"
        Me.XL_MAX_CHECK.UseVisualStyleBackColor = True
        '
        'XR_MAX_CHECK
        '
        Me.XR_MAX_CHECK.AutoSize = True
        Me.XR_MAX_CHECK.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.XR_MAX_CHECK.Location = New System.Drawing.Point(63, 389)
        Me.XR_MAX_CHECK.Name = "XR_MAX_CHECK"
        Me.XR_MAX_CHECK.Size = New System.Drawing.Size(421, 32)
        Me.XR_MAX_CHECK.TabIndex = 29
        Me.XR_MAX_CHECK.Text = "蓄热功率分配计算最大值限制"
        Me.XR_MAX_CHECK.UseVisualStyleBackColor = True
        '
        '蓄能分配计算设置
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 650)
        Me.Controls.Add(Me.XR_MAX_CHECK)
        Me.Controls.Add(Me.XL_MAX_CHECK)
        Me.Controls.Add(Me.清空数据)
        Me.Controls.Add(Me.确定参数)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.XRGL_MAX)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.XRGL_PJ)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.XLGL_MAX)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.XLGL_PJ)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "蓄能分配计算设置"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "蓄能分配计算设置"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents XLGL_PJ As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents XLGL_MAX As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents XRGL_MAX As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents XRGL_PJ As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents 确定参数 As System.Windows.Forms.Button
    Friend WithEvents 清空数据 As System.Windows.Forms.Button
    Friend WithEvents XL_MAX_CHECK As System.Windows.Forms.CheckBox
    Friend WithEvents XR_MAX_CHECK As System.Windows.Forms.CheckBox
End Class
