Module 设备效率曲线
    Function 内燃机发电效率曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的效率相对100%负荷率时的比例
        Dim ans As Double
        ans = -0.0422 * FHL ^ 4 + 0.2262 * FHL ^ 3 - 0.5102 * FHL ^ 2 + 0.5929 * FHL + 0.734
        '返回结果
        Return ans
    End Function
    Function 内燃机余热效率曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的效率相对100%负荷率时的比例
        Dim ans As Double
        ans = 0.0862 * FHL ^ 2 - 0.2543 * FHL + 1.1681
        '返回结果
        Return ans
    End Function
    Function 烟气热水型溴化锂制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        Dim ans As Double
        ans = 0.0762 * FHL ^ 2 - 0.3048 * FHL + 1.2286
        '返回结果
        Return ans
    End Function
    Function 烟气热水型溴化锂制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        Dim ans As Double
        ans = 0.0167 * FHL ^ 2 - 0.0458 * FHL + 1.0292
        '返回结果
        Return ans
    End Function
    Function 生活热水余热锅炉效率曲线(FHL As Double)
        On Error Resume Next
        '生活热水余热锅炉和工业蒸汽余热锅炉效率变化率采用同样的曲线
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的效率相对100%负荷率时的比例
        Dim ans As Double
        ans = -0.0802 * FHL ^ 2 - 0.1327 * FHL + 1.2134
        '返回结果
        Return ans
    End Function
    Function 工业蒸汽余热锅炉效率曲线(FHL As Double)
        On Error Resume Next
        '生活热水余热锅炉和工业蒸汽余热锅炉效率变化率采用同样的曲线
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的效率相对100%负荷率时的比例
        Dim ans As Double
        ans = -0.0802 * FHL ^ 2 - 0.1327 * FHL + 1.2134
        '返回结果
        Return ans
    End Function
    Function 离心式冷水机制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        Dim ans As Double
        '负荷率小于等于0.5
        If FHL <= 0.5 Then
            ans = 31.385 * FHL ^ 3 - 35.663 * FHL ^ 2 + 15.017 * FHL - 0.3516
        Else
            ans = 5.373 * FHL ^ 3 - 9.0827 * FHL ^ 2 + 1.8982 * FHL + 2.8083
        End If
        '返回结果
        Return ans
    End Function
    Function 水冷螺杆机制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        '与离心式冷水机暂时采用同样的曲线
        Dim ans As Double
        '负荷率小于等于0.5
        If FHL <= 0.5 Then
            ans = 31.385 * FHL ^ 3 - 35.663 * FHL ^ 2 + 15.017 * FHL - 0.3516
        Else
            ans = 5.373 * FHL ^ 3 - 9.0827 * FHL ^ 2 + 1.8982 * FHL + 2.8083
        End If
        '返回结果
        Return ans
    End Function

    Function 风冷螺杆式热泵制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        Dim ans As Double
        ans = -2.7857 * FHL ^ 2 + 3.1264 * FHL + 0.6593
        '返回结果
        Return ans
    End Function
    Function 空气源热泵制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        '暂时与风冷螺杆式热泵采用同样的系数曲线
        Dim ans As Double
        ans = -2.7857 * FHL ^ 2 + 3.1264 * FHL + 0.6593
        '返回结果
        Return ans
    End Function
    Function 离心式热泵制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        '与离心式冷水机暂时采用同样的曲线
        Dim ans As Double
        '负荷率小于等于0.5
        If FHL <= 0.5 Then
            ans = 31.385 * FHL ^ 3 - 35.663 * FHL ^ 2 + 15.017 * FHL - 0.3516
        Else
            ans = 5.373 * FHL ^ 3 - 9.0827 * FHL ^ 2 + 1.8982 * FHL + 2.8083
        End If
        '返回结果
        Return ans
    End Function
    Function 水_地源热泵制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        '与离心式冷水机暂时采用同样的曲线
        Dim ans As Double
        '负荷率小于等于0.5
        If FHL <= 0.5 Then
            ans = 31.385 * FHL ^ 3 - 35.663 * FHL ^ 2 + 15.017 * FHL - 0.3516
        Else
            ans = 5.373 * FHL ^ 3 - 9.0827 * FHL ^ 2 + 1.8982 * FHL + 2.8083
        End If
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂制冷COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制冷和蓄冷都用这条曲线
        Dim ans As Double
        ans = 10.159 * FHL ^ 5 - 47.032 * FHL ^ 4 + 85.765 * FHL ^ 3 - 76.941 * FHL ^ 2 + 33.982 * FHL - 4.9397
        '返回结果
        Return ans
    End Function
    Function 天然气采暖锅炉制热效率曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        Dim ans As Double
        ans = -5.2813 * FHL ^ 6 + 16.032 * FHL ^ 5 - 18.931 * FHL ^ 4 + 11.011 * FHL ^ 3 - 3.2783 * FHL ^ 2 + 0.6453 * FHL + 0.802
        '返回结果
        Return ans
    End Function
    Function 风冷螺杆式热泵制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        Dim ans As Double
        ans = -2.4176 * FHL ^ 2 + 2.6661 * FHL + 0.7515
        '返回结果
        Return ans
    End Function
    Function 空气源热泵制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        '暂时与风冷螺杆式热泵采用同样的系数曲线
        Dim ans As Double
        ans = -2.4176 * FHL ^ 2 + 2.6661 * FHL + 0.7515
        '返回结果
        Return ans
    End Function
    Function 离心式热泵制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        Dim ans As Double
        ans = 2.1282 * FHL ^ 3 - 6.7312 * FHL ^ 2 + 5.5523 * FHL + 0.0508
        '返回结果
        Return ans
    End Function
    Function 水_地源热泵制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        '采用与离心式热泵相同的特性曲线
        Dim ans As Double
        ans = 2.1282 * FHL ^ 3 - 6.7312 * FHL ^ 2 + 5.5523 * FHL + 0.0508
        '返回结果
        Return ans
    End Function
    Function 直燃型溴化锂制热COP曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        '制热和蓄热都用这条曲线
        '采用与天然气采暖锅炉相同的效率特性曲线
        Dim ans As Double
        ans = -5.2813 * FHL ^ 6 + 16.032 * FHL ^ 5 - 18.931 * FHL ^ 4 + 11.011 * FHL ^ 3 - 3.2783 * FHL ^ 2 + 0.6453 * FHL + 0.802
        '返回结果
        Return ans
    End Function
    Function 电采暖锅炉制热效率曲线(FHL As Double)
        On Error Resume Next
        '输出结果为小于等于1的小数，用于修正在不同负荷率情况下，设备的COP相对100%负荷率时的比例
        Dim ans As Double
        ans = 0.0198 * FHL + 0.9802
        '返回结果
        Return ans
    End Function
End Module
