Module 读取基础数据
    Function 读取制冷季装机方案参数(ExcelApp As Object)
        '内燃机
        '内燃机（1）装机数量
        Dim NUM1_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(5, 2).Value
        '内燃机（2）装机数量
        Dim NUM2_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(24, 2).Value
        '内燃机（1）100%负荷时发电功率（单台）
        Dim FDGL1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 4).Value
        '内燃机（2）100%负荷时发电功率（单台）
        Dim FDGL2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 4).Value
        '内燃机（1）100%负荷时余热回收功率（单台）
        Dim YRGL1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(11, 4).Value
        '内燃机（2）100%负荷时余热回收功率（单台）
        Dim YRGL2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(30, 4).Value
        '内燃机（1）100%负荷时天然气耗量（单台）
        Dim TRQ1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(12, 4).Value
        '内燃机（2）100%负荷时天然气耗量（单台）
        Dim TRQ2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(31, 4).Value
        '内燃机（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(13, 4).Value
        '内燃机（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(32, 4).Value
        '—————————————————————————————————————————————————————————————————————
        '离心式冷水机
        '设备（1）装机数量
        Dim NUM1_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 10).Value
        '设备（2）装机数量
        Dim NUM2_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 10).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 10).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 10).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 10).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 10).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 10).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_LXSLSJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 10).Value
        '—————————————————————————————————————————————————————————————————————
        '水冷螺杆机
        '设备（1）装机数量
        Dim NUM1_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 12).Value
        '设备（2）装机数量
        Dim NUM2_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 12).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 12).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 12).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 12).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 12).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 12).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_SLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 12).Value
        '—————————————————————————————————————————————————————————————————————
        '风冷螺杆机
        '设备（1）装机数量
        Dim NUM1_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 14).Value
        '设备（2）装机数量
        Dim NUM2_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 14).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 14).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 14).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 14).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 14).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 14).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 14).Value
        '—————————————————————————————————————————————————————————————————————
        '水地源热泵
        '设备（1）装机数量
        Dim NUM1_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 16).Value
        '设备（2）装机数量
        Dim NUM2_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 16).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 16).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 16).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 16).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 16).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 16).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 16).Value
        '—————————————————————————————————————————————————————————————————————
        '离心式热泵
        '设备（1）装机数量
        Dim NUM1_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 18).Value
        '设备（2）装机数量
        Dim NUM2_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 18).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 18).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 18).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 18).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 18).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 18).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 18).Value
        '—————————————————————————————————————————————————————————————————————
        '空气源热泵
        '设备（1）装机数量
        Dim NUM1_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 20).Value
        '设备（2）装机数量
        Dim NUM2_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 20).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 20).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 20).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 20).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 20).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 20).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 20).Value
        '—————————————————————————————————————————————————————————————————————
        '直燃型溴化锂
        '设备（1）装机数量
        Dim NUM1_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(4, 22).Value
        '设备（2）装机数量
        Dim NUM2_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(23, 22).Value
        '设备（1）装机功率（总和）
        Dim ZJLGL1_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(10, 22).Value
        '设备（2）装机功率（总和）
        Dim ZJLGL2_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(29, 22).Value
        '设备（1）100%负荷时本体耗气量（单台）
        Dim BTHQ1_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(6, 22).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(7, 22).Value
        '设备（2）100%负荷时本体耗气量（单台）
        Dim BTHQ2_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(25, 22).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(26, 22).Value
        '—————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(66)
        '内燃机
        ans(0) = NUM1_NRJ
        ans(1) = NUM2_NRJ
        ans(2) = FDGL1_ED_NRJ
        ans(3) = FDGL2_ED_NRJ
        ans(4) = YRGL1_ED_NRJ
        ans(5) = YRGL2_ED_NRJ
        ans(6) = TRQ1_ED_NRJ
        ans(7) = TRQ2_ED_NRJ
        ans(8) = FJHD1_ED_NRJ
        ans(9) = FJHD2_ED_NRJ
        '离心式冷水机
        ans(10) = NUM1_LXSLSJ
        ans(11) = NUM2_LXSLSJ
        ans(12) = ZJLGL1_LXSLSJ
        ans(13) = ZJLGL2_LXSLSJ
        ans(14) = BTHD1_ED_LXSLSJ
        ans(15) = BTHD2_ED_LXSLSJ
        ans(16) = FJHD1_ED_LXSLSJ
        ans(17) = FJHD2_ED_LXSLSJ
        '水冷螺杆机
        ans(18) = NUM1_SLLGJ
        ans(19) = NUM2_SLLGJ
        ans(20) = ZJLGL1_SLLGJ
        ans(21) = ZJLGL2_SLLGJ
        ans(22) = BTHD1_ED_SLLGJ
        ans(23) = BTHD2_ED_SLLGJ
        ans(24) = FJHD1_ED_SLLGJ
        ans(25) = FJHD2_ED_SLLGJ
        '风冷螺杆机
        ans(26) = NUM1_FLLGJ
        ans(27) = NUM2_FLLGJ
        ans(28) = ZJLGL1_FLLGJ
        ans(29) = ZJLGL2_FLLGJ
        ans(30) = BTHD1_ED_FLLGJ
        ans(31) = BTHD2_ED_FLLGJ
        ans(32) = FJHD1_ED_FLLGJ
        ans(33) = FJHD2_ED_FLLGJ
        '水地源热泵
        ans(34) = NUM1_SDYRB
        ans(35) = NUM2_SDYRB
        ans(36) = ZJLGL1_SDYRB
        ans(37) = ZJLGL2_SDYRB
        ans(38) = BTHD1_ED_SDYRB
        ans(39) = BTHD2_ED_SDYRB
        ans(40) = FJHD1_ED_SDYRB
        ans(41) = FJHD2_ED_SDYRB
        '离心式热泵
        ans(42) = NUM1_LXSRB
        ans(43) = NUM2_LXSRB
        ans(44) = ZJLGL1_LXSRB
        ans(45) = ZJLGL2_LXSRB
        ans(46) = BTHD1_ED_LXSRB
        ans(47) = BTHD2_ED_LXSRB
        ans(48) = FJHD1_ED_LXSRB
        ans(49) = FJHD2_ED_LXSRB
        '空气源热泵
        ans(50) = NUM1_KQYRB
        ans(51) = NUM2_KQYRB
        ans(52) = ZJLGL1_KQYRB
        ans(53) = ZJLGL2_KQYRB
        ans(54) = BTHD1_ED_KQYRB
        ans(55) = BTHD2_ED_KQYRB
        ans(56) = FJHD1_ED_KQYRB
        ans(57) = FJHD2_ED_KQYRB
        '直燃型溴化锂
        ans(58) = NUM1_ZRXXHL
        ans(59) = NUM2_ZRXXHL
        ans(60) = ZJLGL1_ZRXXHL
        ans(61) = ZJLGL2_ZRXXHL
        ans(62) = BTHQ1_ED_ZRXXHL
        ans(63) = BTHQ2_ED_ZRXXHL
        ans(64) = FJHD1_ED_ZRXXHL
        ans(65) = FJHD2_ED_ZRXXHL
        '返回结果
        Return ans
    End Function

    Function 读取采暖季装机方案参数(ExcelApp As Object)
        '内燃机
        '内燃机（1）装机数量
        Dim NUM1_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(50, 2).Value
        '内燃机（2）装机数量
        Dim NUM2_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(69, 2).Value
        '内燃机（1）100%负荷时发电功率（单台）
        Dim FDGL1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 4).Value
        '内燃机（2）100%负荷时发电功率（单台）
        Dim FDGL2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 4).Value
        '内燃机（1）100%负荷时余热回收功率（单台）
        Dim YRGL1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(56, 4).Value
        '内燃机（2）100%负荷时余热回收功率（单台）
        Dim YRGL2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(75, 4).Value
        '内燃机（1）100%负荷时天然气耗量（单台）
        Dim TRQ1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(57, 4).Value
        '内燃机（2）100%负荷时天然气耗量（单台）
        Dim TRQ2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(76, 4).Value
        '内燃机（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(58, 4).Value
        '内燃机（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(77, 4).Value
        '—————————————————————————————————————————————————————————————————————
        '天然气锅炉
        '设备（1）装机数量
        Dim NUM1_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 10).Value
        '设备（2）装机数量
        Dim NUM2_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 10).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 10).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 10).Value
        '设备（1）100%负荷时本体耗气量（单台）
        Dim BTHQ1_ED_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 10).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 10).Value
        '设备（2）100%负荷时本体耗气量（单台）
        Dim BTHQ2_ED_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 10).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_TRQGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 10).Value
        '—————————————————————————————————————————————————————————————————————
        '风冷螺杆机
        '设备（1）装机数量
        Dim NUM1_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 12).Value
        '设备（2）装机数量
        Dim NUM2_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 12).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 12).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 12).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 12).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 12).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 12).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_FLLGJ As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 12).Value
        '—————————————————————————————————————————————————————————————————————
        '水地源热泵
        '设备（1）装机数量
        Dim NUM1_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 14).Value
        '设备（2）装机数量
        Dim NUM2_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 14).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 14).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 14).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 14).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 14).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 14).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_SDYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 14).Value
        '—————————————————————————————————————————————————————————————————————
        '空气源热泵
        '设备（1）装机数量
        Dim NUM1_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 16).Value
        '设备（2）装机数量
        Dim NUM2_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 16).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 16).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 16).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 16).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 16).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 16).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_KQYRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 16).Value
        '—————————————————————————————————————————————————————————————————————
        '离心式热泵
        '设备（1）装机数量
        Dim NUM1_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 18).Value
        '设备（2）装机数量
        Dim NUM2_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 18).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 18).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 18).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 18).Value
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 18).Value
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 18).Value
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_LXSRB As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 18).Value
        '—————————————————————————————————————————————————————————————————————
        '直燃型溴化锂
        '设备（1）装机数量
        Dim NUM1_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 20).Value
        '设备（2）装机数量
        Dim NUM2_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 20).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(55, 20).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(74, 20).Value
        '设备（1）100%负荷时本体耗气量
        Dim BTHQ1_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(51, 20).Value
        '设备（1）100%负荷时辅机耗电功率
        Dim FJHD1_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 20).Value
        '设备（2）100%负荷时本体耗气量
        Dim BTHQ2_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(70, 20).Value
        '设备（2）100%负荷时辅机耗电功率
        Dim FJHD2_ED_ZRXXHL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 20).Value
        '—————————————————————————————————————————————————————————————————————
        '电锅炉
        '设备（1）装机数量
        Dim NUM1_DGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(49, 22).Value
        '设备（2）装机数量
        Dim NUM2_DGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(68, 22).Value
        '设备（1）装机功率（总和）
        Dim ZJRGL1_DGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 22).Value
        '设备（2）装机功率（总和）
        Dim ZJRGL2_DGL As Double = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 22).Value
        '设备（1）100%负荷时本体耗电功率（单台）
        Dim BTHD1_ED_DGL As Double
        If NUM1_DGL = 0 Then
            BTHD1_ED_DGL = 0
        Else
            BTHD1_ED_DGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(50, 22).Value / NUM1_DGL
        End If
        '设备（1）100%负荷时辅机耗电功率（单台）
        Dim FJHD1_ED_DGL As Double
        If NUM1_DGL = 0 Then
            FJHD1_ED_DGL = 0
        Else
            FJHD1_ED_DGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(52, 22).Value / NUM1_DGL
        End If
        '设备（2）100%负荷时本体耗电功率（单台）
        Dim BTHD2_ED_DGL As Double
        If NUM2_DGL = 0 Then
            BTHD2_ED_DGL = 0
        Else
            BTHD2_ED_DGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(69, 22).Value / NUM2_DGL
        End If
        '设备（2）100%负荷时辅机耗电功率（单台）
        Dim FJHD2_ED_DGL As Double
        If NUM2_DGL = 0 Then
            FJHD2_ED_DGL = 0
        Else
            FJHD2_ED_DGL = ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(71, 22).Value / NUM2_DGL
        End If
        '—————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(66)
        '内燃机
        ans(0) = NUM1_NRJ
        ans(1) = NUM2_NRJ
        ans(2) = FDGL1_ED_NRJ
        ans(3) = FDGL2_ED_NRJ
        ans(4) = YRGL1_ED_NRJ
        ans(5) = YRGL2_ED_NRJ
        ans(6) = TRQ1_ED_NRJ
        ans(7) = TRQ2_ED_NRJ
        ans(8) = FJHD1_ED_NRJ
        ans(9) = FJHD2_ED_NRJ
        '天然气锅炉
        ans(10) = NUM1_TRQGL
        ans(11) = NUM2_TRQGL
        ans(12) = ZJRGL1_TRQGL
        ans(13) = ZJRGL2_TRQGL
        ans(14) = BTHQ1_ED_TRQGL
        ans(15) = BTHQ2_ED_TRQGL
        ans(16) = FJHD1_ED_TRQGL
        ans(17) = FJHD2_ED_TRQGL
        '电锅炉
        ans(18) = NUM1_DGL
        ans(19) = NUM2_DGL
        ans(20) = ZJRGL1_DGL
        ans(21) = ZJRGL2_DGL
        ans(22) = BTHD1_ED_DGL
        ans(23) = BTHD2_ED_DGL
        ans(24) = FJHD1_ED_DGL
        ans(25) = FJHD2_ED_DGL
        '风冷螺杆机
        ans(26) = NUM1_FLLGJ
        ans(27) = NUM2_FLLGJ
        ans(28) = ZJRGL1_FLLGJ
        ans(29) = ZJRGL2_FLLGJ
        ans(30) = BTHD1_ED_FLLGJ
        ans(31) = BTHD2_ED_FLLGJ
        ans(32) = FJHD1_ED_FLLGJ
        ans(33) = FJHD2_ED_FLLGJ
        '水地源热泵
        ans(34) = NUM1_SDYRB
        ans(35) = NUM2_SDYRB
        ans(36) = ZJRGL1_SDYRB
        ans(37) = ZJRGL2_SDYRB
        ans(38) = BTHD1_ED_SDYRB
        ans(39) = BTHD2_ED_SDYRB
        ans(40) = FJHD1_ED_SDYRB
        ans(41) = FJHD2_ED_SDYRB
        '离心式热泵
        ans(42) = NUM1_LXSRB
        ans(43) = NUM2_LXSRB
        ans(44) = ZJRGL1_LXSRB
        ans(45) = ZJRGL2_LXSRB
        ans(46) = BTHD1_ED_LXSRB
        ans(47) = BTHD2_ED_LXSRB
        ans(48) = FJHD1_ED_LXSRB
        ans(49) = FJHD2_ED_LXSRB
        '空气源热泵
        ans(50) = NUM1_KQYRB
        ans(51) = NUM2_KQYRB
        ans(52) = ZJRGL1_KQYRB
        ans(53) = ZJRGL2_KQYRB
        ans(54) = BTHD1_ED_KQYRB
        ans(55) = BTHD2_ED_KQYRB
        ans(56) = FJHD1_ED_KQYRB
        ans(57) = FJHD2_ED_KQYRB
        '直燃型溴化锂
        ans(58) = NUM1_ZRXXHL
        ans(59) = NUM2_ZRXXHL
        ans(60) = ZJRGL1_ZRXXHL
        ans(61) = ZJRGL2_ZRXXHL
        ans(62) = BTHQ1_ED_ZRXXHL
        ans(63) = BTHQ2_ED_ZRXXHL
        ans(64) = FJHD1_ED_ZRXXHL
        ans(65) = FJHD2_ED_ZRXXHL
        '返回结果
        Return ans
    End Function
    Function 读取制冷季输入的计算系数(ExcelApp As Object)
        '与水换热的设备供冷时本体耗电修正系数
        Dim BTHDXS_GL_water As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(5, 7).Value
        '与水换热的设备蓄冷时本体耗电修正系数
        Dim BTHDXS_XL_water As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(9, 7).Value
        '与空气换热的设备供冷时本体耗电修正系数
        Dim BTHDXS_GL_air As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(6, 7).Value
        '与空气换热的设备蓄冷时本体耗电修正系数
        Dim BTHDXS_XL_air As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(10, 7).Value
        '辅机耗电修正系数
        Dim FJHDXS As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(14, 7).Value
        '直燃型溴化锂制冷耗气量修正系数
        Dim ZLHQXZ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(13, 7).Value
        '除内燃机以外的其它设备内燃机耗量修正系数
        Dim TRQHLXZXS_QT As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(16, 9).Value + 1
        '内燃机天然气耗量修正系数
        Dim TRQHLXZXS_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(15, 9).Value + 1
        '内燃机发电效率修正系数
        Dim FDXLXZ_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(17, 7).Value
        '内燃机余热效率修正系数
        Dim YRXLXZ_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(18, 7).Value
        '溴化锂制冷COP
        Dim XHL_COP_L As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(19, 7).Value
        '返回结果
        Dim ans(11)
        ans(0) = BTHDXS_GL_water
        ans(1) = BTHDXS_XL_water
        ans(2) = BTHDXS_GL_air
        ans(3) = BTHDXS_XL_air
        ans(4) = FJHDXS
        ans(5) = ZLHQXZ
        ans(6) = TRQHLXZXS_QT
        ans(7) = TRQHLXZXS_NRJ
        ans(8) = FDXLXZ_NRJ
        ans(9) = YRXLXZ_NRJ
        ans(10) = XHL_COP_L
        '返回结果
        Return ans
    End Function
    Function 读取采暖季输入的计算系数(ExcelApp As Object)
        '与水换热的设备供热时本体耗电修正系数
        Dim BTHDXS_GR_water As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(7, 7).Value
        '与水换热的设备蓄热时本体耗电修正系数
        Dim BTHDXS_XR_water As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(11, 7).Value
        '与空气换热的设备供热时本体耗电修正系数
        Dim BTHDXS_GR_air As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(8, 7).Value
        '与空气换热的设备蓄热时本体耗电修正系数
        Dim BTHDXS_XR_air As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(12, 7).Value
        '辅机耗电修正系数
        Dim FJHDXS As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(14, 7).Value
        '除内燃机以外的其它设备内燃机耗量修正系数
        Dim TRQHLXZXS_QT As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(16, 9).Value + 1
        '内燃机天然气耗量修正系数
        Dim TRQHLXZXS_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(15, 9).Value + 1
        '内燃机发电效率修正系数
        Dim FDXLXZ_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(17, 7).Value
        '内燃机余热效率修正系数
        Dim YRXLXZ_NRJ As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(18, 7).Value
        '溴化锂制热COP
        Dim XHL_COP_R As Double = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(20, 7).Value
        '返回结果
        Dim ans(10)
        ans(0) = BTHDXS_GR_water
        ans(1) = BTHDXS_XR_water
        ans(2) = BTHDXS_GR_air
        ans(3) = BTHDXS_XR_air
        ans(4) = FJHDXS
        ans(5) = TRQHLXZXS_QT
        ans(6) = TRQHLXZXS_NRJ
        ans(7) = FDXLXZ_NRJ
        ans(8) = YRXLXZ_NRJ
        ans(9) = XHL_COP_R
        '返回结果
        Return ans
    End Function
    Function 读取输入的各种数据并添加报错功能(ExcelApp As Object, FHTJJD As Double, n As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '状态检测
        Dim ZTJC_SHUJU As Integer
        Dim NRJYRLYFS(10000) As String '内燃机余热利用方式
        Dim XSS(10000) As Double '输入的工况时间(天数*小时数*时间频数)
        For a = 1 To n
            '各种负荷量
            LFHZXQL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 12).Value
            RFHZXQL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 21).Value
            XNGLGL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 13).Value
            XNXLGL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 14).Value
            XNGRGL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 22).Value
            XNXRGL(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 23).Value
            NRJYRLYFS(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value
            '输入的工况时间(天数*小时数*时间频数)
            XSS(a) = ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 24).Value * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 25).Value * ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 26).Value
            '添加一些报错功能
            '如果小时数为0，报错
            If XSS(a) = 0 Then
                MsgBox("本工况（负荷段全年天数*负荷段每天小时数*负荷时间频数）的乘积不可以为0，请重新输入！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '如果某工况的用电时间段没有选择，为空，则报错
            If ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 79).Value = Nothing Then
                MsgBox("本工况用电时间段不能为空，请重新选择！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '如果某工况输入的蓄能设备供冷/供热量大于本工况冷/热负荷总需求量，报错
            If XNGLGL(a) > LFHZXQL(a) * (1 + 4 * FHTJJD / 100) Then '允许有2%的误差
                MsgBox("蓄能装置供冷量(kW)不可以大于本工况冷负荷总需求量(kW)，请重新输入！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            If XNGRGL(a) > RFHZXQL(a) * (1 + 4 * FHTJJD / 100) Then '允许有2%的误差
                MsgBox("蓄能装置供热量(kW)不可以大于本工况热负荷总需求量(kW)，请重新输入！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '如果使用内燃机余热制备生活热水或者工业蒸汽，必需使用内燃机可以向外供电的模式计算，手动输入内燃机负荷率，加入相关报错功能
            If (((ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(3, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value <> "溴化锂") Or (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(22, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value <> "溴化锂") Or (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(48, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value <> "溴化锂") Or (ExcelApp.ThisWorkbook.Worksheets("设备选型&负荷分析计算").Cells(67, 2).Value <> "J000GS" And ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 77).Value <> "溴化锂")) And ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value <> "Y") Then
                '报错
                MsgBox("当内燃机余热用于制备生活热水或者工业蒸汽时，必需使用内燃机可以向外供电的模式进行计算！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '如果处于内燃机可以向外供电的模式，检测用户输入的内燃机负荷率，如果负荷率大于1或者小于0，报错
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(27, 9).Value = "Y" Then '处于内燃机可以向外供电的模式
                If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 2).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 2).Value < 0) Then '供冷时间段内燃机(1)
                    MsgBox("供冷时间段内燃机(1)负荷率不可以大于1或者小于0，请重新输入！")
                    ZTJC_SHUJU = 1
                    Call 锁定工作表(ExcelApp)
                    GoTo qqq
                End If
                If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 3).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 3).Value < 0) Then '供冷时间段内燃机(2)
                    MsgBox("供冷时间段内燃机(2)负荷率不可以大于1或者小于0，请重新输入！")
                    ZTJC_SHUJU = 1
                    Call 锁定工作表(ExcelApp)
                    GoTo qqq
                End If
                If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 4).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 4).Value < 0) Then '供热时间段内燃机(1)
                    MsgBox("供热时间段内燃机(1)负荷率不可以大于1或者小于0，请重新输入！")
                    ZTJC_SHUJU = 1
                    Call 锁定工作表(ExcelApp)
                    GoTo qqq
                End If
                If (ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 5).Value > 1 Or ExcelApp.ThisWorkbook.Worksheets("计算输入").Cells(7 + a, 5).Value < 0) Then '供热时间段内燃机(2)
                    MsgBox("供热时间段内燃机(2)负荷率不可以大于1或者小于0，请重新输入！")
                    Call 锁定工作表(ExcelApp)
                    GoTo qqq
                End If
            End If
            '每个计算工况不可以同时计算供冷和供热，如果一个工况不仅有冷负荷需求，还有热负荷需求，报错
            If (LFHZXQL(a) > 0 And RFHZXQL(a) > 0) Then
                MsgBox("同一个工况不可以同时计算供冷和供热，必需分开计算！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '如果有蓄冷蓄热负荷率，判断蓄能装置开关是否开启，如没有开启，则报错
            If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And XNXLGL(a) > 0) Or (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And XNGLGL(a) > 0) Then
                MsgBox("本项目有蓄冷负荷需求，请打开蓄能装置计算开关！！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            If (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And XNXRGL(a) > 0) Or (ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(23, 9).Value <> "Y" And XNGLGL(a) > 0) Then
                MsgBox("本项目有蓄热负荷需求，请打开蓄能装置计算开关！！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '针对内燃机余热利用方式添加报错功能
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(17, 9).Value = "N" And NRJYRLYFS(a) = "溴化锂" Then
                MsgBox("本项目内燃机余热利用方式有溴化锂，请打开对应的计算开关！！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(18, 9).Value = "N" And NRJYRLYFS(a) = "生活热水" Then
                MsgBox("本项目内燃机余热利用方式有生活热水，请打开对应的计算开关！！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(19, 9).Value = "N" And NRJYRLYFS(a) = "工业蒸汽" Then
                MsgBox("本项目内燃机余热利用方式有工业蒸汽，请打开对应的计算开关！！")
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
            '判断每个工况的冷热负荷总需求量和蓄冷蓄热负荷需求量之和是否大于所有设备总制冷制热功率，如果大于，则报错。
            Dim ZTJC_LRFH As Integer = 判断冷热负荷需求量是否大于冷热负荷装机量(ExcelApp, a)(0)
            If ZTJC_LRFH = 1 Then
                ZTJC_SHUJU = 1
                Call 锁定工作表(ExcelApp)
                GoTo qqq
            End If
        Next
qqq:
        '返回结果
        Return ZTJC_SHUJU
    End Function


End Module
