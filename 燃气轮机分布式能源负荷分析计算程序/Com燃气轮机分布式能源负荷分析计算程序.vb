Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
<ComClass(Com燃气轮机分布式能源负荷分析计算程序.ClassId, Com燃气轮机分布式能源负荷分析计算程序.InterfaceId, Com燃气轮机分布式能源负荷分析计算程序.EventsId)>
Public Class Com燃气轮机分布式能源负荷分析计算程序
#Region "COM GUID"
    ' 这些 GUID 提供此类的 COM 标识 
    ' 及其 COM 接口。若更改它们，则现有的
    ' 客户端将不再能访问此类。
    Public Const ClassId As String = "bda7078d-970a-4c34-9496-a767e7ddb375"
    Public Const InterfaceId As String = "c318f9e6-7c79-4e10-97d1-bc179040d3a6"
    Public Const EventsId As String = "0ebdb9c2-8e50-464f-abfd-61e329281c40"
#End Region
    ' 可创建的 COM 类必须具有一个不带参数的 Public Sub New() 
    ' 否则， 将不会在 
    ' COM 注册表中注册此类，且无法通过
    ' CreateObject 创建此类。
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub MainCaculationProgram()
        '本SUB为主程序
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 解锁工作表()
        '将热电分析结果工作表设置为不可以显示
        ExcelApp.ThisWorkbook.Worksheets("热电分析结果").Visible = False
        '清空“变工况结果输出表中已有的数据”
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range("B7:AI306").ClearContents
        '情况变工况计算输入的过程量
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("S7:AL306").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("B3:AL3").ClearContents
        '————————————————————————————————————————————————————————————————————————————————————————
        Call Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定工作表()
            ZTJC = 0
            Exit Sub
        End If
        '计数，统计一共有多少种不同工况
        For i = 110 To 7 Step -1 '行号，从大到小查找
            If ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(i, 16).Value > 0 Then
                n = i - 6 '工况总数
                Exit For '跳出循环
            End If
        Next
        '读取用于计算的燃气轮机（1）（2）基准发电功率(kW)（环境温度20℃，100%负荷），并记录行号
        Call 读取燃气轮机基准发电功率()
        If n > 0 Then
            '将各个工况的温度记录在数组内
            For j = 1 To n
                TEMPArray(j) = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 18).Value
                '针对环境温度添加报错功能
                If (TEMPArray(j) < -20 Or TEMPArray(j) > 40) Then
                    MsgBox("环境温度不可以大于40℃或者低于-20℃")
                    Call 锁定工作表()
                    Exit Sub
                End If
            Next
            '通过输入的抽背蒸汽轮机背压蒸汽需求量，倒算此时需要的背压联合循环燃气轮机负荷率
            Call 通过背压蒸汽需求量反算燃气轮机负荷率()
            '将各个工况的燃机(1)(2)负荷率记录在数组内
            For j = 1 To n
                RJFHL1Array(j) = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 2).Value
                RJFHL2Array(j) = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 3).Value
                '针对燃气轮机负荷率添加报错功能
                If (RJFHL1Array(j) > 1) Then
                    MsgBox("燃气轮机（1）负荷率不可以大于100%")
                    Call 锁定工作表()
                    Exit Sub
                End If
                If (RJFHL1Array(j) < 0) Then
                    MsgBox("燃气轮机（1）负荷率不可以小于0")
                    Call 锁定工作表()
                    Exit Sub
                End If
                If (RJFHL1Array(j) > 0 And RJFHL1Array(j) < 0.4) Then
                    MsgBox("燃气轮机（1）负荷率不可以小于40%")
                    Call 锁定工作表()
                    Exit Sub
                End If
                If (RJFHL2Array(j) > 1) Then
                    MsgBox("燃气轮机（2）负荷率不可以大于100%")
                    Call 锁定工作表()
                    Exit Sub
                End If
                If (RJFHL2Array(j) < 0) Then
                    MsgBox("燃气轮机（2）负荷率不可以小于0")
                    Call 锁定工作表()
                    Exit Sub
                End If
                If (RJFHL2Array(j) > 0 And RJFHL2Array(j) < 0.4) Then
                    MsgBox("燃气轮机（2）负荷率不可以小于40%")
                    Call 锁定工作表()
                    Exit Sub
                End If
            Next
            '计算不同工况温度下的最大100%发电功率（kW)
            '燃气轮机(1)(2)不同工况温度下的最大100%发电功率（kW)
            Call 计算不同环境温度下燃气轮机最大发电功率()
            '计算燃机汽轮（1）（2）计算用基准发电效率（20℃时），并记录在数组中
            '先用燃气轮机负荷率修正，再用环境温度修正
            Call 燃机汽轮计算用基准发电效率()
            '计算余热锅炉效率
            Call 计算余热锅炉效率()
            '计算蒸汽轮机蒸汽循环总效率
            Call 计算蒸汽轮机蒸汽循环总效率()
            '计算主蒸汽焓值修正系数
            Call 计算主蒸汽焓值修正系数()
            '抽背蒸汽轮机背压蒸汽焓值（kj/kg）修正系数
            Call 计算抽背蒸汽轮机背压蒸汽焓值修正系数()
            '抽凝蒸汽轮机抽汽焓值（kj/kg）修正系数，修正系数求取的数据库与背压焓值修正系数一致
            Call 计算抽凝蒸汽轮机或抽背蒸汽轮机抽汽焓值修正系数()
            '抽凝机进入凝汽器的排汽焓值（kj/kg）
            Call 计算抽凝蒸汽轮机排汽焓值()
            '抽背蒸汽轮机计算用基准背压焓值（kj/kg）
            Call 计算抽背蒸汽轮机基准背压蒸汽焓值()
            '变工况计算循环体
            Call 变工况计算循环体()
        End If
        '——————————————————————————————————————————————————————————————————————————————
        '判断天然气耗量计算是否正确，如果出现了不正确则报错
        For i = 12 To 18
            If ExcelApp.ThisWorkbook.Worksheets("成本测算").Cells(i, 2).Value = "不正确" Then
                MsgBox("天然气耗量计算结果中出现了误差较大的不正确结果，请检查！！")
                Exit For
            End If
        Next
        Call 锁定工作表()
    End Sub
    Public Sub 热电分析计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 解锁工作表()
        '清空"变工况计算输入"表和"变工况计算结果输出"表中已经输入的数据
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("B7:AL306").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range("B7:AI306").ClearContents
        '————————————————————————————————————————————————————————————————————————————————————————
        Call Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定工作表()
            ZTJC = 0
            Exit Sub
        End If
        '定义局部变量
        Dim qq, a, w1, b
        '将燃气轮机（1）的负荷率重置为1，进行一次计算
        '常量赋值
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 2).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 16).Value = 2000
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 17).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 18).Value = 20
        '读取环境温度20℃时的最大燃机发电功率kW，发电效率
        '燃气轮机(1)发电效率和最大发电功率
        For q = 3 To 49
            If ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(q, 1).Value = ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(3, 2).Value Then
                '记录下行号
                qq = q
                '最大燃机发电功率
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 19).Value = ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(qq, 5).Value
                '燃机发电效率
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 21).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(qq, 6).Value / 100, 5)
            End If
        Next
        '燃气轮机(2)参数均设置为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 3).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 20).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 22).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 24).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 26), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 28)).Value = 0
        '余热锅炉(1)热效率——抽凝机
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 23).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(qq, 4).Value / 100, 5)
        '汽轮机(1)蒸汽循环总效率——抽凝机
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 25).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(qq, 4).Value / 100, 5)
        '设置蒸汽焓值修正系数
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 29).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 30), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 32)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 33).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 34).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 35).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(qq, 76).Value, 2)
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 36), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 38)).Value = 0
        '蒸汽量均设置为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 4), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 15)).Value = 0
        '进行一次变工况计算
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 38)).Value
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(7, 35)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 35)).Value
        '读取抽凝机最大进汽量，燃气轮机（1）满负荷发电功率(kW)
        a = ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(8, 4).Value
        w1 = ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(7, 4).Value
        '燃气轮机保持100%负荷运行，改变抽凝机抽汽量，比较发电和供热经济性,不计算抽背蒸汽轮机
        '进行"抽凝机"变抽汽量热电分析计算
        '部分定值赋值，燃气轮机（2）全部变量为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 2)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 3), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 3)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 5), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 11)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 12), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 15)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 16), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 16)).Value = 2000
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 17), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 17)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 18), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 18)).Value = 20
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 19), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 19)).Value = ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(qq, 5).Value
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 20), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 20)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 21), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 21)).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(qq, 6).Value / 100, 5)
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 22), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 22)).Value = 0
        '读取余热锅炉（1）——抽凝机热效率，并将余热锅炉（2）热效率设置为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 23), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 23)).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(qq, 4).Value / 100, 5)
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 24), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 24)).Value = 0
        '读取抽凝机（1）蒸汽循环总效率，并将其他蒸汽循环总效率全部设置为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 25), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 25)).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(qq, 4).Value / 100, 5)
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 26), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 28)).Value = 0
        '抽凝机（1）排汽焓值修正系数设置为1，汽轮机（1）主蒸汽焓值修正系数设置为1，其它修正系数全部设置为0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 29), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 29)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 30), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 32)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 33), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 33)).Value = 1
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 34), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 34)).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 36), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 38)).Value = 0
        '抽凝机（1）凝汽器排汽焓值（kj/kg）
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 35), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 35)).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(qq, 76).Value, 2)
        '燃气轮机(1)负荷率保持不变，抽凝机抽汽量从2%变化到80%，每次变化2%
        For i = 1 To 40
            b = 0.02 * i
            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + i, 4), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(46, 4)).Value = ExcelApp.WorksheetFunction.RoundUp(a * b, 4)
            '读取变工况计算输入量
            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + i, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + i, 38)).Value
            '返回变工况计算结果
            ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(6 + i, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(6 + i, 35)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 35)).Value
        Next
        '显示表格
        ExcelApp.ThisWorkbook.Worksheets("热电分析结果").Visible = True
    End Sub
    Public Sub 清空全部数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 解锁工作表()
        Call Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定工作表()
            ZTJC = 0
            Exit Sub
        End If
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("B7:AL306").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range("B7:AI306").ClearContents
        '将所有变量赋值为0，进行一次计算
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 38)).Value = 0
        '读取变工况计算输入量
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range（ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(7, 38)).Value
        '返回变工况计算结果
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(7, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(7, 35)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 35)).Value
        '再清空一次全部数据
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("B7:AL306").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range("B7:AI306").ClearContents
        '将热电分析结果工作表设置为可以不可以显示
        ExcelApp.ThisWorkbook.Worksheets("热电分析结果").Visible = False
        Call 锁定工作表()
    End Sub
    Public Sub 打开表格自动运行()
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
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(3, 20).Value < 20181126 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
        End If
        Call 锁定工作表()
        Call 自保护程序()
    End Sub
    Public Sub Excel版本号验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(3, 20).Value < 20181126 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            ZTJC = 1
        End If
        Call 自保护程序()
    End Sub
    Public Sub 读取燃气轮机基准发电功率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取用于计算的燃气轮机（1）（2）基准发电功率(kW)（环境温度20℃，100%负荷），记录下行号
        For k = 3 To 49
            If ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(k, 1).Value = ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(3, 2).Value Then
                '记录下行号
                kk = k
                rj1kw = ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(k, 5).Value
            End If
        Next
        For l = 3 To 49
            If ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(l, 1).Value = ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(25, 2).Value Then
                '记录下行号
                ll = l
                rj2kw = ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(l, 5).Value
            End If
        Next
    End Sub
    Public Sub 计算不同环境温度下燃气轮机最大发电功率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '插值法计算不同工况温度下的最大100%发电功率（kW)
        '燃气轮机(1)(2)不同工况温度下的最大100%发电功率（kW)
        For a = 1 To n
            For b = 16 To 21
                If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                    '燃气轮机(1)不同温度下实际最大发电功率（kW)
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 19).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value) * rj1kw, 2)
                    '燃气轮机(2)不同温度下实际最大发电功率（kW)
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 20).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value) * rj2kw, 2)
                ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                    XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                    Exit For
                End If
            Next
        Next
    End Sub
    Public Sub 燃机汽轮计算用基准发电效率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '先用燃气轮机负荷率修正，再用环境温度修正
        '插值法计算不同负荷率燃机汽轮（1）计算用基准发电效率（20℃时），并记录在数组中
        '定义不同负荷率情况下，燃气轮机（1）（2）计算用基准发电效率
        Dim RJFDXLJZ1Array(10000) As Single
        Dim RJFDXLJZ2Array(10000) As Single
        For a = 1 To n
            For b = 6 To 12
                If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                    '燃气轮机（1）计算用基准发电效率（20℃时）数组
                    RJFDXLJZ1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value) / 100, 5)
                ElseIf (RJFHL1Array(a) = 0) Then
                    RJFDXLJZ1Array(a) = 0
                ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                    XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                    Exit For
                ElseIf (RJFHL1Array(a) > 1) Then
                    XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                    Exit For
                End If
            Next
        Next
        '插值法计算不同负荷率燃机汽轮（2）计算用基准发电效率（20℃时），并记录在数组中
        For a = 1 To n
            For b = 6 To 12
                If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                    '燃气轮机（2）计算用基准发电效率（20℃时）数组
                    RJFDXLJZ2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value) / 100, 5)
                ElseIf (RJFHL2Array(a) = 0) Then
                    RJFDXLJZ2Array(a) = 0
                ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                    XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                    Exit For
                ElseIf (RJFHL2Array(a) > 1) Then
                    XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                    Exit For
                End If
            Next
        Next
        '用环境温度修正
        '插值法计算不同环境温度下不同负荷率的燃气轮机（1）（2）发电效率
        For a = 1 To n
            For b = 26 To 31
                If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                    '燃气轮机(1)不同温度下不同负荷率的发电效率
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 21).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(kk, b).Value) * RJFDXLJZ1Array(a), 5)
                    '燃气轮机(2)不同温度下不同负荷率的发电效率
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 22).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Cells(ll, b).Value) * RJFDXLJZ2Array(a), 5)
                ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                    XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                    Exit For
                End If
            Next
        Next
    End Sub
    Public Sub 计算余热锅炉效率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算余热锅炉效率
        '判断燃气轮机(1)(2)接的是何种蒸汽轮机，选择不同的余热锅炉效率
        '余热锅炉(1)
        '不同燃气轮机负荷率情况下，插值法计算余热锅炉基准热效率（环境温度20℃）,记录在数组中
        '抽凝联合循环和背压联合循环分别记录
        Dim CNYRGLJZXL1Array(10000) As Single
        Dim BYYRGLJZXL1Array(10000) As Single
        Dim CNYRGLJZXL2Array(10000) As Single
        Dim BYYRGLJZXL2Array(10000) As Single
        '余热锅炉(1)-抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 4 To 10
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        CNYRGLJZXL1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value) / 100, 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        CNYRGLJZXL1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉(1)-抽背机或者简单循环
        If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "无") Then
            For a = 1 To n
                For b = 29 To 35
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        BYYRGLJZXL1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value) / 100, 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        BYYRGLJZXL1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉(2)-抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 4 To 10
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        CNYRGLJZXL2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value) / 100, 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        CNYRGLJZXL2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉(2)-抽背机或者简单循环
        If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "无") Then
            For a = 1 To n
                For b = 29 To 35
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        BYYRGLJZXL2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value) / 100, 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        BYYRGLJZXL2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉（1）-抽凝机，效率环境温度修正，插值法计算不同环境温度下余热锅炉的实际热效率
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 16 To 21
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 23).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value) * CNYRGLJZXL1Array(a), 5)
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉（1）-抽背机或者简单循环，效率环境温度修正，插值法计算不同环境温度下余热锅炉的实际热效率
        If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "无") Then
            For a = 1 To n
                For b = 41 To 46
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 23).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(kk, b).Value) * BYYRGLJZXL1Array(a), 5)
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉（2）-抽凝机，效率环境温度修正，插值法计算不同环境温度下余热锅炉的实际热效率
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 16 To 21
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 24).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value) * CNYRGLJZXL2Array(a), 5)
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '余热锅炉（2）-抽背机或者简单循环，效率环境温度修正，插值法计算不同环境温度下余热锅炉的实际热效率
        If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "无") Then
            For a = 1 To n
                For b = 41 To 46
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 24).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Cells(ll, b).Value) * BYYRGLJZXL2Array(a), 5)
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub 计算蒸汽轮机蒸汽循环总效率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算蒸汽轮机蒸汽循环总效率
        '判断燃气轮机(1)(2)接的是何种蒸汽轮机，选择不同的蒸汽轮机蒸汽循环总效率
        '蒸汽轮机(1)
        '不同燃气轮机负荷率情况下，插值法计算蒸汽轮机蒸汽循环基准总效率（环境温度20℃）,记录在数组中
        '抽凝联合循环和背压联合循环分别记录
        '         '读取用户输入的蒸汽轮机相对内效率，作为计算基准
        '
        '         CNXDNXL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(5, 7).Value
        '
        '         BYXDNXL = ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Cells(6, 7).Value
        '计算不同燃气轮机负荷率情况下的蒸汽轮机蒸汽循环总效率
        '定义数组用于储存计算结果，燃气轮机(1)(2)和抽凝机、抽背机分开储存
        Dim CNZQXHZXL1Array(10000) As Single
        Dim CNZQXHZXL2Array(10000) As Single
        Dim BYZQXHZXL1Array(10000) As Single
        Dim BYZQXHZXL2Array(10000) As Single
        '不同燃气轮机负荷率情况下，蒸汽轮机(1)蒸汽循环总效率——抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 4 To 10
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        CNZQXHZXL1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value) / 100, 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        CNZQXHZXL1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，蒸汽轮机(1)蒸汽循环总效率——抽背机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 29 To 35
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        BYZQXHZXL1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value) / 100, 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        BYZQXHZXL1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，蒸汽轮机(2)蒸汽循环总效率——抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 4 To 10
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        CNZQXHZXL2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value) / 100, 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        CNZQXHZXL2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，蒸汽轮机(2)蒸汽循环总效率——抽背机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 29 To 35
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        BYZQXHZXL2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value) / 100, 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        BYZQXHZXL2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL2Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正蒸汽轮机（1）蒸汽循环总效率——抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 16 To 21
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 25).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value) * CNZQXHZXL1Array(a), 5)
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 27).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正蒸汽轮机（1）蒸汽循环总效率——抽背机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 41 To 46
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 27).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(kk, b).Value) * BYZQXHZXL1Array(a), 5)
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 25).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正蒸汽轮机（2）蒸汽循环总效率——抽凝机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 16 To 21
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 26).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value) * CNZQXHZXL2Array(a), 5)
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 28).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正蒸汽轮机（2）蒸汽循环总效率——抽背机
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 41 To 46
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 28).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Cells(ll, b).Value) * BYZQXHZXL2Array(a), 5)
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 26).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '联合循环（1）为简单循环时，将蒸汽轮机（1）蒸汽循环总效率定义为0
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "无" Then
            For a = 1 To n
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 25).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 27).Value = 0
            Next
        End If
        '联合循环（2）为简单循环时，将蒸汽轮机（2）蒸汽循环总效率定义为0
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "无" Then
            For b = 1 To n
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + b, 26).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + b, 28).Value = 0
            Next
        End If
    End Sub
    Public Sub 计算主蒸汽焓值修正系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算主蒸汽焓值修正系数
        '汽轮机（1）主蒸汽焓值（kj/kg）修正系数
        '定义数组用于储存主蒸汽焓值修正系数计算过程量
        Dim ZZQHZXZXS1Array(10000) As Single
        Dim ZZQHZXZXS2Array(10000) As Single
        '不同燃气轮机负荷率情况下，汽轮机(1)主蒸汽焓值修正系数
        For a = 1 To n
            For b = 4 To 10
                If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                    ZZQHZXZXS1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value), 5)
                ElseIf (RJFHL1Array(a) = 0) Then
                    ZZQHZXZXS1Array(a) = 0
                ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                    XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                    Exit For
                ElseIf (RJFHL1Array(a) > 1) Then
                    XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                    Exit For
                End If
            Next
        Next
        '环境温度修正汽轮机（1）主蒸汽焓值修正系数
        For a = 1 To n
            For b = 16 To 21
                If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 29).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value) * ZZQHZXZXS1Array(a), 5)
                ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                    XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                    Exit For
                End If
            Next
        Next
        '不同燃气轮机负荷率情况下，汽轮机(2)主蒸汽焓值修正系数
        For a = 1 To n
            For b = 4 To 10
                If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                    ZZQHZXZXS2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value), 5)
                ElseIf (RJFHL2Array(a) = 0) Then
                    ZZQHZXZXS2Array(a) = 0
                ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                    XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                    Exit For
                ElseIf (RJFHL2Array(a) > 1) Then
                    XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                    Exit For
                End If
            Next
        Next
        '环境温度修正汽轮机（2）主蒸汽焓值修正系数
        For a = 1 To n
            For b = 16 To 21
                If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 30).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value) * ZZQHZXZXS2Array(a), 5)
                ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                    XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                    Exit For
                End If
            Next
        Next
    End Sub
    Public Sub 计算抽背蒸汽轮机背压蒸汽焓值修正系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '抽背蒸汽轮机背压蒸汽焓值（kj/kg）修正系数
        '定义数组用于储存背压蒸汽焓值修正系数计算过程量
        Dim BYZQHZXZXS1Array(10000) As Single
        Dim BYZQHZXZXS2Array(10000) As Single
        '不同燃气轮机负荷率情况下，抽背蒸汽轮机（1）背压蒸汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 28 To 34
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        BYZQHZXZXS1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value), 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        BYZQHZXZXS1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽背蒸汽轮机（1）背压蒸汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 40 To 45
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 31).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value) * BYZQHZXZXS1Array(a), 5)
                        '抽凝机（1）抽汽焓值修正系数等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 33).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，抽背蒸汽轮机（2）背压蒸汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 28 To 34
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        BYZQHZXZXS2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value), 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        BYZQHZXZXS2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL2Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽背蒸汽轮机（2）背压蒸汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 40 To 45
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 32).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value) * BYZQHZXZXS2Array(a), 5)
                        '抽凝机（2）抽汽焓值修正系数等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 34).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub 计算抽背蒸汽轮机基准背压蒸汽焓值()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '抽背蒸汽轮机（1）计算用基准背压焓值（kj/kg）
        '定义数组，用于储存计算过程量
        Dim JZBYHZ1Array(10000) As Single
        Dim JZBYHZ2Array(10000) As Single
        '不同燃气轮机负荷率情况下，抽背机（1）计算用基准背压蒸汽焓值（kj/kg），计算结果储存在数组中
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 52 To 58
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        JZBYHZ1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value), 2)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        JZBYHZ1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正，背压汽轮机（1）计算用基准背压蒸汽焓值（kj/kg）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 64 To 69
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 37).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value) * JZBYHZ1Array(a), 2)
                        '抽凝机（1）进入凝汽器排汽焓值（kj/kg）等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 35).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，抽背机（2）计算用基准背压蒸汽焓值（kj/kg），计算结果储存在数组中
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 52 To 58
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        JZBYHZ2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value), 2)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        JZBYHZ2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL2Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正，背压汽轮机（2）计算用基准背压蒸汽焓值（kj/kg）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机" Then
            For a = 1 To n
                For b = 64 To 69
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 38).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value) * JZBYHZ2Array(a), 2)
                        '抽凝机（2）进入凝汽器排汽焓值（kj/kg）等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 36).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub 计算抽凝蒸汽轮机排汽焓值()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '抽凝机进入凝汽器的排汽焓值（kj/kg）
        '定义数组，用于储存排汽焓值计算过程量
        Dim CNPQHZ1Array(10000) As Single
        Dim CNPQHZ2Array(10000) As Single
        '不同燃气轮机负荷率情况下，抽凝机（1）进入凝汽器排汽焓值（kj/kg），计算结果储存在数组中
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 76 To 82
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        CNPQHZ1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value), 2)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        CNPQHZ1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽凝机（1）进入凝汽器排汽焓值（kj/kg）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 88 To 93
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 35).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value) * CNPQHZ1Array(a), 2)
                        '抽背机（1）基准背压焓值（kj/kg）等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 37).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，抽凝机（2）进入凝汽器排汽焓值（kj/kg），计算结果储存在数组中
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 76 To 82
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        CNPQHZ2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value), 2)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        CNPQHZ2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL2Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽凝机（2）进入凝汽器排汽焓值（kj/kg）
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
            For a = 1 To n
                For b = 88 To 93
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 36).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value) * CNPQHZ2Array(a), 2)
                        '抽背机（2）背压焓值修正系数等于0
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 38).Value = 0
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub 计算抽凝蒸汽轮机或抽背蒸汽轮机抽汽焓值修正系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '抽凝蒸汽轮机或抽背蒸汽轮机抽汽焓值（kj/kg）修正系数，修正系数求取的数据库与背压焓值修正系数一致
        '定义数组用于储存抽汽焓值修正系数计算过程量
        Dim CQHZXZXS1Array(10000) As Single
        Dim CQHZXZXS2Array(10000) As Single
        '不同燃气轮机负荷率情况下，抽凝蒸汽轮机（1）或抽背蒸汽轮机（1）抽汽蒸汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value <> "无" Then
            For a = 1 To n
                For b = 28 To 34
                    If (RJFHL1Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL1Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        CQHZXZXS1Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL1Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value), 5)
                    ElseIf (RJFHL1Array(a) = 0) Then
                        CQHZXZXS1Array(a) = 0
                    ElseIf (RJFHL1Array(a) > 0 And RJFHL1Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL1Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（1）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽凝蒸汽轮机（1）或抽背蒸汽轮机（1）抽汽焓值修正系数，修正系数求取的数据库与背压焓值修正系数一致
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value <> "无" Then
            For a = 1 To n
                For b = 40 To 45
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 33).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(kk, b).Value) * CQHZXZXS1Array(a), 5)
                        '如果是抽凝机
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽凝机" Then
                            '抽背机（1）背压焓值修正系数等于0
                            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 31).Value = 0
                        End If
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '不同燃气轮机负荷率情况下，抽凝蒸汽轮机（2）或抽背蒸汽轮机（2）抽汽焓值修正系数，修正系数求取的数据库与背压焓值修正系数一致
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value <> "无" Then
            For a = 1 To n
                For b = 28 To 34
                    If (RJFHL2Array(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And RJFHL2Array(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        CQHZXZXS2Array(a) = ExcelApp.WorksheetFunction.RoundUp(((((RJFHL2Array(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value), 5)
                    ElseIf (RJFHL2Array(a) = 0) Then
                        CQHZXZXS2Array(a) = 0
                    ElseIf (RJFHL2Array(a) > 0 And RJFHL2Array(a) < 0.4) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以小于40%", vbOK)
                        Exit For
                    ElseIf (RJFHL2Array(a) > 1) Then
                        XZ = MsgBox("燃气轮机（2）负荷率不可以大于100%", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
        '环境温度修正抽凝蒸汽轮机（2）或抽背蒸汽轮机（2）抽汽焓值修正系数
        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value <> "无" Then
            For a = 1 To n
                For b = 40 To 45
                    If (TEMPArray(a) >= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value And TEMPArray(a) <= ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value) Then
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 34).Value = ExcelApp.WorksheetFunction.RoundUp(((((TEMPArray(a) - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value) / (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(3, b).Value)) * (ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b + 1).Value - ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value)) + ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Cells(ll, b).Value) * CQHZXZXS2Array(a), 5)
                        '如果是抽凝机
                        If ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽凝机" Then
                            '抽背机（2）背压焓值修正系数等于0
                            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + a, 32).Value = 0
                        End If
                    ElseIf (TEMPArray(a) < -20 Or TEMPArray(a) > 40) Then
                        XZ = MsgBox("环境温度不可以大于40℃或者低于-20℃", vbOK)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
    Public Sub 通过背压蒸汽需求量反算燃气轮机负荷率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '定义局部变量
        Dim b1, b2
        '通过输入的抽背蒸汽轮机背压蒸汽需求量，倒算此时需要的背压联合循环燃气轮机负荷率
        '定义数组，用于储存输入的抽背蒸汽轮机背压蒸汽需求量(t/h)
        Dim CBJZQXQL1Array(10000) As Single
        Dim CBJZQXQL2Array(10000) As Single
        '读取抽背蒸汽轮机（1）（2）的背压蒸汽需求量
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 2).Value = "抽背机(1)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机") Then
            '记录背压证汽轮机(1)背压蒸汽需求量
            For j = 1 To n '工况序号
                CBJZQXQL1Array(j) = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 2).Value '抽背机(1)背压蒸汽量(t/h)
            Next
            '清空输入的抽背蒸汽轮机(1)需求量
            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("B7:B306").ClearContents
        End If
        '抽背蒸汽轮机(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 3).Value = "抽背机(2)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机") Then
            '记录背压证汽轮机(2)背压蒸汽需求量
            For j = 1 To n '工况序号
                CBJZQXQL2Array(j) = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 3).Value
            Next
            '清空输入的抽背蒸汽轮机(2)需求量
            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range("C7:C306").ClearContents
        End If
        '通过抽背蒸汽轮机（1）蒸汽需求量反算汽轮机(1)负荷率
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 2).Value = "抽背机(1)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value = "抽背机") Then
            For j = 1 To n '工况序号
                '通过背压蒸汽（1）需求量反算此时燃气轮机（1）负荷率
                For a1 = 80 To 200
                    b1 = 0.005 * a1 '燃气轮机负荷率从0开始增加，最大到1
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 2).Value = b1
                    '读取此时的燃气轮机（1）负荷率储存在数组中
                    RJFHL1Array(j) = b1
                    Call 计算不同环境温度下燃气轮机最大发电功率()
                    Call 燃机汽轮计算用基准发电效率()
                    Call 计算余热锅炉效率()
                    Call 计算蒸汽轮机蒸汽循环总效率()
                    Call 计算主蒸汽焓值修正系数()
                    Call 计算抽背蒸汽轮机背压蒸汽焓值修正系数()
                    Call 计算抽背蒸汽轮机基准背压蒸汽焓值()
                    '进行一次计算
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 19), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 19), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 38)).Value
                    '判断背压蒸汽量是否满足要求
                    If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 8).Value >= CBJZQXQL1Array(j) Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(20, 8).Value >= CBJZQXQL1Array(j)) Then
                        '记录下此时的燃气轮机（1）负荷率
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 2).Value = b1
                        '跳出循环
                        Exit For
                    End If
                Next
            Next
        End If
        '抽背蒸汽轮机(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)(2)
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 3).Value = "抽背机(2)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value = "抽背机") Then
            For j = 1 To n '工况序号
                '通过背压蒸汽（2）需求量反算此时燃气轮机（2）负荷率
                For a2 = 80 To 200
                    b2 = 0.005 * a2 '燃气轮机负荷率从0开始增加，最大到1
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 3).Value = b2
                    '读取此时的燃气轮机（2）负荷率储存在数组中
                    RJFHL2Array(j) = b2
                    Call 计算不同环境温度下燃气轮机最大发电功率()
                    Call 燃机汽轮计算用基准发电效率()
                    Call 计算余热锅炉效率()
                    Call 计算蒸汽轮机蒸汽循环总效率()
                    Call 计算主蒸汽焓值修正系数()
                    Call 计算抽背蒸汽轮机背压蒸汽焓值修正系数()
                    Call 计算抽背蒸汽轮机基准背压蒸汽焓值()
                    '进行一次计算
                    ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 19), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 19), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 38)).Value
                    '判断背压蒸汽量是否满足要求
                    If (ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 8).Value >= CBJZQXQL2Array(j) Or ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(42, 8).Value >= CBJZQXQL2Array(j)) Then
                        '记录下此时的燃气轮机（2）负荷率
                        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + j, 3).Value = b2
                        '跳出循环
                        Exit For
                    End If
                Next
            Next
        End If
        '添加报错功能
        '联合循环（1）
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 2).Value = "抽背机(1)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(11, 2).Value <> "抽背机") Then
            XZ = MsgBox("选择为'抽背机(1)蒸汽量(t/h)'时，联合循环(1)必须为抽背蒸汽轮机。", vbOKCancel)
            Call 锁定工作表()
            Exit Sub
        End If
        '联合循环（2）
        If (ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6, 3).Value = "抽背机(2)蒸汽量(t/h)" And ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Cells(33, 2).Value <> "抽背机") Then
            XZ = MsgBox("选择为'抽背机(2)蒸汽量(t/h)'时，联合循环(2)必须为抽背蒸汽轮机。", vbOKCancel)
            Call 锁定工作表()
            Exit Sub
        End If
    End Sub
    Public Sub 变工况计算循环体()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '变工况计算循环体
        For i = 1 To n
            '读取变工况计算输入量
            ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(3, 38)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + i, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Cells(6 + i, 38)).Value
            '返回变工况计算结果
            ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(6 + i, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(6 + i, 35)).Value = ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Range(ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 2), ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Cells(3, 35)).Value
        Next
    End Sub
    Public Sub 进入维护模式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Call Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定工作表()
            ZTJC = 0
            Exit Sub
        End If
        '读取计数
        If (ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value < 3 And ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value >= 0) Then '最多只可以连续错3次。单元格S1
            Form1.ShowDialog() '窗口显示
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            Form1.TextBox1.Text = Nothing '清空已有的内容
            '剩余步骤在"进入维护模式.确定"
        Else
            MsgBox("已超过最大尝试次数，不可以再尝试输入密码！")
            Call 工作表设置为彻底隐藏()
            Call 锁定工作表()
        End If
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Public Sub 取消工作表彻底隐藏()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.Worksheets("变工况成本计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("变工况蒸汽计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("联合循环变工况计算器").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("联合循环单一工况计算器").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("联合循环综合效率").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("蒸汽信息汇总").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("燃气轮机变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("汽轮机变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("余热锅炉变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("蒸汽换热供暖(生活热水)计算").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("热电分析结果").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("汽轮机蒸汽参数变工况&环境温度").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("空气源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("燃气轮机(1)额定参数数据库").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("燃气轮机(2)额定参数数据库").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("设备组合模式统计").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("热电分析计算过程").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("溴化锂制冷(0.4Mpa)").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("溴化锂制冷(0.6Mpa)").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("溴化锂制冷(0.8Mpa)").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("离心式冷水机组").Visible = Excel.XlSheetVisibility.xlSheetHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetHidden
    End Sub
    Public Sub 工作表设置为彻底隐藏()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.Worksheets("变工况成本计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("变工况蒸汽计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("联合循环变工况计算器").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("联合循环单一工况计算器").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("联合循环综合效率").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("蒸汽信息汇总").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("燃气轮机变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("汽轮机变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("余热锅炉变负荷&环境温度").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("蒸汽换热供暖(生活热水)计算").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("热电分析结果").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("汽轮机蒸汽参数变工况&环境温度").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("空气源热泵（冷水）机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("燃气轮机(1)额定参数数据库").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("燃气轮机(2)额定参数数据库").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("设备组合模式统计").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("热电分析计算过程").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("溴化锂制冷(0.4Mpa)").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("溴化锂制冷(0.6Mpa)").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("溴化锂制冷(0.8Mpa)").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("离心式冷水机组").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("天然气锅炉").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
    End Sub
    Public Sub 锁定工作表()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况成本计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况蒸汽计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环变工况计算器").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环单一工况计算器").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环综合效率").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("蒸汽信息汇总").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本测算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入测算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("蒸汽换热供暖(生活热水)计算").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("热电分析结果").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("空气源热泵（冷水）机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机(1)额定参数数据库").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机(2)额定参数数据库").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备组合模式统计").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("热电分析计算过程").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.4Mpa)").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.6Mpa)").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.8Mpa)").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式冷水机组").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("供能单价的确定").Protect(Password:="wscjc")
    End Sub
    Public Sub 解锁工作表()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Call Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定工作表()
            ZTJC = 0
            Exit Sub
        End If
        ExcelApp.ThisWorkbook.Worksheets("变工况计算输入").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况计算结果输出").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况成本计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("变工况蒸汽计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环变工况计算器").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环单一工况计算器").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("说明&常量设置&数据汇总").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备选型&经济指标").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("联合循环综合效率").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("蒸汽信息汇总").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机变负荷&环境温度").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("汽轮机变负荷&环境温度").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("余热锅炉变负荷&环境温度").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本测算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入测算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("蒸汽换热供暖(生活热水)计算").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("热电分析结果").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("汽轮机蒸汽参数变工况&环境温度").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("空气源热泵（冷水）机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机(1)额定参数数据库").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("燃气轮机(2)额定参数数据库").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("设备组合模式统计").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("热电分析计算过程").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.4Mpa)").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.6Mpa)").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("溴化锂制冷(0.8Mpa)").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("离心式冷水机组").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("天然气锅炉").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("供能单价的确定").unProtect(Password:="wscjc")
    End Sub
    Public Sub 自保护程序()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            '验证本地C盘是否存在白名单文件，用来区分是不是自己的电脑，是否需要进行自保护程序验证
            Dim fs As New FileStream("C:\Windows\WhiteList_CJC.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim WhiteList_PC As String = strTemp '获取TXT文本内容
            sr.Close()
            fs.Close()
            If WhiteList_PC <> "WhiteList_PC" Then '如果不是合格的白名单文件，验证失败，则进行自保护验证
                'Call 获取本地服务器版本信息并验证()
                Call 获取本机MAC地址并验证()
                Call 网络时间和本地时间交替验证()
            End If
        Catch ex As Exception
            '发生任何异常，则进行自保护验证
            'Call 获取本地服务器版本信息并验证()
            Call 获取本机MAC地址并验证()
            Call 网络时间和本地时间交替验证()
            Exit Sub
        End Try
    End Sub
    Public Sub 获取本地服务器版本信息并验证()
        'On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            Dim fs As New FileStream("\\192.168.9.201\user1\CJC-陈俊丞\综合能源多联供项目计算工具\燃气轮机版计算程序\Version.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim FWQBBH_String As String = strTemp '服务器版本号，字符串格式
            sr.Close()
            fs.Close()
            Dim FWQBBH = CInt(FWQBBH_String) '将服务器版本号从字符串格式转为整数型格式
            If FWQBBH > 20190401 Then '如果服务器版本号大于设定的内置版本号信息，则退出程序
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Version Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        Catch ex As Exception
            '发生任何异常，直接退出
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Version Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
            Exit Sub
        End Try
    End Sub
    Public Sub 获取本机MAC地址并验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim wmiObjSet As WbemScripting.SWbemObjectSet
        Dim obj As WbemScripting.SWbemObject
        Dim MAC As String = Nothing
        Dim MACTEST As Integer = 0 'MAC地址检测，1代表通过，0代表不通过
        wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapterConfiguration")
        For Each obj In wmiObjSet
            MAC = obj.MACAddress
            'MAC地址白名单(随便写一个)
            If MAC = "45:40:C1:8F:AD:4A" Then
                MACTEST = 1
                Exit For
            Else
                MACTEST = 0
            End If
        Next
        '————————————————————————————————————————————————————————————
        '针对MAC地址检测结果，进行不同操作
        If MACTEST = 0 Then '如果MAC地址检测不通过
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local MAC Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Public Sub 网络时间和本地时间交替验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '程序开始
        '先验证网络时间
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = "TimeOut" '网络状态默认值
        For i = 1 To 10 '最多Ping网络10次
            A = CStr(Pings()) '反回函数状态
            If A = "Success" Then
                A = "Success"
                Exit For '只要检测出ping通，就跳出循环
            End If
        Next
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then '如果网络是通的
            With CreateObject("MSXML2.ServerXMLHTTP") '获取网络时间
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #01/01/2020# Then '月/日/年，验证网络时间
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                End If
            End With
        Else '如果网络不通，验证系统本地时间
            '进程暂停一段时间（2000分钟）
            Threading.Thread.Sleep(120000000)
            Dim Local_Time = Date.Now '获取系统本地时间
            Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
            If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Internet Ping Error and Local Date Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        End If
    End Sub
    Public Sub 获取系统时间并验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim Local_Time = Date.Now '获取系统本地时间
        Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
        If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local Date Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Public Sub 程序联网验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = CStr(Pings()) '反回函数状态
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then
            With CreateObject("MSXML2.ServerXMLHTTP")
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #01/01/2020# Then '月/日/年
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                    ''最多只能够输入错误3次密码
                    'If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 26).Value < 3 And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 26).Value >= 0 Then
                    '    Dim mima = InputBox("程序已过有效期，请输入密码继续使用：")
                    '    If mima <> "Speri1115" Then
                    '        '密码输入错误
                    '        '密码输入错误次数加1
                    '        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 26).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 26).Value + 1
                    '        '保存表格的改动
                    '        ExcelApp.Application.DisplayAlerts = False
                    '        ExcelApp.ThisWorkbook.Save()
                    '        ExcelApp.Application.DisplayAlerts = True
                    '        '错误提示
                    '        MsgBox("密码输入错误，程序即将关闭！")
                    '        '直接退出表格
                    '        ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    '        ExcelApp.Application.DisplayAlerts = False
                    '        ExcelApp.Application.Quit()
                    '        ExcelApp.Application.DisplayAlerts = True
                    '    Else
                    '        '密码输入正确
                    '        '密码错误计数重置为0
                    '        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 26).Value = 0
                    '        '保存表格变动
                    '        ExcelApp.Application.DisplayAlerts = False
                    '        ExcelApp.ThisWorkbook.Save()
                    '        ExcelApp.Application.DisplayAlerts = True
                    '        '提示
                    '        MsgBox("密码正确，程序可以继续使用！")
                    '    End If
                    'Else
                    '    MsgBox("已超过最大尝试次数，不可以再尝试输入密码！")
                    '    '直接退出表格
                    '    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    '    ExcelApp.Application.DisplayAlerts = False
                    '    ExcelApp.Application.Quit()
                    '    ExcelApp.Application.DisplayAlerts = True
                    'End If
                End If
            End With
        Else
            MsgBox("网络连接异常，请检查网络连接！本程序必需联网验证才可以使用！")
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Function Pings() As String
        Dim _ping As New Net.NetworkInformation.Ping
        Dim _pingreply As Net.NetworkInformation.PingReply = _ping.Send("www.baidu.com") 'ping 百度
        Return _pingreply.Status.ToString()
        '返回以下信息
        'TimedOut 失败
        'Success  成功
    End Function
    '————————————————————————————————————————————————————————————————————————————————————————
    '————————————————————————————————————————————————————————————————————————————————————————
    Public Shared Form1 As New 进入维护模式
    Public ZTJC
    '申明全局数组和变量
    'kk和ll为燃机所在行号
    Public kk As Integer
    Public ll As Integer
    '燃气轮机（1）（2）基准发电功率(kW)
    Public rj1kw As Double
    Public rj2kw As Double
    '工况数量
    Public n As Integer
    '各个工况的环境温度
    Public TEMPArray(10000) As Double
    '燃机汽轮（1）（2）负荷率
    Public RJFHL1Array(10000) As Double
    Public RJFHL2Array(10000) As Double
    Public XZ
End Class
