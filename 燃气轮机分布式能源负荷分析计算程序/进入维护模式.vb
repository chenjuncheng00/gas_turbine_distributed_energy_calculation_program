Public Class 进入维护模式
    Private Sub 确定_Click(sender As Object, e As EventArgs) Handles 确定.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim 燃气轮机分布式能源负荷分析计算程序 As New Com燃气轮机分布式能源负荷分析计算程序
        Dim mima As String '输入的密码
        mima = CType(Com燃气轮机分布式能源负荷分析计算程序.Form1.TextBox1.Text, String)
        If mima = "cjc19920105" Then
            '解锁工作表
            Call 燃气轮机分布式能源负荷分析计算程序.解锁工作表()
            '取消彻底隐藏工作表
            Call 燃气轮机分布式能源负荷分析计算程序.取消工作表彻底隐藏()
            ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value = 0 '密码输入正确后，将计数器重置为0
            Com燃气轮机分布式能源负荷分析计算程序.Form1.Close()
            MsgBox("已成功进入维护模式，程序可以被编辑！")
        Else
            '密码输入错误次数加1
            ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value = ExcelApp.Worksheets("说明&常量设置&数据汇总").Cells(1, 19).Value + 1
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '报错提醒
            MsgBox("密码错误，请重新输入！")
            Com燃气轮机分布式能源负荷分析计算程序.Form1.Close()
            Call 燃气轮机分布式能源负荷分析计算程序.工作表设置为彻底隐藏()
            Call 燃气轮机分布式能源负荷分析计算程序.锁定工作表()
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
        End If
        '清空文本框中的输入内容
        Com燃气轮机分布式能源负荷分析计算程序.Form1.TextBox1.Clear()
        Me.Close()
    End Sub
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim key As String
        key = e.KeyChar
        '检验按键是否为回车键，如果是就把焦点附给按钮1，并执行Click命令
        If key = Microsoft.VisualBasic.ChrW(13) Then
            确定.Focus()
            确定.PerformClick()
        End If
    End Sub
    Private Sub 进入维护模式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyBase.KeyPreview = True
    End Sub
End Class