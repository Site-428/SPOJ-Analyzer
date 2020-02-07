Module OJAnalyzerTests
    ''' <summary>
    ''' 测试/题目集数据
    ''' </summary>
    ''' <remarks></remarks>
    Public Class OJTestInfo
        ''' <summary>
        ''' 测试/题目集编号。
        ''' </summary>
        ''' <remarks></remarks>
        Public TestID As Integer
        ''' <summary>
        ''' 开始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public BeginDate As Date
        ''' <summary>
        ''' 结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public EndDate As Date
        ''' <summary>
        ''' 解析日志时是否出错导致失败。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsParseFailed As Boolean
        ''' <summary>
        ''' 默认的构造函数。
        ''' </summary>
        ''' <remarks>初始化测试号为-1，开始时间为1000年1月1日，结束时间为9999年12月31日，认为处于错误状态。</remarks>
        Public Sub New()
            TestID = -1
            BeginDate = New Date(1000, 1, 1)
            EndDate = New Date(9999, 12, 31)
            IsParseFailed = True
        End Sub
        ''' <summary>
        ''' 构造非错误状态的测试数据的构造函数。
        ''' </summary>
        ''' <param name="ID">测试/题目集编号。</param>
        ''' <param name="DateBegin">开始日期。</param>
        ''' <param name="DateEnd">结束日期。</param>
        ''' <param name="LeaveErrorState">是否保留错误状态，默认为否。</param>
        ''' <remarks>默认认为不处于错误状态。</remarks>
        Public Sub New(ID As Integer, DateBegin As Date, DateEnd As Date, Optional LeaveErrorState As Boolean = False)
            TestID = ID
            BeginDate = DateBegin
            EndDate = DateEnd
            IsParseFailed = LeaveErrorState
        End Sub
    End Class
    ''' <summary>
    ''' 解析文字编码的测试数据。
    ''' </summary>
    ''' <param name="TestLogLine">文字编码的测试数据，使用半角逗号分隔，依次为测试号、开始日期和结束日期。</param>
    ''' <returns>结构化的测试数据。</returns>
    ''' <remarks></remarks>
    Public Function ParseOJTestLogLine(TestLogLine As String) As OJTestInfo
        Dim Temp As New OJTestInfo
        With Temp
            .TestID = -1
            .BeginDate = New Date(1000, 1, 1)
            .EndDate = New Date(9999, 12, 31)
            .IsParseFailed = True
        End With
        Dim LogLineArray() As String
        Dim LogBeginDate() As String
        Dim LogEndDate() As String
        Try
            LogLineArray = Split(TestLogLine, ",")
            If LogLineArray.Length <> 3 Then
                Return Temp
            End If
            LogBeginDate = Split(LogLineArray(1), "-")
            If LogBeginDate.Length <> 3 Then
                Return Temp
            End If
            LogEndDate = Split(LogLineArray(2), "-")
            If LogEndDate.Length <> 3 Then
                ReDim LogEndDate(2)
                LogEndDate(0) = 9999
                LogEndDate(1) = 12
                LogEndDate(2) = 31
            End If
            With Temp
                .TestID = LogLineArray(0)
                .BeginDate = New Date(Int(LogBeginDate(0)), Int(LogBeginDate(1)), Int(LogBeginDate(2)))
                .EndDate = New Date(Int(LogEndDate(0)), Int(LogEndDate(1)), Int(LogEndDate(2)))
                .IsParseFailed = False
            End With
        Catch ex As Exception
            With Temp
                .TestID = -1
                .BeginDate = New Date(1000, 1, 1)
                .EndDate = New Date(9999, 12, 31)
                .IsParseFailed = True
            End With
        End Try
        Return Temp
    End Function
    Public TestDictionary As New Dictionary(Of Integer, OJTestInfo)
End Module
