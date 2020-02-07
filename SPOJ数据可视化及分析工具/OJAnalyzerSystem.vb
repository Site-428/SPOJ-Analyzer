Module OJAnalyzerSystem
    Public Class OJSystemInfo
        ''' <summary>
        ''' 统计起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public StartDate As Date
        ''' <summary>
        ''' 统计结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public EndDate As Date
        ''' <summary>
        ''' 按日记录的新题目数量。
        ''' </summary>
        ''' <remarks></remarks>
        Public NewProblemCount As Dictionary(Of Date, Integer)
        ''' <summary>
        ''' 按每日分时段时记录的提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountByHour() As Integer
        ''' <summary>
        ''' 默认构造函数
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            StartDate = New Date(1000, 1, 1)
            EndDate = New Date(9999, 12, 31)
            NewProblemCount = New Dictionary(Of Date, Integer)
            ReDim SubmitCountByHour(0 To 23)
            Dim i As Integer
            For i = 0 To 23
                SubmitCountByHour(i) = 0
            Next
        End Sub
        ''' <summary>
        ''' 指定了统计起讫日期的构造函数。
        ''' </summary>
        ''' <param name="DateStart">统计起始日期。</param>
        ''' <param name="DateEnd">统计结束日期。</param>
        ''' <remarks></remarks>
        Public Sub New(DateStart As Date, DateEnd As Date)
            StartDate = DateStart
            EndDate = DateEnd
            If DateEnd < DateStart Then
                EndDate = DateStart
            End If
            NewProblemCount = New Dictionary(Of Date, Integer)
            ReDim SubmitCountByHour(0 To 23)
            Dim i As Integer
            For i = 0 To 23
                SubmitCountByHour(i) = 0
            Next
        End Sub
    End Class
    ''' <summary>
    ''' OnlineJudge系统整体资讯。
    ''' </summary>
    ''' <remarks></remarks>
    Public OJSysInfo As New OJSystemInfo
    ''' <summary>
    ''' 用户选择的分析起始日期。
    ''' </summary>
    ''' <remarks></remarks>
    Public UserSpecifiedAnalyzeStartDate As Date
    ''' <summary>
    ''' 用户选择的分析结束日期。
    ''' </summary>
    ''' <remarks></remarks>
    Public UserSpecifiedAnalyzeEndDate As Date
End Module
