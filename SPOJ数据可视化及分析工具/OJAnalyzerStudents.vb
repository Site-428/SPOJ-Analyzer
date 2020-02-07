Module OJAnalyzerStudents
    ''' <summary>
    ''' 参与OnlineJudge的学生的信息。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class OJStudentInfo
        ''' <summary>
        ''' 学生的学号。
        ''' </summary>
        ''' <remarks></remarks>
        Public StudentIDNumber As String
        ''' <summary>
        ''' 总提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCount As Integer
        ''' <summary>
        ''' 通过(AC)次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public ACCount As Integer
        ''' <summary>
        ''' 按日记录的提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountByDay As Dictionary(Of Date, Integer)
        ''' <summary>
        ''' 工作日0:00至12:00提交次数。
        ''' </summary>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
        ''' <remarks></remarks>
        Public SubmitCountOnWorkdayAM As Integer
        ''' <summary>
        ''' 工作日12:00至24:00提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountOnWorkdayPM As Integer
        ''' <summary>
        ''' 休息日0:00至12:00提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountOnRestdayAM As Integer
        ''' <summary>
        ''' 休息日12:00至24:00提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountOnRestdayPM As Integer
        ''' <summary>
        ''' 按每日分时段时记录的提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountByHour() As Integer
        ''' <summary>
        ''' 完成指数Ac。
        ''' </summary>
        ''' <remarks>Ac = Pac+α，Pac为通过数ACCount，α为通过率ACCount/SubmitCount。</remarks>
        Public FittingAC As Double
        ''' <summary>
        ''' 稳定指数Stb。
        ''' </summary>
        ''' <remarks>拟合曲线的调整R平方值。</remarks>
        Public FittingR_Stb As Double
        ''' <summary>
        ''' 使用指数Kb。
        ''' </summary>
        ''' <remarks>拟合曲线的斜率。</remarks>
        Public FittingK_Kb As Double
        ''' <summary>
        ''' 拟合曲线的截距。
        ''' </summary>
        ''' <remarks></remarks>
        Public FittingB As Double
        ''' <summary>
        ''' 聚类分析结果。
        ''' </summary>
        ''' <remarks></remarks>
        Public ClustResult As Double
        ''' <summary>
        ''' 默认的构造函数。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            StudentIDNumber = ""
            SubmitCount = 0
            ACCount = 0
            SubmitCountOnWorkdayAM = 0
            SubmitCountOnWorkdayPM = 0
            SubmitCountOnRestdayAM = 0
            SubmitCountOnRestdayPM = 0
            SubmitCountByDay = New Dictionary(Of Date, Integer)
            ReDim SubmitCountByHour(0 To 23)
            Dim i As Integer
            For i = 0 To 23
                SubmitCountByHour(i) = 0
            Next
            FittingAC = 0
            FittingK_Kb = 0
            FittingR_Stb = 0
            ClustResult = 0
        End Sub
        ''' <summary>
        ''' 指定了学生学号的构造函数。
        ''' </summary>
        ''' <param name="StudentID">学生的学号，为字符串值。</param>
        ''' <remarks></remarks>
        Public Sub New(StudentID As String)
            StudentIDNumber = StudentID
            SubmitCount = 0
            ACCount = 0
            SubmitCountOnWorkdayAM = 0
            SubmitCountOnWorkdayPM = 0
            SubmitCountOnRestdayAM = 0
            SubmitCountOnRestdayPM = 0
            SubmitCountByDay = New Dictionary(Of Date, Integer)
            ReDim SubmitCountByHour(0 To 23)
            Dim i As Integer
            For i = 0 To 23
                SubmitCountByHour(i) = 0
            Next
            FittingAC = 0
            FittingK_Kb = 0
            FittingR_Stb = 0
            ClustResult = 0
        End Sub
        ''' <summary>
        ''' 获取距离指定的起始日期某一日数的提交数。
        ''' </summary>
        ''' <param name="DayCount">距离起始日期的日数，可谓负数、正数或0。</param>
        ''' <param name="UseCustomizedStartDate">可选，是否使用自订的起始日期，默认为False。若为True则需提供起始日期。</param>
        ''' <param name="CustomizedStartDate">可选，自订的起始日期，默认为统计起始日期。</param>
        ''' <returns>指定日期的提交数。</returns>
        ''' <remarks></remarks>
        Public Function GetSubmitCountByDayByDayCount(DayCount As Integer, Optional UseCustomizedStartDate As Boolean = False, Optional CustomizedStartDate As Date = Nothing) As Integer
            If Not UseCustomizedStartDate Then
                If SubmitCountByDay.ContainsKey(OJSysInfo.StartDate.AddDays(DayCount)) Then
                    Return SubmitCountByDay(OJSysInfo.StartDate.AddDays(DayCount))
                Else
                    Return 0
                End If
            Else
                If IsNothing(CustomizedStartDate) Then
                    CustomizedStartDate = OJSysInfo.StartDate
                End If
                If SubmitCountByDay.ContainsKey(CustomizedStartDate.AddDays(DayCount)) Then
                    Return SubmitCountByDay(CustomizedStartDate.AddDays(DayCount))
                Else
                    Return 0
                End If
            End If
        End Function
    End Class
    ''' <summary>
    ''' 学生学号列表。
    ''' </summary>
    ''' <remarks></remarks>
    Public StudentList As New List(Of String)
    ''' <summary>
    ''' 存放学生信息的字典。
    ''' </summary>
    ''' <remarks></remarks>
    Public StudentDictionary As New Dictionary(Of String, OJStudentInfo)
    ''' <summary>
    ''' 处理学生聚类结果到评等结果的映射字典。
    ''' </summary>
    ''' <remarks></remarks>
    Public StudentClustResultMapping As New Dictionary(Of Integer, String)
End Module
