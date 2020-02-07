Module OJAnalyzerProblems
    Public Class OJProblemInfo
        ''' <summary>
        ''' 题目编号。
        ''' </summary>
        ''' <remarks></remarks>
        Public ProblemIDNumber As String
        ''' <summary>
        ''' 作业次序。
        ''' </summary>
        ''' <remarks></remarks>
        Public ProblemTaskSequenceIndex As Integer
        ''' <summary>
        ''' 题目被创建的时间。
        ''' </summary>
        ''' <remarks></remarks>
        Public CreateTime As Date
        ''' <summary>
        ''' 总提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCount As Integer
        ''' <summary>
        ''' 参与人数。
        ''' </summary>
        ''' <remarks></remarks>
        Public ParticipantCount As Integer
        ''' <summary>
        ''' 通过(AC)次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public ACCount As Integer
        ''' <summary>
        ''' 付出指数Jq。
        ''' </summary>
        ''' <remarks></remarks>
        Public EffortValue_Jq As Double
        ''' <summary>
        ''' 参与指数Eq。
        ''' </summary>
        ''' <remarks></remarks>
        Public ParticipateValue_Eq As Double
        ''' <summary>
        ''' 按日记录的提交次数。
        ''' </summary>
        ''' <remarks></remarks>
        Public SubmitCountByDay As Dictionary(Of Date, Integer)
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
            ProblemIDNumber = ""
            CreateTime = New Date(1000, 1, 1)
            SubmitCount = 0
            ParticipantCount = 0
            ACCount = 0
            EffortValue_Jq = 0
            ParticipateValue_Eq = 0
            ClustResult = 0
            SubmitCountByDay = New Dictionary(Of Date, Integer)
        End Sub
        ''' <summary>
        ''' 指定了题目编号的构造函数。
        ''' </summary>
        ''' <param name="ProblemID">题目的编号，为字符串值。</param>
        ''' <remarks></remarks>
        Public Sub New(ProblemID As String)
            ProblemIDNumber = ProblemID
            CreateTime = New Date(1000, 1, 1)
            SubmitCount = 0
            ParticipantCount = 0
            ACCount = 0
            EffortValue_Jq = 0
            ParticipateValue_Eq = 0
            ClustResult = 0
            SubmitCountByDay = New Dictionary(Of Date, Integer)
        End Sub
    End Class
    ''' <summary>
    ''' 题目编号列表。
    ''' </summary>
    ''' <remarks></remarks>
    Public ProblemList As New List(Of String)
    ''' <summary>
    ''' 存放题目信息的字典。
    ''' </summary>
    ''' <remarks></remarks>
    Public ProblemDictionary As New Dictionary(Of String, OJProblemInfo)
End Module
