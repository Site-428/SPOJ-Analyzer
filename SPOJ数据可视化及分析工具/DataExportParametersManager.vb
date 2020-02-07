Module DataExportParametersManager
    ''' <summary>
    ''' 导出的数据的文件格式。
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum AnalyzedDataExportFormat
        JSONFile = 1
        XMLFile = 2
    End Enum
    ''' <summary>
    ''' 导出数据所用的参数。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AnalyzedDataExportParamers
        ''' <summary>
        ''' 导出的数据存放的目录。
        ''' </summary>
        ''' <remarks></remarks>
        Public OutputDirectory As String
        ''' <summary>
        ''' 导出的数据的文件格式。
        ''' </summary>
        ''' <remarks></remarks>
        Public OutputFileFormat As AnalyzedDataExportFormat = AnalyzedDataExportFormat.JSONFile
        ''' <summary>
        ''' 取得或设定是否合并导出的数据到同一个文件中。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsOutputFileMerged As Boolean = False
        ''' <summary>
        ''' 取得或设定需要导出的数据的索引。
        ''' </summary>
        ''' <remarks></remarks>
        Public DataToExport As New List(Of String)
    End Class
    ''' <summary>
    ''' 合并的导出的学生数据的数据集。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MergedStudentDataExportFormat
        ''' <summary>
        ''' 分析起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeStartDate As Date
        ''' <summary>
        ''' 分析结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeEndDate As Date
        ''' <summary>
        ''' 数据是否为已合并的数据，指示数据集为单一数据形式或列表 (数组) 形式。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsDataMerged As Boolean
        ''' <summary>
        ''' 合并的学生数据集。
        ''' </summary>
        ''' <remarks></remarks>
        Public StudentDataSet As New List(Of OJStudentInfo)
    End Class
    ''' <summary>
    ''' 未合并的导出的学生数据的数据集。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SingleStudentDataExportFormat
        ''' <summary>
        ''' 分析起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeStartDate As Date
        ''' <summary>
        ''' 分析结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeEndDate As Date
        ''' <summary>
        ''' 数据是否为已合并的数据，指示数据集为单一数据形式或列表 (数组) 形式。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsDataMerged As Boolean
        ''' <summary>
        ''' 单一学生数据集。
        ''' </summary>
        ''' <remarks></remarks>
        Public StudentDataSet As New OJStudentInfo
    End Class
    ''' <summary>
    ''' 合并的导出的题目数据的数据集。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MergedProblemDataExportFormat
        ''' <summary>
        ''' 分析起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeStartDate As Date
        ''' <summary>
        ''' 分析结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeEndDate As Date
        ''' <summary>
        ''' 数据是否为已合并的数据，指示数据集为单一数据形式或列表 (数组) 形式。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsDataMerged As Boolean
        ''' <summary>
        ''' 合并的题目数据集。
        ''' </summary>
        ''' <remarks></remarks>
        Public ProblemDataSet As New List(Of OJProblemInfo)
    End Class
    ''' <summary>
    ''' 未合并的导出的题目数据的数据集。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SingleProblemDataExportFormat
        ''' <summary>
        ''' 分析起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeStartDate As Date
        ''' <summary>
        ''' 分析结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeEndDate As Date
        ''' <summary>
        ''' 数据是否为已合并的数据，指示数据集为单一数据形式或列表 (数组) 形式。
        ''' </summary>
        ''' <remarks></remarks>
        Public IsDataMerged As Boolean
        ''' <summary>
        ''' 单一题目数据集。
        ''' </summary>
        ''' <remarks></remarks>
        Public ProblemDataSet As New OJProblemInfo
    End Class
    Public Class OJSystemDataExportFormat
        ''' <summary>
        ''' 统计起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public LogStartDate As Date
        ''' <summary>
        ''' 统计结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public LogEndDate As Date
        ''' <summary>
        ''' 用户选择的分析起始日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeStartDate As Date
        ''' <summary>
        ''' 用户选择的分析结束日期。
        ''' </summary>
        ''' <remarks></remarks>
        Public AnalyzeEndDate As Date
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
    End Class
    ''' <summary>
    ''' 学生数据的导出参数组。
    ''' </summary>
    ''' <remarks></remarks>
    Public StudentDataOutputParameters As New AnalyzedDataExportParamers
    ''' <summary>
    ''' 题目数据的导出参数组。
    ''' </summary>
    ''' <remarks></remarks>
    Public ProblemDataOutputParameters As New AnalyzedDataExportParamers
    ''' <summary>
    ''' 生成YYYYMMMDD-HHMM格式的时间戳。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BuildTimeStamp() As String
        Dim TimeStamp As New String(String.Format("{0:D4}", Now.Year) & String.Format("{0:D2}", Now.Month) & String.Format("{0:D2}", Now.Day) & "-" & String.Format("{0:D2}", Now.Hour) & String.Format("{0:D2}", Now.Minute))
        Return TimeStamp
    End Function
End Module
