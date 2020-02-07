Module OJLogParser
    ''' <summary>
    ''' OJ日志中的日期。
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure OJLogDate
        ''' <summary>
        ''' 年。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Year As Integer
        ''' <summary>
        ''' 月。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Month As Integer
        ''' <summary>
        ''' 日。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Day As Integer
    End Structure
    ''' <summary>
    ''' OJ日志中的时间
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure OJLogTime
        ''' <summary>
        ''' 时。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Hour As Integer
        ''' <summary>
        ''' 分。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Minute As Integer
        ''' <summary>
        ''' 秒。
        ''' </summary>
        ''' <remarks></remarks>
        Dim Second As Integer
    End Structure
    ''' <summary>
    ''' OJ日志数据结构。
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure OJLog
        ''' <summary>
        ''' 题号。
        ''' </summary>
        ''' <remarks></remarks>
        Dim ProblemID As String
        ''' <summary>
        ''' 测试号。
        ''' </summary>
        ''' <remarks></remarks>
        Dim TestID As Integer
        ''' <summary>
        ''' 学生学号。
        ''' </summary>
        ''' <remarks></remarks>
        Dim StudentID As String
        ''' <summary>
        ''' 提交日期。
        ''' </summary>
        ''' <remarks></remarks>
        Dim LogDate As OJLogDate
        ''' <summary>
        ''' 提交时间。
        ''' </summary>
        ''' <remarks></remarks>
        Dim LogTime As OJLogTime
        ''' <summary>
        ''' 转化为标准格式的提交日期(仅含提交日期)。
        ''' </summary>
        ''' <remarks></remarks>
        Dim DateSubmit As Date
        ''' <summary>
        ''' 转化为标准格式的提交时间(含提交日期与提交时间)。
        ''' </summary>
        ''' <remarks></remarks>
        Dim TimeSubmit As Date
        ''' <summary>
        ''' 提交时处于星期几，1表示星期日。
        ''' </summary>
        ''' <remarks></remarks>
        Dim WeekdaySubmit As Integer
        ''' <summary>
        ''' 评测是否通过(AC)。
        ''' </summary>
        ''' <remarks></remarks>
        Dim IsPassed As Boolean
        ''' <summary>
        ''' 解析日志时是否出错导致失败。
        ''' </summary>
        ''' <remarks></remarks>
        Dim IsParseFailed As Boolean
        ''' <summary>
        ''' 将已结构化的OJ日志转化为单行文本。
        ''' </summary>
        ''' <returns>转化为单行文本的OJ日志，格式应与标准格式相同。</returns>
        ''' <remarks></remarks>
        Public Overrides Function ToString() As String
            Dim Temp As String
            Temp = TimeSubmit.ToString() & "," & StudentID & "," & TestID & "," & ProblemID & "," & IIf(IsPassed, "1", "0")
            Return Temp
        End Function
    End Structure
    ''' <summary>
    ''' OJ日志总数。
    ''' </summary>
    ''' <remarks></remarks>
    Public OJLogCount As Integer
    ''' <summary>
    ''' 解析OJ日志并转化为结构化数据。
    ''' </summary>
    ''' <param name="LogLine">单行OJ日志。</param>
    ''' <returns>结构化的日志数据。</returns>
    ''' <remarks></remarks>
    Public Function ParseLog(LogLine As String) As OJLog
        Dim LogLineArray() As String
        Dim LogDateTime() As String
        Dim LogDay() As String
        Dim LogTime() As String
        Dim Temp As OJLog = New OJLog
        With Temp
            .IsPassed = False
            .StudentID = ""
            .ProblemID = ""
            .TestID = 0
            .LogDate.Day = 1
            .LogDate.Month = 1
            .LogDate.Year = 1000
            .LogTime.Hour = 0
            .LogTime.Minute = 0
            .LogTime.Second = 0
            .DateSubmit = New Date(1000, 1, 1)
            .TimeSubmit = New Date(1000, 1, 1, 0, 0, 0)
            .WeekdaySubmit = Weekday(.DateSubmit, FirstDayOfWeek.Sunday)
            .IsParseFailed = True
        End With
        Try
            LogLineArray = Split(LogLine, ",")
            If LogLineArray.Length <> 5 Then
                Return Temp
            End If
            LogDateTime = Split(LogLineArray(0), " ")
            If LogDateTime.Length <> 2 Then
                Return Temp
            End If
            LogDay = Split(LogDateTime(0), "-")
            If LogDay.Length <> 3 Then
                Return Temp
            End If
            LogTime = Split(LogDateTime(1), ":")
            If LogTime.Length <> 3 Then
                Return Temp
            End If
            Temp.IsParseFailed = False
            With Temp
                .StudentID = LogLineArray(1)
                .TestID = LogLineArray(2)
                .ProblemID = LogLineArray(3)
                .LogDate.Year = Int(LogDay(0))
                .LogDate.Month = Int(LogDay(1))
                .LogDate.Day = Int(LogDay(2))
                .LogTime.Hour = Int(LogTime(0))
                .LogTime.Minute = Int(LogTime(1))
                .LogTime.Second = Int(LogTime(2))
                If .LogDate.Month > 12 Then
                    .LogDate.Month = 12
                End If
                If .LogDate.Month < 1 Then
                    .LogDate.Month = 1
                End If
                If .LogDate.Day > Date.DaysInMonth(.LogDate.Year, .LogDate.Month) Then
                    .LogDate.Day = Date.DaysInMonth(.LogDate.Year, .LogDate.Month)
                End If
                If .LogDate.Day < 0 Then
                    .LogDate.Day = 1
                End If
                If .LogTime.Hour >= 24 Then
                    .LogTime.Hour = 23
                End If
                If .LogTime.Hour < 0 Then
                    .LogTime.Hour = 0
                End If
                If .LogTime.Minute >= 60 Then
                    .LogTime.Minute = 59
                End If
                If .LogTime.Minute < 0 Then
                    .LogTime.Minute = 0
                End If
                If .LogTime.Second >= 60 Then
                    .LogTime.Second = 59
                End If
                If .LogTime.Second < 0 Then
                    .LogTime.Second = 0
                End If
                .DateSubmit = New Date(.LogDate.Year, .LogDate.Month, .LogDate.Day)
                .TimeSubmit = New Date(.LogDate.Year, .LogDate.Month, .LogDate.Day, .LogTime.Hour, .LogTime.Minute, .LogTime.Second)
                .WeekdaySubmit = Weekday(.DateSubmit, FirstDayOfWeek.Sunday)
                If LogLineArray(4) = "1" Then
                    .IsPassed = True
                Else
                    .IsPassed = False
                End If
                .IsParseFailed = False
            End With
        Catch ex As Exception
            With Temp
                .IsPassed = False
                .StudentID = ""
                .TestID = 0
                .ProblemID = ""
                .LogDate.Day = 1
                .LogDate.Month = 1
                .LogDate.Year = 1000
                .LogTime.Hour = 0
                .LogTime.Minute = 0
                .LogTime.Second = 0
                .DateSubmit = New Date(1000, 1, 1)
                .TimeSubmit = New Date(1000, 1, 1, 0, 0, 0)
                .WeekdaySubmit = Weekday(.DateSubmit, FirstDayOfWeek.Sunday)
                .IsParseFailed = True
            End With
        End Try
        Return Temp
    End Function
End Module
