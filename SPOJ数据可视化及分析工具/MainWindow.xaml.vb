Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Windows.Controls
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Windows.Forms
Imports System.Windows.Window
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports System.IO
Imports MathWorks.MATLAB.NET.Arrays
Imports MathWorks.MATLAB.NET.Utility
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Class MainWindow
    Dim UpdateProgress As Dialogs.ProgressDialogController
    Dim Updater As Process = New Process
    Dim IsUpdateFinished As Boolean = False
    Dim EmptyList As New List(Of String)
    Dim IsUserSpecifiedAnalyzeStartDateEqualsLogStartDate As Boolean = True
    Dim IsUserSpecifiedAnalyzeEndDateEqualsLogEndDate As Boolean = True
    Dim IsMCRFailedStageStudentTotalSubmitByDayLnFit As Boolean = False
    Dim IsMCRFailedStageProblemParticipantFit As Boolean = False
    Dim IsMCRFailedStageProblemClust As Boolean = False
    Dim IsMCRFailedStageStudentClust As Boolean = False
    Dim IsSkippedStageStudentTotalSubmitByDayLnFit As Boolean = False
    Dim IsSkippedStageProblemParticipantFit As Boolean = False
    Dim IsSkippedStageProblemClust As Boolean = False
    Dim IsSkippedStageStudentClust As Boolean = False
    Dim DataSourcePath As String

    Private Sub RefreshList()
        lstProblemList.ItemsSource = EmptyList
        lstStudentList.ItemsSource = EmptyList
        lstStudentList.ItemsSource = StudentList
        lstProblemList.ItemsSource = ProblemList
    End Sub
    Private Function LinearCalculate(Slope As Double, Intercept As Double, XValue As Double) As Double
        Debug.Print("y=" & Slope & "x+" & Intercept & vbCrLf & "x=" & XValue)
        Return Slope * XValue + Intercept
    End Function
    Private Function ExponentialCalculate(Coefficient As Double, Exponent As Double, Constant As Double, XValue As Double) As Double
        Debug.Print("y=" & Coefficient & "e^(" & Exponent & "x)+" & Constant & vbCrLf & "x=" & XValue)
        Return Coefficient * Math.Exp(Exponent * XValue) + Constant
    End Function
    Private Async Sub mnuUpdate_Click(sender As Object, e As RoutedEventArgs) Handles mnuUpdate.Click
        '检查所需文件是否存在，执行灾难恢复。
        GenerateCurrentDirectory()
        Dim CurrentPath As String = GetCurrentDirectory()
        Dim FolderBrowser As New CommonOpenFileDialog
        Dim IsProblemListMissing As Boolean = False
        Dim IsUserListMissing As Boolean = False
        Dim IsTestListMissing As Boolean = False
RequestDataSource:
        With FolderBrowser
            .Title = "请指定 SPOJ 日志文件的位置"
            .IsFolderPicker = True
            .DefaultDirectory = CurrentPath
            .AllowNonFileSystemItems = False
        End With
        If FolderBrowser.ShowDialog() = CommonFileDialogResult.Ok Then
            DataSourcePath = FolderBrowser.FileName
            If DataSourcePath(DataSourcePath.Length - 1) <> "\" Then
                DataSourcePath = DataSourcePath & "\"
            End If
            '开始分析OJ日志。
            IsUpdateFinished = True
            Dim AnalyzeProgress As Dialogs.ProgressDialogController
            '校验数据。
            AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在校验数据", "正在对 OnlineJudge 数据进行校验，这可能需要一点点的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            If Not DirectoryExists(DataSourcePath) Then
                Await AnalyzeProgress.CloseAsync()
                Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
                Exit Sub
            End If
            If Not FileExists(DataSourcePath & "Log.csv") Then
                Await AnalyzeProgress.CloseAsync()
                Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
                Exit Sub
            End If
            If Not FileExists(DataSourcePath & "User.csv") Then
                IsUserListMissing = True
            End If
            If Not FileExists(DataSourcePath & "Problem.csv") Then
                IsProblemListMissing = True
            End If
            If Not FileExists(DataSourcePath & "Test.csv") Then
                IsTestListMissing = True
            End If
            Await AnalyzeProgress.CloseAsync()
            If IsUserListMissing Or IsProblemListMissing Or IsTestListMissing Then
                Await ShowMessageAsync("数据源扩展缺失", "引用的数据源缺失部分扩展数据文件，分析仍将继续执行，但部分分析数据将依赖程序内置的计算算法生成，这可能导致少数分析结果不可靠。" & vbCrLf & vbCrLf & "下列扩展数据文件隶属于数据源：" & vbCrLf & _
                                       DataSourcePath & "User.csv - 描述了 OJ 系统上的用户。" & IIf(IsUserListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                       DataSourcePath & "Problem.csv - 描述了 OJ 系统上的题目。" & IIf(IsProblemListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                       DataSourcePath & "Test.csv - 描述了 OJ 系统上的题目集。" & IIf(IsTestListMissing, "该数据源扩展不存在，题目的创建时间数据可能不准确。", "该数据源扩展存在。") _
                                       )
            End If

            '初始化
            StudentList.Clear()
            StudentDictionary.Clear()
            ProblemList.Clear()
            ProblemDictionary.Clear()
            TestDictionary.Clear()
            lstStudentList.ItemsSource = EmptyList
            lstProblemList.ItemsSource = EmptyList
            txtStudentSearch.Text = "搜索学生学号"
            txtStudentSearch.Foreground = SystemColors.ScrollBarBrush
            txtProblemSearch.Text = "搜索题目编号"
            txtProblemSearch.Foreground = SystemColors.ScrollBarBrush
            icoStudentLink.Foreground = SystemColors.GrayTextBrush
            btnStudentLink.IsEnabled = False
            btnStudentLink.Cursor = System.Windows.Input.Cursors.Arrow
            icoProblemLink.Foreground = SystemColors.GrayTextBrush
            btnProblemLink.IsEnabled = False
            btnProblemLink.Cursor = System.Windows.Input.Cursors.Arrow
            txtStudentID.Text = ""
            txtStudentClustResult.Text = ""
            txtStudentSubmitCount.Text = ""
            txtStudentACCount.Text = ""
            txtStudentACRate.Text = ""
            txtStudentSubmitCountOnWorkdayAM.Text = ""
            txtStudentSubmitCountOnWorkdayPM.Text = ""
            txtStudentSubmitCountOnRestdayAM.Text = ""
            txtStudentSubmitCountOnRestdayPM.Text = ""
            txtFittingAC.Text = ""
            txtFittingK_Kb.Text = ""
            txtFittingR_Stb.Text = ""
            txtProblemID.Text = ""
            txtProblemCreateTime.Text = ""
            txtProblemParticipantCount.Text = ""
            txtProblemSubmitCount.Text = ""
            txtProblemACCount.Text = ""
            txtProblemACRate.Text = ""
            txtEffortValue_Jq.Text = ""
            txtParticipateValue_Eq.Text = ""
            txtStartDate.Text = ""
            txtEndDate.Text = ""
            txtUserSpecifiedAnalyzeStartDate.Text = ""
            txtUserSpecifiedAnalyzeEndDate.Text = ""
            chkCustomAnalyzeEndDate.IsChecked = False
            chkCustomAnalyzeStartDate.IsChecked = False
            dtpUserSpecifiedAnalyzeEndDate.IsEnabled = False
            dtpUserSpecifiedAnalyzeStartDate.IsEnabled = False
            chartStudentACRate.Visibility = Windows.Visibility.Hidden
            chartStudentSubmitByDay.Visibility = Windows.Visibility.Hidden
            chartStudentTotalSubmitByDayLn.Visibility = Windows.Visibility.Hidden
            chartStudentTotalSubmitByTime.Visibility = Windows.Visibility.Hidden
            chartProblemACRate.Visibility = Windows.Visibility.Hidden
            chartProblemSubmitByDay.Visibility = Windows.Visibility.Hidden
            chartProblemParticipant.Visibility = Windows.Visibility.Hidden
            chartSubmitCountByTime.Visibility = Windows.Visibility.Hidden
            chartNewProblemCountByDay.Visibility = Windows.Visibility.Hidden
            pieStudentACRate.ItemsSource = Nothing
            DoEvents()
            pieStudentACRate.Refresh()
            DoEvents()
            linStudentSubmitByDay.ItemsSource = Nothing
            DoEvents()
            linStudentSubmitByDay.Refresh()
            DoEvents()
            sctStudentTotalSubmitByDayLn.ItemsSource = Nothing
            DoEvents()
            sctStudentTotalSubmitByDayLn.Refresh()
            DoEvents()
            linStudentTotalSubmitByDayLnFit.ItemsSource = Nothing
            DoEvents()
            linStudentTotalSubmitByDayLnFit.Refresh()
            DoEvents()
            colStudentTotalSubmitByTime.ItemsSource = Nothing
            DoEvents()
            colStudentTotalSubmitByTime.Refresh()
            DoEvents()
            pieProblemACRate.ItemsSource = Nothing
            DoEvents()
            pieProblemACRate.Refresh()
            DoEvents()
            linProblemSubmitByDay.ItemsSource = Nothing
            DoEvents()
            linProblemSubmitByDay.Refresh()
            DoEvents()
            sctProblemParticipant.ItemsSource = Nothing
            DoEvents()
            sctProblemParticipant.Refresh()
            DoEvents()
            linProblemParticipantFit.ItemsSource = Nothing
            DoEvents()
            linProblemParticipantFit.Refresh()
            DoEvents()
            linNewProblemCountByDay.ItemsSource = Nothing
            DoEvents()
            linNewProblemCountByDay.Refresh()
            DoEvents()
            colSubmitCountByTime.ItemsSource = Nothing
            DoEvents()
            colSubmitCountByTime.Refresh()
            DoEvents()

            Dim OJLogFileReader As System.IO.StreamReader
            Dim OJLogLine As String
            Dim OJLogTemp As New OJLog
            Dim OJTestLogLine As String
            Dim OJTestTemp As New OJTestInfo
            AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在分析数据", "正在对 OnlineJudge 数据进行分析，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            '预存测试数据
            If Not IsTestListMissing Then
                Try
                    OJLogFileReader = New IO.StreamReader(DataSourcePath & "Test.csv")
                    While Not OJLogFileReader.EndOfStream
                        OJTestLogLine = OJLogFileReader.ReadLine()
                        OJTestTemp = ParseOJTestLogLine(OJTestLogLine)
                        If OJTestTemp.IsParseFailed Then
                            Continue While
                        End If
                        If TestDictionary.ContainsKey(OJTestTemp.TestID) Then
                            If OJTestTemp.BeginDate < TestDictionary(OJTestTemp.TestID).BeginDate Then
                                TestDictionary(OJTestTemp.TestID).BeginDate = OJTestTemp.BeginDate
                            End If
                            If OJTestTemp.EndDate > TestDictionary(OJTestTemp.TestID).EndDate Then
                                TestDictionary(OJTestTemp.TestID).EndDate = OJTestTemp.EndDate
                            End If
                        Else
                            TestDictionary.Add(OJTestTemp.TestID, OJTestTemp)
                        End If
                    End While
                    If TestDictionary.Count = 0 Then
                        IsTestListMissing = True
                    End If
                    OJLogFileReader.Close()
                Catch ex As Exception
                    IsTestListMissing = True
                End Try
            End If
            OJLogFileReader = New IO.StreamReader(DataSourcePath & "Log.csv")
            Dim StudentTemp As OJStudentInfo
            Dim ProblemTemp As OJProblemInfo
            Dim ProblemParticipantCheck As New Dictionary(Of String, Dictionary(Of String, Boolean))
            Dim DateList As New List(Of Date)
            StudentList.Clear()
            StudentDictionary.Clear()
            ProblemList.Clear()
            ProblemDictionary.Clear()
            OJSysInfo.NewProblemCount.Clear()
            Dim i As Integer
            For i = 0 To 23
                OJSysInfo.SubmitCountByHour(i) = 0
            Next
            While Not OJLogFileReader.EndOfStream
                OJLogLine = OJLogFileReader.ReadLine()
                OJLogTemp = ParseLog(OJLogLine)
                If OJLogTemp.IsParseFailed Then
                    Continue While
                End If
                '此处加入详细分析代码
                '学生数据
                If StudentDictionary.ContainsKey(OJLogTemp.StudentID) Then
                    With StudentDictionary(OJLogTemp.StudentID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                        If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnWorkdayAM += 1
                            Else
                                .SubmitCountOnWorkdayPM += 1
                            End If
                        Else
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnRestdayAM += 1
                            Else
                                .SubmitCountOnRestdayPM += 1
                            End If
                        End If
                        '拟合计算

                    End With
                Else
                    StudentTemp = New OJStudentInfo(OJLogTemp.StudentID)
                    StudentDictionary(OJLogTemp.StudentID) = StudentTemp
                    StudentList.Add(OJLogTemp.StudentID)
                    With StudentDictionary(OJLogTemp.StudentID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                        If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnWorkdayAM += 1
                            Else
                                .SubmitCountOnWorkdayPM += 1
                            End If
                        Else
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnRestdayAM += 1
                            Else
                                .SubmitCountOnRestdayPM += 1
                            End If
                        End If
                    End With
                End If
                '题目数据
                If Not ProblemParticipantCheck.ContainsKey(OJLogTemp.ProblemID) Then
                    ProblemParticipantCheck.Add(OJLogTemp.ProblemID, New Dictionary(Of String, Boolean))
                End If
                If ProblemDictionary.ContainsKey(OJLogTemp.ProblemID) Then
                    With ProblemDictionary(OJLogTemp.ProblemID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                            .ParticipantCount += 1
                            ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        If IsTestListMissing Then
                            '提高鲁棒性：当一个题目出现更早的提交时更新数据
                            If OJLogTemp.DateSubmit < .CreateTime Then
                                If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                                    If OJSysInfo.NewProblemCount(.CreateTime) > 0 Then
                                        OJSysInfo.NewProblemCount(.CreateTime) -= 1
                                    End If
                                End If
                                .CreateTime = OJLogTemp.DateSubmit
                                If OJSysInfo.NewProblemCount.ContainsKey(OJLogTemp.DateSubmit) Then
                                    OJSysInfo.NewProblemCount(OJLogTemp.DateSubmit) += 1
                                Else
                                    OJSysInfo.NewProblemCount.Add(OJLogTemp.DateSubmit, 1)
                                End If
                            End If
                        End If
                    End With
                Else
                    ProblemTemp = New OJProblemInfo(OJLogTemp.ProblemID)
                    ProblemDictionary(OJLogTemp.ProblemID) = ProblemTemp
                    ProblemList.Add(OJLogTemp.ProblemID)
                    With ProblemDictionary(OJLogTemp.ProblemID)
                        .SubmitCount += 1
                        If IsTestListMissing Then
                            .CreateTime = OJLogTemp.DateSubmit
                        Else
                            .ProblemTaskSequenceIndex = OJLogTemp.TestID
                            .CreateTime = TestDictionary(.ProblemTaskSequenceIndex).BeginDate
                        End If
                        If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                            OJSysInfo.NewProblemCount(.CreateTime) += 1
                        Else
                            OJSysInfo.NewProblemCount.Add(.CreateTime, 1)
                        End If
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                            .ParticipantCount += 1
                            ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                    End With
                End If
                '系统数据
                DateList.Add(OJLogTemp.DateSubmit)
                OJSysInfo.SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                DoEvents()
            End While
            OJLogFileReader.Close()
            ProblemParticipantCheck.Clear()
            OJSysInfo.EndDate = DateList.Max()
            OJSysInfo.StartDate = DateList.Min()
            '配置默认的分析起讫日期
            UserSpecifiedAnalyzeStartDate = OJSysInfo.StartDate
            UserSpecifiedAnalyzeEndDate = OJSysInfo.EndDate
            StudentList.Sort()
            ProblemList.Sort()
            '拟合计算
            '学生拟合
            If StudentList.Count >= 2 Then
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .FittingAC = IIf(.SubmitCount > 0, .ACCount + .ACCount / .SubmitCount, 0)
                        Dim DateLoop As New Date
                        DateLoop = UserSpecifiedAnalyzeStartDate
                        Dim DayLn() As Double
                        Dim SubmitAdjustedTotal() As Double
                        ReDim DayLn((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                        ReDim SubmitAdjustedTotal((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                        Dim StudentTotalSubmitSum As Integer = 0
                        Dim TotalProblems As Integer = 0
                        Dim Index = 0
                        While DateLoop < UserSpecifiedAnalyzeEndDate
                            If OJSysInfo.NewProblemCount.ContainsKey(DateLoop) Then
                                TotalProblems += OJSysInfo.NewProblemCount(DateLoop)
                            End If
                            If .SubmitCountByDay.ContainsKey(DateLoop) Then
                                StudentTotalSubmitSum += .SubmitCountByDay(DateLoop)
                            End If
                            DayLn(Index) = Math.Log((DateLoop - UserSpecifiedAnalyzeStartDate).Days + 1)
                            SubmitAdjustedTotal(Index) = StudentTotalSubmitSum - TotalProblems
                            DateLoop = DateLoop.AddDays(1)
                            Index += 1
                        End While
                        'Dim LinearFitResult As New Tuple(Of Double, Double)(0, 0)
                        'LinearFitResult = MathNet.Numerics.LinearRegression.SimpleRegression.Fit(DayLn, SubmitAdjustedTotal)
                        '.FittingB = LinearFitResult.Item1
                        '.FittingK_Kb = LinearFitResult.Item2
                        '.FittingR_Stb = MathNet.Numerics.Statistics.Correlation.Pearson(DayLn, SubmitAdjustedTotal) ^ 2
                        Try
                            Dim LinearFitReturn(2, 3) As Double
                            Dim LinearFitter As New LinearFit.Class1
                            Dim LinearFitX As New MWNumericArray(DayLn)
                            Dim LinearFitY As New MWNumericArray(SubmitAdjustedTotal)
                            LinearFitX = DayLn
                            LinearFitY = SubmitAdjustedTotal
                            LinearFitReturn = LinearFitter.LinearFit(LinearFitX, LinearFitY).ToArray()
                            LinearFitter.Dispose()
                            LinearFitX.Dispose()
                            LinearFitY.Dispose()
                            .FittingK_Kb = LinearFitReturn(0, 0)
                            .FittingR_Stb = LinearFitReturn(0, 1)
                            .FittingB = LinearFitReturn(0, 2)
                        Catch ex As Exception
                            IsMCRFailedStageStudentTotalSubmitByDayLnFit = True
                            .FittingB = 0
                            .FittingK_Kb = 0
                            .FittingR_Stb = 0
                        End Try
                    End With
                    DoEvents()
                Next
            Else
                IsSkippedStageStudentTotalSubmitByDayLnFit = True
            End If
            '题目拟合
            '产生题目参与人数-作业次序序列
            Dim ProblemCreateTimeList As New List(Of Date)
            Dim FirstProblemCreateTime As New Date
            Dim TaskSequence() As Integer
            Dim ParticipantCount() As Integer
            Dim ExponentialFittingA As Double = 0
            Dim ExponentialFittingB As Double = 0
            Dim ExponentialFittingC As Double = 0
            Dim ExponentialFitReturn(1, 3) As Double
            For i = 0 To ProblemList.Count - 1
                ProblemCreateTimeList.Add(ProblemDictionary(ProblemList(i)).CreateTime)
            Next
            FirstProblemCreateTime = ProblemCreateTimeList.Min()
            ReDim TaskSequence(ProblemList.Count - 1)
            ReDim ParticipantCount(ProblemList.Count - 1)
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    If IsTestListMissing Then
                        .ProblemTaskSequenceIndex = Math.Round((.CreateTime - FirstProblemCreateTime).Days / 7) + 1
                    End If
                    TaskSequence(i) = .ProblemTaskSequenceIndex
                    ParticipantCount(i) = .ParticipantCount
                End With
                DoEvents()
            Next
            If ProblemList.Count >= 3 Then
                Try
                    Dim ExponentialFitter As New e_fit.Class1
                    Dim ExponentialFitX As New MWNumericArray()
                    Dim ExponentialFitY As New MWNumericArray()
                    ExponentialFitX = TaskSequence
                    ExponentialFitY = ParticipantCount
                    ExponentialFitReturn = ExponentialFitter.e_fit(ExponentialFitX, ExponentialFitY).ToArray()
                    ExponentialFitter.Dispose()
                    ExponentialFitX.Dispose()
                    ExponentialFitY.Dispose()
                    ExponentialFittingA = -ExponentialFitReturn(0, 0)
                    ExponentialFittingB = ExponentialFitReturn(0, 1)
                    ExponentialFittingC = ExponentialFitReturn(0, 2)
                Catch ex As Exception
                    IsMCRFailedStageProblemParticipantFit = True
                    ExponentialFittingA = 0
                    ExponentialFittingB = 0
                    ExponentialFittingC = 0
                End Try
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                        .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                    End With
                    DoEvents()
                Next
            Else
                IsSkippedStageProblemParticipantFit = True
                ExponentialFittingA = 0
                ExponentialFittingB = 0
                ExponentialFittingC = 0
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                        .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                    End With
                    DoEvents()
                Next
            End If
            '学生聚类分析
            If StudentList.Count >= 4 Then
                Dim StudentAcStbKbMatrix(,) As Double
                Dim StudentClustResult(,) As Double
                Dim AcAverage As Double = 0
                Dim AcStandardDivision As Double = 0
                Dim StbAvarage As Double = 0
                Dim StbStandardDivision As Double = 0
                Dim KbAverage As Double = 0
                Dim KbStandardDivision As Double = 0
                ReDim StudentAcStbKbMatrix(StudentList.Count - 1, 2)
                ReDim StudentClustResult(StudentList.Count - 1, 1)
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        AcAverage += .FittingAC
                        StbAvarage += .FittingR_Stb
                        KbAverage += .FittingK_Kb
                    End With
                    DoEvents()
                Next
                AcAverage /= StudentList.Count
                StbAvarage /= StudentList.Count
                KbAverage /= StudentList.Count
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        AcStandardDivision += (.FittingAC - AcAverage) ^ 2
                        StbStandardDivision += (.FittingR_Stb - StbAvarage) ^ 2
                        KbStandardDivision += (.FittingK_Kb - KbAverage) ^ 2
                    End With
                    DoEvents()
                Next
                AcStandardDivision /= StudentList.Count
                StbStandardDivision /= StudentList.Count
                KbStandardDivision /= StudentList.Count
                AcStandardDivision = Math.Sqrt(AcStandardDivision)
                StbStandardDivision = Math.Sqrt(StbStandardDivision)
                KbStandardDivision = Math.Sqrt(KbStandardDivision)
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        StudentAcStbKbMatrix(i, 0) = (.FittingAC - AcAverage) / AcStandardDivision
                        StudentAcStbKbMatrix(i, 1) = (.FittingR_Stb - StbAvarage) / StbStandardDivision
                        StudentAcStbKbMatrix(i, 2) = (.FittingK_Kb - KbAverage) / KbStandardDivision
                    End With
                    DoEvents()
                Next
                Try
                    Dim StudentClustExec As New cluster.Class1
                    StudentClustResult = StudentClustExec.cluster(New MWNumericArray(StudentAcStbKbMatrix), 4).ToArray()
                    StudentClustExec.Dispose()
                    For i = 0 To StudentList.Count - 1
                        With StudentDictionary(StudentList(i))
                            .ClustResult = StudentClustResult(i, 0)
                        End With
                        DoEvents()
                    Next
                    StudentClustResultMapping.Clear()
                    StudentClustResultMapping.Add(0, "出错或被跳过")
                    StudentClustResultMapping.Add(1, "未分级，点击以分级")
                    StudentClustResultMapping.Add(2, "未分级，点击以分级")
                    StudentClustResultMapping.Add(3, "未分级，点击以分级")
                    StudentClustResultMapping.Add(4, "未分级，点击以分级")
                Catch ex As Exception
                    IsMCRFailedStageStudentClust = True
                    For i = 0 To StudentList.Count - 1
                        With StudentDictionary(StudentList(i))
                            .ClustResult = 0
                        End With
                        DoEvents()
                    Next
                    StudentClustResultMapping.Clear()
                    StudentClustResultMapping.Add(0, "出错或被跳过")
                    StudentClustResultMapping.Add(1, "未分级，点击以分级")
                    StudentClustResultMapping.Add(2, "未分级，点击以分级")
                    StudentClustResultMapping.Add(3, "未分级，点击以分级")
                    StudentClustResultMapping.Add(4, "未分级，点击以分级")
                End Try
            Else
                IsSkippedStageStudentClust = True
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .ClustResult = 0
                    End With
                    DoEvents()
                Next
                StudentClustResultMapping.Clear()
                StudentClustResultMapping.Add(0, "出错或被跳过")
                StudentClustResultMapping.Add(1, "未分级，点击以分级")
                StudentClustResultMapping.Add(2, "未分级，点击以分级")
                StudentClustResultMapping.Add(3, "未分级，点击以分级")
                StudentClustResultMapping.Add(4, "未分级，点击以分级")
            End If
            '题目聚类分析
            '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
            If ProblemList.Count >= 4 Then
                Dim ProblemEqJqMatrix(,) As Double
                Dim ProblemClustResult(,) As Double
                Dim EqAverage As Double = 0
                Dim EqStandardDivision As Double = 0
                Dim JqAverage As Double = 0
                Dim JqStandardDivision As Double = 0
                ReDim ProblemEqJqMatrix(ProblemList.Count - 1, 1)
                ReDim ProblemClustResult(ProblemList.Count - 1, 1)
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        EqAverage += .ParticipateValue_Eq
                        JqAverage += .EffortValue_Jq
                    End With
                    DoEvents()
                Next
                EqAverage /= ProblemList.Count
                JqAverage /= ProblemList.Count
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        EqStandardDivision += (.ParticipateValue_Eq - EqAverage) ^ 2
                        JqStandardDivision += (.EffortValue_Jq - JqAverage) ^ 2
                    End With
                    DoEvents()
                Next
                EqStandardDivision /= ProblemList.Count
                JqStandardDivision /= ProblemList.Count
                EqStandardDivision = Math.Sqrt(EqStandardDivision)
                JqStandardDivision = Math.Sqrt(JqStandardDivision)
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        ProblemEqJqMatrix(i, 0) = (.EffortValue_Jq - JqAverage) / JqStandardDivision
                        ProblemEqJqMatrix(i, 1) = (.ParticipateValue_Eq - EqAverage) / EqStandardDivision
                    End With
                    DoEvents()
                Next
                Try
                    Dim ProblemClustExec As New cluster.Class1
                    ProblemClustResult = ProblemClustExec.cluster(New MWNumericArray(ProblemEqJqMatrix), 4).ToArray()
                    ProblemClustExec.Dispose()
                    For i = 0 To ProblemList.Count - 1
                        With ProblemDictionary(ProblemList(i))
                            .ClustResult = ProblemClustResult(i, 0)
                        End With
                        DoEvents()
                    Next
                Catch ex As Exception
                    IsMCRFailedStageProblemClust = True
                    For i = 0 To ProblemList.Count - 1
                        ProblemDictionary(ProblemList(i)).ClustResult = 0
                        DoEvents()
                    Next
                End Try
            Else
                IsSkippedStageProblemClust = True
                For i = 0 To ProblemList.Count - 1
                    ProblemDictionary(ProblemList(i)).ClustResult = 0
                    DoEvents()
                Next
            End If
            '呈现列表
            lstStudentList.ItemsSource = StudentList
            lstProblemList.ItemsSource = ProblemList
            txtStartDate.Text = OJSysInfo.StartDate.ToLongDateString()
            txtEndDate.Text = OJSysInfo.EndDate.ToLongDateString()
            txtUserSpecifiedAnalyzeStartDate.Text = UserSpecifiedAnalyzeStartDate.ToLongDateString()
            txtUserSpecifiedAnalyzeEndDate.Text = UserSpecifiedAnalyzeEndDate.ToLongDateString()
            dtpUserSpecifiedAnalyzeStartDate.DisplayDateStart = OJSysInfo.StartDate
            dtpUserSpecifiedAnalyzeStartDate.DisplayDateEnd = OJSysInfo.EndDate
            dtpUserSpecifiedAnalyzeEndDate.DisplayDateStart = OJSysInfo.StartDate
            dtpUserSpecifiedAnalyzeEndDate.DisplayDateEnd = OJSysInfo.EndDate
            dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
            dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
            '题目参与人数与拟合图表
            If Not IsSkippedStageProblemParticipantFit Then
                If Not IsMCRFailedStageProblemParticipantFit Then
                    Dim ProblemParticipantDataSource As New List(Of KeyValuePair(Of Integer, Integer))
                    Dim ProblemParticipantFitDataSource As New List(Of KeyValuePair(Of Integer, Double))
                    For i = 0 To ProblemList.Count - 1
                        ProblemParticipantDataSource.Add(New KeyValuePair(Of Integer, Integer)(TaskSequence(i), ParticipantCount(i)))
                        ProblemParticipantFitDataSource.Add(New KeyValuePair(Of Integer, Double)(TaskSequence(i), ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, TaskSequence(i))))
                        DoEvents()
                    Next
                    sctProblemParticipant.ItemsSource = ProblemParticipantDataSource
                    linProblemParticipantFit.ItemsSource = ProblemParticipantFitDataSource
                    chartProblemParticipant.Visibility = Windows.Visibility.Visible
                End If
            End If
            '分时段总提交数图表
            Dim SubmitCountByTimeDataSource As New List(Of KeyValuePair(Of String, Integer))
            For i = 0 To 23
                SubmitCountByTimeDataSource.Add(New KeyValuePair(Of String, Integer)(i.ToString("00") & ":00" & vbCrLf & "~" & vbCrLf & (i + 1).ToString("00") & ":00", OJSysInfo.SubmitCountByHour(i)))
            Next
            colSubmitCountByTime.ItemsSource = SubmitCountByTimeDataSource
            chartSubmitCountByTime.Visibility = Windows.Visibility.Visible
            '题目发布曲线图表
            Dim j As Date
            j = UserSpecifiedAnalyzeStartDate
            Dim NewProblemCountByDayDataSource As New List(Of KeyValuePair(Of Date, Integer))
            While j <= UserSpecifiedAnalyzeEndDate
                If OJSysInfo.NewProblemCount.ContainsKey(j) Then
                    NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, OJSysInfo.NewProblemCount(j)))
                Else
                    NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, 0))
                End If
                j = j.AddDays(1)
            End While
            linNewProblemCountByDay.ItemsSource = NewProblemCountByDayDataSource
            chartNewProblemCountByDay.Visibility = Windows.Visibility.Visible
            '题目聚类分析结果
            '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
            '聚类图表的四类依次序分为:
            '低Jq低Eq，低Jq高Eq，高Jq低Eq，高Jq高Eq
            If Not IsSkippedStageProblemClust Then
                If Not IsMCRFailedStageProblemClust Then
                    Dim ScatterProblemCluster1DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster2DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster3DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster4DataSource As New List(Of KeyValuePair(Of Double, Double))
                    For i = 0 To ProblemList.Count - 1
                        With ProblemDictionary(ProblemList(i))
                            If .ClustResult = 1 Then
                                ScatterProblemCluster1DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            ElseIf .ClustResult = 2 Then
                                ScatterProblemCluster2DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            ElseIf .ClustResult = 3 Then
                                ScatterProblemCluster3DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            Else
                                ScatterProblemCluster4DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            End If
                        End With
                        DoEvents()
                    Next
                    sctProblemCluster1.ItemsSource = ScatterProblemCluster1DataSource
                    sctProblemCluster2.ItemsSource = ScatterProblemCluster2DataSource
                    sctProblemCluster3.ItemsSource = ScatterProblemCluster3DataSource
                    sctProblemCluster4.ItemsSource = ScatterProblemCluster4DataSource
                    chartProblemCluster.Visibility = Windows.Visibility.Visible
                End If
            End If
            Await AnalyzeProgress.CloseAsync()
            If IsMCRFailedStageStudentTotalSubmitByDayLnFit Or IsMCRFailedStageProblemParticipantFit Or IsMCRFailedStageStudentClust Or IsMCRFailedStageProblemClust Then
                Await Me.ShowMessageAsync("MATLAB Runtime发生问题", "本程序的运行依赖 MathWorks MATLAB Runtime R2018b，但是该组件可能发生问题或无法被调用。请前往 https://ww2.mathworks.cn/products/compiler/matlab-runtime.html 下载并安装适用于您的操作系统的 MathWorks MATLAB Runtime R2018b。如果您已经安装了 MathWorks MATLAB Runtime R2018b，您可能需要将其重新安装。存在边缘情况的数据集也可能引发此错误。" & vbCrLf & vbCrLf & "程序仍可继续运行，但部分功能可能发生异常。" & vbCrLf & vbCrLf & "以下功能依赖 MathWorks MATLAB Runtime R2018b: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsMCRFailedStageStudentTotalSubmitByDayLnFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsMCRFailedStageProblemParticipantFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsMCRFailedStageStudentClust, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsMCRFailedStageProblemClust, "执行时失败或发生错误。", "执行成功。"))
                IsMCRFailedStageStudentTotalSubmitByDayLnFit = False
                IsMCRFailedStageProblemParticipantFit = False
                IsMCRFailedStageStudentClust = False
                IsMCRFailedStageProblemClust = False
            End If
            If IsSkippedStageStudentTotalSubmitByDayLnFit Or IsSkippedStageProblemParticipantFit Or IsSkippedStageStudentClust Or IsSkippedStageProblemClust Then
                Await Me.ShowMessageAsync("一些分析步骤因数据集容量过小被跳过", "一部分拟合/聚类分析步骤被跳过，这可能是由数据集容量过小的问题导致的。" & vbCrLf & vbCrLf & "您可以尝试通过更新日志、扩大分析时间段长度等方式扩展数据集容量。" & vbCrLf & vbCrLf & "该问题会导致一部分图表或分析结果无法显示或显示的结果不正确。" & vbCrLf & vbCrLf & "以下分析流程可能已被跳过: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsSkippedStageStudentTotalSubmitByDayLnFit, "已被跳过。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsSkippedStageProblemParticipantFit, "已被跳过。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsSkippedStageStudentClust, "已被跳过。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsSkippedStageProblemClust, "已被跳过。", "执行成功。"))
                IsSkippedStageStudentTotalSubmitByDayLnFit = False
                IsSkippedStageProblemParticipantFit = False
                IsSkippedStageStudentClust = False
                IsSkippedStageProblemClust = False
            End If
        End If
    End Sub
    Private Async Sub Updater_Exit(sender As Object, e As EventArgs)
        Await UpdateProgress.CloseAsync()
        IsUpdateFinished = True
    End Sub

    Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        '检查所需文件是否存在，执行灾难恢复。
        GenerateCurrentDirectory()
        Dim CurrentPath As String = GetCurrentDirectory()
        Dim FolderBrowser As New CommonOpenFileDialog
        Dim IsProblemListMissing As Boolean = False
        Dim IsUserListMissing As Boolean = False
        Dim IsTestListMissing As Boolean = False
RequestDataSource:
        With FolderBrowser
            .Title = "请指定 SPOJ 日志文件的位置"
            .IsFolderPicker = True
            .DefaultDirectory = CurrentPath
            .AllowNonFileSystemItems = False
        End With
        If FolderBrowser.ShowDialog() = CommonFileDialogResult.Ok Then
            DataSourcePath = FolderBrowser.FileName
            If DataSourcePath(DataSourcePath.Length - 1) <> "\" Then
                DataSourcePath = DataSourcePath & "\"
            End If
            '开始分析OJ日志。
            IsUpdateFinished = True
            Dim AnalyzeProgress As Dialogs.ProgressDialogController
            '校验数据。
            AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在校验数据", "正在对 OnlineJudge 数据进行校验，这可能需要一点点的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            If Not DirectoryExists(DataSourcePath) Then
                Await AnalyzeProgress.CloseAsync()
                Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
                GoTo RequestDataSource
            End If
            If Not FileExists(DataSourcePath & "Log.csv") Then
                Await AnalyzeProgress.CloseAsync()
                Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
                GoTo RequestDataSource
            End If
            If Not FileExists(DataSourcePath & "User.csv") Then
                IsUserListMissing = True
            End If
            If Not FileExists(DataSourcePath & "Problem.csv") Then
                IsProblemListMissing = True
            End If
            If Not FileExists(DataSourcePath & "Test.csv") Then
                IsTestListMissing = True
            End If
            Await AnalyzeProgress.CloseAsync()
            If IsUserListMissing Or IsProblemListMissing Or IsTestListMissing Then
                Await ShowMessageAsync("数据源扩展缺失", "引用的数据源缺失部分扩展数据文件，分析仍将继续执行，但部分分析数据将依赖程序内置的计算算法生成，这可能导致少数分析结果不可靠。" & vbCrLf & vbCrLf & "下列扩展数据文件隶属于数据源：" & vbCrLf & _
                                       DataSourcePath & "User.csv - 描述了 OJ 系统上的用户。" & IIf(IsUserListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                       DataSourcePath & "Problem.csv - 描述了 OJ 系统上的题目。" & IIf(IsProblemListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                       DataSourcePath & "Test.csv - 描述了 OJ 系统上的题目集。" & IIf(IsTestListMissing, "该数据源扩展不存在，题目的创建时间数据可能不准确。", "该数据源扩展存在。") _
                                       )
            End If
            Dim OJLogFileReader As System.IO.StreamReader
            Dim OJLogLine As String
            Dim OJLogTemp As New OJLog
            Dim OJTestLogLine As String
            Dim OJTestTemp As New OJTestInfo
            AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在分析数据", "正在对 OnlineJudge 数据进行分析，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            '预存测试数据
            If Not IsTestListMissing Then
                Try
                    OJLogFileReader = New IO.StreamReader(DataSourcePath & "Test.csv")
                    While Not OJLogFileReader.EndOfStream
                        OJTestLogLine = OJLogFileReader.ReadLine()
                        OJTestTemp = ParseOJTestLogLine(OJTestLogLine)
                        If OJTestTemp.IsParseFailed Then
                            Continue While
                        End If
                        If TestDictionary.ContainsKey(OJTestTemp.TestID) Then
                            If OJTestTemp.BeginDate < TestDictionary(OJTestTemp.TestID).BeginDate Then
                                TestDictionary(OJTestTemp.TestID).BeginDate = OJTestTemp.BeginDate
                            End If
                            If OJTestTemp.EndDate > TestDictionary(OJTestTemp.TestID).EndDate Then
                                TestDictionary(OJTestTemp.TestID).EndDate = OJTestTemp.EndDate
                            End If
                        Else
                            TestDictionary.Add(OJTestTemp.TestID, OJTestTemp)
                        End If
                    End While
                    If TestDictionary.Count = 0 Then
                        IsTestListMissing = True
                    End If
                    OJLogFileReader.Close()
                Catch ex As Exception
                    IsTestListMissing = True
                End Try
            End If
            OJLogFileReader = New IO.StreamReader(DataSourcePath & "Log.csv")
            Dim StudentTemp As OJStudentInfo
            Dim ProblemTemp As OJProblemInfo
            Dim ProblemParticipantCheck As New Dictionary(Of String, Dictionary(Of String, Boolean))
            Dim DateList As New List(Of Date)
            StudentList.Clear()
            StudentDictionary.Clear()
            ProblemList.Clear()
            ProblemDictionary.Clear()
            OJSysInfo.NewProblemCount.Clear()
            Dim i As Integer
            For i = 0 To 23
                OJSysInfo.SubmitCountByHour(i) = 0
            Next
            While Not OJLogFileReader.EndOfStream
                OJLogLine = OJLogFileReader.ReadLine()
                OJLogTemp = ParseLog(OJLogLine)
                If OJLogTemp.IsParseFailed Then
                    Continue While
                End If
                '此处加入详细分析代码
                '学生数据
                If StudentDictionary.ContainsKey(OJLogTemp.StudentID) Then
                    With StudentDictionary(OJLogTemp.StudentID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                        If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnWorkdayAM += 1
                            Else
                                .SubmitCountOnWorkdayPM += 1
                            End If
                        Else
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnRestdayAM += 1
                            Else
                                .SubmitCountOnRestdayPM += 1
                            End If
                        End If
                        '拟合计算

                    End With
                Else
                    StudentTemp = New OJStudentInfo(OJLogTemp.StudentID)
                    StudentDictionary(OJLogTemp.StudentID) = StudentTemp
                    StudentList.Add(OJLogTemp.StudentID)
                    With StudentDictionary(OJLogTemp.StudentID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                        If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnWorkdayAM += 1
                            Else
                                .SubmitCountOnWorkdayPM += 1
                            End If
                        Else
                            If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                                .SubmitCountOnRestdayAM += 1
                            Else
                                .SubmitCountOnRestdayPM += 1
                            End If
                        End If
                    End With
                End If
                '题目数据
                If Not ProblemParticipantCheck.ContainsKey(OJLogTemp.ProblemID) Then
                    ProblemParticipantCheck.Add(OJLogTemp.ProblemID, New Dictionary(Of String, Boolean))
                End If
                If ProblemDictionary.ContainsKey(OJLogTemp.ProblemID) Then
                    With ProblemDictionary(OJLogTemp.ProblemID)
                        .SubmitCount += 1
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                            .ParticipantCount += 1
                            ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                        If IsTestListMissing Then
                            '提高鲁棒性：当一个题目出现更早的提交时更新数据
                            If OJLogTemp.DateSubmit < .CreateTime Then
                                If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                                    If OJSysInfo.NewProblemCount(.CreateTime) > 0 Then
                                        OJSysInfo.NewProblemCount(.CreateTime) -= 1
                                    End If
                                End If
                                .CreateTime = OJLogTemp.DateSubmit
                                If OJSysInfo.NewProblemCount.ContainsKey(OJLogTemp.DateSubmit) Then
                                    OJSysInfo.NewProblemCount(OJLogTemp.DateSubmit) += 1
                                Else
                                    OJSysInfo.NewProblemCount.Add(OJLogTemp.DateSubmit, 1)
                                End If
                            End If
                        End If
                    End With
                Else
                    ProblemTemp = New OJProblemInfo(OJLogTemp.ProblemID)
                    ProblemDictionary(OJLogTemp.ProblemID) = ProblemTemp
                    ProblemList.Add(OJLogTemp.ProblemID)
                    With ProblemDictionary(OJLogTemp.ProblemID)
                        .SubmitCount += 1
                        If IsTestListMissing Then
                            .CreateTime = OJLogTemp.DateSubmit
                        Else
                            .ProblemTaskSequenceIndex = OJLogTemp.TestID
                            .CreateTime = TestDictionary(.ProblemTaskSequenceIndex).BeginDate
                        End If
                        If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                            OJSysInfo.NewProblemCount(.CreateTime) += 1
                        Else
                            OJSysInfo.NewProblemCount.Add(.CreateTime, 1)
                        End If
                        If OJLogTemp.IsPassed Then
                            .ACCount += 1
                        End If
                        If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                            .ParticipantCount += 1
                            ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                        End If
                        If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                            .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                        Else
                            .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                        End If
                    End With
                End If
                '系统数据
                DateList.Add(OJLogTemp.DateSubmit)
                OJSysInfo.SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                DoEvents()
            End While
            OJLogFileReader.Close()
            ProblemParticipantCheck.Clear()
            OJSysInfo.EndDate = DateList.Max()
            OJSysInfo.StartDate = DateList.Min()
            '配置默认的分析起讫日期
            UserSpecifiedAnalyzeStartDate = OJSysInfo.StartDate
            UserSpecifiedAnalyzeEndDate = OJSysInfo.EndDate
            StudentList.Sort()
            ProblemList.Sort()
            '拟合计算
            '学生拟合
            If StudentList.Count >= 2 Then
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .FittingAC = IIf(.SubmitCount > 0, .ACCount + .ACCount / .SubmitCount, 0)
                        Dim DateLoop As New Date
                        DateLoop = UserSpecifiedAnalyzeStartDate
                        Dim DayLn() As Double
                        Dim SubmitAdjustedTotal() As Double
                        ReDim DayLn((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                        ReDim SubmitAdjustedTotal((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                        Dim StudentTotalSubmitSum As Integer = 0
                        Dim TotalProblems As Integer = 0
                        Dim Index = 0
                        While DateLoop < UserSpecifiedAnalyzeEndDate
                            If OJSysInfo.NewProblemCount.ContainsKey(DateLoop) Then
                                TotalProblems += OJSysInfo.NewProblemCount(DateLoop)
                            End If
                            If .SubmitCountByDay.ContainsKey(DateLoop) Then
                                StudentTotalSubmitSum += .SubmitCountByDay(DateLoop)
                            End If
                            DayLn(Index) = Math.Log((DateLoop - UserSpecifiedAnalyzeStartDate).Days + 1)
                            SubmitAdjustedTotal(Index) = StudentTotalSubmitSum - TotalProblems
                            DateLoop = DateLoop.AddDays(1)
                            Index += 1
                        End While
                        'Dim LinearFitResult As New Tuple(Of Double, Double)(0, 0)
                        'LinearFitResult = MathNet.Numerics.LinearRegression.SimpleRegression.Fit(DayLn, SubmitAdjustedTotal)
                        '.FittingB = LinearFitResult.Item1
                        '.FittingK_Kb = LinearFitResult.Item2
                        '.FittingR_Stb = MathNet.Numerics.Statistics.Correlation.Pearson(DayLn, SubmitAdjustedTotal) ^ 2
                        Try
                            Dim LinearFitReturn(2, 3) As Double
                            Dim LinearFitter As New LinearFit.Class1
                            Dim LinearFitX As New MWNumericArray(DayLn)
                            Dim LinearFitY As New MWNumericArray(SubmitAdjustedTotal)
                            LinearFitX = DayLn
                            LinearFitY = SubmitAdjustedTotal
                            LinearFitReturn = LinearFitter.LinearFit(LinearFitX, LinearFitY).ToArray()
                            LinearFitter.Dispose()
                            LinearFitX.Dispose()
                            LinearFitY.Dispose()
                            .FittingK_Kb = LinearFitReturn(0, 0)
                            .FittingR_Stb = LinearFitReturn(0, 1)
                            .FittingB = LinearFitReturn(0, 2)
                        Catch ex As Exception
                            IsMCRFailedStageStudentTotalSubmitByDayLnFit = True
                            .FittingB = 0
                            .FittingK_Kb = 0
                            .FittingR_Stb = 0
                        End Try
                    End With
                    DoEvents()
                Next
            Else
                IsSkippedStageStudentTotalSubmitByDayLnFit = True
            End If
            '题目拟合
            '产生题目参与人数-作业次序序列
            Dim ProblemCreateTimeList As New List(Of Date)
            Dim FirstProblemCreateTime As New Date
            Dim TaskSequence() As Integer
            Dim ParticipantCount() As Integer
            Dim ExponentialFittingA As Double = 0
            Dim ExponentialFittingB As Double = 0
            Dim ExponentialFittingC As Double = 0
            Dim ExponentialFitReturn(1, 3) As Double
            For i = 0 To ProblemList.Count - 1
                ProblemCreateTimeList.Add(ProblemDictionary(ProblemList(i)).CreateTime)
            Next
            FirstProblemCreateTime = ProblemCreateTimeList.Min()
            ReDim TaskSequence(ProblemList.Count - 1)
            ReDim ParticipantCount(ProblemList.Count - 1)
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    If IsTestListMissing Then
                        .ProblemTaskSequenceIndex = Math.Round((.CreateTime - FirstProblemCreateTime).Days / 7) + 1
                    End If
                    TaskSequence(i) = .ProblemTaskSequenceIndex
                    ParticipantCount(i) = .ParticipantCount
                End With
                DoEvents()
            Next
            If ProblemList.Count >= 3 Then
                Try
                    Dim ExponentialFitter As New e_fit.Class1
                    Dim ExponentialFitX As New MWNumericArray()
                    Dim ExponentialFitY As New MWNumericArray()
                    ExponentialFitX = TaskSequence
                    ExponentialFitY = ParticipantCount
                    ExponentialFitReturn = ExponentialFitter.e_fit(ExponentialFitX, ExponentialFitY).ToArray()
                    ExponentialFitter.Dispose()
                    ExponentialFitX.Dispose()
                    ExponentialFitY.Dispose()
                    ExponentialFittingA = -ExponentialFitReturn(0, 0)
                    ExponentialFittingB = ExponentialFitReturn(0, 1)
                    ExponentialFittingC = ExponentialFitReturn(0, 2)
                Catch ex As Exception
                    IsMCRFailedStageProblemParticipantFit = True
                    ExponentialFittingA = 0
                    ExponentialFittingB = 0
                    ExponentialFittingC = 0
                End Try
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                        .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                    End With
                    DoEvents()
                Next
            Else
                IsSkippedStageProblemParticipantFit = True
                ExponentialFittingA = 0
                ExponentialFittingB = 0
                ExponentialFittingC = 0
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                        .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                    End With
                    DoEvents()
                Next
            End If
            '学生聚类分析
            If StudentList.Count >= 4 Then
                Dim StudentAcStbKbMatrix(,) As Double
                Dim StudentClustResult(,) As Double
                Dim AcAverage As Double = 0
                Dim AcStandardDivision As Double = 0
                Dim StbAvarage As Double = 0
                Dim StbStandardDivision As Double = 0
                Dim KbAverage As Double = 0
                Dim KbStandardDivision As Double = 0
                ReDim StudentAcStbKbMatrix(StudentList.Count - 1, 2)
                ReDim StudentClustResult(StudentList.Count - 1, 1)
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        AcAverage += .FittingAC
                        StbAvarage += .FittingR_Stb
                        KbAverage += .FittingK_Kb
                    End With
                    DoEvents()
                Next
                AcAverage /= StudentList.Count
                StbAvarage /= StudentList.Count
                KbAverage /= StudentList.Count
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        AcStandardDivision += (.FittingAC - AcAverage) ^ 2
                        StbStandardDivision += (.FittingR_Stb - StbAvarage) ^ 2
                        KbStandardDivision += (.FittingK_Kb - KbAverage) ^ 2
                    End With
                    DoEvents()
                Next
                AcStandardDivision /= StudentList.Count
                StbStandardDivision /= StudentList.Count
                KbStandardDivision /= StudentList.Count
                AcStandardDivision = Math.Sqrt(AcStandardDivision)
                StbStandardDivision = Math.Sqrt(StbStandardDivision)
                KbStandardDivision = Math.Sqrt(KbStandardDivision)
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        StudentAcStbKbMatrix(i, 0) = (.FittingAC - AcAverage) / AcStandardDivision
                        StudentAcStbKbMatrix(i, 1) = (.FittingR_Stb - StbAvarage) / StbStandardDivision
                        StudentAcStbKbMatrix(i, 2) = (.FittingK_Kb - KbAverage) / KbStandardDivision
                    End With
                    DoEvents()
                Next
                Try
                    Dim StudentClustExec As New cluster.Class1
                    StudentClustResult = StudentClustExec.cluster(New MWNumericArray(StudentAcStbKbMatrix), 4).ToArray()
                    StudentClustExec.Dispose()
                    For i = 0 To StudentList.Count - 1
                        With StudentDictionary(StudentList(i))
                            .ClustResult = StudentClustResult(i, 0)
                        End With
                        DoEvents()
                    Next
                    StudentClustResultMapping.Clear()
                    StudentClustResultMapping.Add(0, "出错或被跳过")
                    StudentClustResultMapping.Add(1, "未分级，点击以分级")
                    StudentClustResultMapping.Add(2, "未分级，点击以分级")
                    StudentClustResultMapping.Add(3, "未分级，点击以分级")
                    StudentClustResultMapping.Add(4, "未分级，点击以分级")
                Catch ex As Exception
                    IsMCRFailedStageStudentClust = True
                    For i = 0 To StudentList.Count - 1
                        With StudentDictionary(StudentList(i))
                            .ClustResult = 0
                        End With
                        DoEvents()
                    Next
                    StudentClustResultMapping.Clear()
                    StudentClustResultMapping.Add(0, "出错或被跳过")
                    StudentClustResultMapping.Add(1, "未分级，点击以分级")
                    StudentClustResultMapping.Add(2, "未分级，点击以分级")
                    StudentClustResultMapping.Add(3, "未分级，点击以分级")
                    StudentClustResultMapping.Add(4, "未分级，点击以分级")
                End Try
            Else
                IsSkippedStageStudentClust = True
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .ClustResult = 0
                    End With
                    DoEvents()
                Next
                StudentClustResultMapping.Clear()
                StudentClustResultMapping.Add(0, "出错或被跳过")
                StudentClustResultMapping.Add(1, "未分级，点击以分级")
                StudentClustResultMapping.Add(2, "未分级，点击以分级")
                StudentClustResultMapping.Add(3, "未分级，点击以分级")
                StudentClustResultMapping.Add(4, "未分级，点击以分级")
            End If
            '题目聚类分析
            '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
            If ProblemList.Count >= 4 Then
                Dim ProblemEqJqMatrix(,) As Double
                Dim ProblemClustResult(,) As Double
                Dim EqAverage As Double = 0
                Dim EqStandardDivision As Double = 0
                Dim JqAverage As Double = 0
                Dim JqStandardDivision As Double = 0
                ReDim ProblemEqJqMatrix(ProblemList.Count - 1, 1)
                ReDim ProblemClustResult(ProblemList.Count - 1, 1)
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        EqAverage += .ParticipateValue_Eq
                        JqAverage += .EffortValue_Jq
                    End With
                    DoEvents()
                Next
                EqAverage /= ProblemList.Count
                JqAverage /= ProblemList.Count
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        EqStandardDivision += (.ParticipateValue_Eq - EqAverage) ^ 2
                        JqStandardDivision += (.EffortValue_Jq - JqAverage) ^ 2
                    End With
                    DoEvents()
                Next
                EqStandardDivision /= ProblemList.Count
                JqStandardDivision /= ProblemList.Count
                EqStandardDivision = Math.Sqrt(EqStandardDivision)
                JqStandardDivision = Math.Sqrt(JqStandardDivision)
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        ProblemEqJqMatrix(i, 0) = (.EffortValue_Jq - JqAverage) / JqStandardDivision
                        ProblemEqJqMatrix(i, 1) = (.ParticipateValue_Eq - EqAverage) / EqStandardDivision
                    End With
                    DoEvents()
                Next
                Try
                    Dim ProblemClustExec As New cluster.Class1
                    ProblemClustResult = ProblemClustExec.cluster(New MWNumericArray(ProblemEqJqMatrix), 4).ToArray()
                    ProblemClustExec.Dispose()
                    For i = 0 To ProblemList.Count - 1
                        With ProblemDictionary(ProblemList(i))
                            .ClustResult = ProblemClustResult(i, 0)
                        End With
                        DoEvents()
                    Next
                Catch ex As Exception
                    IsMCRFailedStageProblemClust = True
                    For i = 0 To ProblemList.Count - 1
                        ProblemDictionary(ProblemList(i)).ClustResult = 0
                        DoEvents()
                    Next
                End Try
            Else
                IsSkippedStageProblemClust = True
                For i = 0 To ProblemList.Count - 1
                    ProblemDictionary(ProblemList(i)).ClustResult = 0
                    DoEvents()
                Next
            End If
            '呈现列表
            lstStudentList.ItemsSource = StudentList
            lstProblemList.ItemsSource = ProblemList
            txtStartDate.Text = OJSysInfo.StartDate.ToLongDateString()
            txtEndDate.Text = OJSysInfo.EndDate.ToLongDateString()
            txtUserSpecifiedAnalyzeStartDate.Text = UserSpecifiedAnalyzeStartDate.ToLongDateString()
            txtUserSpecifiedAnalyzeEndDate.Text = UserSpecifiedAnalyzeEndDate.ToLongDateString()
            dtpUserSpecifiedAnalyzeStartDate.DisplayDateStart = OJSysInfo.StartDate
            dtpUserSpecifiedAnalyzeStartDate.DisplayDateEnd = OJSysInfo.EndDate
            dtpUserSpecifiedAnalyzeEndDate.DisplayDateStart = OJSysInfo.StartDate
            dtpUserSpecifiedAnalyzeEndDate.DisplayDateEnd = OJSysInfo.EndDate
            dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
            dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
            '题目参与人数与拟合图表
            If Not IsSkippedStageProblemParticipantFit Then
                If Not IsMCRFailedStageProblemParticipantFit Then
                    Dim ProblemParticipantDataSource As New List(Of KeyValuePair(Of Integer, Integer))
                    Dim ProblemParticipantFitDataSource As New List(Of KeyValuePair(Of Integer, Double))
                    For i = 0 To ProblemList.Count - 1
                        ProblemParticipantDataSource.Add(New KeyValuePair(Of Integer, Integer)(TaskSequence(i), ParticipantCount(i)))
                        ProblemParticipantFitDataSource.Add(New KeyValuePair(Of Integer, Double)(TaskSequence(i), ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, TaskSequence(i))))
                        DoEvents()
                    Next
                    sctProblemParticipant.ItemsSource = ProblemParticipantDataSource
                    linProblemParticipantFit.ItemsSource = ProblemParticipantFitDataSource
                    chartProblemParticipant.Visibility = Windows.Visibility.Visible
                End If
            End If
            '分时段总提交数图表
            Dim SubmitCountByTimeDataSource As New List(Of KeyValuePair(Of String, Integer))
            For i = 0 To 23
                SubmitCountByTimeDataSource.Add(New KeyValuePair(Of String, Integer)(i.ToString("00") & ":00" & vbCrLf & "~" & vbCrLf & (i + 1).ToString("00") & ":00", OJSysInfo.SubmitCountByHour(i)))
            Next
            colSubmitCountByTime.ItemsSource = SubmitCountByTimeDataSource
            chartSubmitCountByTime.Visibility = Windows.Visibility.Visible
            '题目发布曲线图表
            Dim j As Date
            j = UserSpecifiedAnalyzeStartDate
            Dim NewProblemCountByDayDataSource As New List(Of KeyValuePair(Of Date, Integer))
            While j <= UserSpecifiedAnalyzeEndDate
                If OJSysInfo.NewProblemCount.ContainsKey(j) Then
                    NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, OJSysInfo.NewProblemCount(j)))
                Else
                    NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, 0))
                End If
                j = j.AddDays(1)
            End While
            linNewProblemCountByDay.ItemsSource = NewProblemCountByDayDataSource
            chartNewProblemCountByDay.Visibility = Windows.Visibility.Visible
            '题目聚类分析结果
            '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
            '聚类图表的四类依次序分为:
            '低Jq低Eq，低Jq高Eq，高Jq低Eq，高Jq高Eq
            If Not IsSkippedStageProblemClust Then
                If Not IsMCRFailedStageProblemClust Then
                    Dim ScatterProblemCluster1DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster2DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster3DataSource As New List(Of KeyValuePair(Of Double, Double))
                    Dim ScatterProblemCluster4DataSource As New List(Of KeyValuePair(Of Double, Double))
                    For i = 0 To ProblemList.Count - 1
                        With ProblemDictionary(ProblemList(i))
                            If .ClustResult = 1 Then
                                ScatterProblemCluster1DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            ElseIf .ClustResult = 2 Then
                                ScatterProblemCluster2DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            ElseIf .ClustResult = 3 Then
                                ScatterProblemCluster3DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            Else
                                ScatterProblemCluster4DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                            End If
                        End With
                        DoEvents()
                    Next
                    sctProblemCluster1.ItemsSource = ScatterProblemCluster1DataSource
                    sctProblemCluster2.ItemsSource = ScatterProblemCluster2DataSource
                    sctProblemCluster3.ItemsSource = ScatterProblemCluster3DataSource
                    sctProblemCluster4.ItemsSource = ScatterProblemCluster4DataSource
                    chartProblemCluster.Visibility = Windows.Visibility.Visible
                End If
            End If
            Await AnalyzeProgress.CloseAsync()
            If IsMCRFailedStageStudentTotalSubmitByDayLnFit Or IsMCRFailedStageProblemParticipantFit Or IsMCRFailedStageStudentClust Or IsMCRFailedStageProblemClust Then
                Await Me.ShowMessageAsync("MATLAB Runtime发生问题", "本程序的运行依赖 MathWorks MATLAB Runtime R2018b，但是该组件可能发生问题或无法被调用。请前往 https://ww2.mathworks.cn/products/compiler/matlab-runtime.html 下载并安装适用于您的操作系统的 MathWorks MATLAB Runtime R2018b。如果您已经安装了 MathWorks MATLAB Runtime R2018b，您可能需要将其重新安装。存在边缘情况的数据集也可能引发此错误。" & vbCrLf & vbCrLf & "程序仍可继续运行，但部分功能可能发生异常。" & vbCrLf & vbCrLf & "以下功能依赖 MathWorks MATLAB Runtime R2018b: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsMCRFailedStageStudentTotalSubmitByDayLnFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsMCRFailedStageProblemParticipantFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsMCRFailedStageStudentClust, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsMCRFailedStageProblemClust, "执行时失败或发生错误。", "执行成功。"))
                IsMCRFailedStageStudentTotalSubmitByDayLnFit = False
                IsMCRFailedStageProblemParticipantFit = False
                IsMCRFailedStageStudentClust = False
                IsMCRFailedStageProblemClust = False
            End If
            If IsSkippedStageStudentTotalSubmitByDayLnFit Or IsSkippedStageProblemParticipantFit Or IsSkippedStageStudentClust Or IsSkippedStageProblemClust Then
                Await Me.ShowMessageAsync("一些分析步骤因数据集容量过小被跳过", "一部分拟合/聚类分析步骤被跳过，这可能是由数据集容量过小的问题导致的。" & vbCrLf & vbCrLf & "您可以尝试通过更新日志、扩大分析时间段长度等方式扩展数据集容量。" & vbCrLf & vbCrLf & "该问题会导致一部分图表或分析结果无法显示或显示的结果不正确。" & vbCrLf & vbCrLf & "以下分析流程可能已被跳过: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsSkippedStageStudentTotalSubmitByDayLnFit, "已被跳过。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsSkippedStageProblemParticipantFit, "已被跳过。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsSkippedStageStudentClust, "已被跳过。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsSkippedStageProblemClust, "已被跳过。", "执行成功。"))
                IsSkippedStageStudentTotalSubmitByDayLnFit = False
                IsSkippedStageProblemParticipantFit = False
                IsSkippedStageStudentClust = False
                IsSkippedStageProblemClust = False
            End If
        End If
    End Sub

    Private Sub lstStudentList_GotFocus(sender As Object, e As RoutedEventArgs) Handles lstStudentList.GotFocus
        lstStudentList.BorderBrush = SystemColors.ControlDarkDarkBrush
    End Sub

    Private Sub lstStudentList_LostFocus(sender As Object, e As RoutedEventArgs) Handles lstStudentList.LostFocus
        lstStudentList.BorderBrush = SystemColors.ScrollBarBrush
    End Sub

    Private Sub txtStudentSearch_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtStudentSearch.GotFocus
        If txtStudentSearch.Text = "搜索学生学号" Then
            txtStudentSearch.Text = ""
            txtStudentSearch.Foreground = SystemColors.ControlTextBrush
        End If
    End Sub

    Private Sub txtStudentSearch_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtStudentSearch.LostFocus
        If txtStudentSearch.Text = "" Then
            txtStudentSearch.Text = "搜索学生学号"
            txtStudentSearch.Foreground = SystemColors.ScrollBarBrush
        End If
    End Sub
    Private Sub lstProblemList_GotFocus(sender As Object, e As RoutedEventArgs) Handles lstProblemList.GotFocus
        lstProblemList.BorderBrush = SystemColors.ControlDarkDarkBrush
    End Sub

    Private Sub lstProblemList_LostFocus(sender As Object, e As RoutedEventArgs) Handles lstProblemList.LostFocus
        lstProblemList.BorderBrush = SystemColors.ScrollBarBrush
    End Sub

    Private Sub txtProblemSearch_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtProblemSearch.GotFocus
        If txtProblemSearch.Text = "搜索题目编号" Then
            txtProblemSearch.Text = ""
            txtProblemSearch.Foreground = SystemColors.ControlTextBrush
        End If
    End Sub

    Private Sub txtProblemSearch_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtProblemSearch.LostFocus
        If txtProblemSearch.Text = "" Then
            txtProblemSearch.Text = "搜索题目编号"
            txtProblemSearch.Foreground = SystemColors.ScrollBarBrush
        End If
    End Sub

    Private Sub txtStudentSearch_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtStudentSearch.TextChanged
        If Not txtStudentSearch.IsFocused Then
            Exit Sub
        End If
        If txtStudentSearch.Text = "" Then
            lstStudentList.ItemsSource = StudentList
            Exit Sub
        End If
        Dim SearchTemp As New List(Of String)
        SearchTemp.Clear()
        If StudentList.Count = 0 Then
            Exit Sub
        End If
        Dim i As Integer
        For i = 0 To StudentList.Count - 1
            If StudentList(i).ToUpper.Contains(txtStudentSearch.Text.ToUpper) Then
                SearchTemp.Add(StudentList(i))
            End If
        Next
        SearchTemp.Sort()
        lstStudentList.ItemsSource = SearchTemp
    End Sub

    Private Sub txtProblemSearch_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtProblemSearch.TextChanged
        If Not txtProblemSearch.IsFocused Then
            Exit Sub
        End If
        If txtProblemSearch.Text = "" Then
            lstProblemList.ItemsSource = ProblemList
            Exit Sub
        End If
        Dim SearchTemp As New List(Of String)
        SearchTemp.Clear()
        If ProblemList.Count = 0 Then
            Exit Sub
        End If
        Dim i As Integer
        For i = 0 To ProblemList.Count - 1
            If ProblemList(i).ToUpper.Contains(txtProblemSearch.Text.ToUpper) Then
                SearchTemp.Add(ProblemList(i))
            End If
        Next
        SearchTemp.Sort()
        lstProblemList.ItemsSource = SearchTemp
    End Sub

    Private Sub lstStudentList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstStudentList.SelectionChanged
        If lstStudentList.SelectedIndex <> -1 Then
            txtStudentID.Text = lstStudentList.SelectedItem.ToString
            With StudentDictionary(txtStudentID.Text)
                txtStudentSubmitCount.Text = .SubmitCount
                txtStudentClustResult.Text = IIf(StudentClustResultMapping.ContainsKey(.ClustResult), StudentClustResultMapping(.ClustResult), .ClustResult)
                txtStudentACCount.Text = .ACCount
                txtStudentACRate.Text = (.ACCount / .SubmitCount).ToString("P")
                txtStudentSubmitCountOnWorkdayAM.Text = .SubmitCountOnWorkdayAM
                txtStudentSubmitCountOnWorkdayPM.Text = .SubmitCountOnWorkdayPM
                txtStudentSubmitCountOnRestdayAM.Text = .SubmitCountOnRestdayAM
                txtStudentSubmitCountOnRestdayPM.Text = .SubmitCountOnRestdayPM
                txtFittingAC.Text = .FittingAC.ToString("F5")
                txtFittingK_Kb.Text = .FittingK_Kb.ToString("F5")
                txtFittingR_Stb.Text = .FittingR_Stb.ToString("F5")
                Dim OriginalSelection As Integer
                OriginalSelection = tabStudentCharts.SelectedIndex

                chartStudentACRate.Visibility = Windows.Visibility.Collapsed
                tabStudentCharts.SelectedIndex = 0
                DoEvents()
                pieStudentACRate.ItemsSource = Nothing
                pieStudentACRate.Refresh()
                DoEvents()
                Dim StudentACRateDataSource As New List(Of KeyValuePair(Of String, Integer))
                StudentACRateDataSource.Add(New KeyValuePair(Of String, Integer)("通过  ", .ACCount))
                StudentACRateDataSource.Add(New KeyValuePair(Of String, Integer)("未通过", .SubmitCount - .ACCount))
                pieStudentACRate.ItemsSource = StudentACRateDataSource
                DoEvents()
                chartStudentACRate.Visibility = Windows.Visibility.Visible

                chartStudentSubmitByDay.Visibility = Windows.Visibility.Collapsed
                tabStudentCharts.SelectedIndex = 1
                DoEvents()
                linStudentSubmitByDay.ItemsSource = Nothing
                linStudentSubmitByDay.Refresh()
                DoEvents()
                Dim StudentSubmitByDayDataSource As New List(Of KeyValuePair(Of Date, Integer))
                Dim i As Date
                Dim j As Integer
                i = UserSpecifiedAnalyzeStartDate
                While i <= UserSpecifiedAnalyzeEndDate
                    If .SubmitCountByDay.ContainsKey(i) Then
                        StudentSubmitByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(i, .SubmitCountByDay(i)))
                    Else
                        StudentSubmitByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(i, 0))
                    End If
                    i = i.AddDays(1)
                End While
                linStudentSubmitByDay.ItemsSource = StudentSubmitByDayDataSource
                DoEvents()
                chartStudentSubmitByDay.Visibility = Windows.Visibility.Visible

                chartStudentTotalSubmitByDayLn.Visibility = Windows.Visibility.Collapsed
                tabStudentCharts.SelectedIndex = 2
                DoEvents()
                linStudentTotalSubmitByDayLnFit.ItemsSource = Nothing
                linStudentTotalSubmitByDayLnFit.Refresh()
                DoEvents()
                sctStudentTotalSubmitByDayLn.ItemsSource = Nothing
                sctStudentTotalSubmitByDayLn.Refresh()
                DoEvents()
                Dim StudentTotalSubmitByDayLnDataSource As New List(Of KeyValuePair(Of Double, Integer))
                Dim StudentTotalSubmitByDayLnFitDataSource As New List(Of KeyValuePair(Of Double, Integer))
                Dim StudentTotalSubmitSum As Integer = 0
                Dim TotalProblems As Integer = 0
                i = UserSpecifiedAnalyzeStartDate
                StudentTotalSubmitSum = 0
                TotalProblems = 0
                While i <= UserSpecifiedAnalyzeEndDate
                    If OJSysInfo.NewProblemCount.ContainsKey(i) Then
                        TotalProblems += OJSysInfo.NewProblemCount(i)
                    End If
                    If .SubmitCountByDay.ContainsKey(i) Then
                        StudentTotalSubmitSum += .SubmitCountByDay(i)
                    End If
                    StudentTotalSubmitByDayLnDataSource.Add(New KeyValuePair(Of Double, Integer)(Math.Log((i - UserSpecifiedAnalyzeStartDate).Days + 1), StudentTotalSubmitSum - TotalProblems))
                    i = i.AddDays(1)
                End While
                StudentTotalSubmitByDayLnFitDataSource.Add(New KeyValuePair(Of Double, Integer)(0, .FittingB))
                StudentTotalSubmitByDayLnFitDataSource.Add(New KeyValuePair(Of Double, Integer)(Math.Log((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days + 1), (Math.Log((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days + 1)) * .FittingK_Kb + .FittingB))
                linStudentTotalSubmitByDayLnFit.ItemsSource = StudentTotalSubmitByDayLnFitDataSource
                DoEvents()
                sctStudentTotalSubmitByDayLn.ItemsSource = StudentTotalSubmitByDayLnDataSource
                DoEvents()
                chartStudentTotalSubmitByDayLn.Visibility = Windows.Visibility.Visible

                chartStudentTotalSubmitByTime.Visibility = Windows.Visibility.Collapsed
                tabStudentCharts.SelectedIndex = 3
                DoEvents()
                colStudentTotalSubmitByTime.ItemsSource = Nothing
                colStudentTotalSubmitByTime.Refresh()
                DoEvents()
                Dim StudentTotalSubmitByTimeDataSource As New List(Of KeyValuePair(Of String, Integer))
                For j = 0 To 23
                    StudentTotalSubmitByTimeDataSource.Add(New KeyValuePair(Of String, Integer)(j.ToString("00") & ":00" & vbCrLf & "~" & vbCrLf & (j + 1).ToString("00") & ":00", .SubmitCountByHour(j)))
                Next
                colStudentTotalSubmitByTime.ItemsSource = StudentTotalSubmitByTimeDataSource
                DoEvents()
                chartStudentTotalSubmitByTime.Visibility = Windows.Visibility.Visible
                tabStudentCharts.SelectedIndex = OriginalSelection
            End With
            icoStudentLink.Foreground = SystemColors.WindowTextBrush
            btnStudentLink.IsEnabled = True
            btnStudentLink.Cursor = System.Windows.Input.Cursors.Hand
        End If
    End Sub

    Private Sub lstProblemList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstProblemList.SelectionChanged
        If lstProblemList.SelectedIndex <> -1 Then
            txtProblemID.Text = lstProblemList.SelectedItem.ToString()
            With ProblemDictionary(txtProblemID.Text)
                txtProblemSubmitCount.Text = .SubmitCount
                txtProblemParticipantCount.Text = .ParticipantCount
                txtProblemCreateTime.Text = .CreateTime.ToLongDateString()
                txtProblemACCount.Text = .ACCount
                txtProblemACRate.Text = (.ACCount / .SubmitCount).ToString("P")
                txtEffortValue_Jq.Text = .EffortValue_Jq.ToString("F5")
                txtParticipateValue_Eq.Text = .ParticipateValue_Eq.ToString("F5")
                Dim OriginalSelection As Integer
                OriginalSelection = tabProblemCharts.SelectedIndex

                chartProblemACRate.Visibility = Windows.Visibility.Collapsed
                tabProblemCharts.SelectedIndex = 0
                DoEvents()
                pieProblemACRate.ItemsSource = Nothing
                pieProblemACRate.Refresh()
                DoEvents()
                Dim ProblemACRateDataSource As New List(Of KeyValuePair(Of String, Integer))
                ProblemACRateDataSource.Add(New KeyValuePair(Of String, Integer)("通过  ", .ACCount))
                ProblemACRateDataSource.Add(New KeyValuePair(Of String, Integer)("未通过", .SubmitCount - .ACCount))
                pieProblemACRate.ItemsSource = ProblemACRateDataSource
                chartProblemACRate.Visibility = Windows.Visibility.Visible

                chartProblemSubmitByDay.Visibility = Windows.Visibility.Collapsed
                tabProblemCharts.SelectedIndex = 1
                DoEvents()
                linProblemSubmitByDay.ItemsSource = Nothing
                linProblemSubmitByDay.Refresh()
                DoEvents()
                Dim ProblemSubmitByDayDataSource As New List(Of KeyValuePair(Of Date, Integer))
                Dim i As Date
                i = UserSpecifiedAnalyzeStartDate
                While i <= UserSpecifiedAnalyzeEndDate
                    If .SubmitCountByDay.ContainsKey(i) Then
                        ProblemSubmitByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(i, .SubmitCountByDay(i)))
                    Else
                        ProblemSubmitByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(i, 0))
                    End If
                    i = i.AddDays(1)
                End While
                linProblemSubmitByDay.ItemsSource = ProblemSubmitByDayDataSource
                chartProblemSubmitByDay.Visibility = Windows.Visibility.Visible
                tabProblemCharts.SelectedIndex = OriginalSelection
            End With
            icoProblemLink.Foreground = SystemColors.WindowTextBrush
            btnProblemLink.IsEnabled = True
            btnProblemLink.Cursor = System.Windows.Input.Cursors.Hand
        End If
    End Sub

    Private Sub btnStudentLink_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnStudentLink.MouseLeftButtonUp
        Dim StudentLink As New Process
        If txtStudentID.Text = "" Then
            Exit Sub
        End If
        With StudentLink
            .StartInfo.FileName = "http://oj.bmeonline.cn/userinfo.php?user=" & txtStudentID.Text
            .Start()
        End With
    End Sub

    Private Sub btnProblemLink_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnProblemLink.MouseLeftButtonUp
        Dim ProblemLink As New Process
        If txtProblemID.Text = "" Then
            Exit Sub
        End If
        With ProblemLink
            .StartInfo.FileName = "http://oj.bmeonline.cn/problem.php?id=" & txtProblemID.Text
            .Start()
        End With
    End Sub

    Private Sub btnSystemLink_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnSystemLink.MouseLeftButtonUp
        Dim SystemLink As New Process
        With SystemLink
            .StartInfo.FileName = txtOJLink.Text
            .Start()
        End With
    End Sub

    Private Async Sub chkCustomAnalyzeEndDate_Click(sender As Object, e As RoutedEventArgs) Handles chkCustomAnalyzeEndDate.Click
        IsUserSpecifiedAnalyzeEndDateEqualsLogEndDate = Not chkCustomAnalyzeEndDate.IsChecked
        dtpUserSpecifiedAnalyzeEndDate.IsEnabled = chkCustomAnalyzeEndDate.IsChecked
        If UserSpecifiedAnalyzeEndDate < UserSpecifiedAnalyzeStartDate And chkCustomAnalyzeStartDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析结束日期早于分析起始日期，已将其设为当前指定的分析起始日期。")
            UserSpecifiedAnalyzeEndDate = UserSpecifiedAnalyzeStartDate
            dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
        End If
        If UserSpecifiedAnalyzeStartDate > UserSpecifiedAnalyzeEndDate And chkCustomAnalyzeEndDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析起始日期晚于分析结束日期，已将其设为当前指定的分析结束日期。")
            UserSpecifiedAnalyzeStartDate = UserSpecifiedAnalyzeEndDate
            dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
        End If
    End Sub

    Private Async Sub chkCustomAnalyzeStartDate_Click(sender As Object, e As RoutedEventArgs) Handles chkCustomAnalyzeStartDate.Click
        IsUserSpecifiedAnalyzeStartDateEqualsLogStartDate = Not chkCustomAnalyzeStartDate.IsChecked
        dtpUserSpecifiedAnalyzeStartDate.IsEnabled = chkCustomAnalyzeStartDate.IsChecked
        If UserSpecifiedAnalyzeEndDate < UserSpecifiedAnalyzeStartDate And chkCustomAnalyzeStartDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析结束日期早于分析起始日期，已将其设为当前指定的分析起始日期。")
            UserSpecifiedAnalyzeEndDate = UserSpecifiedAnalyzeStartDate
            dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
        End If
        If UserSpecifiedAnalyzeStartDate > UserSpecifiedAnalyzeEndDate And chkCustomAnalyzeEndDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析起始日期晚于分析结束日期，已将其设为当前指定的分析结束日期。")
            UserSpecifiedAnalyzeStartDate = UserSpecifiedAnalyzeEndDate
            dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
        End If
    End Sub

    Private Async Sub dtpUserSpecifiedAnalyzeEndDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpUserSpecifiedAnalyzeEndDate.SelectedDateChanged
        UserSpecifiedAnalyzeEndDate = dtpUserSpecifiedAnalyzeEndDate.SelectedDate
        If UserSpecifiedAnalyzeEndDate < UserSpecifiedAnalyzeStartDate And chkCustomAnalyzeStartDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析结束日期早于分析起始日期，已将其设为当前指定的分析起始日期。")
            UserSpecifiedAnalyzeEndDate = UserSpecifiedAnalyzeStartDate
            dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
        End If
    End Sub

    Private Async Sub dtpUserSpecifiedAnalyzeStartDate_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpUserSpecifiedAnalyzeStartDate.SelectedDateChanged
        UserSpecifiedAnalyzeStartDate = dtpUserSpecifiedAnalyzeStartDate.SelectedDate
        If UserSpecifiedAnalyzeStartDate > UserSpecifiedAnalyzeEndDate And chkCustomAnalyzeEndDate.IsChecked Then
            Await Me.ShowMessageAsync("提示", "自定义的分析起始日期晚于分析结束日期，已将其设为当前指定的分析结束日期。")
            UserSpecifiedAnalyzeStartDate = UserSpecifiedAnalyzeEndDate
            dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
        End If
    End Sub

    Private Async Sub btnUpdate_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdate.Click
        Dim IsProblemListMissing As Boolean = False
        Dim IsUserListMissing As Boolean = False
        Dim IsTestListMissing As Boolean = False
        If chkCustomAnalyzeStartDate.IsChecked Then
            UserSpecifiedAnalyzeStartDate = dtpUserSpecifiedAnalyzeStartDate.SelectedDate
        End If
        If chkCustomAnalyzeEndDate.IsChecked Then
            UserSpecifiedAnalyzeEndDate = dtpUserSpecifiedAnalyzeEndDate.SelectedDate
        End If
        If DataSourcePath(DataSourcePath.Length - 1) <> "\" Then
            DataSourcePath = DataSourcePath & "\"
        End If
        '开始分析OJ日志。
        IsUpdateFinished = True
        Dim AnalyzeProgress As Dialogs.ProgressDialogController
        '校验数据。
        AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在校验数据", "正在对 OnlineJudge 数据进行校验，这可能需要一点点的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
        If Not DirectoryExists(DataSourcePath) Then
            Await AnalyzeProgress.CloseAsync()
            Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
            Exit Sub
        End If
        If Not FileExists(DataSourcePath & "Log.csv") Then
            Await AnalyzeProgress.CloseAsync()
            Await ShowMessageAsync("数据源错误", "引用的数据源不存在，请重新选择分析数据源。")
            Exit Sub
        End If
        If Not FileExists(DataSourcePath & "User.csv") Then
            IsUserListMissing = True
        End If
        If Not FileExists(DataSourcePath & "Problem.csv") Then
            IsProblemListMissing = True
        End If
        If Not FileExists(DataSourcePath & "Test.csv") Then
            IsTestListMissing = True
        End If
        Await AnalyzeProgress.CloseAsync()
        If IsUserListMissing Or IsProblemListMissing Or IsTestListMissing Then
            Await ShowMessageAsync("数据源扩展缺失", "引用的数据源缺失部分扩展数据文件，分析仍将继续执行，但部分分析数据将依赖程序内置的计算算法生成，这可能导致少数分析结果不可靠。" & vbCrLf & vbCrLf & "下列扩展数据文件隶属于数据源：" & vbCrLf & _
                                   DataSourcePath & "User.csv - 描述了 OJ 系统上的用户。" & IIf(IsUserListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                   DataSourcePath & "Problem.csv - 描述了 OJ 系统上的题目。" & IIf(IsProblemListMissing, "该数据源扩展不存在。", "该数据源扩展存在。") & vbCrLf & _
                                   DataSourcePath & "Test.csv - 描述了 OJ 系统上的题目集。" & IIf(IsTestListMissing, "该数据源扩展不存在，题目的创建时间数据可能不准确。", "该数据源扩展存在。") _
                                   )
        End If

        '初始化
        StudentList.Clear()
        StudentDictionary.Clear()
        ProblemList.Clear()
        ProblemDictionary.Clear()
        TestDictionary.Clear()
        lstStudentList.ItemsSource = EmptyList
        lstProblemList.ItemsSource = EmptyList
        txtStudentSearch.Text = "搜索学生学号"
        txtStudentSearch.Foreground = SystemColors.ScrollBarBrush
        txtProblemSearch.Text = "搜索题目编号"
        txtProblemSearch.Foreground = SystemColors.ScrollBarBrush
        icoStudentLink.Foreground = SystemColors.GrayTextBrush
        btnStudentLink.IsEnabled = False
        btnStudentLink.Cursor = System.Windows.Input.Cursors.Arrow
        icoProblemLink.Foreground = SystemColors.GrayTextBrush
        btnProblemLink.IsEnabled = False
        btnProblemLink.Cursor = System.Windows.Input.Cursors.Arrow
        txtStudentID.Text = ""
        txtStudentClustResult.Text = ""
        txtStudentSubmitCount.Text = ""
        txtStudentACCount.Text = ""
        txtStudentACRate.Text = ""
        txtStudentSubmitCountOnWorkdayAM.Text = ""
        txtStudentSubmitCountOnWorkdayPM.Text = ""
        txtStudentSubmitCountOnRestdayAM.Text = ""
        txtStudentSubmitCountOnRestdayPM.Text = ""
        txtFittingAC.Text = ""
        txtFittingK_Kb.Text = ""
        txtFittingR_Stb.Text = ""
        txtProblemID.Text = ""
        txtProblemCreateTime.Text = ""
        txtProblemParticipantCount.Text = ""
        txtProblemSubmitCount.Text = ""
        txtProblemACCount.Text = ""
        txtProblemACRate.Text = ""
        txtEffortValue_Jq.Text = ""
        txtParticipateValue_Eq.Text = ""
        txtStartDate.Text = ""
        txtEndDate.Text = ""
        txtUserSpecifiedAnalyzeStartDate.Text = ""
        txtUserSpecifiedAnalyzeEndDate.Text = ""
        chartStudentACRate.Visibility = Windows.Visibility.Hidden
        chartStudentSubmitByDay.Visibility = Windows.Visibility.Hidden
        chartStudentTotalSubmitByDayLn.Visibility = Windows.Visibility.Hidden
        chartStudentTotalSubmitByTime.Visibility = Windows.Visibility.Hidden
        chartProblemACRate.Visibility = Windows.Visibility.Hidden
        chartProblemSubmitByDay.Visibility = Windows.Visibility.Hidden
        chartProblemParticipant.Visibility = Windows.Visibility.Hidden
        chartSubmitCountByTime.Visibility = Windows.Visibility.Hidden
        chartNewProblemCountByDay.Visibility = Windows.Visibility.Hidden
        pieStudentACRate.ItemsSource = Nothing
        DoEvents()
        pieStudentACRate.Refresh()
        DoEvents()
        linStudentSubmitByDay.ItemsSource = Nothing
        DoEvents()
        linStudentSubmitByDay.Refresh()
        DoEvents()
        sctStudentTotalSubmitByDayLn.ItemsSource = Nothing
        DoEvents()
        sctStudentTotalSubmitByDayLn.Refresh()
        DoEvents()
        linStudentTotalSubmitByDayLnFit.ItemsSource = Nothing
        DoEvents()
        linStudentTotalSubmitByDayLnFit.Refresh()
        DoEvents()
        colStudentTotalSubmitByTime.ItemsSource = Nothing
        DoEvents()
        colStudentTotalSubmitByTime.Refresh()
        DoEvents()
        pieProblemACRate.ItemsSource = Nothing
        DoEvents()
        pieProblemACRate.Refresh()
        DoEvents()
        linProblemSubmitByDay.ItemsSource = Nothing
        DoEvents()
        linProblemSubmitByDay.Refresh()
        DoEvents()
        sctProblemParticipant.ItemsSource = Nothing
        DoEvents()
        sctProblemParticipant.Refresh()
        DoEvents()
        linProblemParticipantFit.ItemsSource = Nothing
        DoEvents()
        linProblemParticipantFit.Refresh()
        DoEvents()
        linNewProblemCountByDay.ItemsSource = Nothing
        DoEvents()
        linNewProblemCountByDay.Refresh()
        DoEvents()
        colSubmitCountByTime.ItemsSource = Nothing
        DoEvents()
        colSubmitCountByTime.Refresh()
        DoEvents()

        Dim OJLogFileReader As System.IO.StreamReader
        Dim OJLogLine As String
        Dim OJLogTemp As New OJLog
        Dim OJTestLogLine As String
        Dim OJTestTemp As New OJTestInfo
        AnalyzeProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在分析数据", "正在对 OnlineJudge 数据进行分析，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
        '预存测试数据
        If Not IsTestListMissing Then
            Try
                OJLogFileReader = New IO.StreamReader(DataSourcePath & "Test.csv")
                While Not OJLogFileReader.EndOfStream
                    OJTestLogLine = OJLogFileReader.ReadLine()
                    OJTestTemp = ParseOJTestLogLine(OJTestLogLine)
                    If OJTestTemp.IsParseFailed Then
                        Continue While
                    End If
                    If TestDictionary.ContainsKey(OJTestTemp.TestID) Then
                        If OJTestTemp.BeginDate < TestDictionary(OJTestTemp.TestID).BeginDate Then
                            TestDictionary(OJTestTemp.TestID).BeginDate = OJTestTemp.BeginDate
                        End If
                        If OJTestTemp.EndDate > TestDictionary(OJTestTemp.TestID).EndDate Then
                            TestDictionary(OJTestTemp.TestID).EndDate = OJTestTemp.EndDate
                        End If
                    Else
                        TestDictionary.Add(OJTestTemp.TestID, OJTestTemp)
                    End If
                End While
                If TestDictionary.Count = 0 Then
                    IsTestListMissing = True
                End If
                OJLogFileReader.Close()
            Catch ex As Exception
                IsTestListMissing = True
            End Try
        End If
        OJLogFileReader = New IO.StreamReader(DataSourcePath & "Log.csv")
        Dim StudentTemp As OJStudentInfo
        Dim ProblemTemp As OJProblemInfo
        Dim ProblemParticipantCheck As New Dictionary(Of String, Dictionary(Of String, Boolean))
        Dim DateList As New List(Of Date)
        StudentList.Clear()
        StudentDictionary.Clear()
        ProblemList.Clear()
        ProblemDictionary.Clear()
        OJSysInfo.NewProblemCount.Clear()
        Dim i As Integer
        For i = 0 To 23
            OJSysInfo.SubmitCountByHour(i) = 0
        Next
        '预更新日志时间
        While Not OJLogFileReader.EndOfStream
            OJLogLine = OJLogFileReader.ReadLine()
            OJLogTemp = ParseLog(OJLogLine)
            If OJLogTemp.IsParseFailed Then
                Continue While
            End If
            DateList.Add(OJLogTemp.DateSubmit)
        End While
        OJSysInfo.StartDate = DateList.Min
        OJSysInfo.EndDate = DateList.Max
        If IsUserSpecifiedAnalyzeStartDateEqualsLogStartDate Then
            UserSpecifiedAnalyzeStartDate = OJSysInfo.StartDate
        End If
        If IsUserSpecifiedAnalyzeEndDateEqualsLogEndDate Then
            UserSpecifiedAnalyzeEndDate = OJSysInfo.EndDate
        End If
        OJLogFileReader.Dispose()
        OJLogFileReader = New IO.StreamReader(DataSourcePath & "Log.csv")
        While Not OJLogFileReader.EndOfStream
            OJLogLine = OJLogFileReader.ReadLine()
            OJLogTemp = ParseLog(OJLogLine)
            If OJLogTemp.IsParseFailed Then
                Continue While
            End If
            '确定是否在用户指定的分析范围内
            If OJLogTemp.DateSubmit < UserSpecifiedAnalyzeStartDate Or OJLogTemp.DateSubmit > UserSpecifiedAnalyzeEndDate Then
                Continue While
            End If
            '此处加入详细分析代码
            '学生数据
            If StudentDictionary.ContainsKey(OJLogTemp.StudentID) Then
                With StudentDictionary(OJLogTemp.StudentID)
                    .SubmitCount += 1
                    If OJLogTemp.IsPassed Then
                        .ACCount += 1
                    End If
                    If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                        .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                    Else
                        .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                    End If
                    .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                    If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                        If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                            .SubmitCountOnWorkdayAM += 1
                        Else
                            .SubmitCountOnWorkdayPM += 1
                        End If
                    Else
                        If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                            .SubmitCountOnRestdayAM += 1
                        Else
                            .SubmitCountOnRestdayPM += 1
                        End If
                    End If
                    '拟合计算

                End With
            Else
                StudentTemp = New OJStudentInfo(OJLogTemp.StudentID)
                StudentDictionary(OJLogTemp.StudentID) = StudentTemp
                StudentList.Add(OJLogTemp.StudentID)
                With StudentDictionary(OJLogTemp.StudentID)
                    .SubmitCount += 1
                    If OJLogTemp.IsPassed Then
                        .ACCount += 1
                    End If
                    If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                        .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                    Else
                        .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                    End If
                    .SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
                    If OJLogTemp.WeekdaySubmit >= 2 And OJLogTemp.WeekdaySubmit <= 6 Then
                        If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                            .SubmitCountOnWorkdayAM += 1
                        Else
                            .SubmitCountOnWorkdayPM += 1
                        End If
                    Else
                        If OJLogTemp.LogTime.Hour >= 0 And OJLogTemp.LogTime.Hour < 12 Then
                            .SubmitCountOnRestdayAM += 1
                        Else
                            .SubmitCountOnRestdayPM += 1
                        End If
                    End If
                End With
            End If
            '题目数据
            If Not ProblemParticipantCheck.ContainsKey(OJLogTemp.ProblemID) Then
                ProblemParticipantCheck.Add(OJLogTemp.ProblemID, New Dictionary(Of String, Boolean))
            End If
            If ProblemDictionary.ContainsKey(OJLogTemp.ProblemID) Then
                With ProblemDictionary(OJLogTemp.ProblemID)
                    .SubmitCount += 1
                    If OJLogTemp.IsPassed Then
                        .ACCount += 1
                    End If
                    If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                        .ParticipantCount += 1
                        ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                    End If
                    If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                        .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                    Else
                        .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                    End If
                    If IsTestListMissing Then
                        '提高鲁棒性：当一个题目出现更早的提交时更新数据
                        If OJLogTemp.DateSubmit < .CreateTime Then
                            If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                                If OJSysInfo.NewProblemCount(.CreateTime) > 0 Then
                                    OJSysInfo.NewProblemCount(.CreateTime) -= 1
                                End If
                            End If
                            .CreateTime = OJLogTemp.DateSubmit
                            If OJSysInfo.NewProblemCount.ContainsKey(OJLogTemp.DateSubmit) Then
                                OJSysInfo.NewProblemCount(OJLogTemp.DateSubmit) += 1
                            Else
                                OJSysInfo.NewProblemCount.Add(OJLogTemp.DateSubmit, 1)
                            End If
                        End If
                    End If
                End With
            Else
                ProblemTemp = New OJProblemInfo(OJLogTemp.ProblemID)
                ProblemDictionary(OJLogTemp.ProblemID) = ProblemTemp
                ProblemList.Add(OJLogTemp.ProblemID)
                With ProblemDictionary(OJLogTemp.ProblemID)
                    .SubmitCount += 1
                    If IsTestListMissing Then
                        .CreateTime = OJLogTemp.DateSubmit
                    Else
                        .ProblemTaskSequenceIndex = OJLogTemp.TestID
                        .CreateTime = TestDictionary(.ProblemTaskSequenceIndex).BeginDate
                    End If
                    If OJSysInfo.NewProblemCount.ContainsKey(.CreateTime) Then
                        OJSysInfo.NewProblemCount(.CreateTime) += 1
                    Else
                        OJSysInfo.NewProblemCount.Add(.CreateTime, 1)
                    End If
                    If OJLogTemp.IsPassed Then
                        .ACCount += 1
                    End If
                    If Not ProblemParticipantCheck(OJLogTemp.ProblemID).ContainsKey(OJLogTemp.StudentID) Then
                        .ParticipantCount += 1
                        ProblemParticipantCheck(OJLogTemp.ProblemID).Add(OJLogTemp.StudentID, True)
                    End If
                    If .SubmitCountByDay.ContainsKey(OJLogTemp.DateSubmit) Then
                        .SubmitCountByDay(OJLogTemp.DateSubmit) += 1
                    Else
                        .SubmitCountByDay(OJLogTemp.DateSubmit) = 1
                    End If
                End With
            End If
            '系统数据
            OJSysInfo.SubmitCountByHour(OJLogTemp.LogTime.Hour) += 1
            DoEvents()
        End While
        OJLogFileReader.Close()
        ProblemParticipantCheck.Clear()
        StudentList.Sort()
        ProblemList.Sort()
        '拟合计算
        '学生拟合
        If StudentList.Count >= 2 Then
            For i = 0 To StudentList.Count - 1
                With StudentDictionary(StudentList(i))
                    .FittingAC = IIf(.SubmitCount > 0, .ACCount + .ACCount / .SubmitCount, 0)
                    Dim DateLoop As New Date
                    DateLoop = UserSpecifiedAnalyzeStartDate
                    Dim DayLn() As Double
                    Dim SubmitAdjustedTotal() As Double
                    ReDim DayLn((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                    ReDim SubmitAdjustedTotal((UserSpecifiedAnalyzeEndDate - UserSpecifiedAnalyzeStartDate).Days - 1)
                    Dim StudentTotalSubmitSum As Integer = 0
                    Dim TotalProblems As Integer = 0
                    Dim Index = 0
                    While DateLoop < UserSpecifiedAnalyzeEndDate
                        If OJSysInfo.NewProblemCount.ContainsKey(DateLoop) Then
                            TotalProblems += OJSysInfo.NewProblemCount(DateLoop)
                        End If
                        If .SubmitCountByDay.ContainsKey(DateLoop) Then
                            StudentTotalSubmitSum += .SubmitCountByDay(DateLoop)
                        End If
                        DayLn(Index) = Math.Log((DateLoop - UserSpecifiedAnalyzeStartDate).Days + 1)
                        SubmitAdjustedTotal(Index) = StudentTotalSubmitSum - TotalProblems
                        DateLoop = DateLoop.AddDays(1)
                        Index += 1
                    End While
                    'Dim LinearFitResult As New Tuple(Of Double, Double)(0, 0)
                    'LinearFitResult = MathNet.Numerics.LinearRegression.SimpleRegression.Fit(DayLn, SubmitAdjustedTotal)
                    '.FittingB = LinearFitResult.Item1
                    '.FittingK_Kb = LinearFitResult.Item2
                    '.FittingR_Stb = MathNet.Numerics.Statistics.Correlation.Pearson(DayLn, SubmitAdjustedTotal) ^ 2
                    Try
                        Dim LinearFitReturn(2, 3) As Double
                        Dim LinearFitter As New LinearFit.Class1
                        Dim LinearFitX As New MWNumericArray(DayLn)
                        Dim LinearFitY As New MWNumericArray(SubmitAdjustedTotal)
                        LinearFitX = DayLn
                        LinearFitY = SubmitAdjustedTotal
                        LinearFitReturn = LinearFitter.LinearFit(LinearFitX, LinearFitY).ToArray()
                        LinearFitter.Dispose()
                        LinearFitX.Dispose()
                        LinearFitY.Dispose()
                        .FittingK_Kb = LinearFitReturn(0, 0)
                        .FittingR_Stb = LinearFitReturn(0, 1)
                        .FittingB = LinearFitReturn(0, 2)
                    Catch ex As Exception
                        IsMCRFailedStageStudentTotalSubmitByDayLnFit = True
                        .FittingB = 0
                        .FittingK_Kb = 0
                        .FittingR_Stb = 0
                    End Try
                End With
                DoEvents()
            Next
        Else
            IsSkippedStageStudentTotalSubmitByDayLnFit = True
        End If
        '题目拟合
        '产生题目参与人数-作业次序序列
        Dim ProblemCreateTimeList As New List(Of Date)
        Dim FirstProblemCreateTime As New Date
        Dim TaskSequence() As Integer
        Dim ParticipantCount() As Integer
        Dim ExponentialFittingA As Double = 0
        Dim ExponentialFittingB As Double = 0
        Dim ExponentialFittingC As Double = 0
        Dim ExponentialFitReturn(1, 3) As Double
        For i = 0 To ProblemList.Count - 1
            ProblemCreateTimeList.Add(ProblemDictionary(ProblemList(i)).CreateTime)
        Next
        FirstProblemCreateTime = ProblemCreateTimeList.Min()
        ReDim TaskSequence(ProblemList.Count - 1)
        ReDim ParticipantCount(ProblemList.Count - 1)
        For i = 0 To ProblemList.Count - 1
            With ProblemDictionary(ProblemList(i))
                If IsTestListMissing Then
                    .ProblemTaskSequenceIndex = Math.Round((.CreateTime - FirstProblemCreateTime).Days / 7) + 1
                End If
                TaskSequence(i) = .ProblemTaskSequenceIndex
                ParticipantCount(i) = .ParticipantCount
            End With
            DoEvents()
        Next
        If ProblemList.Count >= 3 Then
            Try
                Dim ExponentialFitter As New e_fit.Class1
                Dim ExponentialFitX As New MWNumericArray()
                Dim ExponentialFitY As New MWNumericArray()
                ExponentialFitX = TaskSequence
                ExponentialFitY = ParticipantCount
                ExponentialFitReturn = ExponentialFitter.e_fit(ExponentialFitX, ExponentialFitY).ToArray()
                ExponentialFitter.Dispose()
                ExponentialFitX.Dispose()
                ExponentialFitY.Dispose()
                ExponentialFittingA = -ExponentialFitReturn(0, 0)
                ExponentialFittingB = ExponentialFitReturn(0, 1)
                ExponentialFittingC = ExponentialFitReturn(0, 2)
            Catch ex As Exception
                IsMCRFailedStageProblemParticipantFit = True
                ExponentialFittingA = 0
                ExponentialFittingB = 0
                ExponentialFittingC = 0
            End Try
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                    .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                End With
                DoEvents()
            Next
        Else
            IsSkippedStageProblemParticipantFit = True
            ExponentialFittingA = 0
            ExponentialFittingB = 0
            ExponentialFittingC = 0
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    .EffortValue_Jq = IIf(.ACCount > 0, .SubmitCount / .ACCount, 0)
                    .ParticipateValue_Eq = .ParticipantCount - ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, .ProblemTaskSequenceIndex)
                End With
                DoEvents()
            Next
        End If
        '学生聚类分析
        If StudentList.Count >= 4 Then
            Dim StudentAcStbKbMatrix(,) As Double
            Dim StudentClustResult(,) As Double
            Dim AcAverage As Double = 0
            Dim AcStandardDivision As Double = 0
            Dim StbAvarage As Double = 0
            Dim StbStandardDivision As Double = 0
            Dim KbAverage As Double = 0
            Dim KbStandardDivision As Double = 0
            ReDim StudentAcStbKbMatrix(StudentList.Count - 1, 2)
            ReDim StudentClustResult(StudentList.Count - 1, 1)
            For i = 0 To StudentList.Count - 1
                With StudentDictionary(StudentList(i))
                    AcAverage += .FittingAC
                    StbAvarage += .FittingR_Stb
                    KbAverage += .FittingK_Kb
                End With
                DoEvents()
            Next
            AcAverage /= StudentList.Count
            StbAvarage /= StudentList.Count
            KbAverage /= StudentList.Count
            For i = 0 To StudentList.Count - 1
                With StudentDictionary(StudentList(i))
                    AcStandardDivision += (.FittingAC - AcAverage) ^ 2
                    StbStandardDivision += (.FittingR_Stb - StbAvarage) ^ 2
                    KbStandardDivision += (.FittingK_Kb - KbAverage) ^ 2
                End With
                DoEvents()
            Next
            AcStandardDivision /= StudentList.Count
            StbStandardDivision /= StudentList.Count
            KbStandardDivision /= StudentList.Count
            AcStandardDivision = Math.Sqrt(AcStandardDivision)
            StbStandardDivision = Math.Sqrt(StbStandardDivision)
            KbStandardDivision = Math.Sqrt(KbStandardDivision)
            For i = 0 To StudentList.Count - 1
                With StudentDictionary(StudentList(i))
                    StudentAcStbKbMatrix(i, 0) = (.FittingAC - AcAverage) / AcStandardDivision
                    StudentAcStbKbMatrix(i, 1) = (.FittingR_Stb - StbAvarage) / StbStandardDivision
                    StudentAcStbKbMatrix(i, 2) = (.FittingK_Kb - KbAverage) / KbStandardDivision
                End With
                DoEvents()
            Next
            Try
                Dim StudentClustExec As New cluster.Class1
                StudentClustResult = StudentClustExec.cluster(New MWNumericArray(StudentAcStbKbMatrix), 4).ToArray()
                StudentClustExec.Dispose()
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .ClustResult = StudentClustResult(i, 0)
                    End With
                    DoEvents()
                Next
                StudentClustResultMapping.Clear()
                StudentClustResultMapping.Add(0, "出错或被跳过")
                StudentClustResultMapping.Add(1, "未分级，点击以分级")
                StudentClustResultMapping.Add(2, "未分级，点击以分级")
                StudentClustResultMapping.Add(3, "未分级，点击以分级")
                StudentClustResultMapping.Add(4, "未分级，点击以分级")
            Catch ex As Exception
                IsMCRFailedStageStudentClust = True
                For i = 0 To StudentList.Count - 1
                    With StudentDictionary(StudentList(i))
                        .ClustResult = 0
                    End With
                    DoEvents()
                Next
                StudentClustResultMapping.Clear()
                StudentClustResultMapping.Add(0, "出错或被跳过")
                StudentClustResultMapping.Add(1, "未分级，点击以分级")
                StudentClustResultMapping.Add(2, "未分级，点击以分级")
                StudentClustResultMapping.Add(3, "未分级，点击以分级")
                StudentClustResultMapping.Add(4, "未分级，点击以分级")
            End Try
        Else
            IsSkippedStageStudentClust = True
            For i = 0 To StudentList.Count - 1
                With StudentDictionary(StudentList(i))
                    .ClustResult = 0
                End With
                DoEvents()
            Next
            StudentClustResultMapping.Clear()
            StudentClustResultMapping.Add(0, "出错或被跳过")
            StudentClustResultMapping.Add(1, "未分级，点击以分级")
            StudentClustResultMapping.Add(2, "未分级，点击以分级")
            StudentClustResultMapping.Add(3, "未分级，点击以分级")
            StudentClustResultMapping.Add(4, "未分级，点击以分级")
        End If
        '题目聚类分析
        '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
        If ProblemList.Count >= 4 Then
            Dim ProblemEqJqMatrix(,) As Double
            Dim ProblemClustResult(,) As Double
            Dim EqAverage As Double = 0
            Dim EqStandardDivision As Double = 0
            Dim JqAverage As Double = 0
            Dim JqStandardDivision As Double = 0
            ReDim ProblemEqJqMatrix(ProblemList.Count - 1, 1)
            ReDim ProblemClustResult(ProblemList.Count - 1, 1)
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    EqAverage += .ParticipateValue_Eq
                    JqAverage += .EffortValue_Jq
                End With
                DoEvents()
            Next
            EqAverage /= ProblemList.Count
            JqAverage /= ProblemList.Count
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    EqStandardDivision += (.ParticipateValue_Eq - EqAverage) ^ 2
                    JqStandardDivision += (.EffortValue_Jq - JqAverage) ^ 2
                End With
                DoEvents()
            Next
            EqStandardDivision /= ProblemList.Count
            JqStandardDivision /= ProblemList.Count
            EqStandardDivision = Math.Sqrt(EqStandardDivision)
            JqStandardDivision = Math.Sqrt(JqStandardDivision)
            For i = 0 To ProblemList.Count - 1
                With ProblemDictionary(ProblemList(i))
                    ProblemEqJqMatrix(i, 0) = (.EffortValue_Jq - JqAverage) / JqStandardDivision
                    ProblemEqJqMatrix(i, 1) = (.ParticipateValue_Eq - EqAverage) / EqStandardDivision
                End With
                DoEvents()
            Next
            Try
                Dim ProblemClustExec As New cluster.Class1
                ProblemClustResult = ProblemClustExec.cluster(New MWNumericArray(ProblemEqJqMatrix), 4).ToArray()
                ProblemClustExec.Dispose()
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        .ClustResult = ProblemClustResult(i, 0)
                    End With
                    DoEvents()
                Next
            Catch ex As Exception
                IsMCRFailedStageProblemClust = True
                For i = 0 To ProblemList.Count - 1
                    ProblemDictionary(ProblemList(i)).ClustResult = 0
                    DoEvents()
                Next
            End Try
        Else
            IsSkippedStageProblemClust = True
            For i = 0 To ProblemList.Count - 1
                ProblemDictionary(ProblemList(i)).ClustResult = 0
                DoEvents()
            Next
        End If
        '呈现列表
        lstStudentList.ItemsSource = StudentList
        lstProblemList.ItemsSource = ProblemList
        txtStartDate.Text = OJSysInfo.StartDate.ToLongDateString()
        txtEndDate.Text = OJSysInfo.EndDate.ToLongDateString()
        txtUserSpecifiedAnalyzeStartDate.Text = UserSpecifiedAnalyzeStartDate.ToLongDateString()
        txtUserSpecifiedAnalyzeEndDate.Text = UserSpecifiedAnalyzeEndDate.ToLongDateString()
        dtpUserSpecifiedAnalyzeStartDate.DisplayDateStart = OJSysInfo.StartDate
        dtpUserSpecifiedAnalyzeStartDate.DisplayDateEnd = OJSysInfo.EndDate
        dtpUserSpecifiedAnalyzeEndDate.DisplayDateStart = OJSysInfo.StartDate
        dtpUserSpecifiedAnalyzeEndDate.DisplayDateEnd = OJSysInfo.EndDate
        dtpUserSpecifiedAnalyzeStartDate.SelectedDate = UserSpecifiedAnalyzeStartDate
        dtpUserSpecifiedAnalyzeEndDate.SelectedDate = UserSpecifiedAnalyzeEndDate
        '题目参与人数与拟合图表
        If Not IsSkippedStageProblemParticipantFit Then
            If Not IsMCRFailedStageProblemParticipantFit Then
                Dim ProblemParticipantDataSource As New List(Of KeyValuePair(Of Integer, Integer))
                Dim ProblemParticipantFitDataSource As New List(Of KeyValuePair(Of Integer, Double))
                For i = 0 To ProblemList.Count - 1
                    ProblemParticipantDataSource.Add(New KeyValuePair(Of Integer, Integer)(TaskSequence(i), ParticipantCount(i)))
                    ProblemParticipantFitDataSource.Add(New KeyValuePair(Of Integer, Double)(TaskSequence(i), ExponentialCalculate(ExponentialFittingA, ExponentialFittingB, ExponentialFittingC, TaskSequence(i))))
                    DoEvents()
                Next
                sctProblemParticipant.ItemsSource = ProblemParticipantDataSource
                linProblemParticipantFit.ItemsSource = ProblemParticipantFitDataSource
                chartProblemParticipant.Visibility = Windows.Visibility.Visible
            End If
        End If
        '分时段总提交数图表
        Dim SubmitCountByTimeDataSource As New List(Of KeyValuePair(Of String, Integer))
        For i = 0 To 23
            SubmitCountByTimeDataSource.Add(New KeyValuePair(Of String, Integer)(i.ToString("00") & ":00" & vbCrLf & "~" & vbCrLf & (i + 1).ToString("00") & ":00", OJSysInfo.SubmitCountByHour(i)))
        Next
        colSubmitCountByTime.ItemsSource = SubmitCountByTimeDataSource
        chartSubmitCountByTime.Visibility = Windows.Visibility.Visible
        '题目发布曲线图表
        Dim j As Date
        j = UserSpecifiedAnalyzeStartDate
        Dim NewProblemCountByDayDataSource As New List(Of KeyValuePair(Of Date, Integer))
        While j <= UserSpecifiedAnalyzeEndDate
            If OJSysInfo.NewProblemCount.ContainsKey(j) Then
                NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, OJSysInfo.NewProblemCount(j)))
            Else
                NewProblemCountByDayDataSource.Add(New KeyValuePair(Of Date, Integer)(j, 0))
            End If
            j = j.AddDays(1)
        End While
        linNewProblemCountByDay.ItemsSource = NewProblemCountByDayDataSource
        chartNewProblemCountByDay.Visibility = Windows.Visibility.Visible
        '题目聚类分析结果
        '横轴为付出指数(Jq)，纵轴为参与指数(Eq)
        '聚类图表的四类依次序分为:
        '低Jq低Eq，低Jq高Eq，高Jq低Eq，高Jq高Eq
        If Not IsSkippedStageProblemClust Then
            If Not IsMCRFailedStageProblemClust Then
                Dim ScatterProblemCluster1DataSource As New List(Of KeyValuePair(Of Double, Double))
                Dim ScatterProblemCluster2DataSource As New List(Of KeyValuePair(Of Double, Double))
                Dim ScatterProblemCluster3DataSource As New List(Of KeyValuePair(Of Double, Double))
                Dim ScatterProblemCluster4DataSource As New List(Of KeyValuePair(Of Double, Double))
                For i = 0 To ProblemList.Count - 1
                    With ProblemDictionary(ProblemList(i))
                        If .ClustResult = 1 Then
                            ScatterProblemCluster1DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                        ElseIf .ClustResult = 2 Then
                            ScatterProblemCluster2DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                        ElseIf .ClustResult = 3 Then
                            ScatterProblemCluster3DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                        Else
                            ScatterProblemCluster4DataSource.Add(New KeyValuePair(Of Double, Double)(.EffortValue_Jq, .ParticipateValue_Eq))
                        End If
                    End With
                    DoEvents()
                Next
                sctProblemCluster1.ItemsSource = ScatterProblemCluster1DataSource
                sctProblemCluster2.ItemsSource = ScatterProblemCluster2DataSource
                sctProblemCluster3.ItemsSource = ScatterProblemCluster3DataSource
                sctProblemCluster4.ItemsSource = ScatterProblemCluster4DataSource
                chartProblemCluster.Visibility = Windows.Visibility.Visible
            End If
        End If
        Await AnalyzeProgress.CloseAsync()
        If IsMCRFailedStageStudentTotalSubmitByDayLnFit Or IsMCRFailedStageProblemParticipantFit Or IsMCRFailedStageStudentClust Or IsMCRFailedStageProblemClust Then
            Await Me.ShowMessageAsync("MATLAB Runtime发生问题", "本程序的运行依赖 MathWorks MATLAB Runtime R2018b，但是该组件可能发生问题或无法被调用。请前往 https://ww2.mathworks.cn/products/compiler/matlab-runtime.html 下载并安装适用于您的操作系统的 MathWorks MATLAB Runtime R2018b。如果您已经安装了 MathWorks MATLAB Runtime R2018b，您可能需要将其重新安装。存在边缘情况的数据集也可能引发此错误。" & vbCrLf & vbCrLf & "程序仍可继续运行，但部分功能可能发生异常。" & vbCrLf & vbCrLf & "以下功能依赖 MathWorks MATLAB Runtime R2018b: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsMCRFailedStageStudentTotalSubmitByDayLnFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsMCRFailedStageProblemParticipantFit, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsMCRFailedStageStudentClust, "执行时失败或发生错误。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsMCRFailedStageProblemClust, "执行时失败或发生错误。", "执行成功。"))
            IsMCRFailedStageStudentTotalSubmitByDayLnFit = False
            IsMCRFailedStageProblemParticipantFit = False
            IsMCRFailedStageStudentClust = False
            IsMCRFailedStageProblemClust = False
        End If
        If IsSkippedStageStudentTotalSubmitByDayLnFit Or IsSkippedStageProblemParticipantFit Or IsSkippedStageStudentClust Or IsSkippedStageProblemClust Then
            Await Me.ShowMessageAsync("一些分析步骤因数据集容量过小被跳过", "一部分拟合/聚类分析步骤被跳过，这可能是由数据集容量过小的问题导致的。" & vbCrLf & vbCrLf & "您可以尝试通过更新日志、扩大分析时间段长度等方式扩展数据集容量。" & vbCrLf & vbCrLf & "该问题会导致一部分图表或分析结果无法显示或显示的结果不正确。" & vbCrLf & vbCrLf & "以下分析流程可能已被跳过: " & vbCrLf & "学生对数日总提交曲线拟合: " & IIf(IsSkippedStageStudentTotalSubmitByDayLnFit, "已被跳过。", "执行成功。") & vbCrLf & "题目参与性拟合: " & IIf(IsSkippedStageProblemParticipantFit, "已被跳过。", "执行成功。") & vbCrLf & "学生聚类分析: " & IIf(IsSkippedStageStudentClust, "已被跳过。", "执行成功。") & vbCrLf & "题目聚类分析: " & IIf(IsSkippedStageProblemClust, "已被跳过。", "执行成功。"))
            IsSkippedStageStudentTotalSubmitByDayLnFit = False
            IsSkippedStageProblemParticipantFit = False
            IsSkippedStageStudentClust = False
            IsSkippedStageProblemClust = False
        End If
    End Sub

    Private Async Sub txtStudentClustResult_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtStudentClustResult.MouseUp
        If lstStudentList.SelectedIndex > -1 Then
            Dim MappingSettingValue As String
            Dim MappingKey As Integer = StudentDictionary(StudentList(lstStudentList.SelectedIndex)).ClustResult
            MappingSettingValue = Await ShowInputAsync("设置聚类数值结果" & Chr(34) & MappingKey.ToString() & Chr(34) & "的映射", "您可以在此设置聚类分析的数值结果" & Chr(34) & MappingKey.ToString() & Chr(34) & "的映射目标。所有聚类分析的数值结果" & Chr(34) & MappingKey.ToString() & Chr(34) & "在呈现时都会被映射到您设置的映射目标上。" & vbCrLf & vbCrLf & "推荐将映射目标设为有助于您分析相关学生的学习情况的字符串。")
            If MappingSettingValue = "" Then
                Exit Sub
            End If
            If StudentClustResultMapping.ContainsKey(MappingKey) Then
                StudentClustResultMapping(MappingKey) = MappingSettingValue
                With StudentDictionary(StudentList(lstStudentList.SelectedIndex))
                    txtStudentClustResult.Text = IIf(StudentClustResultMapping.ContainsKey(.ClustResult), StudentClustResultMapping(.ClustResult), .ClustResult)
                End With
            Else
                StudentClustResultMapping.Add(MappingKey, MappingSettingValue)
                With StudentDictionary(StudentList(lstStudentList.SelectedIndex))
                    txtStudentClustResult.Text = IIf(StudentClustResultMapping.ContainsKey(.ClustResult), StudentClustResultMapping(.ClustResult), .ClustResult)
                End With
            End If
        End If
    End Sub

    Private Async Sub mnuExportAnalyzedDataStudents_Click(sender As Object, e As RoutedEventArgs)
        Dim ExportParamDialog As New StudentDataExport
        Dim IsErrorOccurred As Boolean = False
        Dim ExceptionToDisplay As New Exception
        ExportParamDialog.ShowDialog()
        If ExportParamDialog.DialogResult = True Then
            Dim ExportProgress As Dialogs.ProgressDialogController
            ExportProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在导出学生数据", "正在导出学生数据到 " & StudentDataOutputParameters.OutputDirectory & "，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            DoEvents()
            Try
                Select Case StudentDataOutputParameters.OutputFileFormat
                    Case AnalyzedDataExportFormat.JSONFile
                        If Not DirectoryExists(StudentDataOutputParameters.OutputDirectory) Then
                            Directory.CreateDirectory(StudentDataOutputParameters.OutputDirectory)
                        End If
                        If StudentDataOutputParameters.IsOutputFileMerged Then
                            Dim StreamFileWriter As New StreamWriter(StudentDataOutputParameters.OutputDirectory & "StudentDataAnalyzed_" & BuildTimeStamp() & ".json", False)
                            Dim StudentOutputData As New MergedStudentDataExportFormat
                            With StudentOutputData
                                .AnalyzeStartDate = UserSpecifiedAnalyzeStartDate
                                .AnalyzeEndDate = UserSpecifiedAnalyzeEndDate
                                .IsDataMerged = True
                                For Each StudentID In StudentDataOutputParameters.DataToExport
                                    .StudentDataSet.Add(StudentDictionary(StudentID))
                                Next
                            End With
                            Dim JSONContent As String
                            JSONContent = JsonConvert.SerializeObject(StudentOutputData)
                            StreamFileWriter.WriteLine(JSONContent)
                            StreamFileWriter.Flush()
                            StreamFileWriter.Close()
                            StreamFileWriter.Dispose()
                        Else
                            For Each StudentID In StudentDataOutputParameters.DataToExport
                                Dim StreamFileWriter As New StreamWriter(StudentDataOutputParameters.OutputDirectory & "StudentDataAnalyzed_" & StudentID & "_" & BuildTimeStamp() & ".json", False)
                                Dim StudentOutputData As New SingleStudentDataExportFormat
                                With StudentOutputData
                                    .AnalyzeStartDate = UserSpecifiedAnalyzeStartDate
                                    .AnalyzeEndDate = UserSpecifiedAnalyzeEndDate
                                    .IsDataMerged = False
                                    .StudentDataSet = StudentDictionary(StudentID)
                                End With
                                Dim JSONContent As String
                                JSONContent = JsonConvert.SerializeObject(StudentOutputData)
                                StreamFileWriter.WriteLine(JSONContent)
                                StreamFileWriter.Flush()
                                StreamFileWriter.Close()
                                StreamFileWriter.Dispose()
                            Next
                        End If
                    Case AnalyzedDataExportFormat.XMLFile
                        Throw New FileFormatException("不支持导出为所选择的格式。")
                    Case Else
                        Throw New FileFormatException("不支持导出为所选择的格式。")
                End Select
            Catch ex As Exception
                IsErrorOccurred = True
                ExceptionToDisplay = ex
            End Try
            If Not IsErrorOccurred Then
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("学生数据导出成功", "学生数据已成功导出到 " & StudentDataOutputParameters.OutputDirectory & "。")
            Else
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("学生数据导出失败", "无法将学生数据导出到 " & StudentDataOutputParameters.OutputDirectory & "，发生错误：" & vbCrLf & ExceptionToDisplay.Message)
            End If
        End If
    End Sub

    Private Async Sub mnuExportAnalyzedDataProblems_Click(sender As Object, e As RoutedEventArgs)
        Dim ExportParamDialog As New ProblemDataExport
        Dim IsErrorOccurred As Boolean = False
        Dim ExceptionToDisplay As New Exception
        ExportParamDialog.ShowDialog()
        If ExportParamDialog.DialogResult = True Then
            Dim ExportProgress As Dialogs.ProgressDialogController
            ExportProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在导出题目数据", "正在导出题目数据到 " & ProblemDataOutputParameters.OutputDirectory & "，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            DoEvents()
            Try
                Select Case ProblemDataOutputParameters.OutputFileFormat
                    Case AnalyzedDataExportFormat.JSONFile
                        If Not DirectoryExists(ProblemDataOutputParameters.OutputDirectory) Then
                            Directory.CreateDirectory(ProblemDataOutputParameters.OutputDirectory)
                        End If
                        If ProblemDataOutputParameters.IsOutputFileMerged Then
                            Dim StreamFileWriter As New StreamWriter(ProblemDataOutputParameters.OutputDirectory & "ProblemDataAnalyzed_" & BuildTimeStamp() & ".json", False)
                            Dim ProblemOutputData As New MergedProblemDataExportFormat
                            With ProblemOutputData
                                .AnalyzeStartDate = UserSpecifiedAnalyzeStartDate
                                .AnalyzeEndDate = UserSpecifiedAnalyzeEndDate
                                .IsDataMerged = True
                                For Each ProblemID In ProblemDataOutputParameters.DataToExport
                                    .ProblemDataSet.Add(ProblemDictionary(ProblemID))
                                Next
                            End With
                            Dim JSONContent As String
                            JSONContent = JsonConvert.SerializeObject(ProblemOutputData)
                            StreamFileWriter.WriteLine(JSONContent)
                            StreamFileWriter.Flush()
                            StreamFileWriter.Close()
                            StreamFileWriter.Dispose()
                        Else
                            For Each ProblemID In ProblemDataOutputParameters.DataToExport
                                Dim StreamFileWriter As New StreamWriter(ProblemDataOutputParameters.OutputDirectory & "ProblemDataAnalyzed_" & ProblemID & "_" & BuildTimeStamp() & ".json", False)
                                Dim ProblemOutputData As New SingleProblemDataExportFormat
                                With ProblemOutputData
                                    .AnalyzeStartDate = UserSpecifiedAnalyzeStartDate
                                    .AnalyzeEndDate = UserSpecifiedAnalyzeEndDate
                                    .IsDataMerged = False
                                    .ProblemDataSet = ProblemDictionary(ProblemID)
                                End With
                                Dim JSONContent As String
                                JSONContent = JsonConvert.SerializeObject(ProblemOutputData)
                                StreamFileWriter.WriteLine(JSONContent)
                                StreamFileWriter.Flush()
                                StreamFileWriter.Close()
                                StreamFileWriter.Dispose()
                            Next
                        End If
                    Case AnalyzedDataExportFormat.XMLFile
                        Throw New FileFormatException("不支持导出为所选择的格式。")
                    Case Else
                        Throw New FileFormatException("不支持导出为所选择的格式。")
                End Select
            Catch ex As Exception
                IsErrorOccurred = True
                ExceptionToDisplay = ex
            End Try
            If Not IsErrorOccurred Then
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("题目数据导出成功", "题目数据已成功导出到 " & ProblemDataOutputParameters.OutputDirectory & "。")
            Else
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("题目数据导出失败", "无法将题目数据导出到 " & ProblemDataOutputParameters.OutputDirectory & "，发生错误：" & vbCrLf & ExceptionToDisplay.Message)
            End If
        End If
    End Sub

    Private Async Sub mnuExportAnalyzedDataSystem_Click(sender As Object, e As RoutedEventArgs)
        Dim IsErrorOccurred As Boolean = False
        Dim ExceptionToDisplay As New Exception
        Dim FileSaveDialog As New CommonSaveFileDialog
        GenerateCurrentDirectory()
        Dim CurrentPath As String = GetCurrentDirectory()
        With FileSaveDialog
            .Title = "请指定文件的导出位置"
            .DefaultDirectory = CurrentPath
            .DefaultFileName = "OJSystemDataAnalyzed_" & BuildTimeStamp() & ".json"
            .OverwritePrompt = True
            .NavigateToShortcut = True
            .EnsurePathExists = True
            .Filters.Clear()
            .Filters.Add(New CommonFileDialogFilter("JSON 文件", ".json"))
            .Filters.Add(New CommonFileDialogFilter("所有文件", ".*"))
        End With
        If FileSaveDialog.ShowDialog() = CommonFileDialogResult.Ok Then
            Dim ExportProgress As Dialogs.ProgressDialogController
            ExportProgress = Await Dialogs.DialogManager.ShowProgressAsync(Me, "正在导出系统数据", "正在导出系统数据到 " & FileSaveDialog.FileName & "，这可能需要几分钟的时间。" & vbCrLf & vbCrLf & "请勿关闭应用程序。", False)
            DoEvents()
            Try
                Dim StreamFileWriter As New StreamWriter(FileSaveDialog.FileName, False)
                Dim OJSysInfoExport As New OJSystemDataExportFormat
                With OJSysInfoExport
                    .LogStartDate = OJSysInfo.StartDate
                    .LogEndDate = OJSysInfo.EndDate
                    .AnalyzeStartDate = UserSpecifiedAnalyzeStartDate
                    .AnalyzeEndDate = UserSpecifiedAnalyzeEndDate
                    .NewProblemCount = OJSysInfo.NewProblemCount
                    .SubmitCountByHour = OJSysInfo.SubmitCountByHour
                End With
                Dim JSONContent As String = JsonConvert.SerializeObject(OJSysInfoExport)
                StreamFileWriter.WriteLine(JSONContent)
                StreamFileWriter.Flush()
                StreamFileWriter.Close()
                StreamFileWriter.Dispose()
            Catch ex As Exception
                IsErrorOccurred = True
                ExceptionToDisplay = ex
            End Try
            If Not IsErrorOccurred Then
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("系统数据导出成功", "系统数据已成功导出到 " & FileSaveDialog.FileName & "。")
            Else
                Await ExportProgress.CloseAsync()
                System.Threading.Thread.Sleep(500)
                Await Me.ShowMessageAsync("系统数据导出失败", "无法将系统数据导出到 " & ProblemDataOutputParameters.OutputDirectory & "，发生错误：" & vbCrLf & ExceptionToDisplay.Message)
            End If
        End If
    End Sub

    Private Sub mnuExportAnalyzedData_Click(sender As Object, e As RoutedEventArgs) Handles mnuExportAnalyzedData.Click
        mnuExportAnalyzedData.ContextMenu.Placement = Primitives.PlacementMode.Bottom
        mnuExportAnalyzedData.ContextMenu.PlacementTarget = mnuExportAnalyzedData
        mnuExportAnalyzedData.ContextMenu.IsOpen = True
    End Sub
End Class
