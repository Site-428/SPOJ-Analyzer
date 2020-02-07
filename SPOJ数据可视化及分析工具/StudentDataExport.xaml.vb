Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Windows.Controls
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Windows.Forms
Imports System.Windows.Window
Imports System.IO
Imports Microsoft.WindowsAPICodePack.Dialogs
Public Class StudentDataExport
    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        DialogResult = False
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowse.Click
        Dim FolderBrowser As New CommonOpenFileDialog
        Dim ExportPath As New String("")
        GenerateCurrentDirectory()
        Dim CurrentPath As String = GetCurrentDirectory()
        With FolderBrowser
            .Title = "请指定文件的导出位置"
            .IsFolderPicker = True
            .DefaultDirectory = CurrentPath
            .AllowNonFileSystemItems = False
        End With
        If FolderBrowser.ShowDialog() = CommonFileDialogResult.Ok Then
            ExportPath = FolderBrowser.FileName
            If ExportPath(ExportPath.Length - 1) <> "\" Then
                ExportPath = ExportPath & "\"
            End If
            txtExportPath.Text = ExportPath
        End If
        Me.Activate()
    End Sub

    Private Sub StudentDataExport_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        GenerateCurrentDirectory()
        Dim CurrentPath As String = GetCurrentDirectory()
        txtExportPath.Text = CurrentPath
        lstDataToExport.ItemsSource = StudentList
        lstDataToExport.SelectAll()
        If lstDataToExport.Items.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
        If lstDataToExport.SelectedItems.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
    End Sub

    Private Sub btnSelectAll_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnSelectAll.MouseLeftButtonUp
        lstDataToExport.SelectAll()
        If lstDataToExport.SelectedItems.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
    End Sub

    Private Sub btnDeselectAll_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnDeselectAll.MouseLeftButtonUp
        lstDataToExport.UnselectAll()
        If lstDataToExport.SelectedItems.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
    End Sub

    Private Sub btnInvertSelect_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles btnInvertSelect.MouseLeftButtonUp
        For Each ListDataItem In lstDataToExport.Items
            If lstDataToExport.SelectedItems.Contains(ListDataItem) Then
                lstDataToExport.SelectedItems.Remove(ListDataItem)
            Else
                lstDataToExport.SelectedItems.Add(ListDataItem)
            End If
        Next
        If lstDataToExport.SelectedItems.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
    End Sub

    Private Sub lstDataToExport_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstDataToExport.SelectionChanged
        If lstDataToExport.SelectedItems.Count = 0 Then
            btnOK.IsEnabled = False
        Else
            btnOK.IsEnabled = True
        End If
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click
        With StudentDataOutputParameters
            .OutputDirectory = txtExportPath.Text
            .OutputFileFormat = AnalyzedDataExportFormat.JSONFile
            .IsOutputFileMerged = chkIsDataMerged.IsChecked
            .DataToExport.Clear()
            For Each SelectedItemString As String In lstDataToExport.SelectedItems
                .DataToExport.Add(SelectedItemString)
            Next
        End With
        DialogResult = True
    End Sub
End Class
