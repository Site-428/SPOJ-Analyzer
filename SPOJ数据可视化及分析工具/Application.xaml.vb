Class Application

    ' 應用程式層級事件 (例如 Startup、Exit 和 DispatcherUnhandledException)
    ' 可以在此檔案中處理。

    Private Sub Application_DispatcherUnhandledException(sender As Object, e As Windows.Threading.DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException
        e.Handled = True
        MessageBox.Show("发生运行时错误: " & vbCrLf & vbCrLf & e.Exception.Message & vbCrLf & vbCrLf & "-----" & vbCrLf & vbCrLf & "栈痕迹追踪如下: " & vbCrLf & vbCrLf & e.Exception.StackTrace, "错误", MessageBoxButton.OK, MessageBoxImage.Error)
    End Sub
End Class
