Imports Microsoft.Win32
Imports Ookii.Dialogs.Wpf

Class MainWindow
    Private Sub keyTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles keyTextBox.TextChanged

    End Sub

    Private Sub chooseKeyButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles chooseKeyButton.Click
        Dim dlg = New OpenFileDialog
        dlg.Filter = "Word Documents (*.docx, *.doc)|*.docx;*.doc|All Files|*"
        If dlg.ShowDialog Then
            keyTextBox.Text = dlg.FileName
        End If
    End Sub

    Private Sub chooseFolderButton_Click(sender As Object, e As RoutedEventArgs) Handles chooseFolderButton.Click
        Dim dlg = New VistaFolderBrowserDialog
        If dlg.ShowDialog Then
            folderTextBox.Text = dlg.SelectedPath
        End If
    End Sub
End Class
