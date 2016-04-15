Imports Microsoft.Win32
Imports Ookii.Dialogs.Wpf
Imports System.IO

Class MainWindow
    Private worker = New WordWorker
    Private currentFile As String

    Private Sub chooseKeyButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles chooseKeyButton.Click
        Dim dlg = New OpenFileDialog
        dlg.Filter = "Word Documents (*.docx, *.doc)|*.docx;*.doc|All Files|*"
        If dlg.ShowDialog Then
            keyTextBox.Text = dlg.FileName
        End If
    End Sub

    Private Sub chooseFolderButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles chooseFolderButton.Click
        Dim dlg = New VistaFolderBrowserDialog
        If dlg.ShowDialog Then
            folderTextBox.Text = dlg.SelectedPath
        End If
    End Sub

    Private Sub compareFile(file As String)
        Dim key = keyTextBox.Text
        If My.Computer.FileSystem.FileExists(key) AndAlso
           My.Computer.FileSystem.FileExists(file) Then
            worker.compare(file)
        End If
    End Sub

    Private Sub nextButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles nextButton.Click
        filesComboBox.SelectedIndex += 1
    End Sub

    Private Sub prevButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles prevButton.Click
        filesComboBox.SelectedIndex -= 1
    End Sub

    Private Async Sub filesComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles filesComboBox.SelectionChanged
        updatePrevNext()

        Dim sel = filesComboBox.SelectedItem

        If sel IsNot Nothing Then
            comparingStatus()
            ' Re-join the filename with the path name, since it was stripped
            ' to put it in the combobox
            Dim fullPath = Path.Combine(folderTextBox.Text, sel)
            Dim key = keyTextBox.Text
            Await Task.Run(Sub()
                               worker.compare(key, fullPath)
                           End Sub)
            updateStatus()
        End If
    End Sub

    Private Sub folderTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles folderTextBox.TextChanged
        findWordFiles(folderTextBox.Text)
        updateStatus()
    End Sub

    Private Sub findWordFiles(dir As String)
        filesComboBox.Items.Clear()

        Try
            Dim files = Directory.EnumerateFiles(folderTextBox.Text)
            For Each file As String In files
                Dim ext = Path.GetExtension(file)
                Dim rootName = Path.GetFileName(file)
                If ext = ".docx" Or ext = ".doc" Then
                    filesComboBox.Items.Add(rootName)
                End If
            Next
        Catch ex As Exception
            ' Just let it pass, probably a bad directory name
        End Try

        filesComboBox.SelectedIndex = -1
        updatePrevNext()
    End Sub

    Private Sub updatePrevNext()
        Dim idx = filesComboBox.SelectedIndex
        Dim keyPresent = My.Computer.FileSystem.FileExists(keyTextBox.Text)

        prevButton.IsEnabled = (idx > 0 And keyPresent)
        nextButton.IsEnabled = (idx < filesComboBox.Items.Count - 1 And keyPresent)
    End Sub

    Private Sub updateStatus()
        If Not My.Computer.FileSystem.FileExists(keyTextBox.Text) Then
            statusTextBlock.Text = "Choose key"
        ElseIf Not My.Computer.FileSystem.DirectoryExists(folderTextBox.Text) Then
            statusTextBlock.Text = "Choose folder"
        Else
            statusTextBlock.Text = "Ready"
        End If
    End Sub

    Private Sub comparingStatus()
        statusTextBlock.Text = "Comparing..."
    End Sub

    Private Sub keyTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles keyTextBox.TextChanged
        updateStatus()
        updatePrevNext()
    End Sub

    Private Sub aboutButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles aboutButton.Click
        Dim dlg = New TaskDialog
        dlg.MainInstruction = "Word Comparison Tool"
        dlg.Content =
            "By George Hilliard (thirtythreeforty@gmail.com)" + Environment.NewLine +
            "Last updated 15 April 2016" + Environment.NewLine +
            Environment.NewLine +
            "This program is free software: you can redistribute it and/or modify " +
            "it under the terms of the GNU General Public License as published by " +
            "the Free Software Foundation, either version 3 of the License, or " +
            "(at your option) any later version."
        dlg.WindowTitle = "About"
        dlg.Buttons.Add(New TaskDialogButton(ButtonType.Ok))
        dlg.ShowDialog()
    End Sub
End Class
