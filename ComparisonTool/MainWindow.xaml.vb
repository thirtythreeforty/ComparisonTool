Imports Microsoft.Win32
Imports Ookii.Dialogs.Wpf
Imports System.IO

Class MainWindow
    Private worker = New OfficeHandler
    Private WithEvents checker As UpdateChecker = New UpdateChecker

    Private Sub chooseKeyButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles chooseKeyButton.Click
        Dim dlg = New OpenFileDialog
        dlg.Filter =
            "All supported documents (*.docx, *.doc; *.xlsx, *.xls; *.pptx, *.ppt)|*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt|" +
            "Word documents (*.docx, *.doc)|*.docx;*.doc|" +
            "Excel documents (*.xlsx, *.xls)|*.xlsx;*.xls|" +
            "PowerPoint documents (*.pptx, *.ppt)|*.pptx;*.ppt|" +
            "All Files|*"
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

    Private Sub nextButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles nextButton.Click
        filesComboBox.SelectedIndex += 1
    End Sub

    Private Sub prevButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles prevButton.Click
        filesComboBox.SelectedIndex -= 1
    End Sub

    Private Sub filesComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles filesComboBox.SelectionChanged
        CompareCurrentSelection()
    End Sub

    Private Sub folderTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles folderTextBox.TextChanged
        FindSimilarFiles()
    End Sub

    Private Sub FindSimilarFiles()
        filesComboBox.Items.Clear()

        If My.Computer.FileSystem.FileExists(keyTextBox.Text) Then
            Try
                Dim type = OfficeHandler.FindFileType(keyTextBox.Text)
                Dim files = Directory.EnumerateFiles(folderTextBox.Text)
                For Each file As String In files
                    If OfficeHandler.FileTypeMatches(file, type) AndAlso
                            file <> keyTextBox.Text Then
                        Dim rootName = Path.GetFileName(file)
                        filesComboBox.Items.Add(rootName)
                    End If
                Next
            Catch ex As Exception
                ' Just let it pass, probably a bad directory name
            End Try
        End If

        filesComboBox.SelectedIndex = -1
        UpdatePrevNext()
        UpdateStatus()
    End Sub

    Private Sub UpdatePrevNext()
        Dim idx = filesComboBox.SelectedIndex
        Dim keyPresent = My.Computer.FileSystem.FileExists(keyTextBox.Text)

        prevButton.IsEnabled = (idx > 0 And keyPresent)
        nextButton.IsEnabled = (idx < filesComboBox.Items.Count - 1 And keyPresent)
    End Sub

    Private Sub UpdateStatus()
        Dim numFound = filesComboBox.Items.Count
        If Not My.Computer.FileSystem.FileExists(keyTextBox.Text) Then
            statusTextBlock.Text = "Choose key"
        ElseIf Not My.Computer.FileSystem.DirectoryExists(folderTextBox.Text) Then
            statusTextBlock.Text = "Choose folder"
        ElseIf numFound <= 0 Then
            statusTextBlock.Text = "No matching files"
        Else
            statusTextBlock.Text = String.Format("Ready ({0} files)", numFound)
        End If
    End Sub

    Private Sub ComparingStatus()
        statusTextBlock.Text = "Comparing..."
    End Sub

    Private Sub keyTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles keyTextBox.TextChanged
        FindSimilarFiles()
    End Sub

    Private Sub aboutButton_Click(sender As Object, e As RoutedEventArgs) _
            Handles aboutButton.Click
        Dim dlg = New TaskDialog
        dlg.MainInstruction = "Office Comparison Tool"
        dlg.Content =
            "By George Hilliard (thirtythreeforty@gmail.com)" + Environment.NewLine +
            "Last updated 16 April 2016" + Environment.NewLine +
            Environment.NewLine +
            "This program is free software: you can redistribute it and/or modify " +
            "it under the terms of the GNU General Public License as published by " +
            "the Free Software Foundation, either version 3 of the License, or " +
            "(at your option) any later version."
        dlg.WindowTitle = "About"
        dlg.Buttons.Add(New TaskDialogButton(ButtonType.Ok))
        dlg.ShowDialog()
    End Sub

    Private Async Sub CompareCurrentSelection()
        UpdatePrevNext()

        Dim sel = filesComboBox.SelectedItem

        If sel IsNot Nothing Then
            ComparingStatus()
            ' Re-join the filename with the path name, since it was stripped
            ' to put it in the combobox
            Dim fullPath = Path.Combine(folderTextBox.Text, sel)
            Dim key = keyTextBox.Text
            Await Task.Run(Sub()
                               worker.Compare(key, fullPath)
                           End Sub)
            UpdateStatus()
        End If
    End Sub

    Private Sub checker_OnUpdateFound(version As String, url As String) _
                Handles checker.UpdateFound
        Dim dlg = New TaskDialog
        dlg.MainInstruction = "Update Available"
        dlg.Content =
            "A new update, " + version + ", of this program is available! Please visit the project page to download it."
        dlg.WindowTitle = "Office Comparison Tool"
        Dim downloadBtn = New TaskDialogButton("Go to Download Page")
        dlg.Buttons.Add(downloadBtn)
        dlg.Buttons.Add(New TaskDialogButton("Not Now"))
        If dlg.ShowDialog() Is downloadBtn Then
            Process.Start(url)
        End If
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        checker.CheckForUpdates("thirtythreeforty", "ComparisonTool")
    End Sub
End Class
