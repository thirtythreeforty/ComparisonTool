Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text.RegularExpressions

Public Class OfficeHandler
    Public Sub Compare(fileName As String, toCompare As String)
        GetComparator(fileName)(fileName, toCompare)
    End Sub

    Public Enum FileType
        Word
        Excel
        PowerPoint
        Unknown
    End Enum

    Public Shared Function FileTypeMatches(file As String, type As FileType) As Boolean
        Return FindFileType(file) = type AndAlso type <> FileType.Unknown
    End Function

    Public Shared Function FileTypeMatches(file1 As String, file2 As String) As Boolean
        Dim file1type = FindFileType(file1)
        Return file1type = FindFileType(file2) AndAlso file1type <> FileType.Unknown
    End Function

    Public Shared Function FindFileType(file As String) As FileType
        If MatchesStrings(file, "doc", "docx") Then
            Return FileType.Word
        ElseIf MatchesStrings(file, "xls", "xlsx") Then
            Return FileType.Excel
        ElseIf MatchesStrings(file, "ppt", "pptx") Then
            Return FileType.PowerPoint
        Else
            Return FileType.Unknown
        End If
    End Function

    Private Shared Function MatchesStrings(file As String, s1 As String, s2 As String) As Boolean
        Dim ext = Path.GetExtension(file)
        Return ext.ToLower().EndsWith(s1) OrElse ext.EndsWith(s2)
    End Function

    Private Delegate Sub Comparator(base As String, fork As String)

    Private Function GetComparator(file As String) As Comparator
        If FileTypeMatches(file, FileType.Word) Then
            Return AddressOf WordCompare
        ElseIf FileTypeMatches(file, FileType.Excel) Then
            Return AddressOf ExcelCompare
        ElseIf FileTypeMatches(file, FileType.PowerPoint) Then
            Return AddressOf PowerPointCompare
        Else
            Return Sub(x As String, y As String)
                   End Sub
        End If
    End Function

    Private Sub WordCompare(base As String, fork As String)
        Dim app As Word.Application
        Try
            app = Marshal.GetActiveObject("Word.Application")
        Catch ex As Exception
            app = New Word.Application
        End Try
        Dim doc = app.Documents.Open(base)
        If doc IsNot Nothing Then
            doc.Compare(fork)
            doc.Close()
            app.Visible = True
        End If
    End Sub

    Private excelCompareBinary = New Lazy(Of String)(
        Function() As String
            Dim officeLocation = Path.Combine(
                My.Computer.FileSystem.SpecialDirectories.ProgramFiles,
                "Microsoft Office"
            )
            Dim foundCompares =
                My.Computer.FileSystem.GetFiles(officeLocation,
                                                FileIO.SearchOption.SearchAllSubDirectories,
                                                "SPREADSHEETCOMPARE.EXE")
            If foundCompares.Count = 0 Then
                Return Nothing
            End If
            Return foundCompares.ElementAt(0)
        End Function)
    Private Sub ExcelCompare(base As String, fork As String)
        ' Excel does not expose its Inquire add-in that provides spreadsheet
        ' diff via an API.  We must call it manually as a standalone program. See
        ' http://stackoverflow.com/questions/13702663 for details.
        Dim compareBinary = excelCompareBinary.Value
        If compareBinary Is Nothing Then
            MessageBox.Show("Inquire Add-In for Excel seems to be missing! " +
                            "This is only shipped with Professional Plus versions of Office :(")
            Return
        End If

        ' So this is new levels of horrifying.  The arguments are not passed via the command
        ' line -- that would be too straightforward.  No, they are instead written to a file
        ' which is then passed to the comparison program, which deletes the (hopefully
        ' temporary) argument file.  The comparison program crashes if anything is out of place.
        Dim tempFileLocation = Path.Combine(
            My.Computer.FileSystem.SpecialDirectories.Temp,
            "ComparisonToolExcel" + Rnd().ToString() + ".txt"
        )
        Dim compareList = base + Environment.NewLine + fork
        My.Computer.FileSystem.WriteAllText(tempFileLocation, compareList, False)

        Dim proc = Process.Start(compareBinary, EscapePathForShell(tempFileLocation))
    End Sub

    Private Sub PowerPointCompare(base As String, fork As String)
        Dim app As PowerPoint.Application
        Try
            app = Marshal.GetActiveObject("PowerPoint.Application")
        Catch ex As Exception
            app = New PowerPoint.Application
        End Try
        app.Visible = True
        Dim doc = app.Presentations.Open(base)
        doc.Merge(fork)
    End Sub

    Private Shared Function EscapePathForShell(s As String) As String
        Return """" + Regex.Replace(s, "(\\+)$", "$1$1") + """"
    End Function
End Class
