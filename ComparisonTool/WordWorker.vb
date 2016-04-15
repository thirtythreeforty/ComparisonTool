Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class WordWorker
    Private doc As Word.Document

    Public Sub compare(fileName As String, toCompare As String)
        Dim word = getWordInstance()
        doc = word.Documents.Open(fileName)
        If doc IsNot Nothing Then
            doc.Compare(toCompare)
            word.Visible = True
            doc.Close()
        End If
    End Sub

    Public Sub stopComparingCurrent()
        If doc IsNot Nothing Then
            doc.Close()
        End If
    End Sub

    Private Function getWordInstance() As Word.Application
        Try
            Return Marshal.GetActiveObject("Word.Application")
        Catch ex As Exception
            Return New Word.Application
        End Try
    End Function
End Class
