Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Octokit

Class UpdateChecker
    Public Event UpdateFound(newVersion As String, url As String)

    Public Async Sub CheckForUpdates(user As String, repo As String, Optional tries As UInteger = 3)
        Try
            Dim currentVersion = Assembly.GetExecutingAssembly().GetName().Version
            Dim currentMajor = currentVersion.Major
            Dim currentMinor = currentVersion.Minor

            Dim kit = New GitHubClient(New ProductHeaderValue("thirtythreeforty"))
            Dim releases = Await kit.Repository.Release.GetAll(user, repo)

            For Each release As Release In releases
                Dim m = New Regex("(\d+)\.(\d+)").Match(release.Name)
                If m.Groups.Count >= 2 Then
                    Try
                        Dim newMajor = Integer.Parse(m.Groups.Item(1).ToString, 10)
                        Dim newMinor = Integer.Parse(m.Groups.Item(2).ToString, 10)

                        If release.Prerelease = True AndAlso
                                ((currentMajor < newMajor) OrElse
                                 (currentMajor = newMajor AndAlso currentMinor < newMinor)) Then
                            RaiseEvent UpdateFound(release.Name, release.HtmlUrl)
                            Return
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next

        Catch ex As Exception
            If tries > 0 Then
                CheckForUpdates(user, repo, tries - 1)
            End If
        End Try
    End Sub
End Class
