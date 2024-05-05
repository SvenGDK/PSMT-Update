Imports System.ComponentModel
Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Windows.Threading

Class MainWindow

    Private WithEvents WaitTimer As New DispatcherTimer With {.Interval = New TimeSpan(0, 0, 1)}
    Private WithEvents FinishWaitTimer As New DispatcherTimer With {.Interval = New TimeSpan(0, 0, 1)}
    Private WaitedFor As Integer = 0

    Private Sub MainWindow_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        WaitTimer.Start()

        UpdateTextBlock.Text = "Getting Update Package, please wait."
        UpdateProgressBar.IsIndeterminate = False
    End Sub

    Private Sub WaitTimer_Tick(sender As Object, e As EventArgs) Handles WaitTimer.Tick
        WaitedFor += 1

        If WaitedFor = 2 Then
            WaitTimer.Stop()
            WaitedFor = 0

            GetUpdatePackage()
        End If
    End Sub

    Private Sub FinishWaitTimer_Tick(sender As Object, e As EventArgs) Handles FinishWaitTimer.Tick
        WaitedFor += 1

        If WaitedFor = 2 Then
            WaitTimer.Stop()
            StartPSMultiTools()
        End If
    End Sub

    Private Sub GetUpdatePackage()
        If IsURLValid("https://raw.githubusercontent.com/SvenGDK/PS-Multi-Tools/main/Update.zip") Then

            Dim NewWebClient As New WebClient()
            NewWebClient.DownloadFileAsync(New Uri("https://raw.githubusercontent.com/SvenGDK/PS-Multi-Tools/main/Update.zip"), My.Computer.FileSystem.CurrentDirectory + "\Update.zip", Stopwatch.StartNew)

            AddHandler NewWebClient.DownloadProgressChanged, Sub(sender As Object, e As DownloadProgressChangedEventArgs)
                                                                 'Update values
                                                                 Dim ClientSender As WebClient = CType(sender, WebClient)

                                                                 If Dispatcher.CheckAccess() = False Then
                                                                     Dispatcher.BeginInvoke(Sub()
                                                                                                UpdateTextBlock.Text = "Status: " + (e.BytesReceived / (1024 * 1024)).ToString("0.000 MB") + "/" + (e.TotalBytesToReceive / (1024 * 1024)).ToString("0.000 MB")
                                                                                                UpdateProgressBar.Value = e.ProgressPercentage
                                                                                            End Sub)
                                                                 Else
                                                                     UpdateTextBlock.Text = "Status: " + (e.BytesReceived / (1024 * 1024)).ToString("0.000 MB") + "/" + (e.TotalBytesToReceive / (1024 * 1024)).ToString("0.000 MB")
                                                                     UpdateProgressBar.Value = e.ProgressPercentage
                                                                 End If

                                                             End Sub

            AddHandler NewWebClient.DownloadFileCompleted, Sub(sender As Object, e As AsyncCompletedEventArgs)
                                                               Dim ClientSender As WebClient = CType(sender, WebClient)

                                                               If Dispatcher.CheckAccess() = False Then
                                                                   Dispatcher.BeginInvoke(Sub()
                                                                                              UpdateTextBlock.Text = "Update package downloaded. Installing now ..."
                                                                                              UpdateProgressBar.Value = 0

                                                                                              InstallUpdate()
                                                                                          End Sub)
                                                               Else
                                                                   UpdateTextBlock.Text = "Update package downloaded. Installing now ..."
                                                                   UpdateProgressBar.Value = 0

                                                                   InstallUpdate()
                                                               End If

                                                           End Sub
        Else
            MsgBox("Error while updating PS Multi Tools. Description: Could not find the latest Update package." + vbCrLf + "Please check your internet connection.", MsgBoxStyle.Critical, "Error while updating")
            Close()
        End If
    End Sub

    Private Sub InstallUpdate()
        If File.Exists(My.Computer.FileSystem.CurrentDirectory + "\Update.zip") Then
            UpdateProgressBar.Value = 25
            Try
                Using UpdateArchive As ZipArchive = ZipFile.OpenRead(My.Computer.FileSystem.CurrentDirectory + "\Update.zip")
                    For Each ArchiveEntry As ZipArchiveEntry In UpdateArchive.Entries

                        Dim EntryFullname = Path.Combine(My.Computer.FileSystem.CurrentDirectory, ArchiveEntry.FullName)
                        Dim EntryPath = Path.GetDirectoryName(EntryFullname)

                        If Not Directory.Exists(EntryPath) Then
                            Directory.CreateDirectory(EntryPath)
                        End If

                        Dim EntryFn = Path.GetFileName(EntryFullname)

                        If Not String.IsNullOrEmpty(EntryFn) Then
                            ArchiveEntry.ExtractToFile(EntryFullname, True)
                        End If
                    Next
                End Using
            Catch ex As Exception
                MsgBox("Error while updating PS Multi Tools. Description: " + ex.Message, MsgBoxStyle.Critical, "Error updating PS Multi Tools")
                Close()
            End Try

            UpdateTextBlock.Text = "Update installed successfully. PS Multi Tools will start in 2 sec."
            UpdateProgressBar.Value += 75

            FinishWaitTimer.Start()
        Else
            MsgBox("Error while updating PS Multi Tools. Description: " + vbCrLf + "File " + vbCrLf + My.Computer.FileSystem.CurrentDirectory + "\Update.zip" + vbCrLf + "not found.", MsgBoxStyle.Critical, "Error while updating")
            Close()
        End If
    End Sub

    Private Sub StartPSMultiTools()
        If File.Exists(My.Computer.FileSystem.CurrentDirectory + "\PS Multi Tools.exe") Then
            Process.Start(My.Computer.FileSystem.CurrentDirectory + "\PS Multi Tools.exe")
        End If
        Close()
    End Sub

    Public Shared Function IsURLValid(Url As String) As Boolean
        If NetworkInterface.GetIsNetworkAvailable Then
            Try
                Dim request As HttpWebRequest = CType(WebRequest.Create(Url), HttpWebRequest)
                Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        Return True
                    ElseIf response.StatusCode = HttpStatusCode.Found Then
                        Return True
                    ElseIf response.StatusCode = HttpStatusCode.NotFound Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.Unauthorized Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.Forbidden Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.BadGateway Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.BadRequest Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.RequestTimeout Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.GatewayTimeout Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.InternalServerError Then
                        Return False
                    ElseIf response.StatusCode = HttpStatusCode.ServiceUnavailable Then
                        Return False
                    Else
                        Return False
                    End If
                End Using
            Catch Ex As WebException
                Return False
            End Try
        Else
            Return False
        End If
    End Function

End Class
