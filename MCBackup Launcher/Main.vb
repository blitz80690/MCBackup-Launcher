Imports System.IO
Imports System.Net
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.WindowsAPICodePack.Taskbar

Public Class Main
    Dim WebClient As System.Net.WebClient = New System.Net.WebClient
    Dim LatestVersion As String
    Dim CurrentVersion As String
    Dim AppData As String = Environ("APPDATA")
    Dim FileStream As FileStream
    Dim StreamWriter As StreamWriter

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BackgroundWorker.RunWorkerAsync()
    End Sub

    Public Sub BackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
        My.Computer.FileSystem.CreateDirectory(AppData & "\.mcbackup\logs")

        If Not My.Computer.FileSystem.FileExists(AppData + "\.mcbackup\logs\launcher.log") Then
            File.Create(AppData + "\.mcbackup\logs\launcher.log").Dispose()
        End If

        FileStream = New FileStream((AppData + "\.mcbackup\logs\launcher.log"), FileMode.Append, FileAccess.Write)
        StreamWriter = New StreamWriter(FileStream, System.Text.Encoding.Default)

        StreamWriter.WriteLine("------ New Log Session : " & LogTimeStamp() & "------")
        StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Program Initialized")

        UpdateLabel("Connecting...")
        StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Connecting to the download server...")

        Try
            LatestVersion = WebClient.DownloadString("http://content.nicoco007.com/downloads/mcbackup/version.html")
            If Val(LatestVersion) > 0 Then
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Successfully found latest version: " & LatestVersion)
            Else
                MsgBox("Error: Could not fetch latest version number." + vbNewLine + "This might be because the server is down or you are not connected to the internet." + vbNewLine + "MCBackup will launch as usual.", MsgBoxStyle.Critical, "Error!")
                StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] Invalid Server Response")
                StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] Could not connect to the server.")
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Trying to launch...")
                StartProgram()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error: Could not fetch latest version number." + vbNewLine + "This might be because the server is down or you are not connected to the internet." + vbNewLine + "MCBackup will launch as usual.", MsgBoxStyle.Critical, "Error!")
            StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] " & ex.ToString)
            StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] Could not connect to the server.")
            StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Trying to launch...")
            StartProgram()
            Exit Sub
        End Try

        StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Searching for existing installation...")

        If My.Computer.FileSystem.FileExists(AppData + "\.mcbackup\version") Then
            Using SReader As StreamReader = New StreamReader(AppData + "\.mcbackup\version")
                CurrentVersion = SReader.ReadLine
            End Using
            If Val(LatestVersion) > Val(CurrentVersion) Then
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] New version available.")
                StartDownload()
            Else
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Installation up to date.")
                StartProgram()
            End If
        Else
            StreamWriter.WriteLine(LogTimeStamp() & "[ERROR] MCBackup not found.")
            StartDownload()
        End If
    End Sub

    Private Sub StartDownload()
        AddHandler WebClient.DownloadProgressChanged, AddressOf OnDownloadProgressChanged
        AddHandler WebClient.DownloadFileCompleted, AddressOf OnFileDownloadCompleted
        StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Downloading http://content.nicoco007.com/downloads/mcbackup/" + LatestVersion + "/mcbackup.exe...")
        WebClient.DownloadFileAsync(New Uri("http://content.nicoco007.com/downloads/mcbackup/" + LatestVersion + "/mcbackup.exe"), AppData + "\.mcbackup\mcbackup.exe")
    End Sub

    Private Sub OnDownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs)
        Dim TotalSize As Long = e.TotalBytesToReceive
        Dim DownloadedBytes As Long = e.BytesReceived
        Dim Percentage As Integer = e.ProgressPercentage

        UpdateProgressBarStyle(ProgressBarStyle.Blocks)
        UpdateProgressBarValue(Percentage)
        TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal)
        TaskbarManager.Instance.SetProgressValue(Percentage, 100)
        UpdateLabel("Downloading... " + Percentage.ToString + "%")
    End Sub
    Private Sub OnFileDownloadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        If Not e.Error Is Nothing Then
            MsgBox("Error: An error occured while the download." + vbNewLine + "MCBackup will launch as usual. Please try again later.", MsgBoxStyle.Critical, "Error!")
            StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] " & e.Error.ToString)
            StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] Could not download latest version.")
            StartProgram()
        Else
            File.Create(AppData + "\.mcbackup\version").Dispose()
            Using SWriter As StreamWriter = New StreamWriter(AppData + "\.mcbackup\version")
                SWriter.WriteLine(LatestVersion.ToString)
            End Using
            StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Saved new version " & LatestVersion)
            StartProgram()
        End If
    End Sub

    Private Sub StartProgram()
        If My.Computer.FileSystem.FileExists(AppData & "\.mcbackup\mcbackup.exe") Then
            Try
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Starting " & AppData & "\.mcbackup\mcbackup.exe")
                UpdateLabel("Please wait...")
                Process.Start(AppData + "\.mcbackup\mcbackup.exe")
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Done. Program will exit.")
                FormClose()
            Catch ex As Exception
                StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] " & ex.ToString)
                StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] Could not start MCBackup.")
                UpdateLabel("Error!")
                MsgBox("Error: Could not launch MCBackup. Please try redownloading it.", MsgBoxStyle.Critical, "Error!")

                If My.Computer.FileSystem.FileExists(AppData + "\.mcbackup\mcbackup.exe") Then
                    My.Computer.FileSystem.DeleteFile(AppData + "\.mcbackup\mcbackup.exe")
                End If

                If My.Computer.FileSystem.FileExists(AppData + "\.mcbackup\version") Then
                    My.Computer.FileSystem.DeleteFile(AppData + "\.mcbackup\version")
                End If
                StreamWriter.WriteLine(LogTimeStamp() & "[INFO] Files deleted. Program will exit.")
                FormClose()
            End Try
        Else
            StreamWriter.WriteLine(LogTimeStamp() & "[SEVERE] MCBackup was not found. Program will exit.")
            MsgBox("Error: MCBackup is not installed on your computer!" & vbNewLine & "Please connect to the internet and restart the launcher.", MsgBoxStyle.Critical, "Error!")
            FormClose()
        End If
    End Sub

    Private Sub UpdateLabel(ByVal value As String)
        If Label.InvokeRequired Then
            Label.Invoke(Sub() UpdateLabel(value))
        Else
            Label.Text = value
        End If
    End Sub

    Private Sub UpdateProgressBarValue(ByVal value As Integer)
        If ProgressBar.InvokeRequired Then
            ProgressBar.Invoke(Sub() UpdateProgressBarValue(value))
        Else
            ProgressBar.Value = value
        End If
    End Sub

    Private Sub UpdateProgressBarStyle(ByVal value As ProgressBarStyle)
        If ProgressBar.InvokeRequired Then
            ProgressBar.Invoke(Sub() UpdateProgressBarStyle(value))
        Else
            ProgressBar.Style = value
        End If
    End Sub

    Private Sub FormClose()
        If Me.InvokeRequired Then
            Me.Invoke(Sub() FormClose())
        Else
            Me.Close()
        End If
    End Sub

    Private Function LogTimeStamp()
        Dim Day As String = Format(Now(), "dd")
        Dim Month As String = Format(Now(), "MM")
        Dim Year As String = Format(Now(), "yyyy")
        Dim Hours As String = Format(Now(), "hh")
        Dim Minutes As String = Format(Now(), "mm")
        Dim Seconds As String = Format(Now(), "ss")

        Return Year & "-" & Month & "-" & Day & " " & Hours & ":" & Minutes & ":" & Seconds & " "
    End Function

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        StreamWriter.Close()
    End Sub
End Class
