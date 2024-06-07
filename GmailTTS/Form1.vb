Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Gmail.v1
Imports Google.Apis.Gmail.v1.Data
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports Newtonsoft.Json
Imports System.Speech.Synthesis

Public Class Form1
    Private ReadOnly Scopes As String() = {GmailService.Scope.GmailReadonly, GmailService.Scope.GmailModify}

    Private Const ApplicationName As String = "GmailTTS"
    Private service As GmailService
    Private synthesizer As New SpeechSynthesizer()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        AuthenticateGmail()
        Timer1.Interval = 300000 ' Check for new emails every 5 mins
        Timer1.Start()
    End Sub

    Private Sub AuthenticateGmail()
        Dim credential As UserCredential = Nothing

        Try
            Using stream = New FileStream("credentials.json", FileMode.Open, FileAccess.Read)
                Dim credPath As String = "token.json"
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    New FileDataStore(credPath, True)).Result
            End Using

            service = New GmailService(New BaseClientService.Initializer() With {
                .HttpClientInitializer = credential,
                .ApplicationName = ApplicationName
            })
        Catch ex As Exception
            MessageBox.Show("Failed to authenticate: " & ex.Message)
        End Try
    End Sub

    Private Async Sub CheckForNewEmails()
        Try
            Dim request = service.Users.Messages.List("me")
            request.LabelIds = "INBOX"
            request.Q = "is:unread"
            Dim response = Await request.ExecuteAsync()

            If response.Messages IsNot Nothing AndAlso response.Messages.Count > 0 Then
                For Each messageItem In response.Messages
                    Dim message = Await service.Users.Messages.Get("me", messageItem.Id).ExecuteAsync()
                    Dim snippet = message.Snippet

                    ' Display the snippet of unread message in TextBox1
                    TextBox1.AppendText(snippet & Environment.NewLine)

                    'TTS
                    synthesizer.SpeakAsync(snippet)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to check for new emails: " & ex.Message)
        End Try
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        CheckForNewEmails()
    End Sub
End Class
