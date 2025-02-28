Imports System.Security.Principal
Imports System.Drawing
Imports System.Windows.Forms

Public Class Form1
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim username As String = WindowsIdentity.GetCurrent().Name
        Dim welcomename As New Label()
        welcomename.Text = "Welcome, " & username
        welcomename.Location = New Point(60, 60)
        welcomename.Size = New Size(300, 30)
        Me.Controls.Add(welcomename)
    End Sub
End Class