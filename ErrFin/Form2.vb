Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Security.Principal


Public Class Form2
    Dim rulepath As String
    Sub listpopulator()
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.Columns.Add("PTS", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Job Code", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Volume", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Hours", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Shift", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Standard", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Department", 100, HorizontalAlignment.Center)

        Try
            Dim rules() As String = File.ReadAllLines(rulepath)
            For Each rule As String In rules

                Dim words() As String = rule.Split(" "c)

                Dim pts As String = words(0)
                Dim jc As String = words(1)
                Dim vol As String = words(2)
                Dim hr As String = words(3)
                Dim sh As String = words(4)
                Dim st As String = words(5)
                Dim dep As String = words(6)

                Dim item As New ListViewItem(pts)
                item.SubItems.Add(jc)
                item.SubItems.Add(vol)
                item.SubItems.Add(hr)
                item.SubItems.Add(sh)
                item.SubItems.Add(st)
                item.SubItems.Add(dep)
                ListView1.Items.Add(item)

            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox7.Text = "Notes : 
                 -> Use this form to make a new rule for ErrFin
                 -> Please make sure all textbox and dropdowns are filled before submitting
                 -> Do not submit an empty rule
                 -> All the rules will be stored inside your user folder in network drive 'GLSUserData'
                 -> If you want to delete a rule, go to your user folder and delete the rule in the rules folder"

        Dim username As String = WindowsIdentity.GetCurrent().Name
        username = username.Substring(3)
        rulepath = "\\ca07102nt800fil.s07102.ca.wal-mart.com\GLS_UserData\" & username & "\My Documents\rules.txt"
        listpopulator()

        Dim path As String = My.Application.Info.DirectoryPath
        path = path & "\loading.gif"
        PictureBox1.Image = Image.FromFile(path)
        PictureBox1.BringToFront()
        PictureBox1.Visible = False
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        PictureBox1.Visible = True

        Dim alphacheck As String = "^[a-zA-Z]+$"
        Dim numcheck As String = "^[0-9]+$"

        Dim tx1, tx2, tx3, tx4, tx5, tx6, tx7, rule As String

        tx1 = TextBox1.Text
        tx2 = TextBox2.Text
        tx3 = TextBox3.Text
        tx4 = ComboBox6.SelectedItem
        tx5 = ComboBox7.SelectedItem
        tx6 = ComboBox8.SelectedItem
        tx7 = TextBox4.Text

        If tx4 Is Nothing Then
            tx4 = ""
        End If

        If tx5 Is Nothing Then
            tx5 = ""
        End If

        If tx6 Is Nothing Then
            tx6 = ""
        End If

        If tx7 Is Nothing Then
            tx7 = ""
        End If


        If TextBox1.Text = String.Empty AndAlso TextBox2.Text = String.Empty AndAlso TextBox3.Text = String.Empty AndAlso tx4 = String.Empty AndAlso tx5 = String.Empty AndAlso tx6 = String.Empty Then
            MsgBox("Please make a rule!", vbOK, "Rule Error")
            PictureBox1.Visible = False
            Exit Sub
        End If

        If TextBox1.Text = String.Empty AndAlso TextBox2.Text = String.Empty AndAlso TextBox3.Text = String.Empty AndAlso tx4 = String.Empty AndAlso tx5 = String.Empty AndAlso tx6 IsNot String.Empty Then
            MsgBox("PLease make a valid rule!", vbOK, "Rule Error")
            PictureBox1.Visible = False
            Exit Sub
        End If

        If Regex.IsMatch(tx1, numcheck) Then

        Else
            tx1 = "None"
        End If

        If Regex.IsMatch(tx2, numcheck) Then

        Else
            tx2 = "None"
        End If

        If Regex.IsMatch(tx3, numcheck) Then

        Else
            tx3 = "None"
        End If

        If Regex.IsMatch(tx7, numcheck) Then

        Else
            tx7 = "None"
        End If

        If Regex.IsMatch(tx4, numcheck) Then

        Else
            tx4 = "None"
        End If


        Directory.CreateDirectory(Path.GetDirectoryName(rulepath))

        rule = tx7 & " " & tx1 & " " & tx2 & " " & tx3 & " " & tx4 & " " & tx5 & " " & tx6

        Try
            Using writer As StreamWriter = New StreamWriter(rulepath, True)
                writer.WriteLine(rule)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbOK, "Rule Error")
            PictureBox1.Visible = False
        End Try

        ListView1.Clear()
        listpopulator()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox6.SelectedIndex = -1
        ComboBox7.SelectedIndex = -1
        ComboBox8.SelectedIndex = -1
        PictureBox1.Visible = False

        MsgBox("Rule Entered!", vbOK, "Rule Success")

    End Sub

End Class