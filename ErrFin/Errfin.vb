Imports System.Security.Principal
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Text
Imports System.IO
Imports System.Net
Imports System.Diagnostics.Eventing.Reader
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class Form1
    Dim rules As String()
    Dim username As String
    Public Sub ErrFin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        username = WindowsIdentity.GetCurrent().Name
        username = username.Substring(3)
        Label1.Text = "Welcome " & username

        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.Columns.Add("Date", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("PTS", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Job Code", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Volume", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Hours", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Shift", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Standard", 100, HorizontalAlignment.Center)
        ListView1.Columns.Add("Department", 100, HorizontalAlignment.Center)
    End Sub

    Private Sub Open_Rules(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form2 As New Form2()
        form2.Show()
    End Sub

    Function GetResponse(ByVal url As String) As String
        Dim request As WebRequest = WebRequest.Create(url)
        request.Method = "GET"

        Using response As WebResponse = request.GetResponse()
            Using dataStream As Stream = response.GetResponseStream()
                Using reader As New StreamReader(dataStream)
                    Dim responseFromServer As String = reader.ReadToEnd()
                    Return responseFromServer
                End Using
            End Using
        End Using
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListView1.Clear()
        Dim rulepath As String = "\\ca07102nt800fil.s07102.ca.wal-mart.com\GLS_UserData\" & username & "\My Documents\rules.txt"
        Try
            rules = File.ReadAllLines(rulepath)
        Catch ex As Exception
            MsgBox("Current rule list cannot be fetched. Try opening this form again!!", vbOK, "Rule Error")
        End Try

        Dim s_date As String = DateTimePicker1.Value
        Dim e_date As String = DateTimePicker2.Value

        If e_date < s_date Then
            Dim temp_d As Date = e_date
            e_date = s_date
            s_date = temp_d
        End If

        Dim one_date As String() = s_date.Split(" "c)
        s_date = "#" & one_date(0) & "#"

        Dim two_date As String() = e_date.Split(" "c)
        e_date = "#" & two_date(0) & "#"

        Dim curr_date As Date = s_date
        Dim date_range As New List(Of String)

        If s_date = e_date Then
            date_range.Add(curr_date)
        Else
            date_range.Add(curr_date)
            Do While curr_date <= e_date
                curr_date = curr_date.AddDays(1)
                date_range.Add(curr_date)
            Loop
            date_range.RemoveAt(date_range.Count - 1)
        End If


        For Each rule As String In rules
            Dim rule_element As String() = rule.Split(" "c)
            Dim pts As String = rule_element(0)
            Dim jc As String = rule_element(1)
            Dim vol As String = rule_element(2)
            Dim hr As String = rule_element(3)
            Dim sh As String = rule_element(4)
            Dim st As String = rule_element(5)
            Dim dep As String = rule_element(6)

            If sh = "all" Then
                sh = "5"
            End If

            If pts = "None" Or pts = "any" Then
                pts = ""
            End If

            If jc = "None" Or jc = "any" Then
                jc = ""
            End If

            If vol = "None" Or vol = "any" Then
                vol = ""
            End If

            If hr = "None" Or hr = "any" Then
                hr = ""
            End If

            Dim err_date, err_name, err_pts, err_jc, err_vol, err_hr, err_sh, err_st, err_dep As String

            For Each da As String In date_range
                Dim url As String = "https://wcllegacy.ca.wal-mart.com/cgi-bin/qa/TPR/perlTPR_Research.pl?DC=7102&txtSDate=" & da & "&txtLKManifest=&ddlLKEmpDept=" & dep & "&ddlLKJobDept=%7BAll%7D&txtLkVol1=" & vol & "&txtLkHrs1=" & hr & "&txtLkOT1=&ddlLKShift=" & sh & "&ddlLkModified=%7BAll%7D&txtLKTop=&txtEDate=" & da & "&txtLKpts=" & pts & "&txtLKpts2=&txtLKJob=" & jc & "&txtLKJob2=&txtLkVol2=&txtLkHrs2=&txtLkOT2="
                Dim response_string As String = GetResponse(url)
                Console.WriteLine(response_string)
                'Dim brk_response() As String = response_string.Split()
                'Dim upd_response() As String = brk_response.Where(Function(s) s.Trim() <> "").ToArray()

                'For Each upd As String In upd_response
                'ListBox1.Items.Add(upd)
                'Next
            Next
        Next
    End Sub
End Class