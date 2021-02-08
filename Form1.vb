Imports System.Data
Imports System.Data.OleDb

Public Class Form1

    Public sConn As OleDbConnection
    Dim eDS As DataSet = New DataSet
    Dim eDA As OleDbDataAdapter = New OleDbDataAdapter
    Dim eDR As DataRow
    Dim dDS As DataSet = New DataSet
    Dim dDA As OleDbDataAdapter = New OleDbDataAdapter
    Dim dDR As DataRow

    Private Sub Check_Info()
        sConn.Open()
        eDA.SelectCommand = New OleDbCommand("SELECT emp_fname,emp_lname,emp_mname FROM tblEmployee WHERE emp_idno='" & TextBox1.Text.ToString & "' and emp_pass='" & TextBox2.Text & "'", sConn)
        eDS.Clear()
        eDA.Fill(eDS)
        If eDS.Tables(0).Rows.Count > 0 Then
            eDR = eDS.Tables(0).Rows(0)
            TextBox3.Text = eDR("emp_lname") & ", " & eDR("emp_fname") & " " & eDR("emp_mname")
            dDA.SelectCommand = New OleDbCommand("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "' AND date_timein=#" & Format(Now, "MM/d/yyyy") & "# AND time_timeout IS NULL", sConn)
            dDS.Clear()
            dDA.Fill(dDS)

            If dDS.Tables(0).Rows.Count > 0 Then
                dDR = dDS.Tables(0).Rows(0)

                Button1.Enabled = True
                Button1.Text = "&Saida"
                TextBox4.Text = Format(dDR("time_timein"), "h:mm:ss tt")
            Else
                Button1.Enabled = True
                Button1.Text = "&Entrada"
            End If
            dDR = Nothing
            dDS.Dispose()
            dDA.Dispose()
        Else
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox6.Clear()
            Button1.Enabled = False
        End If
        eDR = Nothing
        eDS.Dispose()
        eDA.Dispose()
        sConn.Close()
    End Sub

    Private Sub ReturnFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus, TextBox4.GotFocus
        TextBox1.Focus()
    End Sub

    Private Sub TextboxChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If (TextBox1.Text <> "" And TextBox2.Text <> "") Then
            Check_Info()
        End If
    End Sub

    Private Sub LoginGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus, TextBox2.GotFocus
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=false")
        Label2.Text = Format(Now, "d MMMM, yyyy  h:mm:ss tt")
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label2.Text = Format(Now, "d MMMM, yyyy  h:mm:ss tt")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Time-in and Time-out button
        Dim strSQL As String
        sConn.Open()
        If Button1.Text = "&Entrada" Then
            dDR = dDS.Tables(0).NewRow()
            TextBox4.Text = Format(Now, "h:mm:ss tt")
            strSQL = "INSERT INTO tblDTR (emp_idno, date_timein, time_timein) VALUES ('" & TextBox1.Text & "', #" & Format(Now, "MM/d/yyyy") & "#, #" & TextBox4.Text & "#)"
            Button1.Text = "&Saida"
        Else
            dDR = dDS.Tables(0).NewRow()
            TextBox6.Text = Format(Now, "h:mm:ss tt")
            strSQL = "UPDATE tblDTR SET time_timeout=#" & TextBox6.Text & "# WHERE emp_idno='" & TextBox1.Text & "' AND date_timein=#" & Format(Now, "MM/d/yyyy") & "# and time_timein=#" & TextBox4.Text & "#"
            Button1.Text = "&Entrada"
        End If
            Dim dCmd As OleDbCommand = New OleDbCommand(strSQL, sConn)
        dCmd.ExecuteNonQuery()
        dCmd.Dispose()

        dDR = Nothing
        dDS.Dispose()
        dDA.Dispose()
        sConn.Close()
        Button1.Enabled = False
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' limpa
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()
        Button1.Text = "&Entrada"
        Button1.Enabled = False
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'System Administration
        bExitApplication = True
        Dim f As Form
        f = New Form2
        Me.Close()
        f.Show()

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not bExitApplication Then
            MsgBox("O sistema não pode ser encerrado dessa forma.." & vbCrLf & vbCrLf & "Encerre a aplicação no módulo de administração do sistema.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Erro: Sem previlégios de Administrador")
            e.Cancel = True
        End If
    End Sub


End Class