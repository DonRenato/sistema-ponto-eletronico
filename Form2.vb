Imports System.Data
Imports System.Data.OleDb
Public Class Form2
    Public sConn As OleDbConnection
    Dim lDS As DataSet = New DataSet
    Dim lDA As OleDbDataAdapter = New OleDbDataAdapter
    Dim lDR As DataRow

    Private Sub TextboxChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsuario.TextChanged, txtSenha.TextChanged
        If txtUsuario.Text <> "" And txtSenha.Text <> "" Then
            btnLogin.Enabled = True
        Else
            btnLogin.Enabled = False
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        sConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=False")
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        sConn.Open()
        lDA.SelectCommand = New OleDbCommand("SELECT * FROM tblSysLogin WHERE usr_name='" & txtUsuario.Text.ToString & "' and usr_pass='" & txtSenha.Text & "'", sConn)
        lDS.Clear()
        lDA.Fill(lDS)

        If lDS.Tables(0).Rows.Count > 0 Then
            Dim fx As New Form3
            bExitApplication = False
            Me.Close()
            fx.Show()
        Else
            MsgBox("Você não tem permissão para acessar o sistema.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Acesso Negado")
            txtUsuario.Focus()
        End If

        lDR = Nothing
        lDS.Dispose()
        lDA.Dispose()
        sConn.Close()
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Dim f As New Form1
        bExitApplication = False
        Me.Close()
        f.Show()
    End Sub
End Class