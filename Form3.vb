Imports System.Data
Imports System.Data.OleDb
Imports System.Math
Public Class Form3

#Region " Variable Declarations "

    Public sConn As OleDbConnection
    Public sConn0 As OleDbConnection
    Public sConn1 As OleDbConnection

    Public iCurrentRecord As Integer
    Public iCountRecord As Integer

    Dim eDS0 As DataSet = New DataSet
    Dim eDA0 As OleDbDataAdapter = New OleDbDataAdapter
    Dim eCB0 As OleDb.OleDbCommandBuilder
    Dim eDR0 As DataRow

    Dim dDS As DataSet = New DataSet
    Dim dDA As OleDbDataAdapter = New OleDbDataAdapter
    Dim dCB As OleDb.OleDbCommandBuilder
    Dim dDR As DataRow

    Dim dDS1 As DataSet = New DataSet
    Dim dDA1 As OleDbDataAdapter = New OleDbDataAdapter
    Dim dCB1 As OleDb.OleDbCommandBuilder
    Dim dDR1 As DataRow

    Dim X As Integer


#End Region

#Region " Main Form Events "

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim f As New Form1
        f.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MsgBox("Essa ação irá fechar o sistema." & vbCrLf & vbCrLf & "Deseja realmente encerrar ?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Encerramento da aplicação") = MsgBoxResult.Yes Then
            bExitApplication = True
            Application.Exit()
        End If
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        sConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=False")
        FillListView("SELECT * FROM tblEmployee")

        ComboBox1.Items.Add("Contrato")
        ComboBox1.Items.Add("Permanente")
        ComboBox1.SelectedIndex = -1
        DateTimePicker1.MaxDate = Now

        MonthCalendar1.TodayDate = Now
        MonthCalendar1.MaxDate = Now
        MonthCalendar2.TodayDate = Now
        MonthCalendar2.MaxDate = Now
        Binding()
    End Sub

#End Region

#Region " TAB 1 User-defined Procedures and Functions "

    Private Function FillListView(ByVal SqlString As String)
        Dim strSQL As String
        ListView1.Items.Clear()
        If sConn.State <> ConnectionState.Open Then
            sConn.Open()
        End If

        eDA0.SelectCommand = New OleDbCommand(SqlString, sConn)
        eCB0 = New OleDb.OleDbCommandBuilder(eDA0)

        eDS0.Clear()
        eDA0.Fill(eDS0)
        iCountRecord = Format(eDS0.Tables(0).Rows.Count, "###0")
        If eDS0.Tables(0).Rows.Count > 0 Then
            X = Me.BindingContext(eDS0.Tables(0)).Position
            For Each eDR0 In eDS0.Tables(0).Rows
                iCurrentRecord = Me.BindingContext(eDS0.Tables(0)).Position + 1
                Dim MyItem = ListView1.Items.Add(eDR0("emp_idno"))
                MyItem.tag = X
                X = X + 1
                With MyItem
                    .SubItems.Add(eDR0("emp_fname".ToString))
                    .SubItems.Add(eDR0("emp_mname".ToString))
                    .SubItems.Add(eDR0("emp_lname".ToString))
                    .SubItems.Add(eDR0("emp_addr".ToString))
                    .SubItems.Add(Format(eDR0("emp_dob"), "dd/MM/yyyy"))
                    .SubItems.Add(eDR0("emp_age".ToString))
                    .SubItems.Add(eDR0("emp_pos".ToString))
                    If eDR0("emp_stat") = "cont" Then
                        .SubItems.Add("Contrato")
                    Else
                        .SubItems.Add("Permanente")
                    End If
                End With
            Next
        End If
        sConn.Close()
    End Function

#End Region

#Region " TAB 2 User-defined Procedures and Functions "

    Private Sub Binding()
        eDA0.SelectCommand = New OleDbCommand("SELECT * FROM tblEmployee", sConn0)
        eCB0 = New OleDb.OleDbCommandBuilder(eDA0)
        eDS0.Clear()
        eDA0.Fill(eDS0)
        iCountRecord = Format(eDS0.Tables(0).Rows.Count, "###0")
        If eDS0.Tables(0).Rows.Count > 0 Then
            RefreshData(True)
        End If
    End Sub

    Private Sub RefreshData(Optional ByVal vEnableNavigation As Boolean = False)
        iCurrentRecord = Me.BindingContext(eDS0.Tables(0)).Position + 1
        Label19.Text = iCurrentRecord & "/" & iCountRecord
        TextBox2.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(0)
        TextBox3.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(2)
        TextBox4.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(4)
        TextBox5.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(3)
        TextBox6.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(5)
        TextBox8.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(7)
        TextBox9.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(8)
        DateTimePicker1.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(6)
        If eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(1) = "numsey" Then Label20.Visible = True Else Label20.Visible = False
        If eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(9) = "cont" Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.SelectedIndex = 1
        End If
        TextBox1.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(0)
        TextBox7.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(3) & ", " &
        eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(2) & " " & eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(4)
        TextBox10.Text = eDS0.Tables(0).Rows(iCurrentRecord - 1).Item(8)

        If RadioButton1.Checked Then
            FillListView2("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "'")
        Else
            FillListView2("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "' " &
                        "AND date_timein>=#" & MonthCalendar1.SelectionStart & "# AND date_timein<=#" & MonthCalendar1.SelectionEnd & "#")
        End If


        If vEnableNavigation Then EnableNavigation()
    End Sub

    Private Sub EnableNavigation()
        If CInt(iCountRecord) > 1 Then
            If CInt(iCurrentRecord) = 1 Then
                Button7.Enabled = False
                Button6.Enabled = False
                Button9.Enabled = True
                Button8.Enabled = True
            ElseIf iCurrentRecord = iCountRecord Then
                Button7.Enabled = True
                Button6.Enabled = True
                Button9.Enabled = False
                Button8.Enabled = False
            Else
                Button7.Enabled = True
                Button6.Enabled = True
                Button9.Enabled = True
                Button8.Enabled = True
            End If
        Else
            Button7.Enabled = False
            Button6.Enabled = False
            Button9.Enabled = False
            Button8.Enabled = False
        End If
    End Sub

    Private Sub TextBoxChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged, TextBox12.TextChanged, TextBox13.TextChanged
        If TextBox11.Text <> "" And TextBox12.Text <> "" And TextBox13.Text <> "" Then
            Button11.Enabled = True
        Else
            Button11.Enabled = False
        End If
    End Sub

#End Region

#Region " TAB 1 Form Events "

    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Me.BindingContext(eDS0.Tables(0)).Position = CInt(ListView1.SelectedItems(0).Tag)
        RefreshData(True)
    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        Select Case e.Column
            Case 0
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_idno")
            Case 1
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_fname")
            Case 2
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_mname")
            Case 3
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_lname")
            Case 4
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_addr")
            Case 5
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_dob")
            Case 6
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_age")
            Case 7
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_pos")
            Case 8
                FillListView("SELECT * FROM tblEmployee ORDER BY emp_stat")
        End Select
    End Sub

#End Region

#Region " TAB 2 Form Events "

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim iDagdag As Integer
        iDagdag = CInt(DateDiff(DateInterval.Year, DateTimePicker1.Value, Now) / 4)
        TextBox8.Text = Floor((DateDiff(DateInterval.Day, DateTimePicker1.Value, Now) - iDagdag) / 365)
    End Sub

#End Region

#Region " TAB 2 Nagegacao "

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' prmeiro
        Me.BindingContext(eDS0.Tables(0)).Position = 0
        RefreshData(True)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ' anterior
        Me.BindingContext(eDS0.Tables(0)).Position -= 1
        RefreshData(True)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ' Proximo
        Me.BindingContext(eDS0.Tables(0)).Position += 1
        RefreshData(True)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ' ultimo
        Me.BindingContext(eDS0.Tables(0)).Position = iCountRecord - 1
        RefreshData(True)
    End Sub

#End Region

#Region " TAB 2 Record Manipulation "

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ' Novo/Salvar/Atualizar
        Dim strSQL As String

        If Button3.Text = "&Novo" Then
            'novo 
            Button3.Text = "&Salvar"
            Button4.Text = "&Cancelar"
            Button10.Enabled = False
            Button5.Enabled = False
            GroupBox3.Enabled = False
            Panel1.Enabled = True

            Button1.Enabled = False
            Button2.Enabled = False

            TextBox2.Text = vbNullString
            TextBox3.Text = vbNullString
            TextBox4.Text = vbNullString
            TextBox5.Text = vbNullString
            TextBox6.Text = vbNullString
            TextBox9.Text = vbNullString
            DateTimePicker1.Text = "1/1/1950"
            ComboBox1.SelectedIndex = -1

            TextBox2.Focus()

            Exit Sub
        Else
            sConn.Open()
            Dim sStat As String

            If (TextBox2.Text <> vbNullString And TextBox3.Text <> vbNullString And
                TextBox4.Text <> vbNullString And TextBox5.Text <> vbNullString And
                TextBox6.Text <> vbNullString And TextBox8.Text <> vbNullString And
                TextBox9.Text <> vbNullString And ComboBox1.SelectedIndex >= 0) Then

                If ComboBox1.SelectedIndex = 0 Then sStat = "cont" Else sStat = "perm"
                If Button3.Text = "&Salvar" Then
                    'salvar
                    dDA.SelectCommand = New OleDbCommand("SELECT * FROM tblEmployee WHERE emp_idno='" & TextBox2.Text & "'", sConn)
                    dDS.Clear()
                    dDA.Fill(dDS)
                    If dDS.Tables(0).Rows.Count = 0 Then
                        strSQL = "INSERT INTO tblEmployee VALUES ('" & TextBox2.Text & "','numsey','" & TextBox3.Text & "','" & TextBox5.Text & "','" &
                        TextBox4.Text & "','" & TextBox6.Text & "',#" & DateTimePicker1.Value & "#,'" & TextBox8.Text & "','" & TextBox9.Text & "','" & sStat & "')"
                        Dim dCmdx As OleDbCommand = New OleDbCommand(strSQL, sConn)
                        dCmdx.ExecuteNonQuery()
                        dCmdx.Dispose()

                        MsgBox("Novo Funcionário incluído com sucesso. Altere a senha imediatamente..." & vbCrLf & vbCrLf & "Sua Senha:  numsey ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Sucesso")
                        Binding()
                        FillListView("SELECT * FROM tblEmployee")
                    Else
                        MsgBox("O Código do funcionário ja existe ? Verifique a informação." & vbCrLf & vbCrLf & "Se você achar que isso for um erro, " &
                        "Informa ao administrador.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Erro: Código duplicado encontrado")
                        TextBox2.Focus()
                        dDS.Dispose()
                        dDA.Dispose()
                        sConn.Close()
                        Exit Sub
                    End If
                    dDS.Dispose()
                    dDA.Dispose()
                Else
                    'atualizar
                    strSQL = "UPDATE tblEmployee SET emp_fname='" & TextBox3.Text & "', emp_mname='" & TextBox4.Text & "',emp_lname='" & TextBox5.Text & "',emp_addr='" & TextBox6.Text & "'," &
                    "emp_dob=#" & DateTimePicker1.Value & "#, emp_age='" & TextBox8.Text & "',emp_pos='" & TextBox9.Text & "',emp_stat='" & sStat & "' " &
                    "WHERE emp_idno='" & TextBox2.Text & "'"

                    Dim dCmdx As OleDbCommand = New OleDbCommand(strSQL, sConn)
                    dCmdx.ExecuteNonQuery()
                    dCmdx.Dispose()

                    MsgBox("Informação do funcionário atualizada.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Sucesso")
                    Binding()
                    FillListView("SELECT * FROM tblEmployee")
                End If
                Button3.Text = "&Novo"
                Button4.Text = "&Editar"
                Button10.Enabled = True
                Button5.Enabled = True
                GroupBox3.Enabled = True
                Panel1.Enabled = False

                Button1.Enabled = True
                Button2.Enabled = True

                sConn.Close()
            Else
                MsgBox("Alguns campos estão em branco. Preencha toda a informação.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Erro: Campos em branco")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' EDITAR/CANCELAR
        If Button4.Text = "&Editar" Then
            Button3.Text = "&Atualizar"
            Button4.Text = "&Cancelar"
            Button10.Enabled = False
            Button5.Enabled = False
            GroupBox3.Enabled = False
            Panel1.Enabled = True
            TextBox2.ReadOnly = True
            Button1.Enabled = False
            Button2.Enabled = False
        Else
            Button3.Text = "&Novo"
            Button4.Text = "&Editar"
            Button10.Enabled = True
            Button5.Enabled = True
            GroupBox3.Enabled = True
            Panel1.Enabled = False
            TextBox2.ReadOnly = False
            Button1.Enabled = True
            Button2.Enabled = True
            RefreshData(True)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'ALTERA SENHA
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        Panel2.Enabled = True
        TextBox11.Focus()
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        ' DELETAR REGISTRO
        Dim strSQL As String
        If MsgBox("Deseja deletar este registro ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirma exclusão do registro") = MsgBoxResult.Yes Then
            sConn.Open()
            strSQL = "DELETE FROM tblEmployee WHERE emp_idno='" & TextBox2.Text & "'"
            Dim dCmdx As OleDbCommand = New OleDbCommand(strSQL, sConn)
            dCmdx.ExecuteNonQuery()
            dCmdx.Dispose()

            MsgBox("Informação deletada com sucesso.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Sucesso")
            Binding()
            FillListView("SELECT * FROM tblEmployee")
            sConn.Close()
        End If
    End Sub

#End Region

#Region " TAB 2 Change Password "

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        'Cancelar
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        Panel2.Enabled = False
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        'Alterar
        If TextBox11.Text = eDS0.Tables(0).Rows(BindingContext(eDS0.Tables(0)).Position).Item(1) Then
            If TextBox12.Text = TextBox13.Text Then
                sConn.Open()
                Dim dCmd As OleDbCommand = New OleDbCommand("UPDATE tblEmployee SET emp_pass='" & TextBox12.Text & "' WHERE emp_idno = '" & TextBox2.Text & "'", sConn)

                dCmd.ExecuteNonQuery()
                dCmd.Dispose()

                MsgBox("Senha Alterada com sucesso.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Senha Alterada")
                GroupBox2.Enabled = True
                GroupBox3.Enabled = True
                Panel2.Enabled = False
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox13.Text = ""

                Button1.Enabled = True
                Button2.Enabled = True

                Binding()
                sConn.Close()
            Else
                MsgBox("A nova senha é igual a anterior.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Erro: Senhas iguais")
                TextBox12.Focus()
            End If
        Else
            MsgBox("Senha informada inválida!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Erro: Senha Inválida")
            TextBox11.Focus()
        End If
    End Sub

#End Region

#Region " TAB 4 Form Events "

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        If RadioButton2.Checked Then
            FillListView2("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "' " &
                        "AND date_timein>=#" & MonthCalendar1.SelectionStart & "# AND date_timein<=#" & MonthCalendar1.SelectionEnd & "#")
        End If
    End Sub

    Private Function FillListView2(ByVal SqlString As String)
        Dim iX As Integer
        sConn0 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=False")
        ListView2.Items.Clear()
        sConn0.Open()

        dDA.SelectCommand = New OleDbCommand(SqlString, sConn0)
        dCB = New OleDb.OleDbCommandBuilder(dDA)
        dDS.Clear()
        dDA.Fill(dDS)
        If dDS.Tables(0).Rows.Count > 0 Then
            iX = Me.BindingContext(dDS.Tables(0)).Position
            For Each dDR In dDS.Tables(0).Rows
                Dim MyItem = ListView2.Items.Add(Format(dDR("date_timein"), "dd/MM/yyyy"))
                MyItem.Tag = iX
                iX = iX + 1
                With MyItem
                    .SubItems.Add(Format(dDR("time_timein"), "h:mm:ss tt"))
                    If Not IsDBNull(dDR("time_timeout")) Then
                        .SubItems.Add(Format(dDR("time_timeout"), "h:mm:ss tt"))
                    Else
                        .SubItems.Add("SEM SAÍDA!")
                    End If
                End With
            Next
        End If

        sConn0.Close()
    End Function

    Private Sub RadioButtonChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        If RadioButton1.Checked Then
            FillListView2("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "'")
        Else
            FillListView2("SELECT * FROM tblDTR WHERE emp_idno='" & TextBox1.Text & "' " &
                        "AND date_timein>=#" & MonthCalendar1.SelectionStart & "# AND date_timein<=#" & MonthCalendar1.SelectionEnd & "#")
        End If
    End Sub

#End Region

#Region " TAB 3 Form Events and UD Proc and Func "

    Private Function FillListView3(ByVal SqlString As String)
        Dim iX As Integer
        sConn1 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=False")
        sConn1.Open()
        ListView3.Items.Clear()
        dDA1.SelectCommand = New OleDbCommand(SqlString, sConn1)
        dCB1 = New OleDb.OleDbCommandBuilder(dDA1)

        dDS1.Clear()
        dDA1.Fill(dDS1)

        If dDS1.Tables(0).Rows.Count > 0 Then
            For Each dDR1 In dDS1.Tables(0).Rows
                Dim MyItem = ListView3.Items.Add(dDR1("emp_idno"))
                MyItem.tag = dDR1("emp_idno")
                With MyItem
                    .subitems.Add(Format(dDR1("date_timein"), "dd/MM/yyyy"))
                    .SubItems.Add(Format(dDR1("time_timein"), "h:mm:ss tt"))
                    If Not IsDBNull(dDR1("time_timeout")) Then
                        .SubItems.Add(Format(dDR1("time_timeout"), "h:mm:ss tt"))
                    Else
                        .SubItems.Add("SEM SAÍDA!")
                    End If
                End With
            Next
        End If
        sConn1.Close()
    End Function

    Private Sub ListView3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView3.Click
        TextBox16.Text = ListView3.SelectedItems(0).Tag
        sConn1 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Funci.mdb;Persist Security Info=False")
        sConn1.Open()
        Dim eDSx As DataSet = New DataSet
        Dim eDAx As OleDbDataAdapter = New OleDbDataAdapter
        Dim eDRx As DataRow

        eDAx.SelectCommand = New OleDbCommand("SELECT emp_fname,emp_lname,emp_mname,emp_pos FROM tblEmployee WHERE emp_idno='" & ListView3.SelectedItems(0).Tag & "'", sConn1)
        eDSx.Clear()
        eDAx.Fill(eDSx)
        If eDSx.Tables(0).Rows.Count > 0 Then
            eDRx = eDSx.Tables(0).Rows(0)
            TextBox15.Text = eDRx("emp_lname") & ", " & eDRx("emp_fname") & " " & eDRx("emp_mname")
            TextBox14.Text = eDRx("emp_pos")
        End If
        eDRx = Nothing
        eDSx.Dispose()
        eDAx.Dispose()
        sConn1.Close()
    End Sub

    Private Sub MonthCalendar2_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar2.DateChanged
        If RadioButton4.Checked Then
            FillListView3("SELECT * FROM tblDTR WHERE date_timein>=#" & MonthCalendar2.SelectionStart & "# AND date_timein<=#" & MonthCalendar2.SelectionEnd & "#")
        End If
    End Sub

    Private Sub RadioButtonXChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged, RadioButton4.CheckedChanged
        If RadioButton3.Checked Then
            FillListView3("SELECT * FROM tblDTR")
        Else
            FillListView3("SELECT * FROM tblDTR WHERE date_timein>=#" & MonthCalendar2.SelectionStart & "# AND date_timein<=#" & MonthCalendar2.SelectionEnd & "#")
        End If
    End Sub


#End Region



End Class