
Imports Penpusher.SMS.Encoder.SMS
Imports Penpusher.SMS.Encoder.ConcatenatedShortMessage

Imports System.IO.Ports
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Management

Public Class frmMain
    Private m_SortingColumn As ColumnHeader
    Private eID As Boolean = False
    Private edID As Integer = 0
    Private aID As Boolean = False
    Private taID As Boolean = False
    Private adID As Integer = 0
    Private sGRPA As Integer = 0
    Private tadID As Integer = 0

    Dim SMSObject 'Object To Store SMS or ConcatenatedShortMessage. Late Blinding.
    Dim DataCodingScheme As ENUM_TP_DCS
    Dim ValidPeriod As ENUM_TP_VPF
    Dim PDUCodes() As String


    Private sMODEM As String
    Private sCOMMODEM As String
    Private sMODEMiD As String

    Private Delegate Sub SetTextCallback(text As String)
    ' Private comm_settings As New CommSetting()
    Private Delegate Sub ConnectedHandler(connected As Boolean)

    Private port As String
    Private baudRate As Integer
    Private timeout As Integer

    'Private idGROUP As Integer
    'Private idABONENT As Integer

    'Dim s_antenna As System.Threading.Thread
    'Dim s_conn As System.Threading.Thread

    Shared sp As SerialPort


    Private Sub btnAddA_Click_1(sender As System.Object, e As System.EventArgs) Handles btnAddA.Click
        If Len(txtFIO.Text) = 0 Then
            MsgBox("Введите фамилию абонента", MsgBoxStyle.Information, My.Application.Info.Title)
            txtFIO.Focus()
            Exit Sub
        End If

        If Len(txtPNUMBER.Text) = 0 Then
            MsgBox("Введите номер телефона абонента", MsgBoxStyle.Information, My.Application.Info.Title)
            txtPNUMBER.Focus()
            Exit Sub
        End If

        Dim sSQL As String

        Select Case btnAddA.Text

            Case "Добавить"

                eID = False

                sSQL = "select count(*) as t_n from TBL_ONE where FIO='" & (txtFIO.Text) & "' AND PHONE='" & (txtPNUMBER.Text) & "'"

                Dim rs As Recordset
                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                Dim sCOUNT As Integer

                With rs
                    sCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing

                Select Case sCOUNT

                    Case 0

                        sSQL = "INSERT INTO TBL_ONE(FIO,PHONE,TABEL) VALUES('" & (txtFIO.Text) & "','" & (txtPNUMBER.Text) & "','" & (txtTabel.Text) & "')"

                        DB7.Execute(sSQL)

                    Case Else

                        MsgBox("Абонент с такими данными уже существует," & vbCrLf & "введите другие данные", MsgBoxStyle.Critical, My.Application.Info.Title)
                        txtGroupName.Text = ""

                End Select


            Case Else

                sSQL = "UPDATE TBL_ONE SET FIO='" & (txtFIO.Text) & "',PHONE='" & (txtPNUMBER.Text) & "', TABEL='" & (txtTabel.Text) & "' WHERE id=" & adID

                DB7.Execute(sSQL)

        End Select

        txtFIO.Text = ""
        txtPNUMBER.Text = ""
        txtTabel.Text = ""

        btnAddA.Text = "Добавить"
        adID = 0
        aID = False

        Call Load_abonent()

    End Sub

    Private Sub Load_abonent()

        Dim sCOUNT As Integer
        Dim sSQL As String
        Dim intj As Integer = 0

        sSQL = "SELECT count(*) as t_n FROM TBL_ONE"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        lstAbon.Items.Clear()

        Select Case sCOUNT

            Case 0


            Case Else

                rs = New Recordset
                rs.Open("SELECT * FROM TBL_ONE", DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstAbon.Items.Add(.Fields("id").Value) 'col no. 1
                        lstAbon.Items(CInt(intj)).SubItems.Add((.Fields("FIO").Value))
                        lstAbon.Items(CInt(intj)).SubItems.Add((.Fields("PHONE").Value))
                        lstAbon.Items(CInt(intj)).SubItems.Add((.Fields("TABEL").Value))

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstGroup)


    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        txtFIO.Text = ""
        txtPNUMBER.Text = ""
        txtTabel.Text = ""

        btnAddA.Text = "Добавить"
        adID = 0
        aID = False
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If adID = 0 Or aID = False Then
            Exit Sub
        End If


        Select Case aID

            Case True

                If MsgBox(("Будет произведено удаление абонента" & vbCrLf & "Вы уверены в своих действиях?"), MsgBoxStyle.Exclamation + vbYesNo, My.Application.Info.Title) = vbNo Then Exit Sub

                Dim sSQL As String
                sSQL = "DELETE * FROM TBL_ONE WHERE id=" & adID
                DB7.Execute(sSQL)

                sSQL = "DELETE * FROM TBL_OG WHERE ID_ONE=" & adID
                DB7.Execute(sSQL)

                txtFIO.Text = ""
                txtPNUMBER.Text = ""
                txtTabel.Text = ""

                btnAddA.Text = "Добавить"
                adID = 0
                aID = False

        End Select


        Call Load_abonent()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If Len(cmbAgroup.Text) = 0 Then
            MsgBox("Выберите группу", MsgBoxStyle.Information, My.Application.Info.Title)
            cmbAgroup.Focus()
            Exit Sub
        End If


        Dim sSQL As String
        Dim rs As Recordset
        rs = New Recordset

        'Проверяем есть ли такой абонент в такой группе
        sSQL = "select count(*) as t_n from TBL_OG where ID_ONE=" & adID

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        Dim sCOUNT As Integer
        Dim idGroup As Integer

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        Select Case sCOUNT

            'Если абонент не содержится в таблице то добавляем его
            Case 0

                sSQL = "select count(*) as t_n from TBL_GROUP where GROUPt='" & (cmbAgroup.Text) & "'"

                'Смотрим есть ли такая группа в справочнике
                Dim saCOUNT As Integer

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    saCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing

                Select Case saCOUNT

                    Case 0
                        'Если группы нет то добавляем ее в справочник
                        sSQL = "INSERT INTO TBL_GROUP(GROUPt) VALUES('" & (txtGroupName.Text) & "')"
                        DB7.Execute(sSQL)

                    Case Else

                End Select

                'Находим идентификатор группы для добавления в таблицу
                sSQL = "select id from TBL_GROUP where GROUPt='" & (cmbAgroup.Text) & "'"

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                With rs
                    idGroup = .Fields("id").Value
                End With
                rs.Close()
                rs = Nothing

                sSQL = "INSERT INTO TBL_OG(ID_ONE,ID_GROUP) VALUES(" & adID & "," & idGroup & ")"
                DB7.Execute(sSQL)

            Case Else

                'Если абонент находится в таблице то проверяем есть ли он в данной группе

                sSQL = "select id from TBL_GROUP where GROUPt='" & (cmbAgroup.Text) & "'"

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                With rs
                    idGroup = .Fields("id").Value
                End With
                rs.Close()
                rs = Nothing


                sSQL = "select count(*) as t_n from TBL_OG where ID_ONE=" & adID & " AND ID_GROUP=" & idGroup

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    sCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing


                Select Case sCOUNT

                    Case 0
                        'если абонента в группе нет, то добавляем
                        sSQL = "INSERT INTO TBL_OG(ID_ONE,ID_GROUP) VALUES(" & adID & "," & idGroup & ")"
                        DB7.Execute(sSQL)

                    Case Else
                        'если в группе уже содердится такой абонент то сообщаем
                        MsgBox("В выбранной группе такой абонент уже содержится", MsgBoxStyle.Exclamation, My.Application.Info.Title)


                End Select


        End Select


        Call AG_LOAD(adID)
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        'sGRPA = 0
        If sGRPA = 0 Then
            Exit Sub
        End If

        Select Case aID

            Case True

                If MsgBox(("Будет произведено удаление абонента из группы" & vbCrLf & "Вы уверены в своих действиях?"), MsgBoxStyle.Exclamation + vbYesNo, My.Application.Info.Title) = vbNo Then Exit Sub

                Dim sSQL As String

                sSQL = "DELETE * FROM TBL_OG WHERE ID=" & sGRPA
                DB7.Execute(sSQL)

                sGRPA = 0

        End Select

        AG_LOAD(adID)
    End Sub

    Private Sub AG_LOAD(ByVal id As Integer)

        On Error GoTo err_

        Dim sSQL As String
        Dim rs As Recordset
        Dim sCOUNT As Integer

        sSQL = "SELECT count(*) as t_n FROM TBL_OG"

        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        Select Case sCOUNT

            Case 0

                lstAgr.Items.Clear()

            Case Else

                Dim intj As Integer = 0
                rs = New Recordset

                sSQL = "SELECT TBL_OG.id, TBL_GROUP.GROUPt as GROUPt, TBL_ONE.FIO FROM TBL_ONE INNER JOIN (TBL_GROUP INNER JOIN TBL_OG ON TBL_GROUP.id = TBL_OG.ID_GROUP) ON TBL_ONE.id = TBL_OG.ID_ONE WHERE TBL_ONE.id=" & id

                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                lstAgr.Items.Clear()

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstAgr.Items.Add(.Fields("id").Value)
                        lstAgr.Items(CInt(intj)).SubItems.Add((.Fields("GROUPt").Value))


                        intj = intj + 1
                        .MoveNext()
                    Loop
                End With

        End Select

err_:
    End Sub

    Private Sub btnGroupAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnGroupAdd.Click

        If Len(txtGroupName.Text) = 0 Then
            MsgBox("Введите наименование группы", MsgBoxStyle.Information, My.Application.Info.Title)
            Exit Sub
        End If

        Dim sSQL As String


        Select Case btnGroupAdd.Text

            Case "Добавить"

                eID = False

                sSQL = "select count(*) as t_n from TBL_GROUP where GROUPt='" & (txtGroupName.Text) & "'"

                Dim rs As Recordset
                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                Dim sCOUNT As Integer

                With rs
                    sCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing

                Select Case sCOUNT

                    Case 0

                        sSQL = "INSERT INTO TBL_GROUP(GROUPt) VALUES('" & (txtGroupName.Text) & "')"

                        DB7.Execute(sSQL)

                        txtGroupName.Text = ""

                    Case Else

                        MsgBox("Группа с таким наименвоанием уже существует," & vbCrLf & "введите другое наименование", MsgBoxStyle.Critical, My.Application.Info.Title)
                        txtGroupName.Text = ""

                        btnGroupAdd.Text = "Добавить"


                End Select


            Case Else

                sSQL = "UPDATE TBL_GROUP SET GROUPt='" & (txtGroupName.Text) & "' WHERE id=" & edID

                DB7.Execute(sSQL)

                txtGroupName.Text = ""
                btnGroupAdd.Text = "Добавить"
                edID = 0
                eID = False

        End Select


        Call LOAD_GROUP()
        Call AGroup_Load()
    End Sub

    Private Sub btnGroupClear_Click(sender As System.Object, e As System.EventArgs) Handles btnGroupClear.Click
        txtGroupName.Text = ""
        btnGroupAdd.Text = "Добавить"
    End Sub

    Private Sub btnGroupDel_Click(sender As System.Object, e As System.EventArgs) Handles btnGroupDel.Click

        Select Case eID

            Case True

                Dim sSQL As String
                Dim sCOUNT As Integer

                sSQL = "Select count(*) as t_n FROM TBL_OG where ID_GROUP=" & edID

                Dim rs As Recordset
                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    sCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing

                If MsgBox(("Будет произведено удаление группы" & vbCrLf & "в которой содержится " & sCOUNT & " абонентов" & vbCrLf & "Вы уверены в своих действиях?"), MsgBoxStyle.Exclamation + vbYesNo, My.Application.Info.Title) = vbNo Then Exit Sub

                sSQL = "DELETE * FROM TBL_OG WHERE ID_GROUP=" & edID
                DB7.Execute(sSQL)

                sSQL = "DELETE * FROM TBL_GROUP WHERE id=" & edID
                DB7.Execute(sSQL)

                txtGroupName.Text = ""
                btnGroupAdd.Text = "Добавить"
                edID = 0
                eID = False

        End Select

        Call LOAD_GROUP()
    End Sub

    Private Sub AGroup_Load()

        cmbAgroup.Items.Clear()

        Dim sCOUNT As Integer

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        Select Case sCOUNT

            Case 0

            Case Else

                rs = New Recordset
                rs.Open("SELECT * FROM TBL_GROUP", DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        cmbAgroup.Items.Add((.Fields("GROUPt").Value))
                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstGroup)


    End Sub

    Private Sub LOAD_GROUP()
        On Error GoTo err_

        lstGroup.Items.Clear()
        lstGroup.ListViewItemSorter = Nothing
        lstGroup.Items.Clear()

        Dim intj As Integer = 0
        Dim sCOUNT As Integer

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lstGroup.Items.Clear()

        Select Case sCOUNT

            Case 0

            Case Else

                rs = New Recordset

                sSQL = "SELECT TBL_GROUP.id as gid, TBL_GROUP.GROUPt as GROUPt, (Select count(*) FROM TBL_OG where TBL_GROUP.id=TBL_OG.ID_GROUP) as temp from TBL_GROUP ORDER BY GROUPt"

                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstGroup.Items.Add(.Fields("gid").Value) 'col no. 1
                        lstGroup.Items(CInt(intj)).SubItems.Add((.Fields("GROUPt").Value))

                        lstGroup.Items(CInt(intj)).SubItems.Add(.Fields("temp").Value)
                        'lstGroup.Items(CInt(intj)).SubItems.Add(.Fields("DevelopDev").Value)

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstGroup)

err_:
    End Sub

    Private Sub lstGroup_Click(sender As Object, e As System.EventArgs) Handles lstGroup.Click

        On Error GoTo err_

        Dim sCOUNT As Integer
        If lstGroup.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstGroup.SelectedItems.Count - 1
            sCOUNT = (lstGroup.SelectedItems(z).Text)
        Next

        edID = sCOUNT

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        Select Case sCOUNT

            Case 0

                eID = False

            Case Else

                eID = True
                btnGroupAdd.Text = "Сохранить"

                rs = New Recordset
                rs.Open("SELECT GROUPt FROM TBL_GROUP where id =" & edID, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs

                    txtGroupName.Text = (.Fields("GROUPt").Value)

                End With
                rs.Close()
                rs = Nothing

        End Select

        Exit Sub
err_:
    End Sub

    Private Sub lstGroup_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lstGroup.ColumnClick
        Dim new_sorting_column As ColumnHeader = _
       lstGroup.Columns(e.Column)
        Dim sort_order As System.Windows.Forms.SortOrder
        If m_SortingColumn Is Nothing Then
            sort_order = SortOrder.Ascending
        Else
            If new_sorting_column.Equals(m_SortingColumn) Then
                If m_SortingColumn.Text.StartsWith("> ") Then
                    sort_order = SortOrder.Descending
                Else
                    sort_order = SortOrder.Ascending
                End If
            Else
                sort_order = SortOrder.Ascending
            End If

            m_SortingColumn.Text = m_SortingColumn.Text.Substring(2)
        End If

        m_SortingColumn = new_sorting_column
        If sort_order = SortOrder.Ascending Then
            m_SortingColumn.Text = "> " & m_SortingColumn.Text
        Else
            m_SortingColumn.Text = "< " & m_SortingColumn.Text
        End If

        lstGroup.ListViewItemSorter = New ListViewComparer(e.Column, sort_order)

        lstGroup.Sort()
    End Sub

    Private Sub lstGroup_DoubleClick(sender As Object, e As System.EventArgs) Handles lstGroup.DoubleClick

        Dim sCOUNT As Integer
        If lstGroup.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstGroup.SelectedItems.Count - 1
            sCOUNT = (lstGroup.SelectedItems(z).Text)
        Next

        edID = sCOUNT

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        Select Case sCOUNT

            Case 0

                eID = False

            Case Else

                eID = True
                btnGroupAdd.Text = "Сохранить"

                rs = New Recordset
                rs.Open("SELECT GROUPt FROM TBL_GROUP where id =" & edID, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs

                    txtGroupName.Text = (.Fields("GROUPt").Value)

                End With
                rs.Close()
                rs = Nothing

        End Select


    End Sub

    Private Sub lstAbon_Click(sender As Object, e As System.EventArgs) Handles lstAbon.Click
        Call ABON_CLICK()
    End Sub

    Private Sub lstAbon_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lstAbon.ColumnClick
        Dim new_sorting_column As ColumnHeader = _
  lstAbon.Columns(e.Column)
        Dim sort_order As System.Windows.Forms.SortOrder
        If m_SortingColumn Is Nothing Then
            sort_order = SortOrder.Ascending
        Else
            If new_sorting_column.Equals(m_SortingColumn) Then
                If m_SortingColumn.Text.StartsWith("> ") Then
                    sort_order = SortOrder.Descending
                Else
                    sort_order = SortOrder.Ascending
                End If
            Else
                sort_order = SortOrder.Ascending
            End If

            m_SortingColumn.Text = m_SortingColumn.Text.Substring(2)
        End If

        m_SortingColumn = new_sorting_column
        If sort_order = SortOrder.Ascending Then
            m_SortingColumn.Text = "> " & m_SortingColumn.Text
        Else
            m_SortingColumn.Text = "< " & m_SortingColumn.Text
        End If

        lstAbon.ListViewItemSorter = New ListViewComparer(e.Column, sort_order)

        lstAbon.Sort()
    End Sub

    Private Sub lstAbon_DoubleClick(sender As Object, e As System.EventArgs) Handles lstAbon.DoubleClick

        Call ABON_CLICK()

    End Sub

    Private Sub ABON_CLICK()

        On Error GoTo err_

        Dim sCOUNT As Integer
        If lstAbon.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstAbon.SelectedItems.Count - 1
            sCOUNT = (lstAbon.SelectedItems(z).Text)
        Next

        adID = sCOUNT

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_ONE"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        Select Case sCOUNT

            Case 0

                aID = False

            Case Else

                aID = True
                btnAddA.Text = "Сохранить"

                rs = New Recordset
                rs.Open("SELECT * FROM TBL_ONE where id =" & adID, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs

                    txtFIO.Text = (.Fields("FIO").Value)
                    txtPNUMBER.Text = (.Fields("PHONE").Value)
                    txtTabel.Text = (.Fields("TABEL").Value)

                End With
                rs.Close()
                rs = Nothing

        End Select


        Call AG_LOAD(adID)


        Exit Sub
err_:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, My.Application.Info.Title)
    End Sub

    Private Sub lstAgr_Click(sender As Object, e As System.EventArgs) Handles lstAgr.Click

        sGRPA = 0

        If lstAgr.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstAgr.SelectedItems.Count - 1
            sGRPA = (lstAgr.SelectedItems(z).Text)
        Next


    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click

        lstAbon.Select()
        lstAbon.MultiSelect = False

        Dim item1 As ListViewItem = lstAbon.FindItemWithText(txtSearch.Text, True, 0, True)

        If (item1 IsNot Nothing) Then

            item1.Selected = True
            item1.EnsureVisible()

        Else

        End If

    End Sub

    Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        cmbDataCodingScheme.Items.Add(ENUM_TP_DCS.DefaultAlphabet & ":" & ENUM_TP_DCS.DefaultAlphabet.ToString)
        cmbDataCodingScheme.Items.Add(ENUM_TP_DCS.UCS2 & ":" & ENUM_TP_DCS.UCS2.ToString)
        cmbDataCodingScheme.SelectedIndex = 1

        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.Maximum & ":" & ENUM_TP_VALID_PERIOD.Maximum.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.OneDay & ":" & ENUM_TP_VALID_PERIOD.OneDay.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.OneHour & ":" & ENUM_TP_VALID_PERIOD.OneHour.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.OneWeek & ":" & ENUM_TP_VALID_PERIOD.OneWeek.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.SixHours & ":" & ENUM_TP_VALID_PERIOD.SixHours.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.ThreeHours & ":" & ENUM_TP_VALID_PERIOD.ThreeHours.ToString)
        cmbValidPeriod.Items.Add(ENUM_TP_VALID_PERIOD.TwelveHours & ":" & ENUM_TP_VALID_PERIOD.TwelveHours.ToString)
        cmbValidPeriod.SelectedIndex = 0

        txtMsgRef.Value = 0


        cmbDataCodingSchemeM.Items.Add(ENUM_TP_DCS.DefaultAlphabet & ":" & ENUM_TP_DCS.DefaultAlphabet.ToString)
        cmbDataCodingSchemeM.Items.Add(ENUM_TP_DCS.UCS2 & ":" & ENUM_TP_DCS.UCS2.ToString)
        cmbDataCodingSchemeM.SelectedIndex = 1

        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.Maximum & ":" & ENUM_TP_VALID_PERIOD.Maximum.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.OneDay & ":" & ENUM_TP_VALID_PERIOD.OneDay.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.OneHour & ":" & ENUM_TP_VALID_PERIOD.OneHour.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.OneWeek & ":" & ENUM_TP_VALID_PERIOD.OneWeek.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.SixHours & ":" & ENUM_TP_VALID_PERIOD.SixHours.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.ThreeHours & ":" & ENUM_TP_VALID_PERIOD.ThreeHours.ToString)
        cmbValidPeriodM.Items.Add(ENUM_TP_VALID_PERIOD.TwelveHours & ":" & ENUM_TP_VALID_PERIOD.TwelveHours.ToString)
        cmbValidPeriodM.SelectedIndex = 0

        txtMsgRefM.Value = 0

        'lvModem
        lvModem.Columns.Clear()
        lvModem.Columns.Add("COM порт", 100, HorizontalAlignment.Left)
        lvModem.Columns.Add("Наименование", 240, HorizontalAlignment.Left)

        Call FIND_MODEM()

        '###########################################################################################################
        '
        '###########################################################################################################
        txtFIO.Text = ""
        txtPNUMBER.Text = ""
        txtTabel.Text = ""

        btnAddA.Text = "Добавить"
        adID = 0
        aID = False


        lstGroup.Columns.Clear()
        lstGroup.Columns.Add("id", 1, HorizontalAlignment.Left)
        lstGroup.Columns.Add("Наименование группы", 240, HorizontalAlignment.Left)
        lstGroup.Columns.Add("Колличесто абонентов", 100, HorizontalAlignment.Left)

        'lstAbon

        lstAbon.Columns.Clear()
        lstAbon.Columns.Add("id", 1, HorizontalAlignment.Left)
        lstAbon.Columns.Add("Фамилия, Имя, Отчество", 240, HorizontalAlignment.Left)
        lstAbon.Columns.Add("Номер телефона", 100, HorizontalAlignment.Left)

        lstAgr.Columns.Clear()
        lstAgr.Columns.Add("id", 1, HorizontalAlignment.Left)
        lstAgr.Columns.Add("Наименование группы", 240, HorizontalAlignment.Left)


        lstTemplates.Columns.Clear()
        lstTemplates.Columns.Add("id", 1, HorizontalAlignment.Left)
        lstTemplates.Columns.Add("Наименование шаблона", 240, HorizontalAlignment.Left)
        lstTemplates.Columns.Add("Текст", 240, HorizontalAlignment.Left)


        If DATAB = False Then

            LoadDatabase()
        Else

        End If

        Call LOAD_GROUP()

        Call AGroup_Load()

        Call Load_abonent()

        Call LOAD_TEMPLATES()

        lblSMSkolvo.Text = ""
        lblLenght.Text = ""

    End Sub

    Private Sub FIND_MODEM()
        On Error GoTo err_

        Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_POTSModem")

        sMODEM = ""
        sCOMMODEM = ""
        sMODEMiD = ""

        Dim intj As Integer = 0
        For Each queryObj As ManagementObject In searcher.Get()

            '  sMODEM = queryObj("Model")
            '  sCOMMODEM = queryObj("AttachedTo")
            '  sMODEMiD = queryObj("PNPDeviceID")

            lvModem.Items.Add(queryObj("AttachedTo"))
            lvModem.Items(CInt(intj)).SubItems.Add((queryObj("Model")))

            intj = intj + 1
        Next







        'For i = 0 To 5


        '    lvModem.Items.Add("COM" & i)
        '    lvModem.Items(CInt(intj)).SubItems.Add("_Modem_" & i)

        '    intj = intj + 1
        'Next






        Exit Sub
err_:
        MsgBox(Err.Description)
    End Sub

    Private Sub LOAD_TEMPLATES()
        On Error GoTo err_

        lstTemplates.Items.Clear()
        lstTemplates.ListViewItemSorter = Nothing
        lstTemplates.Items.Clear()

        Dim intj As Integer = 0
        Dim sCOUNT As Integer

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_TEMPLATES"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lstTemplates.Items.Clear()

        Select Case sCOUNT

            Case 0

            Case Else

                rs = New Recordset

                sSQL = "SELECT * FROM TBL_TEMPLATES ORDER BY templateName"

                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstTemplates.Items.Add(.Fields("id").Value) 'col no. 1
                        lstTemplates.Items(CInt(intj)).SubItems.Add((.Fields("templateName").Value))
                        lstTemplates.Items(CInt(intj)).SubItems.Add(.Fields("templateText").Value)

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstTemplates)

err_:
    End Sub

    Private Sub lstTemplates_Click(sender As Object, e As System.EventArgs) Handles lstTemplates.Click

        Call TEMP_LOAD()

    End Sub

    Private Sub lstTemplates_DoubleClick(sender As Object, e As System.EventArgs) Handles lstTemplates.DoubleClick

        Call TEMP_LOAD()

    End Sub

    Private Sub TEMP_LOAD()

        On Error GoTo err_

        Dim sCOUNT As Integer
        If lstTemplates.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstTemplates.SelectedItems.Count - 1
            sCOUNT = (lstTemplates.SelectedItems(z).Text)
        Next

        tadID = sCOUNT

        Dim sSQL As String

        sSQL = "SELECT count(*) as t_n FROM TBL_TEMPLATES"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        Select Case sCOUNT

            Case 0

                taID = False

            Case Else

                taID = True
                Button8.Text = "Сохранить"

                rs = New Recordset
                rs.Open("SELECT * FROM TBL_TEMPLATES where id =" & tadID, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs

                    TextBox1.Text = (.Fields("templateName").Value)
                    txt_message.Text = (.Fields("templateText").Value)
                    cmbDataCodingScheme.Text = .Fields("DataCodingScheme").Value
                    cmbValidPeriod.Text = .Fields("ValidityPeriod").Value
                    txtMsgRef.Value = .Fields("MessageReference").Value
                    chkStatusReport.Checked = .Fields("StatusReport").Value


                End With
                rs.Close()
                rs = Nothing

        End Select

        Exit Sub
err_:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, My.Application.Info.Title)
    End Sub

    Private Sub txt_message_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_message.TextChanged
        'Count for number of PDUs
        Dim i As Integer
        Dim Encoding As Integer '0 for English 1 for Unicode
        Encoding = cmbDataCodingScheme.SelectedIndex
        Dim Text As String = txt_message.Text
        For i = 0 To Text.Length - 1
            If Asc(Text.Chars(i)) < 0 Then
                Encoding = 1
                Exit For
            End If
        Next

        Dim TxtLength As Integer = txt_message.TextLength

        With lblLenght

            .Text = TxtLength
            Dim Piece As Integer

            If Encoding = 0 Then
                If TxtLength <= 160 Then
                    Piece = 1
                    .Text += "/160"
                Else
                    Piece = (TxtLength \ 152) + ((TxtLength Mod 152) = 0) + 1
                    .Text += "/152"
                End If
            End If

            If Encoding = 1 Then
                If TxtLength <= 70 Then
                    Piece = 1
                    .Text += "/70"
                Else
                    Piece = (TxtLength \ 66) + ((TxtLength Mod 66) = 0) + 1
                    .Text += "/66"
                End If
            End If

            lblSMSkolvo.Text = Piece


        End With

    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click

        If Len(TextBox1.Text) = 0 Then
            MsgBox("Введите наименование шаблона", MsgBoxStyle.Information, My.Application.Info.Title)
            Exit Sub
        End If

        Dim sSQL As String

        Select Case Button8.Text

            Case "Добавить"

                taID = False

                sSQL = "select count(*) as t_n from TBL_TEMPLATES where templateName='" & (TextBox1.Text) & "'"

                Dim rs As Recordset
                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                Dim sCOUNT As Integer

                With rs
                    sCOUNT = .Fields("t_n").Value
                End With
                rs.Close()
                rs = Nothing

                'DataCodingScheme - cmbDataCodingScheme.text
                'ValidityPeriod - cmbValidPeriod.text
                'MessageReference - txtMsgRef.value
                'StatusReport - chkStatusReport.checked
                '

                Select Case sCOUNT

                    Case 0

                        sSQL = "INSERT INTO TBL_TEMPLATES(templateName,templateText,DataCodingScheme,ValidityPeriod,MessageReference,StatusReport) VALUES('" & TextBox1.Text & "','" & txt_message.Text & "','" & cmbDataCodingScheme.Text & "','" & cmbValidPeriod.Text & "'," & txtMsgRef.Value & "," & chkStatusReport.Checked & ")"

                        DB7.Execute(sSQL)

                        TextBox1.Text = ""

                    Case Else

                        MsgBox("Шаблон с таким наименвоанием уже существует," & vbCrLf & "введите другое наименование", MsgBoxStyle.Critical, My.Application.Info.Title)
                        txtGroupName.Text = ""

                        Button8.Text = "Добавить"

                End Select


            Case Else
                'DataCodingScheme - cmbDataCodingScheme.text
                'ValidityPeriod - cmbValidPeriod.text
                'MessageReference - txtMsgRef.value
                'StatusReport - chkStatusReport.checked
                'templateName='" & (TextBox1.Text) & "',templateText='" & txt_message.Text & "', DataCodingScheme='" & cmbDataCodingScheme.text & "', ValidityPeriod='" & cmbValidPeriod.text & "', MessageReference=" & txtMsgRef.value & ", StatusReport=" & chkStatusReport.checked & 

                sSQL = "UPDATE TBL_TEMPLATES SET templateName='" & (TextBox1.Text) & "',templateText='" & txt_message.Text & "', DataCodingScheme='" & cmbDataCodingScheme.Text & "', ValidityPeriod='" & cmbValidPeriod.Text & "', MessageReference=" & txtMsgRef.Value & ", StatusReport=" & chkStatusReport.Checked & " WHERE id=" & tadID

                DB7.Execute(sSQL)

                TextBox1.Text = ""
                txt_message.Text = ""
                Button8.Text = "Добавить"

        End Select

        Call LOAD_TEMPLATES()

        '  Call LOAD_TEMPLATES()


        TextBox1.Text = ""
        txt_message.Text = ""
        Button8.Text = "Добавить"
        tadID = 0
        taID = False




    End Sub

    Private Sub lstTemplates_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstTemplates.SelectedIndexChanged

    End Sub

    Private Function GetPDU(ByVal ServiceCenterNumber As String, _
                          ByVal DestNumber As String, _
                          ByVal DataCodingScheme As ENUM_TP_DCS, _
                          ByVal ValidPeriod As ENUM_TP_VALID_PERIOD, _
                          ByVal MsgReference As Integer, _
                          ByVal StatusReport As Boolean, _
                          ByVal UserData As String) As String()
        'Check for SMS type
        Dim Type As Integer '0 for SMS;1 For ConcatenatedShortMessage
        Dim Result() As String
        SMSObject = New SMS.Encoder.SMS
        Select Case DataCodingScheme
            Case ENUM_TP_DCS.DefaultAlphabet
                If UserData.Length > 160 Then
                    SMSObject = New SMS.Encoder.ConcatenatedShortMessage
                    Type = 1
                End If
            Case ENUM_TP_DCS.UCS2
                If UserData.Length > 70 Then
                    SMSObject = New SMS.Encoder.ConcatenatedShortMessage
                    Type = 1
                End If
        End Select

        With SMSObject
            .ServiceCenterNumber = ServiceCenterNumber
            If StatusReport = True Then
                .TP_Status_Report_Request = SMS.Encoder.SMS.ENUM_TP_SRI.Request_SMS_Report
            Else
                .TP_Status_Report_Request = SMS.Encoder.SMS.ENUM_TP_SRI.No_SMS_Report
            End If
            .TP_Destination_Address = DestNumber
            .TP_Data_Coding_Scheme = DataCodingScheme
            .TP_Message_Reference = CInt(txtMsgRef.Text)
            .TP_Validity_Period = ValidPeriod
            .TP_User_Data = UserData
        End With

        If Type = 0 Then
            ReDim Result(0)
            Result(0) = SMSObject.GetSMSPDUCode
        Else
            Result = SMSObject.GetEMSPDUCode            'Note here must use GetEMSPDUCode to get right PDU codes
        End If
        Return Result
    End Function


    Private Sub txt_messageM_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_messageM.TextChanged
        'Count for number of PDUs
        Dim i As Integer
        Dim Encoding As Integer '0 for English 1 for Unicode
        Encoding = cmbDataCodingSchemeM.SelectedIndex
        Dim Text As String = txt_messageM.Text
        For i = 0 To Text.Length - 1
            If Asc(Text.Chars(i)) < 0 Then
                Encoding = 1
                Exit For
            End If
        Next

        Dim TxtLength As Integer = txt_messageM.TextLength

        With lblLenghtM

            .Text = TxtLength
            Dim Piece As Integer

            If Encoding = 0 Then
                If TxtLength <= 160 Then
                    Piece = 1
                    .Text += "/160"
                Else
                    Piece = (TxtLength \ 152) + ((TxtLength Mod 152) = 0) + 1
                    .Text += "/152"
                End If
            End If

            If Encoding = 1 Then
                If TxtLength <= 70 Then
                    Piece = 1
                    .Text += "/70"
                Else
                    Piece = (TxtLength \ 66) + ((TxtLength Mod 66) = 0) + 1
                    .Text += "/66"
                End If
            End If

            lblSMSkolvoM.Text = Piece


        End With
    End Sub

    Private Sub btnSendM_Click(sender As System.Object, e As System.EventArgs) Handles btnSendM.Click

        'Dim result As Boolean
        'result = sendSMS(txt_messageM.Text, txtPhoneM.Text)

        'If result = True Then
        '    OutputM("Сообщение отправлено успешно")
        'Else
        '    OutputM("Произошла ошибка при отправке")
        'End If

        If txtPhoneM.TextLength = 0 Then MsgBox("Введите номер назначения") : Return
        If txt_messageM.TextLength = 0 Then MsgBox("Введите текст в поле") : Return


        'Get PDU Code
        PDUCodes = GetPDU(txtPhoneM.Text, Val(cmbDataCodingSchemeM.Text), Val(cmbValidPeriodM.Text), Val(txtMsgRefM.Text), chkStatusReportM.Checked, txt_messageM.Text)
        'Add PDU Codes to Text

        Try

            Dim i As Integer

            Dim lenMes As Double

            For i = 0 To PDUCodes.Length - 1

                lenMes = PDUCodes(i).Length / 2
                sp.Write("AT+CMGS=" + (Math.Ceiling(lenMes)).ToString() + vbCr & vbLf)
                System.Threading.Thread.Sleep(500)

                sp.Write((PDUCodes(i) & Char.ConvertFromUtf32(26)) + vbCr & vbLf)
                System.Threading.Thread.Sleep(2000)

                OutputM("Сообщение " & i + 1 & " из " & PDUCodes.Length & " отправлено")

            Next

            OutputM("Сообщение отправлено успешно")

        Catch ex As Exception

            OutputM(ex.Message)
            OutputM("Произошла ошибка при отправке")

        End Try

        Try

            Dim recievedData As String
            recievedData = sp.ReadExisting()

            If recievedData.Contains("ERROR") Then

                OutputM("Произошла ошибка при отправке")

            End If

        Catch
        End Try


    End Sub

    Private Function GetPDU( _
                           ByVal DestNumber As String, _
                           ByVal DataCodingScheme As ENUM_TP_DCS, _
                           ByVal ValidPeriod As ENUM_TP_VALID_PERIOD, _
                           ByVal MsgReference As Integer, _
                           ByVal StatusReport As Boolean, _
                           ByVal UserData As String) As String()
        'Check for SMS type
        Dim Type As Integer '0 for SMS;1 For ConcatenatedShortMessage
        Dim Result() As String
        SMSObject = New SMS.Encoder.SMS
        Select Case DataCodingScheme
            Case ENUM_TP_DCS.DefaultAlphabet
                If UserData.Length > 160 Then
                    SMSObject = New SMS.Encoder.ConcatenatedShortMessage
                    Type = 1
                End If
            Case ENUM_TP_DCS.UCS2
                If UserData.Length > 70 Then
                    SMSObject = New SMS.Encoder.ConcatenatedShortMessage
                    Type = 1
                End If
        End Select

        With SMSObject
            ' .ServiceCenterNumber = ServiceCenterNumber
            If StatusReport = True Then
                .TP_Status_Report_Request = SMS.Encoder.SMS.ENUM_TP_SRI.Request_SMS_Report
            Else
                .TP_Status_Report_Request = SMS.Encoder.SMS.ENUM_TP_SRI.No_SMS_Report
            End If
            .TP_Destination_Address = DestNumber
            .TP_Data_Coding_Scheme = DataCodingScheme
            .TP_Message_Reference = CInt(txtMsgRef.Text)
            .TP_Validity_Period = ValidPeriod
            .TP_User_Data = UserData
        End With

        If Type = 0 Then
            ReDim Result(0)
            Result(0) = SMSObject.GetSMSPDUCode
        Else
            Result = SMSObject.GetEMSPDUCode            'Note here must use GetEMSPDUCode to get right PDU codes
        End If
        Return Result
    End Function

    Private Sub lvModem_Click(sender As Object, e As System.EventArgs) Handles lvModem.Click
        On Error GoTo err_

        Dim sCOUNT As Integer
        If lvModem.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lvModem.SelectedItems.Count - 1
            txtModel.Text = (lvModem.SelectedItems(z).SubItems(0).Text)
            sCOMMODEM = lvModem.SelectedItems(z).Text
        Next

        '    lvModem.CheckedItems.Item = True

err_:
    End Sub

    Private Sub lvModem_ItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs) Handles lvModem.ItemCheck
        For Each item In sender.Items
            If Not item.Index = e.Index Then item.Checked = False

            txtModel.Text = lvModem.Items(e.Index).SubItems(1).Text


        Next
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click

        TextBox1.Text = ""
        txt_message.Text = ""
        Button8.Text = "Добавить"
        tadID = 0
        taID = False

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        On Error GoTo err_

        If lstTemplates.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstTemplates.SelectedItems.Count - 1
            tadID = (lstTemplates.SelectedItems(z).Text)
        Next


        Select Case tadID

            Case 0


            Case Else

                If MsgBox(("Будет произведено удаление шаблона" & vbCrLf & "Вы уверены в своих действиях?"), MsgBoxStyle.Exclamation + vbYesNo, My.Application.Info.Title) = vbNo Then Exit Sub

                Dim sSQL As String
                sSQL = "DELETE * FROM TBL_TEMPLATES WHERE id=" & tadID
                DB7.Execute(sSQL)

                TextBox1.Text = ""
                txt_message.Text = ""

                Button8.Text = "Добавить"
                tadID = 0
                taID = False

        End Select

        Call LOAD_TEMPLATES()
        ' Call frmInd.Load_Templates()


        Exit Sub
err_:
    End Sub

    Private connected As Boolean
    Private connectedport As String = "Disconnect"
    Private connectedportname As String = "Disconnect"


    Private Sub btnTestModem_Click(sender As System.Object, e As System.EventArgs) Handles btnTestModem.Click

        Dim t As String = ""
        Dim portname As String = DirectCast(sCOMMODEM, String)

        ' Dim timeout As Integer = 10000
        't: response msg
        sp = New SerialPort

        sp.NewLine = vbCr & vbLf
        sp.BaudRate = 9600
        sp.Parity = Parity.None
        sp.DataBits = 8
        sp.StopBits = StopBits.One
        sp.Handshake = Handshake.None
        sp.DtrEnable = True
        sp.WriteBufferSize = 1024

        If Len(sCOMMODEM) <> 0 AndAlso (connected = False) Then

            Try
                sp.PortName = DirectCast(portname, String)

                ' Dim v As Integer = sp.PortName.IndexOf("COM", 0, sp.PortName.Length)
                ' sp.PortName = sp.PortName.Substring(v, 5)
                Dim port As String = sp.PortName

                'Removing unwanted Leading and Trailing Character's.
                Dim port1 As Char() = New Char(0) {}
                port1(0) = ")"c

                If port.Contains(")") Then
                    port = port.TrimEnd(port1)
                End If

                If port.Length >= 5 Then
                    port = port.Trim()
                End If

                sp.PortName = sCOMMODEM

                If Not sp.IsOpen Then
                    sp.Open()

                    If sp.IsOpen Then
                        ' MessageBox.Show(Convert.ToString("Connected to Port") & portname)

                        OutputM(("Соединение с портом: ") & portname)
                        connected = True
                        connectedport = sp.PortName
                        connectedportname = portname

                        Select Case connected

                            Case True
                                OutputM("ОК")
                                OutputM("-----------------")

                            Case Else

                                OutputM("Соединение отсутствует")
                                OutputM("-----------------")
                        End Select

                        sp.BaseStream.Flush()
                        sp.WriteLine("AT")

                        OutputM("AT")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        'Get the modem's attention

                        sp.WriteLine("ATE1")
                        OutputM("ATE1")
                        OutputM(sp.ReadLine())
                        System.Threading.Thread.Sleep(500)

                        sp.WriteLine("ATZ")
                        OutputM("ATZ")
                        OutputM(sp.ReadLine())
                        System.Threading.Thread.Sleep(500)


                        sp.WriteLine("ATI")
                        OutputM("ATI")
                        OutputM(sp.ReadLine())
                        System.Threading.Thread.Sleep(500)
                        ' Get All Manufacturer Info

                        sp.WriteLine("AT+CGMM")
                        OutputM("AT+CGMM")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        ' Get USB Model
                        sp.WriteLine("AT+CGMI")

                        OutputM("AT+CGMI")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        ' Manufacturer
                        sp.WriteLine("AT+CIMI")

                        OutputM("AT+CIMI")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        ' Get SIM IMSI number
                        sp.WriteLine("AT+CGSN")

                        OutputM("AT+CGSN")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        'Get modem IMEI
                        sp.WriteLine("AT+CGMR")

                        OutputM("AT+CGMR")
                        OutputM(sp.ReadLine())


                        System.Threading.Thread.Sleep(500)
                        sp.WriteLine("AT+CSCS=?")

                        OutputM("AT+CSCS=?")
                        OutputM(sp.ReadLine())

                        System.Threading.Thread.Sleep(500)
                        sp.WriteLine("AT+CSCA")

                        OutputM("AT+CSCA?")
                        OutputM(sp.ReadLine())

                        'AT+CSCA
                        ' Print firmware version of the modem
                        System.Threading.Thread.Sleep(200)


                        timeout = 300


                        While Not ((InlineAssignHelper(t, sp.ReadExisting())).Contains("OK")) AndAlso timeout > 0

                            timeout -= 1
                        End While


                        If Not (t.Equals("")) AndAlso Not (t.Equals(Nothing)) Then

                            Dim z As Char() = New Char(1) {}
                            z(0) = ControlChars.Cr
                            z(1) = ControlChars.Lf
                            Dim f As String() = New String(99) {}
                            Dim c As String() = t.Split(z)
                            Dim m As Integer = 0

                            For i As Integer = 0 To c.Length - 1
                                If Not (c(i).Equals("")) Then
                                    f(m) = c(i)
                                    m += 1
                                End If
                            Next

                            '  Label5.Text = sp.PortName
                            ' Port name 
                            For i As Integer = 0 To m - 1
                                If (f(i).Equals("AT+CGMI")) Then
                                    ' Manufacturer
                                    ' Label6.Text = f(i + 1)
                                End If
                                If (f(i).Equals("AT+CIMI")) Then
                                    ' Get SIM IMSI number
                                    txtIMSI.Text = f(i + 1)
                                End If
                                If (f(i).Equals("AT+CGSN")) Then
                                    'Get modem IMEI
                                    txtIMEI.Text = f(i + 1)
                                End If
                                If (f(i).Equals("AT+CGMM")) Then
                                    ' Get Model of USB 3G
                                    txtModel.Text = f(i + 1)
                                End If

                                'If (f(i).Equals("AT+CGMR")) Then
                                '    ' Print firmware version of the modem
                                '    Label12.Text = f(i + 1)
                                'End If

                                If (f(i).Equals("AT+CSCA?")) Then
                                    ' Print SMSC
                                    txtSMSC.Text = f(i + 1)
                                End If

                            Next

                            sp.BaseStream.Flush()
                        Else

                            sp.Close()
                            connected = False
                            MessageBox.Show(Convert.ToString("Невозможно получить информацию от модема на порту ") & portname)

                            OutputM("Невозможно получить информацию от модема на порту " & portname)

                            MessageBox.Show(Convert.ToString("Подключитесь к другому порту и попробуйте снова. Закрываем порт ") & portname)

                            OutputM("Подключитесь к другому порту и попробуйте снова. Закрываем порт " & portname)

                        End If
                    Else

                        MessageBox.Show(Convert.ToString("Не удается подключиться к порту") & portname)

                    End If
                End If


            Catch ex As Exception
                MessageBox.Show(ex.Message)
                System.Windows.Forms.Application.[Exit]()

            End Try

        Else
            If ((ComboBox1.Items.Count) = 0) Then
                MessageBox.Show("Find the Port")
            End If
        End If

    End Sub

    Public Shared Function StringToUCS2(str As String) As String
        Dim ue As New UnicodeEncoding()
        Dim ucs2 As Byte() = ue.GetBytes(str)

        Dim i As Integer = 0
        While i < ucs2.Length
            Dim b As Byte = ucs2(i + 1)
            ucs2(i + 1) = ucs2(i)
            ucs2(i) = b
            i += 2
        End While
        Return BitConverter.ToString(ucs2).Replace("-", "")
    End Function

    Private Shared Function sendSMS(textsms As String, telnumber As String) As Boolean

        sp = New SerialPort

        If Not sp.IsOpen Then
            Return False
        End If

        Try
            System.Threading.Thread.Sleep(500)
            sp.WriteLine("AT" & vbCr & vbLf)
            ' означает "Внимание!" для модема 
            System.Threading.Thread.Sleep(500)

            sp.Write("AT+CMGF=0" & vbCr & vbLf)
            ' устанавливается цифровой режим PDU для отправки сообщений
            System.Threading.Thread.Sleep(500)
        Catch
            Return False
        End Try

        Try
            telnumber = telnumber.Replace("-", "").Replace(" ", "").Replace("+", "")

            ' 01 это PDU Type или иногда называется SMS-SUBMIT. 01 означает, что сообщение передаваемое, а не получаемое 
            ' цифры 00 это TP-Message-Reference означают, что телефон/модем может установить количество успешных сообщений автоматически
            ' telnumber.Length.ToString("X2") выдаст нам длинну номера в 16-ричном формате
            ' 91 означает, что используется международный формат номера телефона
            telnumber = "01" + "00" + telnumber.Length.ToString("X2") + "91" + EncodePhoneNumber(telnumber)

            textsms = StringToUCS2(textsms)
            ' 00 означает, что формат сообщения неявный. Это идентификатор протокола. Другие варианты телекс, телефакс, голосовое сообщение и т.п.
            ' 08 означает формат UCS2 - 2 байта на символ. Он проще, так что рассмотрим его.
            ' если вместо 08 указать 18, то сообщение не будет сохранено на телефоне. Получится flash сообщение
            Dim leninByte As String = (textsms.Length / 2).ToString("X2")
            textsms = Convert.ToString(Convert.ToString((telnumber & Convert.ToString("00")) + "08") & leninByte) & textsms

            ' посылаем команду с длинной сообщения - количество октет в десятичной системе. то есть делим на два количество символов в сообщении
            ' если октет неполный, то получится в результате дробное число. это дробное число округляем до большего
            Dim lenMes As Double = textsms.Length / 2
            sp.Write("AT+CMGS=" + (Math.Ceiling(lenMes)).ToString() + vbCr & vbLf)
            System.Threading.Thread.Sleep(500)

            ' номер sms-центра мы не указываем, считая, что практически во всех SIM картах он уже прописан
            ' для того, чтобы было понятно, что этот номер мы не указали добавляем к нашему сообщению в начало 2 нуля
            ' добавляем именно ПОСЛЕ того, как подсчитали длинну сообщения
            textsms = Convert.ToString("00") & textsms

            sp.Write((textsms & Char.ConvertFromUtf32(26)) + vbCr & vbLf)
            System.Threading.Thread.Sleep(500)
        Catch
            Return False
        End Try

        Try
            Dim recievedData As String
            recievedData = sp.ReadExisting()

            If recievedData.Contains("ERROR") Then
                Return False

            End If
        Catch
        End Try

        Return True
    End Function

    Private Shared Sub OpenPort()

        sp.BaudRate = 2400
        ' еще варианты 4800, 9600, 28800 или 56000
        sp.DataBits = 7
        ' еще варианты 8, 9
        sp.StopBits = StopBits.One
        ' еще варианты StopBits.Two StopBits.None или StopBits.OnePointFive         
        sp.Parity = Parity.Odd
        ' еще варианты Parity.Even Parity.Mark Parity.None или Parity.Space
        sp.ReadTimeout = 500
        ' еще варианты 1000, 2500 или 5000 (больше уже не стоит)
        sp.WriteTimeout = 500
        ' еще варианты 1000, 2500 или 5000 (больше уже не стоит)
        'port.Handshake = Handshake.RequestToSend;
        'port.DtrEnable = true;
        'port.RtsEnable = true;
        'port.NewLine = Environment.NewLine;

        sp.Encoding = Encoding.GetEncoding("windows-1251")

        sp.PortName = "COM5"

        ' незамысловатая конструкция для открытия порта
        If sp.IsOpen Then
            sp.Close()
        End If
        Try
            sp.Open()
        Catch
        End Try

    End Sub
    ' перекодирование номера телефона для формата PDU
    Public Shared Function EncodePhoneNumber(PhoneNumber As String) As String
        Dim result As String = ""
        If (PhoneNumber.Length Mod 2) > 0 Then
            PhoneNumber += "F"
        End If

        Dim i As Integer = 0
        While i < PhoneNumber.Length
            result += PhoneNumber(i + 1).ToString() + PhoneNumber(i).ToString()
            i += 2
        End While
        Return result.Trim()
    End Function

    '=======================================================
    'Service provided by Telerik (www.telerik.com)
    'Conversion powered by NRefactory.
    'Twitter: @telerik
    'Facebook: facebook.com/telerik
    '=======================================================







    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function


    Private Sub OutputM(ByVal text As String)

        Try

            If Me.txtOutputM.InvokeRequired Then
                Dim stc As New SetTextCallback(AddressOf OutputM)
                Me.Invoke(stc, New Object() {text})
            Else
                txtOutputM.AppendText(text)
                txtOutputM.AppendText(vbCr & vbLf)
                txtOutputM.AppendText(vbCr & vbLf)
            End If

        Catch ex As Exception

        End Try

    End Sub



End Class

