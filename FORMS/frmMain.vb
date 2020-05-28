
Imports Penpusher.SMS.Encoder.SMS
Imports Penpusher.SMS.Encoder.ConcatenatedShortMessage

Imports System.IO.Ports
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Management

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms

Public Class frmMain

    Private m_SortingColumn As ColumnHeader
    Private eID As Boolean = False
    Private edID As Integer = 0
    Private aID As Boolean = False
    Private taID As Boolean = False
    Private adID As Integer = 0
    Private sGRPA As Integer = 0
    Private tadID As Integer = 0
    Private Delegate Sub SetTextCallback(ByVal text As String)
    Private Delegate Sub SetTextCallbackT(ByVal text As String, ByVal sost As String)

    Shared sp As SerialPort

    Private idGROUP As Integer
    Private idABONENT As Integer

    Dim ThMASSSEND As System.Threading.Thread
    Dim sTEXT As String
    Dim sDEstin As String

    Dim sKOL As Integer

    'Dim s_antenna As System.Threading.Thread
    'Dim s_conn As System.Threading.Thread

    Private Sub btnAddA_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddA.Click
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

    Private Sub STAT()

        Dim sAB As Integer
        Dim sOG As Integer
        Dim sSQL As String
        Dim intj As Integer = 0

        sSQL = "SELECT count(*) as t_n FROM TBL_ONE"

        Dim rs As Recordset
        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sAB = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"

        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sOG = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing


        lblContact.Text = "Количество абонентов: [" & sAB & "/" & sOG & "]"


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

        cmbAgroup.Enabled = False
        lstAgr.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False


        ResList(lstGroup)


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        txtFIO.Text = ""
        txtPNUMBER.Text = ""
        txtTabel.Text = ""

        btnAddA.Text = "Добавить"
        adID = 0
        aID = False

        lstAgr.Items.Clear()

        cmbAgroup.Enabled = False
        lstAgr.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

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

    Private Sub btnGroupAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupAdd.Click

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

    Private Sub btnGroupClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupClear.Click
        txtGroupName.Text = ""
        btnGroupAdd.Text = "Добавить"
    End Sub

    Private Sub btnGroupDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupDel.Click

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

    Private Sub lstGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstGroup.Click

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

    Private Sub lstGroup_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstGroup.ColumnClick
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

    Private Sub lstGroup_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstGroup.DoubleClick

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

    Private Sub lstAbon_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAbon.Click
        Call ABON_CLICK()
    End Sub

    Private Sub lstAbon_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstAbon.ColumnClick
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

    Private Sub lstAbon_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAbon.DoubleClick

        Call ABON_CLICK()

    End Sub

    Private Sub ABON_CLICK()

        On Error GoTo err_

        cmbAgroup.Enabled = True
        lstAgr.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True

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

    Private Sub lstAgr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAgr.Click

        sGRPA = 0

        If lstAgr.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lstAgr.SelectedItems.Count - 1
            sGRPA = (lstAgr.SelectedItems(z).Text)
        Next


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        lstAbon.Select()
        lstAbon.MultiSelect = False

        Dim item1 As ListViewItem = lstAbon.FindItemWithText(txtSearch.Text, True, 0, True)

        If (item1 IsNot Nothing) Then

            item1.Selected = True
            item1.EnsureVisible()

        Else

        End If

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        txt_messageM.Text = "Текстовое короткое сообщение SMS"

        'lvModem
        lvModem.Columns.Clear()
        lvModem.Columns.Add("COM порт", 100, HorizontalAlignment.Left)
        lvModem.Columns.Add("Наименование", 240, HorizontalAlignment.Left)

        Call FIND_MODEM()


        ComboBox1.Items.Add("2400")
        ComboBox1.Items.Add("4800")
        ComboBox1.Items.Add("9600")
        ComboBox1.Items.Add("19200")
        ComboBox1.Items.Add("38400")
        ComboBox1.Items.Add("57600")
        ComboBox1.Items.Add("115200")

        baudRate = ComboBox1.Text
        ComboBox1.Text = baudRate.ToString()

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

        lblSearch.Visible = False
        lblSearch.Text = ""
        sFindText.Text = ""

        lstSMS.Columns.Clear()
        lstSMS.Columns.Add("id", 1, HorizontalAlignment.Left)
        lstSMS.Columns.Add("Дата", 240, HorizontalAlignment.Left)
        lstSMS.Columns.Add("Время", 100, HorizontalAlignment.Left)
        lstSMS.Columns.Add("Сообщение", 100, HorizontalAlignment.Left)
        lstSMS.Columns.Add("Абонент", 100, HorizontalAlignment.Left)
        lstSMS.Columns.Add("Номер телефона", 100, HorizontalAlignment.Left)
        lstSMS.Items.Clear()

        rbm.Checked = True

        If DATAB = False Then

            LoadDatabase()
        Else

        End If


        Call OTCHET("mes")
        Call KOLVO()

        Call LOAD_GROUP()

        Call AGroup_Load()

        Call Load_abonent()

        Call LOAD_TEMPLATES()

        Call Load_Templates_()

        Call LOAD_NUM()

        Call STAT()


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
            If queryObj("Status").ToString() = "OK" Then
                lvModem.Items.Add(queryObj("AttachedTo"))
                lvModem.Items(CInt(intj)).SubItems.Add((queryObj("Model")))

                intj = intj + 1
            End If

        Next

        'For i = 3 To 15

        '    lvModem.Items.Add("COM " & i)
        '    lvModem.Items(CInt((i - 3) + intj)).SubItems.Add("Модем " & i)

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

    Private Sub lstTemplates_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTemplates.Click

        Call TEMP_LOAD()

    End Sub

    Private Sub lstTemplates_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTemplates.DoubleClick

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
                    'cmbDataCodingScheme.Text = .Fields("DataCodingScheme").Value
                    'cmbValidPeriod.Text = .Fields("ValidityPeriod").Value
                    'txtMsgRef.Value = .Fields("MessageReference").Value
                    'chkStatusReport.Checked = .Fields("StatusReport").Value


                End With
                rs.Close()
                rs = Nothing

        End Select

        Exit Sub
err_:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, My.Application.Info.Title)
    End Sub

    Private Sub txt_message_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_message.TextChanged
        'Count for number of PDUs
        Dim i As Integer
        Dim Encoding As Integer '0 for English 1 for Unicode
        Encoding = 1
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

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

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

                        sSQL = "INSERT INTO TBL_TEMPLATES(templateName,templateText) VALUES('" & TextBox1.Text & "','" & txt_message.Text & "')"

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

                sSQL = "UPDATE TBL_TEMPLATES SET templateName='" & (TextBox1.Text) & "',templateText='" & txt_message.Text & "' WHERE id=" & tadID

                DB7.Execute(sSQL)

                TextBox1.Text = ""
                txt_message.Text = ""
                Button8.Text = "Добавить"

        End Select

        Call LOAD_TEMPLATES()
        Call Load_Templates_()


        TextBox1.Text = ""
        txt_message.Text = ""
        Button8.Text = "Добавить"
        tadID = 0
        taID = False

    End Sub

    Private Sub txt_messageM_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_messageM.TextChanged
        'Count for number of PDUs
        Dim i As Integer
        Dim Encoding As Integer '0 for English 1 for Unicode
        Encoding = 1

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

    Private Sub btnSendM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendM.Click

        On Error GoTo err_


        If Me.modem.Port.IsOpen Then

            Call SEND_SMS(txt_messageM.Text, txtPhoneM.Text, False)

        Else

            MsgBox("Нет соединения с модемом")


        End If

        Exit Sub
err_:

        MsgBox("Нет соединения с модемом")
    End Sub

    Private Sub SEND_SMS(ByVal text As String, ByVal phone As String, ByVal MESTO As Boolean)

        If phone.Length = 0 Then MsgBox("Введите номер назначения") : Return
        If text.Length = 0 Then MsgBox("Введите текст в поле") : Return

        Dim sMSCreator As New SMSCreator()
        sMSCreator.CreateSMS(phone, text, "")

        Me.Invoke(Sub() GroupBox3.Visible = True)

        Try

            Dim num As Integer = 1
            For Each current As String In sMSCreator.Messages

                Application.DoEvents()
                'Me.Invoke(Sub() Label8.Text = "Отправление сообщения : " & num & " из " & sMSCreator.Messages.Count)

                If Not Me.modem.SendCommand("AT+CMGF=0", vbCr & vbLf & "OK" & vbCr & vbLf) Then
                    Throw New Exception("")
                End If

                If Not Me.modem.SendCommand("AT+CMGS=" + (current.Length / 2).ToString(), ">") Then
                    Throw New Exception("")
                End If

                Select Case MESTO
                    Case False
                        OutputM(Me.modem.LogString)
                    Case True
                        Output(Me.modem.LogString, "")
                End Select

                If Not Me.modem.SendCommand(sMSCreator.SCA + current & Convert.ToString(ChrW(26)), vbCr & vbLf & "OK" & vbCr & vbLf) Then
                    Throw New Exception("")
                End If

                'Select Case MESTO
                '    Case False
                '        OutputM(Me.modem.LogString)
                '    Case True
                '        Output(Me.modem.LogString, "")
                'End Select


                Me.Invoke(Sub() Label19.Text = "SMS: " & num & "/" & sMSCreator.Messages.Count * sKOL)
                num += 1

                '
                Application.DoEvents()

                Dim sSQL As String
                Dim sTmp As DateTime = DateTime.Now

                Select Case RadioButton2.Checked

                    Case True

                        sSQL = "INSERT INTO TBL_ARHIVE (data,times,txtSMS,id_ONE,id_GROUP) VALUES ('" & Date.Today & "','" & sTmp.ToLongTimeString & "','" & (text) & "'," & idABONENT & ",0)"
                        DB7.Execute(sSQL)

                    Case False

                        sSQL = "INSERT INTO TBL_ARHIVE (data,times,txtSMS,id_ONE,id_GROUP) VALUES ('" & Date.Today & "','" & sTmp.ToLongTimeString & "','" & (text) & "'," & idABONENT & "," & idGROUP & ")"
                        DB7.Execute(sSQL)

                End Select

            Next

            Select Case MESTO
                Case False
                    OutputM("Сообщение: " & num - 1 & " из " & sMSCreator.Messages.Count & " для абонента " & phone & " отправлено")
                    OutputM("Пауза 2 сек.")
                Case True
                    Output("Сообщение: " & num - 1 & " из " & sMSCreator.Messages.Count & " для абонента " & phone, "Отправлено")
                    Output("2 секунды ", "пауза")
            End Select

            Thread.Sleep(2000)

        Catch ex As Exception

            Select Case MESTO
                Case False
                    OutputM(ex.Message)
                    OutputM("Произошла ошибка при отправке")
                    OutputM("-----------------------------")

                Case True
                    Output(ex.Message, "")
                    Output("Произошла ошибка при отправке сообщения для абонента " & phone, "Не выполнено")
                    Output("-----------------------------", "")
            End Select

        End Try

    End Sub

    Private Sub lvModem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvModem.Click
        On Error GoTo err_

        Dim sCOUNT As Integer
        If lvModem.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lvModem.SelectedItems.Count - 1

            If lvModem.SelectedItems(z).SubItems(1).Text <> "" Then

                txtModel.Text = (lvModem.SelectedItems(z).SubItems(1).Text)
                sMODEM = lvModem.SelectedItems(z).SubItems(1).Text
            Else


            End If

            sCOMMODEM = lvModem.SelectedItems(z).Text

        Next

        '    lvModem.CheckedItems.Item = True

err_:
    End Sub

    Private Sub lvModem_DoubleClick(sender As Object, e As System.EventArgs) Handles lvModem.DoubleClick

        On Error GoTo err_

        Dim sCOUNT As Integer
        If lvModem.Items.Count = 0 Then Exit Sub

        Dim z As Integer

        For z = 0 To lvModem.SelectedItems.Count - 1

            If lvModem.SelectedItems(z).SubItems(1).Text <> "" Then

                txtModel.Text = (lvModem.SelectedItems(z).SubItems(1).Text)
                sMODEM = lvModem.SelectedItems(z).SubItems(1).Text
            Else


            End If

            sCOMMODEM = lvModem.SelectedItems(z).Text

        Next

        '    lvModem.CheckedItems.Item = True

err_:

    End Sub

    Private Sub lvModem_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvModem.ItemCheck

        For Each item In sender.Items
            If Not item.Index = e.Index Then item.Checked = False

            lvModem.Items(e.Index).Selected = True


            If lvModem.Items(e.Index).SubItems(1).Text <> "" Then

                txtModel.Text = (lvModem.Items(e.Index).SubItems(1).Text)
                sMODEM = lvModem.Items(e.Index).SubItems(1).Text
            Else


            End If

            sCOMMODEM = lvModem.Items(e.Index).Text

        Next

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        TextBox1.Text = ""
        txt_message.Text = ""
        Button8.Text = "Добавить"
        tadID = 0
        taID = False

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
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
    Public modem As ModemClass

    Private Sub btnTestModem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestModem.Click

        lvModem.Select()

        For intj = 0 To lvModem.Items.Count - 1

            lvModem.Items(intj).Selected = True
            lvModem.Items(intj).EnsureVisible()

            If lvModem.Items(intj).Checked = True Then


                Dim z1 As Integer

                For z1 = 0 To lvModem.SelectedItems.Count - 1
                    txtModel.Text = (lvModem.SelectedItems(z1).SubItems(1).Text)
                    sCOMMODEM = lvModem.SelectedItems(z1).Text
                    sMODEM = lvModem.SelectedItems(z1).SubItems(1).Text
                Next



            End If

        Next


        Dim t As String = ""
        Dim portname As String = DirectCast(sCOMMODEM, String)
        baudRate = ComboBox1.Text


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

        Dim sproizv As String

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

                        System.Threading.Thread.Sleep(50)
                        'Get the modem's attention

                        sp.WriteLine("ATE1")
                        System.Threading.Thread.Sleep(50)

                        sp.WriteLine("ATZ")
                        System.Threading.Thread.Sleep(50)


                        sp.WriteLine("ATI")
                        System.Threading.Thread.Sleep(50)
                        ' Get All Manufacturer Info

                        sp.WriteLine("AT+CGMM")

                        System.Threading.Thread.Sleep(50)
                        ' Get USB Model
                        sp.WriteLine("AT+CGMI")

                        System.Threading.Thread.Sleep(50)
                        ' Manufacturer
                        sp.WriteLine("AT+CIMI")

                        System.Threading.Thread.Sleep(50)
                        ' Get SIM IMSI number
                        sp.WriteLine("AT+CGSN")

                        System.Threading.Thread.Sleep(50)
                        'Get modem IMEI
                        sp.WriteLine("AT+CGMR")

                        System.Threading.Thread.Sleep(50)
                        sp.WriteLine("AT+CSCS=?")

                        System.Threading.Thread.Sleep(50)
                        sp.WriteLine("AT+CSCA")

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

                                    sproizv = f(i + 1)

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
                                    sMODEM = f(i + 1)
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
                '  System.Windows.Forms.Application.[Exit]()

            End Try

        Else
            If ((ComboBox1.Items.Count) = 0) Then
                MessageBox.Show("Find the Port")
            End If
        End If

        sp.Close()

        Dim configClass As New ConfigClass()
        Me.modem = New ModemClass(configClass)

        '   Me.modem = sMODEM

        Me.modem.Connect()
        OutputM(Me.modem.modemReply)

        If Not Me.modem.[error] Then

            Try
                If Not Me.modem.SendCommand("ATE0", vbCr & vbLf & "OK" & vbCr & vbLf) Then
                    Throw New Exception("")
                End If
                If Not Me.modem.SendCommand("ATI7", vbCr & vbLf & "OK" & vbCr & vbLf) Then
                    Throw New Exception("")
                End If
                OutputM("Тест прошел успешно, модем ответил:" & vbLf & vbLf + Me.modem.modemReply)

                lblConnect.Text = "Соединение установлено: " & sCOMMODEM & ", модем: " & sproizv & " " & sMODEM

            Catch ex As Exception
                If Me.modem.[error] AndAlso ex.Message.Length = 0 Then
                    MessageBox.Show("Произошла ошибка  при приеме ответа на команду " + Me.modem.lastCommand + ":" & vbLf + Me.modem.modemReply)
                End If
            End Try
            '    Me.modem.Disconnect()
            Return
        End If

        MessageBox.Show("Ошибка открытия порта модема.")

    End Sub

    ' перекодирование номера телефона для формата PDU

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

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Select Case RadioButton2.Checked

            Case True
                Label7.Text = "Абонент"

            Case Else
                Label7.Text = "Группа абонентов"

                'txt_destination_numbers
        End Select

        Call LOAD_NUM()

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Select Case RadioButton1.Checked

            Case True
                Label7.Text = "Группа абонентов"


            Case Else

                Label7.Text = "Абонент"

                'txt_destination_numbers
        End Select

        Call LOAD_NUM()

    End Sub

    Private Sub LOAD_NUM()
        On Error GoTo err_

        txt_destination_numbers.Text = ""
        txt_destination_numbers.Items.Clear()

        If DATAB = False Then
            Exit Sub
            '   LoadDatabase()

        End If

        Dim sCOUNT As Integer
        Dim sSQL As String
        Dim rs As Recordset

        Select Case RadioButton2.Checked

            Case True

                sSQL = "SELECT count(*) as t_n FROM TBL_ONE"
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
                        rs.Open("SELECT * FROM TBL_ONE", DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                        With rs
                            .MoveFirst()
                            Do While Not .EOF

                                txt_destination_numbers.Items.Add((.Fields("FIO").Value))

                                .MoveNext()
                            Loop
                        End With
                        rs.Close()
                        rs = Nothing
                End Select

            Case Else

                sSQL = "SELECT count(*) as t_n FROM TBL_GROUP"
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

                                txt_destination_numbers.Items.Add((.Fields("GROUPt").Value))

                                .MoveNext()
                            Loop
                        End With
                        rs.Close()
                        rs = Nothing

                End Select

        End Select









        'label36
err_:
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        On Error GoTo err_

        txtSMS.Text = ""

        Dim sSQL As String
        Dim rs As Recordset

        sSQL = "select count(*) as t_n from TBL_TEMPLATES WHERE templateName='" & ComboBox3.Text & "'"

        'Смотрим есть ли такой шаблон в справочнике
        Dim sCOUNT As Integer

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
                rs.Open("SELECT * FROM TBL_TEMPLATES WHERE templateName='" & ComboBox3.Text & "'", DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs

                    txtSMS.Text = .Fields("templateText").Value


                End With
                rs.Close()
                rs = Nothing

        End Select

        sCOUNT = 0

err_:
    End Sub

    Public Sub Load_Templates_()
        Dim sSQL As String
        Dim rs As Recordset

        ComboBox3.Items.Clear()

        sSQL = "select count(*) as t_n from TBL_TEMPLATES"

        'Смотрим есть ли такой шаблон в справочнике
        Dim sCOUNT As Integer

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
                rs.Open("SELECT * FROM TBL_TEMPLATES", DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                Dim intj As Integer = 0
                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        ComboBox3.Items.Add(.Fields("templateName").Value)



                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        sCOUNT = 0

    End Sub

    Private Sub BtnClear_Click(sender As System.Object, e As System.EventArgs) Handles BtnClear.Click

        txt_message.Text = ""
        txt_destination_numbers.Text = ""
        txt_message.Focus()

    End Sub

    Private Sub btnSendMessage_Click(sender As System.Object, e As System.EventArgs) Handles btnSendMessage.Click


        Try



            If Len(txtSMS.Text) = 0 Then

                MsgBox("Нечего посылать", MsgBoxStyle.Exclamation, Application.ProductName.ToString)
                Exit Sub

            End If


            If Len(txt_destination_numbers.Text) = 0 Then

                MsgBox("Некому посылать", MsgBoxStyle.Exclamation, Application.ProductName.ToString)
                Exit Sub

            End If

            If Me.modem.Port.IsOpen Then


                btnSendMessageStop.Enabled = True
                btnSendMessage.Enabled = False

                sTEXT = txtSMS.Text
                sDEstin = txt_destination_numbers.Text

                pbSendSMS.Value = 0

                ThMASSSEND = New System.Threading.Thread(AddressOf SEND_SMS_MASS)
                ThMASSSEND.Start()

            Else


                MsgBox("Нет соединения с модемом", MsgBoxStyle.Exclamation, Application.ProductName.ToString)
                Exit Sub

            End If


        Catch ex As Exception
            MsgBox("Нет соединения с модемом", MsgBoxStyle.Exclamation, Application.ProductName.ToString)
        End Try


    End Sub

    Sub SEND_SMS_MASS()

        Try
            '     If CommSetting.comm.IsConnected() = True Then

            Dim sSQL As String
            Dim rs As Recordset

            Select Case RadioButton1.Checked

                Case True

                    'Отправка СМС группе Абонентов

                    sSQL = "select count(*) as t_n from TBL_GROUP where GROUPt='" & (sDEstin) & "'"

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

                            MsgBox("Такой группы нет в справочнике" & vbCrLf & "выберите другую группу", vbExclamation, My.Application.Info.Title)

                        Case Else

                            sSQL = "select id from TBL_GROUP where GROUPt='" & (sDEstin) & "'"

                            Dim sCOUNT As Integer
                            rs = New Recordset
                            rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                            With rs
                                sCOUNT = .Fields("id").Value
                            End With
                            rs.Close()
                            rs = Nothing

                            Dim s1COUNT As Integer

                            sSQL = "select count(*) as t_n from TBL_OG where ID_GROUP=" & sCOUNT

                            rs = New Recordset
                            rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                            With rs
                                s1COUNT = .Fields("t_n").Value
                            End With
                            rs.Close()
                            rs = Nothing

                            Me.Invoke(Sub() pbSendSMS.Maximum = 100)

                            Dim sProc As Integer = 100 / s1COUNT
                            ' sProc = Math.Ceiling(sProc)

                            Select Case s1COUNT

                                Case 0

                                    MsgBox("В выбранной группе не содержится абонентов" & vbCrLf & "выберите другую группу", vbExclamation, My.Application.Info.Title)
                                    Exit Sub

                                Case Else


                                    sKOL = s1COUNT

                                    rs = New Recordset
                                    rs.Open("SELECT TBL_OG.ID_GROUP as groupid, TBL_ONE.PHONE as aPhone, TBL_ONE.ID as oneid, TBL_ONE.FIO as aFIO FROM TBL_ONE INNER JOIN (TBL_GROUP INNER JOIN TBL_OG ON TBL_GROUP.id = TBL_OG.ID_GROUP) ON TBL_ONE.id = TBL_OG.ID_ONE WHERE TBL_OG.ID_GROUP=" & sCOUNT, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                                    Dim intj As Integer = 0
                                    With rs
                                        .MoveFirst()
                                        Do While Not .EOF

                                            idGROUP = .Fields("groupid").Value
                                            idABONENT = .Fields("oneid").Value

                                            Call SEND_SMS(sTEXT, .Fields("aPhone").Value, True)

                                            'Output("Сообщение: 1 из " & Label25.Text & " для абонента " & .Fields("aPhone").Value, "Отправлено")
                                            'Output("2 секунды ", "пауза")

                                            Thread.Sleep(500)

                                            intj = intj + 1
                                            Me.Invoke(Sub() pbSendSMS.Value += sProc)

                                            .MoveNext()
                                        Loop
                                    End With
                                    rs.Close()
                                    rs = Nothing

                                    Me.Invoke(Sub() pbSendSMS.Value = 100)
                                    Output("Отправка " & Label25.Text & " сообщений для " & s1COUNT & " адресатов завершена", "")


                                    '  MsgBox("Осуществлена отправка SMS " & intj & " абонентам", vbInformation, My.Application.Info.Title)

                            End Select

                    End Select

                Case Else

                    'Отправка СМС Абоненту

                    sSQL = "select count(*) as t_n from TBL_ONE where FIO='" & (sDEstin) & "'"

                    'Смотрим есть ли такой абонент в справочнике
                    Dim saCOUNT As Integer

                    rs = New Recordset
                    rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                    With rs
                        saCOUNT = .Fields("t_n").Value
                    End With
                    rs.Close()
                    rs = Nothing

                    Dim sProc As Integer = 100 / saCOUNT
                    ' sProc = Math.Ceiling(sProc)

                    Select Case saCOUNT

                        Case 0

                            MsgBox("Такого абонента не содержится в справочнике" & vbCrLf & "выберите другого абонента", vbExclamation, My.Application.Info.Title)
                            Exit Sub

                        Case Else

                            sKOL = saCOUNT

                            sSQL = "select * from TBL_ONE where FIO='" & (sDEstin) & "'"

                            rs = New Recordset
                            rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                            Dim intj As Integer = 0
                            With rs
                                .MoveFirst()
                                Do While Not .EOF

                                    'idGROUP=.Fields("PHONE").Value
                                    idABONENT = .Fields("id").Value

                                    Call SEND_SMS(sTEXT, .Fields("PHONE").Value, True)

                                    'Output("Сообщение: 1 из " & Label25.Text & " для абонента " & .Fields("PHONE").Value, "Отправлено")
                                    'Output("2 секунды ", "пауза")

                                    Me.Invoke(Sub() pbSendSMS.Value += sProc)
                                    Thread.Sleep(500)
                                    intj = intj + 1

                                    .MoveNext()
                                Loop
                            End With
                            rs.Close()
                            rs = Nothing

                            Me.Invoke(Sub() pbSendSMS.Value = 100)

                            Output("Отправка " & Label25.Text & " сообщения для " & saCOUNT & " адресатов завершена", "")

                            Call SAVE_TEMPLATES(sTEXT)

                    End Select

            End Select

            Me.Invoke(Sub() btnSendMessageStop.Enabled = False)
            Me.Invoke(Sub() btnSendMessage.Enabled = True)
            ThMASSSEND.Abort()

        Catch ex As Exception

            '  Output(ex.Message, "")

        End Try

        Me.Invoke(Sub() btnSendMessageStop.Enabled = False)
        Me.Invoke(Sub() btnSendMessage.Enabled = True)
        ThMASSSEND.Abort()

    End Sub

    Private Sub SAVE_TEMPLATES(ByVal text As String)
        On Error GoTo err_


        Dim sSQL As String
        Dim rs As Recordset

        sSQL = "select count(*) as t_n from TBL_TEMPLATES where templateText='" & text & "'"

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

                'Если шаблона с таким текстом нет то предлогаем сохранить

                If MsgBox(("Хотите сохранить текст SMS как шаблон?"), MsgBoxStyle.Exclamation + vbYesNo, My.Application.Info.Title) = vbYes Then

                    Dim Message, Title, Defaults As String
                    Dim strTmp As String

                    Message = "Укажите наименование шаблона"
                    Title = "Шаблон SMS"
                    Defaults = "Шаблон от " & DateTime.Now
                    strTmp = InputBox(Message, Title, Defaults)

                    Select Case strTmp

                        Case ""
                            MsgBox("Не введено наименование шаблона SMS", MsgBoxStyle.Exclamation, My.Application.Info.Title)

                        Case Else

                            'Сохраняем как шаблон

                            sSQL = "INSERT INTO TBL_TEMPLATES (templateName,templateText) VALUES ('" & strTmp & "','" & text & "')"
                            DB7.Execute(sSQL)

                    End Select

                End If

            Case Else
                'если шаблон с таким текстом есть то ничего не делаем

        End Select

err_:
    End Sub

    Private Sub Output(ByVal text As String, ByVal sost As String)

        Try

            Dim oldcolor As Color
            Me.Invoke(Sub() txtOutput.SelectionStart = txtOutput.Text.Length)
            Me.Invoke(Sub() oldcolor = txtOutput.SelectionColor)

            Dim sTmp As DateTime = DateTime.Now
            Me.Invoke(Sub() txtOutput.AppendText(sTmp.ToShortDateString & "  " & sTmp.ToLongTimeString() & "  -  " & text))


            '  Me.Invoke(Sub() txtOutput.SelectionStart = txtOutput.Text.Length)
            '   Me.Invoke(Sub() txtOutput.ScrollToCaret())

                Select Case sost

                    Case "Отправлено"
                        Me.Invoke(Sub() txtOutput.SelectionColor = Color.Green)
                        Me.Invoke(Sub() txtOutput.AppendText(" - " & "Отправлено"))

                    Case "Не выполнено"

                        Me.Invoke(Sub() txtOutput.SelectionColor = Color.Red)
                        Me.Invoke(Sub() txtOutput.AppendText(" - " & "Не выполнено"))

                    Case "пауза"

                        Me.Invoke(Sub() txtOutput.SelectionColor = Color.Blue)
                        Me.Invoke(Sub() txtOutput.AppendText(" - пауза"))


                End Select

            Me.Invoke(Sub() txtOutput.SelectionColor = oldcolor)
            Me.Invoke(Sub() txtOutput.AppendText(vbCr & vbLf))
            Me.Invoke(Sub() txtOutput.ScrollToCaret())


            'Dim sb As New StringBuilder
            'sb.Append(Environment.NewLine + text)

            'txtOutput.Text = sb.ToString
            'txtOutput.SelectionStart = txtOutput.Text.Length
            'txtOutput.ScrollToCaret()


        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnSendMessageStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendMessageStop.Click

        ThMASSSEND.Abort()

    End Sub

    Private Sub txtSMS_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSMS.TextChanged
        'Count for number of PDUs

        Dim i As Integer
        Dim Encoding As Integer '0 for English 1 for Unicode
        Encoding = 1
        Dim Text As String = txtSMS.Text
        For i = 0 To Text.Length - 1
            If Asc(Text.Chars(i)) < 0 Then
                Encoding = 1
                Exit For
            End If
        Next

        Dim TxtLength As Integer = txtSMS.TextLength

        With txt_text_remaining

            .Text = TxtLength
            Dim Piece As Integer

            If Encoding = 1 Then
                If TxtLength <= 70 Then
                    Piece = 1
                    .Text += "/70"
                Else
                    Piece = (TxtLength \ 66) + ((TxtLength Mod 66) = 0) + 1
                    .Text += "/66"
                End If
            End If

            Label25.Text = Piece


        End With
    End Sub

    Private Sub LOAD_ARHIVE()

        If DATAB = False Then LoadDatabase()

        lstSMS.Items.Clear()

        Dim sCOUNT As Integer
        Dim sSQL As String
        Dim intj As Integer = 0

        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE"

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

                btnSearch.Enabled = False

            Case Else

                btnSearch.Enabled = True

                sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstSMS.Items.Add(.Fields("id").Value) 'col no. 1
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("data").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("times").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("txtSMS").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("FIO").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("PHONE").Value))

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing


        End Select

        ResList(lstSMS)

    End Sub

    Private Sub lstSMS_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lstSMS.ColumnClick
        Dim new_sorting_column As ColumnHeader = _
      lstSMS.Columns(e.Column)
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

        lstSMS.ListViewItemSorter = New ListViewComparer(e.Column, sort_order)

        lstSMS.Sort()
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click

        sFindText.Text = ""
        lblSearch.Text = ""

        If rbn.Checked = True Then
            Call OTCHET("ned")
        End If

        If rbm.Checked = True Then
            Call OTCHET("mes")
        End If

        If rbpg.Checked = True Then
            Call OTCHET("polg")
        End If

        If rbg.Checked = True Then
            Call OTCHET("god")
        End If

    End Sub

    Private Sub OTCHET(ByVal period As String)

        lstSMS.Items.Clear()

        Dim sSQL As String
        Dim sCOUNT As Integer
        Dim intj As Integer = 0

        Dim D1, D2 As String
        Dim dQ() As String

        Dim tmpDat_ As DateTime
        Dim tmpDat As DateTime = Date.Today

        Select Case period

            Case "ned"

                tmpDat_ = Date.Today.AddDays(-7)

            Case "mes"
                tmpDat_ = Date.Today.AddDays(-30)

            Case "polg"
                tmpDat_ = Date.Today.AddDays(-182)

            Case "god"

                tmpDat_ = Date.Today.AddDays(-365)

        End Select

        dQ = Split(tmpDat, ".")

        D1 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"

        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"

        Dim rs As Recordset
        rs = New Recordset

        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE WHERE (data BETWEEN " & D2 & " AND " & D1 & ")"

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        '##############################################################################

        Select Case sCOUNT

            Case 0

                btnSearch.Enabled = False

                MsgBox("Отсутствуют записи за период с " & tmpDat_ & " по " & tmpDat, MsgBoxStyle.Information, My.Application.Info.Title)

            Case Else

                btnSearch.Enabled = True

                sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE where (TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 & ") ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"

                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstSMS.Items.Add(.Fields("id").Value) 'col no. 1
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("data").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("times").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("txtSMS").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("FIO").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("PHONE").Value))

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstSMS)


    End Sub

    Private Sub KOLVO()

        Dim sSQL As String
        Dim sCOUNT As Integer
        Dim rs As Recordset

        Dim D1, D2 As String
        Dim dQ() As String

        Dim tmpDat_ As DateTime
        Dim tmpDat As DateTime = Date.Today

        dQ = Split(tmpDat, ".")

        D1 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"

        tmpDat_ = Date.Today.AddDays(-7)
        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"


        rs = New Recordset
        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE WHERE (data BETWEEN " & D2 & " AND " & D1 & ")"

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lbln.Text = sCOUNT & " SMS."


        'Месяц
        tmpDat_ = Date.Today.AddDays(-30)
        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"


        rs = New Recordset
        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE WHERE (data BETWEEN " & D2 & " AND " & D1 & ")"

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lblm.Text = sCOUNT & " SMS."

        'полгода
        tmpDat_ = Date.Today.AddDays(-182)
        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"


        rs = New Recordset
        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE WHERE (data BETWEEN " & D2 & " AND " & D1 & ")"

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lblp.Text = sCOUNT & " SMS."

        'год
        tmpDat_ = Date.Today.AddDays(-365)
        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"


        rs = New Recordset
        sSQL = "SELECT count(*) as t_n FROM TBL_ARHIVE WHERE (data BETWEEN " & D2 & " AND " & D1 & ")"

        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        lblg.Text = sCOUNT & " SMS."




    End Sub

    Private Sub btnToExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnToExcel.Click

        Dim tmpDat_ As DateTime
        Dim tmpDat As DateTime = Date.Today

        If rbn.Checked = True Then
            tmpDat_ = Date.Today.AddDays(-7)
        End If

        If rbm.Checked = True Then
            tmpDat_ = Date.Today.AddDays(-30)
        End If

        If rbpg.Checked = True Then
            tmpDat_ = Date.Today.AddDays(-182)
        End If

        If rbg.Checked = True Then
            tmpDat_ = Date.Today.AddDays(-365)
        End If

        Call ExportListViewToExcel(lstSMS, "Отчет об отправленных SMS за период с " & tmpDat_ & " по " & tmpDat)

    End Sub

    Private Sub btnSearch_Click(sender As System.Object, e As System.EventArgs) Handles btnSearch.Click
        On Error Resume Next



        ' sFindText.Text = ""

        If rbn.Checked = True Then
            Call OTCHET_search("ned")
        End If

        If rbm.Checked = True Then
            Call OTCHET_search("mes")
        End If

        If rbpg.Checked = True Then
            Call OTCHET_search("polg")
        End If

        If rbg.Checked = True Then
            Call OTCHET_search("god")
        End If


        Exit Sub


        Dim sSQL As String
        Dim rs As Recordset
        Dim intj As Integer = 0
        Dim sCOUNT As Integer

        lstSMS.Items.Clear()

        Dim D1 As String
        Dim dQ As String()
        dQ = Split(sFindText.Text, ".")

        D1 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"
        If D1 Is Nothing Then
            dQ = Split(DateAndTime.Now, ".")

            D1 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"
        End If

        '####################


        sSQL = "SELECT count(*) as t_n FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
" WHERE TBL_ARHIVE.txtSMS like '%" & (sFindText.Text) &
"%' or TBL_ONE.FIO like '%" & (sFindText.Text) &
"%' or TBL_ONE.PHONE like '%" & (sFindText.Text) &
"%' or TBL_ARHIVE.data=" & D1 & ""

        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        If sCOUNT = 0 Then Exit Sub

        lblSearch.Visible = True
        lblSearch.Text = "Найдено: " & sCOUNT & " записей"


        sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
" WHERE TBL_ARHIVE.txtSMS like '%" & (sFindText.Text) &
"%' or TBL_ONE.FIO like '%" & (sFindText.Text) &
"%' or TBL_ONE.PHONE like '%" & (sFindText.Text) &
"%' or TBL_ARHIVE.data=" & D1 &
" ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"

        rs = New Recordset
        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            .MoveFirst()
            Do While Not .EOF

                lstSMS.Items.Add(.Fields("id").Value) 'col no. 1
                lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("data").Value))
                lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("times").Value))
                lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("txtSMS").Value))
                lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("FIO").Value))
                lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("PHONE").Value))

                intj = intj + 1

                .MoveNext()
            Loop
        End With
        rs.Close()
        rs = Nothing


        ResList(lstSMS)



    End Sub

    Private Sub OTCHET_search(ByVal period As String)
        On Error GoTo err_


        lstSMS.Items.Clear()

        Dim sSQL As String
        Dim sCOUNT As Integer
        Dim intj As Integer = 0

        Dim D1, D2 As String
        Dim dQ() As String

        Dim tmpDat_ As DateTime
        Dim tmpDat As DateTime = Date.Today

        Select Case period

            Case "ned"

                tmpDat_ = Date.Today.AddDays(-7)

            Case "mes"
                tmpDat_ = Date.Today.AddDays(-30)

            Case "polg"
                tmpDat_ = Date.Today.AddDays(-182)

            Case "god"

                tmpDat_ = Date.Today.AddDays(-365)

        End Select

        dQ = Split(tmpDat, ".")

        D1 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"

        dQ = Split(tmpDat_, ".")
        D2 = "#" & dQ(1) & "/" & dQ(0) & "/" & dQ(2) & "#"

        Dim rs As Recordset
        rs = New Recordset

        '        sSQL = "SELECT count(*) as t_n FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
        '"WHERE (TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 & ")" &
        '        " AND TBL_ARHIVE.txtSMS like '%" & (sFindText.Text) &
        '"%' or TBL_ONE.FIO like '%" & (sFindText.Text) &
        '"%' or TBL_ONE.PHONE like '%" & (sFindText.Text) & "%'"

        sSQL = "SELECT count(*) as t_n FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
            "WHERE TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 &
            " and TBL_ARHIVE.txtSMS like '%" & sFindText.Text &
            "%'"


        rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

        With rs
            sCOUNT = .Fields("t_n").Value
        End With
        rs.Close()
        rs = Nothing

        '##############################################################################

        lblSearch.Visible = True
        lblSearch.Text = "Найдено: " & sCOUNT & " записей"

        Select Case sCOUNT

            Case 0

                'btnSearch.Enabled = False

                MsgBox("Отсутствуют записи за период с " & tmpDat_ & " по " & tmpDat, MsgBoxStyle.Information, My.Application.Info.Title)

            Case Else

                btnSearch.Enabled = True

                'sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE where (TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 & ") ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"
                '                sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
                '"WHERE (TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 & ")" &
                '        " AND TBL_ARHIVE.txtSMS like '%" & (sFindText.Text) &
                '"%' or TBL_ONE.FIO like '%" & (sFindText.Text) &
                '"%' or TBL_ONE.PHONE like '%" & (sFindText.Text) & "%'" &
                '" ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"

                sSQL = "SELECT TBL_ARHIVE.id, TBL_ARHIVE.data, TBL_ARHIVE.times, TBL_ARHIVE.txtSMS, TBL_ONE.FIO, TBL_ONE.PHONE FROM TBL_ONE INNER JOIN TBL_ARHIVE ON TBL_ONE.id = TBL_ARHIVE.id_ONE " &
                    "WHERE TBL_ARHIVE.data BETWEEN " & D2 & " AND " & D1 &
                    " and TBL_ARHIVE.txtSMS like '%" & sFindText.Text &
                    "%' ORDER BY TBL_ARHIVE.data desc,TBL_ARHIVE.times desc"



                rs = New Recordset
                rs.Open(sSQL, DB7, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                With rs
                    .MoveFirst()
                    Do While Not .EOF

                        lstSMS.Items.Add(.Fields("id").Value) 'col no. 1
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("data").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("times").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("txtSMS").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("FIO").Value))
                        lstSMS.Items(CInt(intj)).SubItems.Add((.Fields("PHONE").Value))

                        intj = intj + 1

                        .MoveNext()
                    Loop
                End With
                rs.Close()
                rs = Nothing

        End Select

        ResList(lstSMS)


        Exit Sub
err_:
        MsgBox(Err.Description)
    End Sub



    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click

        On Error GoTo err_

        lvModem.Items.Clear()


        sMODEM = ""
        sCOMMODEM = ""
        sMODEMiD = ""

        Dim intj As Integer = 0

        Dim ports As String() = SerialPort.GetPortNames()

        Dim port As String
        For Each port In ports

            lvModem.Items.Add(port)
            lvModem.Items(CInt(intj)).SubItems.Add("")

            intj = intj + 1
        Next port


        Exit Sub
err_:
        MsgBox(Err.Description)

    End Sub


End Class

