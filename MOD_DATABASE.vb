Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Module MOD_DATABASE
    Public BasePath As String
    Public Base_Name As String
    Public DB7 As Connection
    Public ConNect As String
    Public DATAB As Boolean = False

    Private b As Byte() = Convert.FromBase64String("0JfQuNC80L7QstCw0LvQuNCX0LLQtdGA0LjQktCv0LzQtTAsNdC30LDRj9GG")
    ' Private sCRTKey As String = System.Text.Encoding.UTF8.GetString(b)


    Public Sub LoadDatabase(Optional ByRef sFile As String = "")

        On Error GoTo ERR1

        Base_Name = "sms.mdb"
        sFile = Base_Name
        BasePath = Directory.GetParent(System.Windows.Forms.Application.ExecutablePath).ToString '& "\database\"

        Dim MyShadowPassword As String
        MyShadowPassword = ""

        DB7 = New Connection

        DB7.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BasePath & "\" & sFile & ";Jet OLEDB:Database Password=" & MyShadowPassword & ";")

        DATAB = True

        Exit Sub
ERR1:
        DATAB = False
        MsgBox(Err.Description)
    End Sub

    Public Sub UnLoadDatabase(Optional ByRef sFile As String = "")

        On Error GoTo Err_

        Select Case DATAB

            Case True

                DB7.Close()
                DB7 = Nothing

            Case Else


        End Select


        Exit Sub
Err_:
    End Sub


    Public Function EncryptString(ByVal sTEXTBLOC As String) As String

        Try

            'Dim DESkey As New DESCryptoServiceProvider()
            'DESkey.IV = UnicodeEncoding.Unicode.GetBytes(Mid(sCRTKey, 1, 8)) 'вектор
            'DESkey.Key = UnicodeEncoding.Unicode.GetBytes(Mid(sCRTKey, 9, 16)) 'Ключ
            'Dim DinBlock() As Byte = UnicodeEncoding.Unicode.GetBytes(sTEXTBLOC)
            'Dim DEStransForm As ICryptoTransform = DESkey.CreateEncryptor()
            'Dim DutBlock() As Byte = DEStransForm.TransformFinalBlock(DinBlock, 0, DinBlock.Length)

            'Return Convert.ToBase64String(DutBlock)




            Dim AESkey As New AesCryptoServiceProvider()
            AESkey.IV = UnicodeEncoding.Unicode.GetBytes(Mid(System.Text.Encoding.UTF8.GetString(b), 1, 8)) 'вектор
            AESkey.Key = UnicodeEncoding.Unicode.GetBytes(Mid(System.Text.Encoding.UTF8.GetString(b), 9, 16)) 'Ключ
            Dim inBlock() As Byte = UnicodeEncoding.Unicode.GetBytes(sTEXTBLOC)
            Dim AEStransForm As ICryptoTransform = AESkey.CreateEncryptor()
            Dim outBlock() As Byte = AEStransForm.TransformFinalBlock(inBlock, 0, inBlock.Length)

            Return Convert.ToBase64String(outBlock)

        Catch ex As Exception

            Return "ERR_" + ex.Message
        End Try

    End Function

    Public Function DecryptBytes(sTEXTBLOC As String) As String

        Try

            Dim AESkey As New AesCryptoServiceProvider()
            AESkey.IV = UnicodeEncoding.Unicode.GetBytes(Mid(System.Text.Encoding.UTF8.GetString(b), 1, 8)) 'вектор
            AESkey.Key = UnicodeEncoding.Unicode.GetBytes(Mid(System.Text.Encoding.UTF8.GetString(b), 9, 16)) 'Ключ
            Dim inBytes() As Byte = Convert.FromBase64String(sTEXTBLOC)
            Dim AEStransForm As ICryptoTransform = AESkey.CreateDecryptor()
            Dim outBlock() As Byte = AEStransForm.TransformFinalBlock(inBytes, 0, inBytes.Length)

            Return UnicodeEncoding.Unicode.GetString(outBlock)

        Catch ex As Exception

            Return "ERR_" + ex.Message
        End Try

    End Function

    Public Sub ResList(ByVal resizingListView As ListView)

        Dim columnIndex As Integer

        For columnIndex = 1 To resizingListView.Columns.Count - 1
            resizingListView.AutoResizeColumn(columnIndex, ColumnHeaderAutoResizeStyle.HeaderSize)
        Next

    End Sub


    Public Function ExportListViewToExcel(ByVal MyListView As ListView, ByVal sTXT As String)

        'Dim ExcelReport As Excel.ApplicationClass

        ' Const MAX_COLOURS As Int16 = 40

        Const MAX_COLUMS As Int16 = 254

        Dim i As Integer

        Dim New_Item As Windows.Forms.ListViewItem

        Dim TempColum As Int16

        Dim ColumLetter As String

        Dim TempRow As Int16

        Dim TempColum2 As Int16

        Dim AddedColours As Int16 = 1

        Dim MyColours As Hashtable = New Hashtable

        Dim AddNewBackColour As Boolean = True

        Dim AddNewFrontColour As Boolean = True

        'Dim BackColour As String

        'Dim FrontColour As String

        '##########################

        Dim chartRange As Excel.Range

        '##########################

        Dim ExcelReport As Excel.Application

        'ExcelReport = New Excel.ApplicationClass

        ExcelReport = New Excel.Application

        ExcelReport.Visible = True

        ExcelReport.Workbooks.Add()

        ColumLetter = ""

        'ExcelReport.Worksheets("Sheet1").Select()

        'ExcelReport.Sheets("Sheet1").Name = sTXT

        i = 0

        Do Until i = MyListView.Columns.Count

            If i > MAX_COLUMS Then

                MsgBox("Too many Colums added")

                Exit Do

            End If

            TempColum = i

            TempColum2 = 0

            Do While TempColum > 25

                TempColum -= 26

                TempColum2 += 1

            Loop

            ColumLetter = Chr(97 + TempColum)

            If TempColum2 > 0 Then ColumLetter = Chr(96 + TempColum2) & ColumLetter

            ExcelReport.Range(ColumLetter & 3).Value = MyListView.Columns(i).Text

            'ExcelReport.Range(ColumLetter & 3).Font.Name = MyListView.Font.Name

            ' ExcelReport.Range(ColumLetter & 3).Font.Size = MyListView.Font.Size + 2

            ExcelReport.Range(ColumLetter & 3).Font.Bold = True

            chartRange = ExcelReport.Range(ColumLetter & 3)

            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

            i += 1

        Loop

        '###############################################

        'Вставляем заголовок

        '###############################################

        'Устанавливаем диапазон ячеек

        chartRange = ExcelReport.Range("A1", ColumLetter & 2)

        'Объединяем ячейки

        chartRange.Merge()

        'Вставляем текст

        chartRange.FormulaR1C1 = sTXT

        'Выравниваем по центру

        chartRange.HorizontalAlignment = 3

        chartRange.VerticalAlignment = 2

        'Устанавливаем шрифт

        ExcelReport.Range("A1").Font.Name = MyListView.Font.Name

        'Увеличиваем шрифт

        ExcelReport.Range("A1").Font.Size = MyListView.Font.Size + 4

        'Делаем шрифт жирным

        ExcelReport.Range("A1").Font.Bold = True

        '###############################################
        '###############################################

        TempRow = 4

        For Each New_Item In MyListView.Items

            i = 0

            Do Until i = New_Item.SubItems.Count

                If i > MAX_COLUMS Then

                    MsgBox("Too many Colums added")

                    Exit Do

                End If

                TempColum = i

                TempColum2 = 0

                Do While TempColum > 25

                    TempColum -= 26

                    TempColum2 += 1

                Loop

                ColumLetter = Chr(97 + TempColum)

                If TempColum2 > 0 Then ColumLetter = Chr(96 + TempColum2) & ColumLetter

                ExcelReport.Range(ColumLetter & TempRow).Value = New_Item.SubItems(i).Text

                chartRange = ExcelReport.Range(ColumLetter & TempRow)

                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

                i += 1

            Loop

            TempRow += 1

        Next

        ExcelReport.Cells.Select()

        ExcelReport.Cells.EntireColumn.AutoFit()

        ExcelReport.Cells.Range("A1").Select()

    End Function

End Module
