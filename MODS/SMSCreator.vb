Imports System.Collections.Generic
Imports System.Text

Public Class SMSCreator
    Public Messages As List(Of String)

    Public TextMode As Boolean

    Public FlashMessage As Boolean

    Public statusReport As Boolean

    Public RecipientPhoneNumberFormat As String

    Public SMSCPhoneNumberFormat As String

    Public LeadingZeroIfSMSCNumberNotPresent As Boolean

    Public SCA As String

    Public Sub New()
        Me.Messages = New List(Of String)()
        Me.Messages.Clear()
        Me.TextMode = False
        Me.FlashMessage = False
        Me.RecipientPhoneNumberFormat = "91"
        Me.SMSCPhoneNumberFormat = "91"
        Me.LeadingZeroIfSMSCNumberNotPresent = True
        Me.SCA = ""
    End Sub

    Public Sub CreateSMS(ByVal RecipientPhoneNumber As String, ByVal MessageText As String, Optional ByVal SMSCenterPhoneNumber As String = "")
        RecipientPhoneNumber = RecipientPhoneNumber.Replace("+", "")
        SMSCenterPhoneNumber = SMSCenterPhoneNumber.Replace("+", "")
        Dim uTF8Encoding As New UTF8Encoding()
        Dim byteCount As Integer = uTF8Encoding.GetByteCount(MessageText)
        Dim flag As Boolean = MessageText.Length <> byteCount
        Me.TextMode = (Not flag AndAlso MessageText.Length <= 160 AndAlso Me.TextMode)
        Me.TextMode = (Not Me.FlashMessage AndAlso Me.TextMode)
        Me.Messages.Clear()
        If Me.TextMode Then
            Me.Messages.Add(MessageText & ChrW(26))
            Return
        End If
        Dim flag2 As Boolean = If(flag, (MessageText.Length <= 70), (MessageText.Length <= 160))
        Dim text As String = If(flag2, (If(Me.statusReport, "21", "01")), (If(Me.statusReport, "61", "41")))
        If SMSCenterPhoneNumber.Length > 0 Then
            Dim text2 As String = Me.EncodePhoneNumber(SMSCenterPhoneNumber)
            Me.SCA = (Me.SMSCPhoneNumberFormat.Length \ 2 + text2.Length \ 2).ToString("X2") & Me.SMSCPhoneNumberFormat & text2
        ElseIf Me.LeadingZeroIfSMSCNumberNotPresent Then
            Me.SCA = "00"
        End If
        Dim text3 As String = "00"
        Dim text4 As String = RecipientPhoneNumber.Length.ToString("X2") & Me.RecipientPhoneNumberFormat & Me.EncodePhoneNumber(RecipientPhoneNumber)
        Dim text5 As String = "00"
        Dim text6 As String = (If(Me.FlashMessage, "1", "0")) & (If(flag, "8", "0"))
        Dim text7 As String = ""
        If flag2 Then
            Me.Messages.Add(MessageText)
        Else
            Dim num As Integer = If(flag, 67, 152)
            While MessageText.Length > 0
                If MessageText.Length > num Then
                    Me.Messages.Add(MessageText.Substring(0, num))
                    MessageText = MessageText.Substring(num)
                Else
                    Me.Messages.Add(MessageText)
                    MessageText = ""
                End If
            End While
        End If
        Dim random As New Random()
        Dim text8 As String = If(flag, ("050003" & random.[Next](256).ToString("X2")), ("060804" & random.[Next](65536).ToString("X4")))
        text8 += Me.Messages.Count.ToString("X2")
        For i As Integer = 0 To Me.Messages.Count - 1
            Dim num2 As Integer = If(flag, (Me.Messages(i).Length * 2), Me.Messages(i).Length)
            Dim text9 As String = If(flag2, num2.ToString("X2"), (num2 + (If(flag, 6, 8))).ToString("X2"))
            Me.Messages(i) = (If(flag, Me.StringToUSC2(Me.Messages(i)), Me.String7To8(Me.Messages(i))))
            text3 = (If(flag2, text3, i.ToString("X2")))
            Dim str As String = String.Concat(New String() {text, text3, text4, text5, text6, text7, _
             text9})
            If Not flag2 Then
                str = str & text8 & (i + 1).ToString("X2")
            End If
            Me.Messages(i) = str & Me.Messages(i)
        Next
    End Sub

    Public Function Decode7bit(ByVal s As String) As String
        Dim array As Byte() = New Byte(s.Length - 1) {}
        For i As Integer = 0 To s.Length - 1 Step 2
            array(i \ 2) = Convert.ToByte(s.Substring(i, 2), 16)
        Next
        For j As Integer = 1 To array.Length - 3
            Dim num As Integer = CInt(array(j - 1) And 128)
            array(j - 1) = (array(j - 1) And 127)
            For k As Integer = j To array.Length - 2
                Dim num2 As Integer = CInt(array(k) And 128)
                array(k) = CByte(array(k) << 1)
                array(k) = (If((num > 0), (array(k) Or 1), array(k)))
                num = num2
            Next
        Next
        s = Encoding.ASCII.GetString(array).Replace(ControlChars.NullChar, " "c).Trim()
        Return s
    End Function

    Public Function EncodePhoneNumber(ByVal PhoneNumber As String) As String
        Dim text As String = ""
        If PhoneNumber.Length < 2 Then
            Return text
        End If
        If PhoneNumber.Length Mod 2 > 0 Then
            PhoneNumber += "F"
        End If
        For i As Integer = 0 To PhoneNumber.Length - 1 Step 2
            text = text & PhoneNumber(i + 1).ToString() & PhoneNumber(i).ToString()
        Next
        Return text.Trim()
    End Function

    Public Function StringToUSC2(ByVal str As String) As String
        Dim unicodeEncoding As New UnicodeEncoding()
        Dim bytes As Byte() = unicodeEncoding.GetBytes(str)
        For i As Integer = 0 To bytes.Length - 1 Step 2
            Dim b As Byte = bytes(i + 1)
            bytes(i + 1) = bytes(i)
            bytes(i) = b
        Next
        Return BitConverter.ToString(bytes).Replace("-", "")
    End Function

    Public Function String7To8_v1(ByVal str As String) As String
        Dim text As String = ""
        Dim aSCIIEncoding As New ASCIIEncoding()
        Dim bytes As Byte() = aSCIIEncoding.GetBytes(str)
        Dim i As Integer
        For i = 1 To bytes.Length - 1
            For j As Integer = bytes.Length - 1 To i Step -1
                Dim b As Byte = If((bytes(j) Mod 2 > 0), 128, 0)
                bytes(j - 1) = ((bytes(j - 1) And 127) Or b)
                bytes(j) = CByte(bytes(j) >> 1)
            Next
        Next
        i = 0
        While i < bytes.Length AndAlso bytes(i) <> 0
            text += bytes(i).ToString("X2")
            i += 1
        End While
        Return text
    End Function

    Public Function String7To8(ByVal str As String) As String
        Dim text As String = ""
        Dim aSCIIEncoding As New ASCIIEncoding()
        Dim bytes As Byte() = aSCIIEncoding.GetBytes(str)
        For i As Integer = 0 To bytes.Length - 2
            For j As Integer = i + 1 To bytes.Length - 1
                bytes(j - 1) = (bytes(j - 1) And 127)
                bytes(j) = bytes(j) And 127
                If (bytes(j) And 1) > 0 Then
                    bytes(j - 1) = (bytes(j - 1) Or 128)
                End If
                If j < bytes.Length - 1 AndAlso (bytes(j + 1) And 1) > 0 Then
                    bytes(j) = bytes(j) Or 128
                End If
                bytes(j) = CByte(bytes(j) >> 1)
            Next
            text += bytes(i).ToString("X2")
        Next
        text += bytes(bytes.Length - 1).ToString("X2")
        Return text.Substring(0, CInt(Math.Truncate(Math.Ceiling(CDbl(str.Length) * 0.875))) * 2)
    End Function

End Class
