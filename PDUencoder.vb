
Namespace SMS
    Namespace Encoder
        Public Class SMS
#Region "Enums"
            Public Enum ENUM_TP_VPF
                Relative_Format = 16    'b4=1 b3=0
            End Enum

            Public Enum ENUM_TP_SRI
                Request_SMS_Report = 32
                No_SMS_Report = 0
            End Enum

            Public Enum ENUM_TP_DCS
                DefaultAlphabet = 0
                UCS2 = 8
            End Enum

            Public Enum ENUM_TP_VALID_PERIOD
                OneHour = 11 '0 to 143:(TP-VP+1)*5Min
                ThreeHours = 29
                SixHours = 71
                TwelveHours = 143
                OneDay = 167
                OneWeek = 196
                Maximum = 255
            End Enum
#End Region

#Region "Protected Data"
            Protected SC_Number As String  'Note the plus!
            Protected TP_MTI As Byte = 1
            Protected TP_RD As Byte = 0
            Protected TP_VPF As Byte = 16
            Protected TP_UDHI As Byte
            Protected TP_SRR As Byte
            Protected TP_MR As Integer
            Protected TP_DA As String
            Protected TP_PID As Byte
            Protected TP_DCS As Byte
            Protected TP_VP As Byte
            Protected TP_UDL As Integer
            Protected TP_UD As String
#End Region

#Region "Properties"

            Public Property ServiceCenterNumber() As String
                Get
                    Return SC_Number
                End Get
                Set(ByVal Value As String)
                    'Convert an ServiceCenterNumber to PDU Code
                    If InStr(Value, "+") Then
                        SC_Number = "91"
                    Else
                        SC_Number = "81"
                    End If

                    Value = Mid(Value, 2)
                    Dim i As Integer
                    If (Value.Length Mod 2) = 1 Then
                        Value += "F"
                    End If
                    For i = 1 To Value.Length Step 2
                        SC_Number += Swap(Mid(Value, i, 2))
                    Next
                    SC_Number = ByteToHex((SC_Number.Length - 2) / 2 + 1) + SC_Number
                End Set

            End Property

            Public Property TP_Status_Report_Request() As ENUM_TP_SRI
                Get
                    Return TP_SRR
                End Get
                Set(ByVal Value As ENUM_TP_SRI)
                    TP_SRR = Value
                End Set
            End Property

            Public Property TP_Message_Reference() As Integer
                Get
                    Return TP_MR
                End Get
                Set(ByVal Value As Integer)
                    TP_MR = Value
                End Set
            End Property

            Public Property TP_Destination_Address() As String
                Get
                    Return TP_DA
                End Get
                Set(ByVal Value As String)
                    TP_DA = ""

                    If InStr(Value, "+") Then
                        TP_DA += "91"
                    Else
                        TP_DA += "81"
                    End If
                    Value = Value.Replace("+", "")
                    TP_DA = Format(Value.Length, "X2") + TP_DA
                    Dim i As Integer
                    If (Value.Length Mod 2) = 1 Then
                        Value += "F"
                    End If
                    For i = 1 To Value.Length Step 2
                        TP_DA += Swap(Mid(Value, i, 2))
                    Next
                End Set
            End Property


            Public Property TP_Data_Coding_Scheme() As ENUM_TP_DCS
                Get
                    Return TP_DCS
                End Get
                Set(ByVal Value As ENUM_TP_DCS)
                    TP_DCS = Value
                End Set
            End Property

            Public Property TP_Validity_Period() As ENUM_TP_VALID_PERIOD
                Get
                    Return TP_VP
                End Get
                Set(ByVal Value As ENUM_TP_VALID_PERIOD)
                    TP_VP = Value
                End Set
            End Property

            Public Overridable Property TP_User_Data() As String
                Get
                    Return TP_UD
                End Get
                Set(ByVal Value As String)
                    Select Case TP_DCS
                        Case Is = ENUM_TP_DCS.DefaultAlphabet
                            TP_UDL = Value.Length
                            TP_UD = Encode7Bit(Value)
                        Case Is = ENUM_TP_DCS.UCS2
                            TP_UDL = Value.Length * 2
                            TP_UD = EncodeUCS2(Value)
                        Case Else
                            TP_UD = Value
                    End Select
                End Set
            End Property
#End Region

#Region "Functions"
            Public Shared Function CheckForEncoding(ByVal Content As String) As SMS.ENUM_TP_DCS
                Dim i As Integer
                For i = 1 To Content.Length
                    If Asc(Mid(Content, i, 1)) < 0 Then
                        Return SMS.ENUM_TP_DCS.UCS2
                    End If
                Next
                Return SMS.ENUM_TP_DCS.DefaultAlphabet
            End Function

            Public Overridable Function GetSMSPDUCode() As String
                Dim PDUCode As String
                'Check User Data Length
                If TP_DCS = ENUM_TP_DCS.DefaultAlphabet Then
                    If TP_UD.Length > 280 Then Throw New Exception("User Data is TOO LONG for SMS")
                End If
                If TP_DCS = ENUM_TP_DCS.UCS2 Then
                    If TP_UD.Length > 280 Then Throw New Exception("User Data is TOO LONG for SMS")
                End If
                'Make PDUCode
                PDUCode = SC_Number
                PDUCode += FirstOctet()
                PDUCode += Format(TP_MR, "X2")
                PDUCode += TP_DA
                PDUCode += Format(TP_PID, "X2")
                PDUCode += Format(TP_DCS, "X2")
                PDUCode += Format(TP_VP, "X2")
                PDUCode += Format(TP_UDL, "X2")
                PDUCode += TP_UD
                Return PDUCode
            End Function

            Public Overridable Function FirstOctet() As String
                Return ByteToHex(TP_MTI + TP_VPF + TP_SRR + TP_UDHI)
            End Function

            Shared Function ByteToHex(ByVal aByte As Byte) As String
                Dim result As String
                result = Format(aByte, "X2")
                Return result
            End Function

#Region "Encode7Bit"
            Shared Function Encode7Bit(ByVal Content As String) As String
                'Prepare
                Dim CharArray As Char() = Content.ToCharArray
                Dim c As Char
                Dim t As String
                For Each c In CharArray
                    t = CharTo7Bits(c) + t
                Next
                'Add "0"
                Dim i As Integer
                If (t.Length Mod 8) <> 0 Then
                    For i = 1 To 8 - (t.Length Mod 8)
                        t = "0" + t
                    Next
                End If
                'Split into 8bits
                Dim result As String
                For i = t.Length - 8 To 0 Step -8
                    result = result + BitsToHex(Mid(t, i + 1, 8))
                Next
                Return result
            End Function

            Shared Function BitsToHex(ByVal Bits As String) As String
                'Convert 8Bits to Hex String
                Dim i, v As Integer
                For i = 0 To Bits.Length - 1
                    v = v + Val(Mid(Bits, i + 1, 1)) * 2 ^ (7 - i)
                Next
                Dim result As String
                result = Format(v, "X2")
                Return result
            End Function

            Shared Function CharTo7Bits(ByVal c As Char) As String
                If c = "@" Then Return "0000000"
                Dim Result As String
                Dim i As Integer
                For i = 0 To 6
                    If (Asc(c) And 2 ^ i) > 0 Then
                        Result = "1" + Result
                    Else
                        Result = "0" + Result
                    End If
                Next
                Return Result
            End Function
#End Region

            Shared Function EncodeUCS2(ByVal Content As String) As String
                Dim i, j, v As Integer
                Dim Result, t As String
                For i = 1 To Content.Length
                    v = AscW(Mid(Content, i, 4))
                    t = Format(v, "X4")
                    Result += t
                Next
                Return Result
            End Function

            Shared Function Swap(ByRef TwoBitStr As String) As String
                'Swap two bit like "EF" TO "FE"
                Dim c() As Char = TwoBitStr.ToCharArray
                Dim t As Char
                t = c(0)
                c(0) = c(1)
                c(1) = t
                Return (c(0) + c(1)).ToString
            End Function
#End Region
        End Class

        Public Class ConcatenatedShortMessage
            Inherits SMS
            Private TotalMessages As Integer

            Public Function GetEMSPDUCode() As String()
                Select Case tp_dcs
                    Case ENUM_TP_DCS.UCS2
                        TotalMessages = (TP_UD.Length / 4) \ 66 + ((TP_UD.Length / 4 Mod 66) = 0)
                    Case ENUM_TP_DCS.DefaultAlphabet
                        TotalMessages = (tp_ud.Length \ 266) - ((tp_ud.Length Mod 266) = 0)
                End Select

                Dim Result(TotalMessages) As String
                Dim tmpTP_UD As String
                Dim i As Integer
                TP_UDHI = 2 ^ 6
                Dim Reference As Integer = Rnd(1) * 65536   '16bit Reference Number 'See 3GPP Document
                For i = 0 To TotalMessages
                    Select Case tp_dcs
                        Case ENUM_TP_DCS.UCS2
                            tmpTP_UD = Mid(TP_UD, i * 66 * 4 + 1, 66 * 4)
                            'When TP_UDL is odd, the max length of an Unicode string in PDU code is 66 Charactor.See [3GPP TS 23.040 V6.5.0 (2004-09] 9.2.3.24.1
                        Case ENUM_TP_DCS.DefaultAlphabet
                            tmpTP_UD = Mid(tp_ud, i * 133 * 2 + 1, 133 * 2)
                    End Select
                    Result(i) = SC_Number
                    Result(i) += FirstOctet()
                    Result(i) += Format(TP_MR, "X2")
                    'Next segement TP_MR must be increased
                    'TP_MR += 1
                    Result(i) += TP_DA
                    Result(i) += Format(TP_PID, "X2")
                    Result(i) += Format(TP_DCS, "X2")
                    Result(i) += Format(TP_VP, "X2")
                    If tp_dcs = ENUM_TP_DCS.UCS2 Then
                        TP_UDL = tmpTP_UD.Length / 2 + 6 + 1 '6:IE
                    End If
                    If tp_dcs = ENUM_TP_DCS.DefaultAlphabet Then
                        TP_UDL = Fix((tmpTP_UD.Length + 7 * 2) * 4 / 7)   '6:length of IE
                        ''#################################
                        ''still problem here:
                        ''when the charcter is several times of 7 of the last message, tp_udl will not correct!
                        ''to correct this problem i write some code below. that's may not perfect solution.
                        ''#################################
                        'If i = TotalMessages And ((tmpTP_UD.Length Mod 14) = 0) Then
                        '    tp_udl -= 1
                        'End If
                    End If
                    Result(i) += Format(TP_UDL, "X2")
                    Result(i) += "060804" 'TP_UDHL and some of Concatenated message
                    Result(i) += Format(Reference, "X4")
                    Result(i) += Format(TotalMessages + 1, "X2")
                    Result(i) += Format(i + 1, "X2")
                    Result(i) += tmpTP_UD
                Next
                Return Result
            End Function

            Public Overrides Function FirstOctet() As String
                Return ByteToHex(TP_MTI + TP_VPF + TP_SRR + TP_UDHI)
            End Function
        End Class
    End Namespace
End Namespace