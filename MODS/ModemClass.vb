
Imports System.IO.Ports
Imports System.Management

Public Class ModemClass
    Public Structure modemLog
        Public command As String

        Public modemReply As String

        Public [error] As Boolean
    End Structure

    Public Config As ConfigClass

    Public Port As SerialPort

    Private mLog As ModemClass.modemLog

    Private devID As String = ""

    Public ReadOnly Property connected() As Boolean
        Get
            Return Me.Port.IsOpen
        End Get
    End Property

    Public ReadOnly Property [error]() As Boolean
        Get
            Return Me.mLog.[error]
        End Get
    End Property

    Public ReadOnly Property lastCommand() As String
        Get
            Return Me.mLog.command
        End Get
    End Property

    Public ReadOnly Property modemReply() As String
        Get
            Return Me.mLog.modemReply
        End Get
    End Property

    Public ReadOnly Property COMPort() As String
        Get
            Return Me.Port.PortName
        End Get
    End Property

    Public ReadOnly Property deviceID() As String
        Get
            Return Me.devID
        End Get
    End Property

    Public ReadOnly Property LogString() As String
        Get
            Return String.Concat(New String() {"Command: ", Me.mLog.command, vbCr & vbLf & "Error: ", Me.mLog.[error].ToString(), vbCr & vbLf & "Modem Reply: ", Me.mLog.modemReply, _
             vbCr & vbLf & "----------" & vbCr & vbLf})
        End Get
    End Property

    Public Sub New(ByVal ModemConfig As ConfigClass)
        Me.Config = ModemConfig
        Me.mLog.modemReply = ""
    End Sub

    Public Function Connect() As Boolean
        Me.mLog.command = "ConnectToModem"
        Me.mLog.modemReply = "Connection error!"
        Me.mLog.[error] = True

        'Dim text As String = ModemClass.GetModemSerialPortName(Me.Config.Modem, Me.devID)

        'If text.Length = 0 Then
        '    text = Me.Config.Modem
        'End If

        If sCOMMODEM.Length > 0 Then
            Try
                Me.Port = New SerialPort(sCOMMODEM)
                Me.Port.BaudRate = Me.Config.PortBaudRate
                Me.Port.DataBits = Me.Config.PortDataBits
                Me.Port.Parity = Me.Config.PortParity
                Me.Port.StopBits = Me.Config.PortStopBits
                Me.Port.NewLine = Environment.NewLine
                Me.Port.Open()
                Me.mLog.modemReply = Convert.ToString("Соединение установлено на порту: ") & sCOMMODEM
                Me.mLog.[error] = False
                Dim result As Boolean = Me.Port.IsOpen
                Return result
            Catch ex As Exception
                Me.mLog.modemReply = (Me.mLog.modemReply & Convert.ToString("/n")) + ex.Message
                Dim result As Boolean = False
                Return result
            End Try
            Return False
        End If
        Return False

    End Function

    Public Sub Disconnect()
        Try
            If Me.Port.IsOpen Then
                Me.Port.Close()
            End If
        Catch
        End Try
    End Sub

    Public Function WaitBeforeContains(ByVal contains As String) As Boolean
        Dim t As DateTime = DateTime.Now.AddSeconds(30.0)
        Dim text As String = ""
        While True
            text += Me.Port.ReadExisting()
            If DateTime.Now > t OrElse text.Contains("ERROR" & vbCr & vbLf) Then
                Exit While
            End If
            If text.Contains(contains) Then
                GoTo Block_2
            End If
        End While
        Me.mLog.modemReply = text
        Return False
Block_2:
        Me.mLog.modemReply = text
        Return True
    End Function

    Public Function SendCommand(ByVal Command As String, Optional ByVal OKCondition As String = vbCr & vbLf & "OK" & vbCr & vbLf) As Boolean
        Me.mLog.modemReply = ""
        Me.mLog.command = Command
        Me.mLog.[error] = True
        Dim result As Boolean
        Try
            Me.Port.WriteLine(Command)
            Me.mLog.[error] = Not Me.WaitBeforeContains(OKCondition)
            result = Not Me.mLog.[error]
        Catch
            result = Me.mLog.[error]
        End Try
        Return result
    End Function

    Public Shared Function GetModemSerialPortName(ByVal ModemName As String, ByRef DeviceID As String) As String
        Dim managementObjectSearcher As New ManagementObjectSearcher((Convert.ToString("select * from Win32_POTSModem where Description like ""%") & ModemName) + "%""")
        Dim managementObjectCollection As ManagementObjectCollection = managementObjectSearcher.[Get]()
        If managementObjectCollection.Count = 0 Then
            Return ""
        End If
        Using enumerator As ManagementObjectCollection.ManagementObjectEnumerator = managementObjectCollection.GetEnumerator()
            While enumerator.MoveNext()
                Dim managementObject As ManagementObject = DirectCast(enumerator.Current, ManagementObject)
                If managementObject("Status").ToString() = "OK" Then
                    DeviceID = managementObject("DeviceID").ToString()
                    Return managementObject("AttachedTo").ToString()
                End If
            End While
        End Using
        Return ""
    End Function
End Class
