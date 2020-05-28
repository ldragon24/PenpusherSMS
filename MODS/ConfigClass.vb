
Imports System.IO.Ports

<Serializable()> _
    Public Class ConfigClass
    Public Modem As String = ""

    Public PortWriteTimeout As Integer = 500

    Public PortReadTimeout As Integer = timeout

    Public PortBaudRate As Integer = baudRate

    Public PortDataBits As Integer = 8

    Public PortParity As Parity

    Public PortStopBits As StopBits = StopBits.One

    Public PortHandshake As Handshake = Handshake.RequestToSend

    Public PortDtrEnable As Boolean = True

    Public PortRtsEnable As Boolean = True

    Public PortNewLine As String = Environment.NewLine

    Public TextMode As Boolean

    Public StatusReport As Boolean

    Public flashMessage As Boolean

    Public LastNumber As String = ""
End Class
