Imports System
Imports System.Windows.Forms

Public Class ProgBar_HorizontalTEXT
    Inherits ProgressBar

    Protected Overrides ReadOnly Property CreateParams As System.Windows.Forms.CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style ' Or &H4
            Return cp
            'Return MyBase.CreateParams
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        MyBase.WndProc(m)
        If m.Msg = &HF Then

            Using g As Graphics = Graphics.FromHwnd(Me.Handle)

                Using sf As New StringFormat(StringFormatFlags.NoWrap)
                    sf.Alignment = StringAlignment.Center
                    sf.LineAlignment = StringAlignment.Center
                    g.DrawString(Me.Value.ToString & "%", Control.DefaultFont, Brushes.Black, Me.ClientRectangle, sf)
                End Using

            End Using



        End If


    End Sub
End Class
