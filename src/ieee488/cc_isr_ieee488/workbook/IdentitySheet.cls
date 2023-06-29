VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_repeatCountCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_receiveTimeoutCell As String = "B4"
Private Const m_errorMessageCell As String = "B5"

Private Const m_socketAddressCell As String = "C2"
Private Const m_connectedSocketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "G2"
Private Const m_receivedMessageCell As String = "H2"
Private Const m_statusByteCell As String = "F3"
Private Const m_averageElapsedTimeCell As String = "H3"
Private Const m_elapsedTimeCell As String = "H4"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_identityCell As String = "H2"

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    Dim timeout As Integer
    
    timeout = Range(m_receiveTimeoutCell).value
    repeatCount = Range(m_repeatCountCell).value
    host = Range(m_inputHostCell).value
    port = Range(m_inputPortCell).value
    
    ' read identity user session read.
    Ieee488SessionReadIdentity host, port, repeatCount
    
End Sub

''' <summary>   Uses the IEEE 4888 session to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="host">   The host IPv4 dotted address. </param>
''' <param name="port">   the host port number. </param>
''' <param name="repeatCount">   The number of time to repeat the query. </param>
Private Sub Ieee488SessionReadIdentity(ByVal host As String, ByVal port As Long, ByVal repeatCount As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = ""
    Range(m_connectedSocketIdCell).value = ""
    Range(m_connectedCell).value = ""
    Range(m_sentMessageCell).value = ""
    Range(m_receivedMessageLengthCell).value = ""
    Range(m_receivedMessageCell).value = ""
    Range(m_averageElapsedTimeCell).value = ""
    Range(m_elapsedTimeCell).value = ""
    Range(m_errorMessageCell).value = ""

    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch()
    Dim totalMilliseconds As Double
    
    Dim Session As New Ieee488Session
    
    Session.OpenConnection host, port
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = Session.SocketId
   
    ' report the connection state
    Range(m_connectedCell).value = Session.Connected
    
    If Session.Connected Then
    
        Range(m_socketAddressCell).value = Session.ViSession.SocketAddress
    
        Dim command As String
        command = "*IDN?"
        Range(m_sentMessageCell).value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
           
            loopCount = loopCount + 1
            sw.Restart
            
            Dim Identity As String
            
            If Me.UsingViSessionIOCheckBox Then
                Identity = Session.ViSession.QueryLine(command)
            Else
                Identity = Session.QueryLine(command)
            End If
            
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(Identity)
            Range(m_receivedMessageCell).value = Identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(totalMilliseconds / loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            errorMessage = "Reading identity failed using Tcp session " & UserDefinedErrors.FormattedErr & _
                           " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    If Not Session Is Nothing Then
        If Session.Connected Then
            Session.CloseConnection
        End If
    End If
    
    ' terminate the session
    Set Session = Nothing
    
    If Len(errorMessage) > 0 Then
        Range(m_errorMessageCell).value = errorMessage
    End If

End Sub


