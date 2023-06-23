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

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    
    repeatCount = Range("B1").value
    host = Range("B2").value
    port = Range("B3").value
    
    Range("C2").value = host
    Range("D2").value = port
    
    ' scpi.preset
    
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
    
    Range("E2").value = ""
    Range("F2").value = ""
    Range("G2").value = ""
    Range("H2").value = ""
    Range("I2").value = ""

    Dim sw As StopWatch: Set sw = cc_isr_Core.Constructor.CreateStopWatch()
    Dim totalMilliseconds As Double
    
    Dim Session As New Ieee488Session
    
    Session.OpenConnection host, port
    
    ' report the socket number
    Range("E2").value = Session.SocketId
   
    ' report the connection state
    Range("F2").value = Session.Connected
    
    If Session.Connected Then
    
        Dim command As String
        command = "*IDN?"
        Range("G2").value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            DoEvents
           
            loopCount = loopCount + 1
            sw.Restart
            
            Session.QueryIdentity
            
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range("H2").value = Len(Session.Identity)
            Range("I2").value = Session.Identity
    
        Loop
  
        ' command = ":SYST:PRES"
        ' count = SendMessage(command)
        ' Call opc
        
        Range("I3").value = Format(totalMilliseconds / loopCount, "0.0") + " ms"
        Range("I4").value = Format(totalMilliseconds, "0.0") + " ms"
    
    Else
        MsgBox "connection failed without reporting an exception"
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
        MsgBox errorMessage
    End If

End Sub


