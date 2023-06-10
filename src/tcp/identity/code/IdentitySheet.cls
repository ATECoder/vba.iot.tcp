VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

''' <summary> Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    
    repeatCount = Range("B1").Value
    host = Range("B2").Value
    port = Range("B3").Value
    
    Range("C2").Value = host
    Range("D2").Value = port
    
    ' scpi.preset
    
    ReadIdentity host, port, repeatCount
    
End Sub

Private Sub ReadIdentity(ByVal host As String, ByVal port As Integer, ByVal repeatCount As Integer)

    On Error GoTo Finally
    
    Range("E2").Value = ""
    Range("F2").Value = ""
    Range("G2").Value = ""
    Range("H2").Value = ""
    Range("I2").Value = ""

    Dim maximumLength As Integer
    maximumLength = 1024
    Dim buffer As String * 1024
    
    Dim readCount As Integer
    
    Dim sw As New StopWatch
    Dim socket As New IPv4StreamSocket
    Dim totalMilliseconds As Double
    
    ' report the socket number
    Range("E2").Value = socket.socketId
    
    socket.OpenConnection host, port
    
    ' report the connection state
    Range("F2").Value = socket.Connected
    
    If socket.Connected Then
    
        Dim command As String
        command = "*IDN?"
        Range("G2").Value = command
        
        Dim loopCount As Integer
        Do While loopCount < repeatCount
            
            loopCount = loopCount + 1
            sw.Restart
            
            ' send the command
            socket.SendCommand (command & vbCrLf)
        
            ' receive the reading
            readCount = socket.Receive(buffer, maximumLength)
    
            totalMilliseconds = totalMilliseconds + sw.ElapsedMilliseconds
            
            Range("H2").Value = readCount
            Range("I2").Value = buffer
    
        Loop
  
        ' command = ":SYST:PRES"
        ' count = SendCommand(command)
        ' Call opc
        
        Range("I3").Value = Format(totalMilliseconds / loopCount, "0.0") + " ms"
        Range("I4").Value = CStr(totalMilliseconds) + " ms"
    
    Else
        MsgBox "connection failed without reporting an exception"
    End If
    
Finally:

    Dim errorMessage As String
    If Not Err Is Nothing And Err.Number <> 0 Then
        errorMessage = "Reading identity failed at " & Err.source & " with error number (0x" _
           & Hex(Err.Number) & "=" & CStr(Err.Number) & ") " & Err.Description & " Last DLL error = 0x" & Hex(Err.LastDllError)
    End If
    
    On Error Resume Next
    If socket.Connected Then
        socket.CloseConnection
    End If
    
    ' terminate the socket, which should release the winsock dll.
    Set socket = Nothing
    
    If Len(errorMessage) > 0 Then
        MsgBox errorMessage
    End If

End Sub

