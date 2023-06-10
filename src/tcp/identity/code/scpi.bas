Attribute VB_Name = "scpi"
Private hostname$
Private PortNumber As Integer
Private Const sheet = "Identity"
Private Const hostCell = "B2"
Private Const portCell = "B3"

Sub GetHostName()
    
    Sheets(sheet).Select
    
    Range("B2").Select
    hostname$ = ActiveCell.FormulaR1C1
    
End Sub

Sub GetPortNumber()
    
    Sheets(sheet).Select
    
    Range("B3").Select
    PortNumber = ActiveCell.Value
    
End Sub

Sub preset()

    On Error GoTo Finally
    
    Dim qpc As New StopWatch
    Dim idnQpc As New StopWatch
    qpc.Restart
    Dim x As Long
    Dim recvBuf As String * 1024
    
    Call StartIt
    Call GetHostName
    Call GetPortNumber
    
    Sheets(sheet).Select
    Range("C2").Value = hostname$
    Range("D2").Value = PortNumber
    
    Dim socketId As Long
    
    socketId = OpenSocket(hostname$, PortNumber)

    Range("E2").Value = socketId

    ' by sending a bad command, such as %IDNX we
    ' verified that the instrument is getting the command.
    
    idnQpc.Restart
    
    Dim count As Integer
    Dim command As String
    command = "*IDN?"
    count = SendCommand(command)
    
    Range("F2").Value = count
    Range("G2").Value = command
    
    ' presently, Winsock crashes here.
    
    count = Receive(recvBuf, 1024)
    
    Range("H2").Value = count
    Range("I2").Value = recvBuf
    
    ' command = ":SYST:PRES"
    ' count = SendCommand(command)
    ' Call opc
    
    Range("I3").Value = CStr(idnQpc.ElapsedMilliseconds) + "ms"
    Range("I4").Value = CStr(qpc.ElapsedMilliseconds) + "ms"

Finally:

    Call CloseConnection
    Call EndIt

End Sub

Sub opc()
'
' wait operation complete
'
    Dim x As Long
    Dim recvBuf As String * 10
    
    x = SendCommand("*OPC?")
    x = Receive(recvBuf, 10)

End Sub
