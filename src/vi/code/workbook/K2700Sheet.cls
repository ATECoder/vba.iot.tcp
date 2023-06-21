VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K2700Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private device As K2700
Private m_session As Ieee488Session
Private sw As StopWatch

Friend Sub Worksheet_Activate()
    Set device = Constructor.CreateK2700()
    Set m_session = device.Session
    Set sw = cc_isr_Core.Constructor.CreateStopWatch()
    ' this toggles the enabled status of the controls.
    Me.ConnectedCheckBox_Click
    Me.CommandsComboBox.Clear
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ClearExecutionStateCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.IdentityQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompleteCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OperationCompletedQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.OptionsQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ResetKnownStateCommand
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.StandardEventEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.StandardEventStatusQueryCommand
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.ServiceRequestEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem Replace(Replace(cc_isr_Ieee488.StandardServiceEnableCommandFormat, _
                                        "{0}", &H7F), "{1}", &H7F)
    Me.CommandsComboBox.AddItem Replace(Replace(cc_isr_Ieee488.StandardServiceEnableCompleteCommandFormat, _
                                        "{0}", &H7F), "{1}", &H7F)
    Me.CommandsComboBox.AddItem Replace(cc_isr_Ieee488.OperationCompleteEnableCommandFormat, "{0}", &H7F)
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestEnableQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.ServiceRequestQueryCommand
    Me.CommandsComboBox.AddItem cc_isr_Ieee488.WaitCommand
    Me.CommandsComboBox.AddItem ":FUNC 'RES',(@101:120)"
    Me.CommandsComboBox.AddItem ":FUNC 'RES',(@201:220)"
   
End Sub

Private Sub Worksheet_Deactivate()
    Set sw = Nothing
    Set m_session = Nothing
    Set device = Nothing
End Sub

Friend Sub RestartStopWatch()
    Range("J2").value = ""
    If sw Is Nothing Then
        Me.Worksheet_Activate
    End If
    sw.Restart
End Sub

Friend Sub ReadStopWatch()
    Range("J2").value = Format(sw.ElapsedMilliseconds, "0.0")
End Sub

Private Sub CLSButton_Click()
    Me.RestartStopWatch
    m_session.ClearExecutionState
    Me.ReadStopWatch
End Sub

Friend Sub ConnectedCheckBox_Click()
    
    Me.ConnectedCheckBox.value = device.Connected
    
    ' report the connection state
    Range("F2").value = device.Connected
        
    Me.QueryIdentityButton.Enabled = device.Connected
    Me.RSTButton.Enabled = device.Connected
    Me.SDCButton.Enabled = device.Connected
    Me.CLSButton.Enabled = device.Connected
    Me.ReadCardsButton.Enabled = device.Connected
    Me.QueryButton.Enabled = device.Connected
    Me.ReadButton.Enabled = device.Connected
    Me.ReadStatusByteButton.Enabled = device.Connected
    Me.WriteButton.Enabled = device.Connected
    Me.SetScansButton.Enabled = device.Connected
    
End Sub

Private Sub QueryIdentityButton_Click()
    Me.RestartStopWatch
    Range("I2").value = m_session.QueryIdentity
    Me.ReadStopWatch
End Sub

Private Sub ReadCardsButton_Click()
    Me.RestartStopWatch
    device.ReadCards
    Me.ReadStopWatch
    Range("C15").value = device.Card1
    Range("C16").value = device.Card2
End Sub

Private Sub RSTButton_Click()
    Me.RestartStopWatch
    m_session.ResetKnownState
    Me.ReadStopWatch
End Sub

Private Sub SDCButton_Click()
    Me.RestartStopWatch
    m_session.SelectiveDeviceClear
    Me.ReadStopWatch
End Sub

Private Sub SetScansButton_Click()
    Me.RestartStopWatch
    device.DefineResistanceScanLists
    Me.ReadStopWatch
    Range("C17").value = device.Card1ScanList
    Range("C18").value = device.Card2ScanList
End Sub

Private Sub ToggleConnectionButton_Click()

    Me.RestartStopWatch

    Dim host As String
    Dim port As Integer
    Dim repeatCount As Integer
    
    repeatCount = Range("B1").value
    host = Range("B2").value
    port = Range("B3").value
    
    If device.Connected Then
        
        device.CloseConnection
        
        Range("C2").value = ""
        Range("D2").value = ""
        Range("E2").value = ""
        
    Else
    
        Range("C2").value = host
        Range("D2").value = port
        
        device.OpenConnection host, port
         
        ' report the socket number
        Range("E2").value = m_session.ViSession.SocketId
    End If
    
    Me.ReadStopWatch
    Me.ConnectedCheckBox_Click
    
End Sub

Private Sub WriteButton_Click()
    Range("G2").value = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    m_session.WriteLine msg
    Me.ReadStopWatch
    Range("G2").value = msg
    If Me.AutoStatusReadCheckBox.value Then
        Me.ReadStatusByteButton_Click
    End If
End Sub

Private Sub ReadButton_Click()
    Range("I2").value = ""
    Me.RestartStopWatch
    Range("I2").value = m_session.Read
    Me.ReadStopWatch
End Sub

Friend Sub ReadStatusByteButton_Click()
    Me.RestartStopWatch
    Range("G3").value = m_session.ReadStatusByte
    Me.ReadStopWatch
End Sub

Private Sub QueryButton_Click()
    Me.SentMessage = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    Me.ReceivedMessage = m_session.QueryLine(msg)
    Me.ReadStopWatch
    Me.SentMessage = msg
End Sub

Friend Property Get host() As String
    host = Range("B2").value
End Property

Friend Property Get port() As Integer
    port = CInt(Range("B3").value)
End Property

Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range("I2").value
End Property

Friend Property Let ReceivedMessage(ByVal value As String)
    Range("I2").value = value
    Range("H2").value = Len(value)
End Property

Friend Property Get SentMessage() As String
    SentMessage = Range("G2").value
End Property

Friend Property Let SentMessage(ByVal value As String)
    Range("G2").value = value
End Property

