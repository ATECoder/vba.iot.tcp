VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ieee488SheetBroken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_receiveTimeoutCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_connectedHostCell As String = "C2"
Private Const m_connectedPortCell As String = "D2"
Private Const m_connectedSocketIdCell As String = "E2"
Private Const m_sentMessageCell As String = "G2"
Private Const m_receivedMessageLengthCell As String = "H2"
Private Const m_receivedMessageCell As String = "I2"
Private Const m_statusByteCell As String = "G3"
Private Const m_elapsedTimeCell As String = "J2"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_connectedCell As String = "F2"
Private Const m_identityCell As String = "I2"

Private WithEvents m_session As cc_isr_Ieee488.Ieee488Session
Private m_stopWatch As StopWatch

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.Ieee488Session"/>. </value>
Public Property Get Session() As cc_isr_Ieee488.Ieee488Session
    If m_session Is Nothing Then
        Set m_session = cc_isr_Ieee488.Constructor.CreateIeee488Session
        ' this toggles the enabled status of the controls.
        Me.ConnectedCheckBox_Click
        ' populate the combo box commands
        PopulateCommandsComboBox
    End If
    Set Session = m_session
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Ieee488.VISession"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Ieee488.VISession"/>. </value>
Public Property Get VISession() As cc_isr_Ieee488.VISession
    Set VISession = Me.Session.VISession
End Property

''' <summary>   Gets the singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </summary>
''' <value>   The singleton instance of the <see cref="cc_isr_Core.StopWatch"/>. </value>
Public Property Get StopWatch() As cc_isr_Core.StopWatch
    If m_stopWatch Is Nothing Then
        Set m_stopWatch = cc_isr_Core.Constructor.CreateStopWatch
    End If
    Set StopWatch = m_stopWatch
End Property

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()
    Set m_stopWatch = Nothing
    If Not m_session Is Nothing Then m_session.Dispose
    Set m_session = Nothing
End Sub

Friend Sub Worksheet_Activate()
End Sub

''' <summary>   Populates the Commands combo box. </summary>
Friend Sub PopulateCommandsComboBox()
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
   
End Sub

''' <summary>   Restarts the stop watch and clears the elapsed time cell. </summary>
Friend Sub RestartStopWatch()
    Range(m_elapsedTimeCell).value = ""
    Me.StopWatch.Restart
End Sub

''' <summary>   Reads the stop watch and updates the elapsed time cell. </summary>
Friend Sub ReadStopWatch()
    Range(m_elapsedTimeCell).value = Format(Me.StopWatch.ElapsedMilliseconds, m_elapsedTimeFormat)
End Sub

''' <summary>   Handles the Clear Know State button click event. </summary>
Private Sub CLSButton_Click()
    Me.RestartStopWatch
    Me.Session.ClearExecutionState
    Me.ReadStopWatch
End Sub


Friend Sub ConnectedCheckBox_Click()
    
    Dim affirmative As Boolean
    
    If m_session Is Nothing Then
        affirmative = False
    Else
        affirmative = m_session.Connected
    End If
    
    ConnectedCheckBox.value = affirmative
    
    ' report the connection state
    Range(m_connectedCell).value = affirmative
        
    Me.QueryIdentityButton.Enabled = affirmative
    Me.RSTButton.Enabled = affirmative
    Me.SDCButton.Enabled = affirmative
    Me.CLSButton.Enabled = affirmative
    Me.QueryButton.Enabled = affirmative
    Me.ReadButton.Enabled = affirmative
    Me.ReadStatusByteButton.Enabled = affirmative
    Me.WriteButton.Enabled = affirmative
    
End Sub

''' <summary>   Queries the device Identity string. </summary>
Private Sub QueryIdentityButton_Click()
    Me.RestartStopWatch
    Range(m_identityCell).value = Me.Session.QueryIdentity
    Me.ReadStopWatch
End Sub

''' <summary>   Handles the Reset Known State button click event. </summary>
Private Sub RSTButton_Click()
    Me.RestartStopWatch
    Me.Session.ResetKnownState
    Me.ReadStopWatch
End Sub

''' <summary>  Handles the Selective device clear button click event. </summary>
Private Sub SDCButton_Click()
    Me.RestartStopWatch
    Me.Session.SelectiveDeviceClear
    Me.ReadStopWatch
End Sub

''' <summary>   Handles the Toggle connection button click event. </summary>
Private Sub ToggleConnectionButton_Click()

    Me.RestartStopWatch

    Dim Host As String
    Dim Port As Integer
    Dim timeout As Integer
    
    timeout = Me.ReceiveTimeout
    Host = Me.Host
    Port = Me.Port
    
    If Me.Session.Connected Then
        
        Me.Session.CloseConnection
        
        Range(m_connectedHostCell).value = ""
        Range(m_connectedPortCell).value = ""
        Range(m_connectedSocketIdCell).value = ""
        
    Else
    
        Range(m_connectedHostCell).value = Host
        Range(m_connectedPortCell).value = Port
        
        Me.Session.OpenConnection Host, Port, timeout
         
        ' report the socket number
        Range(m_connectedSocketIdCell).value = Me.Session.VISession.SocketId
    End If
    
    Me.ReadStopWatch
    Me.ConnectedCheckBox_Click
    
End Sub

''' <summary>   Handles the Write button click event. </summary>
Private Sub WriteButton_Click()
    Range(m_sentMessageCell).value = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    Me.Session.WriteLine msg
    Me.ReadStopWatch
    Range(m_sentMessageCell).value = msg
    If Me.AutoStatusReadCheckBox.value Then
        Me.ReadStatusByteButton_Click
    End If
End Sub

''' <summary>   Handles the Read button click event. </summary>
Private Sub ReadButton_Click()
    Range(m_receivedMessageCell).value = ""
    Me.RestartStopWatch
    Dim msg As String: msg = Me.Session.Read
    Me.ReadStopWatch
    Range(m_receivedMessageLengthCell).value = msg
    Range(m_receivedMessageCell).value = Len(msg)
End Sub

''' <summary>   Handles the Read Status Byte button click event. </summary>
Friend Sub ReadStatusByteButton_Click()
    Me.RestartStopWatch
    Range(m_statusByteCell).value = Me.Session.ReadStatusByte
    Me.ReadStopWatch
End Sub

''' <summary>   Handles the Query button click event. </summary>
Private Sub QueryButton_Click()
    Me.SentMessage = ""
    Dim msg As String: msg = Me.CommandsComboBox.value
    Me.RestartStopWatch
    Me.ReceivedMessage = Me.Session.QueryLine(msg)
    Me.ReadStopWatch
    Me.SentMessage = msg
End Sub

''' <summary>   Gets the host dotted IP Address. </summary>
Friend Property Get Host() As String
    Host = Range(m_inputHostCell).value
End Property

''' <summary>   Gets the port number. </summary>
Friend Property Get Port() As Integer
    Port = CInt(Range(m_inputPortCell).value)
End Property

''' <summary> Get the receive timeout. </summary>
Friend Property Get ReceiveTimeout() As Integer
    ReceiveTimeout = CInt(Range(m_receiveTimeoutCell).value)
End Property

''' <summary>   Gets the received message. </summary>
Friend Property Get ReceivedMessage() As String
    ReceivedMessage = Range(m_receivedMessageCell).value
End Property

''' <summary>   Sets the received message. </summary>
Friend Property Let ReceivedMessage(ByVal value As String)
    Range(m_receivedMessageCell).value = value
    Range(m_receivedMessageLengthCell).value = Len(value)
End Property

''' <summary>   Gerts the sent message. </summary>
Friend Property Get SentMessage() As String
    SentMessage = Range(m_sentMessageCell).value
End Property

''' <summary>   Sets the sent message. </summary>
Friend Property Let SentMessage(ByVal value As String)
    Range(m_sentMessageCell).value = value
End Property

''' <summary>   Handles the session <see cref="cc_isr_VI.ConnectionChanged"/> event. </summary>
''' <param name="eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_session_ConnectionChanged(ByRef eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

End Sub


