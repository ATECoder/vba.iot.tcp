VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionChangingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_connected As Boolean
Private m_cancel As Boolean

''' <summary>   Initializes the event arguments. </summary>
''' <para name="isConnected">   [Boolean] The connection state. </param>
Public Sub Initialize(ByVal isConnected As Boolean)
    m_connected = isConnected
    m_cancel = False
End Sub

''' <summary>   Gets the connection state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = m_connected
End Property

''' <summary>   Gets the cancel request. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Cancel() As Boolean
    Cancel = m_cancel
End Property

''' <summary>   Sets the cancel request. </summary>
''' <param name="value">   [Boolean] True to cancel the event. </param>
Public Property Let Cancel(ByVal value As Boolean)
    Cancel = m_cancel
End Property


