VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScpiSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const thisModuleName As String = "ScpiSystem"

' reference to the IEEE488 session
Private m_session As Ieee488Session


''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the TCP Client socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If m_session Is Nothing Then
        Connected = False
    Else
        Connected = m_session.Connected
    End If

End Property

''' <summary>   Returns a reference to the IEEE488 Session. </summary>
''' <value>   An IEEE488Session. </returns>
Public Property Get Session() As Ieee488Session
    Set Session = m_session
End Property

''' <summary>   Returns a reference to the IEEE488 Session. </summary>
''' <param name="session">    A references to a connected <see cref="Ieee488Session"/> </param>
Public Sub Initialize(ByVal connectedSession As Ieee488Session)
    Set m_session = connectedSession
End Sub

''' <summary>   Issues a beep. </summary>
Public Sub Beep()
    Me.Session.VISession.WriteLine ("SYST:BEEP")
End Sub

''' <summary>   Gets the top error from the error queue. </summary>
''' <value>   A String; the top error on the error queue. </value>
Public Function ErrorDequeue() As Boolean
    ErrorDequeue = Me.Session.VISession.QueryLine(":SYST:ERR?")
End Function

''' <summary>   Clears teh error queue. </summary>
Public Sub ErrorQueueClear()
    Me.Session.VISession.WriteLine (":SYST:CLE")
End Sub

''' <summary>   Query INPUTS switch (0=rear, 1=front). </summary>
''' <value>   A boolean true if the inputs are set to the front pannel. </value>
Public Function FrontSwitch() As Boolean
    FrontSwitch = 1 = Me.Session.VISession.QueryLine(":SYST:FRSW?")
End Function

''' <summary>   Return to :SYST:PRES defaults. </summary>
Public Sub Preset()
    Me.Session.VISession.WriteLine (":SYST:PRES")
End Sub

