VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1
Private pauseRequested As Boolean

Sub ResetClock()
    Range("B4") = TimeValue("00:00:15")
End Sub

Sub StartTimer()
  pauseRequested = False
  If m_eventTimer Is Nothing Then
    Set m_eventTimer = cc_isr_Core.Constructor.CreateEventTimer(100)
    m_eventTimer.StartTimer
  Else
    m_eventTimer.ResumeTimer
  End If
End Sub

Sub PauseTimer()
  pauseRequested = True
End Sub

Sub AbortTimer()
  ' we need to figure out how to recreate the timer
  m_eventTimer.StopTimer
End Sub

Private Sub m_eventTimer_Tick()
    On Error Resume Next
    If Range("B4").value <= 0 Or pauseRequested Then
        m_eventTimer.PauseTimer
    Else
        Range("B4") = Range("B4") - TimeValue("00:00:01")
    End If
    On Error GoTo 0
End Sub

    
