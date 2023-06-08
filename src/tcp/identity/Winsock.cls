VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WINSOCK_VERSION = 257
Private Const COMMAND_ERROR = -1
Private Const RECV_ERROR = -1
Private Const NO_ERROR = 0

' true if winsock was initiated
Private m_initiated As Boolean

' true if winsock was disposed
Private m_disposed As Boolean

''' <summary> Initializes winsock. </summary>
'''< remarks> This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/> </remarks>
Private Sub Class_Initialize()

    On Error GoTo ErrorHandler
    
    Dim startUpInfo As WSADATA
    
    ' Version 1.1 (1*256 + 1) = 257
    ' version 2.0 (2*256 + 0) = 512
    
    ' Initialize Winsock DLL
    Dim result As Long
    result = WSAStartup(WINSOCK_VERSION, startUpInfo)
    
    If result <> 0 Then
        MsgBox ("ERROR starting winsock " & CStr(result))
    End If
    m_initiated = (result = 0)

ErrorHandler:
   MsgBox ("ERROR starting winsock class (" & CStr(err.Number) & "): " & err.Description)

End Sub


''' <summary> Returns true if Winsok was initiated. </summary>
Public Property Get Initiated()
    Initiated = m_initiated
End Property

''' <summary> Returns true if Winsok was Disposed. </summary>
Public Property Get Disposed()
    Disposed = m_disposed
End Property

''' <summary> Terminates Winsock. </summary>
Public Sub Dispose()
    On Error Resume Next
    If m_initiated And Not m_disposed Then
        m_disposed = True
        m_initiated = False
        ' Shutdown Winsock DLL
        Dim result As Long
        result = WSACleanup()
        If result <> 0 Then
            MsgBox ("ERROR stopping winsock " & CStr(result))
        End If
    End If
End Sub

''' <summary> Terminates Winsock. </summary>
''' <remarks> Because this class is declared as a singleron (VB_PredeclaredId=True)
''' Terminate is not executed. So dispose needs to be executed. </remarks>
Private Sub Class_Terminate()
    Dispose
End Sub

