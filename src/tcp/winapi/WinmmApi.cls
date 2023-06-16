VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinmmApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary>   retrieves the system time, in milliseconds. The system time is the time elapsed since Windows was started. </summary>
''' <returns> A Long system time in milliseconds</returns>
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

