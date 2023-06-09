VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary> </summary>
''' <remarks> </remarks>
''' <param name=""> </para>
''' <returns> <returns>
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (stamp As Currency) As Byte

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (freq As Currency) As Byte
   
''' <summary> Suspends the thread for the spefified duration in milleseconds. </summary>
''' <param name="dwMilliseconds"> Sleep time in milliseconds. </para>
Private Declare PtrSafe Sub SleepMilliseconds Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
   
Private m_overheadCounts As Currency
Private m_counting As Boolean
Private m_startCounts As Currency
Private m_endCounts As Currency
Private m_frequency As Currency

''' <summary> Initializes the stop watch </summary>
''' <remarks> Gets the stopwatch clock frequency. </remarks>
Private Sub Class_Initialize()
    QueryPerformanceFrequency m_frequency
    Restart
End Sub

''' <summary> Restarts the stopwatch counter. </summary>
''' <remarks> Sets the stopwatch start and end counts and
''' marks the stopwatch as counting.
''' </remarks>
Public Sub Restart()
    QueryPerformanceCounter m_startCounts
    QueryPerformanceCounter m_endCounts
    m_overheadCounts = m_endCounts - m_startCounts
    m_counting = True
End Sub

''' <summary> Stopes the stopwatch counter. </summary>
''' <remarks> Fixes the stop watch end time and marks it as not counting.
''' </remarks>
Public Sub StopCounter()
    QueryPerformanceCounter m_endCounts
    m_counting = False
End Sub

''' <summary> Returns the elapsed counts. </summary>
Property Get ElapsedCounts() As Double
    If m_counting Then
        QueryPerformanceCounter m_endCounts
    End If
    ElapsedCounts = m_endCounts - m_startCounts - m_overheadCounts
End Property

''' <summary> Returns the elapsed time in seconds. </summary>
Property Get ElapsedSeconds() As Double
    ElapsedSeconds = ElapsedCounts / m_frequency
End Property

''' <summary> Returns the elapsed time in milliseconds. </summary>
Property Get ElapsedMilliseconds() As Double
    ElapsedMilliseconds = 1000# * ElapsedSeconds
End Property

''' <summary> Suspends the thread for the spefified duration in milleseconds. </summary>
''' <param name="duration"> Sleep time in milliseconds. </para>
Public Sub Sleep(ByVal duration As Long)
    SleepMilliseconds duration
End Sub

