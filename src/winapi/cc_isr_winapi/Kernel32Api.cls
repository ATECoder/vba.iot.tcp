VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kernel32Api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary>   Copies the contents of a source memory block to a destination memory block,
''' and supports overlapping source and destination memory blocks. </summary>
''' <param name="a_dest">     [out] A pointer to the destination memory block to copy the bytes to. </param>
''' <param name="a_source">   [in] A pointer to the source memory block to copy the bytes from. </param>
''' <param name="a_length">   [in] The number of bytes to copy from the source to the destination. </param>
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef a_dest As Any, ByRef a_source As Any, ByVal a_length As Long)

''' <summary>   Suspends the thread for the specified duration in milliseconds. </summary>
''' <param name="a_sleeptimeMs">   Sleep time in milliseconds. </param>
Private Declare PtrSafe Sub SleepMilliseconds Lib "kernel32" Alias "Sleep" (ByVal a_sleeptimeMs As Long)

''' <summary>   Copies a memory block form the source to the destination of length bytes. </summary>
''' <param name="a_source">        [Double] The start of the source block. </param>
''' <param name="a_destination">   [Double] The start of the destination block. </param>
''' <param name="a_length">        [Long] The number of bytes which to copy. </param>
Public Sub CopyMemory(ByRef a_source() As Double, ByRef a_destination() As Double, ByVal a_length As Long)
    RtlMoveMemory a_destination(LBound(a_source)), a_destination(LBound(a_source)), a_length
End Sub

''' <summary>   Suspends the thread for the specified duration in milliseconds. </summary>
''' <param name="a_duration">   [Long] Sleep time in milliseconds. </param>
Public Sub Sleep(ByVal a_duration As Long)
    SleepMilliseconds a_duration
End Sub
