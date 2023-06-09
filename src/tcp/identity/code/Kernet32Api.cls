VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kernet32Api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''' <summary> Copies the contents of a source memory block to a destination memory block,
''' and supports overlapping source and destination memory blocks. </summary>
''' <param name="hpvDest">   [out] A pointer to the destination memory block to copy the bytes to.. </para>
''' <param name="hpvSource"> [in] A pointer to the source memory block to copy the bytes from.. </para>
''' <param name="length">     [in] The number of bytes to copy from the source to the destination. </para>
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

''' <summary> Suspends the thread for the spefified duration in milleseconds. </summary>
''' <param name="dwMilliseconds"> Sleep time in milliseconds. </para>
Private Declare PtrSafe Sub SleepMilliseconds Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

''' <summary> Copies a memory block form the source to the destination of length bytes. </summary>
''' <param name="source">      The start of the soruce block. </para>
''' <param name="destination"> The start of the destination block. </para>
''' <param name="length">      The number of bytes which to copy. </para>
Public Sub CopyMemory(ByRef source() As Double, ByRef destination() As Double, ByVal length As Long)
    RtlMoveMemory destination(LBound(source)), destination(LBound(source)), length
End Sub

''' <summary> Suspends the thread for the spefified duration in milleseconds. </summary>
''' <param name="duration"> Sleep time in milliseconds. </para>
Public Sub Sleep(ByVal duration As Long)
    SleepMilliseconds duration
End Sub
