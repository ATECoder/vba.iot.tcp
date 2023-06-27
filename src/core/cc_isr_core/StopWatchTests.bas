Attribute VB_Name = "StopWatchTests"
Option Explicit

''' <summary>   Unit test. Asserts that the <see cref="StopWatch"/>.<see cref="StopWatch.ElapedMilliseconds"/>
''' exceeds the thread sleep time. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestElapsedTimeShouldExceedexpectedMs() As Assert
    
    Dim stopper As StopWatch: Set stopper = New StopWatch
    Dim expectedMs As Long
    expectedMs = 100
    stopper.Sleep expectedMs + 50
    stopper.StopCounter
    Dim actualMs As Long: actualMs = stopper.ElapsedMilliseconds
    Set TestElapsedTimeShouldExceedexpectedMs = Assert.IsTrue(stopper.ElapsedMilliseconds > expectedMs, _
        "elapsed time " & CStr(stopper.ElapsedMilliseconds) & " must exceed sleep time " & CStr(expectedMs))
        
End Function

''' <summary>   Unit test. Asserts that the <see cref="StopWatch"/>.<see cref="StopWatch.Wait"/>
''' exceeds the specified interval. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestTimeShouldExceedexpectedMs() As Assert
    Dim expectedMs As Long: expectedMs = 100
    Dim stopper As StopWatch: Set stopper = New StopWatch
    Dim actualMs As Long: actualMs = stopper.Wait(expectedMs)
    Set TestTimeShouldExceedexpectedMs = Assert.IsTrue(actualMs >= expectedMs, _
        "elapsed time " & CStr(actualMs) & " must exceed " & CStr(expectedMs))
End Function



