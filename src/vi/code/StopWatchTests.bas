Attribute VB_Name = "StopWatchTests"
Option Explicit

''' <summary>   Unit test. Asserts that the <see cref="StopWatch"/>.<see cref="StopWatch.ElapedMilliseconds"/>
''' exceeds the thread sleep time. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestElapsedTimeShouldExceedSleepTime() As Assert
    Dim sw As StopWatch
    Set sw = New StopWatch
    Dim sleepTime As Long
    sleepTime = 100
    sw.Sleep sleepTime + 50
    sw.StopCounter
    Set TestElapsedTimeShouldExceedSleepTime = Assert.IsTrue(sw.ElapsedMilliseconds > sleepTime, _
        "elapsed time " & CStr(sw.ElapsedMilliseconds) & " must exceed sleep time " & CStr(sleepTime))
End Function


