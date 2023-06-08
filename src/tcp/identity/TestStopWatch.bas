Attribute VB_Name = "TestStopWatch"

''' <summary> Tests stopwatch class for elapsed time. </summary>
''' <returns> An instance of the <see cref="Assert"/> class. </returns>
Public Function TestElapsedTime() As Assert
    Dim sw As StopWatch
    Set sw = New StopWatch
    Dim sleepTime As Long
    sleepTime = 100
    Sleep sleepTime + 50
    sw.StopCounter
    Set TestElapsedTime = Assert.IsTrue(sw.ElapsedMilliseconds > sleepTime, _
        "elapsed time " & CStr(sw.ElapsedMilliseconds) & " must exceed sleep time " & CStr(sleepTime))
End Function


