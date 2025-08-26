Attribute VB_Name = "Timer"
Option Explicit

Private Type Timer
    interval As Long
    procedure As String
    times As Long
    enabled As Boolean
    initialized As Boolean
    ticks As Long
End Type

Private Timer As Timer

Public Sub StartTimer(ByRef interval As Long, _
                      ByRef procedure As String, _
                      Optional ByRef times As Long)
With Timer
    .interval = interval
    .procedure = procedure
    .times = times
    .enabled = True
    .initialized = False
    .ticks = 0
End With

InvokeTimer
End Sub

Public Sub StopTimer()
Timer.enabled = False
End Sub

Private Function InvokeTimer()
If Timer.ticks > Timer.times And Not Timer.times = 0 Then StopTimer
If Not Timer.enabled Then Exit Function

If Timer.initialized Then
    ActiveSheet.Evaluate "0+" & Timer.procedure
Else
    Timer.initialized = True
End If

Timer.ticks = Timer.ticks + 1

Application.OnTime Now + 1 / 86400 * Timer.interval, "InvokeTimer"
End Function

