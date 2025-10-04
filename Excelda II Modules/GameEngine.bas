Option Explicit
'Safe timer-driven game engine scaffold.
'Use Application.OnTime to schedule ticks instead of a blocking Sleep/loop.
'This is non-invasive: it doesn't replace the legacy runGame; use StartSafeGameLoop to try a safe, incremental run.

Public g_NextTickTime As Date
Public g_TickIntervalMs As Long
Public g_TickIntervalSeconds As Double
Public g_TickCounter As Long

Sub StartSafeGameLoop()
    On Error Resume Next
    'Default tick interval (ms). We'll try to read from Data!C4 but enforce a safe minimum.
    g_TickIntervalMs = 100
    g_TickIntervalMs = CLng(Sheets("Data").Range("C4").Value)
    If g_TickIntervalMs < 20 Then g_TickIntervalMs = 100 ' don't allow extremely small intervals
    g_TickIntervalSeconds = g_TickIntervalMs / 1000#
    g_TickCounter = 0
    'Initial scheduling
    ScheduleNextTick Now
    MsgBox "Safe game loop started. Tick interval: " & g_TickIntervalMs & " ms"
End Sub

Sub ScheduleNextTick(startTime As Date)
    g_NextTickTime = DateAdd("s", g_TickIntervalSeconds, startTime)
    Application.OnTime EarliestTime:=g_NextTickTime, Procedure:="GameTick", Schedule:=True
End Sub

Sub StopSafeGameLoop()
    On Error Resume Next
    Application.OnTime EarliestTime:=g_NextTickTime, Procedure:="GameTick", Schedule:=False
    MsgBox "Safe game loop stopped. Total ticks: " & g_TickCounter
End Sub

Sub GameTick()
    On Error GoTo TickErr
    g_TickCounter = g_TickCounter + 1

    'Lightweight heartbeat and safe checks only.
    'Do NOT call the legacy runGame here until we've migrated more logic.
    On Error Resume Next
    Sheets("Data").Range("Z1").Value = Now
    Sheets("Data").Range("Z2").Value = g_TickCounter

    'Optionally call legacy per-tick runner when enabled (migration mode)
    If UseLegacyTick = True Then
        On Error Resume Next
        'Call the per-tick method in AA_GameLoop (RunGame_Tick) which contains
        'the migrated, non-blocking portion of the legacy game loop.
        RunGame_Tick
        On Error GoTo TickErr
    End If

    'Schedule next tick
    ScheduleNextTick Now
    Exit Sub

TickErr:
    'Stop scheduling to avoid runaway if errors occur
    StopSafeGameLoop
    MsgBox "GameTick error: " & Err.Number & " - " & Err.Description
End Sub

'Helper to quickly toggle between legacy and safe loop during migration
Sub StartLegacyRunGame_Safely()
    'This is a convenience: DO NOT call runGame directly if it crashes.
    'Use this only after you've verified runGame will behave (and after backing up workbook).
    If MsgBox("Start legacy runGame? This may block Excel. Continue?", vbYesNo) = vbYes Then
        runGame
    End If
End Sub
