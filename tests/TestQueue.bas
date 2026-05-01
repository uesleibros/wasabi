Attribute VB_Name = "TestQueue"
Option Explicit

Private g_Handle As Long

Public Sub Queue_SendReceive()
    If Not SetupConnection() Then Exit Sub
    Dim msg As String
    msg = "TestMessage"
    If Not WebSocketSend(msg, g_Handle) Then TestFail "send failed"
    Dim received As String
    received = WaitForResponse(g_Handle, 5000)
    If received <> msg Then TestFail "expected '" & msg & "', got '" & received & "'"
    CleanupConnection
End Sub

Public Sub Queue_Peek()
    If Not SetupConnection() Then Exit Sub
    Dim msg As String
    msg = "PeekTest"
    If Not WebSocketSend(msg, g_Handle) Then TestFail "send failed"
    Dim startTick As Long
    startTick = GetTickCount()
    Do While WebSocketGetPendingCount(g_Handle) = 0
        If GetTickCount() - startTick > 5000 Then TestFail "timeout waiting for message": Exit Sub
        DoEvents
    Loop
    If WebSocketPeek(g_Handle) <> msg Then TestFail "peek mismatch"
    If WebSocketGetPendingCount(g_Handle) <> 1 Then TestFail "expected 1 pending after peek"
    CleanupConnection
End Sub

Public Sub Queue_Flush()
    If Not SetupConnection() Then Exit Sub
    WebSocketSend "Flush1", g_Handle
    WebSocketSend "Flush2", g_Handle
    Dim startTick As Long
    startTick = GetTickCount()
    Do While WebSocketGetPendingCount(g_Handle) < 2
        If GetTickCount() - startTick > 5000 Then TestFail "timeout waiting for messages": Exit Sub
        DoEvents
    Loop
    WebSocketFlushQueue g_Handle
    If WebSocketGetPendingCount(g_Handle) <> 0 Then TestFail "queue not flushed"
    If WebSocketGetBinaryPendingCount(g_Handle) <> 0 Then TestFail "binary queue not flushed"
    CleanupConnection
End Sub

Private Function SetupConnection() As Boolean
    If Not WebSocketConnect("wss://echo.websocket.org", g_Handle) Then
        TestFail "could not connect to echo server"
        Exit Function
    End If
    SetupConnection = True
End Function

Private Sub CleanupConnection()
    WebSocketDisconnect g_Handle
End Sub

Private Function WaitForResponse(ByVal h As Long, ByVal timeoutMs As Long) As String
    Dim startTick As Long
    startTick = GetTickCount()
    Do
        Dim msg As String
        msg = WebSocketReceive(h)
        If msg <> "" Then WaitForResponse = msg: Exit Function
        If GetTickCount() - startTick > timeoutMs Then Exit Function
        DoEvents
    Loop
End Function
