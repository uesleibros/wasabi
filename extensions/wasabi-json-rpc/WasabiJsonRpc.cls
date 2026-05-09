VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WasabiJsonRpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ============================================================================
' WasabiJsonRpc.cls
' JSON-RPC 2.0 Protocol Handler for Wasabi
'
' Implements the Wasabi Protocol Handler interface and exposes a clean API
' for sending requests, notifications, and batch calls over WebSocket.
'
' Interaction modes
'   Polling   - call Receive() or ReceiveBatch() from your own loop
'   Callbacks - set a callback module name via SetCallbackModule(); the
'               handler will call RpcOnResult, RpcOnError, and RpcOnNotify
'               on that module whenever a message arrives.
'
' Usage
'   Dim rpc As New WasabiJsonRpc
'   rpc.SetCallbackModule "MyModule"          ' optional
'
'   Dim h As Long
'   WebSocketConnect "wss://api.example.com/rpc", h
'   WasabiUseProtocol rpc, h
'
'   Dim id As Long
'   id = rpc.Request("subtract", "{""a"":10,""b"":3}", h)
'
'   ' Polling style
'   Dim result As String
'   result = rpc.Receive(id, timeoutMs:=3000, handle:=h)
'
' Callback module interface (optional, implement in a standard .bas module)
'   Public Sub RpcOnResult(ByVal id As Long, ByVal result As String, ByVal handle As Long)
'   Public Sub RpcOnError(ByVal id As Long, ByVal code As Long, ByVal message As String, ByVal handle As Long)
'   Public Sub RpcOnNotify(ByVal method As String, ByVal params As String, ByVal handle As Long)
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Constants
' ----------------------------------------------------------------------------

Private Const MAX_QUEUE     As Long = 256
Private Const POLL_INTERVAL As Long = 5     ' ms between DoEvents loops

' Response types stored in queue
Private Const RESP_RESULT   As Byte = 1
Private Const RESP_ERROR    As Byte = 2
Private Const RESP_NOTIFY   As Byte = 3

' ----------------------------------------------------------------------------
' Internal types
' ----------------------------------------------------------------------------

Private Type RpcResponse
    RespType    As Byte     ' RESP_RESULT / RESP_ERROR / RESP_NOTIFY
    Id          As Long     ' 0 for notifications
    Result      As String   ' raw JSON result value (for RESP_RESULT)
    ErrCode     As Long     ' JSON-RPC error code (for RESP_ERROR)
    ErrMessage  As String   ' JSON-RPC error message (for RESP_ERROR)
    Method      As String   ' method name (for RESP_NOTIFY)
    Params      As String   ' raw JSON params (for RESP_NOTIFY)
    Consumed    As Boolean
End Type

' ----------------------------------------------------------------------------
' State
' ----------------------------------------------------------------------------

Private m_Handle         As Long
Private m_NextId         As Long
Private m_CallbackModule As String

Private m_Queue(0 To MAX_QUEUE - 1) As RpcResponse
Private m_QueueHead      As Long
Private m_QueueTail      As Long
Private m_QueueCount     As Long

' ----------------------------------------------------------------------------
' Wasabi Protocol Handler interface
' ----------------------------------------------------------------------------

Public Sub OnConnect(ByVal handle As Long)
    m_Handle  = handle
    m_NextId  = 1
    m_QueueHead  = 0
    m_QueueTail  = 0
    m_QueueCount = 0
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    m_Handle = -1
End Sub

Public Sub OnTextMessage(ByVal handle As Long, ByVal message As String)
    ParseAndEnqueue handle, message
End Sub

Public Sub OnBinaryMessage(ByVal handle As Long, ByRef data() As Byte)
    ' JSON-RPC is text-based; binary frames are ignored.
End Sub

' ----------------------------------------------------------------------------
' Public API - configuration
' ----------------------------------------------------------------------------

' SetCallbackModule sets the name of a standard VBA module that Wasabi will
' call with RpcOnResult, RpcOnError, and RpcOnNotify when messages arrive.
' Pass an empty string to disable callbacks and use polling only.
Public Sub SetCallbackModule(ByVal moduleName As String)
    m_CallbackModule = moduleName
End Sub

' ----------------------------------------------------------------------------
' Public API - sending
' ----------------------------------------------------------------------------

' Request sends a JSON-RPC 2.0 request and returns the auto-assigned id.
' params must be a valid JSON value (object, array, or null).
Public Function Request(ByVal method As String, _
                        Optional ByVal params As String = "null", _
                        Optional ByVal handle As Long = -1) As Long
    Dim h   As Long
    Dim id  As Long
    Dim msg As String

    h = ResolveHandle(handle)

    id = m_NextId
    m_NextId = m_NextId + 1

    msg = BuildRequest(method, params, id)
    WebSocketSend msg, h

    Request = id
End Function

' Notify sends a JSON-RPC 2.0 notification (no id, no response expected).
Public Sub Notify(ByVal method As String, _
                  Optional ByVal params As String = "null", _
                  Optional ByVal handle As Long = -1)
    Dim h   As Long
    Dim msg As String

    h = ResolveHandle(handle)
    msg = BuildNotification(method, params)
    WebSocketSend msg, h
End Sub

' Batch sends multiple requests in a single WebSocket frame.
' methods and params must be parallel arrays of equal length.
' Returns an array of assigned ids in the same order.
Public Function Batch(ByRef methods() As String, _
                      ByRef params() As String, _
                      Optional ByVal handle As Long = -1) As Long()
    Dim h       As Long
    Dim i       As Long
    Dim n       As Long
    Dim ids()   As Long
    Dim parts() As String
    Dim msg     As String

    h = ResolveHandle(handle)
    n = UBound(methods) - LBound(methods) + 1

    ReDim ids(0 To n - 1)
    ReDim parts(0 To n - 1)

    Dim base As Long
    base = LBound(methods)

    For i = 0 To n - 1
        ids(i)   = m_NextId
        m_NextId = m_NextId + 1
        parts(i) = BuildRequest(methods(base + i), params(base + i), ids(i))
    Next i

    msg = "[" & Join(parts, ",") & "]"
    WebSocketSend msg, h

    Batch = ids
End Function

' ----------------------------------------------------------------------------
' Public API - receiving (polling)
' ----------------------------------------------------------------------------

' Receive blocks up to timeoutMs waiting for the response to a specific id.
' Returns the raw JSON result string on success.
' Returns an empty string on timeout or if the server returned an error
' (check ReceiveError separately).
Public Function Receive(ByVal id As Long, _
                        Optional ByVal timeoutMs As Long = 5000, _
                        Optional ByVal handle As Long = -1) As String
    Dim h         As Long
    Dim startTick As Long
    Dim i         As Long

    h = ResolveHandle(handle)
    startTick = GetTickCount()

    Do
        ' Pump incoming frames from the socket.
        WebSocketReceive h

        ' Scan the queue for a matching result.
        For i = 0 To m_QueueCount - 1
            Dim idx As Long
            idx = (m_QueueHead + i) Mod MAX_QUEUE
            With m_Queue(idx)
                If Not .Consumed And .Id = id Then
                    If .RespType = RESP_RESULT Then
                        .Consumed = True
                        Receive = .Result
                        Exit Function
                    ElseIf .RespType = RESP_ERROR Then
                        .Consumed = True
                        Exit Function   ' caller checks ReceiveError
                    End If
                End If
            End With
        Next i

        If TickElapsed(startTick) >= timeoutMs Then Exit Do
        Sleep POLL_INTERVAL
        DoEvents
    Loop

    Receive = ""   ' timeout
End Function

' ReceiveError returns the last error that arrived for the given id.
' Call after Receive returns empty string to distinguish timeout from RPC error.
Public Function ReceiveError(ByVal id As Long, _
                             ByRef outCode As Long, _
                             ByRef outMessage As String) As Boolean
    Dim i   As Long
    Dim idx As Long

    For i = 0 To m_QueueCount - 1
        idx = (m_QueueHead + i) Mod MAX_QUEUE
        With m_Queue(idx)
            If .Id = id And .RespType = RESP_ERROR Then
                outCode    = .ErrCode
                outMessage = .ErrMessage
                ReceiveError = True
                Exit Function
            End If
        End With
    Next i
End Function

' ReceivePending returns True and populates the out parameters if any
' unread result or error is in the queue (for non-blocking polling loops).
Public Function ReceivePending(ByRef outId As Long, _
                               ByRef outResult As String, _
                               ByRef outErrCode As Long, _
                               ByRef outErrMsg As String) As Boolean
    Dim i   As Long
    Dim idx As Long

    For i = 0 To m_QueueCount - 1
        idx = (m_QueueHead + i) Mod MAX_QUEUE
        With m_Queue(idx)
            If Not .Consumed And (.RespType = RESP_RESULT Or .RespType = RESP_ERROR) Then
                outId = .Id
                .Consumed = True
                If .RespType = RESP_RESULT Then
                    outResult = .Result
                Else
                    outErrCode = .ErrCode
                    outErrMsg  = .ErrMessage
                End If
                ReceivePending = True
                Exit Function
            End If
        End With
    Next i
End Function

' PendingCount returns the number of unread responses in the queue.
Public Function PendingCount() As Long
    Dim i     As Long
    Dim count As Long

    For i = 0 To m_QueueCount - 1
        Dim idx As Long
        idx = (m_QueueHead + i) Mod MAX_QUEUE
        If Not m_Queue(idx).Consumed Then count = count + 1
    Next i

    PendingCount = count
End Function

' ----------------------------------------------------------------------------
' Private - message parsing
' ----------------------------------------------------------------------------

Private Sub ParseAndEnqueue(ByVal handle As Long, ByVal json As String)
    ' Detect batch responses (array at root).
    Dim trimmed As String
    trimmed = Trim(json)

    If Left(trimmed, 1) = "[" Then
        ParseBatchResponse handle, trimmed
    Else
        ParseSingleResponse handle, trimmed
    End If
End Sub

Private Sub ParseSingleResponse(ByVal handle As Long, ByVal json As String)
    Dim resp As RpcResponse

    If InStr(json, """result""") > 0 Then
        resp.RespType = RESP_RESULT
        resp.Id       = JsonExtractLong(json, "id")
        resp.Result   = JsonExtractRaw(json, "result")
        Enqueue resp
        FireCallback handle, resp

    ElseIf InStr(json, """error""") > 0 And InStr(json, """id""") > 0 Then
        resp.RespType  = RESP_ERROR
        resp.Id        = JsonExtractLong(json, "id")
        Dim errObj As String
        errObj         = JsonExtractRaw(json, "error")
        resp.ErrCode   = JsonExtractLong(errObj, "code")
        resp.ErrMessage = JsonExtractString(errObj, "message")
        Enqueue resp
        FireCallback handle, resp

    ElseIf InStr(json, """method""") > 0 And InStr(json, """id""") = 0 Then
        ' Notification (no id field).
        resp.RespType = RESP_NOTIFY
        resp.Method   = JsonExtractString(json, "method")
        resp.Params   = JsonExtractRaw(json, "params")
        Enqueue resp
        FireCallback handle, resp
    End If
End Sub

Private Sub ParseBatchResponse(ByVal handle As Long, ByVal json As String)
    ' Split the top-level array into individual objects and parse each one.
    ' This is a simple depth-aware splitter; it handles nested objects/arrays.
    Dim i      As Long
    Dim depth  As Long
    Dim inStr  As Boolean
    Dim ch     As String
    Dim start  As Long
    Dim item   As String

    depth = 0
    inStr = False
    start = 2   ' skip the leading [

    For i = 2 To Len(json) - 1
        ch = Mid(json, i, 1)

        If inStr Then
            If ch = """" And Mid(json, i - 1, 1) <> "\" Then inStr = False
        Else
            Select Case ch
                Case """"   : inStr = True
                Case "{", "[" : depth = depth + 1
                Case "}", "]" : depth = depth - 1
                Case ","
                    If depth = 0 Then
                        item = Trim(Mid(json, start, i - start))
                        If Len(item) > 0 Then ParseSingleResponse handle, item
                        start = i + 1
                    End If
            End Select
        End If
    Next i

    ' Last element before the closing ]
    item = Trim(Mid(json, start, Len(json) - start))
    If Len(item) > 0 Then ParseSingleResponse handle, item
End Sub

Private Sub Enqueue(ByRef resp As RpcResponse)
    If m_QueueCount >= MAX_QUEUE Then
        ' Drop the oldest entry to make room.
        m_QueueHead = (m_QueueHead + 1) Mod MAX_QUEUE
        m_QueueCount = m_QueueCount - 1
    End If

    m_Queue(m_QueueTail) = resp
    m_QueueTail  = (m_QueueTail + 1) Mod MAX_QUEUE
    m_QueueCount = m_QueueCount + 1
End Sub

Private Sub FireCallback(ByVal handle As Long, ByRef resp As RpcResponse)
    If m_CallbackModule = "" Then Exit Sub

    On Error Resume Next

    Select Case resp.RespType
        Case RESP_RESULT
            CallByName Nothing, m_CallbackModule & ".RpcOnResult", VbMethod, resp.Id, resp.Result, handle
            ' CallByName does not support module-level calls directly;
            ' use Application.Run for standard modules:
            Application.Run m_CallbackModule & ".RpcOnResult", resp.Id, resp.Result, handle

        Case RESP_ERROR
            Application.Run m_CallbackModule & ".RpcOnError", resp.Id, resp.ErrCode, resp.ErrMessage, handle

        Case RESP_NOTIFY
            Application.Run m_CallbackModule & ".RpcOnNotify", resp.Method, resp.Params, handle
    End Select

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Private - JSON builders
' ----------------------------------------------------------------------------

Private Function BuildRequest(ByVal method As String, _
                              ByVal params As String, _
                              ByVal id As Long) As String
    BuildRequest = "{""jsonrpc"":""2.0""," & _
                   """method"":"  & JsonString(method) & "," & _
                   """params"":"  & params & "," & _
                   """id"":"      & CStr(id) & "}"
End Function

Private Function BuildNotification(ByVal method As String, _
                                   ByVal params As String) As String
    BuildNotification = "{""jsonrpc"":""2.0""," & _
                        """method"":"  & JsonString(method) & "," & _
                        """params"":"  & params & "}"
End Function

' ----------------------------------------------------------------------------
' Private - minimal JSON helpers
' These cover the subset of JSON-RPC 2.0 responses that Wasabi needs to read.
' For production use with deeply nested structures, swap these out for a full
' JSON parser.
' ----------------------------------------------------------------------------

' JsonString returns a JSON-encoded string literal, including surrounding quotes.
Private Function JsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonString = """" & s & """"
End Function

' JsonExtractString returns the string value of a top-level key.
Private Function JsonExtractString(ByVal json As String, ByVal key As String) As String
    Dim pattern As String
    Dim pos     As Long
    Dim endPos  As Long

    pattern = """" & key & """:"
    pos = InStr(json, pattern)
    If pos = 0 Then Exit Function

    pos = pos + Len(pattern)

    ' Skip whitespace.
    Do While Mid(json, pos, 1) = " " : pos = pos + 1 : Loop

    If Mid(json, pos, 1) <> """" Then Exit Function
    pos = pos + 1   ' skip opening quote

    Dim result As String
    Dim ch     As String

    Do While pos <= Len(json)
        ch = Mid(json, pos, 1)
        If ch = "\" Then
            pos = pos + 1
            Select Case Mid(json, pos, 1)
                Case """" : result = result & """"
                Case "\"  : result = result & "\"
                Case "n"  : result = result & vbLf
                Case "r"  : result = result & vbCr
                Case "t"  : result = result & vbTab
                Case Else : result = result & Mid(json, pos, 1)
            End Select
        ElseIf ch = """" Then
            Exit Do
        Else
            result = result & ch
        End If
        pos = pos + 1
    Loop

    JsonExtractString = result
End Function

' JsonExtractLong returns the numeric value of a top-level key.
Private Function JsonExtractLong(ByVal json As String, ByVal key As String) As Long
    Dim pattern As String
    Dim pos     As Long
    Dim endPos  As Long
    Dim numStr  As String

    pattern = """" & key & """:"
    pos = InStr(json, pattern)
    If pos = 0 Then Exit Function

    pos = pos + Len(pattern)
    Do While Mid(json, pos, 1) = " " : pos = pos + 1 : Loop

    endPos = pos
    Do While endPos <= Len(json)
        Dim c As String
        c = Mid(json, endPos, 1)
        If c >= "0" And c <= "9" Or c = "-" Then
            endPos = endPos + 1
        Else
            Exit Do
        End If
    Loop

    numStr = Mid(json, pos, endPos - pos)
    If Len(numStr) > 0 Then JsonExtractLong = CLng(numStr)
End Function

' JsonExtractRaw returns the raw JSON value (object, array, string, or scalar)
' for a top-level key, preserving its original form.
Private Function JsonExtractRaw(ByVal json As String, ByVal key As String) As String
    Dim pattern As String
    Dim pos     As Long

    pattern = """" & key & """:"
    pos = InStr(json, pattern)
    If pos = 0 Then Exit Function

    pos = pos + Len(pattern)
    Do While Mid(json, pos, 1) = " " : pos = pos + 1 : Loop

    Dim startPos As Long
    Dim depth    As Long
    Dim inStr    As Boolean
    Dim ch       As String

    startPos = pos
    depth    = 0
    inStr    = False

    Do While pos <= Len(json)
        ch = Mid(json, pos, 1)

        If inStr Then
            If ch = """" And Mid(json, pos - 1, 1) <> "\" Then inStr = False
        Else
            Select Case ch
                Case """"        : inStr = True
                Case "{", "["   : depth = depth + 1
                Case "}", "]"
                    depth = depth - 1
                    If depth <= 0 And startPos <> pos Then
                        JsonExtractRaw = Mid(json, startPos, pos - startPos + 1)
                        Exit Function
                    End If
                Case ",", "}"
                    If depth = 0 Then
                        JsonExtractRaw = Trim(Mid(json, startPos, pos - startPos))
                        Exit Function
                    End If
            End Select
        End If

        pos = pos + 1
    Loop

    JsonExtractRaw = Trim(Mid(json, startPos))
End Function

' ----------------------------------------------------------------------------
' Private - utilities
' ----------------------------------------------------------------------------

Private Function ResolveHandle(ByVal handle As Long) As Long
    If handle = -1 Then
        ResolveHandle = WebSocketGetDefaultHandle()
    Else
        ResolveHandle = handle
    End If
End Function

Private Function GetTickCount() As Long
    ' Delegates to the Wasabi module-level GetTickCount via kernel32.
    ' Because Wasabi.bas declares it privately, we use a VBA-compatible
    ' alternative: Timer returns seconds since midnight with ms precision.
    GetTickCount = CLng(Timer * 1000) And &H7FFFFFFF
End Function

Private Function TickElapsed(ByVal startTick As Long) As Long
    Dim now As Long
    now = GetTickCount()
    If now >= startTick Then
        TickElapsed = now - startTick
    Else
        TickElapsed = (&H7FFFFFFF - startTick) + now + 1   ' handle wrap-around
    End If
End Function

Private Sub Sleep(ByVal ms As Long)
    Dim t As Long
    t = GetTickCount()
    Do While TickElapsed(t) < ms : DoEvents : Loop
End Sub
