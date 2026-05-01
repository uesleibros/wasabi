Attribute VB_Name = "TestFrameBuilder"
Option Explicit

Private Function FrameHasFinBit(ByRef frame() As Byte) As Boolean
    FrameHasFinBit = ((frame(0) And &H80) <> 0)
End Function

Private Function FrameOpcode(ByRef frame() As Byte) As Byte
    FrameOpcode = frame(0) And &HF
End Function

Private Function FramePayloadLen(ByRef frame() As Byte) As Long
    Dim len7 As Long
    len7 = frame(1) And &H7F
    If len7 < 126 Then
        FramePayloadLen = len7
    ElseIf len7 = 126 Then
        FramePayloadLen = CLng(frame(2)) * 256 + CLng(frame(3))
    Else
        FramePayloadLen = 0
        Dim i As Long
        For i = 2 To 9
            FramePayloadLen = FramePayloadLen * 256 + CLng(frame(i))
        Next i
    End If
End Function

Public Sub BuildWSFrame_SmallPayload()
    Dim payload(0 To 4) As Byte
    payload(0) = Asc("H"): payload(1) = Asc("e"): payload(2) = Asc("l")
    payload(3) = Asc("l"): payload(4) = Asc("o")
    Dim frame() As Byte
    frame = BuildWSFrame(payload, 5, WS_OPCODE_TEXT, True)
    If Not FrameHasFinBit(frame) Then TestFail "FIN bit not set"
    If FrameOpcode(frame) <> WS_OPCODE_TEXT Then TestFail "opcode is not text"
    If FramePayloadLen(frame) <> 5 Then TestFail "expected payload length 5, got " & FramePayloadLen(frame)
End Sub

Public Sub BuildWSFrame_MediumPayload()
    Dim payload(0 To 299) As Byte
    Dim i As Long
    For i = 0 To 299: payload(i) = CByte(i Mod 256): Next i
    Dim frame() As Byte
    frame = BuildWSFrame(payload, 300, WS_OPCODE_BINARY, True)
    If FrameOpcode(frame) <> WS_OPCODE_BINARY Then TestFail "opcode is not binary"
    If FramePayloadLen(frame) <> 300 Then TestFail "expected payload length 300, got " & FramePayloadLen(frame)
End Sub

Public Sub BuildWSFrame_LargePayload()
    Dim payload(0 To 65535) As Byte
    Dim i As Long
    For i = 0 To 65535: payload(i) = CByte(i Mod 256): Next i
    Dim frame() As Byte
    frame = BuildWSFrame(payload, 65536, WS_OPCODE_TEXT, True)
    If FramePayloadLen(frame) <> 65536 Then TestFail "expected payload length 65536, got " & FramePayloadLen(frame)
End Sub

Public Sub BuildWSFrame_Ping()
    Dim frame() As Byte
    frame = BuildWSFrame(NullByteArray(), 0, WS_OPCODE_PING, True)
    If FrameOpcode(frame) <> WS_OPCODE_PING Then TestFail "opcode is not ping"
End Sub

Public Sub BuildWSFrame_Close()
    Dim frame() As Byte
    frame = BuildWSFrame(NullByteArray(), 0, WS_OPCODE_CLOSE, True)
    If FrameOpcode(frame) <> WS_OPCODE_CLOSE Then TestFail "opcode is not close"
End Sub

Public Sub BuildWSFrame_Binary()
    Dim payload(0 To 9) As Byte
    Dim frame() As Byte
    frame = BuildWSFrame(payload, 10, WS_OPCODE_BINARY, True)
    If FrameOpcode(frame) <> WS_OPCODE_BINARY Then TestFail "opcode is not binary"
End Sub

Public Sub BuildWSFrame_FINBit()
    Dim payload(0 To 4) As Byte
    Dim frame() As Byte
    frame = BuildWSFrame(payload, 5, WS_OPCODE_TEXT, False)
    If FrameHasFinBit(frame) Then TestFail "FIN bit should not be set for non-final frame"
    frame = BuildWSFrame(payload, 5, WS_OPCODE_TEXT, True)
    If Not FrameHasFinBit(frame) Then TestFail "FIN bit should be set for final frame"
End Sub

Private Function NullByteArray() As Byte()
    Dim b() As Byte
    NullByteArray = b
End Function
