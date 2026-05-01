Attribute VB_Name = "TestUtils"
Option Explicit

Public Sub Base64Encode_Empty()
    Dim b() As Byte
    Dim result As String
    result = Base64Encode(b)
    If result <> "" Then TestFail "expected empty string, got '" & result & "'"
End Sub

Public Sub Base64Encode_Known()
    Dim b(0 To 2) As Byte
    b(0) = Asc("M"): b(1) = Asc("a"): b(2) = Asc("n")
    Dim result As String
    result = Base64Encode(b)
    If result <> "TWFu" Then TestFail "expected 'TWFu', got '" & result & "'"
End Sub

Public Sub SHA1_Known()
    Dim inputBytes() As Byte
    inputBytes = StrConv("abc", vbFromUnicode)
    Dim hash() As Byte
    hash = SHA1(inputBytes)
    Dim hexStr As String
    Dim i As Long
    For i = 0 To UBound(hash)
        hexStr = hexStr & Right("0" & Hex(hash(i)), 2)
    Next i
    If UCase(hexStr) <> "A9993E364706816ABA3E25717850C26C9CD0D89D" Then
        TestFail "expected A9993E..., got " & hexStr
    End If
End Sub

Public Sub ParseURL_WSS()
    Dim host As String, port As Long, path As String, tls As Boolean
    If Not ParseURL("wss://example.com:8443/chat", host, port, path, tls) Then
        TestFail "ParseURL returned False"
        Exit Sub
    End If
    If host <> "example.com" Then TestFail "expected host 'example.com', got '" & host & "'"
    If port <> 8443 Then TestFail "expected port 8443, got " & port
    If path <> "/chat" Then TestFail "expected path '/chat', got '" & path & "'"
    If Not tls Then TestFail "expected TLS=True"
End Sub

Public Sub ParseURL_WS()
    Dim host As String, port As Long, path As String, tls As Boolean
    If Not ParseURL("ws://localhost/echo", host, port, path, tls) Then
        TestFail "ParseURL returned False"
        Exit Sub
    End If
    If host <> "localhost" Then TestFail "expected host 'localhost', got '" & host & "'"
    If port <> 80 Then TestFail "expected port 80, got " & port
    If path <> "/echo" Then TestFail "expected path '/echo', got '" & path & "'"
    If tls Then TestFail "expected TLS=False"
End Sub

Public Sub StringToUtf8_RoundTrip()
    Dim original As String
    original = "Hello, 世界!"
    Dim utf8() As Byte
    utf8 = StringToUtf8(original)
    Dim result As String
    result = Utf8ToString(utf8, UBound(utf8) - LBound(utf8) + 1)
    If result <> original Then TestFail "expected '" & original & "', got '" & result & "'"
End Sub
