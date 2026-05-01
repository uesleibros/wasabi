Attribute VB_Name = "TestHandshake"
Option Explicit

Public Sub GenerateWSKey_Length()
    Dim key As String
    key = GenerateWSKey()
    Dim decoded() As Byte
    decoded = Base64DecodeBytes(StrConv(key, vbFromUnicode))
    Dim lenDecoded As Long
    lenDecoded = UBound(decoded) - LBound(decoded) + 1
    If lenDecoded <> 16 Then TestFail "expected 16 bytes after Base64 decode, got " & lenDecoded
End Sub

Public Sub ComputeAccept_Known()
    Dim key As String
    key = "dGhlIHNhbXBsZSBub25jZQ=="
    Dim expected As String
    expected = "s3pPLMBiTxaQ9kYGzzhZRbK+xOo="
    Dim accept As String
    accept = ComputeWebSocketAccept(key)
    If accept <> expected Then TestFail "expected '" & expected & "', got '" & accept & "'"
End Sub
