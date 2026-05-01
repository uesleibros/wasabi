Attribute VB_Name = "TestRunner"
Option Explicit

Private m_Passed As Long
Private m_Failed As Long
Private m_CurrentTest As String

Public Sub TestFail(ByVal msg As String)
    Debug.Print "[FAIL] " & m_CurrentTest & ": " & msg
    m_Failed = m_Failed + 1
End Sub

Private Sub TestPass()
    Debug.Print "[PASS] " & m_CurrentTest
    m_Passed = m_Passed + 1
End Sub

Private Sub RunTest(ByVal testName As String, ByVal testProc As String)
    m_CurrentTest = testName
    On Error GoTo ErrHandler
    Application.Run testProc
    TestPass
    Exit Sub
ErrHandler:
    TestFail "Error: " & Err.Description & " (" & Err.Number & ")"
End Sub

Public Sub RunAllTests()
    m_Passed = 0
    m_Failed = 0
    Debug.Print String(60, "=")
    Debug.Print "Wasabi Test Suite"
    Debug.Print String(60, "=")

    ' Utils
    RunTest "Base64Encode_Empty", "TestUtils.Base64Encode_Empty"
    RunTest "Base64Encode_Known", "TestUtils.Base64Encode_Known"
    RunTest "SHA1_Known", "TestUtils.SHA1_Known"
    RunTest "ParseURL_WSS", "TestUtils.ParseURL_WSS"
    RunTest "ParseURL_WS", "TestUtils.ParseURL_WS"
    RunTest "StringToUtf8_RoundTrip", "TestUtils.StringToUtf8_RoundTrip"

    ' Handshake
    RunTest "GenerateWSKey_Length", "TestHandshake.GenerateWSKey_Length"
    RunTest "ComputeAccept_Known", "TestHandshake.ComputeAccept_Known"

    ' Frame Builder
    RunTest "BuildWSFrame_SmallPayload", "TestFrameBuilder.BuildWSFrame_SmallPayload"
    RunTest "BuildWSFrame_MediumPayload", "TestFrameBuilder.BuildWSFrame_MediumPayload"
    RunTest "BuildWSFrame_LargePayload", "TestFrameBuilder.BuildWSFrame_LargePayload"
    RunTest "BuildWSFrame_Ping", "TestFrameBuilder.BuildWSFrame_Ping"
    RunTest "BuildWSFrame_Close", "TestFrameBuilder.BuildWSFrame_Close"
    RunTest "BuildWSFrame_Binary", "TestFrameBuilder.BuildWSFrame_Binary"
    RunTest "BuildWSFrame_FINBit", "TestFrameBuilder.BuildWSFrame_FINBit"

    ' Queue (integration tests)
    RunTest "Queue_SendReceive", "TestQueue.Queue_SendReceive"
    RunTest "Queue_Peek", "TestQueue.Queue_Peek"
    RunTest "Queue_Flush", "TestQueue.Queue_Flush"

    Debug.Print String(60, "=")
    Debug.Print "Passed: " & m_Passed & "  Failed: " & m_Failed
    Debug.Print String(60, "=")
End Sub
