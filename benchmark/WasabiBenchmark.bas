Attribute VB_Name = "WasabiBenchmark"
' ============================================================================
' WasabiBenchmark  –  Version based: Wasabi v2.2.0-vNext
'
' Modernised CPU benchmark for Wasabi's internal operations.
'
' Usage:
'   1. Import this module into the same VBA project as Wasabi.bas.
'   2. Ensure the 7 Public wrapper functions are present in Wasabi.bas
'      (listed at the end of this module).
'   3. Run:  WasabiBenchmark_RunAll
'   4. Results appear in the Immediate Window and in a worksheet named
'      "Wasabi Benchmark Results".
' ============================================================================

Option Explicit

' --------------------------------------------------------------------------
' QueryPerformanceCounter  (high-resolution timer)
' --------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private m_Freq As Currency

Private Sub InitQPC()
    If m_Freq = 0 Then QueryPerformanceFrequency m_Freq
End Sub

Private Function QPCms() As Double
    Dim t As Currency
    QueryPerformanceCounter t
    QPCms = CDbl(t) / CDbl(m_Freq) * 1000#
End Function

' --------------------------------------------------------------------------
' Adaptive iteration counts
' --------------------------------------------------------------------------
Private Function ItersFor(ByVal payloadBytes As Long, ByVal isSlow As Boolean) As Long
    ' Slow operations (SHA1, Deflate): fewer iterations
    ' Fast operations (UTF8, Base64): more iterations for accurate measurement
    Dim base As Long
    If isSlow Then
        Select Case True
            Case payloadBytes <= 64:    base = 5000
            Case payloadBytes <= 1024:  base = 1000
            Case payloadBytes <= 16384: base = 200
            Case Else:                  base = 50
        End Select
    Else
        Select Case True
            Case payloadBytes <= 64:    base = 50000
            Case payloadBytes <= 1024:  base = 10000
            Case payloadBytes <= 16384: base = 2000
            Case Else:                  base = 500
        End Select
    End If
    ItersFor = base
End Function

' --------------------------------------------------------------------------
' Helpers
' --------------------------------------------------------------------------
Private Function MakePayload(ByVal size As Long) As Byte()
    Dim buf() As Byte
    Dim i As Long
    ReDim buf(0 To size - 1)
    For i = 0 To size - 1
        buf(i) = CByte(32 + (i Mod 95))
    Next i
    MakePayload = buf
End Function

Private Function MakeString(ByVal size As Long) As String
    Dim s As String
    Dim i As Long
    s = String(size, " ")
    For i = 1 To size
        Mid(s, i, 1) = Chr(32 + ((i - 1) Mod 95))
    Next i
    MakeString = s
End Function

Private Function FormatSize(ByVal b As Long) As String
    If b >= 1048576 Then
        FormatSize = Format(b / 1048576#, "0.#") & " MB"
    ElseIf b >= 1024 Then
        FormatSize = Format(b / 1024#, "0.#") & " KB"
    Else
        FormatSize = b & " B"
    End If
End Function

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    PadR = Left(s & String(n, " "), n)
End Function

Private Function PadL(ByVal s As String, ByVal n As Long) As String
    PadL = Right(String(n, " ") & s, n)
End Function

Private Sub RecordResult(ByRef results As Collection, _
                          ByVal opName As String, _
                          ByVal payloadBytes As Long, _
                          ByVal iters As Long, _
                          ByVal totalMs As Double)
    Dim avgUs     As Double
    Dim throughMB As Double
    Dim opsPerSec As Double
    Dim row(0 To 5) As Variant

    If totalMs > 0 Then
        avgUs = (totalMs / iters) * 1000#
        throughMB = ((CDbl(payloadBytes) * iters) / (totalMs / 1000#)) / 1048576#
        opsPerSec = iters / (totalMs / 1000#)
    End If

    row(0) = opName
    row(1) = payloadBytes
    row(2) = iters
    row(3) = Format(avgUs, "0.000")
    row(4) = Format(throughMB, "0.00")
    row(5) = Format(opsPerSec / 1000#, "0.00")
    results.Add row

    Debug.Print "  " & PadR(opName, 20) & _
                " | " & PadL(FormatSize(payloadBytes), 7) & _
                " | avg " & Format(avgUs, "0.000") & " µs" & _
                " | " & Format(throughMB, "0.00") & " MB/s" & _
                " | " & Format(opsPerSec / 1000#, "0.00") & "k ops/s" & _
                "  [n=" & iters & "]"
End Sub

' --------------------------------------------------------------------------
' Entry points
' --------------------------------------------------------------------------
Public Sub WasabiBenchmark_RunAll()
    Dim results As Collection
    Set results = New Collection

    InitQPC

    Debug.Print "========================================================"
    Debug.Print " Wasabi v2.2.0-vNext  -  Benchmark"
    Debug.Print " " & now() & "  (QPC, freq=" & Format(CDbl(m_Freq) / 1000000#, "0.0") & " MHz)"
    Debug.Print "========================================================"
    Debug.Print ""

    BenchSHA1 results
    BenchBase64 results
    BenchStringToUtf8 results
    BenchUtf8ToString results
    BenchBuildWSFrame results

    WriteResultsToSheet results

    Debug.Print "========================================================"
    Debug.Print " Sheet 'Wasabi Benchmark Results' updated."
    Debug.Print "========================================================"
End Sub

' --------------------------------------------------------------------------
' SHA-1
' --------------------------------------------------------------------------
Private Sub BenchSHA1(ByRef results As Collection)
    Dim sizes(0 To 3) As Long
    sizes(0) = 64: sizes(1) = 1024: sizes(2) = 16384: sizes(3) = 131072

    Dim si As Long, i As Long, iters As Long
    Dim buf() As Byte, dummy() As Byte
    Dim t0 As Double, t1 As Double

    Debug.Print "[SHA1]"
    For si = 0 To 3
        buf = MakePayload(sizes(si))
        iters = ItersFor(sizes(si), True)

        ' warmup (not measured)
        For i = 1 To 20: dummy = SHA1_Public(buf): Next i

        DoEvents
        t0 = QPCms()
        For i = 1 To iters
            dummy = SHA1_Public(buf)
        Next i
        t1 = QPCms()

        RecordResult results, "SHA1", sizes(si), iters, t1 - t0
        DoEvents
    Next si
End Sub

' --------------------------------------------------------------------------
' Base64Encode
' --------------------------------------------------------------------------
Private Sub BenchBase64(ByRef results As Collection)
    Dim sizes(0 To 3) As Long
    sizes(0) = 64: sizes(1) = 1024: sizes(2) = 16384: sizes(3) = 131072

    Dim si As Long, i As Long, iters As Long
    Dim buf() As Byte, dummy As String
    Dim t0 As Double, t1 As Double

    Debug.Print "[Base64Encode]"
    For si = 0 To 3
        buf = MakePayload(sizes(si))
        iters = ItersFor(sizes(si), False)

        For i = 1 To 50: dummy = Base64Encode_Public(buf): Next i

        DoEvents
        t0 = QPCms()
        For i = 1 To iters
            dummy = Base64Encode_Public(buf)
        Next i
        t1 = QPCms()

        RecordResult results, "Base64Encode", sizes(si), iters, t1 - t0
        DoEvents
    Next si
End Sub

' --------------------------------------------------------------------------
' StringToUtf8
' --------------------------------------------------------------------------
Private Sub BenchStringToUtf8(ByRef results As Collection)
    Dim sizes(0 To 3) As Long
    sizes(0) = 64: sizes(1) = 1024: sizes(2) = 16384: sizes(3) = 131072

    Dim si As Long, i As Long, iters As Long
    Dim src As String, dummy() As Byte
    Dim t0 As Double, t1 As Double

    Debug.Print "[StringToUtf8]"
    For si = 0 To 3
        src = MakeString(sizes(si))
        iters = ItersFor(sizes(si), False)

        For i = 1 To 50: dummy = StringToUtf8_Public(src): Next i

        DoEvents
        t0 = QPCms()
        For i = 1 To iters
            dummy = StringToUtf8_Public(src)
        Next i
        t1 = QPCms()

        RecordResult results, "StringToUtf8", sizes(si), iters, t1 - t0
        DoEvents
    Next si
End Sub

' --------------------------------------------------------------------------
' Utf8ToString
' --------------------------------------------------------------------------
Private Sub BenchUtf8ToString(ByRef results As Collection)
    Dim sizes(0 To 3) As Long
    sizes(0) = 64: sizes(1) = 1024: sizes(2) = 16384: sizes(3) = 131072

    Dim si As Long, i As Long, iters As Long
    Dim buf() As Byte, dummy As String
    Dim t0 As Double, t1 As Double

    Debug.Print "[Utf8ToString]"
    For si = 0 To 3
        buf = MakePayload(sizes(si))
        iters = ItersFor(sizes(si), False)

        For i = 1 To 50: dummy = Utf8ToString_Public(buf, sizes(si)): Next i

        DoEvents
        t0 = QPCms()
        For i = 1 To iters
            dummy = Utf8ToString_Public(buf, sizes(si))
        Next i
        t1 = QPCms()

        RecordResult results, "Utf8ToString", sizes(si), iters, t1 - t0
        DoEvents
    Next si
End Sub

' --------------------------------------------------------------------------
' BuildWSFrame (includes CryptGenRandom mask)
' --------------------------------------------------------------------------
Private Sub BenchBuildWSFrame(ByRef results As Collection)
    Dim sizes(0 To 3) As Long
    sizes(0) = 64: sizes(1) = 1024: sizes(2) = 16384: sizes(3) = 131072

    Dim si As Long, i As Long, iters As Long
    Dim buf() As Byte, dummy() As Byte
    Dim t0 As Double, t1 As Double

    Debug.Print "[BuildWSFrame]"
    For si = 0 To 3
        buf = MakePayload(sizes(si))
        iters = ItersFor(sizes(si), False)

        For i = 1 To 50: dummy = BuildWSFrame_Public(buf, sizes(si), 1, True): Next i

        DoEvents
        t0 = QPCms()
        For i = 1 To iters
            dummy = BuildWSFrame_Public(buf, sizes(si), 1, True)
        Next i
        t1 = QPCms()

        RecordResult results, "BuildWSFrame", sizes(si), iters, t1 - t0
        DoEvents
    Next si
End Sub

' --------------------------------------------------------------------------
' Export to worksheet
' --------------------------------------------------------------------------
Private Sub WriteResultsToSheet(ByRef results As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Wasabi Benchmark Results" Then
            Application.DisplayAlerts = False: ws.Delete: Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Wasabi Benchmark Results"
    WriteHeader ws
    WriteRows ws, results, 5
    ws.Columns("A:F").AutoFit
    ws.Activate
End Sub

Private Sub AppendResultsToSheet(ByRef results As Collection)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Wasabi Benchmark Results")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Wasabi Benchmark Results"
        WriteHeader ws
        WriteRows ws, results, 5
    Else
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        WriteRows ws, results, lastRow
    End If
    ws.Columns("A:F").AutoFit
    ws.Activate
End Sub

Private Sub WriteHeader(ByRef ws As Worksheet)
    ws.Cells(1, 1).Value = "Wasabi v2.2.0-vNext  -  Benchmark Results"
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.size = 13
    ws.Cells(2, 1).Value = "Date: " & now() & "  |  Timer: QueryPerformanceCounter"
    ws.Cells(2, 1).Font.Color = RGB(100, 100, 100)

    Dim r As Long: r = 4
    ws.Cells(r, 1).Value = "Operation"
    ws.Cells(r, 2).Value = "Payload (bytes)"
    ws.Cells(r, 3).Value = "Iterations"
    ws.Cells(r, 4).Value = "Avg Latency (µs)"
    ws.Cells(r, 5).Value = "Throughput (MB/s)"
    ws.Cells(r, 6).Value = "Ops/s (k)"

    With ws.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(13, 148, 136)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub WriteRows(ByRef ws As Worksheet, ByRef results As Collection, ByVal startRow As Long)
    Dim r As Long: r = startRow
    Dim lastOp As String: lastOp = ""
    Dim toggle As Boolean: toggle = False
    Dim item As Variant, row As Variant

    For Each item In results
        row = item
        If row(0) <> lastOp Then
            lastOp = row(0)
            toggle = Not toggle
        End If

        ws.Cells(r, 1).Value = row(0)
        ws.Cells(r, 2).Value = CLng(row(1))
        ws.Cells(r, 3).Value = CLng(row(2))
        ws.Cells(r, 4).Value = CDbl(row(3))
        ws.Cells(r, 5).Value = CDbl(row(4))
        ws.Cells(r, 6).Value = CDbl(row(5))

        ws.Range("A" & r & ":F" & r).Interior.Color = _
            IIf(toggle, RGB(225, 243, 240), RGB(248, 248, 248))
        r = r + 1
    Next item

    ' Summary
    r = r + 2
    ws.Cells(r, 1).Value = "Summary — best throughput per operation"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "Operation": ws.Cells(r, 2).Value = "Best MB/s": ws.Cells(r, 3).Value = "Payload"
    With ws.Range("A" & r & ":C" & r)
        .Font.Bold = True: .Interior.Color = RGB(55, 138, 221): .Font.Color = RGB(255, 255, 255)
    End With
    r = r + 1

    Dim opBest As Object, opPayload As Object
    Set opBest = CreateObject("Scripting.Dictionary")
    Set opPayload = CreateObject("Scripting.Dictionary")

    Dim item2 As Variant, row2 As Variant
    For Each item2 In results
        row2 = item2
        Dim mb As Double: mb = CDbl(row2(4))
        If Not opBest.Exists(row2(0)) Then
            opBest(row2(0)) = mb: opPayload(row2(0)) = CLng(row2(1))
        ElseIf mb > opBest(row2(0)) Then
            opBest(row2(0)) = mb: opPayload(row2(0)) = CLng(row2(1))
        End If
    Next item2

    Dim k As Variant
    For Each k In opBest.Keys
        ws.Cells(r, 1).Value = k
        ws.Cells(r, 2).Value = opBest(k)
        ws.Cells(r, 3).Value = FormatSize(opPayload(k))
        r = r + 1
    Next k
End Sub

' ==========================================================================
' REQUIRED PUBLIC WRAPPERS – add these to Wasabi.bas
' ==========================================================================
'
' Public Function SHA1_Public(ByRef data() As Byte) As Byte()
'     SHA1_Public = SHA1(data)
' End Function
'
' Public Function Base64Encode_Public(ByRef Bytes() As Byte) As String
'     Base64Encode_Public = Base64Encode(Bytes)
' End Function
'
' Public Function StringToUtf8_Public(ByVal str As String) As Byte()
'     StringToUtf8_Public = StringToUtf8(str)
' End Function
'
' Public Function Utf8ToString_Public(ByRef utf8() As Byte, ByVal length As Long) As String
'     Utf8ToString_Public = Utf8ToString(utf8, length)
' End Function
'
' Public Function BuildWSFrame_Public(ByRef payload() As Byte, ByVal payloadLen As Long, _
'                                     ByVal opcode As Byte, ByVal isFinal As Boolean) As Byte()
'     BuildWSFrame_Public = BuildWSFrame(payload, payloadLen, opcode, isFinal)
' End Function
'
' Public Function DeflatePayload_Public(ByVal handle As Long, ByRef data() As Byte, _
'                                       ByVal dataLen As Long, ByRef outLen As Long) As Byte()
'     DeflatePayload_Public = DeflatePayload(handle, data, dataLen, outLen)
' End Function
'
' Public Function InflatePayload_Public(ByVal handle As Long, ByRef data() As Byte, _
'                                       ByVal dataLen As Long, ByRef outLen As Long) As Byte()
'     InflatePayload_Public = InflatePayload(handle, data, dataLen, outLen)
' End Function

