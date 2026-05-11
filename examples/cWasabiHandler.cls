VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWasabiHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/**
' * @description Custom handler class for asynchronous WebSocket events.
' * Implements the required interface for WasabiUseAsync.
' */

Public Sub OnConnect(ByVal handle As Long)
    Debug.Print "[Async] Handle " & handle & " connected successfully."
End Sub

Public Sub OnReceive(ByVal handle As Long)
    Dim msg As String
    ' Retrieve messages from the internal queue
    Do
        msg = Wasabi.WebSocketReceiveText(handle)
        If msg <> "" Then
            Debug.Print "[Async Receive] " & msg
        End If
    Loop While msg <> ""
End Sub

Public Sub OnReadyToSend(ByVal handle As Long)
    ' Called when the socket is ready for more data
End Sub

Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
    Debug.Print "[Async Error] Code: " & errorCode & " on event " & eventType
End Sub

Public Sub OnClose(ByVal handle As Long)
    Debug.Print "[Async] Connection closed for handle " & handle
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    Debug.Print "[Async] Handler detached from handle " & handle
End Sub
