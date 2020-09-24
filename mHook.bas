Attribute VB_Name = "mHook"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private ToolTipTexts                As New Collection
Private ComputerVoice               As New SpVoice

Public Function MsgHandler(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim Temp As String

    Temp = ToolTipTexts(CStr(hWnd))
    Select Case nMsg
      Case 133 'whatever msg 133 is - it is received once on popup
        On Error Resume Next
            ComputerVoice.Speak Mid$(Temp, 9), SVSFlagsAsync 'speak new text asynch (this will not speak if the key cannot be found in the collection)
        On Error GoTo 0
      Case 24 'whatever msg 24 is - it is received once on close
        ComputerVoice.Speak "", SVSFPurgeBeforeSpeak 'mute previous text
    End Select
    MsgHandler = CallWindowProc(Val("&H" & Left$(Temp, 8)), hWnd, nMsg, wParam, lParam)

End Function

Public Sub SaveDetails(ByVal hWnd As Long, Text As String)

    On Error Resume Next
        ToolTipTexts.Remove CStr(hWnd)
    On Error GoTo 0
    ToolTipTexts.Add Text, CStr(hWnd)

End Sub

':) Ulli's VB Code Formatter V2.16.15 (2004-Apr-06 15:28) 5 + 28 = 33 Lines
