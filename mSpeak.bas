Attribute VB_Name = "mSpeak"
Option Explicit

'based on code by ricky_g

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private ToolTipDetails              As New Collection
Private ComputerVoice               As New SpVoice

Public Sub SaveDetails(ByVal hWnd As Long, Text As String)

  'save previous procedure pointer and text to speak

    On Error Resume Next
        ToolTipDetails.Remove CStr(hWnd) 'remove first
    On Error GoTo 0
    ToolTipDetails.Add Text, CStr(hWnd) 'and then add

End Sub

Public Function WinProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  'tooltip window procedure

  Dim Temp As String

    Temp = ToolTipDetails(CStr(hWnd)) 'get saved details
    Select Case nMsg
      Case 133 'whatever msg 133 is - it is received once on tt-popup
        ComputerVoice.Speak Mid$(Temp, 9), SVSFlagsAsync 'speak text asynch
      Case 24 'whatever msg 24 is - it is received once on tt-close
        ComputerVoice.Speak vbNullString, SVSFPurgeBeforeSpeak 'mute current text
    End Select
    WinProc = CallWindowProc(Val("&H" & Left$(Temp, 8)), hWnd, nMsg, wParam, lParam) 'call previous

End Function

':) Ulli's VB Code Formatter V2.16.15 (2004-Apr-06 16:09) 7 + 30 = 37 Lines
