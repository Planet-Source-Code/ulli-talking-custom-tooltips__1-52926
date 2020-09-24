VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Custom Tooltip Class
'''''''''''''''''''''
'This class was inspired by code by Eidos (found at PSC some time ago) and others.
'
'This should be called once on Form Load to create the tooltip window for each
'control. The text uses the vertical bar | as line break character. Tooltip headline,
'individual fontface and fontsize, individual back- and forecolors, an assortment of
'Icons to be displayed in the tooltip, individual hover- and popup-times, and two ways
'of positioning the tooltip window complete the options you have. The .Create function
'returns the hWnd of the created tooltip window or zero if unsuccessful.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Feb30 2004   UMG
'
'Added font modification.
'Limited tooltip box width to one third screen width; this cleared the line break quirk
'with no title text.
'
'Combo Box Quirk: It seems that a combobox consists of several windows and that only
'some of them are implicitly subclassed (TTF_SUBCLASS), so this class will only display
'the tooltip when the cursor is over one of them - the dropdown button in the example.
'Whether this is by desing or by negligence on Micro$oft's part I don't know. The only
'workaround I can think of is to use the same mechanism that is used for windowless
'controls (eg labels - see the example ---> Label1_MouseMove)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun29 2003   UMG
'
'Added tooltip for hWnd-less controls.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jan02 2003   UMG
'
'Three new options have been added - display tooltip always / only if parent form is active / never
'see TTStyle.
'
'Added missing Style private property variable.
'Rearranged code a little.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC           As Long = -4
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private TThWnd                      As Long 'Tooltip window handle
Attribute TThWnd.VB_VarDescription = "Tooltip window handle."
Private TThDC                       As Long 'Tooltip devive context
Private TThFont                     As Long 'Tooltip font
Private Const ToolTipWindowClassName As String = "Tooltips_Class32"
Attribute ToolTipWindowClassName.VB_VarDescription = "Window Style."
Private Enum TTS
    CW_USEDEFAULT = &H80000000
    TTS_STANDARD = 0
    TTS_BALLOON = &H40
    TTS_ALWAYSTIP = 1 'display even if parent window is inactive
    TTS_NOPREFIX = 2 'does not remove "&" from text
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum
Private Const TTF_CENTERTIP         As Long = 2     'center tool on parent
Private Const TTF_SUBCLASS          As Long = &H10  'use implicit subclassing

Public Enum TTStyle
    TTSpeak = 128
    TTStandardIfActive = TTS_STANDARD                   'suppress if parent form is not active
    TTBalloonIfActive = TTS_BALLOON                     'suppress if parent form is not active
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP    'display even if parent form is not active
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP      'display even if parent form is not active
    TTNone = -1                                         'kill tooltip (this is simply treated as illegal)
End Enum
#If False Then
Private TTSpeak, TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone 'to preserve the case
#End If

Public Enum TTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
    'all have a light gray shadow so be careful when selecting the ToolTip BackColor
End Enum
#If False Then 'preserve capitalization
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError
#End If

'default font
Private Const defFontName           As String = "Tahoma"
Private Const defFontSize           As Long = 8

'my properties
Private myStyle                     As TTStyle
Private myVoice                     As TTStyle
Private myIcon                      As TTIcon
Private myForeColor                 As Long
Private myBackColor                 As Long
Private myTitle                     As String 'has the current title
Attribute myTitle.VB_VarDescription = "Private Property Variable."
Private myHoverTime                 As Long 'time im millisecs (-1 = use default)
Private myPopupTime                 As Long 'time im millisecs (-1 = use default)
Private myInitialText               As Variant 'has the initial text
Private myText                      As String
Private myInitialTitle              As Variant 'has the initial title
Private myFontName                  As String
Private myFontSize                  As Long
Private myFontBold                  As Boolean
Private myFontItalic                As Boolean
Private PrevProcPointer             As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum SetWinPos
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOACTIVATE = &H10
    SWP_FLAGS = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    TOPMOST = -1
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As MsgType, ByVal wParam As Long, lParam As Any) As Long
Private Enum MsgType
    WM_USER = &H400
    TTM_SETDELAYTIME = WM_USER + 3
    TTM_ADDTOOL = WM_USER + 4
    TTM_GETTEXT = WM_USER + 11
    TTM_SETTIPBKCOLOR = WM_USER + 19
    TTM_SETTIPTEXTCOLOR = WM_USER + 20
    TTM_SETMAXTIPWIDTH = WM_USER + 24
    TTM_SETTITLE = WM_USER + 32
    WM_SETFONT = &H30
End Enum

Private Type RECTANGLE
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type tToolInfo
    ttSize      As Long
    myFlags     As Long
    ttParhWnd   As Long
    ttId        As Long
    ParentRect  As RECTANGLE
    hInstance   As Long
    pText       As String
    lParam      As Long
End Type
Private ToolInfo                    As tToolInfo

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECTANGLE) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharset As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Const fwDontCare            As Long = 0
Private Const fwBold                As Long = 700

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nMultiplier As Long, ByVal nDivisor As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const DEFAULT_CHARSET       As Long = 1
Private Const LOGPIXELSY            As Long = 90

Public Property Get BackCol() As Long
Attribute BackCol.VB_Description = "Returns the current tooltip backcolor."

  'this returns the current tooltip backcolor

    BackCol = myBackColor

End Property

Public Property Get Centered() As Boolean
Attribute Centered.VB_Description = "Returns the current tooltip alignment."

  'this returns the current tooltip alignment

    Centered = CBool(ToolInfo.myFlags And TTF_CENTERTIP)

End Property

Private Sub Class_Initialize()

    InitCommonControls 'doesn't matter that this is called for every class instance
    myStyle = TTNone

End Sub

Private Sub Class_Terminate()

  'tidy up: kill Font, hDC and tooltip window if they exist

    KillFont
    If TThWnd Then
        SetWindowLong TThWnd, GWL_WNDPROC, PrevProcPointer 'unhook window
        ReleaseDC TThWnd, TThDC
        TThDC = 0
        DestroyWindow TThWnd
        TThWnd = 0
    End If
    myStyle = TTNone

End Sub

Public Function Create(Parent As Control, _
                       Text As String, _
                       Optional ByVal Style As TTStyle = TTBalloonAlways, _
                       Optional ByVal Centered As Boolean = False, _
                       Optional ByVal Icon As TTIcon = TTIconNone, _
                       Optional Title As String = "", _
                       Optional ByVal ForeColor As Long = vbButtonText, _
                       Optional ByVal BackColor As Long = vbInfoBackground, _
                       Optional ByVal HoverTime As Long = -1, _
                       Optional ByVal PopupTime As Long = -1) As Long

  'Create the tooltip window for parent control that has an hWnd
  'This can now also create custom tooltips for hWnd-less controls,
  'just supply a fake hWnd (normally the containing form.hWnd) for windowless controls

    Class_Terminate 'kill font and tooltip window if one exists
    myVoice = Style And TTSpeak
    Style = Style And Not TTSpeak
    With ToolInfo
        On Error Resume Next
            .ttParhWnd = Parent.hWnd 'the control's hWnd
            If Err Then 'has no hWnd
                Err.Clear
                .ttParhWnd = Parent.Parent.hWnd
            End If
            If (Err = 0) And _
               (Style = TTBalloonAlways Or Style = TTStandardAlways Or Style = TTBalloonIfActive Or Style = TTStandardIfActive) And _
               (Icon = TTIconError Or Icon = TTIconInfo Or Icon = TTIconNone Or Icon = TTIconWarning) Then
                'the tooltip parent control has an hWnd and the params are acceptable
                .ttSize = Len(ToolInfo)
                .myFlags = TTF_SUBCLASS Or IIf(Centered, TTF_CENTERTIP, 0&)
                GetClientRect .ttParhWnd, .ParentRect
                .hInstance = App.hInstance
                myTitle = Title
                If myInitialTitle = Empty Then
                    myInitialTitle = myTitle
                End If
                myText = Replace$(Text, "|", vbCrLf) 'the vertical bar is used as line break character
                .pText = myText
                If myInitialText = Empty Then
                    myInitialText = myText
                End If
                If ForeColor < 0 Then
                    ForeColor = GetSysColor(ForeColor And &H7FFFFFFF)
                End If
                If BackColor < 0 Then
                    BackColor = GetSysColor(BackColor And &H7FFFFFFF)
                End If
                If ForeColor = BackColor Then
                    ForeColor = vbButtonText
                    BackColor = vbInfoBackground
                End If
                myForeColor = ForeColor
                myBackColor = BackColor
                myStyle = Style
                myIcon = Icon
                myHoverTime = HoverTime
                If myHoverTime >= 0 And myHoverTime < 50 Then 'at least 50 millisecs
                    myHoverTime = 50
                End If
                myPopupTime = PopupTime
                'create tooltip window and set it's properties
                TThWnd = CreateWindowEx(0&, ToolTipWindowClassName, vbNullString, TTS_NOPREFIX Or Style, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, .ttParhWnd, 0&, .hInstance, 0&)
                TThDC = GetDC(TThWnd)
                SetWindowPos TThWnd, TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS
                PrevProcPointer = SetWindowLong(TThWnd, GWL_WNDPROC, Far(AddressOf WinProc))
                SaveDetails TThWnd, Right$("00000000" & Hex$(PrevProcPointer), 8) & IIf(myVoice, myText, "")
                With Screen
                    SendMessage TThWnd, TTM_SETMAXTIPWIDTH, 0, ByVal .Width / .TwipsPerPixelX / 3
                End With 'SCREEN
                SendMessage TThWnd, TTM_ADDTOOL, 0&, ToolInfo
                SendMessage TThWnd, TTM_SETTITLE, Icon, ByVal myTitle
                SendMessage TThWnd, TTM_SETTIPTEXTCOLOR, myForeColor, ByVal 0&
                SendMessage TThWnd, TTM_SETTIPBKCOLOR, myBackColor, ByVal 0&
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal myHoverTime
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal myPopupTime
                SubstituteFont 'to fill myFontName, myFontSize, myFontBold, and myFontItalic
                Create = TThWnd 'return the tooltip window handle
            End If
        On Error GoTo 0
    End With 'TOOLINFO

End Function

Private Function Far(ByVal Pointer As Long) As Long

  'little helper to be able to use 'AddressOf' in a Class

    Far = Pointer

End Function

Public Property Get FontBold() As Boolean

    FontBold = myFontBold

End Property

Public Property Get FontItalic() As Boolean

    FontItalic = myFontItalic

End Property

Public Property Get FontName() As String

    FontName = myFontName

End Property

Public Property Get FontSize() As Long

    FontSize = myFontSize

End Property

Public Property Get ForeCol() As Long
Attribute ForeCol.VB_Description = "Returns the current tooltip forecolor."

  'this returns the current tooltip forecolor

    ForeCol = myForeColor

End Property

Public Property Get HoverTime() As Long

  'this returns the current mouse hover time time in millicecs (-1 for default)

    HoverTime = myHoverTime

End Property

Public Property Get Icon() As TTIcon
Attribute Icon.VB_Description = "Returns the current tooltip icon."

  'this returns the current tooltip icon

    Icon = myIcon

End Property

Public Property Get InitialText() As String
Attribute InitialText.VB_Description = "Returns the inital tooltip text."

  'this returns the inital tooltip text, ie the one that was supplied on creation

    InitialText = myInitialText

End Property

Public Property Get InitialTitle() As String
Attribute InitialTitle.VB_Description = "Returns the inital tooltip title."

  'this returns the inital tooltip title, ie the one that was supplied on creation

    InitialTitle = myInitialTitle

End Property

Private Sub KillFont()

    If TThFont Then
        DeleteObject TThFont
    End If

End Sub

Public Property Get PopupTime() As Long

  'this returns the current max PopupTime time in millisecs (-1 for default)

    PopupTime = myPopupTime

End Property

Public Property Get Style() As TTStyle
Attribute Style.VB_Description = "Returns the current tooltip style."

  'this returns the current tooltip style

    Style = myStyle Or myVoice

End Property

Public Function SubstituteFont(Optional FontName As String = defFontName, _
                               Optional ByVal FontSize As Long = defFontSize, _
                               Optional ByVal Bold As Boolean = False, _
                               Optional ByVal Italic As Boolean = False) As Long

  'modify tooltip font

    If TThWnd Then 'we have a tooltip window
        Select Case FontSize 'limit fontsize to reasonable values
          Case Is < 8
            FontSize = 8
          Case Is > 36
            FontSize = 36
        End Select
        myFontName = Trim$(FontName)
        myFontSize = FontSize
        myFontBold = Bold
        myFontItalic = Italic
        KillFont 'kill any previous font and create a new one
        TThFont = CreateFont(-MulDiv(myFontSize, GetDeviceCaps(TThDC, LOGPIXELSY), 72), 0, 0, 0, IIf(myFontBold, fwBold, fwDontCare), myFontItalic, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, myFontName)
        If TThFont Then                                 'font was successfully created...
            SendMessage TThWnd, WM_SETFONT, TThFont, 0  'so send it to the tooltip window...
            SubstituteFont = TThFont                    'and return font handle
        End If
    End If

End Function

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns the current tooltip text."

  'this returns the current tooltip text

    Text = myText

End Property

Public Property Get Title() As String
Attribute Title.VB_Description = "Returns the current tooltip title."

  'this returns the current tooltip title

    Title = myTitle

End Property

':) Ulli's VB Code Formatter V2.16.15 (2004-Apr-06 16:09) 164 + 265 = 429 Lines
