VERSION 5.00
Begin VB.Form fDemo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Talking Tooltip Demo"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":000D
      TabIndex        =   9
      Text            =   "Combo1"
      ToolTipText     =   "This is ComboBox1"
      Top             =   2310
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   450
      Index           =   1
      Left            =   195
      TabIndex        =   8
      ToolTipText     =   "This is another Command Button 3"
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2910
      TabIndex        =   7
      Top             =   2325
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   780
      Left            =   2130
      TabIndex        =   4
      ToolTipText     =   "This is Frame Number 1|holding one Option Box  "
      Top             =   1275
      Width           =   1620
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   315
         TabIndex        =   5
         ToolTipText     =   "This is Option 1"
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2130
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "This is Textbox 1 - Enter some text|and hover mouse again"
      Top             =   225
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   450
      Index           =   0
      Left            =   195
      TabIndex        =   2
      ToolTipText     =   "This is Command Button 3"
      Top             =   1605
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   450
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "This is Command Button 2"
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   195
      TabIndex        =   0
      ToolTipText     =   "This is Command Button 1"
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      Height          =   255
      Left            =   2130
      TabIndex        =   6
      ToolTipText     =   "This is Label1 (using the Form-hWnd|to create the Custom ToolTip)|and it is set not to speak"
      Top             =   780
      Width           =   1620
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Tooltips        As New Collection 'keeping references to all tooltip class instances
'Note that this collection is automatically destroyed when this form is
'destroyed causing an avalanche effect such that all class instances are
'destroyed when this collection is destroyed (it having the only reference
'to each class instance), and on Class_Terminate all created tool windows
'are destroyed, so I'm pretty sure there is no memory leak :-)

Private LabelTooltip    As cToolTip
Attribute LabelTooltip.VB_VarHelpID = -1

Private Sub btExit_Click()

    Set Tooltips = Nothing 'only necessary if you want to re-load the form later but doesn't do any harm if not
    Unload Me

End Sub

Private Sub Form_Load()

  'This is called once on Form Load and creates all relevant tooltip windows
  'The original ToolTipText property is used to fill these windows with life.
  'The vertical bar | is used as line break character. Tooltip headline, indi-
  'vidual fontface and fontsize, individual back- and forecolors and an assortment
  'of Icons to be displayed in the tooltip, and individual hover- and popup-times
  'complete the options you have.
  'The .Create function returns the hWnd of the created tooltip window
  'or zero if unsuccessful.

  Dim Control   As Control
  Dim Tooltip   As cToolTip
  Dim TtTxt     As String
  Dim CollKey   As String
  Dim h         As Long

    For Each Control In Controls                    'cycle thru all controls
        h = 0
        With Control
            On Error Resume Next                    'in case the control has no tool tip text property
                TtTxt = Trim$(.ToolTipText)         'try to access the tool tip text
                h = .hWnd                           'try to get the hWnd
            On Error GoTo 0
            If Len(TtTxt) And h <> 0 Then           'got Text and the Control has an hWnd
                CollKey = .Name
                On Error Resume Next                'in case control is not in an array of controls and therefore has no index property
                    CollKey = CollKey & "(" & .Index & ")"
                On Error GoTo 0
                Set Tooltip = New cToolTip
                If Tooltip.Create(Control, TtTxt, TTBalloonAlways Or TTSpeak, (TypeName(Control) = "TextBox"), TTIconInfo, CollKey) Then
                    Tooltips.Add Tooltip, CollKey   'to keep a reference to the current tool tip class instance (prevent it from being destroyed)
                    .ToolTipText = vbNullString     'kill tooltiptext so we don't get two tips
                End If
            End If
        End With 'CONTROL
    Next Control

    'and one indvidual tooltip
    Set Tooltip = New cToolTip 'instantiate
    With Tooltip
        .Create btExit, "Click on this button|to close application", TTBalloonIfActive Or TTSpeak, False, TTIconError, "Exit", vbBlue, RGB(192, 255, 255), 0, 15000  'create tip
        .SubstituteFont "Amaze", 16, Italic:=True 'alter font
    End With 'TOOLTIP
    Tooltips.Add Tooltip, btExit.Name  'don't forget to keep it (we keep it in a collection, but any object variable will do)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'this demonstrates how a tooltip could be altered at runtime
  'uses late binding but who cares? (only happens once during mouse move after the text has changed)

    If Text1.DataChanged Then
        Text1.DataChanged = False
        With Tooltips(Text1.Name)                       'finds reference to Text1 tooltip class instance
            .Create Text1, .InitialText & "||Text has changed:|" & """" & Text1 & """", .Style, .Centered, TTIconWarning, .InitialTitle
        End With 'TOOLTIPS(TEXT1.NAME)
    End If

    'the mouse is outside the windowless control
    If Not LabelTooltip Is Nothing Then                 'and there is a tooltip
        Label1.ToolTipText = LabelTooltip.InitialText   'so restore the tooltiptext to the control
        Set LabelTooltip = Nothing                      'and destroy the class instance
    End If

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'example shows how to create a tooltip for a windowless control
  'when the mouse is over the control

  Dim Tmp As String

    If LabelTooltip Is Nothing Then                     'we have no tooltip yet for this
        Tmp = Label1.ToolTipText                        'so get the control's tooltiptext...
        Label1.ToolTipText = ""                         '...and erase it so we don't get two tips
        Set LabelTooltip = New cToolTip                 'instantiate the class
        LabelTooltip.Create Label1, Tmp, , , _
                            TTIconInfo, _
                            "Trick by somebody @ PSC", _
                            vbRed, _
                            vbYellow, , _
                            30000 'create a tooltip for the containing form instead
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.15 (2004-Apr-06 16:09) 10 + 98 = 108 Lines
