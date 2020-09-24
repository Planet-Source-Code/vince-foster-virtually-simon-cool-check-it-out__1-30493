VERSION 5.00
Begin VB.Form frmSimon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmSimon.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSimon.frx":0E42
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "alienheretic@attbi.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frmSimon.frx":22C2C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Image imgLevel 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "frmSimon.frx":2306E
      Top             =   6240
      Width           =   240
   End
   Begin VB.Image imgLevel 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmSimon.frx":235F8
      Top             =   6240
      Width           =   240
   End
   Begin VB.Image imgStart 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "frmSimon.frx":23B82
      Top             =   5940
      Width           =   240
   End
   Begin VB.Image imgStart 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmSimon.frx":2410C
      Top             =   5940
      Width           =   240
   End
   Begin VB.Image imgPower 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "frmSimon.frx":24696
      Top             =   5640
      Width           =   240
   End
   Begin VB.Image imgPower 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmSimon.frx":24C20
      Top             =   5640
      Width           =   240
   End
   Begin VB.Image imgBlue 
      Height          =   2775
      Index           =   1
      Left            =   8340
      Picture         =   "frmSimon.frx":251AA
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgYellow 
      Height          =   2775
      Index           =   1
      Left            =   8340
      Picture         =   "frmSimon.frx":28058
      Top             =   8280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgRed 
      Height          =   2775
      Index           =   1
      Left            =   5580
      Picture         =   "frmSimon.frx":2AE62
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgGreen 
      Height          =   2775
      Index           =   1
      Left            =   5580
      Picture         =   "frmSimon.frx":2D595
      Top             =   8280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgYellow 
      Height          =   2775
      Index           =   0
      Left            =   8340
      Picture         =   "frmSimon.frx":30364
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgGreen 
      Height          =   2775
      Index           =   0
      Left            =   5580
      Picture         =   "frmSimon.frx":330D6
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgBlue 
      Height          =   2775
      Index           =   0
      Left            =   8340
      Picture         =   "frmSimon.frx":35D98
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgRed 
      Height          =   2775
      Index           =   0
      Left            =   5580
      Picture         =   "frmSimon.frx":38B1E
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgLevel 
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "frmSimon.frx":3B0D6
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgPower 
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "frmSimon.frx":3B660
      Top             =   2340
      Width           =   240
   End
   Begin VB.Image imgStart 
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "frmSimon.frx":3BBEA
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2580
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2310
      TabIndex        =   3
      Top             =   2310
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2310
      TabIndex        =   4
      Top             =   2610
      Width           =   555
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2310
      TabIndex        =   5
      Top             =   2910
      Width           =   735
   End
   Begin VB.Image imgYellow 
      Height          =   2775
      Index           =   2
      Left            =   2760
      Picture         =   "frmSimon.frx":3C174
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Image imgBlue 
      Height          =   2775
      Index           =   2
      Left            =   2760
      Picture         =   "frmSimon.frx":3EF7E
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image imgRed 
      Height          =   2775
      Index           =   2
      Left            =   0
      Picture         =   "frmSimon.frx":41E2C
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image imgGreen 
      Height          =   2775
      Index           =   2
      Left            =   0
      Picture         =   "frmSimon.frx":4455F
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "frmSimon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vincent Foster 01/04/2002 © Copyright 2002 VBVince Software Co.
'Please Give Me credit If You Any Of My Code
'Please Vote
'alienheretic@attbi.com
'http://www.vbvince.com

Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Dim GameOver As Boolean
Dim AINotes(72) As Integer
Dim PlayerNotes(72) As Integer
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim intPlayerTurn As Integer
Dim intPlayerNotesPlayed As Integer
Dim intAITurn As Integer
Dim intAiNotesPlayed As Integer
Const ALTERNATE = 1
Dim Pts(7) As POINTAPI
Dim IntLevel As Integer
Enum ColorEnum
    Red = 0
    Green = 1
    Blue = 2
    Yellow = 3
End Enum
Private Sub Form_Load()
    Paint
    modMidi.InitMidi 'Initallize The Midi Device
    IntLevel = 1
End Sub
Private Sub Paint()
Dim hRgn As Long
'Shapes Our Form To The Outline Of The Simon
'Graphic Making The Corners Transparent
    Pts(0).X = 116
    Pts(0).Y = 0
    Pts(1).X = 252
    Pts(1).Y = 0
    Pts(2).X = 368
    Pts(2).Y = 116
    Pts(3).X = 368
    Pts(3).Y = 252
    Pts(4).X = 252
    Pts(4).Y = 368
    Pts(5).X = 116
    Pts(5).Y = 368
    Pts(6).X = 0
    Pts(6).Y = 252
    Pts(7).X = 0
    Pts(7).Y = 116
    hRgn = CreatePolygonRgn(Pts(0), 8, ALTERNATE)
    SetWindowRgn Me.hwnd, hRgn, True
    DeleteObject hRgn
End Sub
Private Sub Form_Unload(Cancel As Integer)
    modMidi.EndMidi 'Frees Up The Midi Device
    Set frmSimon = Nothing 'Sets The Form To Nothing
End Sub
Public Sub FormDrag(TheForm As Form) 'Drags A Borderless Form
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Private Sub imgBlue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CheckXY ((X / Screen.TwipsPerPixelX) + imgBlue(2).Left), ((Y / Screen.TwipsPerPixelY) + imgBlue(2).Top)
    End If 'Call The CheckXY Mouse Map Sub
End Sub
Private Sub imgYellow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CheckXY ((X / Screen.TwipsPerPixelX) + imgYellow(2).Left), ((Y / Screen.TwipsPerPixelY) + imgYellow(2).Top)
    End If 'Call The CheckXY Mouse Map Sub
End Sub
Private Sub imgGreen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CheckXY ((X / Screen.TwipsPerPixelX) + imgGreen(2).Left), ((Y / Screen.TwipsPerPixelY) + imgGreen(2).Top)
    End If 'Call The CheckXY Mouse Map Sub
End Sub
Private Sub imgRed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CheckXY ((X / Screen.TwipsPerPixelX) + imgRed(2).Left), ((Y / Screen.TwipsPerPixelY) + imgRed(2).Top)
    End If 'Call The CheckXY Mouse Map Sub
End Sub
Private Sub imgLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLevel(0).Picture = imgLevel(1).Picture
End Sub
Private Sub imgLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLevel(0).Picture = imgLevel(2).Picture
End Sub
Private Sub imgLevel_Click(Index As Integer)
Dim K As Integer 'Changes The Level Of Difficulty
    IntLevel = IntLevel + 1
    
    If IntLevel >= 10 Then
        IntLevel = 1
    End If
    For K = 0 To 1
        lblLevel(K).Caption = "Level " & IntLevel
    Next
End Sub
Private Sub imgPower_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPower(0).Picture = imgPower(1).Picture
End Sub
Private Sub imgPower_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPower(0).Picture = imgPower(2).Picture
End Sub
Private Sub imgPower_Click(Index As Integer)
    Unload Me
End Sub
Private Sub imgStart_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgStart(0).Picture = imgStart(1).Picture
End Sub
Private Sub imgStart_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgStart(0).Picture = imgStart(2).Picture
End Sub
Private Sub imgStart_Click(Index As Integer)
    NewGame 'Calls The New Game Sub
End Sub
Public Sub NewGame()
'Begins A New Game
Intro
RandomizeNotes
PauseNote 200
intPlayerTurn = 0
intPlayerNotesPlayed = 0
intAITurn = 0
intAiNotesPlayed = 0
GameOver = False
AiTurn
End Sub
Public Sub CheckXY(X As Single, Y As Single)
'Checks The Color Of The Mouse Map At Varrious X and Y Possitions
'And Either Allows You To Drag The Form Or Play A Simon note
'Check To Make Sure The Player Has Played The Correct Notes Or
'Calls The Youloose Sub Else Calls The YouWin Sub If The Number Of Played Notes
'=50 Times The Level Of Difficulty
Dim PIX As Long
    PIX = GetPixel(Me.hdc, X, Y)

If GameOver = True Then
    Exit Sub
End If

    Select Case PIX
    Case vbBlack
        FormDrag frmSimon
    Case vbBlue
            AddPlayerNotes Blue
            PlaySimon Blue, 200
    Case vbRed
            AddPlayerNotes Red
            PlaySimon Red, 200
    Case vbGreen
            AddPlayerNotes Green
            PlaySimon Green, 200
    Case vbYellow
            AddPlayerNotes Yellow
            PlaySimon Yellow, 200
    End Select
intPlayerNotesPlayed = intPlayerNotesPlayed + 1
intPlayerTurn = intPlayerTurn + 1
If CheckWin = False Then
YouLoose
Exit Sub
ElseIf intPlayerNotesPlayed = 1 + (7 * IntLevel) Then
YouWin
Exit Sub
End If
If intPlayerTurn = intAiNotesPlayed Then
intPlayerTurn = 0
AiTurn
End If
End Sub
Public Sub RandomizeNotes()
'Randomize The computers notes to play
Dim C As Integer
Dim K As Integer
Dim Z As Long
For K = 0 To 72
Randomize GetTickCount
Z = Rnd * 3
AINotes(K) = Z
Debug.Print AINotes(K)
Next
End Sub
Public Sub YouWin()
'Play The winning Sequence
Dim K As Integer
PauseNote 100
    PlaySimon Red, 150
    PlaySimon Green, 150
    PlaySimon Blue, 150
    PlaySimon Yellow, 150
    PlaySimon Red, 150
    PlaySimon Green, 150
    PlaySimon Blue, 150
    PlaySimon Yellow, 150
    PlaySimon Yellow, 150
    PlaySimon Yellow, 150
    PlaySimon Green, 150
    PlaySimon Blue, 150
    PlaySimon Yellow, 150
    
    GameOver = True
End Sub
Public Sub PlaySimon(eColor As ColorEnum, Durration As Long)
'Play The Approiate Note and Highlight The Picture
Select Case eColor
Case 0
    imgRed(2).Picture = imgRed(0).Picture
    imgRed(2).Refresh
    modMidi.PlayNote 36, 1, 127, Durration
Case 1
    imgGreen(2).Picture = imgGreen(0).Picture
    imgGreen(2).Refresh
    modMidi.PlayNote 40, 1, 127, Durration
Case 2
    imgBlue(2).Picture = imgBlue(0).Picture
    imgBlue(2).Refresh
    modMidi.PlayNote 43, 1, 127, Durration
Case 3
    imgYellow(2).Picture = imgYellow(0).Picture
    imgYellow(2).Refresh
    modMidi.PlayNote 48, 1, 127, Durration
End Select
    ResetButtons
End Sub
Public Sub ResetButtons()
'Darken All The Buttons
    imgBlue(2).Picture = imgBlue(1).Picture
    imgYellow(2).Picture = imgYellow(1).Picture
    imgGreen(2).Picture = imgGreen(1).Picture
    imgRed(2).Picture = imgRed(1).Picture
    imgBlue(2).Refresh
    imgYellow(2).Refresh
    imgGreen(2).Refresh
    imgRed(2).Refresh
End Sub
Public Sub LightAllButtons()
'Light All The Buttons
    imgBlue(2).Picture = imgBlue(0).Picture
    imgYellow(2).Picture = imgYellow(0).Picture
    imgGreen(2).Picture = imgGreen(0).Picture
    imgRed(2).Picture = imgRed(0).Picture
    imgBlue(2).Refresh
    imgYellow(2).Refresh
    imgGreen(2).Refresh
    imgRed(2).Refresh
End Sub
Public Sub YouLoose()
'Play The Loosing Sequence
Dim K As Integer
PauseNote 100
    For K = 0 To 2
        LightAllButtons
        modMidi.PlayNote 26, 1, 127, 250
        ResetButtons
        PauseNote 200
    Next
   GameOver = True
End Sub
Public Sub Intro()
'Play An Introduction Of A New Game
Dim K As Integer
    PlaySimon Red, 150
    PlaySimon Blue, 150
    PlaySimon Yellow, 150
    PlaySimon Green, 150
    PlaySimon Red, 150
    PlaySimon Blue, 150
    PlaySimon Yellow, 150
    PlaySimon Green, 150
    LightAllButtons
    modMidi.PlayNote 26, 1, 127, 200
    ResetButtons
    PauseNote 200
End Sub
Public Function CheckWin() As Boolean
'Compare The Players Array Of Played Notes with The Computers
Dim K As Integer
    For K = 0 To (intPlayerNotesPlayed - 1)
        If PlayerNotes(K) = AINotes(K) Then
            CheckWin = True
        Else
            CheckWin = False
        End If
    Next
End Function
Public Sub AddPlayerNotes(eColor As ColorEnum)
'Add Current Note To The Players Note Array
    Select Case eColor
    Case Red
        PlayerNotes(intPlayerNotesPlayed) = 0
    Case Green
        PlayerNotes(intPlayerNotesPlayed) = 1
    Case Blue
        PlayerNotes(intPlayerNotesPlayed) = 2
    Case Yellow
        PlayerNotes(intPlayerNotesPlayed) = 3
    End Select
End Sub
Public Sub AiTurn()
Dim K As Integer
If GameOver = True Then
    Exit Sub
End If
intAiNotesPlayed = 0 'Reset The Number Of Notes The Computer Has Played To 0
intPlayerNotesPlayed = 0 'Reset The Number Of Notes The Player Has Played To 0
PauseNote 600
    For K = 0 To intAITurn
    PlaySimon (AINotes(K)), 200 'Play The Sequence
    PauseNote 200
    intAiNotesPlayed = intAiNotesPlayed + 1 'Add One To The Number Of Notes The Computer Has Played
    Next
intAITurn = intAITurn + IntLevel 'Count The Turns The Computer Has Had
End Sub
Private Sub lblEmail_Click()
sendemail
End Sub
