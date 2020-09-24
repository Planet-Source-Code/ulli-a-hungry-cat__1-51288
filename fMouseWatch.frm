VERSION 5.00
Begin VB.Form fMouseWatch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   FillColor       =   &H00FFFFC0&
   FillStyle       =   0  'Ausgefüllt
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fMouseWatch.frx":0000
   ScaleHeight     =   67
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmr 
      Interval        =   70
      Left            =   435
      Top             =   390
   End
End
Attribute VB_Name = "fMouseWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As tRECT) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_NOACTIVATE    As Long = 16
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_COMBINED      As Long = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Private Const SWP_TOPMOST       As Long = -1

Private Type tPOINT
    X           As Long
    Y           As Long
End Type

Private Type tRECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Trimmer         As New cTrimmer

Private CursorPos       As tPOINT
Private WindowRect      As tRECT

Private LeftEye         As tPOINT 'center of left eye in screen coords
Private RightEye        As tPOINT 'center of right eye in screen coords

Private LeftDistance    As Double 'distance from left eye center to cursor hotspot
Private RightDistance   As Double 'distance from right eye center to cursor hotspot

Private LeftSin         As Double 'sine of left eye to cursor trajectory
Private LeftCos         As Double 'cosine of left eye to cursor trajectory
Private RightSin        As Double 'sine of right eye to cursor trajectory
Private RightCos        As Double 'cosine of right eye to cursor trajectory

'eye properties
Private Const LX            As Long = 31 'left eye center x
Private Const RX            As Long = 55 'right eye center x
Private Const BY            As Long = 27 'both eyes center y
Private Const EyeBackColor  As Long = &HFFF8F8
Private Const PupilColor    As Long = vbBlue
Private Const PupilRad      As Single = 3 'radius of pupil
Private Const MoveRad       As Single = 3 'radius of pupil movement
Private Const EyeRad        As Single = PupilRad + MoveRad + 1 'radius of eyes

Private Sub Form_DblClick()

  'doubleclick to unload

    Unload Me

End Sub

Private Sub Form_Load()

    Trimmer.TrimForm Me
    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        Trimmer.GrabForm Me
    End If

End Sub

Private Sub tmr_Timer()

  'ticks about 14 times a second

    GetWindowRect hWnd, WindowRect 'where am I
    'transform eye positions to screen coords
    With WindowRect
        LeftEye.X = .Left + LX
        LeftEye.Y = .Top + BY
        RightEye.X = .Left + RX
        RightEye.Y = .Top + BY
    End With 'WINDOWRECT

    GetCursorPos CursorPos 'where's the cursor
    'compute distances and angles
    With CursorPos
        LeftDistance = Sqr((LeftEye.X - .X) ^ 2 + (LeftEye.Y - .Y) ^ 2)
        RightDistance = Sqr((RightEye.X - .X) ^ 2 + (RightEye.Y - .Y) ^ 2)
        If LeftDistance = 0 Then 'prevent zero division
            LeftDistance = 1
        End If
        If RightDistance = 0 Then 'prevent zero division
            RightDistance = 1
        End If
        LeftSin = (LeftEye.Y - .Y) / LeftDistance
        LeftCos = (LeftEye.X - .X) / LeftDistance
        RightSin = (RightEye.Y - .Y) / RightDistance
        RightCos = (RightEye.X - .X) / RightDistance
    End With 'CURSORPOS

    'draw eyes
    FillColor = EyeBackColor
    Circle (LX, BY), EyeRad, vbBlack
    Circle (RX, BY), EyeRad, vbBlack
    FillColor = vbBlack
    Circle (LX - MoveRad * LeftCos, BY - MoveRad * LeftSin), PupilRad, PupilColor
    Circle (RX - MoveRad * RightCos, BY - MoveRad * RightSin), PupilRad, PupilColor

End Sub

':) Ulli's VB Code Formatter V2.16.13 (2004-Jan-28 00:47) 47 + 65 = 112 Lines
