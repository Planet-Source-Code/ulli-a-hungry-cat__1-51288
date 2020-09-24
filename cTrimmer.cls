VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrimmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Idea based on code found @ PSC (author unknown but tnx anyway)

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const HTCAPTION         As Long = 2

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_OR            As Long = 2
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private TotalRegion             As Long 'handle of accumulated total region
Private PartRegion              As Long 'handle of part region

Private Sub AddRegion(ByRef Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)

    PartRegion = CreateRectRgn(Left, Top, Right, Bottom) 'create part region
    CombineRgn TotalRegion, TotalRegion, PartRegion, RGN_OR 'add part region to total
    DeleteObject PartRegion 'kill part region
    Left = -1 'invalidate start posn

End Sub

Public Sub GrabForm(Frm As Form)

    ReleaseCapture
    SendMessage Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'grab form

End Sub

Public Sub TrimForm(Frm As Form)

  Dim X                 As Long
  Dim SavedX            As Long
  Dim Y                 As Long
  Dim TransparentColor  As Long
  Dim BackgroundColor   As Long
  Dim SavedSM           As Long
  Dim SavedAR           As Boolean

    With Frm
        If .Picture Then

            SavedSM = .ScaleMode
            SavedAR = .AutoRedraw
            .AutoRedraw = True
            .ScaleMode = vbPixels 'API works with pixels

            TransparentColor = .Point(0, 0) 'get transparent color from top left pixel
            BackgroundColor = .BackColor
            If BackgroundColor < 0 Then
                BackgroundColor = GetSysColor(BackgroundColor And &H7FFFFFFF)
            End If

            TotalRegion = CreateRectRgn(0, 0, 0, 0) 'create null region
            SavedX = -1

            For Y = 0 To .ScaleHeight - 1
                For X = 0 To .ScaleWidth - 1
                    Select Case .Point(X, Y)
                      Case BackgroundColor 'hit the form background, expect no more picture-pixels
                        Exit For '>---> Next
                      Case TransparentColor 'pixel will be invisible
                        If SavedX > -1 Then 'if we have a start also, then we have a valid region
                            AddRegion SavedX, Y, X, Y + 1
                        End If
                      Case Else 'visible part
                        If SavedX = -1 Then 'no begin yet so..
                            SavedX = X      '..save this pix posn as begin
                        End If
                    End Select
                Next X
                If SavedX > -1 Then
                    AddRegion SavedX, Y, X, Y + 1
                End If
            Next Y

            SetWindowRgn .hWnd, TotalRegion, True  'trim window and redraw
            DeleteObject TotalRegion 'tidy up

            .ScaleMode = SavedSM
            .AutoRedraw = SavedAR 'reset original form attributes

        End If
    End With 'FRM

End Sub

':) Ulli's VB Code Formatter V2.16.13 (2004-Jan-28 00:47) 19 + 76 = 95 Lines
