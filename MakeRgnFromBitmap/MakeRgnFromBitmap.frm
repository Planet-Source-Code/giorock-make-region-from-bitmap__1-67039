VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   LinkTopic       =   "Form1"
   Picture         =   "MakeRgnFromBitmap.frx":0000
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   83
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   420
      Top             =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************
' Written by GioRock *
'*********************
'*********************
'  Region as manual  *
'*********************

' -> Use ESC Key to Stop Move <-

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Constant to drag form on mouse down
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' Constant to make region "Or"
Private Const RGN_OR = &H2

' MOVE FORM ON SCREEN
Dim iMX As Integer
Dim iMY As Integer
Dim iW As Integer
Dim iH As Integer

Dim iMoveLeft As Boolean
Dim iMoveDown As Boolean
Dim iMoveRight As Boolean
Dim iMoveUp As Boolean

Const STEP_MOVE = 96

Private bRegion As Boolean
Private bDrag As Boolean


Private Function MakeRegionFromBitmap(hDC As Long, Width As Long, Height As Long, TransparentColor As Long) As Long
' The usual "Make region by bitmap" procedure
' TransparentColor is any you want
' Normally I use GetPixel(hDC, 1, 1)
' to return Transparent Color
Dim X As Long, Y As Long, StartLineX As Long
Dim FullRegion As Long, LineRegion As Long
Dim bInFirstRegion As Boolean
Dim bInLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence

    bInFirstRegion = True
    bInLine = False
    X = Y = StartLineX = 0

    For Y = 0 To Height - 1
        For X = 0 To Width
            If GetPixel(hDC, X, Y) = TransparentColor Or X = Width Then
                ' We reached a transparent pixel
                If bInLine Then
                    bInLine = False
                    ' Make a region for any single pixel
                    ' in Transparent Color
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    If bInFirstRegion Then
                        FullRegion = LineRegion
                        bInFirstRegion = False
                    Else
                        ' Combine any region into one only
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your unused region
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not bInLine Then
                    bInLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next

    ' Finally return a full region
    MakeRegionFromBitmap = FullRegion
    
End Function



Private Sub DragWindow(hWndW As Long)
    Call ReleaseCapture
    Call SendMessage(hWndW, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    ' Scramble Timer
    ' to start in any
    ' screen point randomly
    Randomize Timer
    
    iMX = CSng(Rnd * Screen.Width) / 2
    iMY = CSng(Rnd * Screen.Height) / 2
    
    ' Move form in start up position
    ' I set StartUpPosition to Manual
    Me.Move iMX, iMY
    
    iW = Width - (STEP_MOVE / 2)
    iH = Height - (STEP_MOVE / 2)
    ' Here you can choose
    ' inital movement verse
    iMoveRight = True
    iMoveDown = True
    
    ' You must set this
    ' before creation
    Me.AutoRedraw = True
    
    ' Go To Creation Sub
    CreateOrRestoreRegion
    
    ' After don't import
    Me.AutoRedraw = False

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bDrag = True
        DragWindow hwnd
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CreateOrRestoreRegion
    If bDrag Then: ReleaseCapture
    End
    Set Form1 = Nothing
End Sub


Private Sub Timer1_Timer()

    ' Perform any control to move form on Screen
    
    If iMoveRight Then
        If iMX + iW + STEP_MOVE < Screen.Width Then
            iMX = iMX + STEP_MOVE
        Else
            iMoveRight = False
            iMoveLeft = True
        End If
    End If
    
    If iMoveDown Then
        If iMY + iH + STEP_MOVE < Screen.Height Then
            iMY = iMY + STEP_MOVE
        Else
           iMoveDown = False
           iMoveUp = True
        End If
    End If
    
    If iMoveLeft Then
        If iMX - STEP_MOVE > 0 Then
            iMX = iMX - STEP_MOVE
        Else
            iMoveLeft = False
            iMoveRight = True
        End If
    End If

    If iMoveUp Then
        If iMY - STEP_MOVE > 0 Then
            iMY = iMY - STEP_MOVE
        Else
            iMoveUp = False
            iMoveDown = True
        End If
    End If

    Move iMX, iMY
    
'    Refresh

End Sub



Private Sub CreateOrRestoreRegion()
Dim hRgn As Long

    ' I use this Sub to create
    ' and restore region
    ' at occurrency
    If Not bRegion Then
        ' Create a Region from Bitmap
        hRgn = MakeRegionFromBitmap(hDC, ScaleWidth, ScaleHeight, GetPixel(hDC, 1, 1))
        ' Start timer to move form
        ' If you don't start timer
        ' can move form on screen
        ' Draggin' with left button mouse
        Timer1.Enabled = True
    Else
        ' Restore original rectangular Region
        hRgn = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
        ' Stop movement
        Timer1.Enabled = False
    End If
    
    ' set creation flag
    ' to prevent double work
    bRegion = Not (bRegion)
    
    ' You can adjust region here
    ' if necessary
    'OffsetRgn hRgn, -1, -1
    
    ' Set desired Region in form
    SetWindowRgn hwnd, hRgn, True
    
    ' clean up created region
    ' if you don't use longer
    DeleteObject hRgn
    
End Sub
