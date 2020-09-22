VERSION 5.00
Begin VB.Form Fancyfrm 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fancy Form"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9045
   Icon            =   "Fancyfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBody 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   30
      ScaleHeight     =   307
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   0
      Top             =   390
      Width           =   9000
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   5
      Top             =   0
      Width           =   9045
      Begin VB.CommandButton cmdClock 
         Height          =   345
         Left            =   330
         Picture         =   "Fancyfrm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Turn clock off"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdMusic 
         Height          =   345
         Left            =   0
         Picture         =   "Fancyfrm.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Turn music off"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdExit 
         Height          =   345
         Left            =   660
         Picture         =   "Fancyfrm.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   435
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   3
      Top             =   5055
      Width           =   9045
      Begin VB.PictureBox picSBPanel1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   0
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   83
         TabIndex        =   8
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lblMusic 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "lblMusic"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   1440
         TabIndex        =   10
         Top             =   60
         Width           =   825
      End
      Begin VB.Label lblClock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "lblClock"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   7800
         TabIndex        =   4
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1020
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   390
      Picture         =   "Fancyfrm.frx":0CD0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   450
      Top             =   1500
   End
   Begin VB.Timer Timer2 
      Left            =   1020
      Top             =   1500
   End
   Begin VB.Image imgClock2 
      Height          =   240
      Left            =   1800
      Picture         =   "Fancyfrm.frx":1912
      Top             =   4620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgClock 
      Height          =   240
      Left            =   1380
      Picture         =   "Fancyfrm.frx":1C54
      Top             =   4620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgMusic2 
      Height          =   240
      Left            =   990
      Picture         =   "Fancyfrm.frx":1F96
      Top             =   4620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMusic 
      Height          =   240
      Left            =   600
      Picture         =   "Fancyfrm.frx":22D8
      Top             =   4620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgExit 
      Height          =   240
      Left            =   2220
      Picture         =   "Fancyfrm.frx":261A
      Top             =   4620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMusic 
         Caption         =   "&Music"
      End
      Begin VB.Menu mnuFileClock 
         Caption         =   "&Clock"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Fancyfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FancyFrm.frm
'
' By Herman Liu
'
' A fancy form - just for fun.
'
' The form is "fancy" because it consists of (1) a title bar which is of a color continued
' from the gradient of form body colors; (2) scrolling text and images in the form body,
' with stepped changes in text and image colors; (3) bitmaps in menus; (4) a colored status
' bar; as well as (5) the accompanying background music of a matching duration.
'
' All is done with a few simple API calls, and using only one form (as I always try to do).
'
 
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, _
    lpSysColor As Long, lpColorValues As Long) As Long
    
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, _
    ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
    ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
    ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, _
    ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) As Long
    
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal Z As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_CAPTIONTEXT = 9

Private Const MF_BITMAP = &H4&

Private Const SND_SYNC = &H0              '  play synchronously (default)
Private Const SND_ASYNC = &H1             '  play asynchronously
Private Const SND_LOOP = &H8              '  loop the sound until next sndPlaySound
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found

Dim arrStdPic(1) As New StdPicture
Dim hdctemp(1) As Long
Dim hbmSrc(1) As Long
Dim hPalOld(1) As Long
Dim hPal(1) As Long
Dim Wbody As Long
Dim Hbody As Long
Dim Ydefault As Long
Dim YbodyPos As Long
Dim mLineSpace As Integer
Dim arrMsgLine() As String
Dim MusicFile As String
Dim HasMusicFile As Boolean
Dim MusicOnFlag As Boolean
Dim ClockOnFlag As Boolean
Dim origTitleBarColor
Dim origTitleBarTextColor
Dim mAbort As Boolean
Dim mOK As Boolean



Private Sub Form_Load()
    Dim tmp As String
    Dim i
    Dim j
    Dim R, g, B
    Dim interval
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
       ' Set align property
    picBody.Align = 1         ' top
    picStatusBar.Align = 2    ' bottom
    
       ' Properties of picStatusBar
    picStatusBar.ScaleMode = vbPixels
    picStatusBar.AutoRedraw = True
    
       ' Some positioning
    picSBPanel1.Move 1, 0
    lblMusic.Left = picSBPanel1.Left + picSBPanel1.Width + 4
    lblClock.Left = picStatusBar.Width - lblClock.Width - 6
    
       ' Properties of form body display medium
    picBody.ScaleMode = vbPixels
    picBody.AutoRedraw = True
    picBody.Visible = False
    
       ' Pictureboxes used to provide faded images of logo
    For i = 0 To 1
        picLogo(i).ScaleMode = vbPixels
        picLogo(i).Appearance = 0
        picLogo(i).AutoRedraw = True
        If i > 0 Then
            picLogo(i).Width = picLogo(0).Width
            picLogo(i).Height = picLogo(0).Height
        End If
        picLogo(i).AutoSize = True
        picLogo(i).Visible = False
    Next i
    
       ' Text and images to display in form body
    ReDim arrMsgLine(15, 1)
    arrMsgLine(0, 0) = "fuCK YoU aIMAn"
    arrMsgLine(1, 0) = Space(1)
    arrMsgLine(2, 0) = "* Title bar of a color continued from"
    arrMsgLine(3, 0) = "gradient form body colors"
    arrMsgLine(4, 0) = Space(1)
    arrMsgLine(5, 0) = "* Scrolling text and images in form body"
    arrMsgLine(6, 0) = "with stepped changes in text and image colors"
    arrMsgLine(7, 0) = Space(1)
    arrMsgLine(8, 0) = "* Bipmaps in form menu"
    arrMsgLine(9, 0) = Space(1)
    arrMsgLine(10, 0) = "* Colored status bar"
    arrMsgLine(11, 0) = Space(1)
    arrMsgLine(12, 0) = "* Background music of a matching duration ......"
    arrMsgLine(13, 0) = Space(1)
    arrMsgLine(14, 0) = "Herman Liu"
    arrMsgLine(15, 0) = CStr(Date)
    arrMsgLine(0, 1) = 16
    For i = 1 To 12
         arrMsgLine(i, 1) = 14
    Next i
    arrMsgLine(14, 1) = 12
    arrMsgLine(15, 1) = 8
    
    For i = 0 To UBound(arrStdPic)
        Set arrStdPic(i) = New StdPicture
    Next i
    
       ' Store original title bar colors before changing them
    origTitleBarColor = GetSysColor(COLOR_ACTIVECAPTION)
    origTitleBarTextColor = GetSysColor(COLOR_CAPTIONTEXT)
    
    Wbody = picBody.ScaleWidth
    Hbody = picBody.ScaleHeight
      ' We set to start from this position
    Ydefault = Hbody
    YbodyPos = Ydefault
    
      ' Control line spacing of body display
    mLineSpace = 10
        
      ' Form body colors
    B = 255
    interval = 7
    j = interval * (0 - B) / picBody.ScaleHeight
    For i = 0 To picBody.ScaleHeight + 1 Step interval
         If B < 0 Then B = 0
         If B > 255 Then B = 255
         picBody.Line (0, i)-(picBody.ScaleWidth, i + interval), RGB(0, 0, B), BF
         B = B + j
    Next i
    picBody.Picture = picBody.Image
    
      ' Store body colors in a memory DC, ready for repeated use later
    Set arrStdPic(0) = picBody.Picture
    mOK = CreateDC(picBody, 0)
    If mOK = True Then
          ' Prepare faded pictures for use later (Note we have to do this one first
          ' before setting stdpic=picLogo(0).Picture)
         FadePic picLogo(0), picLogo(1), 1.3
         
         Set arrStdPic(1) = picLogo(0).Picture
         mOK = CreateDC(arrStdPic(1), 1)
         
         If mOK Then
              FadePic picLogo(1), picLogo(0), 0.4
         End If
    End If
    
       ' Add bitmaps to menu
    AddBitMap
    
    ClockOnFlag = True
    
    MusicFile = App.Path & "\Beetvn9.wav"
    If IsFileThere(MusicFile) Then
         HasMusicFile = True
         MusicOnFlag = True
         lblMusic.Caption = "Music on"
         sndPlaySound MusicFile, SND_ASYNC Or SND_LOOP
    Else
         HasMusicFile = False
         MusicOnFlag = False
         mnuFileMusic.Enabled = False
         cmdMusic.Enabled = False
         lblMusic.Caption = "Music off"
    End If
End Sub



Private Sub Form_activate()
    If Not mOK Then
         MsgBox "StdPicture not of bitmap type"
         Exit Sub
    End If
    DrawMusicDispBox
    SetSysColors 1, COLOR_ACTIVECAPTION, RGB(0, 0, 255)
    SetSysColors 1, COLOR_CAPTIONTEXT, &HFFFF&
    Timer1.interval = 800
    Timer1.Enabled = True
    Timer2.interval = 10
    Timer2.Enabled = True
    mAbort = False
End Sub



Private Sub Form_Unloadx(Cancel As Integer)
    On Error Resume Next
    Dim i
    mAbort = True
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    DoEvents
    
    SetSysColors 1, COLOR_ACTIVECAPTION, origTitleBarColor
    SetSysColors 1, COLOR_CAPTIONTEXT, origTitleBarTextColor
    For i = 0 To UBound(hdctemp)
         Call SelectObject(hdctemp(i), hbmSrc(i))
         Call DeleteDC(hdctemp(i))
         Call DeleteObject(hdctemp(i))
    Next i
    For i = 0 To UBound(arrStdPic)
         Set arrStdPic(i) = Nothing
    Next i
    If MusicOnFlag Then
         sndPlaySound vbNullString, SND_ASYNC
    End If
End Sub



Private Sub mnuFileMusic_Click()
    cmdMusic_Click
End Sub


Private Sub mnuFileClock_Click()
    cmdClock_Click
End Sub



Private Sub cmdMusic_Click()
    If Not HasMusicFile Then
          Exit Sub
    End If
    MusicOnFlag = Not MusicOnFlag
    If MusicOnFlag Then
         cmdMusic.Picture = imgMusic
         cmdMusic.ToolTipText = "Turn music off"
         lblMusic.Caption = "Music on"
         sndPlaySound MusicFile, SND_ASYNC Or SND_LOOP
    Else
         cmdMusic.ToolTipText = "Turn music on"
         cmdMusic.Picture = ImgMusic2
         lblMusic.Caption = "Music off"
         sndPlaySound vbNullString, SND_ASYNC
    End If
    MnuImageOnOff 0, MusicOnFlag
End Sub



Private Sub DrawMusicDispBox()
    On Error Resume Next
    Dim m As Long
    Dim x As Long, y As Long, W As Long, H As Long
    m = 2
    x = lblMusic.Left - m
    y = lblMusic.Top - m
    W = lblMusic.Width + m * 2
    H = lblMusic.Height + m * 2
    DrawBevelledBox picStatusBar, x, y, W, H, 2, RGB(100, 100, 100), RGB(200, 200, 200)
End Sub


Private Sub DrawBevelledBox(ByVal inPic As PictureBox, inLeft As Long, inTop As Long, _
    inWidth As Long, inHeight As Long, inBevel As Long, inColor1 As Long, inColor2 As Long)
    On Error Resume Next
    Dim i As Long
    If inBevel < 1 Then inBevel = 1
    For i = 1 To inBevel
         MoveToEx inPic.hdc, inLeft - i, (inTop + inHeight + (i - 1)), 0
           ' Note LineTo draws a line up to one point before the specified point
         inPic.ForeColor = inColor1
         LineTo inPic.hdc, inLeft - i, inTop - i
         LineTo inPic.hdc, (inLeft + inWidth + (i - 1)), inTop - i
         inPic.ForeColor = inColor2
         LineTo inPic.hdc, (inLeft + inWidth + (i - 1)), (inTop + inHeight + (i - 1))
         LineTo inPic.hdc, (inLeft - i), (inTop + inHeight + (i - 1))
         DoEvents
    Next i
End Sub



Private Sub cmdClock_Click()
    ClockOnFlag = Not ClockOnFlag
    If ClockOnFlag Then
         cmdClock.Picture = imgClock
         cmdClock.ToolTipText = "Turn clock off"
         lblClock.Visible = True
         Timer1.Enabled = True
    Else
         cmdClock.ToolTipText = "Turn clock on"
         cmdClock.Picture = imgClock2
         lblClock.Visible = False
         Timer1.Enabled = False
    End If
    MnuImageOnOff 1, ClockOnFlag
End Sub



Private Sub mnuFileExit_Click()
    Unload Me
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim i
    mAbort = True
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    DoEvents
    
    SetSysColors 1, COLOR_ACTIVECAPTION, origTitleBarColor
    SetSysColors 1, COLOR_CAPTIONTEXT, origTitleBarTextColor
    DoEvents
    For i = 0 To UBound(hdctemp)
         Call SelectObject(hdctemp(i), hbmSrc(i))
         Call DeleteDC(hdctemp(i))
         Call DeleteObject(hdctemp(i))
    Next i
    For i = 0 To UBound(arrStdPic)
         Set arrStdPic(i) = Nothing
    Next i
    If MusicOnFlag Then
         sndPlaySound vbNullString, SND_ASYNC
    End If
End Sub



Private Sub AddBitMap()
    Dim i As Integer
    Dim mMenu As Long, mSubMenu As Long, mSubMenuID As Long
    mMenu = GetMenu(Me.hwnd)
    mSubMenu = GetSubMenu(mMenu, 0)
      ' Remarks: Normally we would use ImageList/array and loop through
      ' submenus to load images, but here we only have a few.
    mSubMenuID = GetMenuItemID(mSubMenu, 0)
    SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgMusic.Picture, imgMusic.Picture
    mSubMenuID = GetMenuItemID(mSubMenu, 1)
    SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgClock.Picture, imgClock.Picture
    mSubMenuID = GetMenuItemID(mSubMenu, 2)
    SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgExit.Picture, imgExit.Picture
End Sub



Private Sub MnuImageOnOff(ByVal inItem As Integer, ByVal OnOff As Boolean)
    Dim mMenu As Long, mSubMenu As Long, mSubMenuID As Long
    mMenu = GetMenu(Me.hwnd)
    mSubMenu = GetSubMenu(mMenu, 0)               ' We only have 1 menu here this time.
    mSubMenuID = GetMenuItemID(mSubMenu, inItem)
      ' Remarks: Normally we would use ImageList/array, here we have a few images only
    If inItem = 0 Then
        If OnOff = True Then
            SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgMusic.Picture, imgMusic.Picture
        Else
            SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, ImgMusic2.Picture, ImgMusic2.Picture
        End If
    Else
        If OnOff = True Then
            SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgClock.Picture, imgClock.Picture
        Else
            SetMenuItemBitmaps mMenu, mSubMenuID, MF_BITMAP, imgClock2.Picture, imgClock2.Picture
        End If
    End If
End Sub



Private Sub FadePic(ByVal inPic1 As PictureBox, ByVal inPic2 As PictureBox, _
        ByVal fraction As Single)
    If fraction = 1 Then
         Exit Sub
    End If
    Dim i, j, c
    Dim B, g, R
    For i = 0 To inPic1.ScaleWidth
        For j = 0 To inPic1.ScaleHeight
            c = inPic1.Point(i, j)
            R = Abs(c Mod &H100) * fraction
            g = Abs((c \ &H100) Mod &H100) * fraction
            B = Abs((c \ &H10000) Mod &H100) * fraction
            inPic2.PSet (i, j), RGB(R, g, B)
        Next j
        DoEvents
    Next i
    inPic2.Picture = inPic2.Image
End Sub



Private Function CreateDC(ByVal inStdPic As StdPicture, ByVal inIndex As Integer) As Boolean
    On Error GoTo errHandler
    CreateDC = False
    If inStdPic.Type = vbPicTypeBitmap Then
         hPal(inIndex) = CreateHalftonePalette(0&)
         hdctemp(inIndex) = CreateCompatibleDC(0&)
         hPalOld(inIndex) = SelectPalette(hdctemp(inIndex), hPal(inIndex), True)
         hbmSrc(inIndex) = SelectObject(hdctemp(inIndex), inStdPic.Handle)
         CreateDC = True
    End If
    Exit Function
errHandler:
    CreateDC = False
End Function



Private Sub Timer1_Timer()
    lblClock.Caption = Format$(Now, "hh:mm:ss AM/PM") & Space(2)
End Sub



Private Sub Timer2_Timer()
    Dim i
    Dim R, g, B
    Dim DispTally
    Dim x, y, W, H
    On Error Resume Next
    
      ' Renew background. We copy it from DC
    BitBlt picBody.hdc, 0, 0, Wbody, Hbody, hdctemp(0), 0, 0, vbSrcCopy
    
    DispTally = 0
    For i = 0 To UBound(arrMsgLine)
        picBody.FontSize = CInt(arrMsgLine(i, 1))
           ' Center horizontally
        picBody.CurrentX = (Wbody / 2) - (picBody.TextWidth(arrMsgLine(i, 0)) / 2)
           ' Calculate height position for the line display
        picBody.CurrentY = YbodyPos + DispTally
        
        W = picLogo(0).ScaleWidth
        H = picLogo(0).ScaleHeight
        x = picBody.CurrentX - W / 2
        y = picBody.CurrentY + 10
        If i = 1 Then               ' Display an image in between
                ' See if to display original or a less sharp one
            If picBody.CurrentY <= (Hbody / 100 * 40) Then
                 If picBody.CurrentY > Hbody / 10 Then
                      BitBlt picBody.hdc, x, y, x + W, y + H, hdctemp(1), 0, 0, vbSrcCopy
                 Else
                      BitBlt picBody.hdc, x, y, x + W, y + H, picLogo(0).hdc, _
                             0, 0, vbSrcCopy
                 End If
            ElseIf picBody.CurrentY >= (Hbody / 100 * 80) Then
                 BitBlt picBody.hdc, x, y, x + W, y + H, picLogo(0).hdc, 0, 0, vbSrcCopy
            Else
                     ' The sharpest one
                  BitBlt picBody.hdc, x, y, x + W, y + H, picLogo(1).hdc, 0, 0, vbSrcCopy
            End If
                ' Adjust height position for the image display
            DispTally = DispTally + picLogo(0).ScaleHeight + 20
            picBody.CurrentY = YbodyPos + DispTally
        End If
           
           ' Vary ForeColor according to line position reached
        If picBody.CurrentY <= (Hbody / 100 * 40) Then
            If picBody.CurrentY > Hbody / 10 Then
                R = (255 / 225) * picBody.CurrentY
                g = (255 / 225) * picBody.CurrentY
                B = (255 / 30) * picBody.CurrentY
            Else
                R = 0: g = 0:  B = 255
                If i = UBound(arrMsgLine) Then
                    If picBody.CurrentY < picBody.TextHeight(arrMsgLine(i, 0)) Then
                         YbodyPos = Ydefault + 1
                         Exit For
                    End If
                End If
            End If
        ElseIf picBody.CurrentY >= (Hbody / 100 * 80) Then
            R = 80 + (Hbody - picBody.CurrentY) * 3
            g = R
            B = R
        Else
            R = 255: g = 255: B = 255
        End If
           ' Tally the height position ready for displaying next line
        DispTally = DispTally + (picBody.TextHeight(arrMsgLine(i, 0)) + mLineSpace)
               
        If R > 255 Then R = 255
        If g > 255 Then g = 255
        If B > 255 Then B = 255
        picBody.ForeColor = RGB(R, g, B)
           ' Note if text in line is too long, some characters would not be printed
        If i = 0 Then
            BeginPath picBody.hdc
            picBody.Print arrMsgLine(i, 0)
            EndPath picBody.hdc
            StrokePath picBody.hdc
            picBody.CurrentY = picBody.CurrentY
        Else
            picBody.Print arrMsgLine(i, 0)
        End If
        If mAbort Then
            Exit For
        End If
    Next
       ' Change starting position.  This controls display pace; almost invariably 1.
    YbodyPos = YbodyPos - 1
       ' Display on screen the text and images printed earlier
    x = (Me.ScaleWidth - Wbody) / 2
    y = (Me.ScaleHeight - Hbody) / 2
    BitBlt Me.hdc, x, y, x + Wbody, y + Hbody, picBody.hdc, 0, 0, vbSrcCopy
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    'Dim mFile As String
    'mFile = LongToShort(inFileSpec)
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function




 
