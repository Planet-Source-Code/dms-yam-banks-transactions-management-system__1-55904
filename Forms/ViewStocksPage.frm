VERSION 5.00
Object = "{F49365FC-E8A5-4E38-9DBC-DAA7D889B8A3}#1.6#0"; "pbxpbutton.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form ViewStocksPage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "This is www.Tadawul.com.sa"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin PB_XP_Button.PBXPButton PBXPButton1 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Go"
      BorderColorOver =   6956042
      BorderColorDown =   6956042
      BackColor       =   49152
      BackColorOver   =   8454016
      BackColorDown   =   49152
      BackColorIcon   =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   1
      ShowShadowOver  =   -1  'True
      ShowFocus       =   -1  'True
      CheckedColor    =   14211029
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   0
      TabIndex        =   1
      Text            =   "http://"
      Top             =   0
      Width           =   5175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "ViewStocksPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multimedia As New Mmedia
Private Sub Form_Load()
MyURL = "http://tadawul.com.sa/quotes/allmarketwatch_HTML_ar_eg.asp"
WebBrowser1.Navigate "" & MyURL & ""
End Sub

Private Sub PBXPButton1_Click()
MyURL = Text1.Text
WebBrowser1.Navigate "" & MyURL & ""
End Sub
