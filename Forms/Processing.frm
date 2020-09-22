VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormProcessing 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   4005
   ClientTop       =   3495
   ClientWidth     =   4530
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Processing.."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6615
      Left            =   -240
      Picture         =   "Processing.frx":0000
      Top             =   -2760
      Width           =   11445
   End
End
Attribute VB_Name = "FormProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormProcessing.Top = (Screen.Height - FormProcessing.Height) / 2
FormProcessing.Left = (Screen.Width - FormProcessing.Width) / 2
MainForm.lblCancel.Enabled = False
MainForm.lblSubmit.Enabled = False
ProgressBar1.Value = ProgressBar1.Min
FormProcessing.MousePointer = 11
End Sub

Private Sub Timer1_Timer()
Dim finish
    If ProgressBar1.Value <> 100 Then
    ProgressBar1.Value = ProgressBar1.Value + 30
    ProgressBar1.Value = ProgressBar1.Value + 70
    End If
    MainForm.lblLoggedInUser.Caption = MainForm.txtUserName.Text
    FormUserWizard.lblLoggedInUser.Caption = MainForm.txtUserName.Text
    MainForm.lblWelcome.Visible = True
    FormUserWizard.lblWelcome.Visible = True
    If MainForm.lblSecurityLevel.Caption = False Then
       FormUserWizard.Show
       'MainForm.frmUserWizard(0).ZOrder (0)
       'MainForm.frmLogin.Visible = False
       Unload Me
    ElseIf MainForm.lblSecurityLevel.Caption = True Then
       FormAdminWizard.Show
       FormAdminWizard.lblLoggedInUser.Caption = MainForm.txtUserName.Text
       FormAdminWizard.lblWelcome.Visible = True
       'MainForm.frmAdminWizard.ZOrder (0)
       'MainForm.frmLogin.Visible = False
       Unload Me
    End If
End Sub
