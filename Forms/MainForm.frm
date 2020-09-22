VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   9000
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   5317.496
   ScaleMode       =   0  'User
   ScaleWidth      =   11267.35
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLogin 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   11295
      Begin VB.Data Login 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Users"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   376
         Left            =   3840
         TabIndex        =   0
         Top             =   1920
         Width           =   4569
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   376
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2640
         Width           =   4569
      End
      Begin VB.Label lblSecurityLevel 
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUserID 
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         MouseIcon       =   "MainForm.frx":1232A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblSubmit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Submit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         MouseIcon       =   "MainForm.frx":1247C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Image ImgLogin 
         Height          =   2265
         Left            =   2160
         Picture         =   "MainForm.frx":125CE
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Image ImgFrmLoginBack 
         Height          =   6570
         Index           =   0
         Left            =   0
         Picture         =   "MainForm.frx":16E28
         Top             =   0
         Width           =   11370
      End
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      MouseIcon       =   "MainForm.frx":1B6B9
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   11640
      MouseIcon       =   "MainForm.frx":1B80B
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblExplainButtom 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   271
      Left            =   373
      TabIndex        =   5
      Top             =   8683
      Width           =   3735
   End
   Begin VB.Label lblLoggedInUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   4410
   End
   Begin VB.Image ImgFormBack 
      Height          =   9390
      Left            =   0
      Picture         =   "MainForm.frx":1B95D
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Login_Attempts As Integer
  Dim Multimedia As New Mmedia
Private Sub Form_Load()
On Error GoTo notfound
Dim source, source1, desti, h, h2, path As String, path1 As String, path2 As String, recover As String, recover1 As String
path = "c:\BTMS backup utility (Backed up).log"
path2 = "c:\BTMS backup utility (Retrieved).log"
path1 = App.path & "\BFN.log"
h = MsgBox("Do you want to back up your data? Press cancel for data recovery options", vbYesNoCancel, "BTMS Backup Utility")
source = App.path & "\btms.dll"
If h = 6 Then
 desti = App.path & "\BackUp\" & Format(Now(), "d-mmmm-yyyy h-mm") & ".dll"
 FileCopy source, desti
 MsgBox ("Your data has been backed up successfully to " & desti + " in your backup folder." & vbNewLine + " For more information please view the log file " & path)
 Open path For Append As #1
 Write #1, "Your data was backed up susccefully in " & Format(Now(), "d/mmmm/yyyy h:mm")
 Close #1
 Open path1 For Append As #2
 Write #2, desti
 Close #2
End If
If h = 2 Then
    h2 = MsgBox("Press yes to retrieve the last data you backed up" & vbNewLine + "Press no to undo the last retrieve.", vbYesNoCancel)
        If h2 = 6 Then
            Open path1 For Input As #3
            Do While Not EOF(3)
            Input #3, recover
            Loop
            Close #3
            recover1 = App.path & "\BackUp\" & "btms.dll"
            FileCopy source, recover1
            FileCopy recover, source
            MsgBox ("Your data has been retrieved successfully with the last backed up data for more information please view the log file " & path2)
            Open path2 For Append As #4
            Write #4, "Your data was retrieved susccefully in " & Format(Now(), "d/mmmm/yyyy h:mm")
            Close #4
        End If
        If h2 = 7 Then
            source = App.path & "\BackUp\" & "btms.dll"
            recover = App.path & "\btms.dll"
            FileCopy source, recover
            MsgBox ("Undo of last retreved data was successfully done.")
        End If
End If

MainForm.Top = (Screen.Height - MainForm.Height) / 2
MainForm.Left = (Screen.Width - MainForm.Width) / 2
MsgBox ("For best resolution, please use 800x600 ok??")
lblTime.Caption = Format(Now, "Long Date")
Login_Attempts = 0
notfound:
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox ("Sorry, you need to backup before you retrieve." & vbNewLine + "If you have any other problem please reinstall BTMS")
        MainForm.Top = (Screen.Height - MainForm.Height) / 2
        MainForm.Left = (Screen.Width - MainForm.Width) / 2
        MsgBox ("For best resolution, please use 800x600 ok??")
        lblTime.Caption = Format(Now, "Long Date")
        Login_Attempts = 0
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = ""
End Sub
Private Sub frmLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = ""
End Sub
Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = ""
End Sub
Private Sub Label17_Click()
Fancyfrm.Show
End Sub
Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = "Click here to exit"
End Sub
Private Sub lblSubmit_Click()
  Dim Record_Found As Boolean
  'Dim LoginDatabase As Database
  'Dim LoginRecordset As Recordset
  'Dim Database_Path, Database_Password, Database_Name As String
  'Database_Name = "BTMS.dll"
  'Database_Path = ""
  'Database_Password = ""
  Record_Found = False
      'Set LoginDatabase = OpenDatabase(Database_Name, False, False, ";pwd=" & Database_Password)
  Login_Attempts = Login_Attempts + 1
  'Set LoginRecordset = LoginDatabase.OpenRecordset("Users")
    
  Do While Not Login.Recordset.EOF
     If Login.Recordset.Fields("LoginID") = MainForm.txtUserName.Text And _
        Login.Recordset.Fields("LoginPassword") = MainForm.txtPassword.Text Then
        Record_Found = True
        txtUserName.Text = Login.Recordset.Fields("UserName")
        lblUserID.Caption = Login.Recordset.Fields("UserID")
        lblSecurityLevel.Caption = Login.Recordset.Fields("SecurityLevel")
        Login.Recordset.Close
        Login.Database.Close
        Exit Do
       Else
        Login.Recordset.MoveNext
      End If
  Loop
  
  If Record_Found = True Then
     FormProcessing.Show
    Else ' Record_Found =False
     If Login_Attempts < 4 Then
        MsgBox "The entries that you have made are invalid. Note: Values are Case Sensitive.", vbInformation + vbOKOnly
       Else
        MsgBox "Contact an Administrator for a valid login name and password." & vbNewLine + "Good bye", vbInformation
        Login.Recordset.Close
        Login.Database.Close
        Unload Me
     End If
    End If
End Sub
Private Sub lblSubmit_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = "Click here when finished"
End Sub
Private Sub lblCancel_Click()
txtUserName.Text = ""
txtPassword.Text = ""
End Sub
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = "Click here to exit"
End Sub
Private Sub lblExit_Click()
Multimedia.mmOpen "expand.wav"
Multimedia.mmPlay
Dim answer
answer = MsgBox("This will log you off and then exit BTMS. Are you sure?", vbYesNo, "Question")
If answer = 6 Then Unload Me
End Sub
Private Sub txtPassword_Change()
If txtUserName.Text = "" Or txtPassword = "" Then
lblSubmit.Enabled = False
Else
lblSubmit.Enabled = True
End If
End Sub
Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = "Type your Password here"
End Sub
Private Sub txtUserName_Change()
If txtUserName.Text = "" Or txtPassword = "" Then
lblSubmit.Enabled = False
Else
lblSubmit.Enabled = True
End If
End Sub
Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lblExplainButtom.Caption = "Type your login ID here"
End Sub
