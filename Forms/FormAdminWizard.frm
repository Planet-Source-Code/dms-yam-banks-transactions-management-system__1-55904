VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormAdminWizard 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frm0 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6500
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   11295
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transaction Type Manager"
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
         Height          =   705
         Index           =   3
         Left            =   360
         Picture         =   "FormAdminWizard.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4560
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bank Manager"
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
         Height          =   585
         Index           =   1
         Left            =   360
         Picture         =   "FormAdminWizard.frx":4891
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Company Manager"
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
         Height          =   585
         Index           =   2
         Left            =   360
         Picture         =   "FormAdminWizard.frx":9122
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3600
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "User Manager"
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
         Height          =   585
         Index           =   0
         Left            =   360
         Picture         =   "FormAdminWizard.frx":D9B3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdNext0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":12244
         DownPicture     =   "FormAdminWizard.frx":125C5
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":1305B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label lblQ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "What Would You Like To Do, Sir?!!"
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
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7695
      End
      Begin VB.Image ImgFrm0 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":13A55
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.OptionButton Op1 
         Caption         =   "Modification"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         Picture         =   "FormAdminWizard.frx":189E8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2400
         Width           =   3855
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Addition"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         Picture         =   "FormAdminWizard.frx":1D97B
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CommandButton cmdNext1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":2290E
         DownPicture     =   "FormAdminWizard.frx":22C8F
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":23725
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":2411F
         DownPicture     =   "FormAdminWizard.frx":243AD
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":24E8D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label lblQ1 
         BackStyle       =   0  'Transparent
         Caption         =   "What Would You Like To Do Exactly?"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   7095
      End
      Begin VB.Image Img1 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":2590D
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         DataField       =   "SecurityLevel"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   80
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton CmdBack2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":2A8A0
         DownPicture     =   "FormAdminWizard.frx":2AB2E
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":2B60E
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5640
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFC0C0&
         DataField       =   "UserName"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   33
         Top             =   960
         Width           =   3975
      End
      Begin VB.CommandButton cmdFinish2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":2C08E
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":2C3CF
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data Confirm2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Users"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtPhoneField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "Phone"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   37
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox txtEmailField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "Email"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   36
         Top             =   3480
         Width           =   4575
      End
      Begin VB.TextBox txtPasswordField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LoginPassword"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtLoginIDField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LoginID"
         DataSource      =   "Confirm2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   34
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Level?"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   240
         TabIndex        =   79
         Top             =   5280
         Width           =   2355
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Login ID :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   78
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   1965
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   2760
         Width           =   1965
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone    :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   4440
         Width           =   1965
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Email    :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   3600
         Width           =   1965
      End
      Begin VB.Label lblQ6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Now Update The Profile"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   4305
      End
      Begin VB.Image ImgFrm2 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":2CB97
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdNext4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":31B2A
         DownPicture     =   "FormAdminWizard.frx":31EAB
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":32941
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":3333B
         DownPicture     =   "FormAdminWizard.frx":335C9
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":340A9
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data UserListConn 
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Users"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDBCtls.DBCombo UserList 
         Bindings        =   "FormAdminWizard.frx":34B29
         DataField       =   "UserID"
         DataSource      =   "UserListConn"
         Height          =   360
         Left            =   1080
         TabIndex        =   40
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ListField       =   "UserName"
         BoundColumn     =   "UserID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Users List"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   42
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Choose a User To Edit"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   360
         Width           =   7095
      End
      Begin VB.Image Imgfrm4 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":34B44
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFC0C0&
      Height          =   6495
      Left            =   360
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdFinish3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":39AD7
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":39E18
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":3A5E0
         DownPicture     =   "FormAdminWizard.frx":3A86E
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":3B34E
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data Confirm3 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Company"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFC0C0&
         DataField       =   "CompanyName"
         DataSource      =   "Confirm3"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   66
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Enter The Company Name"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   7815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   67
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image Imgfrm3 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":3BDCE
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm10 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   60
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":40D61
         DownPicture     =   "FormAdminWizard.frx":40FEF
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":41ACF
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":4254F
         DownPicture     =   "FormAdminWizard.frx":428D0
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":43366
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data TTListConn 
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TransactionType"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo TTList 
         Bindings        =   "FormAdminWizard.frx":43D60
         DataField       =   "TransactionTypeID"
         DataSource      =   "TTListConn"
         Height          =   360
         Left            =   1080
         TabIndex        =   62
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ListField       =   "TransactionTypeName"
         BoundColumn     =   "TransactionTypeID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Choose A Transaction To Edit"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   7095
      End
      Begin VB.Image Imgfrm10 
         Height          =   6555
         Left            =   120
         Picture         =   "FormAdminWizard.frx":43D79
         Top             =   120
         Width           =   11385
      End
   End
   Begin VB.Frame frm9 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   48
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":48D0C
         DownPicture     =   "FormAdminWizard.frx":48F9A
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":49A7A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdFinish9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":4A4FA
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":4A83B
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data confirm9 
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Bank"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtBankName 
         BackColor       =   &H00FFC0C0&
         DataField       =   "BankName"
         DataSource      =   "confirm9"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   52
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Taype The Bank Name That You Want"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   50
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Image Imgfrm9 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":4B003
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm8 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   49
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         DataField       =   "BalanceFlag"
         DataSource      =   "Confirm8"
         Height          =   255
         Left            =   1440
         TabIndex        =   74
         Top             =   3120
         Width           =   210
      End
      Begin VB.CommandButton cmdFinish8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":4FF96
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":502D7
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":50A9F
         DownPicture     =   "FormAdminWizard.frx":50D2D
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":5180D
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5640
         Width           =   2295
      End
      Begin VB.TextBox txtTTName 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TransactionTypeName"
         DataSource      =   "Confirm8"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   56
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Data Confirm8 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TransactionType"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         DataField       =   "DebitFlag"
         DataSource      =   "Confirm8"
         Height          =   255
         Left            =   1440
         TabIndex        =   72
         Top             =   2520
         Width           =   210
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FormAdminWizard.frx":5228D
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
         Height          =   1815
         Left            =   1440
         TabIndex        =   75
         Top             =   3720
         Width           =   6495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Are you receiving(check) or paying(Uncheck)?"
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
         Left            =   1920
         TabIndex        =   73
         Top             =   2520
         Width           =   6495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Is it a balance(Check) or stock transaction(Uncheck)?"
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
         Left            =   1920
         TabIndex        =   59
         Top             =   3120
         Width           =   8055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Fill The Fields"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   8655
      End
      Begin VB.Image Imgfrm8 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":523D9
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":5736C
         DownPicture     =   "FormAdminWizard.frx":575FA
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":580DA
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":58B5A
         DownPicture     =   "FormAdminWizard.frx":58EDB
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":59971
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data adoCompany 
         Caption         =   "adoCompany"
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   5160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Company"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstCompany 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lblListofCompanies 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Companies"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label lblCompanyID 
         Height          =   615
         Left            =   5160
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblQ8 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Choose The Appropriate Company"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   7815
      End
      Begin VB.Image imgfrm6 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":5A36B
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton CmdNext5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":5F2FE
         DownPicture     =   "FormAdminWizard.frx":5F67F
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":60115
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormAdminWizard.frx":60B0F
         DownPicture     =   "FormAdminWizard.frx":60D9D
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormAdminWizard.frx":6187D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Data BankListConn 
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   615
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Bank"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo BankList 
         Bindings        =   "FormAdminWizard.frx":622FD
         DataField       =   "BankID"
         DataSource      =   "BankListConn"
         Height          =   360
         Left            =   600
         TabIndex        =   20
         Top             =   1680
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ListField       =   "BankName"
         BoundColumn     =   "BankID"
         Text            =   "DBCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Choose A Bank To Edit"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   5670
      End
      Begin VB.Image Img5 
         Height          =   6555
         Left            =   0
         Picture         =   "FormAdminWizard.frx":62318
         Top             =   0
         Width           =   11385
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
      TabIndex        =   77
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3120
      Picture         =   "FormAdminWizard.frx":672AB
      Top             =   360
      Width           =   450
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
      Left            =   4800
      TabIndex        =   65
      Top             =   960
      Width           =   4410
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
      Left            =   7560
      TabIndex        =   26
      Top             =   1560
      Width           =   3975
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
      Left            =   3240
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Go To Tadawul"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      MouseIcon       =   "FormAdminWizard.frx":678E1
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   11640
      MouseIcon       =   "FormAdminWizard.frx":67A33
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ImgFormBack 
      Height          =   9390
      Left            =   0
      Picture         =   "FormAdminWizard.frx":67B85
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "FormAdminWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IgnorarListaClick, IgnorarListaClickComp, IgnorarListaClickTT As Boolean
Dim Multimedia As New Mmedia
Dim counter As Integer, counter1 As Integer, itsdebit As Integer, itsdebit1 As Integer
Dim PassedFrom0 As Integer, PassedFrom1 As Integer, PassedFrom1GoBack As Integer, PassedFrom2 As Integer, PassedFrom2GoBack As Integer, PassedFrom2bankid As Integer, PassedFrom3 As Integer, PassedFrom3GoBack As Integer, PassedFrom4 As Integer, PassedFrom4GoBack As Integer, PassedFrom5 As Integer, PassedFrom5Goback As Integer, PassedFrom6 As Integer, PassedFrom6GoBack As Integer, PassedFrom7 As Integer, PassedFrom7GoBack As Integer, PassedFrom8 As Integer, PassedFrom8GoBack As Integer, PassedFrom9 As Integer, PassedFrom9GoBack As Integer, PassedFrom10 As Integer, PassedFrom10GoBack As Integer, PassedFrom11 As Integer, PassedFrom11GoBack As Integer, PassedFrom12 As Integer, PassedFrom12GoBack As Integer, PassedFrom13 As Integer, PassedFrom13GoBack As Integer, PassedFrom14 As Integer, PassedFrom14GoBack As Integer, PassedFrom15 As Long, PassedFrom15GoBack As Integer, PassedFrom16 As Integer, PassedFrom16GoBack As Integer, PassedFrom23accountid As Integer, PassedFrom23GoBack As Integer
Dim PassedFrom23bankid As Integer, PassedFrom14debit As Integer, PassedFrom14Balance As Boolean, PassedFrom9numberofstocks As Long, PassedFrom9priceperstock As Long, PassedFrom14TransactionTypeName As String, PassedFrom23AccountNumber As String, PassedFrom8CompanyName As String, ttblclicked As Integer, ttslclicked As Integer

Private Sub backo(goback As Integer)
If goback = 0 Then frm0.Visible = True
If goback = 1 Then frm1.Visible = True
If goback = 2 Then frm2.Visible = True
'If goback = 23 Then frm23.Visible = True
If goback = 4 Then frm4.Visible = True
If goback = 5 Then frm5.Visible = True
If goback = 6 Then frm6.Visible = True
'If goback = 7 Then frm7.Visible = True
If goback = 8 Then frm8.Visible = True
If goback = 9 Then frm9.Visible = True
If goback = 10 Then frm10.Visible = True
'If goback = 11 Then frm11.Visible = True
'If goback = 12 Then frm12.Visible = True
'If goback = 13 Then frm13.Visible = True
'If goback = 14 Then frm14.Visible = True
'If goback = 15 Then frm15.Visible = True
'If goback = 16 Then frm16.Visible = True
'If goback = 17 Then frm17.Visible = True
End Sub


Private Sub BankList_Click(Area As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub





Private Sub CityList_Change()
'txtCityIDField.Text = Val(CityList.BoundText)
End Sub

Private Sub CityList_KeyPress(KeyAscii As Integer)
'txtCityIDField.Text = Val(CityList.BoundText)
End Sub

Private Sub cmdBack1_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm1.Visible = False
Call backo(PassedFrom1GoBack)
End Sub

Private Sub cmdBack10_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm10.Visible = False
Call backo(PassedFrom10GoBack)
End Sub


Private Sub cmdBack2_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm2.Visible = False
Confirm2.Recordset.CancelUpdate
Call backo(PassedFrom2GoBack)
End Sub
Private Sub cmdBack3_Click()
Multimedia.mmOpen "Back.wav"
Multimedia.mmPlay
frm3.Visible = False
Call backo(PassedFrom3GoBack)
Confirm3.Recordset.CancelUpdate
End Sub

Private Sub cmdBack4_Click()
Multimedia.mmOpen "Back.wav"
Multimedia.mmPlay
frm4.Visible = False
Call backo(PassedFrom4GoBack)
End Sub

Private Sub cmdBack5_Click()
Multimedia.mmOpen "Back.wav"
Multimedia.mmPlay

frm5.Visible = False
Call backo(PassedFrom5Goback)
End Sub

Private Sub cmdBack6_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm6.Visible = False
'Confirm3.Recordset.CancelUpdate
Call backo(PassedFrom6GoBack)
End Sub


Private Sub cmdBack8_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm8.Visible = False
Confirm8.Recordset.CancelUpdate
Call backo(PassedFrom8GoBack)
End Sub

Private Sub cmdBack9_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm9.Visible = False
confirm9.Recordset.CancelUpdate
Call backo(PassedFrom9GoBack)
End Sub


Private Sub cmdFinish2_Click()
Confirm2.Recordset.Update
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
End Sub

Private Sub cmdFinish3_Click()
Confirm3.Recordset.Update
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
End Sub

Private Sub cmdFinish6_Click()
Confirm2.Recordset.Update
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
End Sub

Private Sub cmdFinish8_Click()
Confirm8.Recordset.Update
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
End Sub

Private Sub cmdFinish9_Click()
confirm9.Recordset.Update
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
End Sub

Private Sub cmdNext0_Click()
If op0(0).Value = False And op0(1).Value = False And op0(2).Value = False And op0(3).Value = False Then
MsgBox ("Choose first")
Else
Multimedia.mmOpen "Back.wav"
Multimedia.mmPlay
If op0(0) = True Then PassedFrom0 = 1
If op0(1) = True Then PassedFrom0 = 2
If op0(2) = True Then PassedFrom0 = 3
If op0(3) = True Then PassedFrom0 = 4

frm0.Visible = False
frm1.Visible = True
PassedFrom1GoBack = 0


End If
End Sub

Private Sub cmdNext1_Click()
If Op1(0).Value = False And Op1(1).Value = False Then
MsgBox ("Choose first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay

If Op1(0) = True Then PassedFrom1 = 1
If Op1(1) = True Then PassedFrom1 = 2


If PassedFrom0 = 1 And PassedFrom1 = 1 Then
frm1.Visible = False
frm2.Visible = True
PassedFrom2GoBack = 1
Confirm2.Recordset.AddNew
End If
If PassedFrom0 = 1 And PassedFrom1 = 2 Then
frm1.Visible = False
frm4.Visible = True
PassedFrom4GoBack = 1
End If
If PassedFrom0 = 2 And PassedFrom1 = 2 Then
frm1.Visible = False
frm5.Visible = True
PassedFrom5Goback = 1
End If
If PassedFrom0 = 2 And PassedFrom1 = 1 Then
frm1.Visible = False
frm9.Visible = True
PassedFrom9GoBack = 1
confirm9.Recordset.AddNew
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 Then
frm1.Visible = False
frm3.Visible = True
PassedFrom3GoBack = 1
Confirm3.Recordset.AddNew
End If
If PassedFrom0 = 3 And PassedFrom1 = 2 Then
frm1.Visible = False
frm6.Visible = True
PassedFrom6GoBack = 1
End If
If PassedFrom0 = 4 And PassedFrom1 = 1 Then
frm1.Visible = False
frm8.Visible = True
PassedFrom8GoBack = 1
Confirm8.Recordset.AddNew
End If
If PassedFrom0 = 4 And PassedFrom1 = 2 Then
frm1.Visible = False
frm10.Visible = True
PassedFrom10GoBack = 1

End If
End If
End Sub

Private Sub cmdNext10_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm10.Visible = False
frm8.Visible = True
PassedFrom8GoBack = 10
Confirm8.RecordSource = "select * From TransactionType where TransactionTypeID = " & TTList.BoundText
Confirm8.Refresh
Confirm8.Recordset.Edit
End Sub



Private Sub cmdNext4_Click()
Dim h As Integer
On Error GoTo DelErr
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
Confirm2.RecordSource = "Select * from Users where UserID=" & UserList.BoundText
Confirm2.Refresh
h = MsgBox("Would you like to delete?", vbYesNo)
If h = 6 Then
Confirm2.Recordset.Delete

Else

Confirm2.RecordSource = "Select * from Users where UserID=" & UserList.BoundText
Confirm2.Refresh
Confirm2.Recordset.Edit
frm4.Visible = False
frm2.Visible = True
PassedFrom4GoBack = 2
End If
DelErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub

Private Sub cmdNext5_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
confirm9.RecordSource = "Select * from Bank where BankID=" & BankList.BoundText
confirm9.Refresh
confirm9.Recordset.Edit

frm5.Visible = False
frm9.Visible = True
PassedFrom9GoBack = 5

End Sub

Private Sub cmdNext6_Click()
Dim h
On Error GoTo DelErr
Multimedia.mmOpen "next.wav"
        Multimedia.mmPlay
Confirm3.RecordSource = "select * from company where companyID = " & lblCompanyID.Caption
Confirm3.Refresh
h = MsgBox("Would you like to delete? This company is in some transactions and this will delete those transactions too, are you sure?", vbYesNo)
If h = 6 Then
Confirm3.Recordset.Delete
Unload FormAdminWizard
Load FormAdminWizard
FormAdminWizard.Show
Else
Confirm3.Recordset.Edit
frm6.Visible = False
frm3.Visible = True
PassedFrom3GoBack = 6
End If
DelErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub

Private Sub Image1_Click()
Multimedia.mmOpen "next.wav"
        Multimedia.mmPlay
ViewStocksPage.Show
End Sub

Private Sub Label14_Click()
End Sub

Private Sub Label17_Click()
Fancyfrm.Show
End Sub

Private Sub lblExit_Click()
Multimedia.mmOpen "expand.wav"
        Multimedia.mmPlay
Unload MainForm
Unload FormAdminWizard
Unload FormUserWizard
End Sub
Private Sub Form_Load()
    '======
    'Fills The Company list
    adoCompany.RecordSource = "SELECT * FROM Company where CompanyID <> 5"
    adoCompany.Refresh
    While Not adoCompany.Recordset.EOF
        lstCompany.AddItem adoCompany.Recordset("CompanyName")
        adoCompany.Recordset.MoveNext
    Wend
    '======
FormAdminWizard.Top = (Screen.Height - FormAdminWizard.Height) / 2
FormAdminWizard.Left = (Screen.Width - FormAdminWizard.Width) / 2
lblTime.Caption = Format(Now, "Long Date")
End Sub

Private Sub lstCompany_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
If Not IgnorarListaClickComp Then
            adoCompany.RecordSource = "SELECT * FROM Company WHERE CompanyName='" & lstCompany.Text & "'"
            adoCompany.Refresh
            lblCompanyID.Caption = adoCompany.Recordset.Fields("CompanyID")
    End If
End Sub

Private Sub lstTType_Click()

Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
End Sub

Private Sub op0_Click(Index As Integer)
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
End Sub

Private Sub Op1_Click(Index As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub

Private Sub op5_Click(Index As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub

Private Sub Op6_Click(Index As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub

Private Sub Op7_Click(Index As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub


Private Sub txtPhoneField_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub
