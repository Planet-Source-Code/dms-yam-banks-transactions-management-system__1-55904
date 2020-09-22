VERSION 5.00
Object = "{56183D41-61D7-11D6-BD5C-0010A4F59E39}#23.0#0"; "XPCalendar.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F49365FC-E8A5-4E38-9DBC-DAA7D889B8A3}#1.6#0"; "pbxpbutton.ocx"
Begin VB.Form FormUserWizard 
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
   Begin VB.Frame frm13 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   98
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":0000
         DownPicture     =   "FormUserWizard.frx":028E
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":0D6E
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdFinish13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":17EE
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":1B2F
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtNS 
         DataField       =   "NumberOfStocks"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   7920
         TabIndex        =   141
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPS 
         DataField       =   "PricePerStock"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   6720
         TabIndex        =   140
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox c1 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   139
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox d1 
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
         Height          =   375
         Left            =   6000
         TabIndex        =   138
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtEditAmount1 
         BackColor       =   &H00FFC0C0&
         DataField       =   "Amount"
         DataSource      =   "TListConn1"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtEditDate11 
         BackColor       =   &H00FFC0C0&
         DataField       =   "DateOfTransaction"
         DataSource      =   "TListConn1"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Data TTConn1 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Query3"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data AccConn1 
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Account"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtEditNS 
         BackColor       =   &H00FFC0C0&
         DataField       =   "NumberOfStocks"
         DataSource      =   "TListConn1"
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
         Left            =   3000
         TabIndex        =   106
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtEditTT1 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TransactionTypeName"
         DataSource      =   "TListConn1"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   3840
         Width           =   3855
      End
      Begin VB.TextBox txtEditAcc1 
         BackColor       =   &H00FFC0C0&
         DataField       =   "AccountNumber"
         DataSource      =   "TListConn1"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Data TListConn1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4560
         MouseIcon       =   "FormUserWizard.frx":22F7
         MousePointer    =   99  'Custom
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Query2"
         Top             =   5520
         Width           =   1150
      End
      Begin VB.TextBox txtEditPS 
         BackColor       =   &H00FFC0C0&
         DataField       =   "PricePerStock"
         DataSource      =   "TListConn1"
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
         Left            =   3000
         TabIndex        =   107
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Data Data31 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   8880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Transaction"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox dd1 
         DataField       =   "Debit"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   6720
         TabIndex        =   113
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox acc1 
         DataField       =   "AccountID"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   6720
         TabIndex        =   112
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox cc1 
         DataField       =   "Credit"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   7920
         TabIndex        =   100
         Top             =   5640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox tt1 
         DataField       =   "TransactionTypeID"
         DataSource      =   "Data31"
         Height          =   285
         Left            =   7920
         TabIndex        =   99
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Data Data41 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TransactionType"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin PB_XP_Button.PBXPButton PBXPButton11 
         Height          =   495
         Left            =   5280
         TabIndex        =   118
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Search"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin XP_Calendar.XPCalendar TSDate1 
         Height          =   375
         Left            =   2880
         TabIndex        =   117
         Top             =   2040
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         BorderColorOver =   0
         BorderColorDown =   0
         TextMaxLength   =   0
         CalendarBackSelectedG2=   16777215
         Value           =   37984
         MousePointer    =   99
         MouseIcon       =   "FormUserWizard.frx":2449
         TodayForeColor  =   16777215
         TodayFontName   =   "MS Sans Serif"
         TodayFontSize   =   8.25
         TodayFontBold   =   0   'False
         TodayFontItalic =   0   'False
         TodayPictureWidth=   16
         TodayPictureHeight=   16
         TodayPictureSize=   0
         TodayOriginalPicSizeW=   0
         TodayOriginalPicSizeH=   0
         GridLineColor   =   0
         CalendarBdHighlightColour=   0
         CalendarBdHighlightDKColour=   0
         CalendarBdShadowColour=   0
         CalendarBdShadowDKColour=   0
      End
      Begin MSDBCtls.DBCombo AccSList1 
         Bindings        =   "FormUserWizard.frx":25AB
         DataField       =   "AccountID"
         DataSource      =   "AccConn"
         Height          =   360
         Left            =   2880
         TabIndex        =   115
         Top             =   1560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
         BackColor       =   16761024
         ListField       =   "AccountNumber"
         BoundColumn     =   "AccountID"
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
         MouseIcon       =   "FormUserWizard.frx":25C2
      End
      Begin MSDBCtls.DBCombo TTSList1 
         Bindings        =   "FormUserWizard.frx":2724
         DataField       =   "TransactionTypeID"
         DataSource      =   "TTConn1"
         Height          =   360
         Left            =   2880
         TabIndex        =   114
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
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
         MouseIcon       =   "FormUserWizard.frx":273A
      End
      Begin PB_XP_Button.PBXPButton PBXPButton21 
         Height          =   495
         Left            =   7320
         TabIndex        =   116
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Reset"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton btnEdit1 
         Height          =   495
         Left            =   8040
         TabIndex        =   109
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Edit"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin MSDBCtls.DBCombo TTEditList1 
         Bindings        =   "FormUserWizard.frx":289C
         DataField       =   "TransactionTypeID"
         DataSource      =   "TTConn1"
         Height          =   360
         Left            =   3000
         TabIndex        =   103
         Top             =   3840
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
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
         MouseIcon       =   "FormUserWizard.frx":28B2
      End
      Begin MSDBCtls.DBCombo AccEditList1 
         Bindings        =   "FormUserWizard.frx":2A14
         DataField       =   "AccountID"
         DataSource      =   "AccConn1"
         Height          =   360
         Left            =   3000
         TabIndex        =   105
         Top             =   4320
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
         BackColor       =   16761024
         ListField       =   "AccountNumber"
         BoundColumn     =   "AccountID"
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
         MouseIcon       =   "FormUserWizard.frx":2A2B
      End
      Begin XP_Calendar.XPCalendar txtEditDate111 
         Height          =   375
         Left            =   3000
         TabIndex        =   119
         Top             =   3360
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         BorderColorOver =   0
         BorderColorDown =   0
         TextMaxLength   =   0
         CalendarBackSelectedG2=   16777215
         Value           =   37984
         MousePointer    =   99
         MouseIcon       =   "FormUserWizard.frx":2B8D
         TodayForeColor  =   16777215
         TodayFontName   =   "MS Sans Serif"
         TodayFontSize   =   8.25
         TodayFontBold   =   0   'False
         TodayFontItalic =   0   'False
         TodayPictureWidth=   16
         TodayPictureHeight=   16
         TodayPictureSize=   0
         TodayOriginalPicSizeW=   0
         TodayOriginalPicSizeH=   0
         GridLineColor   =   0
         CalendarBdHighlightColour=   0
         CalendarBdHighlightDKColour=   0
         CalendarBdShadowColour=   0
         CalendarBdShadowDKColour=   0
      End
      Begin PB_XP_Button.PBXPButton btnCancel1 
         Height          =   495
         Left            =   8040
         TabIndex        =   120
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Cancel"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton btnSave1 
         Height          =   495
         Left            =   8040
         TabIndex        =   121
         Top             =   4800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Save"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin VB.Label lblTID 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Index           =   3
         Left            =   360
         TabIndex        =   148
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "All Fields Are Required"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   145
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "All Fields Are Required"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   144
         Top             =   6240
         Width           =   3255
      End
      Begin VB.Label Label8 
         DataField       =   "Debit"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   6360
         TabIndex        =   143
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         DataField       =   "Debit"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   7680
         TabIndex        =   142
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTID 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Per Stock"
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
         Index           =   2
         Left            =   360
         TabIndex        =   137
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label lblQ12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Which Transaction Would You Like To Modify?"
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
         Left            =   360
         TabIndex        =   136
         Top             =   360
         Width           =   8400
      End
      Begin VB.Label lblS1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Index           =   1
         Left            =   360
         TabIndex        =   135
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblS2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Index           =   1
         Left            =   360
         TabIndex        =   134
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblS3 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Index           =   1
         Left            =   360
         TabIndex        =   133
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   360
         X2              =   10680
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblTID 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Stocks"
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
         Index           =   1
         Left            =   360
         TabIndex        =   132
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tranaction Type"
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
         Index           =   1
         Left            =   360
         TabIndex        =   131
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Index           =   1
         Left            =   360
         TabIndex        =   130
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Index           =   1
         Left            =   360
         TabIndex        =   129
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label RecordCount1 
         BackStyle       =   0  'Transparent
         Caption         =   " "
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
         Left            =   5880
         TabIndex        =   128
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label RC1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   5400
         TabIndex        =   127
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label lbldd1 
         DataField       =   "Debit"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   6360
         TabIndex        =   126
         Top             =   5640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblcc1 
         DataField       =   "Credit"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   7680
         TabIndex        =   125
         Top             =   5640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblAcc1 
         DataField       =   "AccountID"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   6360
         TabIndex        =   124
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbltt1 
         DataField       =   "TransactionTypeID"
         DataSource      =   "Data31"
         Height          =   255
         Left            =   7680
         TabIndex        =   123
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label tid1 
         DataField       =   "TransactionsID"
         DataSource      =   "TListConn1"
         Height          =   135
         Left            =   2640
         TabIndex        =   122
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Img13 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":2CEF
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm12 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   61
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":7C82
         DownPicture     =   "FormUserWizard.frx":7F10
         Height          =   735
         Left            =   6720
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":89F0
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CommandButton cmdFinish12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":9470
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":97B1
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TransactionType"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox tt 
         DataField       =   "TransactionTypeID"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   7920
         TabIndex        =   91
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox cc 
         DataField       =   "Credit"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   7920
         TabIndex        =   90
         Top             =   5520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox acc 
         DataField       =   "AccountID"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   6720
         TabIndex        =   89
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox dd 
         DataField       =   "Debit"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   6720
         TabIndex        =   88
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   8880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Transaction"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox c 
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
         Height          =   375
         Left            =   3000
         TabIndex        =   86
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox d 
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
         Height          =   375
         Left            =   3000
         TabIndex        =   85
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin PB_XP_Button.PBXPButton PBXPButton1 
         Height          =   495
         Left            =   5280
         TabIndex        =   77
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Search"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin VB.Data TListConn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3000
         MouseIcon       =   "FormUserWizard.frx":9F79
         MousePointer    =   99  'Custom
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Query1"
         Top             =   6000
         Width           =   1150
      End
      Begin VB.TextBox txtEditDate 
         BackColor       =   &H00FFC0C0&
         DataField       =   "DateOfTransaction"
         DataSource      =   "TListConn"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtEditAcc 
         BackColor       =   &H00FFC0C0&
         DataField       =   "AccountNumber"
         DataSource      =   "TListConn"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtEditTT 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TransactionTypeName"
         DataSource      =   "TListConn"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox txtEditAmount 
         BackColor       =   &H00FFC0C0&
         DataField       =   "Amount"
         DataSource      =   "TListConn"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   5160
         Width           =   1215
      End
      Begin XP_Calendar.XPCalendar TSDate 
         Height          =   375
         Left            =   2880
         TabIndex        =   68
         Top             =   2040
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         BorderColorOver =   0
         BorderColorDown =   0
         TextMaxLength   =   0
         CalendarBackSelectedG2=   16777215
         Value           =   37984
         MousePointer    =   99
         MouseIcon       =   "FormUserWizard.frx":A0CB
         TodayForeColor  =   16777215
         TodayFontName   =   "MS Sans Serif"
         TodayFontSize   =   8.25
         TodayFontBold   =   0   'False
         TodayFontItalic =   0   'False
         TodayPictureWidth=   16
         TodayPictureHeight=   16
         TodayPictureSize=   0
         TodayOriginalPicSizeW=   0
         TodayOriginalPicSizeH=   0
         GridLineColor   =   0
         CalendarBdHighlightColour=   0
         CalendarBdHighlightDKColour=   0
         CalendarBdShadowColour=   0
         CalendarBdShadowDKColour=   0
      End
      Begin VB.Data AccConn 
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Account"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo AccSList 
         Bindings        =   "FormUserWizard.frx":A22D
         DataField       =   "AccountID"
         DataSource      =   "AccConn"
         Height          =   360
         Left            =   2880
         TabIndex        =   65
         Top             =   1560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
         BackColor       =   16761024
         ListField       =   "AccountNumber"
         BoundColumn     =   "AccountID"
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
         MouseIcon       =   "FormUserWizard.frx":A243
      End
      Begin MSDBCtls.DBCombo TTSList 
         Bindings        =   "FormUserWizard.frx":A3A5
         DataField       =   "TransactionTypeID"
         DataSource      =   "TTConn"
         Height          =   360
         Left            =   2880
         TabIndex        =   64
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
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
         MouseIcon       =   "FormUserWizard.frx":A3BA
      End
      Begin VB.Data TTConn 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Query4"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1140
      End
      Begin PB_XP_Button.PBXPButton PBXPButton2 
         Height          =   495
         Left            =   7320
         TabIndex        =   80
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Reset"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton btnEdit 
         Height          =   495
         Left            =   8040
         TabIndex        =   81
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Edit"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin MSDBCtls.DBCombo TTEditList 
         Bindings        =   "FormUserWizard.frx":A51C
         DataField       =   "TransactionTypeID"
         DataSource      =   "TTConn"
         Height          =   360
         Left            =   3000
         TabIndex        =   82
         Top             =   3960
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
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
         MouseIcon       =   "FormUserWizard.frx":A531
      End
      Begin MSDBCtls.DBCombo AccEditList 
         Bindings        =   "FormUserWizard.frx":A693
         DataField       =   "AccountID"
         DataSource      =   "AccConn"
         Height          =   360
         Left            =   3000
         TabIndex        =   83
         Top             =   4560
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   635
         _Version        =   393216
         MousePointer    =   99
         BackColor       =   16761024
         ListField       =   "AccountNumber"
         BoundColumn     =   "AccountID"
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
         MouseIcon       =   "FormUserWizard.frx":A6A9
      End
      Begin XP_Calendar.XPCalendar txtEditDate1 
         Height          =   375
         Left            =   3000
         TabIndex        =   84
         Top             =   3360
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         BorderColorOver =   0
         BorderColorDown =   0
         TextMaxLength   =   0
         CalendarBackSelectedG2=   16777215
         Value           =   37984
         MousePointer    =   99
         MouseIcon       =   "FormUserWizard.frx":A80B
         TodayForeColor  =   16777215
         TodayFontName   =   "MS Sans Serif"
         TodayFontSize   =   8.25
         TodayFontBold   =   0   'False
         TodayFontItalic =   0   'False
         TodayPictureWidth=   16
         TodayPictureHeight=   16
         TodayPictureSize=   0
         TodayOriginalPicSizeW=   0
         TodayOriginalPicSizeH=   0
         GridLineColor   =   0
         CalendarBdHighlightColour=   0
         CalendarBdHighlightDKColour=   0
         CalendarBdShadowColour=   0
         CalendarBdShadowDKColour=   0
      End
      Begin PB_XP_Button.PBXPButton btnCancel 
         Height          =   495
         Left            =   8040
         TabIndex        =   87
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Cancel"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton btnSave 
         Height          =   495
         Left            =   8040
         TabIndex        =   97
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Save"
         BorderColor     =   0
         BorderColorOver =   6956042
         BorderColorDown =   6956042
         BackColor       =   16761024
         BackColorOver   =   16761024
         BackColorDown   =   11899525
         BackColorIcon   =   16761024
         BackColorIconOver=   16761024
         BackColorIconDown=   11899525
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         AlignCaption    =   2
         ShowFocus       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "All Fields Are Required"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   147
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "All Fields Are Required"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   146
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label tid 
         DataField       =   "TransactionsID"
         DataSource      =   "TListConn"
         Height          =   135
         Left            =   2640
         TabIndex        =   96
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbltt 
         DataField       =   "TransactionTypeID"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   7680
         TabIndex        =   95
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblAcc 
         DataField       =   "AccountID"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   6360
         TabIndex        =   94
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblcc 
         DataField       =   "Credit"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   7680
         TabIndex        =   93
         Top             =   5520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbldd 
         DataField       =   "Debit"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   6360
         TabIndex        =   92
         Top             =   5520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label RC 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   5400
         TabIndex        =   79
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label RecordCount 
         BackStyle       =   0  'Transparent
         Caption         =   " "
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
         Left            =   5880
         TabIndex        =   78
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Index           =   0
         Left            =   360
         TabIndex        =   75
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Index           =   0
         Left            =   360
         TabIndex        =   73
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tranaction Type"
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
         Index           =   0
         Left            =   360
         TabIndex        =   71
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label lblTID 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Index           =   0
         Left            =   360
         TabIndex        =   69
         Top             =   5160
         Width           =   975
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   10680
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblS3 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
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
         Index           =   0
         Left            =   360
         TabIndex        =   67
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblS2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Index           =   0
         Left            =   360
         TabIndex        =   66
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblS1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Index           =   0
         Left            =   360
         TabIndex        =   63
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblQ12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Which Transaction Would You Like To Modify?"
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
         Left            =   360
         TabIndex        =   62
         Top             =   360
         Width           =   8400
      End
      Begin VB.Image Imgfrm12 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":A96D
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm0 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   11295
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "View Reports"
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
         Index           =   3
         Left            =   360
         Picture         =   "FormUserWizard.frx":F900
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   4320
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Account Manager"
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
         Picture         =   "FormUserWizard.frx":14191
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2400
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transaction Manager"
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
         Picture         =   "FormUserWizard.frx":18A22
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3360
         Width           =   3855
      End
      Begin VB.OptionButton op0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Profile Manager"
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
         Picture         =   "FormUserWizard.frx":1D2B3
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdBack0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":21B44
         DownPicture     =   "FormUserWizard.frx":21DD2
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":228B2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":23332
         DownPicture     =   "FormUserWizard.frx":236B3
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":24149
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5760
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
         Picture         =   "FormUserWizard.frx":24B43
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm15 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   48
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":29AD6
         DownPicture     =   "FormUserWizard.frx":29D64
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":2A844
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":2B2C4
         DownPicture     =   "FormUserWizard.frx":2B645
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":2C0DB
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   840
         TabIndex        =   50
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label SR 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SR"
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
         TabIndex        =   51
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblQ15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "How Much is the Amount?"
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
         Left            =   360
         TabIndex        =   49
         Top             =   360
         Width           =   4500
      End
      Begin VB.Image Imgfrm15 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":2CAD5
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm14 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   42
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.Data Datas 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo TTBL 
         Bindings        =   "FormUserWizard.frx":31A68
         DataField       =   "transactiontypeID"
         DataSource      =   "TTConn"
         Height          =   360
         Left            =   840
         TabIndex        =   179
         Top             =   2640
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ListField       =   "TransactionTypeName"
         BoundColumn     =   "transactiontypeID"
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
      Begin VB.CommandButton cmdBack14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":31A9F
         DownPicture     =   "FormUserWizard.frx":31D2D
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":3280D
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":3328D
         DownPicture     =   "FormUserWizard.frx":3360E
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":340A4
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5760
         Width           =   2295
      End
      Begin MSDBCtls.DBCombo TTSL 
         Bindings        =   "FormUserWizard.frx":34A9E
         DataField       =   "transactiontypeID"
         DataSource      =   "TTConn1"
         Height          =   360
         Left            =   840
         TabIndex        =   180
         Top             =   2640
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ListField       =   "TransactionTypeName"
         BoundColumn     =   "transactiontypeID"
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
      Begin VB.Label lblTTBalanceFlag 
         DataField       =   "BalanceFlag"
         DataSource      =   "Datas"
         Height          =   375
         Left            =   5400
         TabIndex        =   155
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblTTDebitFlag 
         DataField       =   "DebitFlag"
         DataSource      =   "Datas"
         Height          =   375
         Left            =   5400
         TabIndex        =   154
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTTypeID 
         DataField       =   "TransactionTypeID"
         DataSource      =   "Datas"
         Height          =   615
         Left            =   5400
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Transaction Types"
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
         Left            =   840
         TabIndex        =   44
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label lblQ14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "What Information Would You Like To Modify?"
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
         Left            =   360
         TabIndex        =   43
         Top             =   360
         Width           =   8205
      End
      Begin VB.Image Imgfrm14 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":34AB4
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm11 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   181
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.Data Confirm11 
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Transaction"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdBack11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":39A47
         DownPicture     =   "FormUserWizard.frx":39CD5
         Height          =   735
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":3A7B5
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdFinish11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":3B235
         Height          =   735
         Left            =   2520
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":3B576
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Not Shown Here."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   208
         Top             =   5280
         Width           =   2265
      End
      Begin VB.Label final1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Simply, your transaction is"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   5880
         TabIndex        =   205
         Top             =   3600
         Width           =   5295
      End
      Begin VB.Label creditfield 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "Credit"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8520
         TabIndex        =   204
         Top             =   2280
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label debitfield 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "Debit"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5880
         TabIndex        =   203
         Top             =   2280
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label UserIDField1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "UserID"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         TabIndex        =   202
         Top             =   960
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label CommentsField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "Comments"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   201
         Top             =   5040
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label DateOfTransactionField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "DateOfTransaction"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   200
         Top             =   4800
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments        :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   199
         Top             =   5280
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date            :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   198
         Top             =   4800
         Width           =   2565
      End
      Begin VB.Label lblCompanyNameField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   197
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label TransactionTypeIDField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "TransactionTypeID"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         TabIndex        =   196
         Top             =   1440
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label AccountIDField1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "AccountID"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         TabIndex        =   195
         Top             =   1920
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label NumberOfStocksField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "NumberOfStocks"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   194
         Top             =   3720
         Width           =   2565
      End
      Begin VB.Label PricePerStockField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "PricePerStock"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   193
         Top             =   4320
         Width           =   2565
      End
      Begin VB.Label CompanyIDField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         DataField       =   "CompanyID"
         DataSource      =   "Confirm11"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6960
         TabIndex        =   192
         Top             =   3120
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label AmountField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   191
         Top             =   2520
         Width           =   2565
      End
      Begin VB.Label lblAccountNumberField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   190
         Top             =   1920
         Width           =   2565
      End
      Begin VB.Label lblTransactionTypeNameField 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   189
         Top             =   1440
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Company         :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   188
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Price Per Stock :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   187
         Top             =   4320
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Stocks:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   186
         Top             =   3720
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount          :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   185
         Top             =   2520
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number  :"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   184
         Top             =   1920
         Width           =   2565
      End
      Begin VB.Label lblViewFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   183
         Top             =   1440
         Width           =   2565
      End
      Begin VB.Label lblQ11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please, These Are The Options Chosen"
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
         Left            =   360
         TabIndex        =   182
         Top             =   360
         Width           =   7035
      End
      Begin VB.Image ImgFrm11 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":3BD3E
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm10 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   54
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdNext10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":40CD1
         DownPicture     =   "FormUserWizard.frx":41052
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":41AE8
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":424E2
         DownPicture     =   "FormUserWizard.frx":42770
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":43250
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtComment 
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
         Height          =   2895
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   2280
         Width           =   4455
      End
      Begin XP_Calendar.XPCalendar XPCalendar1 
         Height          =   495
         Left            =   840
         TabIndex        =   56
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         IconSizeWidth   =   30
         IconSizeHeight  =   30
         BorderColorDown =   16761024
         TextMaxLength   =   0
         CalendarBackSelectedG2=   16777215
         Value           =   37984
         TodayForeColor  =   16777215
         TodayFontName   =   "MS Sans Serif"
         TodayFontSize   =   8.25
         TodayFontBold   =   0   'False
         TodayFontItalic =   0   'False
         TodayPictureWidth=   16
         TodayPictureHeight=   16
         TodayPictureSize=   0
         TodayOriginalPicSizeW=   0
         TodayOriginalPicSizeH=   0
         GridLineColor   =   0
         CalendarBdHighlightColour=   0
         CalendarBdHighlightDKColour=   0
         CalendarBdShadowColour=   0
         CalendarBdShadowDKColour=   0
      End
      Begin VB.Label lblDate 
         Height          =   495
         Left            =   3240
         TabIndex        =   57
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblQ10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Fill In The Following"
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
         Left            =   360
         TabIndex        =   55
         Top             =   360
         Width           =   5670
      End
      Begin VB.Image Imgfrm10 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":43CD0
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm9 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtStock 
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
         Height          =   615
         Index           =   1
         Left            =   3000
         TabIndex        =   33
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton cmdBack9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Default         =   -1  'True
         DisabledPicture =   "FormUserWizard.frx":48C63
         DownPicture     =   "FormUserWizard.frx":48EF1
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":499D1
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":4A451
         DownPicture     =   "FormUserWizard.frx":4A7D2
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":4B268
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtStock 
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
         Height          =   615
         Index           =   0
         Left            =   3000
         TabIndex        =   32
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label lblNumberOfStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Per Stock"
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
         Index           =   1
         Left            =   480
         TabIndex        =   37
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblNumberOfStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Stocks"
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
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Specify the Following Amounts "
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
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   8055
      End
      Begin VB.Image Imgfrm9 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":4BC62
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm8 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   360
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":50BF5
         DownPicture     =   "FormUserWizard.frx":50E83
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":51963
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":523E3
         DownPicture     =   "FormUserWizard.frx":52764
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":531FA
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   5760
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
         TabIndex        =   26
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
         TabIndex        =   28
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label lblCompanyID 
         Height          =   615
         Left            =   5160
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   360
         Width           =   7815
      End
      Begin VB.Image imgfrm8 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":53BF4
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm7 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.OptionButton Op7 
         Caption         =   "Balance"
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
         Picture         =   "FormUserWizard.frx":58B87
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   1440
         Width           =   3855
      End
      Begin VB.OptionButton Op7 
         Caption         =   "Stock"
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
         Picture         =   "FormUserWizard.frx":5DB1A
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   2400
         Width           =   3855
      End
      Begin VB.CommandButton cmdNext7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":62AAD
         DownPicture     =   "FormUserWizard.frx":62E2E
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":638C4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton CmdBack7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":642BE
         DownPicture     =   "FormUserWizard.frx":6454C
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":6502C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label lblQ7 
         BackStyle       =   0  'Transparent
         Caption         =   "What Kind of Transaction Exactly?"
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
         TabIndex        =   13
         Top             =   360
         Width           =   8175
      End
      Begin VB.Image Img7 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":65AAC
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   209
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFC0C0&
         DataField       =   "UserName"
         DataSource      =   "Confirm6"
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
         TabIndex        =   218
         Top             =   960
         Width           =   3975
      End
      Begin VB.CommandButton cmdFinish6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":6AA3F
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":6AD80
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":6B548
         DownPicture     =   "FormUserWizard.frx":6B7D6
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":6C2B6
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Data Confirm6 
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
         DataSource      =   "Confirm6"
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
         TabIndex        =   222
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox txtEmailField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "Email"
         DataSource      =   "Confirm6"
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
         TabIndex        =   221
         Top             =   3480
         Width           =   4575
      End
      Begin VB.TextBox txtPasswordField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LoginPassword"
         DataSource      =   "Confirm6"
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
         TabIndex        =   220
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtLoginIDField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LoginID"
         DataSource      =   "Confirm6"
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
         TabIndex        =   219
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label lblFields 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name:"
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
         TabIndex        =   217
         Top             =   1080
         Width           =   1965
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
         TabIndex        =   214
         Top             =   1920
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   210
         Top             =   360
         Width           =   4305
      End
      Begin VB.Image ImgFrm6 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":6CD36
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "585"
      Height          =   6615
      Left            =   360
      TabIndex        =   38
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.OptionButton Op6 
         Caption         =   "Just Password"
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
         Left            =   840
         Picture         =   "FormUserWizard.frx":71CC9
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   2520
         Width           =   3855
      End
      Begin VB.OptionButton Op6 
         Caption         =   "All"
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
         Left            =   840
         Picture         =   "FormUserWizard.frx":76C5C
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CommandButton cmdNext5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":7BBEF
         DownPicture     =   "FormUserWizard.frx":7BF70
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":7CA06
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":7D400
         DownPicture     =   "FormUserWizard.frx":7D68E
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":7E16E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label lblQ5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "What Information Would You Like To Modify?"
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
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   8205
      End
      Begin VB.Image Imgfrm5 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":7EBEE
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   156
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":83B81
         DownPicture     =   "FormUserWizard.frx":83E0F
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":848EF
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Data Confirm4 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   "btms.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Account"
         Top             =   4200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdFinish4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":8536F
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":856B0
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox AccountNumberField 
         BackColor       =   &H00FFC0C0&
         DataField       =   "AccountNumber"
         DataSource      =   "Confirm4"
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
         Left            =   240
         TabIndex        =   158
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label AccountIDField 
         DataField       =   "AccountID"
         DataSource      =   "Confirm4"
         Height          =   495
         Left            =   2640
         TabIndex        =   169
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label BankIDField 
         DataField       =   "BankID"
         DataSource      =   "Confirm4"
         Height          =   495
         Left            =   240
         TabIndex        =   162
         Top             =   3600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label UserIDField 
         DataField       =   "UserID"
         DataSource      =   "Confirm4"
         Height          =   495
         Left            =   240
         TabIndex        =   161
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "You Can Write A Number Too"
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
         TabIndex        =   160
         Top             =   840
         Width           =   5085
      End
      Begin VB.Label lblQ4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Give Your Account A Name"
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
         TabIndex        =   157
         Top             =   360
         Width           =   6255
      End
      Begin VB.Image Img4 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":85E78
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   360
      TabIndex        =   164
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton CmdNext2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":8AE0B
         DownPicture     =   "FormUserWizard.frx":8B18C
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":8BC22
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":8C61C
         DownPicture     =   "FormUserWizard.frx":8C8AA
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":8D38A
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   5760
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
         Bindings        =   "FormUserWizard.frx":8DE0A
         DataField       =   "BankID"
         DataSource      =   "BankListConn"
         Height          =   360
         Left            =   600
         TabIndex        =   166
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
         Caption         =   "Please, Choose A Bank For Your Account"
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
         TabIndex        =   165
         Top             =   360
         Width           =   7425
      End
      Begin VB.Image Img2 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":8DE25
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm23 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdBack23 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":92DB8
         DownPicture     =   "FormUserWizard.frx":93046
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":93B26
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdNext23 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":945A6
         DownPicture     =   "FormUserWizard.frx":94927
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":953BD
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ListBox List1 
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
         Height          =   2460
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   2640
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   360
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "BTMS.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   615
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Bank"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Accounts For The Selected Bank"
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
         Left            =   720
         TabIndex        =   29
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblAutozisemsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Size"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   23
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblAutoSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                         Click Here For   Auto Sizing"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   4800
         TabIndex        =   22
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblBankID 
         DataField       =   "BankID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblAccountID 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblQ9 
         BackStyle       =   0  'Transparent
         Caption         =   "Please, Choose The Appropriate Account"
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
         TabIndex        =   17
         Top             =   360
         Width           =   7815
      End
      Begin VB.Image Img23 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":95DB7
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   360
      TabIndex        =   8
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
         Picture         =   "FormUserWizard.frx":9AD4A
         Style           =   1  'Graphical
         TabIndex        =   171
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
         Picture         =   "FormUserWizard.frx":9FCDD
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CommandButton cmdNext1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":A4C70
         DownPicture     =   "FormUserWizard.frx":A4FF1
         Height          =   735
         Left            =   8880
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":A5A87
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBack1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DisabledPicture =   "FormUserWizard.frx":A6481
         DownPicture     =   "FormUserWizard.frx":A670F
         Height          =   735
         Left            =   6480
         MaskColor       =   &H00000000&
         Picture         =   "FormUserWizard.frx":A71EF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5760
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
         TabIndex        =   9
         Top             =   360
         Width           =   7095
      End
      Begin VB.Image Img1 
         Height          =   6555
         Left            =   0
         Picture         =   "FormUserWizard.frx":A7C6F
         Top             =   0
         Width           =   11385
      End
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   11640
      MouseIcon       =   "FormUserWizard.frx":ACC02
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3120
      MouseIcon       =   "FormUserWizard.frx":ACD54
      MousePointer    =   99  'Custom
      Picture         =   "FormUserWizard.frx":ACEA6
      Top             =   360
      Width           =   450
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
      MouseIcon       =   "FormUserWizard.frx":AD4DC
      MousePointer    =   99  'Custom
      TabIndex        =   226
      Top             =   360
      Width           =   2655
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
      TabIndex        =   176
      Top             =   1560
      Width           =   3975
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
      TabIndex        =   175
      Top             =   960
      Width           =   4410
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
      TabIndex        =   174
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
      MouseIcon       =   "FormUserWizard.frx":AD62E
      MousePointer    =   99  'Custom
      TabIndex        =   149
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image ImgFormBack 
      Height          =   9390
      Left            =   0
      Picture         =   "FormUserWizard.frx":AD780
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "FormUserWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IgnorarListaClick, IgnorarListaClickComp, IgnorarListaClickTT As Boolean
Dim Multimedia As New Mmedia
Dim counter As Integer, counter1 As Integer, itsdebit As Boolean, itsdebit1 As Boolean
Dim PassedFrom0 As Integer, PassedFrom1 As Integer, PassedFrom1GoBack As Integer, PassedFrom2 As Integer, PassedFrom2GoBack As Integer, PassedFrom2bankid As Integer, PassedFrom3 As Integer, PassedFrom3GoBack As Integer, PassedFrom4 As Integer, PassedFrom4GoBack As Integer, PassedFrom5 As Integer, PassedFrom5Goback As Integer, PassedFrom6 As Integer, PassedFrom6GoBack As Integer, PassedFrom7 As Integer, PassedFrom7GoBack As Integer, PassedFrom8 As Integer, PassedFrom8GoBack As Integer, PassedFrom9 As Integer, PassedFrom9GoBack As Integer, PassedFrom10 As Integer, PassedFrom10GoBack As Integer, PassedFrom11 As Integer, PassedFrom11GoBack As Integer, PassedFrom12 As Integer, PassedFrom12GoBack As Integer, PassedFrom13 As Integer, PassedFrom13GoBack As Integer, PassedFrom14 As Integer, PassedFrom14GoBack As Integer, PassedFrom15 As Long, PassedFrom15GoBack As Integer, PassedFrom16 As Integer, PassedFrom16GoBack As Integer, PassedFrom23accountid As Integer, PassedFrom23GoBack As Integer
Dim PassedFrom23bankid As Integer, PassedFrom14debit As Boolean, PassedFrom14Balance As Boolean, PassedFrom9numberofstocks As Long, PassedFrom9priceperstock As Long, PassedFrom14TransactionTypeName As String, PassedFrom23AccountNumber As String, PassedFrom8CompanyName As String, ttblclicked As Integer, ttslclicked As Integer
Private Sub backo(goback As Integer)
If goback = 0 Then frm0.Visible = True
If goback = 1 Then frm1.Visible = True
If goback = 2 Then frm2.Visible = True
If goback = 23 Then frm23.Visible = True
If goback = 4 Then frm4.Visible = True
If goback = 5 Then frm5.Visible = True
'If goback = 6 Then frm6.Visible = True
If goback = 7 Then frm7.Visible = True
If goback = 8 Then frm8.Visible = True
If goback = 9 Then frm9.Visible = True
If goback = 10 Then frm10.Visible = True
'If goback = 11 Then frm11.Visible = True
If goback = 12 Then frm12.Visible = True
If goback = 13 Then frm13.Visible = True
If goback = 14 Then frm14.Visible = True
If goback = 15 Then frm15.Visible = True
'If goback = 16 Then frm16.Visible = True
'If goback = 17 Then frm17.Visible = True
End Sub
Private Sub backup()
Dim source, source1, desti, h, h2, path As String, path1 As String, path2 As String, recover As String, recover1 As String
path = "c:\BTMS backup utility (Backed up).log"
path2 = "c:\BTMS backup utility (Retrieved).log"
path1 = App.path & "\BFN.log"
h = MsgBox("Do you want to back up your data before you leave?", vbYesNo, "BTMS Backup Utility")
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
 Else
 Exit Sub
End If
End Sub
Private Sub BankList_Click(Area As Integer)
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub

Private Sub btnCancel_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
txtEditDate1.Visible = False
txtEditDate.Visible = True
btnCancel.Visible = False
btnEdit.Visible = True
btnSave.Visible = False
    TTEditList.Visible = False
    AccEditList.Visible = False
    txtEditAcc.Visible = True
    txtEditTT.Visible = True
    c.Visible = False
    d.Visible = False

cc.Text = lblcc.Caption
dd.Text = lbldd.Caption
acc.Text = lblAcc.Caption
tt.Text = lbltt.Caption
    Me.Refresh
End Sub

Private Sub btnCancel1_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
txtEditDate111.Visible = False
txtEditDate11.Visible = True
btnCancel1.Visible = False
btnEdit1.Visible = True
btnSave1.Visible = False
    TTEditList1.Visible = False
    AccEditList1.Visible = False
    txtEditAcc1.Visible = True
    txtEditTT1.Visible = True
    c1.Visible = False
    d1.Visible = False

cc1.Text = lblcc1.Caption
dd1.Text = lbldd1.Caption
acc1.Text = lblAcc1.Caption
tt1.Text = lbltt1.Caption
    Me.Refresh
End Sub

Private Sub btnEdit_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
If txtEditAmount.Text <> "" And txtEditAcc.Text <> "" And txtEditTT <> "" Then
txtEditDate1.Visible = True
txtEditDate.Visible = False
btnSave.Visible = True
TTEditList.Visible = True
AccEditList.Visible = True
txtEditAcc.Visible = False
txtEditTT.Visible = False
btnCancel.Visible = True
btnEdit.Visible = False
Data3.RecordSource = "SELECT TransactionsID,TransactionTypeID,Debit,credit,AccountID From [Transaction]Where TransactionsID =" & Val(tid.Caption)
Data3.Refresh
If txtEditAmount.Text = Val(dd.Text) Then
'd.Visible = True
d.Text = Val(txtEditAmount.Text)
c.Text = 0
c.Visible = False
itsdebit = True
TTEditList.Text = txtEditTT.Text
AccEditList.Text = txtEditAcc.Text
txtEditDate1.Value = txtEditDate.Text
Else
If txtEditAmount.Text = Val(cc.Text) Then
'c.Visible = True
c.Text = Val(txtEditAmount.Text)
d.Text = 0
d.Visible = False
itsdebit = False
TTEditList.Text = txtEditTT.Text
AccEditList.Text = txtEditAcc.Text
txtEditDate1.Value = txtEditDate.Text
End If
End If
Else
RecordCount.Caption = "No Records to Edit"
End If
End Sub

Private Sub btnEdit1_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
If txtEditAmount1.Text <> "" And txtEditAcc1.Text <> "" And txtEditTT1.Text <> "" And txtEditNS.Text <> "" And txtEditPS.Text <> "" Then
txtEditDate111.Visible = True
txtEditDate11.Visible = False
btnSave1.Visible = True
TTEditList1.Visible = True
AccEditList1.Visible = True
txtEditAcc1.Visible = False
txtEditTT1.Visible = False
btnCancel1.Visible = True
btnEdit1.Visible = False
Data31.RecordSource = "SELECT TransactionsID,TransactionTypeID,Debit,credit,AccountID ,Transaction.NumberOfStocks, Transaction.PricePerStock From Transaction Where TransactionsID =" & Val(tid1.Caption)
Data31.Refresh
If txtEditAmount1.Text = Val(dd1.Text) Then
'd1.Visible = True
'd1.Text = Val(txtEditNS.Text * txtEditPS.Text)
c1.Text = 0
c1.Visible = False
itsdebit1 = True
TTEditList1.Text = txtEditTT1.Text
AccEditList1.Text = txtEditAcc1.Text
txtEditDate111.Value = txtEditDate11.Text
Else
If txtEditAmount1.Text = Val(cc1.Text) Then
'c1.Visible = True
'c1.Text = Val(txtEditNS.Text * txtEditPS.Text)
d1.Text = 0
d1.Visible = False
itsdebit1 = False
TTEditList1.Text = txtEditTT1.Text
AccEditList1.Text = txtEditAcc1.Text
txtEditDate111.Value = txtEditDate11.Text
End If
End If
Else
RecordCount1.Caption = "No Records To Edit"
End If
End Sub

Private Sub btnSave1_Click()
Dim h
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
Data31.Recordset.Edit
txtNS.Text = txtEditNS.Text
txtPS.Text = txtEditPS.Text
Data41.Recordset.MoveFirst
If itsdebit1 = True Then d1.Text = Val(txtEditNS.Text * txtEditPS.Text)
If itsdebit1 = False Then c1.Text = Val(txtEditNS.Text * txtEditPS.Text)
While Not Data41.Recordset.EOF
If TTEditList1.Text = Data41.Recordset.Fields("TransactionTypeName") Then
    If itsdebit1 = Data41.Recordset.Fields("DebitFlag") Then
    'Data3.Recordset.Edit
    MsgBox ("since " & txtEditTT1.Text + " and " & TTEditList1.Text + " are of the same type, there is no change i'll just paste values and update the transaction type  to " & TTEditList1.Text)
    
    cc1.Text = Val(c1.Text)
    dd1.Text = Val(d1.Text)
    MsgBox ("pasted successfully")
    MsgBox ("now click the update button")
    
    'acc1.Text = Val(AccEditList1.BoundText)
    Data41.Recordset.MoveLast
    Else
    MsgBox ("ohhh, " & TTEditList1.Text + " is a different transaction type, we will do some changes")
        If itsdebit1 = True Then
        MsgBox ("" & TTEditList1.Text + " is credit but " & txtEditTT1.Text + " is debit, i am gonna switch values")
        'd1.Text = 0
        'd1.Text = Val(txtEditNS.Text * txtEditPS.Text)
        cc1.Text = Val(d1.Text)
        dd1.Text = 0
        MsgBox ("switched successfully")
        Else
        MsgBox ("" & TTEditList1.Text + " is debit but " & txtEditTT1.Text + " is credit, i am gonna switch values")
        'c1.Text = 0
        
        dd1.Text = Val(c1.Text)
        cc1.Text = 0
        MsgBox ("switched successfully")
        End If
        MsgBox ("now click the update button")
        'acc1.Text = Val(AccEditList1.BoundText)
        Data41.Recordset.MoveLast
    End If
End If
Data41.Recordset.MoveNext
Wend
        acc1.Text = Val(AccEditList1.BoundText)
        tt1.Text = Val(TTEditList1.BoundText)
h = MsgBox("are you sure", vbYesNo)
If h = 6 Then
    
    Data31.Recordset.Update
    TListConn1.Refresh
    MsgBox ("Seccess")
    btnCancel1_Click
    c1.Visible = False
    d1.Visible = False
    Me.Refresh
Else
btnCancel1_Click
'itsdebit1 = 0
d1.Text = 0
c1.Text = 0
'Me.Refresh
End If
End Sub

Private Sub CityList_Change()
'txtCityIDField.Text = Val(CityList.BoundText)
End Sub

Private Sub CityList_KeyPress(KeyAscii As Integer)
'txtCityIDField.Text = Val(CityList.BoundText)
End Sub



Private Sub c_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

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

Private Sub cmdBack11_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm11.Visible = False
Call backo(PassedFrom11GoBack)
End Sub

Private Sub cmdBack12_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm12.Visible = False
Call backo(PassedFrom12GoBack)
End Sub

Private Sub cmdBack13_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm13.Visible = False
Call backo(PassedFrom13GoBack)
End Sub

Private Sub cmdBack14_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm14.Visible = False
Call backo(PassedFrom14GoBack)
End Sub

Private Sub cmdBack15_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm15.Visible = False
Call backo(PassedFrom15GoBack)
End Sub

Private Sub cmdBack2_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm2.Visible = False
Call backo(PassedFrom2GoBack)
End Sub

Private Sub cmdBack23_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm23.Visible = False
Call backo(PassedFrom23GoBack)
End Sub

Private Sub cmdBack4_Click()
Multimedia.mmOpen "Back.wav"
Multimedia.mmPlay
frm4.Visible = False
Call backo(PassedFrom4GoBack)
Confirm4.Recordset.CancelUpdate
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
txtLoginIDField.Visible = True
txtEmailField.Visible = True
txtPhoneField.Visible = True

lblFields(0).Visible = True
lblFields(1).Visible = True
lblFields(2).Visible = True
lblFields(4).Visible = True
txtUserName.Visible = True
Confirm6.Recordset.CancelUpdate
frm6.Visible = False
Call backo(PassedFrom6GoBack)
End Sub

Private Sub CmdBack7_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm7.Visible = False
Call backo(PassedFrom7GoBack)
End Sub

Private Sub cmdBack8_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm8.Visible = False
Call backo(PassedFrom8GoBack)
End Sub

Private Sub cmdBack9_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
frm9.Visible = False
Call backo(PassedFrom9GoBack)
End Sub

Private Sub cmdFinish11_Click()
Confirm11.Recordset.Update
Unload FormUserWizard
Load FormUserWizard
FormUserWizard.Show
FormUserWizard.lblLoggedInUser.Caption = MainForm.lblLoggedInUser.Caption
End Sub

Private Sub cmdFinish12_Click()
Unload FormUserWizard
Load FormUserWizard
FormUserWizard.Show
FormUserWizard.lblLoggedInUser.Caption = MainForm.lblLoggedInUser.Caption
End Sub

Private Sub cmdFinish13_Click()
Unload FormUserWizard
Load FormUserWizard
FormUserWizard.Show
FormUserWizard.lblLoggedInUser.Caption = MainForm.lblLoggedInUser.Caption
End Sub

Private Sub cmdFinish4_Click()
If AccountNumberField.Text = "" Then
MsgBox ("Please enter a value first")
Else
If PassedFrom0 = 2 And PassedFrom1 = 1 Then
'Confirm4.Recordset.AddNew
Confirm4.Recordset.Update
End If
If PassedFrom0 = 2 And PassedFrom1 = 2 Then
Confirm4.Recordset.Edit
Confirm4.Recordset.Update
End If
Unload FormUserWizard
Load FormUserWizard
FormUserWizard.Show
FormUserWizard.lblLoggedInUser.Caption = MainForm.lblLoggedInUser.Caption
End If
End Sub

Private Sub cmdFinish6_Click()

Confirm6.Recordset.Update
Unload FormUserWizard
Load FormUserWizard
FormUserWizard.Show
FormUserWizard.lblLoggedInUser.Caption = MainForm.lblLoggedInUser.Caption
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
If PassedFrom0 = 1 Then
    frm0.Visible = False
    frm5.Visible = True
    PassedFrom5Goback = 0
End If
If PassedFrom0 = 4 Then
    If Reports.Adodc0.Recordset.RecordCount = 0 And Reports.Adodc1.Recordset.RecordCount = 0 And Reports.Adodc2.Recordset.RecordCount = 0 And Reports.Adodc3.Recordset.RecordCount = 0 And Reports.Adodc4.Recordset.RecordCount = 0 And Reports.Adodc5.Recordset.RecordCount = 0 Then
    MsgBox ("No reports available, please make some transactions first")
    Exit Sub
    Else
    Reports.Show
    Reports.lblLoggedInUser.Caption = lblLoggedInUser.Caption
    cmdNext0.Enabled = False
    Unload Me
    End If
End If
If PassedFrom0 = 2 Or PassedFrom0 = 3 Then
    frm0.Visible = False
    frm1.Visible = True
    PassedFrom1GoBack = 0
End If

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


If PassedFrom0 = 2 And PassedFrom1 = 1 Then
frm1.Visible = False
frm2.Visible = True
PassedFrom2GoBack = 1
End If
If PassedFrom0 = 2 And PassedFrom1 = 2 Then
frm1.Visible = False
frm23.Visible = True
PassedFrom23GoBack = 1
End If
If PassedFrom0 = 3 Then
frm1.Visible = False
frm7.Visible = True
PassedFrom7GoBack = 1
End If
End If
End Sub

Private Sub cmdNext10_Click()
Dim ttt As String
If lblDate.Caption = "" Then
MsgBox ("Fill all fields first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay

If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 2 Then
    Confirm11.Recordset.AddNew
    UserIDField1.Caption = MainForm.lblUserID.Caption
    AccountIDField1.Caption = PassedFrom23accountid
    lblAccountNumberField.Caption = PassedFrom23AccountNumber
    NumberOfStocksField.Caption = PassedFrom9numberofstocks
    'NumberOfStocksField.Caption = PassedFrom9numberofstocks
    PricePerStockField.Caption = PassedFrom9priceperstock
    'lblPricePerStockField.Caption = PassedFrom9priceperstock
    TransactionTypeIDField.Caption = PassedFrom14
    lblTransactionTypeNameField.Caption = PassedFrom14TransactionTypeName
    DateOfTransactionField.Caption = lblDate.Caption
    CommentsField.Caption = txtComment.Text
    lblCompanyNameField.Caption = PassedFrom8CompanyName
    CompanyIDField.Caption = PassedFrom8
    If PassedFrom14debit = True Then
    debitfield.Caption = PassedFrom9numberofstocks * PassedFrom9priceperstock
    creditfield.Caption = 0
    AmountField.Caption = Val(debitfield.Caption)
    End If
    If PassedFrom14debit = False Then
    creditfield.Caption = PassedFrom9numberofstocks * PassedFrom9priceperstock
    debitfield.Caption = 0
    AmountField.Caption = Val(creditfield.Caption)
    End If
    frm10.Visible = False
    frm11.Visible = True
    
    PassedFrom11GoBack = 10
    ttt = Left$(lblTransactionTypeNameField.Caption, 3)
    final1.Caption = "Simply, You Want To " & ttt + " " & NumberOfStocksField.Caption + " Stocks From " & lblCompanyNameField.Caption + vbNewLine + " With A Price of " & PricePerStockField.Caption + " Each. This Will Charge the" + vbNewLine + "" & lblAccountNumberField.Caption + " Account with " & AmountField.Caption + " Saudi Riyals." + vbNewLine + "Thanks For Using BTMS."
End If

If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 1 Then
    Confirm11.Recordset.AddNew
    UserIDField1.Caption = MainForm.lblUserID.Caption
    AccountIDField1.Caption = PassedFrom23accountid
    lblAccountNumberField.Caption = PassedFrom23AccountNumber
    TransactionTypeIDField.Caption = PassedFrom14
    lblTransactionTypeNameField.Caption = PassedFrom14TransactionTypeName
    DateOfTransactionField.Caption = lblDate.Caption
    CommentsField.Caption = txtComment.Text
    If PassedFrom14debit = True Then
    debitfield.Caption = PassedFrom15
    creditfield.Caption = 0
    AmountField.Caption = Val(debitfield.Caption)
    End If
    If PassedFrom14debit = False Then
    creditfield.Caption = PassedFrom15
    debitfield.Caption = 0
    AmountField.Caption = Val(creditfield.Caption)
    End If
    frm10.Visible = False
    frm11.Visible = True
    
    PassedFrom11GoBack = 10
    NumberOfStocksField.Visible = False
    PricePerStockField.Visible = False
    lblCompanyNameField.Visible = False
    lblViewFields(5).Visible = False
    lblViewFields(3).Visible = False
    lblViewFields(4).Visible = False
    CompanyIDField.Visible = False
    CompanyIDField.Caption = 5
    final1.Caption = "Simply, You Want To Make A " & lblTransactionTypeNameField.Caption + vbNewLine + ". This Will Charge the" + vbNewLine + "" & lblAccountNumberField.Caption + " Account with " & AmountField.Caption + " Saudi Riyals." + vbNewLine + "Thanks For Using BTMS."
End If
End If
End Sub

Private Sub cmdNext14_Click()
If lblTTypeID.Caption = "" Or lblTTDebitFlag.Caption = "" Or lblTTBalanceFlag.Caption = "" Or ttblclicked = 0 Then
MsgBox ("Choose a transaction type first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
PassedFrom14 = lblTTypeID.Caption
PassedFrom14debit = lblTTDebitFlag.Caption
PassedFrom14Balance = lblTTBalanceFlag.Caption
If TTBL.Visible = True Then
PassedFrom14TransactionTypeName = TTBL.Text
Else
PassedFrom14TransactionTypeName = TTSL.Text
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 1 Then
frm15.Visible = True
frm14.Visible = False
PassedFrom15GoBack = 14
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 2 Then
frm14.Visible = False
frm9.Visible = True
PassedFrom9GoBack = 14
End If
End If
ttblclicked = 1
End Sub

Private Sub cmdNext15_Click()
If txtAmount.Text = "" Then
MsgBox ("Please enter an amount first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
PassedFrom15 = Val(txtAmount.Text)
frm15.Visible = False
frm10.Visible = True
PassedFrom10GoBack = 15
End If
End Sub

Private Sub CmdNext2_Click()
If BankList.BoundText = "" Then
MsgBox ("Choose a Bank First")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
Confirm4.Recordset.AddNew
PassedFrom2bankid = BankList.BoundText
BankIDField.Caption = PassedFrom2bankid
UserIDField.Caption = MainForm.lblUserID.Caption
'AccountIDField.DataSource = Confirm4
'AccountIDField.DataField = AccountID
'AccountIDField.Caption = PaasedFrom2AccountID
frm2.Visible = False
frm4.Visible = True
PassedFrom4GoBack = 2
End If
End Sub

Private Sub cmdNext23_Click()
If lblAccountID.Caption = "" Then
MsgBox ("Choose an Account first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
PassedFrom23accountid = lblAccountID.Caption
PassedFrom23bankid = lblBankID.Caption
PassedFrom23AccountNumber = List1.Text
If PassedFrom0 = 2 And PassedFrom1 = 2 Then
    Confirm4.RecordSource = "select * from account where accountid = " & PassedFrom23accountid
    Confirm4.Refresh
    Confirm4.Recordset.Edit
    lblQ4.Caption = "Enter The New Name For Your Account"
    frm23.Visible = False
    frm4.Visible = True
    PassedFrom4GoBack = 23
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 1 Then
frm23.Visible = False
TTBL.Visible = True
TTSL.Visible = False
frm14.Visible = True
PassedFrom14GoBack = 23
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 2 Then
frm23.Visible = False
TTSL.Visible = True
TTBL.Visible = False
frm14.Visible = True
PassedFrom14GoBack = 23
End If
End If
End Sub

Private Sub cmdNext5_Click()
If Op6(0).Value = False And Op6(1).Value = False Then
MsgBox ("Choose one first")
Else
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
Confirm6.RecordSource = "Select * from Users where UserID=" & MainForm.lblUserID.Caption
Confirm6.Refresh
Confirm6.Recordset.Edit
If Op6(0).Value = True Then
frm5.Visible = False
frm6.Visible = True
PassedFrom6GoBack = 5

End If
If Op6(1).Value = True Then
txtLoginIDField.Visible = False
txtEmailField.Visible = False
txtPhoneField.Visible = False
lblFields(0).Visible = False
lblFields(1).Visible = False
lblFields(2).Visible = False
lblFields(4).Visible = False
txtUserName.Visible = False
frm5.Visible = False
frm6.Visible = True
PassedFrom6GoBack = 5
lblQ6.Caption = "Enter The New Password Please"
End If
End If
End Sub

Private Sub Command3_Click()
counter = 0
TListConn.Refresh
counter = TListConn.Recordset.RecordCount
RC.Caption = counter
RecordCount.Caption = "Record(s) Found."
End Sub
Private Sub cmdNext7_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
If Op7(0).Value = False And Op7(1).Value = False Then
MsgBox ("Choose first")
Else
If Op7(0) = True Then PassedFrom7 = 1
If Op7(1) = True Then PassedFrom7 = 2

If PassedFrom0 = 3 And PassedFrom1 = 2 And PassedFrom7 = 1 Then
frm12.Visible = True
frm7.Visible = False
PassedFrom12GoBack = 7
End If
If PassedFrom0 = 3 And PassedFrom1 = 2 And PassedFrom7 = 2 Then
frm13.Visible = True
frm7.Visible = False
PassedFrom13GoBack = 7
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 1 Then
frm23.Visible = True
frm7.Visible = False
PassedFrom23GoBack = 7
End If
If PassedFrom0 = 3 And PassedFrom1 = 1 And PassedFrom7 = 2 Then
frm8.Visible = True
frm7.Visible = False
PassedFrom8GoBack = 7
End If
End If
End Sub

Private Sub DBCombo1_Click(Area As Integer)

End Sub

Private Sub cmdNext8_Click()
Multimedia.mmOpen "Back.wav"
        Multimedia.mmPlay
PassedFrom8 = lblCompanyID.Caption
PassedFrom8CompanyName = lstCompany.Text
frm23.Visible = True
frm8.Visible = False
PassedFrom23GoBack = 8

End Sub

Private Sub cmdNext9_Click()
If txtStock(0).Text = "" Or txtStock(1) = "" Then
MsgBox ("fill fields first")
Else
Multimedia.mmOpen "back.wav"
        Multimedia.mmPlay
PassedFrom9numberofstocks = Val(txtStock(0).Text)
PassedFrom9priceperstock = Val(txtStock(1).Text)
frm10.Visible = True
frm9.Visible = False
PassedFrom10GoBack = 9
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub d_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub Image1_Click()
Multimedia.mmOpen "next.wav"
        Multimedia.mmPlay
ViewStocksPage.Show
End Sub

Private Sub Label10_Click()
frm13.Visible = False
frm12.Visible = True
End Sub

Private Sub Label17_Click()
Fancyfrm.Show
End Sub
Private Sub lblAutoSize_Click()
If lblAutozisemsg.Caption = "Normal Size" Then
AutoSizeComboBoxDropDown Combo1
lblAutozisemsg.Caption = "Auto Sized"
End If
End Sub

Private Sub lblExit_Click()
Multimedia.mmOpen "expand.wav"
        Multimedia.mmPlay

Call backup
Unload Me
Unload MainForm
End Sub


Private Sub Combo1_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
    If Not IgnorarListaClick Then
            List1.Clear
            Data1.RecordSource = "SELECT * FROM Bank WHERE BankName='" & Combo1.Text & "'"
            Data1.Refresh
            lblBankID.Caption = Data1.Recordset.Fields("BankID")
            Data1.RecordSource = "SELECT * FROM Account WHERE BankID=" & lblBankID.Caption + " and UserID = " & MainForm.lblUserID.Caption
            Data1.Refresh
            While Not Data1.Recordset.EOF
                List1.AddItem Data1.Recordset("AccountNumber")
            Data1.Recordset.MoveNext
            Wend
            Data1.RecordSource = "SELECT * FROM Account WHERE BankID=" & lblBankID.Caption + " and UserID = " & MainForm.lblUserID.Caption
            Data1.Refresh
            'lblBankID.Caption = Data1.Recordset.Fields("BankID")
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim BuscarCadena As String
Dim Retorno As Long
        
    If KeyAscii = 13 Then
        Combo1_Click
        KeyAscii = 0
    Else
        BuscarCadena = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
        'Retorno = SendMessage(Combo1.hWnd, CB_FINDSTRING, -1, ByVal BuscarCadena)
        If Retorno <> CB_ERR Then
            IgnorarListaClick = True
            Combo1.ListIndex = Retorno
            IgnorarListaClick = False
            Combo1.Text = Combo1.List(Retorno)
            Combo1.SelStart = Len(BuscarCadena)
            Combo1.SelLength = Len(Combo1.Text)
            KeyAscii = 0
            Data1.RecordSource = "SELECT * FROM Bank WHERE BankName='" & Combo1.Text & "'"
            Data1.Refresh
            lblBankID.Caption = Data1.Recordset.Fields("BankID")
        End If
    End If
End Sub

Private Sub Form_Load()
Dim LI As ListItem
Dim g, f
    Data1.RecordSource = "SELECT * FROM Bank"
    Data1.Refresh
    '======
    'Fills The Banks ListBox
    While Not Data1.Recordset.EOF
        Combo1.AddItem Data1.Recordset("BankName")
        
        Data1.Recordset.MoveNext
    Wend
    '======
    '======
    'Fills The Company list
    adoCompany.RecordSource = "SELECT * FROM Company where CompanyID <> 5"
    adoCompany.Refresh
    While Not adoCompany.Recordset.EOF
        lstCompany.AddItem adoCompany.Recordset("CompanyName")
        adoCompany.Recordset.MoveNext
    Wend
    TListConn1.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, Transaction.NumberOfStocks, Transaction.PricePerStock, (Debit + Credit) AS Amount, DateOfTransaction FROM [Transaction], Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and TransactionType.BalanceFlag=No and Transaction.UserID = " & MainForm.lblUserID.Caption
    TListConn1.Refresh
    TListConn.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, Transaction.NumberOfStocks, Transaction.PricePerStock, (Debit + Credit) AS Amount, DateOfTransaction FROM [Transaction], Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and TransactionType.BalanceFlag=Yes and Transaction.UserID = " & MainForm.lblUserID.Caption
    TListConn.Refresh
    '======
FormUserWizard.Top = (Screen.Height - FormUserWizard.Height) / 2
FormUserWizard.Left = (Screen.Width - FormUserWizard.Width) / 2
lblTime.Caption = Format(Now, "Long Date")
TSDate.Value = Date
TSDate1.Value = Date
txtEditDate111.Value = Date
txtEditDate1.Value = Date
XPCalendar1.Value = Date
    End Sub

Private Sub List1_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
    'If Not IgnorarListaClick Then
            
            lblAccountID.Caption = List1.Text
            Data1.RecordSource = "select AccountID from Account Where AccountNumber = '" & lblAccountID.Caption & "' and UserID = " & MainForm.lblUserID.Caption
            Data1.Refresh
            lblAccountID.Caption = Data1.Recordset.Fields("AccountID")
            'lblAccountID.Caption = Data1.Recordset.Fields("AccountID")
    'End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Dim BuscarCadena As String
Dim Retorno As Long
'
    If KeyAscii = 13 Then
        List1_Click
        KeyAscii = 0
    Else
'        BuscarCadena = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
'        Retorno = SendMessage(List1.hWnd, CB_FINDSTRING, -1, ByVal BuscarCadena)
'        If Retorno <> CB_ERR Then
'            IgnorarListaClick = True
'            List1.ListIndex = Retorno
'            IgnorarListaClick = False
'            List1.Text = Combo1.List(Retorno)
'            List1.SelStart = Len(BuscarCadena)
'            List1.SelLength = Len(List1.Text)
'            KeyAscii = 0
'            Data1.RecordSource = "SELECT * FROM Bank WHERE BankName='" & List1.Text & "'"
'            Data1.Refresh
'            lblBankID.Caption = Data1.Recordset.Fields("BankID")
'        End If
   End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lstCompany_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
If Not IgnorarListaClickComp Then
            adoCompany.RecordSource = "SELECT * FROM Company WHERE CompanyName='" & lstCompany.Text + "'"
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

Private Sub PBXPButton1_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay

TListConn.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, (Debit + Credit) AS Amount, DateOfTransaction From Transaction, Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and Transactiontype.BalanceFlag=yes and Transaction.TransactionTypeID= " & TTSList.BoundText + " and Transaction.AccountID= " & AccSList.BoundText + " and DateOfTransaction = #" & Format(TSDate.OutputText, "mm/dd/yyyy") & "# and transaction.UserID = " & MainForm.lblUserID.Caption
TListConn.Refresh
counter = 0
While Not TListConn.Recordset.EOF
counter = counter + 1
TListConn.Recordset.MoveNext
Wend
TListConn.Refresh
RC.Caption = counter
RC.Visible = True
RecordCount.Caption = "Record(s) Found."
If TListConn.Recordset.RecordCount = 0 Then
RC = "No "
RecordCount.Caption = "Record(s) Found."

End If

End Sub

Private Sub PBXPButton11_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
TListConn1.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, Transaction.NumberOfStocks, Transaction.PricePerStock, (Debit + Credit) AS Amount, DateOfTransaction From Transaction, Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and Transaction.TransactionTypeID= " & TTSList1.BoundText + " and Transaction.AccountID= " & AccSList1.BoundText + " and DateOfTransaction = #" & Format(TSDate1.OutputText, "mm/dd/yyyy") & "# and Transaction.UserID = " & MainForm.lblUserID.Caption
TListConn1.Refresh
counter1 = 0
While Not TListConn1.Recordset.EOF
counter1 = counter1 + 1
TListConn1.Recordset.MoveNext
Wend
TListConn1.Refresh
RC1.Caption = counter1
RC1.Visible = True
RecordCount1.Caption = "Record(s) Found."
If TListConn1.Recordset.RecordCount = 0 Then
RC1 = "No "
RecordCount1.Caption = "Record(s) Found."

End If
End Sub

Private Sub PBXPButton2_Click()

Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
TListConn.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, Transaction.NumberOfStocks, Transaction.PricePerStock, (Debit + Credit) AS Amount, DateOfTransaction FROM [Transaction], Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and TransactionType.BalanceFlag=Yes and Transaction.UserID = " & MainForm.lblUserID.Caption
counter = 0
TListConn.Refresh
While Not TListConn.Recordset.EOF
counter = counter + 1
TListConn.Recordset.MoveNext
Wend
TListConn.Refresh
RC.Caption = counter
RC.Visible = True
RecordCount.Caption = "Record(s) Found."

End Sub

Private Sub PBXPButton4_Click()

End Sub

Private Sub btnSave_Click()
Dim h
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay

Data4.Recordset.MoveFirst
Data3.Recordset.Edit
While Not Data4.Recordset.EOF
If TTEditList.Text = Data4.Recordset.Fields("TransactionTypeName") Then
    If itsdebit = Data4.Recordset.Fields("DebitFlag") Then
    'Data3.Recordset.Edit
    'MsgBox ("since " & txtEditTT.Text + " and " & TTEditList.Text + " are of the same type, there is no change i'll just paste values and update the transaction type to " & TTEditList.Text)
    cc.Text = Val(c.Text)
    dd.Text = Val(d.Text)
    'MsgBox ("pasted successfully")
    'MsgBox ("now click the update button")
    'acc.Text = Val(AccEditList.BoundText)
    
    Data4.Recordset.MoveLast
    Else
    'MsgBox ("ohhh, " & TTEditList.Text + " is a different transaction type, we will do some changes")
        If itsdebit = True Then
        'MsgBox ("" & TTEditList.Text + " is credit but " & txtEditTT.Text + " is debit, i am gonna switch values")
        cc.Text = Val(d.Text)
        dd.Text = 0
        'MsgBox ("switched successfully")
        Else
        'MsgBox ("" & TTEditList.Text + " is debit but " & txtEditTT.Text + " is credit, i am gonna switch values")
        dd.Text = Val(c.Text)
        cc.Text = 0
        'MsgBox ("switched successfully")
        End If
        'MsgBox ("now click the update button")
        acc.Text = Val(AccEditList.BoundText)
        tt.Text = Val(TTEditList.BoundText)
        Data4.Recordset.MoveLast
    End If
End If
Data4.Recordset.MoveNext
Wend
acc.Text = Val(AccEditList.BoundText)
tt.Text = Val(TTEditList.BoundText)
h = MsgBox("are you sure", vbYesNo)
If h = 6 Then
    
    Data3.Recordset.Update
    TListConn.Refresh
    MsgBox ("Seccess")
    btnCancel_Click
    c.Visible = False
    d.Visible = False
    Me.Refresh
Else
btnCancel_Click
Me.Refresh
End If
End Sub

Private Sub PBXPButton21_Click()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay

TListConn1.RecordSource = "SELECT TransactionsID, AccountNumber, TransactionTypeName, Transaction.NumberOfStocks, Transaction.PricePerStock, (Debit + Credit) AS Amount, DateOfTransaction FROM [Transaction], Account, TransactionType WHERE Transaction.TransactionTypeID=TransactionType.TransactionTypeID and Transaction.AccountID = Account.AccountID and TransactionType.BalanceFlag=no and Transaction.UserID = " & MainForm.lblUserID.Caption
counter1 = 0
TListConn1.Refresh
While Not TListConn1.Recordset.EOF
counter1 = counter1 + 1
TListConn1.Recordset.MoveNext
Wend
TListConn1.Refresh
RC1.Caption = counter1
RC1.Visible = True
RecordCount1.Caption = "Record(s) Found."

End Sub

Private Sub TTBL_Click(Area As Integer)
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay

'While TTBL.BoundText <> ""
Datas.RecordSource = "Select * From TransactionType where TransactionTypeID = " & TTBL.BoundText
Datas.Refresh
lblTTDebitFlag.Caption = Datas.Recordset.Fields("DebitFlag")
lblTTBalanceFlag.Caption = Datas.Recordset.Fields("BalanceFlag")
lblTTypeID.Caption = Datas.Recordset.Fields("TransactionTypeID")
'Wend


End Sub

Private Sub TTSL_Click(Area As Integer)
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay

'While TTSL.BoundText <> ""
Datas.RecordSource = "Select * From TransactionType where TransactionTypeID = " & TTSL.BoundText
Datas.Refresh
lblTTDebitFlag.Caption = Datas.Recordset.Fields("DebitFlag")
lblTTBalanceFlag.Caption = Datas.Recordset.Fields("BalanceFlag")
lblTTypeID.Caption = Datas.Recordset.Fields("TransactionTypeID")
'Wend

End Sub



Private Sub txtAmount_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtComment_Change()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub

Private Sub txtEditAmount_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtPhoneField_KeyPress(KeyAscii As Integer)
 
   'If KeyAscii < 48 Or KeyAscii > 57 Then
    '  KeyAscii = 0  '0 is Ascii Null
   'End If
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub



Private Sub txtStock_KeyPress(Index As Integer, KeyAscii As Integer)
 If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub XPCalendar1_Change()
Multimedia.mmOpen "sound.wav"
Multimedia.mmPlay
lblDate.Caption = XPCalendar1.OutputText
End Sub

Private Sub XPCalendar1_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
End Sub
