VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Reports 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   6
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   2040
         Top             =   1560
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "TransactionCreditQuery"
         Caption         =   "Adodc5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid5 
         Bindings        =   "Reports.frx":0000
         Height          =   5895
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image Image5 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":0015
         Top             =   240
         Width           =   11385
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   735
         Left            =   1680
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "CompanyTansactionsQuery"
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid2 
         Bindings        =   "Reports.frx":4FA8
         Height          =   5895
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Company Name"
            Caption         =   "Company Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Number ofTransactions"
            Caption         =   "Number ofTransactions"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3165.166
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":4FBD
         Top             =   120
         Width           =   11385
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   2040
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1296
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "BalanceTransactionQuery"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Bindings        =   "Reports.frx":9F50
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Account Number"
            Caption         =   "Account Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Transaction Type"
            Caption         =   "Transaction Type"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Comments"
            Caption         =   "Comments"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Total Amount"
            Caption         =   "Total Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3465.071
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":9F65
         Top             =   120
         Width           =   11385
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc0 
         Height          =   495
         Left            =   840
         Top             =   1200
         Visible         =   0   'False
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "AccountBalanceQuery"
         Caption         =   "Adodc0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid0 
         Bindings        =   "Reports.frx":EEF8
         Height          =   5895
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         RowDividerStyle =   5
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "AccountNumber"
            Caption         =   "AccountNumber"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "BankName"
            Caption         =   "BankName"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Balance"
            Caption         =   "Balance"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Count Of Transactions"
            Caption         =   "Count Of Transactions"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3165.166
            EndProperty
         EndProperty
      End
      Begin VB.Image Imgfrmr0 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":EF0D
         Top             =   120
         Width           =   11385
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   5
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   1920
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "TransactionDebitQuery"
         Caption         =   "Adodc4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid4 
         Bindings        =   "Reports.frx":13EA0
         Height          =   5895
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image Image4 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":13EB5
         Top             =   120
         Width           =   11385
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   11295
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   1800
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=btms.dll;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "StockTransactionQuery"
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Grid3 
         Bindings        =   "Reports.frx":18E48
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Account Number"
            Caption         =   "Account Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Company Name"
            Caption         =   "Company Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Total Amount"
            Caption         =   "Total Amount"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Comments"
            Caption         =   "Comments"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Number of Stocks"
            Caption         =   "Number of Stocks"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Price Per Stock"
            Caption         =   "Price Per Stock"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "TransactionTypeName"
            Caption         =   "TransactionTypeName"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2264.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3465.071
            EndProperty
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   6555
         Left            =   120
         Picture         =   "Reports.frx":18E5D
         Top             =   120
         Width           =   11385
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6495
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11456
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Account Balances"
            Key             =   "AB"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Balance Transactions"
            Key             =   "BT"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Company Transactions"
            Key             =   "CT"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock Transactions"
            Key             =   "ST"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sold Stocks"
            Key             =   "SS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bought Stocks"
            Key             =   "BS"
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Reports.frx":1DDF0
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   11640
      MouseIcon       =   "Reports.frx":1DF52
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ImgFormBack 
      Height          =   9390
      Left            =   0
      Picture         =   "Reports.frx":1E0A4
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Numtabs = 6 'Set the number of tabs
Dim X As Integer  'For/Next loop variable
Dim Multimedia As New Mmedia

Private Sub Form_Load()
'Generic error trapping
  On Error Resume Next
  '------------------------------------------------
  'This sets up your containers - the pictureboxes
  'so they are the same size as the tabstrip
  '------------------------------------------------
  For X = 1 To Numtabs 'Loop through the tabs
    
    With frm(X)
      .BorderStyle = 0
      .Left = TabStrip1.ClientLeft
      .Top = TabStrip1.ClientTop
      .Width = TabStrip1.ClientWidth
      .Height = TabStrip1.ClientHeight
      .Visible = False
    End With
    
  Next X
  
  '-----------------------------------
  'Form loads with first tab/picturebox selected
  '-----------------------------------
  frm(1).Visible = True
  
Adodc0.RecordSource = "SELECT DISTINCTROW Account.AccountNumber, Bank.BankName, Sum(Transaction.Debit)- Sum(Transaction.Credit) AS Balance, Count(*) AS [Count Of Transactions] FROM Bank INNER JOIN (Account INNER JOIN [Transaction] ON Account.AccountID = Transaction.AccountID) ON Bank.BankID = Account.BankID where Account.userID = " & MainForm.lblUserID.Caption & " GROUP BY Account.AccountNumber, Bank.BankName "
Adodc0.Refresh
Adodc1.RecordSource = "SELECT DISTINCTROW Account.AccountNumber AS [Account Number], TransactionType.TransactionTypeName AS [Transaction Type], Transaction.DateOfTransaction AS [Date], Transaction.Comments AS Comments, Sum(Transaction.Debit)-Sum(Transaction.Credit) AS [Total Amount] FROM TransactionType INNER JOIN (Account INNER JOIN [Transaction] ON Account.AccountID = Transaction.AccountID) ON TransactionType.TransactionTypeID = Transaction.TransactionTypeID Where (((TransactionType.BalanceFlag) = Yes)) and transaction.userID = " & MainForm.lblUserID.Caption + " GROUP BY Account.AccountNumber, TransactionType.TransactionTypeName, Transaction.DateOfTransaction, Transaction.Comments"
Adodc1.Refresh
Adodc2.RecordSource = "SELECT DISTINCTROW Company.CompanyName AS [Company Name], count ([Transaction].[CompanyID]) AS [Number ofTransactions] FROM Company INNER JOIN [Transaction] ON Company.CompanyID = Transaction.CompanyID where transaction.userID = " & MainForm.lblUserID.Caption + "  GROUP BY Company.CompanyName"
Adodc2.Refresh
Adodc3.RecordSource = "SELECT DISTINCTROW Account.AccountNumber AS [Account Number], Company.CompanyName AS [Company Name], Transaction.DateOfTransaction AS [Date], (Transaction.NumberOfStocks*Transaction.PricePerStock) AS [Total Amount], First(Transaction.Comments) AS Comments, [Transaction.NumberOfStocks] AS [Number of Stocks], Transaction.PricePerStock AS [Price Per Stock], TransactionType.TransactionTypeName FROM TransactionType INNER JOIN (Company INNER JOIN (Account INNER JOIN [Transaction] ON Account.AccountID = Transaction.AccountID) ON Company.CompanyID = Transaction.CompanyID) ON TransactionType.TransactionTypeID = Transaction.TransactionTypeID Where (((TransactionType.BalanceFlag) = No)) and transaction.userID = " & MainForm.lblUserID.Caption + "  GROUP BY Account.AccountNumber, Company.CompanyName, Transaction.DateOfTransaction, Transaction.PricePerStock, Transaction.NumberOfStocks, TransactionType.TransactionTypeName"
Adodc3.Refresh
Adodc4.RecordSource = "SELECT DISTINCTROW Sum([Transaction].[NumberOfStocks]) AS [Number Of Sold Stocks], Sum(Transaction.PricePerStock*Transaction.NumberOfStocks) AS [Total Debit], Company.CompanyName AS [Company Name] FROM Company INNER JOIN (TransactionType INNER JOIN [Transaction] ON TransactionType.TransactionTypeID = Transaction.TransactionTypeID) ON Company.CompanyID = Transaction.CompanyID Where (((TransactionType.BalanceFlag) = 0) And ((TransactionType.DebitFlag) = Yes)) and transaction.userID = " & MainForm.lblUserID.Caption + " GROUP BY Company.CompanyName"
Adodc4.Refresh
Adodc5.RecordSource = "SELECT DISTINCTROW Sum(Transaction.NumberOfStocks) AS [Number Of Bought Stocks], Sum(Transaction.PricePerStock*Transaction.NumberOfStocks) AS [Total Credit], Company.CompanyName AS [Company Name] FROM TransactionType INNER JOIN (Company INNER JOIN [Transaction] ON Company.CompanyID = Transaction.CompanyID) ON TransactionType.TransactionTypeID = Transaction.TransactionTypeID Where (((TransactionType.BalanceFlag) = 0) And ((TransactionType.DebitFlag) = No)) and transaction.userID = " & MainForm.lblUserID.Caption + "  GROUP BY Company.CompanyName"
Adodc5.Refresh
Reports.Top = (Screen.Height - Reports.Height) / 2
Reports.Left = (Screen.Width - Reports.Width) / 2
End Sub

Private Sub lblExit_Click()
Multimedia.mmOpen "expand.wav"
        Multimedia.mmPlay
Unload Me
Load FormUserWizard
FormUserWizard.Show
End Sub

Private Sub TabStrip1_Click()
Multimedia.mmOpen "sound.wav"
        Multimedia.mmPlay
Static PrevTab As Integer
  PrevTab = Switch(PrevTab = 0, 1, PrevTab >= 1 And PrevTab <= Numtabs, PrevTab)
  frm(PrevTab).Visible = False
  frm(TabStrip1.SelectedItem.Index).Visible = True
  frm(TabStrip1.SelectedItem.Index).Refresh
  PrevTab = TabStrip1.SelectedItem.Index
End Sub
