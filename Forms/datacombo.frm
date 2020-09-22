VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "BTMS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Bank"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblAccountNumber 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblBankID 
      Caption         =   "Label1"
      DataField       =   "BankID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IgnorarListaClick As Boolean

Private Sub Combo1_Click()
    If Not IgnorarListaClick Then
            List1.Clear
            Data1.RecordSource = "SELECT * FROM Bank WHERE BankName='" & Combo1.Text & "'"
            Data1.Refresh
            lblBankID.Caption = Data1.Recordset.Fields("BankID")
            Data1.RecordSource = "SELECT * FROM Account WHERE BankID=" & lblBankID.Caption
            Data1.Refresh
            List1.AddItem Data1.Recordset("AccountNumber")
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
        Retorno = SendMessage(Combo1.hWnd, CB_FINDSTRING, -1, ByVal BuscarCadena)
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
Data1.RecordSource = "SELECT * FROM Bank"
Data1.Refresh
    While Not Data1.Recordset.EOF
        Combo1.AddItem Data1.Recordset("BankName")
        'List1.AddItem Data1.Recordset("BankName")
        Data1.Recordset.MoveNext
    Wend
End Sub
Private Sub List1_Click()
    If Not IgnorarListaClick Then
            Data1.RecordSource = "SELECT * FROM Account WHERE BankID=" & lblBankID.Caption
            Data1.Refresh
            lblAccountNumber.Caption = Data1.Recordset.Fields("AccountNumber")
    End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Dim BuscarCadena As String
Dim Retorno As Long
        
    If KeyAscii = 13 Then
        List1_Click
        KeyAscii = 0
    Else
        BuscarCadena = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
        Retorno = SendMessage(List1.hWnd, CB_FINDSTRING, -1, ByVal BuscarCadena)
        If Retorno <> CB_ERR Then
            IgnorarListaClick = True
            List1.ListIndex = Retorno
            IgnorarListaClick = False
            List1.Text = Combo1.List(Retorno)
            List1.SelStart = Len(BuscarCadena)
            List1.SelLength = Len(List1.Text)
            KeyAscii = 0
            Data1.RecordSource = "SELECT * FROM Bank WHERE BankName='" & List1.Text & "'"
            Data1.Refresh
            lblBankID.Caption = Data1.Recordset.Fields("BankID")
        End If
    End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub
