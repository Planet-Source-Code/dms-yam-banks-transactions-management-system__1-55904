VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl DataEditGrid 
   Alignable       =   -1  'True
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ScaleHeight     =   4245
   ScaleWidth      =   5760
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   510
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   1995
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Index           =   0
      Left            =   2850
      TabIndex        =   3
      Top             =   2460
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   57671681
      CurrentDate     =   36877
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3465
      Left            =   300
      TabIndex        =   5
      Top             =   420
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   6112
      _Version        =   393216
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgHolders 
      Height          =   165
      Index           =   2
      Left            =   1380
      Picture         =   "DataEditrGrid.ctx":0000
      Top             =   3810
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHolders 
      Height          =   165
      Index           =   1
      Left            =   690
      Picture         =   "DataEditrGrid.ctx":01FA
      Top             =   3720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHolders 
      Height          =   105
      Index           =   0
      Left            =   360
      Picture         =   "DataEditrGrid.ctx":03F4
      Top             =   3870
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "DataEditGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const mcstrMod$ = "DataEditGrid"
Private m_iSortCol As Integer
Private m_iSortType As Integer
Private m_iSortCustomAscending As Boolean

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer
Private ctl As Control
Private m_Datasource As COMEXDataSource
Public Enum FieldControlType
    fcEditBx
    fcComboBx
    fcMaskBx
    fcCheckBx
    fcDateTimePick
End Enum

Public Enum FieldControlAlign
    fcLeft
    fcCenter
    fcRight
End Enum

Private Enum PictureType
    ptArrow = 0
    ptPen
    ptStar
    ptNone
End Enum

Private Type ColDataType
    cdObjectIndex As Long
    cdControlType As FieldControlType
    cdComboMaskList As String
    cdComboMaskIndex As Integer
    cdHidden As Boolean
    cdAutoNumber As Boolean
End Type

Private m_ComboCount As Integer
Private m_MaskCount As Integer
Private m_flgLoading As Boolean
Private m_ColDataType() As ColDataType
Private m_LastRow As Long
Event FetchColumnSetup(ByRef ColName As String, ByRef ControlType As FieldControlType, ByRef ComboMaskList As String, ByRef Alignment As FieldControlAlign, ByRef Hidden As Boolean, ByRef AutoNumber As Boolean)
Event Dirty()
'Default Property Values:
Const m_def_AllowAddNew = 0
'Property Variables:
Dim m_AllowAddNew As Boolean

Public Property Get CurrentRecord() As Long
    With MSHFlexGrid1
        CurrentRecord = .RowData(.Row)
    End With
End Property
    
Public Property Get hwnd()
    hwnd = UserControl.hwnd
End Property

Public Property Set DataSource(vDatasource As COMEXDataSource)
    Dim i As Long, j As Long, m_ControlType As FieldControlType, m_ComboList As String, _
        m_Align As FieldControlAlign, m_Hidden As Boolean, m_Caption As String, _
        m_AutoNumber As Boolean
    Dim strName As String, iCount As Long, iRowCount As Long
    Dim cProg As CProgBar32
    
    On Error GoTo Err_DataSource
    
    Screen.MousePointer = vbHourglass
    Set m_Datasource = vDatasource
    Reset
    With MSHFlexGrid1
        .Redraw = False
        .Clear
        'If vDatasource.GetRecordCount <= 0 Then GoTo Done_DataSource
        .Cols = vDatasource.GetFieldCount + 1
        If m_AllowAddNew Then
            .Rows = vDatasource.GetRecordCount + 2
        Else
            .Rows = vDatasource.GetRecordCount + 1
        End If
        iRowCount = .Rows
        iCount = .Cols * iRowCount
        .ColWidth(0) = 200
        ReDim m_ColDataType(vDatasource.GetFieldCount)
        Set cProg = New CProgBar32
        CreateProgessBar cProg
        fMainForm.SetStatus "Loading Data..."
        For i = 1 To .Cols - 1
            m_Caption = m_Datasource.GetFieldName(i)
            m_ControlType = fcEditBx
            m_Hidden = False
            m_AutoNumber = False
            RaiseEvent FetchColumnSetup(m_Caption, m_ControlType, m_ComboList, m_Align, m_Hidden, m_AutoNumber)
            .TextMatrix(0, i) = m_Caption
            .ColWidth(i) = UserControl.TextWidth(m_Caption) + 400
            If m_ControlType = fcComboBx Then
                LoadComboState i, m_ComboList
            ElseIf m_ControlType = fcMaskBx Then
                LoadMaskState i, m_ComboList
            End If
            With m_ColDataType(i - 1)
                .cdObjectIndex = i
                .cdComboMaskList = m_ComboList
                .cdControlType = m_ControlType
                .cdAutoNumber = m_AutoNumber
                .cdHidden = m_Hidden
            End With
            Select Case m_Align
                Case fcLeft
                    .ColAlignment(i) = flexAlignLeftCenter
                Case fcRight
                    .ColAlignment(i) = flexAlignRightCenter
                Case fcCenter
                    .ColAlignment(i) = flexAlignCenterCenter
            End Select
            For j = 1 To .Rows - 1
                .TextMatrix(j, i) = vDatasource.GetData(i, j) & ""
                If ((i - 1) * iRowCount + j) Mod 320 = 1 Then
                    cProg.SetProgBarPos (i * iRowCount + j) / (iCount) * 100
                    cProg.DelayProgBar 50
                End If
            Next
        Next
        If .Rows > 1 Then
            .FillStyle = flexFillRepeat
            If m_AllowAddNew Then
                    For i = .FixedRows To .Rows - 2
                        .Row = i
                        .Col = 0
                        Set .CellPicture = Nothing
                        'don't set the last line index
                        .RowData(i) = i
                    Next
                    DrawIcons ptStar, .Rows - 1, False
                Else
                        For i = .FixedRows To .Rows - 1
                        .Row = i
                        .Col = 0
                        Set .CellPicture = Nothing
                        'don't set the last line index
                        .RowData(i) = i
                    Next
                End If
            For i = .FixedRows + 1 To .Rows - 1 Step 2
                .Row = i
                .Col = .FixedCols
                .ColSel = .Cols() - .FixedCols
                .CellBackColor = vbInfoBackground  ' light grey
            Next i
            .FillStyle = flexFillSingle
        End If
        .RowHeight(-1) = 315
        If .Rows > 1 Then .Row = 1
        If .Cols > 1 Then .Col = 1
        .Redraw = True
    End With
    
Done_DataSource:
    Screen.MousePointer = vbDefault
    cProg.SetProgBarPos 0
    cProg.DestroyProgBar
    fMainForm.SetStatus
    Exit Property
Err_DataSource:
    ErrorMsg Err.Number, Err.Description, "DataSource", mcstrMod
    Resume Done_DataSource
End Property

Public Sub Save()
    m_Datasource.Save
End Sub

Public Sub delete()
    On Error GoTo Err_Delete
    
    Dim iColOld As Long, iColOldSel As Long, i As Long
    Dim iRowOld As Long, iRowOldSel As Long
    With MSHFlexGrid1
        .Redraw = False
        .FillStyle = flexFillRepeat
        iColOldSel = .ColSel
        iColOld = .Col
        iRowOld = .Row
        iRowOldSel = .RowSel
        .Row = iRowOld
        .Col = .FixedCols
        .RowSel = iRowOldSel
        .ColSel = .Cols - .FixedCols - 1
        .CellFontStrikeThrough = True
        .CellForeColor = vbRed
        .FillStyle = flexFillSingle
        For i = .Row To .RowSel
            m_Datasource.delete .RowData(i)
        Next
        .Row = iRowOld
        .Col = iColOld
        .RowSel = iRowOldSel
        .ColSel = iColOldSel
        .Redraw = True
    End With
    RaiseEvent Dirty
Done_Delete:
    Exit Sub
Err_Delete:
    ErrorMsg Err.Number, Err.Description, "Delete", mcstrMod
    Resume Done_Delete
End Sub

Private Sub Check1_Click()
    On Error Resume Next
    With MSHFlexGrid1
        If .RowData(.Row) = 0 Then
                .RowData(.Row) = m_Datasource.GetRecordCount + 1
                .Rows = .Rows + 1
        End If
        m_Datasource.SetData .Col, .RowData(.Row), Check1
        .TextMatrix(.Row, .Col) = CBool(Check1)
    End With
    RaiseEvent Dirty
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            With MSHFlexGrid1
                .TextMatrix(.Row, .Col) = Check1.Tag
            End With
            Check1.Visible = False
        Case vbKeyReturn
            Check1.Visible = False
    End Select
End Sub

Private Sub Check1_LostFocus()
    On Error Resume Next
    Check1.Visible = False
End Sub

Private Sub Combo1_Change(Index As Integer)
    On Error Resume Next
    If m_flgLoading Then Exit Sub
    Combo1_Click Index
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error Resume Next
    Dim strAll() As String, strRow() As String
    With MSHFlexGrid1
        strAll = Split(m_ColDataType(.Col - 1).cdComboMaskList, "|")
        strRow = Split(strAll(Combo1(Index).ListIndex), vbTab)
        If .RowData(.Row) = 0 Then
                .RowData(.Row) = m_Datasource.GetRecordCount + 1
                .Rows = .Rows + 1
        End If
        m_Datasource.SetData .Col, .RowData(.Row), strRow(0)
        .TextMatrix(.Row, .Col) = Combo1(Index)
    End With
    RaiseEvent Dirty
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            With MSHFlexGrid1
                .TextMatrix(.Row, .Col) = Combo1(Index).Tag
            End With
             Combo1(Index).Visible = False
        Case vbKeyReturn
             Combo1(Index).Visible = False
    End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    On Error Resume Next
    Combo1(Index).Visible = False
End Sub

Private Sub DTPicker1_Change()
    On Error Resume Next
    With MSHFlexGrid1
        If .RowData(.Row) = 0 Then
                .RowData(.Row) = m_Datasource.GetRecordCount + 1
                .Rows = .Rows + 1
        End If
        m_Datasource.SetData .Col, .RowData(.Row), DTPicker1.Value
        .TextMatrix(.Row, .Col) = DTPicker1.Value
    End With
    RaiseEvent Dirty
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            With MSHFlexGrid1
                .TextMatrix(.Row, .Col) = DTPicker1.Tag
            End With
            DTPicker1.Visible = False
        Case vbKeyReturn
            DTPicker1.Visible = False
    End Select
End Sub

Private Sub DTPicker1_LostFocus()
    On Error Resume Next
    DTPicker1.Visible = False
End Sub

Private Sub MaskEdBox1_Change(Index As Integer)
    On Error Resume Next
    If m_flgLoading Then Exit Sub
    With MSHFlexGrid1
        If .RowData(.Row) = 0 Then
                .RowData(.Row) = m_Datasource.GetRecordCount + 1
                .Rows = .Rows + 1
        End If
        m_Datasource.SetData .Col, .RowData(.Row), MaskEdBox1(Index)
        .TextMatrix(.Row, .Col) = MaskEdBox1(Index)
    End With
    RaiseEvent Dirty
End Sub

Private Sub MaskEdBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            With MSHFlexGrid1
                .TextMatrix(.Row, .Col) = MaskEdBox1(Index).Tag
            End With
            MaskEdBox1(Index).Visible = False
        Case vbKeyReturn
            MaskEdBox1(Index).Visible = False
    End Select
End Sub

Private Sub MaskEdBox1_LostFocus(Index As Integer)
    On Error Resume Next
    MaskEdBox1(Index).Visible = False
End Sub

Private Sub MSHFlexGrid1_Click()
    On Error Resume Next
    MSHFlexGrid1_EnterCell
End Sub

Private Sub MSHFlexGrid1_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    On Error Resume Next
    Dim dtmRow1 As Date, dtmRow2 As Date
    With MSHFlexGrid1
        dtmRow1 = CDate(.TextMatrix(Row1, m_iSortCol))
        dtmRow2 = CDate(.TextMatrix(Row2, m_iSortCol))
        If dtmRow1 > dtmRow2 Then
            Cmp = IIf(m_iSortCustomAscending, 1, -1)
        ElseIf dtmRow1 = dtmRow2 Then
            Cmp = 0
        Else
            Cmp = IIf(m_iSortCustomAscending, -1, 1)
        End If
    End With
End Sub

Private Sub MSHFlexGrid1_EnterCell()
    On Error GoTo Err_MSHFlexGrid1_EnterCell
    Dim strType As String
    
    m_flgLoading = True
    With MSHFlexGrid1
        If .Col < 1 Or .Row < 1 Then Exit Sub
        ' uneditable if it's a foreign key field or an autonumber field
        With m_ColDataType(.Col - 1)
            If .cdHidden Or .cdAutoNumber Then Exit Sub
        End With
        Set ctl = GetColControl(.Col)
        strType = TypeName(ctl)
        'combox cannot set height , so ..
        If strType = "ComboBox" Then
            Dim ctlCombo As ComboBox
            Set ctlCombo = ctl
            ctlCombo.Top = .CellTop
            ctlCombo.Left = .CellLeft
            ctlCombo.Width = .CellWidth
        Else
            ctl.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        End If
        Select Case strType
            Case "CheckBox"
                ctl = Abs(Val(.TextMatrix(.Row, .Col)))
            Case "MaskEdBox"
                Dim s As String
                s = ctl.Mask
                ctl.Mask = vbNullString
                ctl = .TextMatrix(.Row, .Col)
                ctl.Mask = s
            Case Else
                ctl = MSHFlexGrid1.TextMatrix(.Row, .Col)
        End Select
        ctl.Tag = MSHFlexGrid1.TextMatrix(.Row, .Col)
        ctl.ZOrder 0
        ctl.Visible = True
        ctl.SetFocus
    End With
Done_MSHFlexGrid1_EnterCell:
    m_flgLoading = False
    Exit Sub
Err_MSHFlexGrid1_EnterCell:
        ErrorMsg Err.Number, Err.Description, "MSHFlexGrid1_EnterCell", mcstrMod
        Resume Done_MSHFlexGrid1_EnterCell
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        MsgBox "hi"
    End If
End Sub

Private Sub MSHFlexGrid1_RowColChange()
    On Error Resume Next
    With MSHFlexGrid1
        If m_LastRow <> -1 Then
            If m_LastRow = .Rows - 1 Then
                If AllowAddNew Then
                    DrawIcons ptStar, m_LastRow
                Else
                    DrawIcons ptNone, m_LastRow
                End If
            Else
                DrawIcons ptNone, m_LastRow
            End If
        End If
        m_LastRow = .Row
        DrawIcons ptArrow, m_LastRow
    End With
End Sub

Private Sub DrawIcons(ByVal PicType As PictureType, ByVal vRow As Long, Optional UseRedraw As Boolean = True)
    On Error GoTo Err_DrawIcons
    Dim iRow&, iCol&, iRowsel&, iColSel&
    With MSHFlexGrid1
        If UseRedraw Then .Redraw = False
        iRow = .Row
        iCol = .Col
        iRowsel = .RowSel
        iColSel = .ColSel
        .Row = vRow
        .Col = 0
        .CellPictureAlignment = flexAlignCenterCenter
        If PicType = ptNone Then
            Set .CellPicture = Nothing
        Else
            Set .CellPicture = imgHolders(PicType).Picture
        End If
        .Row = iRow
        .Col = iCol
        .RowSel = iRowsel
        .ColSel = iColSel
        If UseRedraw Then .Redraw = True
    End With
    Exit Sub
Err_DrawIcons:
    ErrorMsg Err.Number, Err.Description, "DrawIcons", mcstrMod
End Sub
Private Sub MSHFlexGrid1_Scroll()
    On Error Resume Next
    'MSHFlexGrid1_EnterCell
    ctl.Visible = False
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    If m_flgLoading Then Exit Sub
    With MSHFlexGrid1
        If .RowData(.Row) = 0 Then
                .RowData(.Row) = m_Datasource.GetRecordCount + 1
                .Rows = .Rows + 1
        End If
        m_Datasource.SetData .Col, .RowData(.Row), IIf(Text1 = ".", "0.", Text1)
        .TextMatrix(.Row, .Col) = Text1
    End With
    RaiseEvent Dirty
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            With MSHFlexGrid1
                .TextMatrix(.Row, .Col) = Text1.Tag
            End With
            Text1.Visible = False
        Case vbKeyReturn
            Text1.Visible = False
    End Select
End Sub

Private Sub Text1_LostFocus()
    On Error Resume Next
    Text1.Visible = False
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    m_iSortCustomAscending = False
    Set ctl = Nothing
    m_LastRow = -1
    With MSHFlexGrid1
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    MSHFlexGrid1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Function GetColControl(ByVal ColIndex As Long) As Control
    On Error Resume Next
    Select Case m_ColDataType(ColIndex - 1).cdControlType
        Case fcEditBx
            Set GetColControl = Text1
        Case fcComboBx
            Set GetColControl = Combo1(m_ColDataType(ColIndex - 1).cdComboMaskIndex)
        Case fcMaskBx
            Set GetColControl = MaskEdBox1(m_ColDataType(ColIndex - 1).cdComboMaskIndex)
        Case fcCheckBx
            Set GetColControl = Check1
        Case fcDateTimePick
            Set GetColControl = DTPicker1
    End Select
End Function

Private Sub ResetComboState()
    On Error GoTo Err_ResetComboState
    Dim i As Long
    m_ComboCount = 1
    Do While Combo1.Count > 1
        Unload Combo1(Combo1.Count - 1)
    Loop
    Combo1(0).Clear
    Exit Sub
Err_ResetComboState:
    ErrorMsg Err.Number, Err.Description, "ResetComboState", mcstrMod
End Sub

Private Sub LoadComboState(ByVal Col As Long, ByVal ComboList As String)
    On Error GoTo Err_LoadComboState
    Dim strAll() As String, i&
    If m_ComboCount > 1 Then Load Combo1(m_ComboCount - 1)
    m_ColDataType(Col - 1).cdComboMaskIndex = m_ComboCount - 1
    strAll = Split(ComboList, "|")
    For i = 0 To UBound(strAll)
        Combo1(m_ComboCount - 1).AddItem strAll(i)
    Next
    m_ComboCount = m_ComboCount + 1
    Exit Sub
Err_LoadComboState:
    ErrorMsg Err.Number, Err.Description, "LoadComboState", mcstrMod
End Sub

Private Sub ResetMaskState()
    On Error GoTo Err_ResetMaskState
    Dim i As Long
    m_MaskCount = 1
    Do While MaskEdBox1.Count > 1
        Unload MaskEdBox1(MaskEdBox1.Count - 1)
    Loop
    MaskEdBox1(0).Mask = vbNullString
    Exit Sub
Err_ResetMaskState:
    ErrorMsg Err.Number, Err.Description, "ResetMaskState", mcstrMod
End Sub

Private Sub LoadMaskState(ByVal Col As Long, ByVal Mask As String)
    On Error GoTo Err_LoadMaskState
    m_flgLoading = True
    If m_MaskCount > 1 Then Load MaskEdBox1(m_MaskCount - 1)
    m_ColDataType(Col - 1).cdComboMaskIndex = m_MaskCount - 1
    MaskEdBox1(m_MaskCount - 1).Mask = Mask
    m_MaskCount = m_MaskCount + 1
    m_flgLoading = False
    Exit Sub
Err_LoadMaskState:
    ErrorMsg Err.Number, Err.Description, "LoadMaskState", mcstrMod
End Sub

Private Sub Reset()
    Erase m_ColDataType
    ResetMaskState
    ResetComboState
End Sub

Private Sub MSHFlexGrid1_DblClick()
'-------------------------------------------------------------------------------------------
' code in grid's DblClick event enables column sorting
'-------------------------------------------------------------------------------------------
        On Error Resume Next
    Dim i As Integer

    ' sort only when a fixed row is clicked
    If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub

    If Not (ctl Is Nothing) Then ctl.Visible = False
    i = m_iSortCol                  ' save old column
    m_iSortCol = MSHFlexGrid1.Col   ' set new column

    ' increment sort type
    If i <> m_iSortCol Then
        ' if clicking on a new column, start with ascending sort
        If m_ColDataType(m_iSortCol - 1).cdControlType = fcDateTimePick Then
            m_iSortCustomAscending = Not m_iSortCustomAscending
            m_iSortType = 9
        Else
            m_iSortType = 1
            m_iSortCustomAscending = False
        End If
    Else
        ' if clicking on the same column, toggle between ascending and descending sort
        
        Select Case m_iSortType
            Case 3
                m_iSortType = 1
            Case 9
                m_iSortCustomAscending = Not m_iSortCustomAscending
            Case Else
                m_iSortType = m_iSortType + 1
        End Select
    End If

    DoColumnSort

End Sub

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------
        On Error Resume Next
    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType
        .FillStyle = flexFillRepeat
        .Col = .FixedCols
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &HFFFFFF
        ' grey every other row
        Dim iLoop As Integer
        For iLoop = .FixedRows + 1 To .Rows - 1 Step 2
            .Row = iLoop
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
            .CellBackColor = vbInfoBackground  ' light grey
        Next iLoop
        .FillStyle = flexFillSingle
        .Redraw = True
    End With

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AllowAddNew() As Boolean
    AllowAddNew = m_AllowAddNew
End Property

Public Property Let AllowAddNew(ByVal New_AllowAddNew As Boolean)
    m_AllowAddNew = New_AllowAddNew
    PropertyChanged "AllowAddNew"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowAddNew = m_def_AllowAddNew
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AllowAddNew = PropBag.ReadProperty("AllowAddNew", m_def_AllowAddNew)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AllowAddNew", m_AllowAddNew, m_def_AllowAddNew)
End Sub

Public Sub Update()

End Sub
