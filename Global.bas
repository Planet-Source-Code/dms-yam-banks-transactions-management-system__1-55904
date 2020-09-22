Attribute VB_Name = "Global"
Option Explicit

'*************************************************************************
'API declaration

Public Type GUID
    PartOne As Long
    PartTwo As Integer
    PartThree As Integer
    PartFour(7) As Byte
End Type
Public Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type


Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CoCreateGuid Lib "OLE32.DLL" (ptrGuid As GUID) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE = 2
    Public Const TRANSPARENT = 1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Public Const DT_CENTER = &H1&
    Public Const DT_TOP = &H0&
    Public Const DT_LEFT = &H0&
    Public Const DT_BOTTOM = &H8&
    Public Const DT_SINGLELINE = &H20&
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    Public Const DI_NORMAL = &H3
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

'end API declaration
'*************************************************************************

Public Const CON_ROW_DELIM$ = "|"
Public Const CON_COL_DELIM$ = vbTab

Private Const mcstrMod$ = "Global"

Public fMainForm As frmMain
Public m_Connect As Connect
Public m_Toolbar As cToolbar

Public Function ThinBorder(ByVal lHwnd As Long, ByVal bState As Integer)
On Error Resume Next
Dim lS As Long

    lS = GetWindowLong(lHwnd, GWL_EXSTYLE)
    Select Case bState
        Case 1
            lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
        Case 0
            lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
        Case -1
            lS = lS And Not WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    End Select
    SetWindowLong lHwnd, GWL_EXSTYLE, lS
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function

Public Sub Main()
    On Error GoTo Err_Main
    frmSplash.Show vbModal
    'frmSplash.Refresh
    Set m_Connect = New Connect
    m_Connect.Connect
    
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash

    fMainForm.Show
    
    Set m_Toolbar = New cToolbar
    m_Toolbar.Attach fMainForm, fMainForm.Name
Done_Main:
    Exit Sub
Err_Main:
    ErrorMsg Err.Number, Err.Description, "Main", mcstrMod
    Resume Done_Main
End Sub

Public Sub FillCombo(ctl As ComboBox, ByVal ComboList As String)
    On Error Resume Next
    Dim strAll() As String, i&
    strAll = Split(ComboList, "|")
    For i = 0 To UBound(strAll)
        ctl.AddItem strAll(i)
    Next
End Sub

Public Function GetComboValue(ctl As ComboBox, ByVal ComboList As String)
    On Error Resume Next
    Dim strAll() As String, strRow() As String
    strAll = Split(ComboList, "|")
    strRow = Split(strAll(ctl.ListIndex), vbTab)
    GetComboValue = strRow(0)
End Function

Function ErrorMsg(ErrNum As Long, ErrDesc As String, _
    strFunction As String, strModule As String)
    On Error Resume Next
    Dim anErrorMessage As String
    anErrorMessage = "Error Number: " & ErrNum & "." & vbCrLf & _
        "Error Description: " & ErrDesc & vbCrLf & _
        "Module Name: " & strModule & vbCrLf & _
        "Sub/Function: " & strFunction & vbCrLf
    MsgBox anErrorMessage, vbCritical
End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    On Error Resume Next
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function


'Why:This function is to help us to generate for a unique number for forms
'for example, frmOrders might have instances running so we need to identifier to keep
'the toolbar status
'RETURNS:  GUID if successful; blank string otherwise.
'Unlike the GUIDS in the registry, this function returns GUID
'without "-" characters.  See comments for how to modify if you
'want the dash.

Public Function GUID() As String
    Dim lRetVal As Long
    Dim udtGuid As GUID
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
   
    On Error GoTo errorhandler
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
    
       'First 8 chars
        sPartOne = Hex$(udtGuid.PartOne)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.PartOne)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartOne)
        
        'Next 4 Chars
        sPartTwo = Hex$(udtGuid.PartTwo)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.PartTwo)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartTwo)
           
        'Next 4 Chars
        sPartThree = Hex$(udtGuid.PartThree)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.PartThree)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartThree)   'Next 2 bytes (4 hex digits)
           
        'Final 16 chars
        For iCtr = 0 To 7
            sPartFour = sPartFour & _
            Format$(Hex$(udtGuid.PartFour(iCtr)), "00")
        Next
 
     'To create GUID with "-", change line below to:
     'sAns = sPartOne & "-" & sPartTwo & "-" & sPartThree _
     '& "-" & sPartFour
       
       sAns = sPartOne & sPartTwo & sPartThree & sPartFour
            
    End If
        
    GUID = sAns
Exit Function


errorhandler:
'return a blank string if there's an error
Exit Function
End Function

Function GetFindString() As String
    On Error Resume Next
    Dim frmX As New frmFind
    With frmX
        .Show vbModal
        If .m_OK Then GetFindString = .Text1
    End With
    Unload frmX
    Set frmX = Nothing
End Function

Function IsLoaded(ByVal Frm As Form) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = Frm.Name Then
            IsLoaded = True
            Exit Function
        End If
    Next
    IsLoaded = False
End Function

Sub Draw3dRect(hDC As Long, rc As RECT, clrTopLeft As OLE_COLOR, _
    clrBottomRight As OLE_COLOR)
    Dim x As Long, Y As Long, cx As Long, cy As Long
    x = rc.Left
    Y = rc.Top
    cx = rc.right - rc.Left
    cy = rc.bottom - rc.Top
    
    FillSolidRect hDC, x, Y, cx - 1, 1, clrTopLeft
    FillSolidRect hDC, x, Y, 1, cy - 1, clrTopLeft
    FillSolidRect hDC, x + cx, Y, -1, cy, clrBottomRight
    FillSolidRect hDC, x, Y + cy, cx, -1, clrBottomRight

End Sub

Sub FillSolidRect(hDC As Long, x As Long, Y As Long, cx As Long, _
    cy As Long, clr As OLE_COLOR)
    Dim hBr As Long, rc As RECT
    rc.Left = x
    rc.Top = Y
    rc.right = x + cx
    rc.bottom = Y + cy
    hBr = CreateSolidBrush(TranslateColor(clr))
    FillRect hDC, rc, hBr
    DeleteObject hBr
End Sub

Public Sub CreateProgessBar(cProg As CProgBar32)
    On Error GoTo Err_CreateProgessBar
    Dim o As StatusBar
    With cProg
        Set o = fMainForm.sbStatusBar
        Set .Parent = o
        .Create 100, 2, o.Panels(1).Width / Screen.TwipsPerPixelX - _
            100, o.Height / Screen.TwipsPerPixelY - 3
    End With
done_CreateProgessBar:
    Set o = Nothing
    Exit Sub
Err_CreateProgessBar:
    Resume done_CreateProgessBar
End Sub
