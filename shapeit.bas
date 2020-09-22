Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4


Public Function AddOfficeBorder(ByVal hWnd As Long)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function


Public Sub changeForm(ByRef frmChng As Object)
    On Error GoTo EH
    Dim objCtrl As Control
    
    frmChng.Appearance = 0
    AddOfficeBorder (frmChng.hWnd)
    frmChng.BackColor = &H80000016
    For Each objCtrl In frmChng.Controls
    
        If Not TypeOf objCtrl Is Label Then
            objCtrl.Appearance = 0
            
            If TypeOf objCtrl Is TextBox Or TypeOf objCtrl Is CommandButton Or TypeOf objCtrl Is ComboBox Or TypeOf objCtrl Is Frame Then
                
                
                AddOfficeBorder (objCtrl.hWnd)
            End If
            If TypeOf objCtrl Is TextBox Or _
                TypeOf objCtrl Is CommandButton Or _
                TypeOf objCtrl Is ComboBox Or _
                TypeOf objCtrl Is CheckBox Or _
                TypeOf objCtrl Is OptionButton Then
                
                objCtrl.BackColor = &H80000016
                objCtrl.BorderStyle = 0
            End If
            
            If TypeOf objCtrl Is CheckBox Or _
                TypeOf objCtrl Is OptionButton Then
                
                objCtrl.BackColor = &H8000000F
            End If
            If TypeOf objCtrl Is Frame Then
                objCtrl.BackColor = &H8000000F
            End If
        End If
    Next
    
EH:
    If Err.Number = 438 Then
        Resume Next
    End If
End Sub



