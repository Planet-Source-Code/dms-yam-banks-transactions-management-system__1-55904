Attribute VB_Name = "API"
Option Explicit
    Global Const CB_ERR = -1
    Global Const CB_FINDSTRING = &H14C
    Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As Long
