Attribute VB_Name = "Module1"
Option Explicit

'general declarations
Private Declare Function SendMessageBynum _
    Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Const CB_SETDROPPEDWIDTH = &H160

'this function sets the width of the list portion of the combo box
Public Sub SetComboBoxDropDownWidth(hWnd As Long, WidthPx As Long)
    'Parameters: hWnd - handle to the ComboBox (gotten through ComboBox.hWnd)
    '            WidthPx - width of the list portion of the combo box in pixels

    SendMessageBynum hWnd, CB_SETDROPPEDWIDTH, WidthPx, 0

End Sub

'this function autosizes the list portion of the combo box
Public Sub AutoSizeComboBoxDropDown(cmb As ComboBox)
    'Parameters: cmb - ComboBox control/object to perform the Autosize on

    Dim CurrentEntryWidth As Integer
    Dim PixelLength As Long
    Dim x As Integer
    Dim oFormFont As StdFont
    Dim iScaleMode As Integer 'find the longest string in the list portion of the combobox
        
    'temporarily set the form font to the combo box font
    'First cache the font
    Set oFormFont = cmb.Parent.Font
    
    'now set the combo box font to the form font
    Set cmb.Parent.Font = cmb.Font
    
    'temporarily change the ScaleMode of the form to Pixel
    'first cache the ScaleMode
    iScaleMode = cmb.Parent.ScaleMode
    
    'now set the ScaleMode to Pixel
    cmb.Parent.ScaleMode = vbPixels
    
    'find out the length in pixels of the longest string in the combo box
    For x = 0 To cmb.ListCount - 1
        CurrentEntryWidth = cmb.Parent.TextWidth(cmb.List(x))
        If CurrentEntryWidth > PixelLength Then
            PixelLength = CurrentEntryWidth
        End If
    Next
    
    'then add 10 pixels for a good measure (actually, to account for combobox margins)
    SetComboBoxDropDownWidth cmb.hWnd, PixelLength + 10
    
    'reset the ScaleMode to its original value
    cmb.Parent.ScaleMode = iScaleMode
    
    'reset to the original form font
    Set cmb.Parent.Font = oFormFont

End Sub


