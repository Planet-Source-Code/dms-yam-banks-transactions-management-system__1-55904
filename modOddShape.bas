Attribute VB_Name = "Module1"
Option Explicit

'Only 2 phase types available (Square & Ellipse)
Public Enum PhaseType
    pEllipse = 0
    pSquare = 1
End Enum

'Scalar determines by how much the form is shrunk each loop
Public Enum Scalar
    scalarX = 15
    scalarY = 8
End Enum

'API declarations
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'  MODULE NAME: modOddShape
'
'  PURPOSE: To allow users to unload forms in a couple of cool ways: 1.)The form
'           shrinks into a disappearing square, 2.) The form shrinks into a disappearing
'           ellipse. The function requires that you choose which way you want to
'           'phase out' the form and pass it the form you want to 'phase out' as
'           an object. It will then 'phase out' the form, set its visible property
'           to FALSE, and then unload the form.
'
'  WIN32 API FUNCTIONS USED:
'           Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'           Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'           Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'           Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'  PUBLIC FUNCTIONS:
'           Public Function PhaseOutForm(pType As PhaseType, oForm As Object)
'
'  PUBLIC ENUMS:
'           Public Enum PhaseType
'               pEllipse = 0
'               pSquare = 1
'           End Enum
'
'           Public Enum Scalar
'               scalarX = 15
'               scalarY = 8
'           End Enum

'  PRIVATE FUNCTIONS:
'           NONE
'
'  EXTERNAL DEPENDENCIES:
'           NONE
'
'  USE: (examples only.  Use can vary)
'       WHEN 'CLOSE' TERMINATES THE APP AND UNLOADS THE FORM:
'
'           Private Sub Close_Click()
'               Call PhaseOutForm(pSquare, Me) 'For shrinking square
'           End sub
'
'           Private Sub Close_Click()
'               Call PhaseOutForm(pEllipse, Me) 'For shrinking ellipse
'           End sub
'
'
'  MODIFICATION HISTORY:
'    MODIFIED ON:                    BY:
'    CHANGES:
'    ------------------------------------------------------
'    CREATED BY: Dan McLeran
'    CREATED ON: 29APR1999
'
'----------------------------------------------------------------------------


'Only 2 phase types available (Square & Ellipse)
Public Enum PhaseType
    pEllipse = 0
    pSquare = 1
End Enum

'Scalar determines by how much the form is shrunk each loop
Public Enum Scalar
    scalarX = 15
    scalarY = 8
End Enum

'API declarations
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'This function does all of the work in creating the shrinking shapes and
'merging them with the form passed to the function.
Public Function PhaseOutForm(pType As PhaseType, oForm As Object)
Dim lMyHandle As Long 'Handle to the form passed in.
Dim lMyRgn As Long 'Handle to the created region
Dim l As Long 'variable used to call api functions
Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long 'x,y parameters
Dim iTime As Long 'sleep time in mS used by the API call 'Sleep'.

iTime = 1 'define sleep time in millisec
lMyHandle = oForm.hwnd 'get handle to the form to phase out

'Set up initial values
X1 = 0
Y1 = 0
X2 = oForm.ScaleX(oForm.Width, vbTwips, vbPixels) 'convert form width to pixels
Y2 = oForm.ScaleX(oForm.Height, vbTwips, vbPixels) 'convert form height to pixels
'Loop to 'phase out' the form
Do
    If (pType = pEllipse) Then 'if ellipse was chosen
        lMyRgn = CreateEllipticRgn(X1, Y1, X2, Y2) 'create elliptic region
        If (lMyRgn = 0) Then GoTo ErrCreateRgn 'if error occurs go here
    ElseIf (pType = pSquare) Then 'if square was chosen
        lMyRgn = CreateRectRgn(X1, Y1, X2, Y2) 'create square region
        If (lMyRgn = 0) Then GoTo ErrCreateRgn 'if error occurs go here
    End If
    'Merge the created region with the form passed into the function
    l = SetWindowRgn(lMyHandle, lMyRgn, True)
    If (l = 0) Then GoTo ErrSetRgn 'if error occurs go here
    DoEvents 'Speeds up the visual changes made to the form
    Sleep (iTime) 'Delay the app for the time specified
    X1 = X1 + Scalar.scalarX 'Change the shape of the region by the scalar amts
    Y1 = Y1 + Scalar.scalarY
    X2 = X2 - Scalar.scalarX
    Y2 = Y2 - Scalar.scalarY
Loop Until Y2 - Y1 < 0 'loop until the region is very small
oForm.Visible = False
Unload oForm
Exit Function

ErrCreateRgn:
    MsgBox "An error has occurred while creating the region." & vbCrLf & _
    "Error number: " & Err.Number & " " & Err.Description, vbCritical, App.ProductName
    SetWindowRgn oForm.hwnd, 0, True ' restore original window shape
    Exit Function

ErrSetRgn:
    MsgBox "An error has occurred while setting the window region." & vbCrLf & _
    "Error number: " & Err.Number & " " & Err.Description, vbCritical, App.ProductName
    SetWindowRgn oForm.hwnd, 0, True ' restore original window shape
    Exit Function

End Function


