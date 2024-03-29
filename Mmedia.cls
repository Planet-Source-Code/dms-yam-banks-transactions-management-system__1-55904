VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mmedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'   Name    :   MMedia.cls
'   Author  :   Peter Wright, For BG2VB4 & BG2VB5
'
'   Notes   :   A multimedia class, which when turned
'           :   into an object lets you load and play
'           :   multimedia files, such as sound and
'           :   video.
'-----------------------------------------------------

' -=-=-=- PROPERTIES -=-=-=-
' Filename      Determines the name of the current file
' Length        The length of the file (Read Only)
' Position      The current position through the file
' Status        The current status of the object (Read Only)
' Wait          True/False...tells VB to wait until play done

' -=-=-=- METHODS -=-=-=-=-
' mmOpen <Filename>   Opens the requested filename
' mmClose             Closes the current file
' mmPause             Pauses playback of the current file
' mmStop              Stops playback ready for closedown
' mmSeek <Position>   Seeks to a position in the file
' mmPlay              Plays the open file

'-------------------------------------------------------------
' NOTES
' -----
'
' Open a file, then play it. Pause it in response to a request
' from the user. Stop if you intend to seek to the start and
' play again. Close when you no longer want to play the file
'--------------------------------------------------------------

Private sAlias As String        ' Used internally to give an alias name to
                          ' the multimedia resource

Private sFilename As String     ' Holds the filename internally
Private nLength As Single       ' Holds the length of the filename
                          ' internally
Private nPosition As Single     ' Holds the current position internally
Private sStatus As String       ' Holds the current status as a string
Private bWait As Boolean        ' Determines if VB should wait until play
                        ' is complete before returning.

'------------ API DECLARATIONS -------------
'note that this is all one code line:
Private Declare Function mciSendString Lib "winmm.dll" _
   Alias "mciSendStringA" (ByVal lpstrCommand As String, _
   ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
   ByVal hwndCallback As Long) As Long

Public Sub mmOpen(ByVal sTheFile As String)

    ' Declare a variable to hold the value returned by mciSendString
    Dim nReturn As Long
    
    ' Declare a string variable to hold the file type
    Dim sType As String

    ' Opens the specified multimedia file, and closes any
    ' other that may be open
    If sAlias <> "" Then
        mmClose
    End If
    
    ' Determine the type of file from the file extension
    'Select Case UCase$(Right$(sTheFile, 3))
       'Case "WAV"
          sType = "Waveaudio"
       'Case "AVI"
        '  sType = "AviVideo"
       'Case "MID"
        '  sType = "Sequencer"
       'Case Else
          ' If the file extension is not known then exit the subroutine
          'Exit Sub
    'End Select
    sAlias = "wav" & Minute(Now)

    ' At this point there is no file open, and we have determined the
    ' file type. Now would be a good time to open the new file.
    ' Note: if the name contains a space we have to enclose it in quotes
    If InStr(sTheFile, " ") Then sTheFile = Chr(34) & sTheFile & Chr(34)
    nReturn = mciSendString("Open " & sTheFile & " ALIAS " & sAlias _
            & " TYPE " & sType & " wait", "", 0, 0)
End Sub

Public Sub mmClose()
    ' Closes the currently opened multimedia file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long

    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
    
End Sub

Public Sub mmPause()
    ' Pause playback of the file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Pause " & sAlias, "", 0, 0)

End Sub

Public Sub mmPlay()
    ' Plays the currently open file, from the current position

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    ' If there is no file currently open, then exit the routine
    If sAlias = "" Then Exit Sub
    
    ' Now play the file
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("Play " & sAlias, "", 0, 0)
    End If
End Sub

Public Sub mmStop()
    ' Stop using a file totally, be it playing or whatever

    ' Declare a variable to hold the return value from mciSendString
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
    
End Sub

Public Sub mmSeek(ByVal nPosition As Single)
    ' Seeks to a specific position within the file

    ' Declare a variable to hold the return value from the mciSendString
    ' function
    Dim nReturn As Long
    
    nReturn = mciSendString("Seek " & sAlias & " to " & nPosition, "", 0, 0)

End Sub

Property Get Filename() As String
' Routine to return a value when the programmer asks the
' object for the value of its Filename property
    Filename = sFilename
End Property

Property Let Filename(ByVal sTheFile As String)
' Routine to set the value of the filename property, should the programmer
' wish to do so. This implies that the programmer actually wants to open
' a file as well so control is passed to the mmOpen routine
   mmOpen sTheFile
End Property

Property Get Wait() As Boolean
' Routine to return the value of the object's wait property.
   Wait = bWait
End Property

Property Let Wait(bWaitValue As Boolean)
' Routine to set the value of the object's wait property
   bWait = bWaitValue
End Property

Property Get Length() As Single
   ' Routine to return the length of the currently opened multimedia file

   ' Declare a variable to hold the return value from the mciSendString
   Dim nReturn As Long, nLength As Integer

   ' Declare a string to hold the returned length from the mci Status call
   Dim sLength As String * 255
    
   ' If there is no file open then return 0
   If sAlias = "" Then
      Length = 0
      Exit Property
   End If

  nReturn = mciSendString("Status " & sAlias & " length", sLength, 255, 0)
  nLength = InStr(sLength, Chr$(0))
  'Length = Left$(sLength, (nLength - 1))
End Property

Property Let Position(ByVal nPosition As Single)
' Sets the Position property effectively by seeking
    mmSeek nPosition
End Property

Property Get Position() As Single
   ' Returns the current position in the file
    
   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer
    
   ' Declare a variable to hold the position returned
   ' by the mci Status position command
   Dim sPosition As String * 255

   ' If there is no file currently opened then exit the subroutine
   If sAlias = "" Then Exit Property
    
   ' Get the position and return
   nReturn = mciSendString("Status " & sAlias & " position", sPosition, 255, 0)
   nLength = InStr(sPosition, Chr$(0))
   'Position = Val(Left$(sPosition, nLength - 1))

End Property

Property Get Status() As String
   ' Returns the playback/record status of the current file

   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer
    
   ' Declare a variable to hold the return string from mciSendString
   Dim sStatus As String * 255
    
   ' If there is no file currently opened, then exit the subroutine
   If sAlias = "" Then Exit Property

   nReturn = mciSendString("Status " & sAlias & " mode", sStatus, 255, 0)
    
   nLength = InStr(sStatus, Chr$(0))
   'Status = Left$(sStatus, nLength - 1)
    
End Property


