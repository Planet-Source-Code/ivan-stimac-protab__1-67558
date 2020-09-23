Attribute VB_Name = "modFile"
'File Open Dialog Related Declarations
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Type OPENFILENAME      'for GetOpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


' Function used to dispaly fileopen dialog. I didn't used
' MS Common Dialog Control bcozSince i didn't wanted to
' use any 3rd party control...
Public Function ShowFileOpenDialog(lhWndOwner As Long, Optional ByVal sInitDir As String = "", Optional ByVal sFilter As String = "") As String
  On Error Resume Next
    
  Dim utOFName As OPENFILENAME
    
  With utOFName
    
    .lStructSize = Len(utOFName)
      
    .flags = 0
      
    .hwndOwner = lhWndOwner
      
    .hInstance = App.hInstance
      
    If sFilter <> "" Then
      .lpstrFilter = Replace$(sFilter, "|", vbNullChar)
    Else
      .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    End If
    'create a buffer
    .lpstrFile = Space$(254)
    'set the maximum length of a returned file (important)
    .nMaxFile = 255
      
    .lpstrFileTitle = Space$(254)
      
    .nMaxFileTitle = 255
      
    .lpstrInitialDir = sInitDir
    .lpstrTitle = "Open File"

  End With
    
  'Show the dialog
  If GetOpenFileName(utOFName) Then
    ShowFileOpenDialog = Trim$(utOFName.lpstrFile)
  Else
    'Cancel Pressed
    ShowFileOpenDialog = ""
  End If
End Function


