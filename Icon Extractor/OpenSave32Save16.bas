Attribute VB_Name = "OpenSave32Save16"
Option Explicit
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long


Dim icon_index As Integer
Dim glLargeicons() As Long
Dim glsmallicons() As Long
Dim total_icons As Integer
Private Const DI_NORMAL = 3
Dim fso As New FileSystemObject
Dim ram As New Shell
Dim file_title As String
Dim Reply As String

Const large_icon = 32
Const small_icon = 16

Public Function ShowOpen()

With frmMain.CommonDialog1
    .DialogTitle = "Open"
    .Filter = "All Supported Formats|*.exe;*.dll|Executable Files (*.exe)|*.exe|Application Extension (*.dll)|*.dll"
    .ShowOpen
End With


If Not frmMain.CommonDialog1.FileName = "" Then
   On Error GoTo err1:
    frmMain.Text1 = frmMain.CommonDialog1.FileName
    icon_index = 0
    total_icons = ExtractIconEx(frmMain.CommonDialog1.FileName, -1, 0, 0, 0)
    If total_icons = 0 Then GoTo err1
    ReDim glsmallicons(total_icons)
    ReDim glLargeicons(total_icons)
    frmMain.path1 = frmMain.CommonDialog1.FileName
    file_title = frmMain.CommonDialog1.FileName
    Call geticon
End If
Exit Function
err1:
frmMain.path1 = "This File contains No Icons"
frmMain.icon1 = " 0 Icons "
frmMain.Picture1.Picture = Nothing
frmMain.Picture2.Picture = Nothing
frmMain.Picture3.Picture = Nothing
frmMain.Picture4.Picture = Nothing

End Function
Public Function ShowSave32()
On Error Resume Next

With frmMain.CommonDialog1
    .DialogTitle = "Save As "
    .Filter = "Bitmap (*.bmp)|*.bmp|Cursors (*.cur)|*.cur"
    .DefaultExt = "bmp"
    .FileName = "Untitled"
    .ShowSave
If frmMain.CommonDialog1.FileName = "" Then
    MsgBox "Please enter a title!", vbInformation
Else
 If fso.FileExists(frmMain.CommonDialog1.FileName) = False Then
     SavePicture frmMain.Picture3.Image, frmMain.CommonDialog1.FileName
     Call geticon
     MsgBox "Saved Successfully !!!", vbInformation
 ElseIf fso.FileExists(frmMain.CommonDialog1.FileName) = True Then
     Reply = MsgBox("File Already Exists , Overwrite it Now ? ", vbQuestion + vbYesNo)
     If Reply = vbYes Then
                SavePicture frmMain.Picture3.Image, frmMain.CommonDialog1.FileName
                MsgBox "Saved Successfully !!!", vbInformation
     ElseIf Reply = vbNo Then
                Resume
     End If
 End If
End If
End With

End Function
Public Function ShowSave16()
On Error Resume Next

With frmMain.CommonDialog1
    .DialogTitle = "Save As "
    .Filter = "Bitmap (*.bmp)|*.bmp|Cursors (*.cur)|*.cur"
    .DefaultExt = "bmp"
    .FileName = "Untitled"
    .ShowSave
If frmMain.CommonDialog1.FileName = "" Then
    MsgBox "Please enter a title!", vbInformation
Else
 If fso.FileExists(frmMain.CommonDialog1.FileName) = False Then
     SavePicture frmMain.Picture3.Image, frmMain.CommonDialog1.FileName
     Call geticon
     MsgBox "Saved Successfully !!!", vbInformation
 ElseIf fso.FileExists(frmMain.CommonDialog1.FileName) = True Then
     Reply = MsgBox("File Already Exists , Overwrite it Now ? ", vbQuestion + vbYesNo)
     If Reply = vbYes Then
                SavePicture frmMain.Picture4.Image, frmMain.CommonDialog1.FileName
                MsgBox "Saved Successfully !!!", vbInformation
     ElseIf Reply = vbNo Then
                Resume
     End If
 End If
End If
End With

End Function

Public Function ShowNext()
If icon_index + 1 < total_icons Then
icon_index = icon_index + 1
Call geticon
Else
End If
End Function

Public Function ShowPreview()
If icon_index <= total_icons And icon_index > 0 Then
icon_index = icon_index - 1
Call geticon
Else
End If
End Function

Public Sub geticon()
Call ExtractIconEx(file_title, icon_index, glLargeicons(icon_index), glsmallicons(icon_index), 1)
With frmMain.Picture1
    .Picture = LoadPicture("")
    .AutoRedraw = True
      Call DrawIconEx(.hDC, 0, 0, glLargeicons(icon_index), large_icon, large_icon, 0, 0, DI_NORMAL)
    .Refresh
End With
  With frmMain.Picture2
    .Picture = LoadPicture("")
    .AutoRedraw = True
     Call DrawIconEx(.hDC, 0, 0, glsmallicons(icon_index), small_icon, small_icon, 0, 0, DI_NORMAL)
    .Refresh
End With
  With frmMain.Picture3
    .Picture = LoadPicture("")
    .AutoRedraw = True
     Call DrawIconEx(.hDC, 0, 0, glLargeicons(icon_index), large_icon, large_icon, 0, 0, DI_NORMAL)
    .Refresh
End With
  With frmMain.Picture4
    .Picture = LoadPicture("")
    .AutoRedraw = True
     Call DrawIconEx(.hDC, 0, 0, glsmallicons(icon_index), small_icon, small_icon, 0, 0, DI_NORMAL)
    .Refresh
End With

frmMain.icon1 = "No.Of.Icons : ( " & icon_index + 1 & " of " & total_icons & " )"
End Sub

