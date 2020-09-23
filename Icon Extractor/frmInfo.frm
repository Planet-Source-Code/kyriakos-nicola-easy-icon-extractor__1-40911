VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info . . ."
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bye"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Form"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'the actual text to scroll. This could also be loaded in from a text file
Const ScrollText As String = "Icon Extractor" & vbCrLf & _
                             vbCrLf & vbCrLf & _
                             "Author : MySelf" & vbCrLf & _
                             "Producer: MySelf" & vbCrLf & _
                             "Executive Producer: MySelf" & _
                             vbCrLf & "Main programmer: MySelf" & _
                             vbCrLf & "Main graphic artist: MySelf" & _
                             vbCrLf & vbCrLf & _
                             "Special Thanks to MySelf !!!" & _
                             vbCrLf & vbCrLf & _
                             "- Kyriakos Nicola -"
                             
Dim EndingFlag As Boolean

Private Sub Command1_Click()
Unload Me
EndingFlag = True
End Sub

Private Sub Form_Activate()
RunMain
End Sub

Private Sub Form_Load()

EndingFlag = False
picScroll.ForeColor = vbGreen
picScroll.FontSize = 14

End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmInfo.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hDC, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hDC, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set frmAbout = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    EndingFlag = True

End Sub

