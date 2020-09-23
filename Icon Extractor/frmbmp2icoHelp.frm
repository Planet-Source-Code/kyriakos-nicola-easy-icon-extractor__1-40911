VERSION 5.00
Begin VB.Form frmbmp2icoHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmbmp2icoHelp.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Application's Path"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   240
      Picture         =   "frmbmp2icoHelp.frx":4F0DA
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   480
      Picture         =   "frmbmp2icoHelp.frx":4F490
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   240
      Picture         =   "frmbmp2icoHelp.frx":5035A
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Browse for Bitmap Image (*.bmp)"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "frmbmp2icoHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
