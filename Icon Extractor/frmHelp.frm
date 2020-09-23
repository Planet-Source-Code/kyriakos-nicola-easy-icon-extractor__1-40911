VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   Picture         =   "frmHelp.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prev. Icon"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save as bmp 16x16 and 32x32"
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next Icon"
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Browse for a *.exe or a *.dll file"
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   2145
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   3
      Left            =   120
      Picture         =   "frmHelp.frx":4F0DA
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   2
      Left            =   480
      Picture         =   "frmHelp.frx":4FFB8
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   1
      Left            =   480
      Picture         =   "frmHelp.frx":509CE
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "frmHelp.frx":513E4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   840
      Picture         =   "frmHelp.frx":522C2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
