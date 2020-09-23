VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2955
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6825
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   3960
         Top             =   2400
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   4800
         MouseIcon       =   "frmSplash.frx":000C
         Picture         =   "frmSplash.frx":0316
         ScaleHeight     =   600
         ScaleWidth      =   720
         TabIndex        =   7
         Top             =   960
         Width           =   720
      End
      Begin VB.Image imgLogo 
         Height          =   1065
         Left            =   120
         Picture         =   "frmSplash.frx":11E0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright : 2002 - 2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "KN Productions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Warning : All rights reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   2115
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   765
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   2430
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Brought To You By"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   3270
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me
End Sub

Private Sub lblCompany_Click()
Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
Unload Me
End Sub

Private Sub lblCopyright_Click()
Unload Me
End Sub

Private Sub lblProductName_Click()
Unload Me
End Sub

Private Sub lblVersion_Click()
Unload Me
End Sub

Private Sub lblWarning_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
