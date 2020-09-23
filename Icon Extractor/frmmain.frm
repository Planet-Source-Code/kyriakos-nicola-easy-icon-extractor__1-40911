VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extractor By KN . . . . ."
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   3840
      MouseIcon       =   "frmmain.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "frmmain.frx":11D4
      ScaleHeight     =   600
      ScaleWidth      =   720
      TabIndex        =   18
      ToolTipText     =   "Info..."
      Top             =   3840
      Width           =   720
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   3120
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   1215
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   1335
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse *.exe or *.dll file"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4575
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1920
         Top             =   2400
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1800
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         Height          =   600
         Left            =   480
         ScaleHeight     =   600
         ScaleMode       =   0  'User
         ScaleWidth      =   855
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "< &Prev"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Previews Icon"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Next >"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Next Icon"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S&ave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save As... |16x16|"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save As... |32x32|"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Small Icons ( 16 x 16 )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   2520
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Large Icons  (32 x 32)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1770
      End
      Begin VB.Label icon1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No .Of  Icons :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label path1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path : No Files Selected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   135
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   1320
      MouseIcon       =   "frmmain.frx":209E
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "About Program |Please Read!|"
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   360
      MouseIcon       =   "frmmain.frx":23A8
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Exit Program"
      Top             =   3960
      Width           =   645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter a Location of a  "" Dll "" or "" Exe ""  File :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave32 
         Caption         =   "Save (32x32)"
      End
      Begin VB.Menu mnuSave16 
         Caption         =   "Save (16x16)"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnubmp2ico 
         Caption         =   "bmp2ico"
      End
      Begin VB.Menu mnuIcoEdit 
         Caption         =   "IcoEdit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpD 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Icon Extractor..."
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The most of the commands are in a module so take a look at it.
'I decided to write it like this(module) cause I wanted to use
'the same commands for my menu.

'Notice: I have insert a label under the TitleBar.You can't see it
'cause its black like the form.It's used to make visible the menu.

Const large_icon = 32
Const small_icon = 16

Private Sub Command1_Click()
ShowOpen 'see OpenSave32Save16
End Sub

Private Sub Command2_Click()
ShowSave32 'see OpenSave32Save16
End Sub

Private Sub Command3_Click()
ShowSave16 'see OpenSave32Save16
End Sub

Private Sub Command4_Click()
ShowNext 'see OpenSave32Save16
End Sub

Private Sub Command5_Click()
ShowPreview 'see OpenSave32Save16
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
Picture1.BackColor = vbBlack
Picture2.BackColor = vbBlack
Picture1.Height = large_icon * Screen.TwipsPerPixelY
Picture1.Width = large_icon * Screen.TwipsPerPixelX
Picture2.Height = small_icon * Screen.TwipsPerPixelY
Picture2.Width = small_icon * Screen.TwipsPerPixelX
Picture3.Height = Picture1.Height
Picture3.Width = Picture1.Width
Picture4.Height = Picture2.Height
Picture4.Width = Picture2.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Graphic Efects
Label9.ForeColor = &HC0&
Label1.ForeColor = &HC0&
mnuFile.Visible = False
frmMain.Height = 5040
mnuHelp.Visible = False
mnuTools.Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Graphic Efects
Label9.ForeColor = &HC0&
Label1.ForeColor = &HC0&
mnuTools.Visible = False
mnuFile.Visible = False
mnuHelp.Visible = False
frmMain.Height = 5040
End Sub

Private Sub Label1_Click()
    frmAbout.Show 1 'loads frmAbout form
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Graphic Efects
    Label1.ForeColor = vbRed
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mnuTools.Visible = True
    mnuFile.Visible = True
    mnuHelp.Visible = True
    frmMain.Height = 5340
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mnuTools.Visible = False
    mnuFile.Visible = False
    mnuHelp.Visible = False
    frmMain.Height = 5040
End Sub

Private Sub label9_Click()
    Unload Me 'take a wild guess :)
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Graphic Efects
    Label9.ForeColor = vbRed
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1 'again loads frmAbout form
End Sub

Private Sub mnubmp2ico_Click()
frmBmp2Ico.Show 1 'loads frmBmp2Ico form
End Sub

Private Sub mnuExit_Click()
Unload Me 'take a wild guess :)
End Sub

Private Sub mnuHelpD_Click()
frmHelp.Show 1 'loads frmHelp form
End Sub

Private Sub mnuIcoEdit_Click()
MsgBox "It will be added in a newer version!", vbInformation, "Info..."
End Sub

Private Sub mnuInfo_Click()
frmInfo.Show 1 'I think it's not nescesary to explain this again
End Sub

Private Sub mnuOpen_Click()
ShowOpen 'see OpenSave32Save16
End Sub

Private Sub mnuSave16_Click()
ShowSave16 'see OpenSave32Save16
End Sub

Private Sub mnuSave32_Click()
ShowSave32 'see OpenSave32Save16
End Sub

Private Sub Picture5_Click()
    frmInfo.Show 1
End Sub

Private Sub Timer1_Timer()
If Text1.Text <> "" Then
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    mnuSave32.Enabled = True
    mnuSave16.Enabled = True
Else
    mnuSave32.Enabled = False
    mnuSave16.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
End If
End Sub
