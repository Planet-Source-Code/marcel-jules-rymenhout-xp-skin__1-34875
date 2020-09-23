VERSION 5.00
Begin VB.Form FRMPREVIEW 
   Appearance      =   0  'Flat
   BackColor       =   &H00E4EBEB&
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ONTOPBOX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1080
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox CLOSEBOX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2160
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox MAXRESBOX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1800
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox MINBOX 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1440
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox TOPRIGHT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E4EBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox RIGHTTOP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E4EBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox RIGHTBOT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E4EBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox LEFTBOT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E4EBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox LEFTTOP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E4EBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JSSKIN PREVIEW"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Image TOPLEFT 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   240
      Top             =   120
      Width           =   375
   End
   Begin VB.Image BOT 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image RIGHTMID 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image LEFTMID 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image TOPMID 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FRMPREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim frmontop As Boolean

Private Sub CLOSEBOX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLOSEBOX.Picture = FRMMAIN.CLOSEBOX(2).Picture
End Sub

Private Sub CLOSEBOX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLOSEBOX.Picture = FRMMAIN.CLOSEBOX(1).Picture
End Sub

Private Sub CLOSEBOX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLOSEBOX.Picture = FRMMAIN.CLOSEBOX(0).Picture
End Sub

Private Sub Form_DblClick()
FRMMAIN.CommonDialog1.ShowColor
FRMPREVIEW.BackColor = FRMMAIN.CommonDialog1.Color
End Sub

Private Sub Form_Load()
frmontop = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MINBOX.Picture = FRMMAIN.MINBOX(0).Picture
CLOSEBOX.Picture = FRMMAIN.CLOSEBOX(0).Picture
If Me.WindowState = 2 Then
    Me.MAXRESBOX.Picture = FRMMAIN.RESBOX(0).Picture
Else
    Me.MAXRESBOX.Picture = FRMMAIN.MAXBOX(0).Picture
End If
If frmontop = False Then
    ONTOPBOX.Picture = FRMMAIN.ONTOP(0).Picture
End If
End Sub

Private Sub Form_Resize()
DO_skin Me
End Sub

Private Sub MAXRESBOX_Click()
If Me.WindowState = 2 Then
    Me.WindowState = 0
Else
    Me.WindowState = 2
End If
End Sub

Private Sub MAXRESBOX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 0 Then
    MAXRESBOX.Picture = FRMMAIN.MAXBOX(2).Picture
Else
    MAXRESBOX.Picture = FRMMAIN.RESBOX(2).Picture
End If
End Sub

Private Sub MAXRESBOX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 0 Then
    MAXRESBOX.Picture = FRMMAIN.MAXBOX(1).Picture
Else
    MAXRESBOX.Picture = FRMMAIN.RESBOX(1).Picture
End If
End Sub

Private Sub MAXRESBOX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 0 Then
    MAXRESBOX.Picture = FRMMAIN.MAXBOX(0).Picture
Else
    MAXRESBOX.Picture = FRMMAIN.RESBOX(0).Picture
End If
End Sub

Private Sub MINBOX_Click()
Me.WindowState = 1
End Sub

Private Sub MINBOX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MINBOX.Picture = FRMMAIN.MINBOX(2).Picture
End Sub

Private Sub MINBOX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MINBOX.Picture = FRMMAIN.MINBOX(1).Picture
End Sub

Private Sub MINBOX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MINBOX.Picture = FRMMAIN.MINBOX(0).Picture
End Sub

Private Sub ONTOPBOX_Click()
If ONTOPBOX.Picture = FRMMAIN.ONTOP(0).Picture Then
    frmontop = True
    ONTOPBOX.Picture = FRMMAIN.ONTOP(2).Picture
Else
    frmontop = False
    ONTOPBOX.Picture = FRMMAIN.ONTOP(0).Picture
End If
End Sub

Private Sub ONTOPBOX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmontop = False Then
End If
End Sub

Private Sub TOPMID_DblClick()
If Me.WindowState = 0 Then
    Me.WindowState = 2
Else
    Me.WindowState = 0
End If
End Sub

