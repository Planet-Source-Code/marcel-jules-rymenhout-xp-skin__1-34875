VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMMAIN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "JSSKIN BUILDER Version 1.1"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Borders"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Control Box"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Shape1"
      Tab(1).Control(8)=   "Line1"
      Tab(1).Control(9)=   "Line2"
      Tab(1).Control(10)=   "Line3"
      Tab(1).Control(11)=   "Line4"
      Tab(1).Control(12)=   "Line5"
      Tab(1).Control(13)=   "Label9"
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(15)=   "Label11"
      Tab(1).Control(16)=   "Label18"
      Tab(1).Control(17)=   "Label19"
      Tab(1).Control(18)=   "Label20"
      Tab(1).Control(19)=   "MINBOX(0)"
      Tab(1).Control(20)=   "MAXBOX(0)"
      Tab(1).Control(21)=   "RESBOX(0)"
      Tab(1).Control(22)=   "CLOSEBOX(0)"
      Tab(1).Control(23)=   "MINBOX(1)"
      Tab(1).Control(24)=   "MAXBOX(1)"
      Tab(1).Control(25)=   "RESBOX(1)"
      Tab(1).Control(26)=   "CLOSEBOX(1)"
      Tab(1).Control(27)=   "MINBOX(2)"
      Tab(1).Control(28)=   "MAXBOX(2)"
      Tab(1).Control(29)=   "RESBOX(2)"
      Tab(1).Control(30)=   "CLOSEBOX(2)"
      Tab(1).Control(31)=   "TXTRIGHT"
      Tab(1).Control(32)=   "TXTGAP"
      Tab(1).Control(33)=   "TXTTOP"
      Tab(1).Control(34)=   "ONTOP(2)"
      Tab(1).Control(35)=   "ONTOP(0)"
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Caption"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "Label12"
      Tab(2).Control(4)=   "Label1"
      Tab(2).Control(5)=   "Label16"
      Tab(2).Control(6)=   "Text1"
      Tab(2).Control(7)=   "Picture1"
      Tab(2).Control(8)=   "TXTCTOP"
      Tab(2).Control(9)=   "TXTCLEFT"
      Tab(2).ControlCount=   10
      Begin VB.PictureBox ONTOP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   -74400
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   3120
         Width           =   375
      End
      Begin VB.PictureBox ONTOP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   -73320
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   38
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox TXTCLEFT 
         Height          =   285
         Left            =   -73680
         TabIndex        =   31
         Text            =   "10"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TXTCTOP 
         Height          =   285
         Left            =   -73680
         TabIndex        =   30
         Text            =   "5"
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   -73680
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   29
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74880
         TabIndex        =   28
         Text            =   "JSSKIN PREVIEW"
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox TXTTOP 
         Height          =   285
         Left            =   -72840
         TabIndex        =   17
         Text            =   "2"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox TXTGAP 
         Height          =   285
         Left            =   -72840
         TabIndex        =   16
         Text            =   "10"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox TXTRIGHT 
         Height          =   285
         Left            =   -72840
         TabIndex        =   15
         Text            =   "10"
         Top             =   3720
         Width           =   855
      End
      Begin VB.PictureBox CLOSEBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   -72240
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox RESBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   -72840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox MAXBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   -73440
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox MINBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   -74040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox CLOSEBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   -72240
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox RESBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   -72840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox MAXBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   -73440
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox MINBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   -74040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox CLOSEBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   -72240
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox RESBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   -72840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox MAXBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   -73440
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox MINBOX 
         AutoSize        =   -1  'True
         BackColor       =   &H00FBB17D&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   -74040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Down"
         Height          =   195
         Left            =   -73920
         TabIndex        =   42
         Top             =   3120
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Up"
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   3120
         Width           =   210
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Icons For Always on top"
         Height          =   195
         Left            =   -74760
         TabIndex        =   40
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "DoubleClick on preview form to change its back color"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -74880
         TabIndex        =   37
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "The Caption is just a preview this is not saved in the skin as JSSKIN gets its caption from the form itself...."
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -74880
         TabIndex        =   36
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Caption Left"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Caption Top"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Caption Color"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         Height          =   195
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ControlBox from Top"
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   4680
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Distance Between Icons"
         Height          =   195
         Left            =   -74760
         TabIndex        =   25
         Top             =   4200
         Width           =   1740
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ControlBox from Right"
         Height          =   195
         Left            =   -74760
         TabIndex        =   24
         Top             =   3720
         Width           =   1530
      End
      Begin VB.Line Line5 
         X1              =   -74160
         X2              =   -71760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line4 
         X1              =   -74160
         X2              =   -71760
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   -72360
         X2              =   -72360
         Y1              =   720
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   -72960
         X2              =   -72960
         Y1              =   720
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   -73560
         X2              =   -73560
         Y1              =   720
         Y2              =   2520
      End
      Begin VB.Shape Shape1 
         Height          =   1815
         Left            =   -74160
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Close"
         Height          =   195
         Left            =   -72240
         TabIndex        =   23
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Restore"
         Height          =   195
         Left            =   -72960
         TabIndex        =   22
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Max"
         Height          =   195
         Left            =   -73440
         TabIndex        =   21
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Left            =   -74040
         TabIndex        =   20
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Highlight"
         Height          =   195
         Left            =   -74880
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Normal"
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Down"
         Height          =   195
         Left            =   -74880
         TabIndex        =   27
         Top             =   2040
         Width           =   420
      End
   End
   Begin VB.PictureBox SAMPLECONTAINER 
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   3600
      ScaleHeight     =   4515
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "View"
      Begin VB.Menu MnuRestore 
         Caption         =   "Reset Position"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FRMMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private pb As PropertyBag

Sub CLEARBORDER()
    FRMPREVIEW.TOPLEFT.BorderStyle = 0
    FRMPREVIEW.TOPMID.BorderStyle = 0
    FRMPREVIEW.TOPRIGHT.BorderStyle = 0
    FRMPREVIEW.LEFTTOP.BorderStyle = 0
    FRMPREVIEW.LEFTMID.BorderStyle = 0
    FRMPREVIEW.LEFTBOT.BorderStyle = 0
    FRMPREVIEW.RIGHTTOP.BorderStyle = 0
    FRMPREVIEW.RIGHTMID.BorderStyle = 0
    FRMPREVIEW.RIGHTBOT.BorderStyle = 0
    FRMPREVIEW.BOT.BorderStyle = 0
End Sub

Private Sub CLOSEBOX_Click(Index As Integer)
Me.CommonDialog1.ShowOpen
CLOSEBOX(Index).Picture = LoadPicture(CommonDialog1.FileName)
FRMPREVIEW.CLOSEBOX.Picture = CLOSEBOX(0).Picture
End Sub


Private Sub Form_Load()
SetParent FRMPREVIEW.hWnd, Me.SAMPLECONTAINER.hWnd
FRMPREVIEW.Show
With Me.List1
.AddItem "TOPLEFT"
.AddItem "TOPMID"
.AddItem "TOPRIGHT"
.AddItem "LEFTTOP"
.AddItem "LEFTMID"
.AddItem "LEFTBOT"
.AddItem "RIGHTTOP"
.AddItem "RIGHTMID"
.AddItem "RIGHTBOT"
.AddItem "BOTTOM"
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Me.SAMPLECONTAINER
.Width = Me.ScaleWidth - .Left
.Height = Me.ScaleHeight - .Top
End With
FRMPREVIEW.Move (Me.SAMPLECONTAINER.Width / 2) - (FRMPREVIEW.Width / 2), (Me.SAMPLECONTAINER.Height / 2) - (FRMPREVIEW.Height / 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload FRMPREVIEW
End Sub

Private Sub List1_Click()
CLEARBORDER
If Me.List1.Text = "TOPLEFT" Then
    FRMPREVIEW.TOPLEFT.BorderStyle = 1
ElseIf Me.List1.Text = "TOPMID" Then
    FRMPREVIEW.TOPMID.BorderStyle = 1
ElseIf Me.List1.Text = "TOPRIGHT" Then
    FRMPREVIEW.TOPRIGHT.BorderStyle = 1
ElseIf Me.List1.Text = "LEFTTOP" Then
    FRMPREVIEW.LEFTTOP.BorderStyle = 1
ElseIf Me.List1.Text = "LEFTMID" Then
    FRMPREVIEW.LEFTMID.BorderStyle = 1
ElseIf Me.List1.Text = "LEFTBOT" Then
    FRMPREVIEW.LEFTBOT.BorderStyle = 1
ElseIf Me.List1.Text = "RIGHTTOP" Then
    FRMPREVIEW.RIGHTTOP.BorderStyle = 1
ElseIf Me.List1.Text = "RIGHTMID" Then
    FRMPREVIEW.RIGHTMID.BorderStyle = 1
ElseIf Me.List1.Text = "RIGHTBOT" Then
    FRMPREVIEW.RIGHTBOT.BorderStyle = 1
ElseIf Me.List1.Text = "BOTTOM" Then
    FRMPREVIEW.BOT.BorderStyle = 1
End If
End Sub

Private Sub List1_DblClick()
Me.CommonDialog1.ShowOpen
If Me.List1.Text = "TOPLEFT" Then
    FRMPREVIEW.TOPLEFT.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "TOPMID" Then
    FRMPREVIEW.TOPMID.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "TOPRIGHT" Then
    FRMPREVIEW.TOPRIGHT.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "LEFTTOP" Then
    FRMPREVIEW.LEFTTOP.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "LEFTMID" Then
    FRMPREVIEW.LEFTMID.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "LEFTBOT" Then
    FRMPREVIEW.LEFTBOT.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "RIGHTTOP" Then
    FRMPREVIEW.RIGHTTOP.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "RIGHTMID" Then
    FRMPREVIEW.RIGHTMID.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "RIGHTBOT" Then
    FRMPREVIEW.RIGHTBOT.Picture = LoadPicture(Me.CommonDialog1.FileName)
ElseIf Me.List1.Text = "BOTTOM" Then
    FRMPREVIEW.BOT.Picture = LoadPicture(Me.CommonDialog1.FileName)
End If
CLEARBORDER
DO_skin FRMPREVIEW
End Sub

Private Sub MAXBOX_Click(Index As Integer)
Me.CommonDialog1.ShowOpen
MAXBOX(Index).Picture = LoadPicture(CommonDialog1.FileName)
If FRMPREVIEW.WindowState = 2 Then
    FRMPREVIEW.MAXRESBOX.Picture = Me.RESBOX(0).Picture
Else
    FRMPREVIEW.MAXRESBOX.Picture = Me.MAXBOX(0).Picture
End If
End Sub

Private Sub MINBOX_Click(Index As Integer)
Me.CommonDialog1.ShowOpen
MINBOX(Index).Picture = LoadPicture(CommonDialog1.FileName)
FRMPREVIEW.MINBOX.Picture = MINBOX(0).Picture
End Sub

Private Sub MnuAbout_Click()
FrmAbout.Show 1
End Sub

Private Sub MnuExit_Click()
End
End Sub

Private Sub MnuOpen_Click()
Dim varTemp As Variant
   Dim byteArr() As Byte
 On Error Resume Next
   Set pb = New PropertyBag
  CommonDialog1.ShowOpen
   Open CommonDialog1.FileName For Binary As #1
   Get #1, , varTemp
   Close #1
   byteArr = varTemp
   pb.Contents = byteArr
   With pb
   FRMPREVIEW.TOPLEFT.Picture = .ReadProperty("TOPLEFT")
   FRMPREVIEW.TOPMID.Picture = .ReadProperty("TOPMID")
   FRMPREVIEW.TOPRIGHT.Picture = .ReadProperty("TOPRIGHT")
   Me.MINBOX(0).Picture = .ReadProperty("MIN")
    Me.MAXBOX(0).Picture = .ReadProperty("MAX")
   Me.CLOSEBOX(0).Picture = .ReadProperty("CLOSE")
   Me.MINBOX(1).Picture = .ReadProperty("MIN3")
   Me.MAXBOX(1).Picture = .ReadProperty("MAX3")
   Me.CLOSEBOX(1).Picture = .ReadProperty("CLOSE3")
   Me.MINBOX(2).Picture = .ReadProperty("MIN2")
   Me.MAXBOX(2).Picture = .ReadProperty("MAX2")
   Me.CLOSEBOX(2).Picture = .ReadProperty("CLOSE2")
   Me.RESBOX(0).Picture = .ReadProperty("RES1")
   Me.RESBOX(1).Picture = .ReadProperty("RES2")
   Me.RESBOX(2).Picture = .ReadProperty("RES3")
   Me.ONTOP(0).Picture = .ReadProperty("ONTOP1")
   Me.ONTOP(2).Picture = .ReadProperty("ONTOP3")
   FRMPREVIEW.LEFTTOP.Picture = .ReadProperty("LEFTTOP")
   FRMPREVIEW.LEFTMID.Picture = .ReadProperty("LEFTMID")
   FRMPREVIEW.LEFTBOT.Picture = .ReadProperty("LEFTBOT")
   FRMPREVIEW.RIGHTTOP.Picture = .ReadProperty("RIGHTTOP")
   FRMPREVIEW.RIGHTMID.Picture = .ReadProperty("RIGHTMID")
   FRMPREVIEW.RIGHTBOT.Picture = .ReadProperty("RIGHTBOT")
    FRMPREVIEW.BOT.Picture = .ReadProperty("BOTTOM")
   FRMPREVIEW.Label1.ForeColor = .ReadProperty("FORECOLOR")
   Me.Picture1.BackColor = .ReadProperty("FORECOLOR")
    Me.TXTGAP.Text = .ReadProperty("ICONSPACE")
   Me.TXTRIGHT.Text = .ReadProperty("FROMRIGHT")
   Me.TXTTOP.Text = .ReadProperty("FROMTOP")
    Me.TXTCTOP.Text = .ReadProperty("YOFFSET")
   Me.TXTCLEFT.Text = .ReadProperty("XOFFSET")
   End With
   FRMPREVIEW.MINBOX.Picture = FRMMAIN.MINBOX(0).Picture
   FRMPREVIEW.CLOSEBOX.Picture = FRMMAIN.CLOSEBOX(0).Picture
   FRMPREVIEW.ONTOPBOX.Picture = FRMMAIN.ONTOP(0).Picture
    If FRMPREVIEW.WindowState = 2 Then
        FRMPREVIEW.MAXRESBOX.Picture = FRMMAIN.RESBOX(0).Picture
    Else
        FRMPREVIEW.MAXRESBOX.Picture = FRMMAIN.MAXBOX(0).Picture
    End If
   DO_skin FRMPREVIEW
End Sub

Private Sub MnuRestore_Click()
FRMPREVIEW.WindowState = 0
FRMPREVIEW.Height = 3360
FRMPREVIEW.Width = 4800
FRMPREVIEW.Move (Me.SAMPLECONTAINER.Width / 2) - (FRMPREVIEW.Width / 2), (Me.SAMPLECONTAINER.Height / 2) - (FRMPREVIEW.Height / 2)
End Sub

Private Sub MnuSave_Click()
   Dim varTemp As Variant
   On Error GoTo errhandler
   CommonDialog1.ShowSave
   Set pb = New PropertyBag
 With pb
   .WriteProperty "TOPLEFT", FRMPREVIEW.TOPLEFT.Picture
   .WriteProperty "TOPMID", FRMPREVIEW.TOPMID.Picture
   .WriteProperty "TOPRIGHT", FRMPREVIEW.TOPRIGHT.Picture
   .WriteProperty "MIN", Me.MINBOX(0).Picture
   .WriteProperty "MAX", Me.MAXBOX(0).Picture
   .WriteProperty "CLOSE", Me.CLOSEBOX(0).Picture
   .WriteProperty "MIN3", Me.MINBOX(1).Picture
   .WriteProperty "MAX3", Me.MAXBOX(1).Picture
   .WriteProperty "CLOSE3", Me.CLOSEBOX(1).Picture
   .WriteProperty "MIN2", Me.MINBOX(2).Picture
   .WriteProperty "MAX2", Me.MAXBOX(2).Picture
   .WriteProperty "CLOSE2", Me.CLOSEBOX(2).Picture
   .WriteProperty "RES1", Me.RESBOX(0).Picture
   .WriteProperty "RES2", Me.RESBOX(1).Picture
   .WriteProperty "RES3", Me.RESBOX(2).Picture
   .WriteProperty "ONTOP1", Me.ONTOP(0).Picture
   .WriteProperty "ONTOP3", Me.ONTOP(2).Picture
   .WriteProperty "LEFTTOP", FRMPREVIEW.LEFTTOP.Picture
   .WriteProperty "LEFTMID", FRMPREVIEW.LEFTMID.Picture
   .WriteProperty "LEFTBOT", FRMPREVIEW.LEFTBOT.Picture
   .WriteProperty "RIGHTTOP", FRMPREVIEW.RIGHTTOP.Picture
   .WriteProperty "RIGHTMID", FRMPREVIEW.RIGHTMID.Picture
   .WriteProperty "RIGHTBOT", FRMPREVIEW.RIGHTBOT.Picture
   .WriteProperty "BOTTOM", FRMPREVIEW.BOT.Picture
   .WriteProperty "FORECOLOR", FRMPREVIEW.Label1.ForeColor
   .WriteProperty "ICONSPACE", Me.TXTGAP.Text
   .WriteProperty "FROMRIGHT", Me.TXTRIGHT.Text
   .WriteProperty "FROMTOP", Me.TXTTOP.Text
   .WriteProperty "YOFFSET", Me.TXTCTOP.Text
   .WriteProperty "XOFFSET", Me.TXTCLEFT.Text
End With
   varTemp = pb.Contents
   Open CommonDialog1.FileName For Binary As #1
   Put #1, , varTemp
   Close #1
errhandler:
   Exit Sub
End Sub

Private Sub ONTOP_Click(Index As Integer)
Me.CommonDialog1.ShowOpen
ONTOP(Index).Picture = LoadPicture(CommonDialog1.FileName)
FRMPREVIEW.ONTOPBOX.Picture = ONTOP(0).Picture
End Sub

Private Sub Picture1_Click()
Me.CommonDialog1.ShowColor
Picture1.BackColor = Me.CommonDialog1.Color
FRMPREVIEW.Label1.ForeColor = Picture1.BackColor
End Sub

Private Sub RESBOX_Click(Index As Integer)
Me.CommonDialog1.ShowOpen
RESBOX(Index).Picture = LoadPicture(CommonDialog1.FileName)
FRMPREVIEW.MAXRESBOX.Picture = RESBOX(0).Picture
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
CLEARBORDER
End Sub

Private Sub Text1_Change()
FRMPREVIEW.Label1.Caption = Text1.Text
End Sub

Private Sub TXTCLEFT_Change()
DO_skin FRMPREVIEW
End Sub

Private Sub TXTCTOP_Change()
DO_skin FRMPREVIEW
End Sub

Private Sub TXTGAP_Change()
DO_skin FRMPREVIEW
End Sub

Private Sub TXTRIGHT_Change()
DO_skin FRMPREVIEW
End Sub

Private Sub TXTTOP_Change()
DO_skin FRMPREVIEW
End Sub

