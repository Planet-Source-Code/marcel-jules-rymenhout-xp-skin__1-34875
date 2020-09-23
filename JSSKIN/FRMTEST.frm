VERSION 5.00
Object = "*\AJSSKIN.vbp"
Begin VB.Form FRMTEST 
   BackColor       =   &H00D67563&
   BorderStyle     =   0  'None
   Caption         =   "Js Skin Preview"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "FRMTEST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Perfect Whistler"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Luna"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin JSSKIN.JSBORDER JSBORDER3 
      Align           =   3  'Align Left
      Height          =   3990
      Left            =   0
      TabIndex        =   3
      Top             =   405
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   7038
      BORDERTYPE      =   1
   End
   Begin JSSKIN.JSBORDER JSBORDER2 
      Align           =   4  'Align Right
      Height          =   3990
      Left            =   4575
      TabIndex        =   2
      Top             =   405
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   7038
      BORDERTYPE      =   2
   End
   Begin JSSKIN.JSBORDER JSBORDER1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   1
      Top             =   4395
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   185
      BORDERTYPE      =   3
   End
   Begin JSSKIN.JSCAPTION JSCAPTION1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
      ONTOP           =   -1  'True
      SHOWONTOP       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FRMTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.JSCAPTION1.Path = App.Path & "\luna.jss"
Me.JSBORDER1.Path = App.Path & "\luna.jss"
Me.JSBORDER2.Path = App.Path & "\luna.jss"
Me.JSBORDER3.Path = App.Path & "\luna.jss"
Me.JSCAPTION1.REDRAW
End Sub

Private Sub Command2_Click()
Me.JSCAPTION1.Path = App.Path & "\PW.jss"
Me.JSBORDER1.Path = App.Path & "\PW.jss"
Me.JSBORDER2.Path = App.Path & "\PW.jss"
Me.JSBORDER3.Path = App.Path & "\PW.jss"
Me.JSCAPTION1.REDRAW
End Sub

Private Sub Form_Load()
Me.JSCAPTION1.Path = App.Path & "\luna.jss"
Me.JSBORDER1.Path = App.Path & "\luna.jss"
Me.JSBORDER2.Path = App.Path & "\luna.jss"
Me.JSBORDER3.Path = App.Path & "\luna.jss"

End Sub


