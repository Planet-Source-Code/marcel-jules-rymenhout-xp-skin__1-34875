VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00E4EBEB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4140
   ForeColor       =   &H00E4EBEB&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "mrymenhout@redirack.be"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send any comments or questions to "
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   2565
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hope you like this "
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dont forget to vote for me thanks"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   120
      Picture         =   "FrmAbout.frx":0000
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1"
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JSSKIN BUILDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D67563&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2565
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


