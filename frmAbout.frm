VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2295
   ClientLeft      =   6240
   ClientTop       =   3930
   ClientWidth     =   4590
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4590
   Begin VB.Frame fraAbout 
      Caption         =   "About my proggy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Caption         =   "Hello! Just a simple Encodind/Decoding program."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image imgAbout 
         Height          =   480
         Left            =   480
         Picture         =   "frmAbout.frx":0442
         Top             =   1080
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MailToMe As Long

Private Sub cmdOK_Click()

    frmMain.Enabled = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub lblMail_Click()
    MailToMe = Shell("Start.exe " & "mailto:giannis@usa.com", 0)
End Sub
