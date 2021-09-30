VERSION 5.00
Begin VB.Form frmSets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extras"
   ClientHeight    =   2550
   ClientLeft      =   6000
   ClientTop       =   3330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   Icon            =   "frmSets.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4455
   Begin VB.Timer tmrSaver 
      Interval        =   1000
      Left            =   120
      Top             =   3600
   End
   Begin VB.Frame fraSets 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Frame fraSec 
         Caption         =   "Security"
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3975
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   255
            Left            =   2760
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            Height          =   255
            Left            =   1440
            TabIndex        =   7
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdGen 
            Caption         =   "&Generate random password level"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtNewLevel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "&AutoSave encoded/decoded text after 1 minute"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "Enable Always on &Top function"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAlwaysOnTop_Click()

    If chkAlwaysOnTop.Value = vbUnchecked Then
     chkAlwaysOnTop.Value = vbUnchecked
     DiaTop = False
    ElseIf chkAlwaysOnTop.Value = vbChecked Then
     chkAlwaysOnTop.Value = vbChecked
     DiaTop = True
    End If

End Sub

Private Sub chkAutoSave_Click()

    If chkAutoSave.Value = vbUnchecked Then
     chkAutoSave.Value = vbUnchecked
     AutoSave = False
     Counter = 60
    ElseIf chkAutoSave.Value = vbChecked Then
     chkAutoSave.Value = vbChecked
     AutoSave = True
    End If

End Sub

Private Sub cmdApply_Click()

    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
     Call nonTop(frmAbout)
    ElseIf DiaTop = False Then
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     Call nonBottom(frmAbout)
    End If
    frmMain.txtPass.Text = txtNewLevel.Text

End Sub

Private Sub cmdCancel_Click()
    frmSets.Hide
End Sub

Private Sub cmdGen_Click()

    txtNewLevel.Text = (CDbl(CInt(((255 + Rnd * Rnd) * (Rnd + 1)) + Rnd) * 50))

End Sub

Private Sub cmdOK_Click()

    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
     Call nonTop(frmAbout)
    ElseIf DiaTop = False Then
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     Call nonBottom(frmAbout)
    End If
    frmMain.txtPass.Text = txtNewLevel.Text
    frmSets.Hide

End Sub

Private Sub Form_Load()

    If chkAlwaysOnTop.Value = True Then
     Call nonTop(Me)
    End If

End Sub

Private Sub tmrSaver_Timer()

    On Error GoTo SaveError

    If AutoSave = True Then
     If Counter > 0 Then
      Counter = Counter - 1
      Debug.Print Counter
     End If
     If Counter = 0 Then
      Open App.path & "\EncryptLog.txt" For Output As #3
       Print #3, frmMain.txtMain.Text
      Close #3
      Counter = 60
     End If
    End If

SaveError:
    If Err.Number <> 0 Then
     Counter = 60
     Exit Sub
    End If
End Sub

