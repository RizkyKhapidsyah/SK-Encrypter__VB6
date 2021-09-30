VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   2610
   ClientLeft      =   5415
   ClientTop       =   3855
   ClientWidth     =   6375
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   6375
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrState 
      Interval        =   1
      Left            =   4560
      Top             =   2400
   End
   Begin VB.TextBox txtPass 
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
      Left            =   5160
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtMain 
      Height          =   1995
      Left            =   0
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note: CHAR encryption method chosed."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2360
      Width           =   6375
   End
   Begin VB.Label lblTotalCharacters 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total: 0  Max: 65535"
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
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblLevel 
      Caption         =   "Password:"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSEseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtras 
         Caption         =   "E&xtras..."
      End
      Begin VB.Menu mnuESseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Ed&it"
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy &All"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCPseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clea&r"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "&Select All"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuCFseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuColor 
         Caption         =   "&Color"
         Begin VB.Menu mnuBackcolor 
            Caption         =   "&Back color..."
         End
         Begin VB.Menu mnuForecolor 
            Caption         =   "&Fore color..."
         End
      End
   End
   Begin VB.Menu mnuEncryption 
      Caption         =   "&Encryption"
      Begin VB.Menu mnuHex 
         Caption         =   "By &Hex"
      End
      Begin VB.Menu mnuChar 
         Caption         =   "By &Char"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEDseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEncrypt 
         Caption         =   "&Encrypt"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDecrypt 
         Caption         =   "&Decrypt"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Counter = 60
End Sub

Private Sub Form_Resize()

    On Error GoTo ResError

    If Not Me.WindowState = 2 Then frmSets.WindowState = Me.WindowState
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 6495 Then Me.Width = 6495
    If Me.Height < 3015 Then Me.Height = 3015
    txtMain.Height = Me.Height - 1275
    txtMain.Width = Me.Width - 120
    txtPass.Left = (txtMain.Left + txtMain.Width) - txtPass.Width
    txtPass.Top = (txtMain.Top + txtMain.Height) + 40
    lblTotalCharacters.Top = txtPass.Top
    lblLevel.Top = txtPass.Top
    lblLevel.Left = (txtPass.Left - lblLevel.Width) - 120
    lblType.Top = txtPass.Top + 300
    lblType.Width = Me.Width - 120

ResError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub mnuAbout_Click()

    Load frmAbout
    frmAbout.Show
    frmMain.Enabled = False

End Sub

Private Sub mnuBackcolor_Click()

    On Error GoTo FontError
    
    With OpenDialog
     frmMain.Enabled = False
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     .CancelError = True
     .RGBResult = txtMain.BackColor
     .ShowColor
    End With
    txtMain.BackColor = OpenDialog.RGBResult
    Me.Show

FontError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    frmMain.Enabled = True
    Me.Show
    Exit Sub

End Sub

Private Sub mnuChar_Click()

    If mnuChar.Checked = False Then
     mnuChar.Checked = True
     mnuHex.Checked = False
    End If

End Sub

Private Sub mnuClear_Click()
    txtMain.Text = ""
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText (txtMain.SelText)
End Sub


Private Sub mnuCopyAll_Click()
    Clipboard.SetText (txtMain.Text)
End Sub

Private Sub mnuCut_Click()

    Clipboard.SetText (txtMain.SelText)
    txtMain.Text = Left(txtMain.Text, txtMain.SelStart) & Right(txtMain.Text, Len(txtMain.Text) - txtMain.SelStart - Len(Clipboard.GetText))

End Sub

Private Sub mnuDecrypt_Click()

    lblTotalCharacters.Caption = "Total: " & Len(txtMain.Text) & "  " & "Max: 65535"
    If mnuChar.Checked = True Then
     txtMain.Text = ChrDecrypt(txtMain.Text, txtPass.Text)
     txtMain.Text = XORDecrypt(txtMain.Text, txtPass.Text)
    ElseIf mnuHex.Checked = True Then
     txtMain.Text = HexDecrypt(txtMain.Text)
     txtMain.Text = XORDecrypt(txtMain.Text, txtPass.Text)
    End If

End Sub

Private Sub mnuDelete_Click()

End Sub

Private Sub mnuEncrypt_Click()

    lblTotalCharacters.Caption = "Total: " & Len(txtMain.Text) & "  " & "Max: 65535"
    If mnuChar.Checked = True Then
     txtMain.Text = XOREncrypt(txtMain.Text, txtPass.Text)
     txtMain.Text = ChrEncrypt(txtMain.Text, txtPass.Text)
    ElseIf mnuHex.Checked = True Then
     txtMain.Text = XOREncrypt(txtMain.Text, txtPass.Text)
     txtMain.Text = HexEncrypt(txtMain.Text)
    End If

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExtras_Click()

    frmSets.txtNewLevel.Text = frmMain.txtPass.Text
    frmSets.Show

End Sub

Private Sub mnuFont_Click()

    On Error GoTo FontError
    
    With OpenDialog
     frmMain.Enabled = False
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     .CancelError = True
     .FontName = txtMain.FontName
     .FontSize = txtMain.FontSize
     .Bold = txtMain.FontBold
     .Italic = txtMain.FontItalic
     .StrikeThru = txtMain.FontStrikethru
     .Underline = txtMain.FontUnderline
     .ShowFont
    End With

    txtMain.FontName = OpenDialog.FontName
    txtMain.FontSize = OpenDialog.FontSize
    txtMain.FontBold = OpenDialog.Bold
    txtMain.FontItalic = OpenDialog.Italic
    txtMain.FontStrikethru = OpenDialog.StrikeThru
    txtMain.FontUnderline = OpenDialog.Underline
    Me.Show

FontError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    frmMain.Enabled = True
    Me.Show
    Exit Sub

End Sub

Private Sub mnuForecolor_Click()

    On Error GoTo FontError

    With OpenDialog
     frmMain.Enabled = False
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     .CancelError = True
     .RGBResult = txtMain.ForeColor
     .ShowColor
    End With
    txtMain.ForeColor = OpenDialog.RGBResult
    Me.Show

FontError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    frmMain.Enabled = True
    Me.Show
    Exit Sub

End Sub

Private Sub mnuHex_Click()

    If mnuHex.Checked = False Then
     mnuHex.Checked = True
     mnuChar.Checked = False
    End If

End Sub

Private Sub mnuOpen_Click()
    
    On Error GoTo OpError

    Dim FName As String

    With OpenDialog
     frmMain.Enabled = False
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     .CancelError = True
     .DialogTitle = "Open a file..."
     .FileName = ""
     .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
     .InitDir = App.path
     .ShowOpen
    End With

    If OpenDialog.FileName <> "" Then
     Dim a As String
     FName = OpenDialog.FileName
      Close #1
      Open FName For Input As #1
       Line Input #1, a$
       txtMain.Text = txtMain.Text & a & vbCrLf
      Close #1
    Else
     GoTo OpError
    End If

   If FileExists(FName) = False Then
    MsgBox "File does not exist!", vbInformation, "Message"
    GoTo OpError
   End If

   If Err.Number = 0 Then NamOp = OpenDialog.FileName
   Me.Show

OpError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    frmMain.Enabled = True
    Me.Show
    Exit Sub

End Sub

Private Sub mnuPaste_Click()
    txtMain.SelText = Clipboard.GetText
End Sub
Private Sub mnuSave_Click()

    On Error GoTo SavError

    If FileExists(NamOp) = True Then
      Close #1
      Open NamOp For Output As #1
       Print #1, txtMain.Text
      Close #1
     ElseIf FileExists(NamOp) = False Then
      Call nonBottom(frmMain)
      Call nonBottom(frmSets)
      If MsgBox("File does not exist! (" & NamOp & ") Create a new one? ", vbQuestion & vbYesNo, "Message") = vbYes Then
       Close #1
       Open NamOp For Output As #1
        Print #1, txtMain.Text
       Close #1
      Else
       Exit Sub
      End If
    End If

SavError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    Me.Show
    Exit Sub

End Sub

Private Sub mnuSaveAs_Click()

    On Error GoTo SavError

    With OpenDialog
     frmMain.Enabled = False
     Call nonBottom(frmMain)
     Call nonBottom(frmSets)
     .CancelError = True
     .DialogTitle = "Save to a file..."
     .FileName = ""
     .Filter = "Text files (*.txt)|*.txt"
     .InitDir = App.path
     .ShowSave
    End With

    If OpenDialog.FileName <> "" Then
    Else
     GoTo SavError
    End If

    If FileExists(OpenDialog.FileName) = True Then
     If MsgBox("Overwrite the existing file?", vbQuestion & vbYesNo, "Question") = vbYes Then
      If Right(OpenDialog.FileName, 3) = "txt" Then
       Open OpenDialog.FileName For Output As #1
        Print #1, txtMain.Text
       Close #1
       GoTo Name
      Else
       Open OpenDialog.FileName & ".txt" For Output As #1
        Print #1, txtMain.Text
       Close #1
       GoTo Name
      End If
     End If
    ElseIf FileExists(OpenDialog.FileName) = False Then
     Open OpenDialog.FileName & ".txt" For Output As #1
      Print #1, txtMain.Text
     Close #1
     GoTo Name
    End If

Name:
    If Err.Number = 0 Then
     If Right(OpenDialog.FileName, 3) = "txt" Then
      NamOp = OpenDialog.FileName
     Else
      NamOp = OpenDialog.FileName & ".txt"
     End If
     frmMain.Enabled = True
     Me.Show
     Exit Sub
    End If

SavError:
    If DiaTop = True Then
     Call nonTop(frmMain)
     Call nonTop(frmSets)
    End If
    frmMain.Enabled = True
    Me.Show
    Exit Sub

End Sub


Private Sub mnuSelAll_Click()

    txtMain.SetFocus
    txtMain.SelStart = 0
    txtMain.SelLength = Len(txtMain.Text)

End Sub


Private Sub tmrState_Timer()

    On Error Resume Next

    If txtMain.Text = "" Then
     mnuCopyAll.Enabled = False
     mnuSelAll.Enabled = False
    ElseIf txtMain.Text <> "" Then
     mnuCopyAll.Enabled = True
     mnuSelAll.Enabled = True
    End If

    If txtMain.SelText = "" Then
     mnuCopy.Enabled = False
     mnuCut.Enabled = False
    ElseIf txtMain.SelText <> "" Then
     mnuCopy.Enabled = True
     mnuCut.Enabled = True
    End If

    If Clipboard.GetText = "" Then
     mnuPaste.Enabled = False
    ElseIf Clipboard.GetText <> "" Then
     mnuPaste.Enabled = True
    End If

    If NamOp <> "" Then
     mnuSave.Enabled = True
    ElseIf NamOp = "" Then
     mnuSave.Enabled = False
    End If

    If mnuChar.Checked = True Then
     lblType.Caption = "Note: CHAR encryption method chosed."
    ElseIf mnuHex.Checked = True Then
     lblType.Caption = "Note: HEX encryption method chosed."
    End If

End Sub

Private Sub txtMain_Change()
    lblTotalCharacters.Caption = "Total: " & Len(txtMain.Text) & "  " & "Max: 65535"
End Sub


