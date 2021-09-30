Attribute VB_Name = "mdlEnc"
Option Explicit

Public Enc As Boolean
Public Dec As Boolean
Public AutoSave As Boolean
Public DiaTop As Boolean

Public OpenDialog As New clsCDEx
Public Counter As Integer
Public NamOp As String

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const conSwNormal = 1
Public Const HWND_TOPMOST = -1

Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function XOREncrypt(Text As String, Pass As String) As String

    On Error GoTo XORError
    Dim passLen As Double

    For passLen = 0 To Len(Text) - 1
     XOREncrypt = XOREncrypt + CStr(Asc(Mid$(Text, passLen + 1, 1)) Xor Asc(Mid$(Pass, (passLen Mod Len(Pass)) + 1, 1))) & " "
    Next passLen
    XOREncrypt = Trim$(XOREncrypt)

XORError:
    If Err.Number <> 0 Then
     XOREncrypt = Text
     Exit Function
    End If

End Function
Public Function XORDecrypt(Text As String, Pass As String) As String

    On Error GoTo XORError
    Dim lenText As Double
    Dim HexArray As Variant
    HexArray = Split(Text, " ")

    For lenText = 0 To UBound(HexArray)
     XORDecrypt = XORDecrypt + Chr(Int(HexArray(lenText)) Xor Asc(Mid$(Pass, (lenText Mod Len(Pass)) + 1, 1)))
    Next lenText

XORError:
    If Err.Number <> 0 Then
     XORDecrypt = Text
     Exit Function
    End If

End Function
Public Function nonTop(frmToTop As Form)
    
    On Error GoTo TopError
    Dim onTop%
    onTop% = SetWindowPos(frmToTop.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)

TopError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Function FileExists(FileName As String) As String

    On Error Resume Next
    FileExists = Dir(FileName, vbHidden) <> ""

End Function
Public Function nonBottom(frmToBottom As Form)
    
    On Error GoTo BotError
    Dim onTop%
    onTop% = SetWindowPos(frmToBottom.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)

BotError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Function ChrDecrypt(Text As String, Optional Pass As String) As String

    On Error GoTo EncError
    Dim LastPass As Double, i As Double, a As Double
    
    If Len(Pass) > 0 Then
     LastPass = Asc(Mid$(Pass, Len(Pass), Len(Pass)))
    ElseIf Len(Pass) < 0 Then
     LastPass = Asc(Mid$(Pass, Len(Text), 1))
    End If

    For i = 1 To Len(Text)
     a = Asc(Mid$(Text, i, 1)) + 2 * LastPass
     If a > 255 Then a = a - ((a \ 255) * 255)
     LastPass = a
     ChrDecrypt = ChrDecrypt & Chr$(a)
    Next i

EncError:
    If Err.Number <> 0 Then
     ChrDecrypt = Text
     Exit Function
    End If

End Function
Public Function HexEncrypt(Text As String) As String

    On Error GoTo HexError
    Dim HexCount As Double
    Dim HexTemp As String

    For HexCount = 1 To Len(Text)
     HexTemp = Hex$(Asc(Mid$(Text, HexCount, 1)))
     If Len(HexTemp) < 2 Then HexTemp = "0" & HexTemp
     HexTemp = Right(HexTemp, 1) & Left(HexTemp, 1)
     HexEncrypt = HexEncrypt & HexTemp
    Next HexCount

HexError:
    If Err.Number <> 0 Then
     HexEncrypt = Text
     Exit Function
    End If

End Function
Public Function HexDecrypt(Text As String) As String

    On Error GoTo HexError
    Dim HexDecryptCount As Double

    For HexDecryptCount = 1 To Len(Text) Step 2
     HexDecrypt = HexDecrypt & Chr$(Val("&H" & (Right(Mid$(Text, HexDecryptCount, 2), 1) & Left(Mid$(Text, HexDecryptCount, 2), 1))))
    Next HexDecryptCount

HexError:
    If Err.Number <> 0 Then
     HexDecrypt = Text
     Exit Function
    End If

End Function
Public Function ChrEncrypt(Text As String, Optional Pass As String) As String

    On Error GoTo DecError
    Dim FirstPass As Double, i As Double, a As Double

    If Len(Pass) > 0 Then
     FirstPass = Asc(Mid$(Pass, Len(Pass), Len(Pass)))
    ElseIf Len(Pass) < 0 Then
     FirstPass = Asc(Mid$(Pass, Len(Text), 1))
    End If

    For i = 1 To Len(Text)
     a = Asc(Mid$(Text, i, 1)) - 2 * FirstPass
     If a < 0 Then a = 255 - (-a + ((a \ 255) * 255))
     FirstPass = Asc(Mid$(Text, i, 1))
     ChrEncrypt = ChrEncrypt & Chr$(a)
    Next i

DecError:
    If Err.Number <> 0 Then
     ChrEncrypt = Text
     Exit Function
    End If

End Function
