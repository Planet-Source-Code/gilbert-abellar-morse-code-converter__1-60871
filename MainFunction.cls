VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PlayMorseCode(ByVal Code As String, TXT As TextBox, L1 As PictureBox, L2 As PictureBox, P1 As PictureBox, P2 As PictureBox)
    Dim X As Long
    For X = conSwNormal To Len(Code)
        If Can = True Then Exit For
        L1.Picture = P2.Picture
        L2.Picture = P2.Picture
        L1.BackColor = P2.BackColor
        L2.BackColor = P2.BackColor
        Select Case Mid$(Code, X, conSwNormal)
            Case ".": Call Beep(2000, 80)
            Case "_": Call Beep(2000, 150)
        End Select
        L1.Picture = P1.Picture
        L2.Picture = P1.Picture
        L1.BackColor = P1.BackColor
        L2.BackColor = P1.BackColor
        Select Case Mid$(Code, X, conSwNormal)
            Case "*": Call Sleep(100)
            Case Else: Call Sleep(50)
        End Select
        With TXT
            .SelStart = X
            .SelLength = conSwNormal
        End With
        DoEvents
    Next X
End Sub

Public Function GetMorseCode(ByVal TString As String) As String
    Dim Wrd As String, Wrd1 As String
    Wrd = Replace$(TString, vbNewLine, vbNullChar)
    Wrd = Replace$(Wrd, vbTab, vbNullChar)
    Wrd = Replace$(Wrd, Space(conSwNormal), vbNullChar)
    Wrd1 = vbNullString
    For X = conSwNormal To Len(Wrd)
        If Mid$(Wrd, X, conSwNormal) = vbNullChar And Mid$(Wrd, X + conSwNormal, conSwNormal) = vbNullChar Then
        Else
            If Not X = Len(Wrd) Then Wrd1 = Wrd1 & Mid$(Wrd, X, conSwNormal)
        End If
        DoEvents
    Next X
    Wrd1 = Wrd1 & Mid$(Wrd, Len(Wrd), conSwNormal)
    If Left$(Wrd1, conSwNormal) = vbNullChar Then Wrd1 = Right$(Wrd1, Len(Wrd1) - conSwNormal)
    Wrd1 = Replace$(Wrd1, vbNullChar, Space$(conSwNormal))
    For X = conSwNormal To Len(Wrd1)
        If IsAlpha(Mid$(Wrd1, X, conSwNormal)) = True And IsNumeric(Mid$(Wrd1, X, conSwNormal)) = False Then
            If Not Mid$(Wrd1, X, conSwNormal) = Space$(conSwNormal) Then GetMorseCode = GetMorseCode & Char1(Val(Trim$(Asc(UCase$(Mid$(Wrd1, X, conSwNormal))))))
        ElseIf IsAlpha(Mid$(Wrd1, X, conSwNormal)) = False And IsNumeric(Mid$(Wrd1, X, conSwNormal)) = True Then
            If Not Mid$(Wrd1, X, conSwNormal) = Space$(conSwNormal) Then GetMorseCode = GetMorseCode & Num1(Val(Trim$(Asc(UCase$(Mid$(Wrd1, X, conSwNormal))))))
        Else
            GetMorseCode = GetMorseCode & "*"
        End If
        DoEvents
    Next X
End Function
