VERSION 5.00
Begin VB.Form frmMorse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conveting Morse Code"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmMorse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Light2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   540
      Left            =   4160
      Picture         =   "frmMorse.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   540
      Left            =   4080
      Picture         =   "frmMorse.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4800
      Picture         =   "frmMorse.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Light1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   540
      Left            =   120
      Picture         =   "frmMorse.frx":1108
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Width           =   4695
      Begin VB.CommandButton cmdPlay 
         Caption         =   "&Play"
         BeginProperty Font 
            Name            =   "Orbus Multiserif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMorse 
         Caption         =   "Convert To &Morse Code"
         BeginProperty Font 
            Name            =   "Orbus Multiserif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   4695
      Begin VB.TextBox txtString1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Schindler"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox txtMorse1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Schindler"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enter &String To Convert"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1020
      TabIndex        =   11
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmMorse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PromptOnce As Boolean

Private Sub cmdMorse_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    txtString1.Enabled = False
    txtMorse1.Enabled = False
    cmdMorse.Enabled = False
    cmdPlay.Enabled = False
    txtMorse1 = GetSecretWritting(txtString1, TString)
    txtString1.Enabled = True
    txtMorse1.Enabled = True
    cmdMorse.Enabled = True
    cmdPlay.Enabled = True
    Screen.MousePointer = 0
    With txtString1
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If Shift = 2 Then
        If KeyCode = vbKeyA Then
            On Error Resume Next
            With txtString1
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        ElseIf KeyCode = vbKeyC Then
            Clipboard.Clear
            Clipboard.SetText txtString1
        ElseIf KeyCode = vbKeyX Then
            Clipboard.Clear
            Clipboard.SetText txtString1
            txtString1 = vbNullString
        ElseIf KeyCode = vbKeyV Then
            If txtString1.SelStart = 0 And txtString1.SelLength = Len(txtString1) Then
                txtString1 = Clipboard.GetText
                On Error Resume Next
                With txtString1
                    .SetFocus
                    .SelStart = 0
                    .SelStart = Len(.Text)
                End With
            Else
                txtString1 = txtString1 & Clipboard.GetText
                txtString1.SelStart = Len(txtString1)
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call Sleep(1000)
    If App.PrevInstance = True Then
        Call Beep(3000, 100)
        Call Sleep(100)
        Call Beep(2000, 100)
        MsgBox Caption & " is already running!", vbOKOnly + vbCritical + vbApplicationModal, "Application Start"
        End
    End If
    PromptOnce = False
    Char = Split(CharCont, "|")
    Num = Split(NumCont, "|")
    Y = 0
    For X = 65 To 90
        Char1(X) = Char(Y)
        Y = Y + conSwNormal
    Next X
    Y = 0
    For X = 48 To 57
        Num1(X) = Num(Y)
        Y = Y + conSwNormal
    Next X
    Extras = False
    Cancelling = False
    If Len(txtString1) <= 0 Then txtMorse1 = vbNullString
    cmdMorse.Enabled = Not (txtString1 = vbNullString)
    cmdPlay.Enabled = Not (txtString1 = vbNullString)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancelling = True Then
        Call Beep(3000, 100)
        Call Sleep(100)
        Call Beep(2000, 100)
        If MsgBox("Program is still processing your request. Are you sure you want to cancel now?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Confirm Action") = vbYes Then
            Can = True
            Screen.MousePointer = 0
        End If
    Else
        Call Beep(3000, 100)
        Call Sleep(100)
        Call Beep(2000, 100)
        If MsgBox("Are you sure you want to quit now?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Confirm Action") = vbYes Then
            Call Sleep(1000)
            End
        End If
    End If
    Cancel = conSwNormal
End Sub

Private Sub cmdPlay_Click()
    On Error Resume Next
    With txtMorse1
        .SetFocus
        .SelStart = 0
        .SelLength = conSwNormal
    End With
    If cmdPlay.Caption = "&Play" Then
        Screen.MousePointer = 11
        cmdMorse.Enabled = False
        txtString1.Enabled = False
        cmdPlay.Caption = "&Stop"
        Can = False
        Cancelling = True
        Call ClassMod.PlayMorseCode(txtMorse1, txtMorse1, Light1, Light2, Picture1, Picture2)
        cmdPlay.Caption = "&Play"
        Can = True
        Cancelling = False
        cmdMorse.Enabled = True
        txtString1.Enabled = True
        Screen.MousePointer = 0
    Else
        cmdPlay.Caption = "&Play"
        Can = True
        Cancelling = False
        cmdMorse.Enabled = True
        txtString1.Enabled = True
        Screen.MousePointer = 0
    End If
    With txtString1
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtString1_Change()
    If InStr(txtString1, vbNewLine) > 0 Then Exit Sub
    cmdMorse.Enabled = Not (txtString1 = vbNullString)
    Screen.MousePointer = 11
    If Not txtString1 = vbNullString Then
        If IsValidChar(txtString1, TString) = False Then
            If PromptOnce = False Then
                MsgBox "It seems to me that you are trying to paste a string having an invalid character, this will not be permitted!", vbOKOnly + vbExclamation + vbSystemModal, Caption
                PromptOnce = True
            End If
            txtString1 = vbNullString
            PromptOnce = False
        End If
    End If
    If Len(txtString1) <= 0 Then txtMorse1 = vbNullString
    If Not Len(txtMorse1) <= 0 Then cmdPlay.Enabled = True Else cmdPlay.Enabled = False
    Screen.MousePointer = 0
End Sub

Private Sub txtString1_GotFocus()
    On Local Error Resume Next
    With txtString1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtString1_KeyPress(KeyAscii As Integer)
    Dim KEY As String
    If KeyAscii = vbKeyEscape Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeySpace Then Exit Sub
    KEY = Chr$(KeyAscii)
    Select Case KEY
        Case "A" To "Z"
        Case "a" To "z": KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
        Case "0" To "9"
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtMorse1_Change()
    If Not Screen.MousePointer = 11 Then cmdPlay.Enabled = Not (txtMorse1 = vbNullString) Else cmdPlay.Enabled = False
End Sub

Private Sub txtMorse1_GotFocus()
    On Error Resume Next
    If Not Cancelling = True Then
        With txtString1
            .SetFocus
        End With
    End If
End Sub

Private Sub txtString2_KeyPress(KeyAscii As Integer)
    Dim KEY As String
    If KeyAscii = vbKeyEscape Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeySpace Then Exit Sub
    KEY = Chr$(KeyAscii)
    Select Case KEY
        Case "*"
        Case "."
        Case "_"
        Case Else: KeyAscii = 0
    End Select
End Sub
