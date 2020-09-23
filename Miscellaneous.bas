Attribute VB_Name = "Miscellaneous"
Option Explicit

Global Const conSwNormal As Long = 1

Public Enum EffectConstant
    TString = 0
End Enum

Public I As Long
Public J As Long
Public X As Long
Public Y As Long
Public Z As Long

Public ClassMod As MainFunction

Public Can As Boolean
Public Extras As Boolean
Public Cancelling As Boolean

Public Num() As String
Public Char() As String

Public Num1(48 To 57) As String
Public Char1(65 To 90) As String

Public Const NumCont As String = "_ _ _ _ _|._ _ _ _|.._ _ _|..._ _|...._|.....|_....|_ _...|_ _ _..|_ _ _ _."
Public Const CharCont As String = "._|_...|_._.|_..|.|.._.|_ _.|....|..|._ _ _|_._|._..|_ _|_.|_ _ _|._ _.|_ _._|._.|...|_|.._|..._|._ _|_.._|_._ _|_ _.."

Public Declare Sub Sleep Lib "kernel32" (ByVal lpMilliSeconds As Long)
Public Declare Sub Beep Lib "kernel32" (ByVal dwFrequency As Long, ByVal dwLength As Long)

Public Function TrimString(ByVal TString As String) As String
    TrimString = Replace$(TString, Space$(conSwNormal), vbNullString)
End Function

Public Function IsAlpha(ByVal TString As String) As Boolean
    Dim X As Long
    For X = conSwNormal To Len(TString)
        IsAlpha = (Mid$(TString, X, conSwNormal) Like "[A-Z]" Or Mid$(TString, X, conSwNormal) Like "[a-z]")
        If IsAlpha = False Then Exit Function
    Next X
End Function

Public Function IsNumeric(ByVal TString As String) As Boolean
    Dim X As Long
    For X = conSwNormal To Len(TString)
        IsNumeric = (Mid$(TString, X, conSwNormal) Like "[0-9]")
        If IsNumeric = False Then Exit Function
    Next X
End Function

Public Function ConvertToUpperAlpha(ByVal Str As String) As String
    ConvertToUpperAlpha = UCase$(Str)
End Function

Public Function IsValidChar(ByVal Str As String, Effect As EffectConstant) As Boolean
    Dim X As Long, StrTemp As String
    Select Case Effect
        Case TString
            StrTemp = ConvertToUpperAlpha(Str)
            For X = conSwNormal To Len(StrTemp)
                IsValidChar = (Mid$(StrTemp, X, conSwNormal) Like "[A-Z]") Or (Mid$(StrTemp, X, conSwNormal) Like "[0-9]") Or (Mid$(StrTemp, X, conSwNormal) = Space$(conSwNormal))
                If IsValidChar = False Then Exit Function
            Next X
    End Select
End Function

Public Function GetSecretWritting(ByVal Str As String, Effect As EffectConstant) As String
    Set ClassMod = New MainFunction
    Select Case Effect
        Case TString: GetSecretWritting = ClassMod.GetMorseCode(Str)
    End Select
End Function
