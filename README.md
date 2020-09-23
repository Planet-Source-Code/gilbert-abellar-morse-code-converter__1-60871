<div align="center">

## Morse Code Converter


</div>

### Description

This program converts Alpha-Numeric characters into Morse Code, it also plays Morse sound. Very usefull for decrypting or encrypting strings, especially passwords.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-06-01 23:48:08
**By**             |[Gilbert Abellar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gilbert-abellar.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Morse\_Code189599622005\.zip](https://github.com/Planet-Source-Code/gilbert-abellar-morse-code-converter__1-60871/archive/master.zip)

### API Declarations

```
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
```





