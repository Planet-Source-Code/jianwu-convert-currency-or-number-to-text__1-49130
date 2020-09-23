<div align="center">

## Convert Currency or Number to Text


</div>

### Description

Tow functions provide to convert the number or currency into English Text.
 
### More Info
 
amount The amount to be converted

The converted English string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[jianwu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jianwu.md)
**Level**          |Advanced
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jianwu-convert-currency-or-number-to-text__1-49130/archive/master.zip)





### Source Code

```
'=================================================
'
' Functions used to convert number or currency into English String
'
' Author: Chen Jianwu (jianwu_chen@yahoo.com)
' Create Date: 2003-10-10
'
'=================================================
Option Explicit
Dim suffix
Dim digitName
Dim namety
'=================================================
'
' Convert number to English String
'
' @param num The number to be converted
' @param units Optional parameter, the unit string which will be append to the result string
' @return The converted English string
'
'=================================================
Public Function Number_2_String(num As Long, Optional units As String = "units") As String
  Dim i As Integer
  suffix = Array("", "thousand", "million", "milliard", "Tera", "Peta", "Exa")
  digitName = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
  namety = Array("twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
  suffix(0) = units
  i = 0
  Do Until num = 0
    Number_2_String = Small_Number_2_String(num Mod 1000, i) + " " + suffix(i) + " " + Number_2_String
    i = i + 1
    num = num \ 1000
  Loop
End Function
'=================================================
'
' Convert currency to English String
'
' @param amount The amount to be converted
' @return The converted English string
'
'=================================================
Public Function Currency_2_String(amount As Double) As String
  Dim dollars As Long
  Dim cents As Long
  dollars = Int(amount) ' Note Int mean floor and Fix means ceiling on which i spend lot of time .
  cents = (amount - dollars) * 100
  Currency_2_String = Number_2_String(dollars, "dollars") + Number_2_String(cents, "cents")
End Function
Private Function Small_Number_2_String(num As Long, k As Integer) As String
  If num = 0 Then
    Small_Number_2_String = ""
    Exit Function
  End If
  Dim needSpace As Boolean
  needSpace = False
  If num > 99 Then
    Small_Number_2_String = Small_Number_2_String + digitName(num \ 100) + " hundred"
    needSpace = True
    num = num Mod 100
    If (Small_Number_2_String <> "" And num > 0 And k = 0) Then
      Small_Number_2_String = Small_Number_2_String + " and"
    End If
  End If
  If num > 19 Then
    Small_Number_2_String = Small_Number_2_String + IIf(needSpace, " ", "") + namety(num \ 10 - 2)
    needSpace = True
    num = num Mod 10
    If num > 0 Then
      Small_Number_2_String = Small_Number_2_String + "-"
      needSpace = False
    End If
  End If
  If num > 0 Then
    Small_Number_2_String = Small_Number_2_String + IIf(needSpace, " ", "") + digitName(num)
  End If
End Function
'=================================================
'
' Only for Testing
'
'=================================================
Public Sub test_number_2_string()
  MsgBox "10023034 = " + (Number_2_String(10023034))
  MsgBox "231002314 = " + (Number_2_String(231002314))
  MsgBox "90219.11 = " + (Currency_2_String(90219.11))
  MsgBox "384721911.48 = " + (Currency_2_String(384721911.48))
End Sub
```

