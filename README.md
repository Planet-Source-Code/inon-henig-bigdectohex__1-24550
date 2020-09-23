<div align="center">

## BigDecToHex


</div>

### Description

Converting decimal value to hexadecimal WITHOUT! using the VB functions or mod (which are limited to long data type value), therefore the function can handle BIG numbers, untill 15,000,000,000,000,000 (1.5E+16).
 
### More Info
 
decimal value

hexadecimal value

The functions might not be accurate beyond 15,000,000,000,000,000 (1.5E+16).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Inon Henig](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/inon-henig.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/inon-henig-bigdectohex__1-24550/archive/master.zip)





### Source Code

```
Public Function BigDecToHex(ByVal DecNum) As String
  ' This function is 100% accurate untill 15,000,000,000,000,000 (1.5E+16)
  Dim NextHexDigit As Double
  Dim HexNum As String
  HexNum = ""
  While DecNum <> 0
    NextHexDigit = DecNum - (Int(DecNum / 16) * 16)
    If NextHexDigit < 10 Then
      HexNum = Chr(Asc(NextHexDigit)) & HexNum
    Else
      HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
    End If
    DecNum = Int(DecNum / 16)
  Wend
  If HexNum = "" Then HexNum = "0"
  BigDecToHex = HexNum
End Function
```

