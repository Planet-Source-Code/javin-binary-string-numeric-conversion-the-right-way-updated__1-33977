<div align="center">

## Binary String\<\-\-\> Numeric conversion \- THE RIGHT WAY\!\!\!\! \{Updated\!\}


</div>

### Description

First some clarification: This takes a number (say, 65) and the number of bits you want returned (say 8) and returns a STRING of 1's and 0's (such as "01000001"). Also converts the other way, too. Since VB doesn't have any bitshifting capability, this MIGHT be the only way to do it. (If someone has a better method, by all means PLEASE let me know.) <p> I came in here looking for a quick and dirty method of doing this. I don't mean to be rude, people, but if you don't know what you're doing, don't upload the code. I poured through well over 2 dozen horribly ugly "methods" of conversion that all involved nested loops, select case, if/then trees, and all other sorts of nonsense. Finally, I gave up and wrote it myself. Once you understand the maths that go into making a binary number, you'll understand how simple this code really is. It can manage to convert any Long Integer to binary, and back again, quick, easy, and with a minimum of overhead. I'm sure there's probably even APIs to do this, now, but they likely wouldn't be available on legacy systems, and this method will work in any language. Note that this is marked as "Advanced" code. If you can't figure out how to use it, that's on you. In reality, it's pretty straightforward.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Javin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/javin.md)
**Level**          |Advanced
**User Rating**    |4.1 (41 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/javin-binary-string-numeric-conversion-the-right-way-updated__1-33977/archive/master.zip)





### Source Code

```
'*
'*
'*
'***NOTE: This is the highly optimized code. If you haven't been following this uphill battle, you won't get it. But this is (I believe) the fastest method of finding binary values through VB. - This was a collective effort of myself and
'Kaverin (Of #VB on DalNet)
'*
Public Function ConvertToBinary(ByVal Number As Long, ByVal Bits As Byte) As String
 Dim intCount As Byte
 ConvertToBinary = String$(Bits, "0")
 For intCount = 1 To Bits
 If Number And 1 Then Mid$(ConvertToBinary, Bits + 1 - intCount, 1) = "1"
 Number = Number \ 2
 Next intCount
End Function
Public Function ConvertFromBinary(ByVal Binary As String) As Long
 Dim intCount As Integer
 For intCount = Len(Binary) To 1 Step -1
 If Mid$(Binary, intCount, 1) = "1" Then ConvertFromBinary = ConvertFromBinary + (&H2 ^ (Len(Binary) - intCount))
 Next intCount
End Function
```

