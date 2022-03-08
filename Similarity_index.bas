
Function mostSimilar(val As String, vals() As String)

Dim size As Integer

size = Len(vals)

For valLooper = 1 To size

Next

End Function

Function codeVal(str As String) As Double
' Assess the value of a string based on the place-value of characters from left to right
Dim looper As Integer
Dim chars() As Byte
Dim size As Integer

size = Len(str)

Dim temp As Double
temp = 0


chars = StrConv(str, vbfrmunicode)

' looping the chars array
For looper = 1 To size

  temp = temp + chars(looper) * 10 ^ (looper)

Next

codeVal = temp

End Function

