'(The MIT License)
'
'Copyright (c) 2017 Edwin Njeru <edwin.njeru@abcthebank.com>
'
'Permission is hereby granted, free of charge, to any person
'obtaining a copy of this software and associated documentation
' files (the 'Software'), to deal in the Software without restriction,
'including without limitation the rights to use, copy, modify,
'merge, publish, distribute, sublicense, and/or sell copies of the Software,
'and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be
'included in all copies or substantial portions of the Software.
'
'The SOFTWARE Is PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
'EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
'OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
'IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
'DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
'TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
'SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


' InPlaceTrimArray algorithm
' Trim all cells in selection...
' By treating them as arrays the algorithm is much faster hence efficient
' for huge data sets
' as opposed to looping through each individual cell in selection.
' This is especily so if the cells selected are arranged into 2
' dimensions.

Sub InPlaceMultiplyByConstant()

Dim inputData As Double

inputData = InputBoxTest()

Dim inputDataString As String

inputDataString = Format(inputData, "")

MsgBox ("You are about to multiply everything by :" + inputDataString)

Dim rng As Excel.Range

  Set rng = Selection
  
      rng.Formula = Application.Evaluate("=" & rng.Address & "*" & inputData)

  Set rng = Nothing
End Sub

Function InputBoxTest() As Double

    Dim inputData As Double
    '
    ' Get the data
    '
    inputData = InputBox("Enter the amount with which you wish to multiply the range :", "Input Box Text")
    
    InputBoxTest = inputData
End Function