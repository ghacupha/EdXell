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


' InPlaceRoundArray() algorithm
' Round to 2 decimal places all cells in selection...
' By treating them as arrays the algorithm is much faster hence efficient
' for huge data sets
' as opposed to looping through each individual cell in selection.
' This is especily so if the cells selected are arranged into 2
' dimensions.
private Sub InPlaceRoundArray()

Dim arrData() As Variant
Dim arrReturnData() As Variant
Dim rng As Excel.Range
Dim lRows As Long
Dim lCols As Long
Dim i As Long, j As Long

  lRows = Selection.Rows.Count
  lCols = Selection.Columns.Count

  ReDim arrData(1 To lRows, 1 To lCols)
  ReDim arrReturnData(1 To lRows, 1 To lCols)

  Set rng = Selection
  arrData = rng.Value

  For j = 1 To lCols
    For i = 1 To lRows
      arrReturnData(i, j) = Round(arrData(i, j),2)
      ' Not working
      ' Debug.Print arrData.Address
      ' Debug.Print arrReturnData.Address
    Next i
  Next j

  rng.Value = arrReturnData

  Set rng = Nothing
End Sub