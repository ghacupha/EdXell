
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

Sub ClearCBKReport()

' SHORTCUT: ctrl+shift+d

Dim r As Range
' generic range of cells in all workbooks


'Step 1:  Declare your variables
    Dim ws As Worksheet
    Dim cell As Object
'Step 2: Start looping through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    Set r = ws.Range("A1:U500")
    
    ' Now to loop through cells in range
    For Each cell In r.Cells
    
    ' conditional for clearing contents
    If cell.Interior.Color = RGB(255, 255, 153) Or cell.Interior.Color = RGB(255, 255, 204) Then
    
    If Len(cell) > 0 Then
    
    cell.ClearContents
    
    End If
    
    End If
    
    Next cell
    
'Step 5:  Loop to next worksheet
    Next ws
End Sub
