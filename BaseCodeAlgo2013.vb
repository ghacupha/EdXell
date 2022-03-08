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
Sub unlocksh()

    Dim i As Integer, j As Integer, k As Integer

    Dim l As Integer, m As Integer, n As Integer

    Dim i1 As Integer, i2 As Integer, i3 As Integer

    Dim i4 As Integer, i5 As Integer, i6 As Integer

    Dim pwd As String

    On Error Resume Next

    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66

        For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66

            For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66

                For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126

                    pwd = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

                    ActiveWorkbook.Unprotect pwd

                    If ActiveWorkbook.ProtectStructure = False Then

                        MsgBox "One usable password is " & pwd

                        ActiveWorkbook.Sheets(1).Select

                        Range("a1").FormulaR1C1 = pwd

                        Exit Sub

                    End If

                Next: Next: Next

            Next: Next: Next

        Next: Next: Next

    Next: Next: Next

End Sub