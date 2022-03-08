Sub FlipRows()
 Dim vTop As Variant
 Dim vEnd As Variant
 Dim iStart As Integer
 Dim iEnd As Integer
 Application.ScreenUpdating = False
 iStart = 1
 iEnd = Selection.Rows.Count
 Do While iStart < iEnd
 vTop = Selection.Rows(iStart)
 vEnd = Selection.Rows(iEnd)
 Selection.Rows(iEnd) = vTop
 Selection.Rows(iStart) = vEnd
 iStart = iStart + 1
 iEnd = iEnd - 1
 Loop
 Application.ScreenUpdating = True
End Sub