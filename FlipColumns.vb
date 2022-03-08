Sub FlipColumns()
 Dim vTop As Variant
 Dim vEnd As Variant
 Dim iStart As Integer
 Dim iEnd As Integer
 Application.ScreenUpdating = False
 iStart = 1
 iEnd = Selection.Columns.Count
 Do While iStart < iEnd
 vTop = Selection.Columns(iStart)
 vEnd = Selection.Columns(iEnd)
 Selection.Columns(iEnd) = vTop
 Selection.Columns(iStart) = vEnd
 iStart = iStart + 1
 iEnd = iEnd - 1
 Loop
 Application.ScreenUpdating = True
End Sub