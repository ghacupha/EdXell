' This algorith iterates the columns in a given selection and then, for each columns
' it will fill down the gaps using the data in previous cells
' for purposes for which it was originally created the rows selected, the rows selected
' are limited between row 15 and 1500 (an estimation of usefulness), according to the
' template configuration of the report we are working on. This limitation will enable the
' user to select entire columns in a workbook and run the filldown algorithm without
' limiting the selection manually
Sub fill_Down_With_Previous()

    Dim c As Range
    
    For Each c In Selection.Columns
    
      Dim columnValues  As Range, i As Long

      Set columnValues = c

      ' To limit selection for the report template comment the For loop head and
      ' uncomment the following line
      ' For i = 15 To 1500
      For i = 1 To columnValues.Rows.Count
          If columnValues.Cells(i, 1).Value = "" Then
              columnValues.Cells(i, 1).Value = columnValues.Cells(i - 1, 1).Value
          End If
      Next
    
    Next c

End Sub