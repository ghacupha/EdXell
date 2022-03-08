Function FilterUniqueSort(rng As Range)

Dim ucoll As New Collection, Value As Variant, temp() As Variant
Dim iRows As Single, i As Single
ReDim temp(0)

On Error Resume Next
For Each Value In rng
    If Len(Value) > 0 Then ucoll.Add Value, CStr(Value)
Next Value
On Error GoTo 0

For Each Value In ucoll
    temp(UBound(temp)) = Value
    ReDim Preserve temp(UBound(temp) + 1)
Next Value

ReDim Preserve temp(UBound(temp) - 1)

iRows = Range(Application.Caller.Address).Rows.Count

SelectionSort temp

For i = UBound(temp) To iRows
  ReDim Preserve temp(UBound(temp) + 1)
  temp(UBound(temp)) = ""
Next i

FilterUniqueSort = Application.Transpose(temp)

End Function





Function SelectionSort(TempArray As Variant)
          Dim MaxVal As Variant
          Dim MaxIndex As Integer
          Dim i, j As Integer

          ' Step through the elements in the array starting with the
          ' last element in the array.
          For i = UBound(TempArray) To 0 Step -1

              ' Set MaxVal to the element in the array and save the
              ' index of this element as MaxIndex.
              MaxVal = TempArray(i)
              MaxIndex = i

              ' Loop through the remaining elements to see if any is
              ' larger than MaxVal. If it is then set this element
              ' to be the new MaxVal.
              For j = 0 To i
                  If TempArray(j) > MaxVal Then
                      MaxVal = TempArray(j)
                      MaxIndex = j
                  End If
              Next j

              ' If the index of the largest element is not i, then
              ' exchange this element with element i.
              If MaxIndex < i Then
                  TempArray(MaxIndex) = TempArray(i)
                  TempArray(i) = MaxVal
              End If
          Next i

      End Function