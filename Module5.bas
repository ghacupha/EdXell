Attribute VB_Name = "Module5"
Function sum_range(r As Range)


Dim value As Double

For Each cell In r

value = value + cell.value

Next cell

sum_range = value

End Function
