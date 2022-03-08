' Loan Account type definition
Public Type LoanAccount
     CustId As String
     AccountId As String
     AccountName As String
     GroupId As String
     PropertySecurity As String
     VehicleSecurity As String
     FixedDepositSecurity As String
     DebentureSecurity As String
     SharesSecurity As String
End Type




' This algorithm loops Table1 and creates an array of items which are considered to be
' group loans.
' These are then evaluated for individual sums of collateral and then divided by the number
' of companies in the group.
' This is from the assumption that the security provided for the group is iterated evenly in
' every record that has to do with the group
Sub loan_securityEvaluation()

Dim rng As Range
Set rng = ActiveSheet.ListObjects("Table1").Range

Dim rowCount As Long
rowCount = rng.Rows.Count

' An Array of records or selected row objects
Dim loan_list() As LoanAccount
ReDim loan_list(1 To rowCount)

For i = 1 To rowCount

If Cells(i, 1) = "" Then GoTo lastline

     With loan_list(i)
   
      .CustId = Cells(i, 7).Value
      .AccountId = Cells(i, 8).Value
      .AccountName = Cells(i, 9).Value
      .GroupId = Cells(i, 29).Value
      .PropertySecurity = Cells(i, 23).Value
      .VehicleSecurity = Cells(i, 24).Value
      .FixedDepositSecurity = Cells(i, 25).Value
      .DebentureSecurity = Cells(i, 26).Value
      .SharesSecurity = Cells(i, 27).Value
    
     End With
     
     'MsgBox loan_list(i).AccountName & vbNewLine & loan_list(i).AccountId & vbNewLine & loan_list(i).GroupId, vbInformation
lastline:
Next

Dim length As Long
length = arraySize(loan_list)

' show content in the list
MsgBox "first item " + loan_list(7).AccountName & vbNewLine & " loan list size: " & length, vbInformation

' An Array of loans that are part of a group
Dim grouped_loans() As LoanAccount

' Get number of loans in groups
Dim no_of_grouped_loans As Long
no_of_grouped_loans = numberOfGroupLoans(loan_list)

' initialized grouped loans array
ReDim grouped_loans(1 To no_of_grouped_loans)
Dim loan As LoanAccount

MsgBox "Size of grouped loans: " & arraySize(grouped_loans), vbInformation

'todo populate the group loans array
Call populateGroupLoansArr(grouped_loans, loan_list)

' Small Tests to see if group loans array is populated
MsgBox "first item in group loans: " + grouped_loans(1).AccountName & vbNewLine & " group loans array size: " & arraySize(grouped_loans), vbInformation

End Sub



' populate the gouped loans array
Private Sub populateGroupLoansArr(ByRef group_loans_arr() As LoanAccount, all_loans_list() As LoanAccount)
  For x = 1 To arraySize(all_loans_list)
  
   Dim counter As Long
   
   counter = 1
   
   If toNumber(all_loans_list(x).GroupId) <> 0 Then
       
     group_loans_arr(counter) = all_loans_list(x)
    
     counter = counter + 1
    
   End If
   
  Next
  
End Sub



' get number of loans in groups
Private Function numberOfGroupLoans(arr() As LoanAccount) As Long
  Dim counter As Long
    For x = 1 To arraySize(arr)
    If toNumber(arr(x).GroupId) = 0 Then GoTo nextitem
  counter = counter + 1
nextitem:
  Next
numberOfGroupLoans = counter
End Function




' convert string to number
Private Function toNumber(str As String) As Double

toNumber = Val(str)

End Function



' get size of loanAccount array
Private Function arraySize(arr() As LoanAccount) As Integer

arraySize = UBound(arr) - LBound(arr) + 1

End Function
