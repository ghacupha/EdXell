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




' Method for classification of deposits as per the following sectors:
'
' a) Real Estate
' b) Personal/Household
' c) Trade
' d) Manufacturing
' e) Building and Construction
' f) Financial Services
' g) Tourism,Restaurant and Hotel
' h) Energy and water
' g) Mining and Quarrying
' h) Agriculture
' i) Transport and Communication
Function SectoralClassDeposits(Args As String) As String

' convert input into lower case
searchString = LCase(Args)

' Variable returned by the function
Dim result As String

' 1)Criterion for Real Estate
If InStr(searchString, "real") > 0 Then

result = "Real Estate"


' 2)Criterion for Personal/Household
ElseIf InStr(searchString, "personal") > 0 Then

result = "Personal/Household"

' 3)Criterion for Trade
ElseIf InStr(searchString, "trade") > 0 or InStr(searchString, "dealer") > 0 or InStr(searchString, "forwarding") > 0 or InStr(searchString, "distributor") > 0 or InStr(searchString, "consume") > 0 Then

result = "Trade"

' 4)Criterion for Manufacturing
ElseIf InStr(searchString, "manufacturing") > 0 or InStr(searchString, "epz") > 0 or InStr(searchString, "miller") > 0 Then

result = "Manufacturing"

' 5)Criterion for Building and Construction
ElseIf InStr(searchString, "building") > 0 Or InStr(searchString, "construction") > 0 Or InStr(searchString, "contractor") > 0 Then

result = "Building and Construction"

' 6)Criterion for Financial Services
ElseIf InStr(searchString, "finance") > 0 or InStr(searchString, "broker") > 0 or InStr(searchString, "sacco") > 0 or InStr(searchString, "forex") > 0 or InStr(searchString, "fund") > 0 Then

result = "Financial Services"

' 7)Criterion for Tourism,Restaurant and Hotel
ElseIf InStr(searchString, "tourism") > 0 Or InStr(searchString, "restaurant") > 0 Or InStr(searchString, "hotel") > 0 Then

result = "Tourism,Restaurant and Hotel"

' 8)Criterion for Energy and water
ElseIf InStr(searchString, "energy") > 0 Or InStr(searchString, "water") > 0 Then

result = "Energy and water"

' 9)Criterion for Mining and Quarrying
ElseIf InStr(searchString, "mining") > 0 Or InStr(searchString, "quarrying") > 0 Then

result = "Mining and Quarrying"

' 10)Criterion for Agriculture
ElseIf InStr(searchString, "agriculture") > 0 Then

result = "Agriculture"

' 11)Criterion for Transport and Communication
ElseIf InStr(searchString, "transport") > 0 or InStr(searchString, "telecom") > 0 or InStr(searchString, "media") > 0 or InStr(searchString, "shipping") > 0 Then

result = "Transport and Communication"

' 12)If all the above are not in the search string we classify as
'    Trade
Else

result = "Trade"

' TODO: include sectoral analysis in the production reporting template






End If

' Return result to the main function
SectoralClassDeposits = result

End Function
