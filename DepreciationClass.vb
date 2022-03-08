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

Imports Microsoft.VisualBasic

Public Class Class1

    'useful life function for the calculation of useful life
    'for any asset given the category of the asset
    Function UsefulLife(ByVal category As Text)

        Dim ComputerSoftware, Computers, ElectronicEquipment,
            FurnitureFitings, MotorVehicles, OfficeRenovation,
            WestlandsBuildings As Double
        ' depreciation rates
        ComputerSoftware = 0.1
        Computers = 0.3
        ElectronicEquipment = 0.3
        FurnitureFitings = 0.125
        MotorVehicles = 0.2
        OfficeRenovation = 0.125
        WestlandsBuildings = 0.02

        Select Case category

            Case Is = "COMPUTER SOFTWARE"
                UsefulLife = 1 / ComputerSoftware
            Case Is = "COMPUTERS"
                UsefulLife = 1 / Computers
            Case Is = "ELECTRONIC EQUIPMENT"
                UsefulLife = 1 / ElectronicEquipment
            Case Is = "ELECTRONIC EQUIPMENT"
                UsefulLife = 1 / ElectronicEquipment
            Case Is = "FURNITURE & FITTINGS"
                UsefulLife = 1 / FurnitureFitings
            Case Is = "MOTOR VEHICLES"
                UsefulLife = 1 / MotorVehicles
            Case Is = "OFFICE RENOVATION"
                UsefulLife = 1 / OfficeRenovation
            Case Is = "WESTLANDS BUILDING OFFICES"
                UsefulLife = 1 / WestlandsBuildings


        End Select



    End Function

End Class
