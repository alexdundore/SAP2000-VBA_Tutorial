Attribute VB_Name = "Example_5"
Sub API_Example_5()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variable
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '
' setting units to kips and inches
ret = SapModel.SetPresentUnits(eUnits_kip_in_F)

' first, deselect all cases and combos for output
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

' set the combo "1.2D + W" as the selected output combo
ret = SapModel.Results.Setup.SetComboSelectedForOutput("1.2D + W", True)

' get drift (joint displacement) at each right-most roof node
' then writing to spreadsheet table
Dim Name As String
Dim NumberResults As Long
Dim Obj() As String
Dim Elm() As String
Dim LoadCase() As String
Dim StepType() As String
Dim StepNum() As Double
Dim U1() As Double
Dim U2() As Double
Dim U3() As Double
Dim R1() As Double
Dim R2() As Double
Dim R3() As Double

Dim WS As Worksheet ' dimension the worksheet
Set WS = Worksheets("Sheet1")
Dim OutputTable As Range ' setting area in spreadsheet to print results to
Set OutputTable = WS.Range(WS.Cells(1, 1), WS.Cells(n_stories + 1, 2))
OutputTable.Clear ' clears current values
OutputTable(1, 1).Value = "Node Name" ' setting column headers
OutputTable(1, 2).Value = "Lateral Drift [in]"

Dim c As Integer
c = n_bays ' only interested in right-most column line
For b = 1 To n_stories
    Name = "Node_" + CStr(b) + "_" + CStr(c) ' node name for results
    ret = SapModel.Results.JointDispl(Name, 0, NumberResults, Obj, Elm, LoadCase, _
    StepType, StepNum, U1, U2, U3, R1, R2, R3) ' underscore is a line break for clarity
    OutputTable(1 + b, 1).Value = Name ' writing name to output table
    OutputTable(1 + b, 2).Value = U1(0) ' writing lateral node displacement to table
    OutputTable.Rows(1 + b).NumberFormat = "0.000" ' truncating decimals
Next b


' get base reactions of the ground level nodes
Dim F1() As Double
Dim F2() As Double
Dim F3() As Double
Dim M1() As Double
Dim M2() As Double
Dim M3() As Double

Dim OutputTable2 As Range ' setting area in spreadsheet to print results
Set OutputTable2 = Worksheets("Sheet1").Range(WS.Cells(1, 4), WS.Cells(n_bays + 1, 7))
OutputTable2.Clear ' clears current values
OutputTable2(1, 1).Value = "Node Name" ' setting column headers
OutputTable2(1, 2).Value = "Fx [kips]"
OutputTable2(1, 3).Value = "Fz [kips]"
OutputTable2(1, 4).Value = "M [kip-in]"

For c = 0 To n_bays
    Name = "Node_0_" + CStr(c) ' node name for results
    ret = SapModel.Results.JointReact(Name, 0, NumberResults, Obj, Elm, LoadCase, _
    StepType, StepNum, F1, F2, F3, M1, M2, M3) ' underscore is a line break for clarity
    OutputTable2(2 + c, 1).Value = Name ' writing name to output table
    OutputTable2(2 + c, 2).Value = F1(0) ' writing lateral force to table
    OutputTable2(2 + c, 3).Value = F3(0) ' writing vertical force to table
    OutputTable2(2 + c, 4).Value = M2(0) ' writing moment to table
    OutputTable2.Rows(2 + c).NumberFormat = "0.000" ' truncating decimals
Next c

' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub




















