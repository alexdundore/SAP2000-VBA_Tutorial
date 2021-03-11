Attribute VB_Name = "Example_6"
Sub API_Example_6()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variable
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '

' turn off automatically generated combinations
ret = SapModel.DesignSteel.SetComboAutoGenerate(False)

' turn user-defined combination for strength design
ret = SapModel.DesignSteel.SetComboStrength("1.2D + W", True)

' set design code and execute the steel design module
ret = SapModel.DesignSteel.SetCode("AISC 360-10")
ret = SapModel.DesignSteel.StartDesign()

' get a summary of the results
Dim Name As String
Dim NumberItems As Long
Dim FrameName() As String
Dim Ratio() As Double
Dim RatioType() As Long
Dim Location() As Double
Dim ComboName() As String
Dim ErrorSummary() As String
Dim WarningSummary() As String
Dim ItemType As eItemType

Name = "Frames" ' name of the group of interest
ItemType = eItemType_Group ' indicating results for a group, not an object
ret = SapModel.DesignSteel.GetSummaryResults(Name, NumberItems, FrameName, _
Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary, ItemType)

' creating an output table and writing the results
Dim WS As Worksheet ' dimension the worksheet
Set WS = Worksheets("Sheet1")
Dim OutputTable3 As Range ' setting area in spreadsheet to print results to
Set OutputTable3 = WS.Range(WS.Cells(1, 9), WS.Cells(NumberItems + 1, 10))
OutputTable3.Clear ' clears current values
OutputTable3(1, 1).Value = "Frame Name" ' setting column headers
OutputTable3(1, 2).Value = "DCR"

For i = 1 To NumberItems
    OutputTable3(i + 1, 1).Value = FrameName(i - 1)
    OutputTable3(i + 1, 2).Value = Ratio(i - 1)
    OutputTable3.Rows(i + 1).NumberFormat = "0.000" ' truncating decimals
Next i

' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub


