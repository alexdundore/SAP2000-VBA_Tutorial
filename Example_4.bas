Attribute VB_Name = "Example_4"
Sub API_Example_4()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variable
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '
ret = SapModel.SetPresentUnits(eUnits_lb_ft_F)

' unlock the model if it is already locked
' analysis only runs when model is unlocked
ret = SapModel.SetModelIsLocked(False)

' run the analysis
ret = SapModel.Analyze.RunAnalysis()

' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub

