Attribute VB_Name = "Example_1"
Sub API_Example_1()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variableS
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '


' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub
