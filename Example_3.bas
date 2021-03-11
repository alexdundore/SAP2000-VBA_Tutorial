Attribute VB_Name = "Example_3"
Sub API_Example_3()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variable
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '
ret = SapModel.SetPresentUnits(eUnits_lb_ft_F)

' specify load intensities
Dim dead_load As Double
Dim wind_load As Double
dead_load = 50 ' 50 plf
wind_load = 50000 ' 50 kip point load

' declare variables relating to dead loads on frames
Dim Name As String
Dim LoadPat As String
Dim MyType As Double
Dim Dir As Long
Dim Dist1, Dist2 As Double
Dim Val1, Val2 As Double
MyType = 1 ' indicates Force/Length load type
Dir = 10 ' indicates gravity load direction
Dist1 = 0 ' start of distributed load at member start
Dist2 = 1 ' end of distributed load at member end
Val1 = dead_load ' load will be uniformly distributed
Val2 = dead_load

' add distributed dead loads to the beams
LoadPat = "dead"
For b = 1 To n_stories
    For c = 0 To n_bays - 1
        Name = "Beam_" + CStr(b) + "_" + CStr(c)
        ret = SapModel.FrameObj.SetLoadDistributed(Name, LoadPat, MyType, Dir, Dist1, Dist2, Val1, Val2)
    Next c
Next b

' declare variables relating to wind loads on nodes
Dim Value(0 To 5) As Double
Dim Replace As Boolean
Value(0) = wind_load ' Force in DOF 1
Value(1) = 0 ' Force in DOF 2
Value(2) = 0 ' Force in DOF 3
Value(3) = 0 ' Moment in DOF 4
Value(4) = 0 ' Moment in DOF 5
Value(5) = 0 ' Moment in DOF 6
Replace = True ' will replace the loads already applied to node

' add lateral point loads to the left-most nodes
LoadPat = "wind"
For b = 1 To n_stories
    Name = "Node_" + CStr(b) + "_0"
    ret = SapModel.PointObj.SetLoadForce(Name, LoadPat, Value, Replace)
Next b

' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub
