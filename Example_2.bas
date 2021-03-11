Attribute VB_Name = "Example_2"
Global n_bays As Integer
Global n_stories As Integer
Global width As Double
Global height As Double
Global Col_Size As String
Global Beam_Size As String

Sub API_Example_2()

' dimension the sap object and sap model
Dim SapObject As cOAPI
Dim SapModel As cSapModel

' set the sap object as the current instance
Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")

' attach current sap model to the variableS
Set SapModel = SapObject.SapModel

' --- START YOUR CODE HERE --- '
' unlock the model if it is already locked!
ret = SapModel.SetModelIsLocked(False)

' set units and select & delete existing framing
ret = SapModel.SetPresentUnits(eUnits_lb_ft_F)
ret = SapModel.SelectObj.All(False) ' selects all objects in model
ret = SapModel.FrameObj.Delete("ignore", eItemType.eItemType_SelectedObjects)
ret = SapModel.PointObj.DeleteSpecialPoint("ignore", eItemType_SelectedObjects)

' create parameters to characterize a moment frame
' these can be easily changed to affect the whole geometry
n_bays = 2
n_stories = 2
width = 20 ' 15 feet
height = 14 ' 12 feet
Col_Size = "W14X53"
Beam_Size = "W12X40"

' add points to the model, setting restraints to ground level nodes
Dim x As Double
Dim y As Double
Dim z As Double
Dim Name As String
Dim UserName As String
Dim Value(0 To 5) As Boolean
Value(0) = True ' U1 translation fixity
Value(1) = True ' U2 translation fixity
Value(2) = True ' U3 translation fixity
Value(3) = True ' R1 translation fixity
Value(4) = True ' R2 translation fixity
Value(5) = True ' R3 translation fixity

For b = 0 To n_stories
    For c = 0 To n_bays
        ' create nodes at each beam & column junction
        x = 0 + c * width
        y = 0
        z = 0 + b * height
        UserName = "Node_" + CStr(b) + "_" + CStr(c)
        ret = SapModel.PointObj.AddCartesian(x, y, z, Name, UserName)
        ' set full fixity if on base level
        If b = 0 Then
            ret = SapModel.PointObj.SetRestraint(Name, Value)
        End If
    Next c
Next b

' create a group in which all frame elements will be added to
ret = SapModel.GroupDef.SetGroup("Frames")

' add columns to the model by specifying start & end nodes
Dim Point1, Point2 As String
Dim PropName As String
PropName = Col_Size

For b = 0 To n_stories
    For c = 0 To n_bays
        Point1 = "Node_" + CStr(b) + "_" + CStr(c)
        Point2 = "Node_" + CStr(b + 1) + "_" + CStr(c)
        UserName = "Col_" + CStr(b) + "_" + CStr(c)
        ret = SapModel.FrameObj.AddByPoint(Point1, Point2, Name, PropName, UserName)
        ret = SapModel.FrameObj.SetGroupAssign(Name, "Frames")
    Next c
Next b

' add beams to the model by specifying start & end nodes
PropName = Beam_Size
For b = 1 To n_stories
    For c = 0 To n_bays - 1
        Point1 = "Node_" + CStr(b) + "_" + CStr(c)
        Point2 = "Node_" + CStr(b) + "_" + CStr(c + 1)
        UserName = "Beam_" + CStr(b) + "_" + CStr(c)
        ret = SapModel.FrameObj.AddByPoint(Point1, Point2, Name, PropName, UserName)
        ret = SapModel.FrameObj.SetGroupAssign(Name, "Frames")
    Next c
Next b

' --- END OF USER CODE --- '

' releases memory resources
Set SapModel = Nothing
Set SapObject = Nothing

End Sub


















