VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vsoWindow As Window
Attribute vsoWindow.VB_VarHelpID = -1
Dim myMouseListener As MouseListener
Public mouse_x As Double
Public mouse_y As Double
Public is_add_point As Boolean
Dim ge_row_num As Integer
 
Public Sub Reroute()
Dim vsoShapeUID As String


For Each vsoShape In ActiveWindow.Selection
    vsoShapeUID = vsoShape.UniqueID(1)
    Set vsoShape = ActivePage.Shapes(vsoShapeUID)
    vsoShape.CellsSRC(visSectionFirstComponent, 0, 0).FormulaForceU = "0"
    vsoShape.CellsSRC(visSectionFirstComponent, 0, 1).FormulaForceU = "0"
    vsoShape.CellsSRC(visSectionObject, visRowFill, visFillPattern).FormulaU = "0"
    Exit For
Next
 
 Set vsoWindow = ActiveWindow
 Call AddPoint(vsoShapeUID)
 
exitsub:
End Sub

Public Sub vsoWindow_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
 
 Debug.Print "Button is: "; Button; x; y
 mouse_x = x
 mouse_y = y
 is_add_point = True
End Sub

Sub AddPoint(vsoShapeUID As String)
Dim PauseTime, Start, Finish, TotalTime
Dim PreShapeName As String
ge_row_num = 1

PauseTime = 600    ' Set duration.
Start = Timer   ' Set start time.

Do While Timer < Start + PauseTime
DoEvents
If stop_sub Then GoTo exitsub
If is_add_point Then
    Call AddPoint2(ge_row_num, vsoShapeUID)
    ge_row_num = ge_row_num + 1
End If
Loop
exitsub:
stop_sub = False
End Sub

Sub AddPoint2(ge_row_num As Integer, vsoShapeUID As String)
Debug.Print ge_row_num
If ge_row_num >= vsoShape.RowCount(visSectionFirstComponent) Then
    Debug.Print ge_row_num, vsoShape.RowCount(visSectionFirstComponent)
    ge_row_num = vsoShape.RowCount(visSectionFirstComponent) - 1
End If
Dim UndoScopeID1 As Long
UndoScopeID1 = Application.BeginUndoScope("Manual Edit")
Set vsoShape = ActivePage.Shapes(vsoShapeUID)
vsoShape.CellsSRC(visSectionObject, visRowShapeLayout, visSLOConFixedCode).FormulaForceU = "2"
vsoShape.CellsSRC(visSectionFirstComponent, ge_row_num, 0).FormulaForceU = mouse_x - vsoShape.Cells("BeginX").Result(visNone)
vsoShape.CellsSRC(visSectionFirstComponent, ge_row_num, 1).FormulaForceU = mouse_y - vsoShape.Cells("BeginY").Result(visNone)
vsoShape.CellsSRC(visSectionObject, visRowFill, visFillPattern).FormulaU = "0"
is_add_point = False
End Sub

Public Sub ReturnLine()
Dim vsoShapeUID As String
If ActiveWindow.Selection.Count <> 1 Then
    MsgBox "Plase select only one shape."
    GoTo exitsub
End If

For Each vsoShape In ActiveWindow.Selection
    vsoShapeUID = vsoShape.UniqueID(1)
    Set vsoShape = ActivePage.Shapes(vsoShapeUID)
    Exit For
Next
vsoShape.CellsSRC(visSectionObject, visRowShapeLayout, visSLOConFixedCode).FormulaForceU = "0"
vsoShapeGlueToUID = vsoShape.Connects(1).ToSheet.UniqueID(1)
vsoShapeGlueToPort = CutPort(vsoShape.Cells("BeginX").Formula)
vsoShape.Cells("BeginX").Formula = "1"
vsoShape.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 0
vsoShape.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 1

Set vsoCellGlueFromEnd = vsoShape.Cells("BeginX")
Set vsoCellGlueToObject = ActivePage.Shapes(vsoShapeGlueToUID).Cells(vsoShapeGlueToPort)
    vsoCellGlueFromEnd.GlueTo vsoCellGlueToObject
ge_row_num = 1
exitsub:
End Sub

Function CutPort(str As String) As String

string1 = Left(str, Len(str) - InStr(1, str, ".X,") + 2)
CutPort = Right(string1, Len(string1) - InStr(1, string1, "!"))
End Function
