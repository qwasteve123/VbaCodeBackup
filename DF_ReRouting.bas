Attribute VB_Name = "DF_ReRouting"
Public Sub LineToScale()
Dim vsoShapeUID As String
Dim layout_scale As Double

For Each vsoShape In ActivePage.Shapes
    If InStr(vsoShape, "Layout Scale") Then
        layout_scale = vsoShape.Cells("Prop.scale").Result("mm^-1")
        Exit For
    End If
Next

If layout_scale = 0 Then
    MsgBox "Scale line is not found or scale has not been defined."
    End
End If

For Each vsoShape In ActiveWindow.Selection
    If vsoShape.CellExists("Prop.component_type", 0) = 0 Then GoTo skiploop
    If vsoShape.Cells("Prop.component_type").ResultStr(visNone) = "Connector" Then
        Debug.Print ToLength(vsoShape, layout_scale)
        'add 1m as buffer
        vsoShape.Cells("Prop.feeder_length") = Format(ToLength(vsoShape, layout_scale) / 1000 + 1, "####")
    End If
skiploop:
Next
 
exitsub:
End Sub

Function ToLength(sh As Shape, layout_scale As Double) As Double
Debug.Print sh.RowCount(visSectionFirstComponent)
For i = 1 To sh.RowCount(visSectionFirstComponent) - 2
ToLength = ToLength + TwoPtDistance(i, sh)
Next
ToLength = ToLength * layout_scale
End Function

Function TwoPtDistance(i As Integer, sh As Shape) As Double
Dim x_1, x_2 As Double
Dim y_1, y_2 As Double

x_1 = sh.CellsSRC(visSectionFirstComponent, i, 0).Result("mm")
y_1 = sh.CellsSRC(visSectionFirstComponent, i, 1).Result("mm")

x_2 = sh.CellsSRC(visSectionFirstComponent, i + 1, 0).Result("mm")
y_2 = sh.CellsSRC(visSectionFirstComponent, i + 1, 1).Result("mm")

TwoPtDistance = ((x_2 - x_1) ^ 2 + (y_2 - y_1) ^ 2) ^ (1 / 2)

End Function
