Attribute VB_Name = "F_ErrorList"
Public Sub CheckError(ErrorID As Integer)
    
Select Case ErrorID

    Case ErrRemoveCircle
        Call ErrorRmCircle
        
    Case ErrConnector
            If vsoConnects.Count = 1 Then 'ErrorConnector
                Call ErrorConnectors
            End If

        
    Case ErrWalkGlue
        If InStr(vsoConnectfromCell.Formula, "WALKGLUE") Then
            Call ErrorWalkGlue
        End If
                  
    Case ErrLinkage
        For TestLinkRow = LinkSecRow + 1 To ShapeMaxRow
            If ShapeDataList(TestLinkRow, shddata3) = ShapeDataList(LinkSecRow, shddata3) Then
                Call ErrorLinkages
            End If
        Next
        
'    Case ErrSamePort
'        For j = 1 To 2
'            For k = 1 To PortRowCount
'                If

    Case ErrItemNumNotInteger
        Call ErrorNotInteger
        
        
End Select
End Sub
Public Sub ErrorRmCircle()
    
    ActiveWindow.DeselectAll
        Set vsoShapes = ActivePage.Shapes
        For Each vsoShape In vsoShapes
            If InStr(vsoShape.Name, "Error Circle") Then
                vsoShape.Delete
            End If
        Next

End Sub


Public Sub ErrorConnectors()

Dim floor As String
Dim ItemNo As String
Dim ErrorShape As Visio.Shape
Dim CoorX As Double
Dim CoorY As Double

floor = vsoShape.CellsSRC(visSectionProp, 1, visCustPropsValue)
ItemNo = vsoShape.CellsSRC(visSectionProp, 0, visCustPropsValue)

    MsgBox "Connector" & vbNewLine & _
    "Name:" & " Page" & PageNum & "_" & vsoShape.Name & vbNewLine & _
    "Label: S" & floor & "." & ItemNo & vbNewLine & _
    "is connected to nothing or unknown shape"

    If InStr(vsoConnectfromCell.Name, "Begin") Then
        CoorX = vsoShape.Cells("EndX").Result("")
        CoorY = vsoShape.Cells("EndY").Result("")
    Else
        CoorX = vsoShape.Cells("BeginX").Result("")
        CoorY = vsoShape.Cells("BeginY").Result("")
    End If
    
ActiveDocument.Pages.item(PageNum).Drop ActiveDocument.Masters.ItemU("Error Circle"), CoorX, CoorY
Application.ActiveWindow.DeselectAll

End
End Sub

Public Sub ErrorWalkGlue()

Dim floor As String
Dim ItemNo As String

floor = vsoShape.CellsSRC(visSectionProp, 1, visCustPropsValue)
ItemNo = vsoShape.CellsSRC(visSectionProp, 0, visCustPropsValue)

    MsgBox "Connector" & vbNewLine & _
    "Name:" & " Page" & PageNum & "_" & vsoShape.Name & vbNewLine & _
    "Label: S" & floor & "." & ItemNo & vbNewLine & _
    "is connected to nothing or unknown shape"

    If InStr(vsoConnectfromCell.Name, "Begin") Then
        CoorX = vsoShape.Cells("BeginX").Result("")
        CoorY = vsoShape.Cells("BeginY").Result("")
    Else
        CoorX = vsoShape.Cells("EndX").Result("")
        CoorY = vsoShape.Cells("EndY").Result("")
    End If

    ActiveDocument.Pages.item(PageNum).Drop ActiveDocument.Masters.ItemU("Error Circle"), CoorX, CoorY
    Application.ActiveWindow.DeselectAll

End
End Sub

Public Sub ErrorLinkages()
      
MsgBox "Error Exist: Same Linkage Name " & ShapeDataList(LinkSecRow, shddata3) & vbNewLine & _
ShapeDataList(LinkRow, shdCompName) & vbNewLine & _
ShapeDataList(LinkSecRow, shdCompName) & vbNewLine & _
ShapeDataList(TestLinkRow, shdCompName) & vbNewLine
End

End Sub

Public Sub ErrorNotInteger()

    For Each vsoShape In ActiveWindow.Selection
        If (vsoShape.CellsSRC(visSectionProp, 1, visCustPropsValue)) > 0 Then
            If Not IsNumeric(vsoShape.Cells("Prop.item_no").Formula) Then
                MsgBox "Shape: " & vsoShape.Name & " item no. is in string format or empty" & vbNewLine & _
                "Please enter an integer in item no. in shape data for labelling"
                
                CoorX = vsoShape.Cells("PinX").Result("")
                CoorY = vsoShape.Cells("PinY").Result("")
                
                ActivePage.Drop ActiveDocument.Masters.ItemU("Error Circle"), CoorX, CoorY
                Application.ActiveWindow.DeselectAll

                End
            End If
        End If
    Next
    
End Sub
