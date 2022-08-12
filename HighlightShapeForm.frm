VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HighlightShapeForm 
   Caption         =   "Highlight Shapes"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7110
   OleObjectBlob   =   "HighlightShapeForm.frx":0000
End
Attribute VB_Name = "HighlightShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Button_Find_Click()
    DeleteCircles
    Me.ListBox_Comp.Clear
    Call HighLight(Combo_floor.Value, Combo_type.Value, Combo_cable.Value)
    LabelCompFound.Caption = ListBox_Comp.ListCount
End Sub


Private Sub MultiPage1_Change()
Dim page_num As Integer
    If MultiPage1.SelectedItem.Name <> "PgFindShape" Then
        page_num = MultiPage1.SelectedItem.Index
        Unload HighlightShapeForm
        Call DC_AutoNumbering.ActivateForm
        UserFormAutoNum.MultiPage1.Value = page_num
    End If
End Sub

Private Sub UserForm_Initialize()
    MultiPage1.Pages(3).Visible = True
End Sub

Private Sub UserForm_Terminate()

Call DeleteCircles

End Sub

Private Sub DeleteCircles()

Dim PageNum As Integer
Dim coll As Collection
Dim vsoShapes As Visio.Shapes
Dim vsoShape As Visio.Shape
    
    Set coll = New Collection
    
    For PageNum = 1 To ActiveDocument.Pages.Count
        
         Set vsoShapes = ActiveDocument.Pages.item(PageNum).Shapes
         
        For Each vsoShape In vsoShapes
            If InStr(vsoShape.Name, "Error Circle") Then
                coll.add vsoShape
                Debug.Print vsoShape.Name
            End If
        Next
    Next
    
    For Each vsoShape In coll
        vsoShape.Delete
    Next

End Sub
