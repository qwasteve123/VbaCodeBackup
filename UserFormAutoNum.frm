VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAutoNum 
   Caption         =   "Auto Numbering"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   OleObjectBlob   =   "UserFormAutoNum.frx":0000
End
Attribute VB_Name = "UserFormAutoNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub ButtonAdd_Click()
If IsNumeric(TextBoxAddUp.Value) Then
    Call PlusNum(TextBoxAddUp.Value)
End If
End Sub

Private Sub ButtonChangeNum_Click()
If IsNumeric(TextBoxAntNum.Value) And TextBoxAntNum > 0 Then
    AutoNum = TextBoxAntNum.Value - 1
    LabelNextNum.Caption = "Next Number: " & AutoNum + 1
Else
    MsgBox "Please enter integer greater than 0."
    Exit Sub
End If
End Sub


Private Sub ButtonExportToXl_Click()
Dim is_export As Integer

 is_export = MsgBox("Export link budget to excel form?", vbYesNo)
 If is_export = vbYes Then
    Call A_Main.ExportLinkBudget
 End If
End Sub

Private Sub ButtonLength_Click()
Call DF_ReRouting.LineToScale
End Sub

Private Sub ButtonRoute_Click()
If ActiveWindow.Selection.Count <> 1 And CheckButtonState = 0 Then
    MsgBox "Plase select only one shape."
    GoTo exitsub
End If

If CheckButtonState = 0 Then
    ButtonRoute.Caption = "End"
    CheckButtonState = 1
    Call ThisDocument.Reroute
Else
    stop_sub = True
    CheckButtonState = 0
    Unload UserFormAutoNum
    Call DC_AutoNumbering.ActivateForm
End If
exitsub:
End Sub

Private Sub ButtonShowLinkBud_Click()
 Call A_Main.ShowLinkBudget
End Sub

Private Sub ButtonStartEnd_Click()

If CheckButtonState = 0 Then
    ButtonStartEnd.Caption = "End"
    CheckButtonState = 1
    Call AutoShapeNum
Else
    stop_sub = True
    Unload UserFormAutoNum
    Call DC_AutoNumbering.ActivateForm
End If
End Sub
Private Sub CommandButton2_Click()

If ActiveWindow.Selection.Count <> 1 Then
    MsgBox "Plase select only one shape."
    GoTo exitsub
End If

Call ThisDocument.ReturnLine
exitsub:
End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub ListLayer_Change()

    'Dim item As Variant
    Dim vsoLayer As Visio.Layer
    
On Error Resume Next
    
        If AutoPage = ActivePage.Name Then
            If ActivePage.Layers.Count > 0 Then
                For i = LBound(ListLayer.list) To UBound(ListLayer.list)
                    If ListLayer.list(i) = ActivePage.Layers(i + 1).Name Then
                        If ListLayer.Selected(i) Then
                            Set vsoLayer = Application.ActiveWindow.Page.Layers.item(i + 1)
                            vsoLayer.CellsC(visLayerLock).FormulaU = "1"
                        Else
                            Set vsoLayer = Application.ActiveWindow.Page.Layers.item(i + 1)
                            vsoLayer.CellsC(visLayerLock).FormulaU = "0"
                        End If
                    End If
                Next
            End If
        Else
            UserFormAutoNum.Hide
            ListLayer.Clear
            
            If ActivePage.Layers.Count > 0 Then
                For i = 1 To ActivePage.Layers.Count
                    UserFormAutoNum.ListLayer.AddItem ActivePage.Layers(i).Name
                    Set vsoLayer = Application.ActiveWindow.Page.Layers.item(i)
                    If vsoLayer.CellsC(visLayerLock).FormulaU = "1" Then
                        UserFormAutoNum.ListLayer.Selected(i - 1) = True
                    End If
                Next
            End If
            
            AutoPage = ActivePage.Name
            UserFormAutoNum.Show vbModeless
        End If
End Sub

Private Sub MultiPage1_Change()
    If MultiPage1.SelectedItem.Name = "PgFindShape" Then
        Unload UserFormAutoNum
        Call DD_HighlightComponent.ActivateForm
    End If
End Sub

Private Sub UserForm_Initialize()

LabelNextNum.Caption = "Next Number:" & AutoNum + 1
MultiPage1.Pages(1).Visible = True

End Sub

Private Sub Op1_Click()
    If Op2.Value = True Then
        Op2.Value = False
    End If
End Sub
Private Sub Op2_Click()
    If Op1.Value = True Then
        Op1.Value = False
    End If
End Sub
Private Sub ButtonLabel_Click()
Dim label_is_lift As Boolean

If Op1.Value = False And Op2.Value = False Then
    MsgBox "Please choose a naming format."
    GoTo exitsub
End If

label_is_lift = False
If Op1.Value = True Then
    label_is_lift = True
End If
    
Call DE_ChangeLabelFormat.ChangeLabelName(label_is_lift)

exitsub:
End Sub

Private Sub UserForm_Terminate()

End Sub
