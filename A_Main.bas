Attribute VB_Name = "A_Main"
Public Sub ExportLinkBudget()
Attribute ExportLinkBudget.VB_ProcData.VB_Invoke_Func = "F"

    CheckError (ErrRemoveCircle)
    
    Call GetInformation

    For PageNum = 1 To ActiveDocument.Pages.Count

        Call ShapeData(PageNum)
        Call ShapeConnection(PageNum)
    
    Next PageNum
    
    Call CountFloor
    Call SetLabelID
    
    PageName = ""
    
    Call GetSector
    Call findlinkage
    Call ShuffleRelation
    
    AntCount = AntCount - 1
    
    Call ExcelFile
    
    Call TerminateSub
    


End Sub

Public Sub ShowLinkBudget()
Attribute ShowLinkBudget.VB_ProcData.VB_Invoke_Func = "D"

    CheckError (ErrRemoveCircle)
    
    Application.ScreenUpdating = False
    
    Call GetInformation

'    For PageNum = 1 To ActiveDocument.Pages.Count
'
'        Call ShapeData(PageNum)
'        Call ShapeConnection(PageNum)
'
'    Next PageNum

    PageNum = ActivePage.Index
        Call ShapeData(PageNum)
        Call ShapeConnection(PageNum)
    
    Call CountFloor
    Call SetLabelID

    PageName = ""
    
    Call GetSector
    Call CountFloor
    Call findlinkage
    Call ShuffleRelation

    AntCount = AntCount - 1
    
    Call LabelToLinkBud
    
    Application.ScreenUpdating = True
    
    Call TerminateSub
    
End Sub

Sub TerminateSub()
    Erase Relation, LinkPath, ShapeDataList, MaterialList, LinkageList, FloorList, SectorList, BondMatList
    RelationMaxNo = 0
    ShapeMaxRow = 0
    LinkNumOfRow = 0
    LinkSecRow = 0
    FloorMaxRow = 0
    SectorMaxRow = 0
    SectorMaxNum = 0
    RSRP_output = 0
    FSPL = 0
    FSPL_lift = 0
    FreqChoice = 0
End Sub
