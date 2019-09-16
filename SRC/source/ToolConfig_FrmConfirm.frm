VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ToolConfig_FrmConfirm 
   Caption         =   "Confirm"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4788
   OleObjectBlob   =   "ToolConfig_FrmConfirm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ToolConfig_FrmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Const CON_STAR As String = "*"
Public m_strForRun_JobSourceName As String
Public m_strForRun_JobTargetName As String
Public m_strForRun_JobType As String

Private Const CON_JOB As String = "DTJobListSheet"
Public m_lConfigFunctionIndex As Long

Private m_lUsedRangeCount As Long
Private m_lUsedRangeCount2 As Long

Public FrmResult As String

Private Sub RunAction()
    ConfigInfor_Initialize
    
    FlowAction
    InstancesAction
    JobAction
    
    MsgBox "Processing is complete!", vbOKOnly, "Information"
    
    FrmConfirmResult = "Done"
    
    Unload Me
End Sub

Private Sub JobAction()
    If OptionButtonYesNewJob.Value = False Then
        Exit Sub
    End If

    Dim i As Long
    Dim l_sheetJob As Worksheet
    
    Set l_sheetJob = ThisWorkbook.Worksheets(getSheetNameByKey(CON_JOB, 1, 1))
    
    Dim l_strSourceJobName As String
    Dim l_strTargetJobName As String
    
    l_strSourceJobName = ListBoxJobSelected.List(0)
    l_strTargetJobName = txtboxJobName.Text
    
    Dim l_lSourceJobIndex As Long
    Dim l_lTargetJobIndex As Long
    
    m_lUsedRangeCount = getMaxRow(l_sheetJob)
    l_lTargetJobIndex = m_lUsedRangeCount + 1
    
    For i = 5 To m_lUsedRangeCount
        If l_sheetJob.Cells(i, 2) = l_strSourceJobName Then
            l_lSourceJobIndex = i
            Exit For
        End If
    Next
    
    copyRange l_sheetJob, l_lSourceJobIndex, 1, l_lSourceJobIndex, 30, l_sheetJob, l_lTargetJobIndex, 1
    l_sheetJob.Cells(l_lTargetJobIndex, 2) = l_strTargetJobName
    l_sheetJob.Cells(l_lTargetJobIndex, 4) = ToolConfig_FrmMain.m_strForRun_InstanceTargetSheeName
    l_sheetJob.Cells(l_lTargetJobIndex, 5) = ToolConfig_FrmMain.m_strForRun_FlowTargetSheeName
End Sub

Private Sub SetupMyMultiColumn(p_ListBox As MSForms.ListBox)
    Dim width As Integer
    Dim i As Integer
    Dim l_strMaxString As String
    If p_ListBox.ListCount > 0 Then
        For i = 0 To p_ListBox.ListCount - 1
            If Len(l_strMaxString) < Len(p_ListBox.List(i)) Then
                l_strMaxString = p_ListBox.List(i)
            End If
        Next
        width = Len(l_strMaxString) * 6
        p_ListBox.ColumnWidths = width
    End If
End Sub

Private Function getSheetNameByKey(p_strKey As String, row As Integer, col As Integer) As String
    Dim i As Long
    Dim l_iSheetCount As Integer
    Dim l_iCount As Integer
    Dim l_strListResult() As String
    l_iSheetCount = 0
    l_iCount = 0

    For i = 1 To ActiveWorkbook.Sheets.Count
'''        ThisWorkbook.Sheets(i).Activate
        If Len(ActiveWorkbook.Sheets(i).Cells(row, col).Text) >= Len(p_strKey) Then
            If p_strKey = VBA.Mid(VBA.Trim(ActiveWorkbook.Sheets(i).Cells(row, col).Text), 1, Len(p_strKey)) Then
                getSheetNameByKey = ActiveWorkbook.Sheets(i).Name
                Exit Function
            End If
        End If
    Next i
    getSheetNameByKey = ""
End Function

Private Sub FlowAction()
    Dim i As Long
    Dim l_sheetFlow As Worksheet
    Dim l_sheetTagertFlow As Worksheet
    Dim l_iTagertFlowUsedRange As Long
    Dim l_iTagertFlowIndexForAdd As Long
    Set l_sheetFlow = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_FlowSourceSheeName)
    
    Dim l_iStartNewIndex As Long

    'If new then new
    If ToolConfig_FrmMain.OptionButtonNewFlow.Value = True Then
        createSheet ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_FlowSourceSheeName), ToolConfig_FrmMain.m_strForRun_FlowTargetSheeName
        'set target sheet
        l_iTagertFlowUsedRange = 4
        l_iTagertFlowIndexForAdd = 5
        Set l_sheetTagertFlow = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_FlowTargetSheeName)
    Else
        Set l_sheetTagertFlow = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_FlowTargetSheeName)
        l_iTagertFlowUsedRange = getMaxRow(l_sheetTagertFlow)
        l_iTagertFlowIndexForAdd = l_iTagertFlowUsedRange + 1
    End If
    
    Dim l_iNowRowIndex As Long
    m_lUsedRangeCount = getMaxRow(l_sheetFlow)
    For i = 5 To m_lUsedRangeCount
        If checkTestItem(l_sheetFlow.Cells(i, 8)) = True Then
            m_lUsedRangeCount2 = getMaxRow(l_sheetTagertFlow)
            If ToolConfig_FrmMain.OptionButtonReplacetheSelectFlow.Value = True Then
                l_iNowRowIndex = findTestItemFromFLowSheet(l_sheetTagertFlow, 5, l_sheetFlow.Cells(i, 8))
                If l_iNowRowIndex > m_lUsedRangeCount2 Then
                    copyRange l_sheetFlow, i, 1, i, 145, l_sheetTagertFlow, l_iNowRowIndex, 1
                End If
            Else
                l_iNowRowIndex = m_lUsedRangeCount2 + 1
                'copy
                copyRange l_sheetFlow, i, 1, i, 50, l_sheetTagertFlow, l_iNowRowIndex, 1
            End If
            replaceFlow l_sheetTagertFlow, l_iNowRowIndex
        End If
    Next
End Sub

Private Function findTestItemFromFLowSheet(p_sheetFlow As Worksheet, p_iStartRow As Long, p_strTestItemName As String) As Long
    Dim i As Long
    m_lUsedRangeCount = getMaxRow(p_sheetFlow)
    For i = p_iStartRow To m_lUsedRangeCount
        If p_sheetFlow.Cells(i, 8) = p_strTestItemName Then
            findTestItemFromFLowSheet = i
            Exit Function
        End If
    Next
    findTestItemFromFLowSheet = m_lUsedRangeCount + 1
End Function

Private Sub replaceFlow(p_sheetTarget As Worksheet, p_iTRow As Long)
    'change TestItemName
    Dim l_strTestItemName As String
    Dim l_sheetConfig As Worksheet
    Set l_sheetConfig = ThisWorkbook.Sheets("ToolConfig")
    Dim l_strConvertedHead As String
    Dim l_strConvertedEND As String
    
    l_strConvertedHead = l_sheetConfig.Cells(m_lConfigFunctionIndex, 3)
    l_strConvertedEND = l_sheetConfig.Cells(m_lConfigFunctionIndex, 4)
    l_strTestItemName = p_sheetTarget.Cells(p_iTRow, 8)
    l_strTestItemName = l_strConvertedHead + l_strTestItemName
    If l_strConvertedEND <> "" Then
        l_strTestItemName = l_strTestItemName + l_strConvertedEND 'VBA.Mid(l_strTestItemName, 1, Len(l_strTestItemName) - l_strConvertedEND)
    End If
    p_sheetTarget.Cells(p_iTRow, 8) = l_strTestItemName
End Sub

Private Sub copyRange(p_sheetSource As Worksheet, p_iSRow As Long, p_iSCol As Integer, p_iSRow2 As Long, p_iSCol2 As Integer, _
                        p_sheetTarget As Worksheet, p_iTRow As Long, p_iTCol As Integer)
                        
    p_sheetTarget.Activate
    p_sheetTarget.Rows(p_iTRow & ":" & p_iTRow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        
    p_sheetSource.Activate
    p_sheetSource.Range(p_sheetSource.Cells(p_iSRow, p_iSCol), p_sheetSource.Cells(p_iSRow2, p_iSCol2)).Select
    p_sheetSource.Range(p_sheetSource.Cells(p_iSRow, p_iSCol), p_sheetSource.Cells(p_iSRow2, p_iSCol2)).Copy
    p_sheetTarget.Activate
    p_sheetTarget.Cells(p_iTRow, p_iTCol).PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False

End Sub

Private Sub InstancesAction()
    Dim i As Long
    Dim l_sheetInstances As Worksheet
    Dim l_sheetTagertInstances As Worksheet
    Dim l_iTagertInstancesUsedRange As Long
    Dim l_iTagertInstancesIndexForAdd As Long
    Set l_sheetInstances = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_InstanceSourceSheeName)
    
    Dim l_iStartNewIndex As Long

    'If new then new
    If ToolConfig_FrmMain.OptionButtonNewInstance.Value = True Then
        createSheet ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_InstanceSourceSheeName), ToolConfig_FrmMain.m_strForRun_InstanceTargetSheeName
        'set target sheet
        l_iTagertInstancesUsedRange = 4
        l_iTagertInstancesIndexForAdd = 5
        Set l_sheetTagertInstances = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_InstanceTargetSheeName)
    Else
        Set l_sheetTagertInstances = ActiveWorkbook.Worksheets(ToolConfig_FrmMain.m_strForRun_InstanceTargetSheeName)
        l_iTagertInstancesUsedRange = getMaxRow(l_sheetTagertInstances)
        l_iTagertInstancesIndexForAdd = l_iTagertInstancesUsedRange + 1
    End If
    
    Dim j As Long
    Dim l_iNowRowIndex As Long
    m_lUsedRangeCount = getMaxRow(l_sheetInstances)
    
    For i = 5 To m_lUsedRangeCount
        If checkTestItem(l_sheetInstances.Cells(i, 2)) = True Then
            m_lUsedRangeCount2 = getMaxRow(l_sheetTagertInstances)
            If ToolConfig_FrmMain.OptionButtonReplaceInstance.Value = True Then
                l_iNowRowIndex = findTestItemFromInstancesSheet(l_sheetTagertInstances, 5, l_sheetInstances.Cells(i, 2))
                If l_iNowRowIndex > m_lUsedRangeCount2 Then
                    copyRange l_sheetInstances, i, 1, i, 145, l_sheetTagertInstances, l_iNowRowIndex, 1
                End If
            Else
                l_iNowRowIndex = m_lUsedRangeCount2 + 1
                'copy, add one more line o avoid low effective production, Jeremy, 9/10/2019
                copyRange l_sheetInstances, i, 1, i, 145, l_sheetTagertInstances, l_iNowRowIndex + 1, 1
            End If
            'add one more line o avoid low effective production, Jeremy, 9/10/2019
            replaceInstance l_sheetTagertInstances, l_iNowRowIndex + 1
        End If
    Next
    
End Sub

Private Function findTestItemFromInstancesSheet(p_sheetInstances As Worksheet, p_iStartRow As Long, p_strTestItemName As String) As Long
    Dim i As Long
    m_lUsedRangeCount = getMaxRow(p_sheetInstances)
    For i = p_iStartRow To m_lUsedRangeCount
        If p_sheetInstances.Cells(i, 2) = p_strTestItemName Then
            findTestItemFromInstancesSheet = i
            Exit Function
        End If
    Next
    findTestItemFromInstancesSheet = m_lUsedRangeCount + 1
End Function

Private Sub replaceInstance(p_sheetTarget As Worksheet, p_iTRow As Long)
    'change TestItemName
    Dim l_strTestItemName As String
    Dim l_sheetConfig As Worksheet
    Set l_sheetConfig = ThisWorkbook.Sheets("ToolConfig")
    Dim l_strConvertedHead As String
    Dim l_strConvertedEND As String
    Dim l_strDC_Specs_Category As String
    Dim l_strFunctionName As String
    
    l_strFunctionName = l_sheetConfig.Cells(m_lConfigFunctionIndex, 2)
    l_strConvertedHead = l_sheetConfig.Cells(m_lConfigFunctionIndex, 3)
    l_strConvertedEND = l_sheetConfig.Cells(m_lConfigFunctionIndex, 4)
    l_strDC_Specs_Category = l_sheetConfig.Cells(m_lConfigFunctionIndex, 5)
    l_strTestItemName = p_sheetTarget.Cells(p_iTRow, 2)
    l_strTestItemName = l_strConvertedHead + l_strTestItemName
    If l_strConvertedEND <> "" Then
        l_strTestItemName = l_strTestItemName + l_strConvertedEND 'VBA.Mid(l_strTestItemName, 1, Len(l_strTestItemName) - l_strConvertedEND)
    End If
    p_sheetTarget.Cells(p_iTRow, 2) = l_strTestItemName
    
    p_sheetTarget.Cells(p_iTRow, 4) = l_strFunctionName
    
    If l_strDC_Specs_Category <> CON_STAR Then
        p_sheetTarget.Cells(p_iTRow, 6) = l_strDC_Specs_Category
    End If
    
    Dim i As Long
    Dim l_iblankCount As Integer
    l_iblankCount = 0
    For i = 6 To 145
        If l_iblankCount > 10 Then
            Exit For
        End If
        If l_sheetConfig.Cells(m_lConfigFunctionIndex, i) = "" Then
            l_iblankCount = l_iblankCount + 1
        Else
            l_iblankCount = 0
        End If
        If l_sheetConfig.Cells(m_lConfigFunctionIndex, i) <> CON_STAR Then
            p_sheetTarget.Cells(p_iTRow, i + 9) = l_sheetConfig.Cells(m_lConfigFunctionIndex, i)
        End If
    Next
End Sub

Private Sub createSheet(p_sheetSource As Worksheet, p_strSheetName As String)
    ThisWorkbook.Worksheets.Add().Name = p_strSheetName
    ThisWorkbook.Worksheets(p_strSheetName).Cells(1, 1) = p_sheetSource.Cells(1, 1)
End Sub

Private Function setInstanceDate(p_sheetInstance As Worksheet, p_iRow As Long)
    Dim l_strFunctionName As String
    Dim l_sheetConfig As Worksheet
    Set l_sheetConfig = ThisWorkbook.Sheets("ToolConfig")
    Dim l_strarg As String
    Dim i As Integer
    
    'DC_Specs/Category
    If l_sheetConfig.Cells(m_lConfigFunctionIndex, 5) <> CON_STAR Then
        p_sheetInstance.Cells(p_iRow, 6) = l_sheetConfig.Cells(m_lConfigFunctionIndex, 5)
    End If
    
    For i = 6 To 145
        l_strarg = l_sheetConfig.Cells(m_lConfigFunctionIndex, 6)
        If l_strarg <> CON_STAR Then
            p_sheetInstance.Cells(p_iRow, 14 + i) = l_strarg
        End If
    Next
End Function

Private Function checkTestItem(p_strTestItem As String)
    Dim i As Long
    For i = 0 To ToolConfig_FrmMain.ListBoxTestItemTarget.ListCount - 1
        If ToolConfig_FrmMain.ListBoxTestItemTarget.List(i) = p_strTestItem Then
            checkTestItem = True
            Exit Function
        End If
    Next
    checkTestItem = False
End Function

Private Sub ConfigInfor_Initialize()
    Dim i As Long
    Dim l_sheetConfig As Worksheet
    Set l_sheetConfig = ThisWorkbook.Sheets("ToolConfig")
    m_lUsedRangeCount = getMaxRow(l_sheetConfig)
    For i = 1 To m_lUsedRangeCount
        If l_sheetConfig.Cells(i, 2) = ToolConfig_FrmMain.ListBoxFunctionSelected.List(0) Then
            m_lConfigFunctionIndex = i
            Exit For
        End If
    Next
End Sub

Private Function changeFunctionName(p_strName As String) As String
    Dim l_strFunctionName As String
    Dim l_sheetConfig As Worksheet
    Set l_sheetConfig = ThisWorkbook.Sheets("ToolConfig")
    
    Dim l_strConvertedHead As String
    Dim l_strConvertedEND As String
    Dim l_strResult As String
    
    l_strConvertedHead = l_sheetConfig.Cells(m_lConfigFunctionIndex, 3)
    l_strConvertedEND = l_sheetConfig.Cells(m_lConfigFunctionIndex, 4)
    
    l_strResult = p_strName
    
    l_strResult = l_strConvertedHead + l_strConvertedHead
    
    
    If l_strConvertedEND <> "" Then
        l_strResult = Left(l_strResult, Len(l_strResult) - l_strConvertedEND)
    End If
    
    changeFunctionName = l_strResult
End Function

Private Sub cmdbtnBack_Click()
   Unload Me
End Sub

Private Sub cmdbtnDone_Click()
    RunAction
End Sub

Private Sub cmdbtnFunctionMoveL_Click()
    SingalmoveToLeft ListBoxJobAll, ListBoxJobSelected, cmdbtnFunctionMoveL, cmdbtnFunctionMoveR
End Sub

Private Sub cmdbtnFunctionMoveR_Click()
    SingalmoveToRight ListBoxJobAll, ListBoxJobSelected, cmdbtnFunctionMoveL, cmdbtnFunctionMoveR
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub OptionButtonNoNewJob_Click()
    setDoneButtonState
End Sub

Private Sub OptionButtonYesNewJob_Click()
    setDoneButtonState
End Sub

Private Sub txtboxJobName_Change()
    setDoneButtonState
End Sub

Private Sub UserForm_Initialize()
    OptionButtonNoNewJob.Value = True
    Dim i As Long
    Dim l_sheetJob As Worksheet
    Dim l_strJobName As String
    
'''    ThisWorkbook.Worksheets(getSheetNameByKey(CON_JOB, 1, 1)).Activate
    Set l_sheetJob = ActiveWorkbook.Worksheets(getSheetNameByKey(CON_JOB, 1, 1))
    m_lUsedRangeCount = getMaxRow(l_sheetJob)
    For i = 5 To m_lUsedRangeCount
        l_strJobName = l_sheetJob.Cells(i, 2)
        If l_strJobName <> "" Then
            ListBoxJobAll.AddItem l_strJobName
        End If
    Next
    
    SetupMyMultiColumn ListBoxJobAll
    ListBoxJobSelected.ColumnWidths = ListBoxJobAll.ColumnWidths
End Sub

Private Sub setDoneButtonState()
    If OptionButtonYesNewJob.Value = True Then
        If ListBoxJobSelected.ListCount > 0 And txtboxJobName.Text <> "" Then
            cmdbtnDone.Enabled = True
        Else
            cmdbtnDone.Enabled = False
        End If
    Else
        cmdbtnDone.Enabled = True
    End If
End Sub

Private Function getMaxRow(p_sheet As Worksheet) As Long
    Dim i As Long
    For i = 5 To 65000
        If p_sheet.Cells(i, 1) = "" And p_sheet.Cells(i, 2) = "" And p_sheet.Cells(i, 3) = "" And p_sheet.Cells(i, 4) = "" And _
            p_sheet.Cells(i, 5) = "" And p_sheet.Cells(i, 6) = "" And p_sheet.Cells(i, 7) = "" And p_sheet.Cells(i, 8) = "" And _
            p_sheet.Cells(i, 9) = "" And p_sheet.Cells(i, 10) = "" Then
            getMaxRow = i - 1
            Exit Function
        End If
    Next
End Function

Private Sub SingalmoveToLeft(p_ListBoxL As MSForms.ListBox, p_ListBoxR As MSForms.ListBox, p_btnMoveToL As MSForms.CommandButton, p_btnMoveToR As MSForms.CommandButton)
    p_ListBoxL.AddItem p_ListBoxR.List(0)
    p_ListBoxR.RemoveItem (0)
    p_btnMoveToL.Enabled = False
    p_btnMoveToR.Enabled = True
    setDoneButtonState
End Sub

Private Sub SingalmoveToRight(p_ListBoxL As MSForms.ListBox, p_ListBoxR As MSForms.ListBox, p_btnMoveToL As MSForms.CommandButton, p_btnMoveToR As MSForms.CommandButton)
    Dim i As Long
    For i = 0 To p_ListBoxL.ListCount - 1
        If p_ListBoxL.Selected(i) = True Then
            p_ListBoxR.AddItem p_ListBoxL.List(i), 0
            p_ListBoxL.RemoveItem i
            p_btnMoveToL.Enabled = True
            p_btnMoveToR.Enabled = False
            setDoneButtonState
            Exit For
        End If
    Next
    setDoneButtonState
End Sub
