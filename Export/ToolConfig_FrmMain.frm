VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ToolConfig_FrmMain 
   Caption         =   "ToolSetup"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18435
   OleObjectBlob   =   "ToolConfig_FrmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ToolConfig_FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Const CON_FLOW As String = "DTFlowtableSheet"
Private Const CON_INSTANCE As String = "DTTestInstancesSheet"

Private Const CON_CONFIGSHEETNAME As String = "ToolConfig"
Private Const CON_CONFIGSHEETNAME2 As String = "ToolConfig (2)"

Private m_strBeforeFlowName As String
Private m_lUsedRangeCount As Long
Private m_lUsedRangeCount2 As Long

Public m_strForRun_FunctionSheeName As String
Public m_strForRun_FlowSourceSheeName As String
Public m_strForRun_FlowTargetSheeName As String
Public m_strForRun_FlowType As String
Public m_strForRun_InstanceSourceSheeName As String
Public m_strForRun_InstanceTargetSheeName As String
Public m_strForRun_InstanceType As String

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

Private Sub ListBoxTestItemAll_Click()

End Sub

Private Sub UserForm_Initialize()
    Call disableAll
    If findSheetByName(CON_CONFIGSHEETNAME) = True Then
        cmdbtnFunctionSearch.Enabled = True
    End If
End Sub

Private Function setSearchItemButtonState() As Boolean
    If ListBoxFunctionSelected.Enabled = False Or ListBoxFunctionSelected.ListCount <= 0 Then
        setSearchItemButtonDisable
        setSearchItemButtonState = False
        Exit Function
    End If
    
    If ListBoxFlowSource.Enabled = False Or ListBoxFlowSource.ListCount <= 0 Then
        setSearchItemButtonDisable
        setSearchItemButtonState = False
        Exit Function
    End If
    
    If ListBoxInstanceSource.Enabled = False Or ListBoxInstanceSource.ListCount <= 0 Then
        setSearchItemButtonDisable
        setSearchItemButtonState = False
        Exit Function
    End If
    
    If (OptionButtonNewInstance.Value = True And TextBoxInstanceName.Text = "") Or _
        (OptionButtonNewInstance.Value = False And (ListBoxInstanceTarget.Enabled = False Or ListBoxInstanceTarget.ListCount <= 0)) Then
        setSearchItemButtonDisable
        setSearchItemButtonState = False
        Exit Function
    End If
    
    If (OptionButtonNewFlow.Value = True And TextBoxFlowName.Text = "") Or _
        (OptionButtonNewFlow.Value = False And (ListBoxFlowTarget.Enabled = False Or ListBoxFlowTarget.ListCount <= 0)) Then
        setSearchItemButtonDisable
        setSearchItemButtonState = False
        Exit Function
    End If
    
    setSearchItemButtonEnable
    
    setSearchItemButtonState = True
End Function

Private Sub setSearchItemButtonEnable()
    If cmdbtnItemSearch.Enabled = True Then
        Exit Sub
    End If
    If ListBoxFlowSource.ListCount = 0 Then
        m_strBeforeFlowName = ""
        ListBoxTestItemAll.Clear
        ListBoxTestItemTarget.Clear
    Else
        If m_strBeforeFlowName <> ListBoxFlowSource.List(0) Then
            m_strBeforeFlowName = ListBoxFlowSource.List(0)
            ListBoxTestItemAll.Clear
            ListBoxTestItemTarget.Clear
        End If
    End If
    cmdbtnItemSearch.Enabled = True
    ListBoxTestItemAll.Enabled = True
    cmdbtnTestItemMoveSR.Enabled = True
    cmdbtnTestItemMoveSL.Enabled = False
    cmdbtnTestItemMoveSRAll.Enabled = True
    cmdbtnTestItemMoveSLAll.Enabled = False
    ListBoxTestItemTarget.Enabled = True
    
    If ListBoxTestItemAll.ListCount > 0 Or ListBoxTestItemTarget.ListCount > 0 Then
        cmdbtnNext.Enabled = True
    Else
        cmdbtnNext.Enabled = False
    End If
End Sub

Private Sub setSearchItemButtonDisable()
    If cmdbtnItemSearch.Enabled = False Then
        Exit Sub
    End If
    cmdbtnItemSearch.Enabled = False
    If ListBoxFlowSource.ListCount = 0 Then
        m_strBeforeFlowName = ""
        ListBoxTestItemAll.Clear
        ListBoxTestItemTarget.Clear
    Else
        If m_strBeforeFlowName <> ListBoxFlowSource.List(0) Then
            m_strBeforeFlowName = ListBoxFlowSource.List(0)
            ListBoxTestItemAll.Clear
            ListBoxTestItemTarget.Clear
        End If
    End If

    cmdbtnItemSearch.Enabled = False
    ListBoxTestItemAll.Enabled = False
    cmdbtnTestItemMoveSR.Enabled = False
    cmdbtnTestItemMoveSL.Enabled = False
    cmdbtnTestItemMoveSRAll.Enabled = False
    cmdbtnTestItemMoveSLAll.Enabled = False
    ListBoxTestItemTarget.Enabled = False
    cmdbtnNext.Enabled = False
End Sub

Private Sub cmdbtnCancel_Click()
    Unload Me
End Sub

Private Sub cmdbtnItemSearch_Click()
    If ListBoxTestItemAll.ListCount > 0 Then
        Exit Sub
    End If

    Dim i As Long
    Dim l_strSourceFlowSheetName As String
    Dim l_sheetSourceFlow As Worksheet
    Dim l_strItem As String
    
    l_strSourceFlowSheetName = ListBoxFlowSource.List(0)
    m_strBeforeFlowName = l_strSourceFlowSheetName
    
    Set l_sheetSourceFlow = ActiveWorkbook.Sheets(l_strSourceFlowSheetName)
    m_lUsedRangeCount = getMaxRow(l_sheetSourceFlow)
    For i = 5 To m_lUsedRangeCount
        l_strItem = l_sheetSourceFlow.Cells(i, 8)
        If l_strItem <> "" Then
            ListBoxTestItemAll.AddItem l_strItem
        End If
    Next
    
    If ListBoxTestItemAll.ListCount > 0 Then
        ListBoxTestItemAll.Enabled = True
        cmdbtnTestItemMoveSR.Enabled = True
        cmdbtnTestItemMoveSL.Enabled = False
        cmdbtnTestItemMoveSRAll.Enabled = True
        cmdbtnTestItemMoveSLAll.Enabled = False
        ListBoxTestItemTarget.Enabled = True
        
        SetupMyMultiColumn ListBoxTestItemAll
        ListBoxTestItemTarget.ColumnWidths = ListBoxTestItemAll.ColumnWidths
    End If
End Sub

Private Sub cmdbtnNext_Click()

    m_strForRun_FunctionSheeName = ListBoxFunctionSelected.Text
    
    m_strForRun_FlowSourceSheeName = ListBoxFlowSource.List(0)
    If OptionButtonAddedtoFlow.Value = True Then
        m_strForRun_FlowType = "Add"
    ElseIf OptionButtonReplacetheSelectFlow.Value = True Then
        m_strForRun_FlowType = "Replace"
    Else
        m_strForRun_FlowType = "New"
    End If
    
    If m_strForRun_FlowType = "New" Then
        m_strForRun_FlowTargetSheeName = TextBoxFlowName.Text
    Else
        m_strForRun_FlowTargetSheeName = ListBoxFlowTarget.List(0)
    End If
    
    m_strForRun_InstanceSourceSheeName = ListBoxInstanceSource.List(0)
    If OptionButtonAddedtoSelectinstance.Value = True Then
        m_strForRun_InstanceType = "Add"
    ElseIf OptionButtonReplaceInstance.Value = True Then
        m_strForRun_InstanceType = "Replace"
    Else
        m_strForRun_InstanceType = "New"
    End If
    
    If m_strForRun_InstanceType = "New" Then
        m_strForRun_InstanceTargetSheeName = TextBoxInstanceName.Text
    Else
        m_strForRun_InstanceTargetSheeName = ListBoxInstanceTarget.List(0)
    End If
    
    'ReDim m_strForRun_TestItems(ListBoxTestItemTarget.ListCount)
    
    FrmConfirmResult = "Cancel"
    
    ToolConfig_FrmConfirm.Show 1
    
    If FrmConfirmResult = "Done" Then
        Unload Me
    End If
    
End Sub

Private Sub cmdbtnTestItemMoveSLAll_Click()
    MultimoveToLeftAtAll ListBoxTestItemAll, ListBoxTestItemTarget, cmdbtnTestItemMoveSL, cmdbtnTestItemMoveSR, cmdbtnTestItemMoveSLAll, cmdbtnTestItemMoveSRAll
End Sub

Private Sub cmdbtnTestItemMoveSRAll_Click()
    MultimoveToRightAtAll ListBoxTestItemAll, ListBoxTestItemTarget, cmdbtnTestItemMoveSL, cmdbtnTestItemMoveSR, cmdbtnTestItemMoveSLAll, cmdbtnTestItemMoveSRAll
End Sub

Private Sub cmdbtnTestItemMoveSR_Click()
    MultimoveToRightAtSingal ListBoxTestItemAll, ListBoxTestItemTarget, cmdbtnTestItemMoveSL, cmdbtnTestItemMoveSR, cmdbtnTestItemMoveSLAll, cmdbtnTestItemMoveSRAll
End Sub

Private Sub cmdbtnTestItemMoveSL_Click()
    MultimoveToLeftAtSingal ListBoxTestItemAll, ListBoxTestItemTarget, cmdbtnTestItemMoveSL, cmdbtnTestItemMoveSR, cmdbtnTestItemMoveSLAll, cmdbtnTestItemMoveSRAll
End Sub

Private Sub OptionButtonNewFlow_Click()
    If ListBoxFlowTarget.Enabled = True Then
        ListBoxFlowTarget.Enabled = False
        cmdbtnFlowMoveTL.Enabled = False
        cmdbtnFlowMoveTR.Enabled = False
    End If
    TextBoxFlowName.Enabled = True
    setSearchItemButtonState
End Sub

Private Sub OptionButtonReplacetheSelectFlow_Click()
    If ListBoxFlowTarget.Enabled = False Then
        ListBoxFlowTarget.Enabled = True
        If ListBoxFlowTarget.ListCount = 1 Then
            cmdbtnFlowMoveTL.Enabled = True
        Else
            If ListBoxFlowAll.ListCount > 0 Then
                cmdbtnFlowMoveTR.Enabled = True
            End If
        End If
    End If
    TextBoxFlowName.Enabled = False
    setSearchItemButtonState
End Sub

Private Sub OptionButtonAddedtoFlow_Click()
    If ListBoxFlowTarget.Enabled = False Then
        ListBoxFlowTarget.Enabled = True
        If ListBoxFlowTarget.ListCount = 1 Then
            cmdbtnFlowMoveTL.Enabled = True
        Else
            If ListBoxFlowAll.ListCount > 0 Then
                cmdbtnFlowMoveTR.Enabled = True
            End If
        End If
    End If
    TextBoxFlowName.Enabled = True
    setSearchItemButtonState
End Sub

Private Sub OptionButtonAddedtoSelectinstance_Click()
    If ListBoxInstanceTarget.Enabled = False Then
        ListBoxInstanceTarget.Enabled = True
        If ListBoxInstanceTarget.ListCount = 1 Then
            cmdbtnInstanceMoveTL.Enabled = True
        Else
            If ListBoxInstanceAll.ListCount > 0 Then
                cmdbtnInstanceMoveTR.Enabled = True
            End If
        End If
    End If
    TextBoxInstanceName.Enabled = False
    setSearchItemButtonState
End Sub

Private Sub OptionButtonNewInstance_Click()
    If ListBoxInstanceTarget.Enabled = True Then
        ListBoxInstanceTarget.Enabled = False
        cmdbtnInstanceMoveTR.Enabled = False
        cmdbtnInstanceMoveTL.Enabled = False
    End If
    TextBoxInstanceName.Enabled = True
    setSearchItemButtonState
End Sub

Private Sub OptionButtonReplaceInstance_Click()
    If ListBoxInstanceTarget.Enabled = False Then
        ListBoxInstanceTarget.Enabled = True
        If ListBoxInstanceTarget.ListCount = 1 Then
            cmdbtnInstanceMoveTL.Enabled = True
        Else
            If ListBoxInstanceAll.ListCount > 0 Then
                cmdbtnInstanceMoveTR.Enabled = True
            End If
        End If
    End If
    TextBoxInstanceName.Enabled = False
    setSearchItemButtonState
End Sub

Private Sub TextBoxFlowName_Change()
    setSearchItemButtonState
End Sub

Private Sub TextBoxInstanceName_Change()
    setSearchItemButtonState
End Sub

Private Sub cmdbtnFlowMoveSL_Click()
     SingalmoveToLeft ListBoxFlowAll, ListBoxFlowSource, cmdbtnFlowMoveSL, cmdbtnFlowMoveSR
End Sub

Private Sub cmdbtnFlowMoveSR_Click()
    SingalmoveToRight ListBoxFlowAll, ListBoxFlowSource, cmdbtnFlowMoveSL, cmdbtnFlowMoveSR
End Sub

Private Sub cmdbtnFlowMoveTL_Click()
    SingalmoveToLeft ListBoxFlowAll, ListBoxFlowTarget, cmdbtnFlowMoveTL, cmdbtnFlowMoveTR
End Sub

Private Sub cmdbtnFlowMoveTR_Click()
    SingalmoveToRight ListBoxFlowAll, ListBoxFlowTarget, cmdbtnFlowMoveTL, cmdbtnFlowMoveTR
End Sub

Private Sub cmdbtnFunctionMoveL_Click()
    SingalmoveToLeft ListBoxFunctionAll, ListBoxFunctionSelected, cmdbtnFunctionMoveL, cmdbtnFunctionMoveR
End Sub

Private Sub cmdbtnFunctionMoveR_Click()
    SingalmoveToRight ListBoxFunctionAll, ListBoxFunctionSelected, cmdbtnFunctionMoveL, cmdbtnFunctionMoveR
End Sub

Private Sub MultimoveToLeftAtAll(p_ListBoxL As MSForms.ListBox, _
                                p_ListBoxR As MSForms.ListBox, _
                                p_btnMoveToL1 As MSForms.CommandButton, _
                                p_btnMoveToR1 As MSForms.CommandButton, _
                                p_btnMoveToL2 As MSForms.CommandButton, _
                                p_btnMoveToR2 As MSForms.CommandButton)
                                
    Dim i As Long
    For i = 0 To p_ListBoxR.ListCount - 1
        p_ListBoxL.AddItem p_ListBoxR.List(i), 0
        p_btnMoveToR1.Enabled = True
        p_btnMoveToL1.Enabled = False
        p_btnMoveToR2.Enabled = True
        p_btnMoveToL2.Enabled = False
    Next
    p_ListBoxR.Clear
    
    If p_ListBoxL.ListCount = 0 Then
        p_btnMoveToR1.Enabled = False
        p_btnMoveToR2.Enabled = False
    Else
        p_btnMoveToR1.Enabled = True
        p_btnMoveToR2.Enabled = True
    End If

    If p_ListBoxR.ListCount = 0 Then
        p_btnMoveToL1.Enabled = False
        p_btnMoveToL2.Enabled = False
        cmdbtnNext.Enabled = False
    Else
        p_btnMoveToL1.Enabled = True
        p_btnMoveToL2.Enabled = True
        cmdbtnNext.Enabled = True
    End If
End Sub

Private Sub MultimoveToRightAtAll(p_ListBoxL As MSForms.ListBox, _
                                p_ListBoxR As MSForms.ListBox, _
                                p_btnMoveToL1 As MSForms.CommandButton, _
                                p_btnMoveToR1 As MSForms.CommandButton, _
                                p_btnMoveToL2 As MSForms.CommandButton, _
                                p_btnMoveToR2 As MSForms.CommandButton)
    Dim i As Long
    For i = 0 To p_ListBoxL.ListCount - 1
        p_ListBoxR.AddItem p_ListBoxL.List(i), i '0
        p_btnMoveToL1.Enabled = True
        p_btnMoveToR1.Enabled = False
        p_btnMoveToL2.Enabled = True
        p_btnMoveToR2.Enabled = False
    Next
    p_ListBoxL.Clear
    
    If p_ListBoxL.ListCount = 0 Then
        p_btnMoveToR1.Enabled = False
        p_btnMoveToR2.Enabled = False
    Else
        p_btnMoveToR1.Enabled = True
        p_btnMoveToR2.Enabled = True
    End If

    If p_ListBoxR.ListCount = 0 Then
        p_btnMoveToL1.Enabled = False
        p_btnMoveToL2.Enabled = False
        cmdbtnNext.Enabled = False
    Else
        p_btnMoveToL1.Enabled = True
        p_btnMoveToL2.Enabled = True
        cmdbtnNext.Enabled = True
    End If
End Sub

Private Sub MultimoveToLeftAtSingal(p_ListBoxL As MSForms.ListBox, _
                                p_ListBoxR As MSForms.ListBox, _
                                p_btnMoveToL1 As MSForms.CommandButton, _
                                p_btnMoveToR1 As MSForms.CommandButton, _
                                p_btnMoveToL2 As MSForms.CommandButton, _
                                p_btnMoveToR2 As MSForms.CommandButton)
                                
    Dim i As Long
    For i = p_ListBoxR.ListCount - 1 To 0 Step -1
        If p_ListBoxR.Selected(i) = True Then
            p_ListBoxL.AddItem p_ListBoxR.List(i), 0
            p_ListBoxR.RemoveItem i
        End If
    Next
    
    If p_ListBoxL.ListCount = 0 Then
        p_btnMoveToR1.Enabled = False
        p_btnMoveToR2.Enabled = False
    Else
        p_btnMoveToR1.Enabled = True
        p_btnMoveToR2.Enabled = True
    End If

    If p_ListBoxR.ListCount = 0 Then
        p_btnMoveToL1.Enabled = False
        p_btnMoveToL2.Enabled = False
        cmdbtnNext.Enabled = False
    Else
        p_btnMoveToL1.Enabled = True
        p_btnMoveToL2.Enabled = True
        cmdbtnNext.Enabled = True
    End If
    
End Sub

Private Sub MultimoveToRightAtSingal(p_ListBoxL As MSForms.ListBox, _
                                p_ListBoxR As MSForms.ListBox, _
                                p_btnMoveToL1 As MSForms.CommandButton, _
                                p_btnMoveToR1 As MSForms.CommandButton, _
                                p_btnMoveToL2 As MSForms.CommandButton, _
                                p_btnMoveToR2 As MSForms.CommandButton)
    Dim i As Long
    For i = p_ListBoxL.ListCount - 1 To 0 Step -1
        If p_ListBoxL.Selected(i) = True Then
            p_ListBoxR.AddItem p_ListBoxL.List(i), 0
            p_ListBoxL.RemoveItem i
        End If
    Next
    
    If p_ListBoxL.ListCount = 0 Then
        p_btnMoveToR1.Enabled = False
        p_btnMoveToR2.Enabled = False
    Else
        p_btnMoveToR1.Enabled = True
        p_btnMoveToR2.Enabled = True
    End If

    If p_ListBoxR.ListCount = 0 Then
        p_btnMoveToL1.Enabled = False
        p_btnMoveToL2.Enabled = False
        cmdbtnNext.Enabled = False
    Else
        p_btnMoveToL1.Enabled = True
        p_btnMoveToL2.Enabled = True
        cmdbtnNext.Enabled = True
    End If
End Sub

Private Sub SingalmoveToLeft(p_ListBoxL As MSForms.ListBox, p_ListBoxR As MSForms.ListBox, p_btnMoveToL As MSForms.CommandButton, p_btnMoveToR As MSForms.CommandButton)
    p_ListBoxL.AddItem p_ListBoxR.List(0)
    p_ListBoxR.RemoveItem (0)
    p_btnMoveToL.Enabled = False
    p_btnMoveToR.Enabled = True
    setSearchItemButtonState
End Sub

Private Sub SingalmoveToRight(p_ListBoxL As MSForms.ListBox, p_ListBoxR As MSForms.ListBox, p_btnMoveToL As MSForms.CommandButton, p_btnMoveToR As MSForms.CommandButton)
    Dim i As Long
    For i = 0 To p_ListBoxL.ListCount - 1
        If p_ListBoxL.Selected(i) = True Then
            p_ListBoxR.AddItem p_ListBoxL.List(i), 0
            p_ListBoxL.RemoveItem i
            p_btnMoveToL.Enabled = True
            p_btnMoveToR.Enabled = False
            setSearchItemButtonState
            Exit For
        End If
    Next
    setSearchItemButtonState
End Sub

Private Sub cmdbtnFunctionSearch_Click()
    ListBoxFunctionAll.Clear
    readConfigSheetToListBox
    readFlowSheetToListBox
    readInstancesSheetToListBox
    
    ListBoxFunctionAll.Enabled = True
    ListBoxFunctionSelected.Enabled = True
    OptionButtonAddedtoSelectinstance.Enabled = True
    OptionButtonReplaceInstance.Enabled = True
    OptionButtonNewInstance.Enabled = True
    OptionButtonAddedtoFlow.Enabled = True
    OptionButtonReplacetheSelectFlow.Enabled = True
    OptionButtonNewFlow.Enabled = True
    OptionButtonAddedtoSelectinstance.Value = True
    OptionButtonAddedtoFlow.Value = True
    ListBoxFlowAll.Enabled = True
    ListBoxFlowSource.Enabled = True
    ListBoxFlowTarget.Enabled = True
    ListBoxInstanceAll.Enabled = True
    ListBoxInstanceSource.Enabled = True
    ListBoxInstanceTarget.Enabled = True
    
    checkCoEnable
    
    SetupMyMultiColumn ListBoxFunctionAll
    ListBoxFunctionSelected.ColumnWidths = ListBoxFunctionAll.ColumnWidths
    SetupMyMultiColumn ListBoxFlowAll
    ListBoxFlowSource.ColumnWidths = ListBoxFlowAll.ColumnWidths
    ListBoxFlowTarget.ColumnWidths = ListBoxFlowAll.ColumnWidths
    SetupMyMultiColumn ListBoxInstanceAll
    ListBoxInstanceSource.ColumnWidths = ListBoxInstanceAll.ColumnWidths
    ListBoxInstanceSource.ColumnWidths = ListBoxInstanceAll.ColumnWidths
End Sub

Private Sub checkCoEnable()

    checkCoEnableListBox ListBoxFunctionAll, cmdbtnFunctionMoveR
    checkCoEnableListBox ListBoxFunctionSelected, cmdbtnFunctionMoveL
    
    If OptionButtonNewInstance.Enabled = True And _
        OptionButtonNewInstance.Value = True Then
        ListBoxInstanceTarget.Enabled = False
        cmdbtnInstanceMoveTR.Enabled = False
        cmdbtnInstanceMoveTL.Enabled = False
    End If
    
    If OptionButtonNewFlow.Enabled = True And _
        OptionButtonNewFlow.Value = True Then
        ListBoxFlowTarget.Enabled = False
        cmdbtnFlowMoveTR.Enabled = False
        cmdbtnFlowMoveTL.Enabled = False
    End If
    
    checkCoEnableListBox ListBoxFlowAll, cmdbtnFlowMoveSR
    checkCoEnableListBox ListBoxFlowSource, cmdbtnFlowMoveSL
    checkCoEnableListBox ListBoxFlowAll, cmdbtnFlowMoveTR
    checkCoEnableListBox ListBoxFlowTarget, cmdbtnFlowMoveTL
    
    checkCoEnableListBox ListBoxInstanceAll, cmdbtnInstanceMoveSR
    checkCoEnableListBox ListBoxInstanceSource, cmdbtnInstanceMoveSL
    checkCoEnableListBox ListBoxInstanceAll, cmdbtnInstanceMoveTR
    checkCoEnableListBox ListBoxInstanceTarget, cmdbtnInstanceMoveTL
    
    checkCoEnableListBox ListBoxTestItemAll, cmdbtnTestItemMoveSR
    checkCoEnableListBox ListBoxTestItemAll, cmdbtnTestItemMoveSRAll
    checkCoEnableListBox ListBoxInstanceTarget, cmdbtnTestItemMoveSL
    checkCoEnableListBox ListBoxInstanceTarget, cmdbtnTestItemMoveSLAll
    
    If ListBoxInstanceTarget.Enabled = True And ListBoxInstanceTarget.ListCount > 0 Then
        cmdbtnNext.Enabled = True
    Else
        cmdbtnNext.Enabled = False
    End If

End Sub

Private Sub checkCoEnableListBox(p_ListBox As MSForms.ListBox, p_btnMove As MSForms.CommandButton)
    If p_ListBox.Enabled = True And p_ListBox.ListCount > 0 Then
        p_btnMove.Enabled = True
    Else
        p_btnMove.Enabled = False
    End If
End Sub

Private Sub cmdbtnInstanceMoveSL_Click()
    SingalmoveToLeft ListBoxInstanceAll, ListBoxInstanceSource, cmdbtnInstanceMoveSL, cmdbtnInstanceMoveSR
End Sub

Private Sub cmdbtnInstanceMoveSR_Click()
    SingalmoveToRight ListBoxInstanceAll, ListBoxInstanceSource, cmdbtnInstanceMoveSL, cmdbtnInstanceMoveSR
End Sub

Private Sub cmdbtnInstanceMoveTL_Click()
    SingalmoveToLeft ListBoxInstanceAll, ListBoxInstanceTarget, cmdbtnInstanceMoveTL, cmdbtnInstanceMoveTR
End Sub

Private Sub cmdbtnInstanceMoveTR_Click()
    SingalmoveToRight ListBoxInstanceAll, ListBoxInstanceTarget, cmdbtnInstanceMoveTL, cmdbtnInstanceMoveTR
End Sub

Private Sub cmdBtnLoadConfig_Click()
    If findSheetByName(CON_CONFIGSHEETNAME) = True Then
        If MsgBox("Do you want to read another Config to replace current config?", vbOKCancel, "Question") = vbOK Then
            If LoadConfigSheet = True Then
                cmdbtnFunctionSearch.Enabled = True
                MsgBox "read Config is done.", vbOKOnly, "Inforamtion"
            End If
        End If
    Else
        If LoadConfigSheet = True Then
            cmdbtnFunctionSearch.Enabled = True
            MsgBox "read Config is done.", vbOKOnly, "Inforamtion"
        End If
    End If
End Sub

Private Function LoadConfigSheet() As Boolean
    Dim fname As String
    Dim SourceWorkbook As Object
    Dim Configworkbook As Object
    Dim configWorkSheet As Worksheet
    fname = Application.GetOpenFilename("Microsoft Excel(*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm")
    If fname = "False" Then
        LoadConfigSheet = False
    Else
        If findSheetByName(CON_CONFIGSHEETNAME) = False Then
            ThisWorkbook.Worksheets.Add().Name = CON_CONFIGSHEETNAME
        End If
        
        Set configWorkSheet = ThisWorkbook.Worksheets(CON_CONFIGSHEETNAME)
        configWorkSheet.Cells.Clear

        Set SourceWorkbook = ThisWorkbook
        
        Application.ScreenUpdating = False
        Application.ShowWindowsInTaskbar = False
        
        Set Configworkbook = Workbooks.Open(fname, 0, True)
        
        Configworkbook.Sheets(CON_CONFIGSHEETNAME).Range("A1:XFD65535").Copy configWorkSheet.Range("A1")
        Configworkbook.Close savechanges:=False
        
        Application.ShowWindowsInTaskbar = True
        Application.ScreenUpdating = True
        
        LoadConfigSheet = True
    End If
End Function

Private Sub createSheet(p_strSheetName As String)
    ThisWorkbook.Worksheets.Add().Name = p_strSheetName
End Sub

Private Function findSheetByName(p_strSheetName As String) As Boolean
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = p_strSheetName Then
            findSheetByName = True
            Exit Function
        End If
    Next i
    findSheetByName = False
End Function

Private Function findSheetByKey(p_strSheetKey As String) As Boolean
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Cells(1, 1) = p_strSheetKey Then
            findSheetByKey = True
            Exit Function
        End If
    Next i
    findSheetByKey = False
End Function

Private Function getSheetIndex(p_strSheetName As String) As Integer
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = p_strSheetName Then
            getSheetIndex = i
            Exit Function
        End If
    Next i
    getSheetIndex = -1
End Function

Private Sub readFlowSheetToListBox()
    Dim i As Long
    Dim l_strListFlow() As String
    ListBoxFlowAll.Clear
    l_strListFlow = getSheetNameByKey(CON_FLOW, 1, 1)
    For i = 0 To UBound(l_strListFlow) - 1
        ListBoxFlowAll.AddItem l_strListFlow(i)
    Next
End Sub

Private Sub readInstancesSheetToListBox()
    Dim i As Long
    Dim l_strListFlow() As String
    ListBoxInstanceAll.Clear
    l_strListFlow = getSheetNameByKey(CON_INSTANCE, 1, 1)
    For i = 0 To UBound(l_strListFlow) - 1
        ListBoxInstanceAll.AddItem l_strListFlow(i)
    Next
End Sub

Private Sub readConfigSheetToListBox()
    Dim l_shtConfig As Worksheet
    Dim l_iConfigSheetIndex As Integer
    Dim l_strFunctionName As String
    
    l_iConfigSheetIndex = getSheetIndex(CON_CONFIGSHEETNAME)
    
    Set l_shtConfig = ThisWorkbook.Worksheets(l_iConfigSheetIndex)
    
    Dim i As Long
    Dim l_blnBlank As Boolean
    l_blnBlank = False
    m_lUsedRangeCount = getMaxRow(l_shtConfig)
    For i = 3 To m_lUsedRangeCount
        l_strFunctionName = l_shtConfig.Cells(i, 2)
        If VBA.Trim(l_strFunctionName) <> "" Then
            l_blnBlank = False
            ListBoxFunctionAll.AddItem VBA.Trim(l_strFunctionName)
        Else
            If l_blnBlank = True Then
                Exit Sub
            Else
                l_blnBlank = True
            End If
        End If
    Next
End Sub

Private Function getSheetNameByKey(p_strKey As String, row As Integer, col As Integer) As String()
    Dim i As Long
    Dim l_iSheetCount As Integer
    Dim l_iCount As Integer
    Dim l_strListResult() As String
    l_iSheetCount = 0
    l_iCount = 0
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Len(ActiveWorkbook.Sheets(i).Cells(row, col).Text) >= Len(p_strKey) Then
            If p_strKey = VBA.Mid(VBA.Trim(ActiveWorkbook.Sheets(i).Cells(row, col).Text), 1, Len(p_strKey)) Then
                l_iSheetCount = l_iSheetCount + 1
            End If
        End If
    Next
    
    ReDim l_strListResult(l_iSheetCount)
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Len(ActiveWorkbook.Sheets(i).Cells(row, col).Text) >= Len(p_strKey) Then
            If p_strKey = VBA.Mid(VBA.Trim(ActiveWorkbook.Sheets(i).Cells(row, col).Text), 1, Len(p_strKey)) Then
                l_strListResult(l_iCount) = ActiveWorkbook.Sheets(i).Name
                l_iCount = l_iCount + 1
            End If
        End If
    Next
    getSheetNameByKey = l_strListResult
End Function

Private Sub disableAll()
    cmdbtnFunctionSearch.Enabled = False
    
    ListBoxFunctionAll.Enabled = False
    ListBoxFunctionSelected.Enabled = False
    cmdbtnFunctionMoveR.Enabled = False
    cmdbtnFunctionMoveL.Enabled = False

    OptionButtonAddedtoSelectinstance.Enabled = False
    OptionButtonReplaceInstance.Enabled = False
    OptionButtonNewInstance.Enabled = False
    TextBoxInstanceName.Enabled = False
    OptionButtonAddedtoFlow.Enabled = False
    OptionButtonReplacetheSelectFlow.Enabled = False
    OptionButtonNewFlow.Enabled = False
    TextBoxFlowName.Enabled = False
    
    ListBoxFlowAll.Enabled = False
    ListBoxFlowSource.Enabled = False
    cmdbtnFlowMoveSR.Enabled = False
    cmdbtnFlowMoveSL.Enabled = False
    ListBoxFlowTarget.Enabled = False
    cmdbtnFlowMoveTR.Enabled = False
    cmdbtnFlowMoveTL.Enabled = False
    
    ListBoxInstanceAll.Enabled = False
    ListBoxInstanceSource.Enabled = False
    cmdbtnInstanceMoveSR.Enabled = False
    cmdbtnInstanceMoveSL.Enabled = False
    ListBoxInstanceTarget.Enabled = False
    cmdbtnInstanceMoveTR.Enabled = False
    cmdbtnInstanceMoveTL.Enabled = False

    cmdbtnItemSearch.Enabled = False
    ListBoxTestItemAll.Enabled = False
    cmdbtnTestItemMoveSR.Enabled = False
    cmdbtnTestItemMoveSL.Enabled = False
    cmdbtnTestItemMoveSRAll.Enabled = False
    cmdbtnTestItemMoveSLAll.Enabled = False
    ListBoxTestItemTarget.Enabled = False
    
    cmdbtnNext.Enabled = False
End Sub

Private Sub clearAll()

    ListBoxFunctionAll.Clear
    ListBoxFunctionSelected.Clear

    OptionButtonAddedtoSelectinstance.Value = True
    OptionButtonAddedtoFlow.Value = True
    TextBoxInstanceName.Text = ""
    TextBoxFlowName.Text = ""

    
    ListBoxFlowAll.Clear
    ListBoxFlowSource.Clear
    ListBoxFlowTarget.Clear
    
    ListBoxInstanceAll.Clear
    ListBoxInstanceSource.Clear
    ListBoxInstanceTarget.Clear

    ListBoxFlowTarget.Clear
    ListBoxInstanceAll.Clear
    ListBoxInstanceSource.Clear
    ListBoxInstanceTarget.Clear
    ListBoxTestItemAll.Clear
    
    ListBoxTestItemAll.Clear
    ListBoxTestItemTarget.Clear

End Sub

Private Function getMaxRow(p_sheet As Worksheet) As Long
    Dim i As Long
    For i = 5 To 65000
        If p_sheet.Cells(i, 1) = "" And p_sheet.Cells(i, 2) = "" And p_sheet.Cells(i, 3) = "" And p_sheet.Cells(i, 4) = "" And _
            p_sheet.Cells(i, 5) = "" And p_sheet.Cells(i, 6) = "" And p_sheet.Cells(i, 7) = "" And p_sheet.Cells(i, 8) = "" And _
            p_sheet.Cells(i, 9) = "" And p_sheet.Cells(i, 10) = "" Then
            getMaxRow = i
            Exit Function
        End If
    Next
End Function

