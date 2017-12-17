Attribute VB_Name = "Common_Facilities"
Option Explicit
Option Base 1

'======================================================================================================
Sub Sub_ListActiveXControlOnActiveSheet()
    Dim obj As Object
    Dim sStr As String
    
    For Each obj In ActiveSheet.DrawingObjects
        sStr = sStr & vbCr & obj.Name
    Next
     
    Set obj = Nothing
    
    MsgBox sStr
End Sub

Sub sub_ExportModulesSourceCodeToFolder()
    Dim sFolder As String
    Dim sMsg As String
    Dim i As Integer
    Dim iCnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set vbProj = ThisWorkbook.VBProject
    
    iCnt = vbProj.VBComponents.Count
    
    fGetFSO
        
    For i = 1 To 1
        If i = 1 Then
            sFolder = ThisWorkbook.Path & "\" & "Source_Code"
        Else
        End If
        
        sMsg = sMsg & vbCr & vbCr & sFolder
        
        If Not gFSO.FolderExists(sFolder) Then gFSO.CreateFolder (sFolder)
        
        'call fCheckPath(sfolder, true)
        fDeleteAllFilesInFolder (sFolder)
        
        iCnt = 0
        For Each vbComp In vbProj.VBComponents
            If UCase(vbComp.Name) Like "SHEET*" Then GoTo Next_mod
            If vbComp.Type = 1 Or vbComp.Type = 3 Or vbComp.Type = 100 Then
                vbComp.Export sFolder & "\" & vbComp.Name & ".bas"
            End If
            
Next_mod:
        Next
    Next
    
    MsgBox "Done"
End Sub

Sub sub_ListAllFunctionsOfThisWorkbook()
    Dim shtOutput As Worksheet
    If Not fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtOutput) Then Exit Sub
    
    Dim arrModules()
    Dim arrFunctions()
    
    arrModules = fGetListAllModulesOfThisWorkbook()
    arrFunctions = fGetListAllSubFunctionsInThisWorkbook(arrModules)
    
    Call fWriteArray2Sheet(shtOutput, arrFunctions)
    
    Erase arrModules: Erase arrFunctions
    
    shtOutput.Cells(1, 1) = "Type"
    shtOutput.Cells(1, 2) = "Modules"
    shtOutput.Cells(1, 3) = "Functions"
    
    Call fAutoFilterAutoFitSheet(shtOutput)
    Call fFreezeSheet(shtOutput)
    Call fSortDataInSheetSortSheetData(shtOutput, Array(3))
    
    Set shtOutput = Nothing
End Sub

Sub Sub_ToHomeSheet()
    If shtMenu.Visible = xlSheetVisible Then
        shtMenu.Activate
    Else
        ThisWorkbook.Worksheets(1).Activate
    End If
End Sub

Sub sub_ResetOnError_Initialize()
    Err.Clear
    
    'On Error GoTo err_exit
    
    gsEnv = fGetEnvFromSysConf
    
    Call fEnableExcelOptionsAll
    Call sub_RemoveAllCommandBars
    
   ' Call ThisWorkbook.fRefreshGetAllCommandbarsList
    
    Call ThisWorkbook.sub_WorkBookInitialization
    Call fSetIntialValueForShtMenuInitialize
err_exit:
    Err.Clear
    ThisWorkbook.CheckCompatibility = False
    End
End Sub
Function fGetEnvFromSysConf() As String
    gsEnv = fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE", False)
    fGetEnvFromSysConf = gsEnv
End Function

Sub sub_SwitchDevProdMode()
    gsEnv = fGetEnvFromSysConf
    
    If gsEnv = "DEV" Then
        gsEnv = "PROD"
    ElseIf gsEnv = "PROD" Then
        gsEnv = "DEV"
    End If
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE" _
                                    , gsEnv, False)
    
    shtMenu.Activate
    Range("A1").Select
End Sub

Function fSetDEVUATPRODNotificationInSheetMenu()
    Const sDevNotifi = "This is DEV mode, please switch to PROD vresion by click the button above ""Switch Dev/Prod Mode"""
    
    Dim sNotifi As String
    Dim iColor As Long
    Dim iFontSize As Long
    Dim bBold As Boolean
    
    If gsEnv = "DEV" Then
        sNotifi = sDevNotifi
        
        iColor = RGB(0, 0, 255)
        iFontSize = 20
        bBold = True
    ElseIf gsEnv = "PROD" Then
        sNotifi = ""
        
        iColor = RGB(0, 0, 0)
        iFontSize = 10
        bBold = False
    Else
    End If
    
    shtMenu.Range("A1").Value = sNotifi
    shtMenu.Range("A1").Font.Size = iFontSize
    shtMenu.Range("A1").Font.Color = iColor
    shtMenu.Range("A1").Font.Bold = bBold
End Function

'*************************************************************************

Function fGetListAllModulesOfThisWorkbook() As Variant
    Dim arrOut()
    Dim iCnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set vbProj = ThisWorkbook.VBProject
    
    iCnt = vbProj.VBComponents.Count
    ReDim arrOut(1 To iCnt, 3)
    
    iCnt = 0
    For Each vbComp In vbProj.VBComponents
        iCnt = iCnt + 1
        arrOut(iCnt, 1) = "Modules"
        arrOut(iCnt, 2) = fVBEComponentTypeToString(vbComp.Type)
        arrOut(iCnt, 3) = vbComp.Name
    Next
    
    fGetListAllModulesOfThisWorkbook = arrOut
    Erase arrOut
End Function

Function fVBEComponentTypeToString(aType As VBIDE.vbext_ComponentType) As String
    Dim sOut As String
    
    Select Case aType
        Case VBIDE.vbext_ct_ActiveXDesigner
            sOut = "ActiveX Designer"
        Case VBIDE.vbext_ct_ClassModule
            sOut = "Class"
        Case VBIDE.vbext_ct_StdModule
            sOut = "Module"
        Case VBIDE.vbext_ct_Document
            sOut = "Document"
        Case VBIDE.vbext_ct_MSForm
            sOut = "User Form"
        Case Else
            sOut = "Unknown type: " & CStr(aType)
    End Select
    
    fVBEComponentTypeToString = sOut
End Function

Function fGetListAllSubFunctionsInThisWorkbook(arrModules()) As Variant
    Dim arrOut()
    Dim i As Long
    Dim iCnt As Long
    Dim sMod As String
    Dim lineNo As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim procKind As VBIDE.vbext_ProcKind
    Dim funcName As String
    
    Set vbProj = ThisWorkbook.VBProject
    
    iCnt = 0
    ReDim arrOut(1 To 10000, 4)
    
    For i = LBound(arrModules, 1) To UBound(arrModules, 1)
        sMod = arrModules(i, 3)
        
        Set vbComp = vbProj.VBComponents(sMod)
        Set codeMod = vbComp.CodeModule
        
        lineNo = codeMod.CountOfDeclarationLines + 1
        
        Do Until lineNo >= codeMod.CountOfLines + 1
            funcName = codeMod.ProcOfLine(lineNo, procKind)
            
            If Not UCase(funcName) Like "CB*_CLICK" Then
                iCnt = iCnt + 1
                arrOut(iCnt, 1) = "Functions"
                arrOut(iCnt, 2) = sMod
                arrOut(iCnt, 3) = funcName
                arrOut(iCnt, 4) = ProcKindString(procKind)
            End If
            
            lineNo = codeMod.ProcStartLine(funcName, procKind) + codeMod.ProcCountLines(funcName, procKind) + 1
        Loop
    Next
    fGetListAllSubFunctionsInThisWorkbook = arrOut
    Erase arrOut
End Function

Function ProcKindString(procKind As VBIDE.vbext_ProcKind) As String
    Dim sOut As String
    
    Select Case procKind
        Case VBIDE.vbext_pk_Get
            sOut = "Property Get"
        Case VBIDE.vbext_pk_Let
            sOut = "Property Let"
        Case VBIDE.vbext_pk_Proc
            sOut = "Sub/Function"
        Case VBIDE.vbext_pk_Set
            sOut = "Property Set"
        Case Else
            sOut = "Unknown type: " & CStr(procKind)
    End Select
    ProcKindString = sOut
End Function

Function fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtTmp As Worksheet, Optional wb As Workbook) As Boolean
    Dim sTmp As String
    Dim response As VbMsgBoxResult
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    sTmp = "tmpOutput"
    
    If SheetExists(sTmp) Then
        wb.Worksheets(sTmp).Activate
        
        response = MsgBox("There is an existing sheet " & sTmp & ", to delete it, please press yes" _
                    & vbCr, vbCritical + vbYesNoCancel)
        If response = vbNo Then
            Set shtTmp = wb.Worksheets(sTmp)
        ElseIf response = vbYes Then    'vbYes
            Call fDeleteSheet(sTmp)
            Set shtTmp = fAddNewSheet(sTmp)
        Else
            fGetTmpSheetInWorkbookWhenNotExistsCreateIt = False
            Exit Function
        End If
    Else
        Set shtTmp = fAddNewSheet(sTmp)
    End If
    
    fGetTmpSheetInWorkbookWhenNotExistsCreateIt = True
End Function

Function fShtSysConf_SheetChange_DevProdChange(Target As Range)
    Dim rgAimed As Range
    Dim rgIntersect As Range
    
    Set rgAimed = fGetRangeFromExternalAddress(fGetSpecifiedConfigCellAddress(shtSysConf, "[Facility For Testing]", "Value" _
                        , "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE"))
    Set rgIntersect = Intersect(Target, rgAimed)
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then fErr "Please select only one cell."
        
        gsEnv = rgIntersect.Value
        
        Call fRemoveAllCommandbarsByConfig
        Call ThisWorkbook.sub_WorkBookInitialization
        Call fSetIntialValueForShtMenuInitialize
        Call fSetDEVUATPRODNotificationInSheetMenu
    End If
    
    Set rgAimed = Nothing
    Set rgIntersect = Nothing
End Function

Sub sub_GenAlpabetList()
    Dim maxNum
    Dim lMax As Long
    Dim sMaxcol As String
    Dim arrList()
    
    If Not fPromptToOverWrite() Then Exit Sub
    
    maxNum = InputBox("How many letters to you want to generate? (either number or letter is ok, e.g., 20 or AF)", "Max Number letter")
    
    If fZero(maxNum) Then Exit Sub
    
    maxNum = Trim(maxNum)
    
    On Error Resume Next
    lMax = CLng(maxNum)
    sMaxcol = CStr(maxNum)
    Err.Clear
    
    If lMax > 0 Then
    ElseIf Len(sMaxcol) > 0 Then
        lMax = fLetter2Num(sMaxcol)
    End If
    
    If lMax <= 0 Or lMax > Columns.Count Then
        fMsgBox "the number you input is too small or too large, which should be with 1 - " & Columns.CountLarge
        Exit Sub
    End If
    
    Dim i As Long
    ReDim arrList(1 To lMax, 1)
    For i = 1 To lMax
        arrList(i, 1) = fNum2Letter(i)
    Next
    
    ActiveCell.Resize(UBound(arrList, 1), 1).Value = arrList
    Erase arrList
End Sub

Sub sub_GenNumberList()
    Dim maxNum
    Dim lMax As Long
    Dim sMaxcol As String
    Dim arrList()
    
    If Not fPromptToOverWrite() Then Exit Sub
    
    maxNum = InputBox("How many letters to you want to generate? ( e.g., 20 , 100)", "Max Number")
    If fZero(maxNum) Then Exit Sub
    
    maxNum = Trim(maxNum)
    
    On Error Resume Next
    lMax = CLng(maxNum)
    Err.Clear

    If lMax <= 0 Then
        fMsgBox "the number you input is too small or too large, which should be with 1 - " & Columns.CountLarge
        Exit Sub
    End If
    
    Dim i As Long
    ReDim arrList(1 To lMax, 1)
    For i = 1 To lMax
        arrList(i, 1) = i
    Next
    
    ActiveCell.Resize(UBound(arrList, 1), 1).Value = arrList
    Erase arrList

End Sub

Function fPromptToOverWrite() As Boolean
    fPromptToOverWrite = fPromptToConfirmToContinue("Data will be write to the current cell:" _
                & Replace(ActiveCell.Address, "$", "") & vbCr & "are you sure to continue?")
End Function
Function fPromptToConfirmToContinue(asAskMsg As String _
            , Optional aBBbMsgboxStyle As VbMsgBoxStyle = vbYesNoCancel + vbCritical + vbDefaultButton3 _
            , Optional bDoubleConfirm As Boolean = False) As Boolean
    fPromptToConfirmToContinue = False
    
    Dim response As VbMsgBoxResult
    response = MsgBox(prompt:=asAskMsg, Buttons:=aBBbMsgboxStyle)
    
    If response <> vbYes Then Exit Function
    
    If bDoubleConfirm Then
        response = MsgBox(prompt:="Are you sure to continue?", Buttons:=vbYesNoCancel + vbCritical + vbDefaultButton3)
        If response <> vbYes Then Exit Function
    End If
    
    fPromptToConfirmToContinue = True
End Function

'Sub AddFaceIDs()
'    Dim GName As String
'    Dim I As Integer, J As Single
'
'    For I = 6 To 1 Step -1 'Display from bottom to top
'        GName = "Group" & 600 * (I - 1) + 1 & "_" & 600 * I
'        On Error GoTo Endline
'        With Application.CommandBars.Add(GName)
'            .Visible = True
'            With .Controls
'                For J = 600 * (I - 1) + 1 To 600 * I
'                On Error Resume Next
'                With .Add(msoControlButton)
'                .FaceId = J
'                .Caption = J
'                End With
'                Next
'            End With
'        End With
'Endline:
'        With CommandBars(GName)
'            .Visible = True
'            .Width = 720 'contains 30�20 icons
'            .Left = 50 + (6 - I) * 20
'            .Top = 90 + (6 - I) * 20
'        End With
'    Next I
'End Sub
Sub Sub_FilterByActiveCell()
    Dim lMaxCol As Long
    lMaxCol = ActiveSheet.Cells(1, 1).End(xlToRight).Column
    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(ActiveSheet)

    If ActiveSheet.AutoFilterMode Then  'auto filter
        ActiveSheet.AutoFilter.ShowAllData
    Else
        fGetRangeByStartEndPos(ActiveSheet, 1, 1, 1, lMaxCol).AutoFilter
    End If
    
    Dim aActiveCellValue
    Dim lColToFilter As Long
    aActiveCellValue = ActiveCell.Value
    lColToFilter = ActiveCell.Column
    
    fGetRangeByStartEndPos(ActiveSheet, 1, 1, lMaxRow, lMaxCol).AutoFilter _
                Field:=lColToFilter _
                , Criteria1:="=*" & aActiveCellValue & "*" _
                , Operator:=xlAnd
End Sub

Sub Sub_RemoveFilterForAcitveSheet()
    Call fRemoveFilterForSheet(ActiveSheet)
End Sub

Sub sub_SortBySelectColumn()
    Dim sSelectContent As String
    Dim lSelectCol As Long
    sSelectContent = ActiveCell.Value
    lSelectCol = ActiveCell.Column
    
    Call Sub_RemoveFilterForAcitveSheet
    Call fSortDataInSheetSortSheetData(ActiveSheet, Array(ActiveCell.Column))
    
    Dim rgFound As Range
    Set rgFound = fFindInWorksheet(ActiveSheet.Columns(lSelectCol), sSelectContent, True, True)
    
    If Not rgFound Is Nothing Then rgFound.Select
    Set rgFound = Nothing
End Sub
