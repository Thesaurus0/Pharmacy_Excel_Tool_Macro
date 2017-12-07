Attribute VB_Name = "Common_Facilities"
Option Explicit
Option Base 1

'======================================================================================================

Sub sub_ExportModulesSourceCodeToFolder()
    Dim sFolder As String
    Dim sMsg As String
    Dim i As Integer
    Dim icnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set vbProj = ThisWorkbook.VBProject
    
    icnt = vbProj.VBComponents.Count
        
    For i = 1 To 1
        If i = 1 Then
            sFolder = ThisWorkbook.Path & "\" & "Source_Code"
        Else
        End If
        
        sMsg = sMsg & vbCr & vbCr & sFolder
        
        'call fCheckPath(sfolder, true)
        fDeleteAllFilesInFolder (sFolder)
        
        icnt = 0
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
    Call fSortDataInSheetSortSheetData(shtOutput, 3)
    
    Set shtOutput = Nothing
End Sub


'*************************************************************************

Function fGetListAllModulesOfThisWorkbook() As Variant
    Dim arrOUt()
    Dim icnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    Set vbProj = ThisWorkbook.VBProject
    
    icnt = vbProj.VBComponents.Count
    ReDim arrOUt(1 To icnt, 3)
    
    icnt = 0
    For Each vbComp In vbProj.VBComponents
        icnt = icnt + 1
        arrOUt(icnt, 1) = "Modules"
        arrOUt(icnt, 2) = fVBEComponentTypeToString(vbComp.Type)
        arrOUt(icnt, 3) = vbComp.Name
    Next
    
    fGetListAllModulesOfThisWorkbook = arrOUt
    Erase arrOUt
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
    Dim arrOUt()
    Dim i As Long
    Dim icnt As Long
    Dim sMod As String
    Dim lineNo As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim procKind As VBIDE.vbext_ProcKind
    Dim funcName As String
    
    Set vbProj = ThisWorkbook.VBProject
    
    icnt = 0
    ReDim arrOUt(1 To 10000, 4)
    
    For i = LBound(arrModules, 1) To UBound(arrModules, 1)
        sMod = arrModules(i, 3)
        
        Set vbComp = vbProj.VBComponents(sMod)
        Set codeMod = vbComp.CodeModule
        
        lineNo = codeMod.CountOfDeclarationLines + 1
        
        Do Until lineNo >= codeMod.CountOfLines + 1
            funcName = codeMod.ProcOfLine(lineNo, procKind)
            
            If Not UCase(funcName) Like "CB*_CLICK" Then
                icnt = icnt + 1
                arrOUt(icnt, 1) = "Functions"
                arrOUt(icnt, 2) = sMod
                arrOUt(icnt, 3) = funcName
                arrOUt(icnt, 4) = ProcKindString(procKind)
            End If
            
            lineNo = codeMod.ProcStartLine(funcName, procKind) + codeMod.ProcCountLines(funcName, procKind) + 1
        Loop
    Next
    fGetListAllSubFunctionsInThisWorkbook = arrOUt
    Erase arrOUt
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
