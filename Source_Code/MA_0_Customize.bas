Attribute VB_Name = "MA_0_Customize"
Option Explicit
Option Base 1

Function fOverWriteGDictVariables_gDictInputfiles()
    Dim sFile As String
    
    sFile = Trim(shtMenu.Range("rngSalesFilePath_GY").Value)
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Input Files]", "File Full Path", "File Tag=GY", sFile)
    Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, "GY", InputFile.FilePath - InputFile.FileTag, sFile)
    
End Function
