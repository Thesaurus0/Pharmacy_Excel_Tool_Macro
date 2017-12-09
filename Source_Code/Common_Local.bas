Attribute VB_Name = "Common_Local"
Option Explicit
Option Base 1
 

Function fUpdateGDictInputFile_FileName(asFileTag As String, asFileName As String)
    Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, asFileTag, InputFile.FilePath - InputFile.FileTag, asFileName)
End Function

Function fSetValueBackToSysConf_InputFile_FileName(asFileTag As String, asFileName As String)
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Input Files]", "File Full Path", "File Tag=" & asFileTag, asFileName)
End Function

