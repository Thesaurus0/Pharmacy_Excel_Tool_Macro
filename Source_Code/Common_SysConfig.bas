Attribute VB_Name = "Common_SysConfig"
Option Explicit
Option Base 1

Public Const ERROR_NUMBER = 1000
Public Const DELIMITER = "|"
Public gsEnv As String

Public gbNoData As Boolean
Public gbBusinessError As Boolean
Public gbUserCanceled As Boolean

Public gbCheckCompatibility As Boolean
'=======================================
Public gsRptID As String

Public gDictInputFiles As Dictionary
Public gDictTxtFileSpec As Dictionary

'=======================================
Public arrMaster()
Public dictMstColIndex As Dictionary
Public dictMstDisplayName As Dictionary
Public dictMstRawType As Dictionary
Public dictMstCellFormat As Dictionary
Public dictMstDataFormat As Dictionary
'=======================================
Public arrOutput()
Public gsRptType As String
Public gsRptFilePath As String
Public dictRptColIndex As Dictionary
Public dictRptDisplayName As Dictionary
Public dictRptRawType As Dictionary
Public dictRptCellFormat As Dictionary
Public dictRptDataFormat As Dictionary
Public dictRptColWidth As Dictionary
Public dictRptColAttr As Dictionary
'=======================================
Public gFSO As FileSystemObject
Public gRegExp As VBScript_RegExp_55.RegExp


Function fIsDev() As Boolean
    If fZero(gsEnv) Then
        gsEnv = fGetEnvFromSysConf
        Debug.Print "gsenv is blank in fIsDev. re-call fGetEnvFromSysConf " & Now()
    End If
    
    fIsDev = (gsEnv = "DEV")
End Function

Function fSetNoDataAndError(sMsg As String, Optional bStop As Boolean = True)
    gbNoData = True
    
    If bStop Then
        fErr sMsg
    Else
        fMsgBox sMsg
    End If
End Function

Function fSetUserCanceledAndError(sMsg As String, Optional bStop As Boolean = True)
    gbUserCanceled = True
    If bStop Then fErr
End Function

