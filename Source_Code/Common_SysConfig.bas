Attribute VB_Name = "Common_SysConfig"
Option Explicit
Option Base 1

Public Const ERROR_NUMBER = 1000
Public Const DELIMITER = "|"
Public gsEnv As String
'=======================================
Public gsRptID As String

Public gDictInputFiles As Dictionary
'=======================================
Public gFSO As FileSystemObject

