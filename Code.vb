
    Select Case sVariable
        Case "CMDBAR"
            
        Case Else
        
    End Select
    
    
Function fCType(aValue, asToType As String, asFormat As String) As Variant
    Dim aOut As Variant
    Dim sDataType As String
    Dim bOrigToAreSame As Boolean
    
    If IsEmpty(aValue) Then fCType = aValue: Exit Function
    
    asToType = UCase(asToType)
    sDataType = UCase(TypeName(aValue))
    
    bOrigToAreSame = False
    Select Case asToType
        Case "STRING", "TEXT"
            If sDataType = "STRING" Then bOrigToAreSame = True
        Case "DATE"
            If aValue = 0 Then fCType = 0: Exit Function
            If sDataType = "DATE" Then bOrigToAreSame = True
        Case "DECIMAL"
            If sDataType = "DECIMAL" Or sDataType = "DOUBLE" Or sDataType = "SINGLE" Or sDataType = "CURRENCY" Then
                bOrigToAreSame = True
            End If
        Case "NUMBER"
            If sDataType = "BYTE" Or sDataType = "INTEGER" Or sDataType = "LONG" Or sDataType = "LONGLONG" Or sDataType = "LONGPRT" Then
                bOrigToAreSame = True
            End If
        Case "STRING_PERCENTAGE"
            
        Case Else
            fErr "wrong param asToType"
    End Select
    
    If bOrigToAreSame Then fCType = aValue: Exit Function
    
    Select Case asToType
        Case "STRING", "TEXT"
            fCType = CStr(aValue)
        Case "DATE"
            Dim dtTmp As Date
            dtTmp = fCdateStr(CStr(aValue), asFormat)
            
            If dtTmp <= 0 Then
                fErr "Wrong date value: " & aValue & ", please check your data, or contact with IT support."
            End If
            fCType = dtTmp
        Case "DECIMAL"
            fCType = CDbl(aValue)
        Case "NUMBER"
            fCType = CLng(aValue)
        Case "STRING_PERCENTAGE"
            fCType = fCPercentage2Dbl(aValue)
        Case Else
            fErr "wrong param asToType"
    End Select
End Function

Function fCdateStr(sDate As String, Optional sFormat As String = "") As Date
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    
    sDate = Trim(sDate)
    If Len(sDate) <= 0 Then Exit Function
    
    Dim bSplit As Boolean
    Dim sDelimiters As String
    Const DATE_DELIMITERS = "-/._."
    
    bSplit = False
    
    Dim i As Integer
    For i = 1 To Len(DATE_DELIMITERS)
        If InStr(sDate, Mid(DATE_DELIMITERS, i, 1)) > 0 Then
            sDelimiter = Mid(DATE_DELIMITERS, i, 1)
            bSplit = True
            Exit For
        End If
    Next
    
    sFormat = Replace(sFormat, ">", "")
    sFormat = Replace(sFormat, "<", "")
    
    If bSplit Then sFormat = Replace(sDate, sDelimiter, "/")
    
    Select Case UCase(sFormat)
        Case "DDMMMYY", "DDMMMYYYY"
            sYear = Mid(sDate, 6)
            sMonth = fConvertMMM2Num(Mid(sDate, 3, 3))
            sDay = Left(sDate, 2)
        Case Else
        
    End Select
End Function
