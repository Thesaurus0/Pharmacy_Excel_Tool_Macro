

Function fValidateBlankInArray(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fMsgRaiseErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fMsgRaiseErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call fValidateBlankInArrayForMultipleCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateBlankInArrayForMultipleCols(arrParam, arrKeyCols _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to function fValidateBlankInArrayIndividually
    Const DELI = " " & DELIMITER & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not fZero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & fNum2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
            
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & DELI & Trim(CStr(arrParam(lEachRow, lEachCol)))
        Next
        
        If Not bAllowBlank Then
            If fZero(Replace(sKeyStr, DELI, "")) Then
                'sPos = sPos & "[" & lActualRow & ", " & sColLetterStr & "]"
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fMsgRaiseErr "Keys [" & sKeyStr & "] is blank!" & sPos
            End If
        End If
        
        sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(DELI))
        
        If dict.Exists(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fMsgRaiseErr "Duplicate key [" & sKeyStr & " was found:" & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fValidateBlankInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional bAllowBlank As Boolean = False _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fMsgRaiseErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row/Column: "
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"
         
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = Trim(CStr(arrParam(lEachRow, lKeyCol)))
        
        If Not bAllowBlank Then
            If fZero(sKeyStr) Then
                'sPos = sPos & lActualRow & " / " & sColLetter
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fMsgRaiseErr "Keys [" & sKeyStr & "] is blank!" & sPos
            End If
        End If
        
        If dict.Exists(sKeyStr) Then
            'sPos = sPos & lActualRow & " / " & sColLetter
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fMsgRaiseErr "Duplicate key [" & sKeyStr & "] was found " & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fValidateBlankInArrayIndividually(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional arrColNames)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fMsgRaiseErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fMsgRaiseErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
        Next
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
    End If
End Function


