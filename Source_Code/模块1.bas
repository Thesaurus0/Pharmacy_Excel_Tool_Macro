Attribute VB_Name = "Ä£¿é1"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function _
    CloseClipboard& Lib "user32" ()
    Private Declare PtrSafe Function _
    OpenClipboard& Lib "user32" (ByVal hWnd&)
    Private Declare PtrSafe Function _
    EmptyClipboard& Lib "user32" ()
    Private Declare PtrSafe Function _
    GetClipboardData& Lib "user32" (ByVal wFormat&)
    Private Declare PtrSafe Function _
    GlobalSize& Lib "kernel32" (ByVal hMem&)
    Private Declare PtrSafe Function _
    GlobalLock& Lib "kernel32" (ByVal hMem&)
    Private Declare PtrSafe Function _
    GlobalUnlock& Lib "kernel32" (ByVal hMem&)
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias _
    "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length&)
#Else
    Private Declare Function _
    CloseClipboard& Lib "user32" ()
    Private Declare Function _
    OpenClipboard& Lib "user32" (ByVal hWnd&)
    Private Declare Function _
    EmptyClipboard& Lib "user32" ()
    Private Declare Function _
    GetClipboardData& Lib "user32" (ByVal wFormat&)
    Private Declare Function _
    GlobalSize& Lib "kernel32" (ByVal hMem&)
    Private Declare Function _
    GlobalLock& Lib "kernel32" (ByVal hMem&)
    Private Declare Function _
    GlobalUnlock& Lib "kernel32" (ByVal hMem&)
    Private Declare Sub CopyMem Lib "kernel32" Alias _
    "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length&)
#End If

Function GetData(ByVal Format&, abData() As Byte) As Boolean
    Dim hWnd&, Size&, Ptr&
    If OpenClipboard(0&) Then
        ' Get memory handle to the data
        hWnd = GetClipboardData(Format)
        ' Get size of this memory block
        If hWnd Then Size = GlobalSize(hWnd)
            ' Get pointer to the locked memory
        If Size Then Ptr = GlobalLock(hWnd)
        
        If Ptr Then
            ' Resize the byte array to hold the data
            ReDim abData(0 To Size - 1) As Byte
            ' Copy from the pointer into the array
            CopyMem abData(0), ByVal Ptr, Size
            ' Unlock the memory
            Call GlobalUnlock(hWnd)
            GetData = True
        End If
        EmptyClipboard
        CloseClipboard
        DoEvents
    End If
End Function

