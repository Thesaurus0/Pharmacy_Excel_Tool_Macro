Attribute VB_Name = "ģ��1"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function _
    CloseClipboard& Lib "User32" ()
    Private Declare PtrSafe Function _
    OpenClipboard& Lib "User32" (ByVal hwnd&)
    Private Declare PtrSafe Function _
    EmptyClipboard& Lib "User32" ()
    Private Declare PtrSafe Function _
    GetClipboardData& Lib "User32" (ByVal wFormat&)
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
    CloseClipboard& Lib "User32" ()
    Private Declare Function _
    OpenClipboard& Lib "User32" (ByVal hwnd&)
    Private Declare Function _
    EmptyClipboard& Lib "User32" ()
    Private Declare Function _
    GetClipboardData& Lib "User32" (ByVal wFormat&)
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
    Dim hwnd&, size&, Ptr&
    If OpenClipboard(0&) Then
        ' Get memory handle to the data
        hwnd = GetClipboardData(Format)
        ' Get size of this memory block
        If hwnd Then size = GlobalSize(hwnd)
            ' Get pointer to the locked memory
        If size Then Ptr = GlobalLock(hwnd)
        
        If Ptr Then
            ' Resize the byte array to hold the data
            ReDim abData(0 To size - 1) As Byte
            ' Copy from the pointer into the array
            CopyMem abData(0), ByVal Ptr, size
            ' Unlock the memory
            Call GlobalUnlock(hwnd)
            GetData = True
        End If
        EmptyClipboard
        CloseClipboard
        DoEvents
    End If
End Function

