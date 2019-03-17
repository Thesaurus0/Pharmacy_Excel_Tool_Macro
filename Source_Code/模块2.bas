Attribute VB_Name = "Ä£¿é2"
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

'Public Function GetClipboard() As String
'    Dim iStrPtr As LongPtr
'    Dim iLen As LongPtr
'    Dim iLock As LongPtr
'    Dim sUniText As String
'    Const CF_UNICODETEXT As Long = 13&
'    OpenClipboard 0&
'    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
'        iStrPtr = GetClipboardData(CF_UNICODETEXT)
'        If iStrPtr Then
'            iLock = GlobalLock(iStrPtr)
'            iLen = GlobalSize(iStrPtr)
'            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
'            'lstrcpy StrPtr(sUniText), iLock
'            GlobalUnlock iStrPtr
'        End If
'        GetClipboard = sUniText
'    End If
'    CloseClipboard
'End Function
Sub testaxxccac()
    ClipBoard_SetData "aæùsdfxxxafxxx"
    Dim a
    'a = GetClipboard
    
    a = GetClipboard
    SetClipboard "asdfasdfdddddddd"
    a = GetClipboard
    'ClipBoard_SetData ""
   ' ClearClipboard
'    Dim a, b() As Byte
'     Call GetData(a, b)
End Sub
Sub asdfasdfddd()
    Dim a As String
    a = "abc"
    
    Dim c$
    Dim d&
    Dim e As LongPtr
    
    Dim b
    b = StrPtr(a)
    
End Sub
