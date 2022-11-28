Attribute VB_Name = "mTDHelper"
Option Explicit
'mTDHelper: Helper module for cTaskDialog.cls
'Must be included with the class.
#Const CTASKDIALOG_DEFINED = 1
#If (VBA7 = 0) Then 'Adds LongPtr variable support to VB6
Public Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
#Else
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
#End If
Public Function TaskDialogCallbackProc(ByVal hwnd As LongPtr, ByVal uNotification As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal lpRefData As LongPtr) As LongPtr
Dim cTD As cTaskDialog
CopyMemory cTD, lpRefData, LenB(lpRefData)
TaskDialogCallbackProc = cTD.zz_ProcessCallback(hwnd, uNotification, wParam, lParam)
ZeroMemory cTD, LenB(lpRefData)
End Function
Public Function TaskDialogEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
Dim cTD As cTaskDialog
CopyMemory cTD, lParam, LenB(lParam)
TaskDialogEnumChildProc = cTD.zz_ProcessEnumCallback(hwnd)
ZeroMemory cTD, LenB(lParam)
End Function
Public Function TaskDialogSubclassProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
Dim cTD As cTaskDialog
CopyMemory cTD, dwRefData, LenB(dwRefData)
TaskDialogSubclassProc = cTD.zz_ProcessSubclass(hwnd, uMsg, wParam, lParam, uIdSubclass)
ZeroMemory cTD, LenB(dwRefData)
End Function