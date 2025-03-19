Attribute VB_Name = "mTDHelper"
Option Explicit
'mTDHelper: Helper module for cTaskDialog.cls
'Must be included with the class.
#If (VBA7 = 0) Then 'Adds LongPtr variable support to VB6
Public Enum LongPtr
    [_]
End Enum
#End If
Public Sub MagicalTDInitFunction()
	'The trick is a GENIUS!
    'He identified the bug in VBA64 that had been causing the crashing.
    'As if by magic, calling this from Class_Initialize resolves the problem.
End Sub
Public Function TaskDialogCallbackProc(ByVal hwnd As LongPtr, ByVal uNotification As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal lpRefData As cTaskDialog) As LongPtr
TaskDialogCallbackProc = lpRefData.zz_ProcessCallback(hwnd, uNotification, wParam, lParam)
End Function
Public Function TaskDialogEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As cTaskDialog) As Long
TaskDialogEnumChildProc = lParam.zz_ProcessEnumCallback(hwnd)
End Function
Public Function TaskDialogSubclassProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As cTaskDialog) As LongPtr
TaskDialogSubclassProc = dwRefData.zz_ProcessSubclass(hwnd, uMsg, wParam, lParam, uIdSubclass)
End Function