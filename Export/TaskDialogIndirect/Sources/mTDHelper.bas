Attribute VB_Name = "mTDHelper"
Option Explicit

Public Function TaskDialogCallbackProc(ByVal hwnd As LongPtr, ByVal uNotification As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal lpRefData As cTaskDialog) As Long: TaskDialogCallbackProc = lpRefData.zz_ProcessCallback(hwnd, uNotification, wParam, lParam): End Function
Public Function TaskDialogEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As cTaskDialog) As Long: TaskDialogEnumChildProc = lParam.zz_ProcessEnumCallback(hwnd): End Function
Public Function TaskDialogSubclassProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As cTaskDialog) As LongPtr: TaskDialogSubclassProc = dwRefData.zz_ProcessSubclass(hwnd, uMsg, wParam, lParam, uIdSubclass): End Function


