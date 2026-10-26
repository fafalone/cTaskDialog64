Attribute VB_Name = "mTDHelper"
Option Explicit
'mTDHelper: Helper module for cTaskDialog.cls
'Must be included with the class.
#If (VBA7 = 0) Then 'Adds LongPtr variable support to VB6
Public Enum LongPtr
    [_]
End Enum
#End If
#If VBA7 Then
Private Declare PtrSafe Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hhk As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#Else
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hhk As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If
Private Const WH_GETMESSAGE As Long = 3
Private Const HC_ACTION As Long = 0
Private Const PM_REMOVE As Long = 1
Private mhTDGetMsgHook As LongPtr
Private mTDTabClients As Collection
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

'Tab-order support for the controls cTaskDialog adds; see cTaskDialog.zz_ProcessGetMsg.
'One thread hook serves every open dialog (nested ones included) and goes away with the last.
Public Sub TaskDialogTabHookAdd(ByVal oTD As cTaskDialog)
If mTDTabClients Is Nothing Then Set mTDTabClients = New Collection
On Error Resume Next 'already registered
mTDTabClients.Add oTD, CStr(ObjPtr(oTD))
On Error GoTo 0
If mhTDGetMsgHook = 0 Then
    mhTDGetMsgHook = SetWindowsHookExW(WH_GETMESSAGE, AddressOf TaskDialogGetMsgProc, 0, GetCurrentThreadId())
End If
End Sub
Public Sub TaskDialogTabHookRemove(ByVal oTD As cTaskDialog)
If mTDTabClients Is Nothing Then Exit Sub
On Error Resume Next 'not registered
mTDTabClients.Remove CStr(ObjPtr(oTD))
On Error GoTo 0
If (mTDTabClients.Count = 0) And (mhTDGetMsgHook <> 0) Then
    UnhookWindowsHookEx mhTDGetMsgHook
    mhTDGetMsgHook = 0
End If
End Sub
Public Function TaskDialogGetMsgProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Dim oTD As cTaskDialog
If (nCode = HC_ACTION) And ((wParam And PM_REMOVE) <> 0) Then
    If Not (mTDTabClients Is Nothing) Then
        For Each oTD In mTDTabClients
            If oTD.zz_ProcessGetMsg(lParam) Then Exit For 'lParam is the MSG*
        Next
    End If
End If
TaskDialogGetMsgProc = CallNextHookEx(mhTDGetMsgHook, nCode, wParam, lParam)
End Function