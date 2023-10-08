VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TaskDialogIndirect Sample Project"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Misc"
      Height          =   1245
      Left            =   90
      TabIndex        =   58
      Top             =   5040
      Width           =   2985
      Begin VB.CommandButton Command44 
         Caption         =   "Auto-close"
         Height          =   300
         Left            =   90
         TabIndex        =   63
         Top             =   585
         Width           =   1350
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Logo Image B"
         Height          =   300
         Left            =   1545
         TabIndex        =   62
         Top             =   585
         Width           =   1350
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Advanced Multi Page"
         Height          =   300
         Left            =   450
         TabIndex        =   61
         Top             =   915
         Width           =   2040
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Logo Image A"
         Height          =   300
         Left            =   1545
         TabIndex        =   60
         Top             =   255
         Width           =   1350
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Dropdown Button"
         Height          =   300
         Left            =   90
         TabIndex        =   59
         Top             =   255
         Width           =   1350
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "With Slider"
      Height          =   675
      Left            =   105
      TabIndex        =   49
      Top             =   4305
      Width           =   2970
      Begin VB.CommandButton Command35 
         Caption         =   "Full Options"
         Height          =   270
         Left            =   1575
         TabIndex        =   53
         Top             =   255
         Width           =   1335
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Basic Slider"
         Height          =   300
         Left            =   135
         TabIndex        =   52
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "With DateTime"
      Height          =   1710
      Left            =   120
      TabIndex        =   36
      Top             =   2520
      Width           =   2970
      Begin VB.CommandButton Command34 
         Caption         =   "Double-check"
         Height          =   285
         Left            =   60
         TabIndex        =   48
         Top             =   945
         Width           =   1380
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Limited Range"
         Height          =   285
         Left            =   1545
         TabIndex        =   47
         Top             =   945
         Width           =   1365
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Triple Controls"
         Height          =   285
         Left            =   1545
         TabIndex        =   46
         Top             =   1320
         Width           =   1365
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Dual controls"
         Height          =   285
         Left            =   60
         TabIndex        =   45
         Top             =   1320
         Width           =   1365
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Date + Time"
         Height          =   285
         Left            =   1545
         TabIndex        =   44
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Checkbox"
         Height          =   285
         Left            =   60
         TabIndex        =   43
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Basic Time"
         Height          =   285
         Left            =   1545
         TabIndex        =   40
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Basic Date"
         Height          =   285
         Left            =   60
         TabIndex        =   37
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "With ComboBox"
      Height          =   945
      Left            =   75
      TabIndex        =   27
      Top             =   1545
      Width           =   2925
      Begin VB.CommandButton Command26 
         Caption         =   "Dual controls B"
         Height          =   285
         Left            =   1515
         TabIndex        =   35
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Dual controls A"
         Height          =   285
         Left            =   75
         TabIndex        =   34
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Dropdown List"
         Height          =   285
         Left            =   1515
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Editable Combo"
         Height          =   285
         Left            =   75
         TabIndex        =   28
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "With Input Box"
      Height          =   1395
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   2910
      Begin VB.CommandButton Command21 
         Caption         =   "Command Links"
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   990
         Width           =   1365
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Minimalist B"
         Height          =   285
         Left            =   45
         TabIndex        =   25
         Top             =   607
         Width           =   1365
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Default Text"
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Password Demo"
         Height          =   285
         Left            =   45
         TabIndex        =   23
         Top             =   990
         Width           =   1365
      End
      Begin VB.CommandButton Command17 
         Caption         =   "As footer"
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   607
         Width           =   1365
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Minimalist A"
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Button Icons"
      Height          =   285
      Left            =   3135
      TabIndex        =   17
      Top             =   780
      Width           =   1485
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Multiple Page"
      Height          =   285
      Left            =   4659
      TabIndex        =   16
      Top             =   765
      Width           =   1485
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Plain TaskDialog"
      Height          =   285
      Left            =   6180
      TabIndex        =   15
      Top             =   765
      Width           =   1485
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Everything"
      Height          =   390
      Left            =   3150
      TabIndex        =   14
      Top             =   1920
      Width           =   1485
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Windows DLL Icon"
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   1455
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   360
      Left            =   6375
      TabIndex        =   12
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5310
      Top             =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Progress Bar"
      Height          =   285
      Left            =   6180
      TabIndex        =   11
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command Links"
      Height          =   285
      Left            =   4659
      TabIndex        =   10
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton Command8 
      Caption         =   "All Text Fields"
      Height          =   285
      Left            =   3141
      TabIndex        =   9
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Custom Icons"
      Height          =   285
      Left            =   4635
      TabIndex        =   8
      Top             =   1455
      Width           =   1485
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hyperlinks"
      Height          =   285
      Left            =   6180
      TabIndex        =   7
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Radio Buttons"
      Height          =   285
      Left            =   4659
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Custom Buttons"
      Height          =   285
      Left            =   3135
      TabIndex        =   5
      Top             =   1095
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Header text"
      Height          =   285
      Left            =   4635
      TabIndex        =   4
      Top             =   1110
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simple Dialog"
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Top             =   120
      Width           =   1485
   End
   Begin VB.CommandButton Command36 
      Caption         =   "InputDbg"
      Height          =   315
      Left            =   705
      TabIndex        =   54
      Top             =   3030
      Width           =   870
   End
   Begin VB.CommandButton Command37 
      Caption         =   "ComboDbg"
      Height          =   315
      Left            =   825
      TabIndex        =   55
      Top             =   3270
      Width           =   885
   End
   Begin VB.CommandButton Command38 
      Caption         =   "DateDbg"
      Height          =   270
      Left            =   705
      TabIndex        =   56
      Top             =   2790
      Width           =   930
   End
   Begin VB.CommandButton Command39 
      Caption         =   "SliderDbg"
      Height          =   300
      Left            =   825
      TabIndex        =   57
      Top             =   3630
      Width           =   855
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3630
      TabIndex        =   65
      Top             =   4365
      Width           =   225
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page:"
      Height          =   195
      Left            =   3195
      TabIndex        =   64
      Top             =   4335
      Width           =   420
   End
   Begin VB.Label Label15 
      Height          =   240
      Left            =   3705
      TabIndex        =   51
      Top             =   4125
      Width           =   1410
   End
   Begin VB.Label Label14 
      Caption         =   "Slider: "
      Height          =   195
      Left            =   3210
      TabIndex        =   50
      Top             =   4125
      Width           =   525
   End
   Begin VB.Label Label13 
      Height          =   240
      Left            =   3945
      TabIndex        =   42
      Top             =   3915
      Width           =   3435
   End
   Begin VB.Label Label12 
      Caption         =   "DT Check:"
      Height          =   240
      Left            =   3195
      TabIndex        =   41
      Top             =   3900
      Width           =   750
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   4005
      TabIndex        =   39
      Top             =   3690
      Width           =   3270
   End
   Begin VB.Label Label10 
      Caption         =   "Date/Time:"
      Height          =   225
      Left            =   3180
      TabIndex        =   38
      Top             =   3690
      Width           =   870
   End
   Begin VB.Label Label9 
      Height          =   225
      Left            =   4275
      TabIndex        =   32
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label8 
      Caption         =   "Combo index:"
      Height          =   255
      Left            =   3195
      TabIndex        =   31
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label Label7 
      Height          =   240
      Left            =   4170
      TabIndex        =   30
      Top             =   3255
      Width           =   3420
   End
   Begin VB.Label Label6 
      Caption         =   "Combo text:"
      Height          =   255
      Left            =   3210
      TabIndex        =   29
      Top             =   3255
      Width           =   930
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   4485
      TabIndex        =   21
      Top             =   3015
      Width           =   3180
   End
   Begin VB.Label Label4 
      Caption         =   "Input box result:"
      Height          =   240
      Left            =   3225
      TabIndex        =   20
      Top             =   3030
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "(checkbox result)"
      Height          =   240
      Left            =   3195
      TabIndex        =   2
      Top             =   2805
      Width           =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "(radio result)"
      Height          =   255
      Left            =   3180
      TabIndex        =   1
      Top             =   2595
      Width           =   3030
   End
   Begin VB.Label Label1 
      Caption         =   "(main result)"
      Height          =   210
      Left            =   3180
      TabIndex        =   0
      Top             =   2385
      Width           =   2985
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'cTaskDialog Samples
'Written by fafalone
'Feel free to use as you wish, with due credit



Private WithEvents TaskDialog1 As cTaskDialog
Attribute TaskDialog1.VB_VarHelpID = -1
Private WithEvents TaskDialog2 As cTaskDialog
Attribute TaskDialog2.VB_VarHelpID = -1
Private WithEvents TaskDialog3 As cTaskDialog
Attribute TaskDialog3.VB_VarHelpID = -1
Private WithEvents TaskDialogPW As cTaskDialog
Attribute TaskDialogPW.VB_VarHelpID = -1
Private WithEvents TaskDialogPW2 As cTaskDialog
Attribute TaskDialogPW2.VB_VarHelpID = -1
Private WithEvents TaskDialogSC As cTaskDialog
Attribute TaskDialogSC.VB_VarHelpID = -1
Private WithEvents TaskDialogAC As cTaskDialog
Attribute TaskDialogAC.VB_VarHelpID = -1
Private WithEvents TaskDialogMPX1 As cTaskDialog
Attribute TaskDialogMPX1.VB_VarHelpID = -1
Private WithEvents TaskDialogMPX2 As cTaskDialog
Attribute TaskDialogMPX2.VB_VarHelpID = -1
Private WithEvents TaskDialogMPX3 As cTaskDialog
Attribute TaskDialogMPX3.VB_VarHelpID = -1

Private bRunProgress As Boolean
Private bRunMarquee As Boolean
Private bRunMarquee2 As Boolean
Private lSecs As Long
Private himlSys As LongPtr
Private bPageExampleEx As Boolean
Private sMPLogin As String

Private sMPName As String

Private Enum ShowWindowTypes
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

#If (VBA7 = 0) Then
Private Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As ShowWindowTypes) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As SysBeeps) As Long
#Else
Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal lpOperation As LongPtr, ByVal lpFile As LongPtr, ByVal lpParameters As LongPtr, ByVal lpDirectory As LongPtr, ByVal nShowCmd As ShowWindowTypes) As LongPtr
Private Declare PtrSafe Function MessageBeep Lib "user32" (ByVal wType As SysBeeps) As Long
#End If
Private Enum SysBeeps
    MB_DEFAULTBEEP = -1    ' the default beep sound
    MB_ERROR = 16         ' for critical errors/problems
    MB_WARNING = 48       ' for conditions that might cause problems in the future
    MB_INFORMATION = 64   ' for informative messages only
    MB_QUESTION = 32      ' (no longer recommended to be used)
        
End Enum
Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Command10_Click()
With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "Are you absolutely sure you want to continue with this really bad idea? I'll give you a minute to think about it."
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .Footer = "Really, think about it."
    .Flags = TDF_USE_COMMAND_LINKS Or TDF_SHOW_PROGRESS_BAR Or TDF_CALLBACK_TIMER
    .ParenthWnd = Me.hWnd
    .AddCustomButton 101, "YeeHaw!" & vbLf & "Put some additional information about the command here."
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?"
    .VerifyText = "Hold up!"
    bRunProgress = True
    
    .ShowDialog

    bRunProgress = False
    
    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
End Sub

Private Sub Command11_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Show me the icons!"
    .Content = "Yeah, that's the stuff."
    .Footer = "Got some footer icon action here too."
    .Flags = TDF_USE_SHELL32_ICONID
    .IconMain = 18
    .IconFooter = 35
    .Title = "cTaskDialog Project"
    .CommonButtons = TDCBF_CLOSE_BUTTON

    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain

End With
End Sub

Private Sub Command12_Click()
Dim hIconM As LongPtr, hIconF As LongPtr
hIconM = ResIconToHICON("ICO_CLOCK", 32, 32)
hIconF = ResIconToHICON("ICO_HEART", 16, 16)
With TaskDialog1
    .Init
    .MainInstruction = "Let's see it all!"
    .Content = "Lots and lots of features are possible, thanks <a href=" & Chr(34) & "http://www.microsoft.com" & Chr(34) & ">Microsoft</a> for everything!"
'    .Content = "Lots and blah blah blah no link here"
    .IconMain = hIconM
    .IconFooter = hIconF
    .Flags = TDF_USE_HICON_MAIN Or TDF_USE_HICON_FOOTER Or TDF_ENABLE_HYPERLINKS Or TDF_USE_COMMAND_LINKS Or TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_CAN_BE_MINIMIZED Or TDF_DATETIME
    .DateTimeType = dttDateTimeWithCheck
    .Title = "cTaskDialog Project"
    .Footer = "Have some footer text."
    .CollapsedControlText = "Click here for some more info."
    .ExpandedControlText = "Click again to hide that extra info."
    .ExpandedInfo = "Here's a whole bunch more information you probably don't need."
    .VerifyText = "Never ever show me this dialog again!"
    .CommonButtons = TDCBF_RETRY_BUTTON Or TDCBF_CANCEL_BUTTON Or TDCBF_CLOSE_BUTTON Or TDCBF_YES_BUTTON
    .AddCustomButton 101, "YeeHaw!" & vbLf & "Some more information describing YeeHaw"
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?" & vbLf & "Or do i?"
    .AddRadioButton 110, "Let's do item 1"
    .AddRadioButton 111, "Or maybe 2"
    .AddRadioButton 112, "super secret option"
    .EnableRadioButton 112, 0
    .EnableButton 102, 0
    .SetButtonElevated TD_RETRY, 1
    bRunMarquee = True
    .ShowDialog
    bRunMarquee = False

    Label1.Caption = "ID of button clicked: " & .ResultMain
    Label2.Caption = "ID of radio button selected: " & .ResultRad
    Label3.Caption = "Verification box checked? " & .ResultVerify
End With
End Sub

Private Sub Command13_Click()
Dim td As TDBUTTONS
td = TaskDialog1.SimpleDialog("Is TaskDialogIndirect going to be better than this?", TDCBF_YES_BUTTON, App.Title, "This is regular old TaskDialog", TD_SHIELD_GRAY_ICON, Me.hWnd, App.hInstance)
Label1.Caption = "ID of button clicked: " & td

End Sub

Private Sub Command14_Click()
With TaskDialog2
    .Init
    .Content = "Here's a whole new dialog with all the options."
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_RETRY_BUTTON
    .IconMain = TD_SHIELD_OK_ICON
    .Title = "cTaskDialog Project - Page 2"
End With
With TaskDialog1
    .Init
    .MainInstruction = "You can now have multiple pages."
    .Content = "Click Next Page to continue."
    .Flags = TDF_USE_COMMAND_LINKS
    .AddCustomButton 200, "Next Page" & vbLf & "Click here to continue to the next TaskDialog"
    .CommonButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
    .IconMain = TD_SHIELD_WARNING_ICON
    .ParenthWnd = Me.hWnd
    .SetButtonHold 200
    .Title = "cTaskDialog Project - Page 1"
    .ShowDialog
End With
Label1.Caption = TaskDialog1.ResultMain
End Sub


Private Sub Command15_Click()
With TaskDialog1
    .Init
    .Content = "Input Required"
    .Flags = TDF_INPUT_BOX
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With

End Sub

Private Sub Command16_Click()
Dim hIcon1 As LongPtr, hIcon2 As LongPtr
hIcon1 = ResIconToHICON("ICO_CLOCK", 16, 16)
'hIcon2 = ResIconToHICON("ICO_HEART", 32, 32)
hIcon2 = ResIconToHICON("ICO_HEART", 16, 16)
With TaskDialog1
    .Init
    .MainInstruction = "Look at the pretty icons."
    .IconMain = TD_SHIELD_GRADIENT_ICON
    .Title = "cTaskDialog Project"
'    .Flags = TDF_USE_COMMAND_LINKS_NO_ICON
    .CommonButtons = TDCBF_CLOSE_BUTTON Or TDCBF_NO_BUTTON
    .AddCustomButton 103, "Button 1", hIcon2
    .AddCustomButton 102, "Button 2"
    .SetCommonButtonIcon TDCBF_NO_BUTTON, hIcon1
    .ShowDialog
Call DestroyIcon(hIcon1)

    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
End Sub

Private Sub Command17_Click()

With TaskDialog1
    .Init
    .Content = "Something somesuch hows-it what-eva" '& vbCrLf & vbCrLf & vbCrLf & vbCrLf
    .Flags = TDF_INPUT_BOX Or TDF_USE_COMMAND_LINKS 'Or TDF_EXPAND_FOOTER_AREA
    .InputAlign = TDIBA_Footer
    .AddCustomButton 101, "Test" & vbLf & "blah"
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
'    .IconFooter = TD_INFORMATION_ICON
    .VerifyText = "Check mate"
    .ExpandedControlText = "Gimme some more"
    .ExpandedInfo = "Here you are sir."
    .Title = "cTaskDialog Project"
    .Footer = "$input"
    .IconFooter = TD_INFORMATION_ICON
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command18_Click()
Set TaskDialogPW = New cTaskDialog
With TaskDialogPW
    .Init
    .MainInstruction = "Authorization Required"
    .Content = "The password is: password"
    .Flags = TDF_INPUT_BOX
    .InputIsPassword = True
    .InputAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_NO_BUTTON 'Or TDCBF_CANCEL_BUTTON
    .SetButtonElevated TD_OK, 1
    .SetButtonHold TD_OK
    .Footer = "Enter your password then press OK to continue."
    .IconFooter = TD_INFORMATION_ICON
    .IconMain = TD_SHIELD_ERROR_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command19_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Duplicates"
    .Content = "If you want to exclude an Artists name from the search:" & vbCrLf & vbCrLf
    .Flags = TDF_INPUT_BOX Or TDF_VERIFICATION_FLAG_CHECKED
    .AddCustomButton 100, "Continue"
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_SHIELD_ICON
    .Title = "cTaskDialog Project"
    .InputText = "Enter Artist name here."
    .VerifyText = "Exclude Jingles"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = 100 Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With

End Sub

Private Sub Command2_Click()
With TaskDialog1
    .Init
    .MainInstruction = "test"
'    .Flags = TDF_CAN_BE_MINIMIZED 'TDF_KILL_SHIELD_ICON
'    .Flags = TDF_ALLOW_DIALOG_CANCELLATION
    .Content = "This is a simple dialog."
    .CommonButtons = TDCBF_YES_BUTTON Or TDCBF_CLOSE_BUTTON 'Or TDCBF_CANCEL_BUTTON
    .IconMain = IDI_ERROR
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
'    .hinst = 0
    .ShowDialog

    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    ElseIf .ResultMain = TD_NO Then
        Label1.Caption = "Nope. No. Non. Nein."
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command20_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Input Required"
    .Content = "Tell me what I want to know!" & vbCrLf & vbCrLf
    .Flags = TDF_INPUT_BOX
    .InputAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command21_Click()
With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "First, tell me why?"
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .Flags = TDF_USE_COMMAND_LINKS Or TDF_INPUT_BOX
    .AddCustomButton 101, "YeeHaw!" & vbLf & "Put some additional information about the command here."
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?"
    
    .ShowDialog

    Label5.Caption = .ResultInput
    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
End Sub

Private Sub Command22_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Sliding on down"
    .Content = "Pick a number" '& vbCrLf & vbCrLf
    .Flags = TDF_SLIDER Or TDF_INPUT_BOX ' Or TDF_EXPANDED_BY_DEFAULTTDF_EXPAND_FOOTER_AREA Or
    .SliderAlign = TDIBA_Buttons
    .Footer = "$input"
    .InputAlign = TDIBA_Footer
    .InputWidth = -1
    .IconFooter = TD_INFORMATION_ICON
'    .ExpandedControlText = "Show more"
'    .ExpandedInfo = "Line1"
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label15.Caption = .ResultSlider
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command23_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
With TaskDialog3
    .Init
    .MainInstruction = "Duplicates"
    .Content = "If you want to exclude an Artists name from the search:"
    .Flags = TDF_VERIFICATION_FLAG_CHECKED Or TDF_COMBO_BOX 'Or TDF_INPUT_BOX
'    .InputAlign = TDIBA_Footer
    .AddCustomButton 100, "Continue"
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_SHIELD_ICON
    .Title = "cTaskDialog Project"
    .ComboCueBanner = "Cue Banner Text"
    .ComboSetInitialState "", 5
'    .ComboSetInitialItem 1
    .ComboImageList = himlSys
    .ComboAddItem "Item 1", 6
    .ComboAddItem "Item 2", 7
    .ComboAddItem "Item 3", 8
    .VerifyText = "Exclude Jingles"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label3.Caption = "Checked? " & .ResultVerify
    Label7.Caption = .ResultComboText
    Label9.Caption = .ResultComboIndex
    If .ResultMain = 100 Then
        Label1.Caption = "Continue!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command24_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
With TaskDialog1
    .Init
    .MainInstruction = "Making a list..."
    .Content = "...and checking it twice" & vbCrLf & vbCrLf
    .Flags = TDF_COMBO_BOX
    .ComboStyle = cbtDropdownList
    .AddCustomButton 100, "Continue"
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ComboSetInitialItem 0
    .ComboImageList = himlSys
    .ComboAddItem "Item 1", 6
    .ComboAddItem "Item 2", 7
    .ComboAddItem "Item 3", 8
'    .Footer = "Have you been naughty or nice?"
'    .IconFooter = IDI_QUESTION
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label7.Caption = .ResultComboText
    Label9.Caption = .ResultComboIndex
    If .ResultMain = 100 Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With

End Sub

Private Sub Command25_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
Set TaskDialogPW2 = New cTaskDialog
With TaskDialogPW2
    .Init
    .MainInstruction = "Authorization Required"
    .Content = "The password is: 'password' + user number, e.g. password1" '& vbCrLf & vbCrLf
    .Flags = TDF_INPUT_BOX Or TDF_COMBO_BOX
    .ComboStyle = cbtDropdownList
    .InputIsPassword = True
    .InputAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .SetButtonElevated TD_OK, 1
    .SetButtonHold TD_OK
    .ComboAlign = TDIBA_Content
    .ComboSetInitialItem 0
    .ComboImageList = himlSys
    .ComboAddItem "User 1", 6
    .ComboAddItem "User 2", 7
    .ComboAddItem "User 3", 8
    .Footer = "Enter your password then press OK to continue."
    .IconFooter = TD_INFORMATION_ICON
    .IconMain = TD_SHIELD_ERROR_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    Label9.Caption = .ResultComboIndex
    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command26_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
Set TaskDialogPW2 = New cTaskDialog
With TaskDialogPW2
    .Init
    .MainInstruction = "Authorization Required"
    .Content = "Select a user and password." & vbCrLf & "The password is: 'password' + user number, e.g. password1"
    .Flags = TDF_INPUT_BOX Or TDF_COMBO_BOX
    .InputIsPassword = True
    .InputAlign = TDIBA_Footer
    .InputWidth = -1
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON Or TDCBF_RETRY_BUTTON
    .SetButtonElevated TD_OK, 1
    .SetButtonHold TD_OK
    .ComboSetInitialItem 0
    .ComboAlign = TDIBA_Buttons
    .ComboImageList = himlSys
    .ComboStyle = cbtDropdownList
    .ComboAddItem "User 1", 6
    .ComboAddItem "User 2", 7
    .ComboAddItem "User 3", 8
    .Footer = "$input"
    .IconFooter = TD_INFORMATION_ICON
    .IconMain = TD_SHIELD_ERROR_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    Label9.Caption = .ResultComboIndex
    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command27_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Hello World"
    .Content = "Pick a day, any day"
    .Flags = TDF_DATETIME
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command28_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Hello World"
    .Content = "Yo u got the time bro?" '& vbCrLf & vbCrLf
    .Flags = TDF_DATETIME
    .DateTimeType = dttTime
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With

End Sub

Private Sub Command29_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Hello World"
    .Content = "Hey when u wanna do dis?" '& vbCrLf & vbCrLf
    .Flags = TDF_DATETIME
    .DateTimeType = dttDateWithCheck
    .DateTimeAlign = TDIBA_Footer
    .IconFooter = TD_INFORMATION_ICON
    .Footer = "$input"
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    Label13.Caption = .ResultDateTimeChecked
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command3_Click()
With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "Are you absolutely sure you want to continue with this really bad idea?"
    .CommonButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
    .IconMain = TD_SHIELD_WARNING_ICON 'TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    
    .ShowDialog

    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    ElseIf .ResultMain = TD_NO Then
        Label1.Caption = "Nope. No. Non. Nein."
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command30_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Hello World"
    .Content = "Pick a day, any day"
    .Flags = TDF_DATETIME Or TDF_USE_COMMAND_LINKS
    .AddCustomButton 100, "CmdLnk"
    .DateTimeType = dttDateTime
'    .DateTimeAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command31_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
With TaskDialog1
    .Init
    .MainInstruction = "Schedule Event"
    .Content = "Pick action to schedule:" '& vbCrLf & vbCrLf
    .Flags = TDF_DATETIME Or TDF_COMBO_BOX 'Or TDF_USE_COMMAND_LINKS
    '.AddCustomButton 101, "CommandL"
    .DateTimeType = dttDateTime
    .DateTimeAlign = TDIBA_Buttons
    .Width = 200 * .DPIScaleX
    .ComboStyle = cbtDropdownList
    .ComboSetInitialItem 0
    .ComboImageList = himlSys
    .ComboAddItem "Do One Thing", 6
    .ComboAddItem "Do Something Else", 7
    .ComboAddItem "Run and hide", 8
    .ComboAlign = TDIBA_Content
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .VerifyText = "Verify"
    .Footer = "Some reminder about these actions."
    .IconMain = TD_SHIELD_ICON
    .IconFooter = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog
    Label7.Caption = .ResultComboText
    Label9.Caption = .ResultComboIndex
    Label11.Caption = .ResultDateTime
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub AddCbxItems(cdg As cTaskDialog)

End Sub
Private Sub Command32_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
Dim hIconF As LongPtr
hIconF = IconToHICON(LoadResData("ICO_CLIP", "CUSTOM"), 16, 16)
    Dim hBmp As LongPtr
    Dim sImg As String
    sImg = App.Path & "\vbf.jpg"
    Dim CX As Long, CY As Long
    hBmp = hBitmapFromFile(sImg, CX, CY)
With TaskDialog1
    .Init
    .MainInstruction = "Perform Event"
    .Content = "Pick action to perform. You can schedule execution for later or enter a custom label below."
    .Flags = TDF_USE_COMMAND_LINKS Or TDF_COMBO_BOX Or TDF_DATETIME Or TDF_USE_HICON_FOOTER Or TDF_USE_SHELL32_ICONID Or TDF_KILL_SHIELD_ICON Or TDF_CAN_BE_MINIMIZED
'    .ExpandedControlText = "Expando ABCDEFGHIJKL" Or TDF_INPUT_BOX
'    .ExpandedInfo = "Test"
    .DateTimeType = dttDateTimeWithCheckTimeOnly
    .DateTimeAlign = TDIBA_Buttons
    .DateTimeAlignInButtons = tdcaRight
    .ComboAlign = TDIBA_Content
    .ComboStyle = cbtDropdownList
    .ComboSetInitialItem 1
    .ComboImageList = himlSys
    .ComboAddItem "Do Thing #1", 2
    .ComboAddItem "Do Thing #2", 7
    .ComboAddItem "Do Thing #3", 8
    .CommonButtons = TDCBF_CANCEL_BUTTON Or TDCBF_OK_BUTTON 'Or TDCBF_CLOSE_BUTTON Or TDCBF_OK_BUTTON
'    .InputText = "New Event 1"
'    .InputAlign = TDIBA_Buttons
'    .InputWidth = 140
'    .InputAlignInFooter = tdcaCenter
    .Footer = "Now you can say something else here."
'    .VerifyText = "Perform event later:"
    .IconMain = TD_SHIELD_GRADIENT_ICON
    .IconFooter = hIconF
    .IconReplaceGradient = 276
    .Title = "cTaskDialog Project"
'    .ParenthWnd = Me.hwnd
    .AddCustomButton 102, "Schedule" & vbLf & "Additional information here."
    .AddRadioButton 110, "Apply to this account only."
    .AddRadioButton 111, "Apply to all accounts."
    .SetLogoImage hBmp, LogoBitmap, LogoTopRight, 0, 0
    .ShowDialog

    Label2.Caption = "Radio: " & .ResultRad
    Label5.Caption = .ResultInput
    Label7.Caption = .ResultComboText
    Label9.Caption = .ResultComboIndex
    Label11.Caption = .ResultDateTime
    If .ResultDateTimeChecked = 0 Then
        Label13.Caption = "Time unchecked."
    Else
        Label13.Caption = "Time checked."
    End If
    If .ResultMain = 102 Then
        Label1.Caption = "Scheduled."
    Else
        Label1.Caption = "Cancelled."
    End If
End With
DeleteObject hBmp
End Sub

Private Sub Command33_Click()
Dim dTimeMin As Date, dTimeMax As Date

dTimeMin = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(13, 0, 0)
dTimeMax = DateAdd("d", 7, dTimeMin)
dTimeMax = DateAdd("h", 4, dTimeMax)

With TaskDialog1
    .Init
    .MainInstruction = "Date Ranges"
    .Content = "Pick a time, limited to sometime in the next 7 days, between 1pm and 6pm"
    .Flags = TDF_DATETIME Or TDF_INPUT_BOX Or TDF_USE_COMMAND_LINKS
    .DateTimeType = dttDateTime
    .DateTimeAlign = TDIBA_Content
    .DateTimeSetRange True, True, dTimeMin, dTimeMax
    .DateTimeSetInitial dTimeMin
    .InputAlign = TDIBA_Buttons
    .InputCueBanner = "Add an optional note to whatever."
    .AddCustomButton 101, "Set Date" & vbLf & "Apply this date and time to whatever it is you're doing."
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    If .ResultMain = 101 Then
        Label1.Caption = "Date Set"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command34_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Sup"
    .Content = "Note that if you want date/time in the buttons, there may not be enough room depending on number of buttons and whether there's checkboxes." '& vbCrLf & vbCrLf
    .Flags = TDF_DATETIME
    .DateTimeType = dttDateTimeWithCheck 'TimeOnly
    .DateTimeAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    Select Case .ResultDateTimeChecked
        Case 0: Label13.Caption = "Neither box checked."
        Case 2: Label13.Caption = "Time checked, date unchecked."
        Case 3: Label13.Caption = "Date checked, time unchecked."
        Case 4: Label13.Caption = "Both checked."
    End Select
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command35_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Sliding on down"
    .Content = "Pick a number"
    .Flags = TDF_SLIDER Or TDF_USE_COMMAND_LINKS
    .SliderSetRange 0, 100, 10
    .SliderSetChangeValues 10, 20
    .SliderTickStyle = SldTickStyleBoth
    .SliderValue = 50
    .SliderAlign = TDIBA_Content
    .ExpandedControlText = "ExpandMe"
    .ExpandedInfo = "Expanded"
    .AddCustomButton 100, "CommandLink"
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label15.Caption = .ResultSlider
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command36_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Hello World"
    .Content = "Input Required"
    .Flags = TDF_INPUT_BOX Or TDF_EXPAND_FOOTER_AREA Or TDF_EXPANDED_BY_DEFAULT  ' Or TDF_SHOW_PROGRESS_BAROr TDF_USE_COMMAND_LINKS '
'    .AddCustomButton 101, "CommandLink1" & vbLf & "Desc1"
'    .AddCustomButton 102, "CommandLink2"
    .AddRadioButton 103, "Radio 1"
    .AddRadioButton 104, "Radio 2"
    .ExpandedControlText = "Expando"
    .ExpandedInfo = "Expanded information."
'    .VerifyText = "Verification check."
    .InputAlign = TDIBA_Footer
'    .InputAlignInFooter = tdcaCenter
    
'    .InputWidth = 100
'    .Footer = "$input"
    .IconFooter = TD_INFORMATION_ICON
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON 'Or TDCBF_RETRY_BUTTON Or TDCBF_CLOSE_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label5.Caption = .ResultInput
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command37_Click()
himlSys = GetSystemImagelist(SHGFI_SMALLICON)
With TaskDialog3
    .Init
    .MainInstruction = "Main Instruct"
    .Content = "Content goes here."
    .Flags = TDF_COMBO_BOX Or TDF_USE_COMMAND_LINKS Or TDF_SHOW_MARQUEE_PROGRESS_BAR 'Or TDF_EXPANDED_BY_DEFAULT  Or TDF_EXPAND_FOOTER_AREA  '
    .CommonButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
    .IconMain = TD_SHIELD_ICON
    .Title = "cTaskDialog Project"
    .ComboCueBanner = "Cue Banner Text"
    .ComboSetInitialState "", 5
    .ComboAlign = TDIBA_Footer
'    .ComboAlignInFooter = tdcaCenter
'    .ComboSetInitialItem 1
    .ComboImageList = himlSys
'    .ComboStyle = cbtDropdownList
    .ComboAddItem "Item 1", 6
    .ComboAddItem "Item 2", 7
    .ComboAddItem "Item 3", 8
    .AddCustomButton 101, "CommandLink1" & vbLf & "Desc1"
    .AddCustomButton 102, "CommandLink2"
'    .AddRadioButton 103, "Radio 1"
'    .AddRadioButton 104, "Radio 2"
    .ExpandedControlText = "Expando"
    .ExpandedInfo = "Expanded information."
    .VerifyText = "Verification check."
    .IconFooter = TD_ERROR_ICON
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label7.Caption = .ResultComboText
    Label9.Caption = .ResultComboIndex
    If .ResultMain = 100 Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command38_Click()
With TaskDialog1
    .Init
'    .MainInstruction = "Hello World"
    .Content = "Pick a day, any day."
    .Flags = TDF_DATETIME Or TDF_EXPANDED_BY_DEFAULT Or TDF_USE_COMMAND_LINKS Or TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_EXPANDED_BY_DEFAULT 'TDF_EXPAND_FOOTER_AREA  '
    .DateTimeType = dttDateTimeWithCheckTimeOnly
    .DateTimeAlign = TDIBA_Footer
    .DateTimeAlignInFooter = tdcaRight
    .AddCustomButton 101, "CommandLink1" & vbLf & "Desc1"
    .AddCustomButton 102, "CommandLink2"
    .AddRadioButton 103, "Radio 1"
    .AddRadioButton 104, "Radio 2"
    .ExpandedControlText = "Expando blah blah"
    .ExpandedInfo = "Expanded information."
'    .VerifyText = "Verification check.sggsgdggggggg"
    
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .IconFooter = TD_ERROR_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label11.Caption = .ResultDateTime
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command39_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Sliding on down"
    .Content = "Pick a number"
    .Flags = TDF_SLIDER Or TDF_USE_COMMAND_LINKS Or TDF_EXPANDED_BY_DEFAULT ' Or TDF_EXPAND_FOOTER_AREA  TDF_SHOW_MARQUEE_PROGRESS_BAR  Or
'    .SliderTickStyle = SldTickStyleBoth
'    .SliderAlign = TDIBA_Footer
    .AddCustomButton 101, "CommandLink1" & vbLf & "Desc1"
    .AddCustomButton 102, "CommandLink2"
'    .AddRadioButton 103, "Radio 1"
'    .AddRadioButton 104, "Radio 2"
    .ExpandedControlText = "Expando"
    .ExpandedInfo = "Expanded information."
'    .VerifyText = "Verification check."
    .IconFooter = TD_INFORMATION_ICON
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label15.Caption = .ResultSlider
    If .ResultMain = TD_OK Then
        Label1.Caption = "Yes Yes Yes!"
    Else
        Label1.Caption = "Cancelled."
    End If
End With

End Sub

Private Sub Command4_Click()
With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "Are you absolutely sure you want to continue with this really bad idea?"
    .IconMain = TD_ERROR_ICON
    .Title = "cTaskDialog Project"
    .AddCustomButton 101, "YeeHaw!"
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?"
    
    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
End Sub

Private Sub Command40_Click()
Dim hIco16 As LongPtr
hIco16 = ResIconToHICON("ICO_HEART", 16, 16) 'IconToHICON(LoadResData("ICO_CLIP", "CUSTOM"), 16, 16)
Set TaskDialogSC = New cTaskDialog
With TaskDialogSC
    .Init
    .Flags = TDF_INPUT_BOX 'TDF_KILL_SHIELD_ICON 'Or TDF_USE_IMAGERES_ICONID
'    .CommonButtons = TDCBF_NO_BUTTON
    .Title = "TestTitle"
    .Content = "TestContent"
    .ParenthWnd = Me.hWnd
    .MainInstruction = "TestInstruction"
    .IconMain = TD_INFORMATION_ICON
'    .AddCustomButton 122, "Button 1"
    .AddCustomButton 123, "SuperButton ", hIco16
'    .AddCustomButton 124, "Button 3"
    .SetSplitButton 123
    .ShowDialog
Label1.Caption = .ResultMain
Label5.Caption = .ResultInput

End With

End Sub



Private Sub Command41_Click()
Dim dTimeMin As Date, dTimeMax As Date
himlSys = GetSystemImagelist(SHGFI_SMALLICON)

dTimeMin = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(13, 0, 0)
dTimeMax = DateAdd("d", 7, dTimeMin)
dTimeMax = DateAdd("h", 4, dTimeMax)
    Dim hBmp As LongPtr
    Dim sImg As String
    sImg = App.Path & "\disc32.png"
    Dim CX As Long, CY As Long
    hBmp = hBitmapFromFile(sImg, CX, CY)
'    hBmp = LoadImageW(0, StrPtr(simg), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    Debug.Print "hBmp=" & hBmp '& ",cx=" & cx & ",cy=" & cy
With TaskDialog1
    .Init
    .MainInstruction = "Set Action"
'    .Content = "Pick a time, limited to sometime in the next 7 days, between 1pm and 6pm"
    .Content = "Execute this action now or choose a new time below." & vbCrLf & "For additional help: <A HREF=""www.microsoft.com"">Microsoft</A> on the web - <A HREF=""http://msdn.microsoft.com/"">MSDN</A> on the web"
    .Flags = TDF_DATETIME Or TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_INPUT_BOX Or TDF_KILL_SHIELD_ICON Or TDF_ENABLE_HYPERLINKS Or TDF_COMBO_BOX  'Or TDF_USE_COMMAND_LINKS
'    .AddRadioButton 501, "Radio 1"
'    .AddRadioButton 502, "Radio 2"
'    .ExpandedControlText = "ExpandMe!"
'    .ExpandedInfo = "blahdy blah blah"
    .DateTimeType = dttDateTime
    .DateTimeAlign = TDIBA_Footer
'    .DateTimeAlignInContent = tdcaCenter
    .DateTimeAlignInFooter = tdcaRight
    .DateTimeSetRange True, True, dTimeMin, dTimeMax
    .DateTimeSetInitial dTimeMin
    .InputAlign = TDIBA_Content
    .InputCueBanner = "Add an optional note to whatever."
    .ComboAlign = TDIBA_Buttons
    .ComboCueBanner = "Cue Banner Text"
    .ComboSetInitialState "", 5
'    .ComboSetInitialItem 2
    .ComboImageList = himlSys
    .ComboAddItem "Item 1", 6
    .ComboAddItem "Item 2", 7
    .ComboAddItem "Item 3", 8
    .ComboWidth = -1
'    .DefaultButton = TD_CANCEL
'    .VerifyText = "Confirm something or another."
    .IconFooter = TD_INFORMATION_ICON
    .Footer = "<A HREF=""www.microsoft.com"">Choose</A> date and time:"
    .AddCustomButton 101, "Set Date" ' & vbLf & "Apply this date and time to whatever it is you're doing."
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_SHIELD_GRAY_ICON
'    .hinst = 0
'    .Footer = "<A HREF=""www.microsoft.com"">Microsoft</A> on the web" & _
'                                                  " - <A HREF=""http://msdn.microsoft.com/"">MSDN</A> on the web"
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .SetLogoImage hBmp, LogoBitmap, LogoTopRight, 4, 4 'LogoButtons
    bRunMarquee = True
    .ShowDialog
    bRunMarquee = False

    Label11.Caption = .ResultDateTime
    If .ResultMain = 101 Then
        Label1.Caption = "Date Set"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
Call DeleteObject(hBmp)

End Sub

Private Sub Command42_Click()
Dim dTimeMin As Date, dTimeMax As Date
himlSys = GetSystemImagelist(SHGFI_SMALLICON)

dTimeMin = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(13, 0, 0)
dTimeMax = DateAdd("d", 7, dTimeMin)
dTimeMax = DateAdd("h", 4, dTimeMax)
    Dim hBmp As LongPtr
    Dim sImg As String
    sImg = App.Path & "\vbf.jpg"
    Dim CX As Long, CY As Long
    hBmp = hBitmapFromFile(sImg, CX, CY)
'    hBmp = LoadImageW(0, StrPtr(simg), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    Debug.Print "hBmp=" & hBmp '& ",cx=" & cx & ",cy=" & cy
With TaskDialog1
    .Init
    .MainInstruction = "Set Action"
'    .Content = "Pick a time, limited to sometime in the next 7 days, between 1pm and 6pm"
    .Content = "Execute this action now or choose a new time below." & vbCrLf & "For additional help: <A HREF=""www.microsoft.com"">Microsoft</A> on the web - <A HREF=""http://msdn.microsoft.com/"">MSDN</A> on the web"
    .Flags = TDF_DATETIME Or TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_INPUT_BOX Or TDF_ENABLE_HYPERLINKS  ' Or TDF_COMBO_BOX 'Or TDF_USE_COMMAND_LINKS
'    .AddRadioButton 501, "Radio 1"
'    .AddRadioButton 502, "Radio 2"
'    .ExpandedControlText = "ExpandMe!"
'    .ExpandedInfo = "blahdy blah blah"
    .DateTimeType = dttDateTime
    .DateTimeAlign = TDIBA_Footer
'    .DateTimeAlignInContent = tdcaCenter
    .DateTimeAlignInFooter = tdcaRight
    .DateTimeSetRange True, True, dTimeMin, dTimeMax
    .DateTimeSetInitial dTimeMin
    .InputAlign = TDIBA_Content
    .InputCueBanner = "Add an optional note to whatever."
    .ComboAlign = TDIBA_Content
    .ComboCueBanner = "Cue Banner Text"
    .ComboSetInitialState "", 5
'    .ComboSetInitialItem 2
    .ComboImageList = himlSys
    .ComboAddItem "Item 1", 6
    .ComboAddItem "Item 2", 7
    .ComboAddItem "Item 3", 8
    .ComboWidth = -1
'    .DefaultButton = TD_CANCEL
'    .VerifyText = "Confirm something or another."
    .IconFooter = TD_INFORMATION_ICON
    .Footer = "<A HREF=""www.microsoft.com"">Choose</A> date and time:"
    .AddCustomButton 101, "Set Date" ' & vbLf & "Apply this date and time to whatever it is you're doing."
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_ERROR_ICON
'    .hinst = 0
'    .Footer = "<A HREF=""www.microsoft.com"">Microsoft</A> on the web" & _
'                                                  " - <A HREF=""http://msdn.microsoft.com/"">MSDN</A> on the web"
    .Title = "cTaskDialog Project"
    .ParenthWnd = Me.hWnd
    .SetLogoImage hBmp, LogoBitmap, LogoButtons
    bRunMarquee = True
    .ShowDialog
    bRunMarquee = False

    Label11.Caption = .ResultDateTime
    If .ResultMain = 101 Then
        Label1.Caption = "Date Set"
    Else
        Label1.Caption = "Cancelled."
    End If
End With
Call DeleteObject(hBmp)
End Sub

Private Sub Command43_Click()
Set TaskDialogMPX1 = New cTaskDialog
Set TaskDialogMPX2 = New cTaskDialog
Set TaskDialogMPX3 = New cTaskDialog
Debug.Print "1: " & ObjPtr(TaskDialogMPX1) & ", 2: " & ObjPtr(TaskDialogMPX2) & ", 3: " & ObjPtr(TaskDialogMPX3)

sMPLogin = ""
With TaskDialogMPX3
    .Init
    .PageIndex = 3
    .MainInstruction = "dummy"
    .Content = "We're now doing stuff..."
    .CommonButtons = TDCBF_OK_BUTTON
    .IconMain = TD_SHIELD_OK_ICON
    .Flags = TDF_SHOW_MARQUEE_PROGRESS_BAR Or TDF_USE_COMMAND_LINKS
    .AddCustomButton 310, "Restart process" & vbLf & "Click to return to the previous page."
    .SetButtonHold 310
    .Title = "cTaskDialog Project - Page 3"
End With
With TaskDialogMPX2
    .Init
    .PageIndex = 2
    .MainInstruction = "Log In"
    .Content = "The password is: 'password' + user number, e.g. password1" '& vbCrLf & vbCrLf
    .Flags = TDF_INPUT_BOX Or TDF_COMBO_BOX
    .ComboStyle = cbtDropdownList
    .InputIsPassword = True
    .InputAlign = TDIBA_Buttons
    .CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    .SetButtonElevated TD_OK, 1
    .SetButtonHold TD_OK
    .ComboAlign = TDIBA_Content
    .ComboSetInitialItem 0
    If (himlSys = 0) Then himlSys = GetSystemImagelist(SHGFI_SMALLICON)
    .ComboImageList = himlSys
    .ComboAddItem "User 1", 6
    .ComboAddItem "User 2", 7
    .ComboAddItem "User 3", 8
    .Footer = "Enter your password then press OK to continue."
    .IconFooter = TD_INFORMATION_ICON
    .IconMain = TD_SHIELD_GRAY_ICON
    .Title = "cTaskDialog Project - Page 2"
    .ParenthWnd = Me.hWnd
End With
With TaskDialogMPX1
    .Init
    .PageIndex = 1
    .MainInstruction = "Mutli-page Testing"
    .Content = "Choose how you want to proceed."
    .Flags = TDF_USE_COMMAND_LINKS
    .AddCustomButton 200, "Proceed anonymously" & vbLf & "Click here to continue without logging in."
    .AddCustomButton 201, "Set log in information" & vbLf & "Select your username."
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .IconMain = TD_SHIELD_ICON
    .ParenthWnd = Me.hWnd
    .SetButtonHold 200
    .SetButtonHold 201
    .Title = "cTaskDialog Project - Page 1"
    bPageExampleEx = True
    .ShowDialog
    bPageExampleEx = False
    Label1.Caption = .ResultMain
    Label5.Caption = .ResultInput
    Label17.Caption = .PageIndex
End With
Label1.Caption = TaskDialog1.ResultMain
End Sub

Private Sub Command44_Click()
With TaskDialogAC
    .Init
    .MainInstruction = "Do you wish to do somethingsomesuch?"
    .Flags = TDF_CALLBACK_TIMER Or TDF_USE_COMMAND_LINKS Or TDF_SHOW_PROGRESS_BAR
    .Content = "Execute it then, otherwise I'm gonna peace out."
    .AddCustomButton 101, "Let's Go!" & vbLf & "Really, let's go."
    .CommonButtons = TDCBF_CLOSE_BUTTON
    .IconMain = IDI_QUESTION
    .IconFooter = TD_ERROR_ICON
    .Footer = "Closing in 15 seconds..."
    .Title = "cTaskDialog Project"
    .AutocloseTime = 15 'seconds
    .ParenthWnd = Me.hWnd
'    .hinst = 0
    .ShowDialog

    If .ResultMain = TD_YES Then
        Label1.Caption = "Yes Yes Yes!"
    ElseIf .ResultMain = TD_NO Then
        Label1.Caption = "Nope. No. Non. Nein."
    Else
        Label1.Caption = "Cancelled."
    End If
End With
End Sub

Private Sub Command5_Click()
With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "Are you absolutely sure you want to continue with this really bad idea? So just exactly how damn wide are you son of bitching bastards planning on making this before you get around to wrapping my text?"
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .AddCustomButton 101, "YeeHaw!"
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?"
    .AddRadioButton 110, "Let's do item 1"
    .AddRadioButton 111, "Or maybe 2"
    .AddRadioButton 112, "super secret option"
    .Flags = TDF_SIZE_TO_CONTENT
    .Width = 50
    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain
    Label2.Caption = "ID of radio button selected: " & .ResultRad
    
End With
End Sub

Private Sub Command6_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Let's see some hyperlinking!"
    .Content = "Where else to link to but <a href=""http://www.microsoft.com"">Microsoft.com</a>"
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .CommonButtons = TDCBF_CLOSE_BUTTON
    .Flags = TDF_ENABLE_HYPERLINKS
    .ParenthWnd = Me.hWnd
    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain
    Label2.Caption = "ID of radio button selected: " & .ResultRad
    
End With
End Sub

Private Sub Command7_Click()
Dim hIconM As LongPtr, hIconF As LongPtr
hIconM = IconToHICON(LoadResData("ICO_CLIP", "CUSTOM"), 32, 32)
'hIconM = ResIconToHICON("ICO_CLOCK", 32, 32)
hIconF = ResIconToHICON("ICO_HEART", 16, 16)
With TaskDialog1
    .Init
    .MainInstruction = "What time is it?"
    .Content = "Is is party time yet???"
    .Footer = "Don't you love TaskDialogIndirect?"
    .Flags = TDF_USE_HICON_MAIN Or TDF_USE_HICON_FOOTER
    .IconMain = hIconM
    .IconFooter = hIconF
    .Title = "cTaskDialog Project"
    .CommonButtons = TDCBF_CLOSE_BUTTON
    
    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
Call DestroyIcon(hIconM)
Call DestroyIcon(hIconF)

End Sub

Private Sub Command8_Click()
With TaskDialog1
    .Init
    .MainInstruction = "Let's see all the basic fields."
    .Content = "We can really fit in a lot of organized information now."
    .Title = "cTaskDialog Project"
    .Footer = "Have some footer text."
'    .CollapsedControlText = "Click here for some more info."
    .ExpandedControlText = "Click again to hide that extra info."
    .ExpandedInfo = "Here's some more info we don't really need."
    .VerifyText = "Never ever show me this dialog again!"
    
    .IconMain = TD_INFORMATION_ICON
    .IconFooter = TD_ERROR_ICON
    
    .ShowDialog
    
    Label1.Caption = "ID of button clicked: " & .ResultMain
    Label2.Caption = "Box checked? " & .ResultVerify
End With
End Sub

Private Sub Command9_Click()

With TaskDialog1
    .Init
    .MainInstruction = "You're about to do something stupid."
    .Content = "Are you absolutely sure you want to continue with this really bad idea?"
    .IconMain = TD_INFORMATION_ICON
    .Title = "cTaskDialog Project"
    .CommonButtons = TDCBF_CANCEL_BUTTON
    .Flags = TDF_USE_COMMAND_LINKS
    .AddCustomButton 101, "YeeHaw!" & vbLf & "Put some additional information about the command here."
    .AddCustomButton 102, "NEVER!!!"
    .AddCustomButton 103, "I dunno?"
    
    .ShowDialog

    Label1.Caption = "ID of button clicked: " & .ResultMain
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set TaskDialog1 = Nothing
Set TaskDialog2 = Nothing
FreeGDIPlus gdipInitToken

End Sub


Private Sub TaskDialog1_ButtonClick(ByVal ButtonID As Long)
Debug.Print "TaskDialog1_ButtonClick " & ButtonID
If ButtonID = 200 Then
    TaskDialog1.NavigatePage TaskDialog2
End If
End Sub


Private Sub TaskDialog1_ComboItemChanged(ByVal iNewItem As Long)
Debug.Print "ComboItmChg " & iNewItem
End Sub

Private Sub TaskDialog1_DateTimeChange(ByVal dtNew As Date, ByVal lCheckStatus As Long)
Debug.Print "DateTimeChange " & dtNew

End Sub

Private Sub TaskDialog1_DialogDestroyed()
Timer1.Enabled = False
bRunProgress = False
End Sub

Private Sub TaskDialog1_HyperlinkClick(ByVal lPtr As LongPtr)

Call ShellExecuteW(0, 0, lPtr, 0, 0, SW_SHOWNORMAL)

End Sub
Private Sub Form_Load()
gdipInitToken = InitGDIPlus
Set TaskDialog1 = New cTaskDialog
Set TaskDialog2 = New cTaskDialog
Set TaskDialog3 = New cTaskDialog
Set TaskDialogAC = New cTaskDialog
Set TaskDialogMPX1 = New cTaskDialog
Set TaskDialogMPX2 = New cTaskDialog
End Sub

Private Sub TaskDialog1_DialogCreated(ByVal hWnd As LongPtr)
If bRunProgress Then
    Timer1.Enabled = True
    TaskDialog1.ProgressSetRange 0, 60
End If
If bRunMarquee Then
    TaskDialog1.ProgressStartMarquee
End If

End Sub


Private Sub TaskDialog1_InputBoxChange(sText As String)
Debug.Print "InputChange=" & sText
End Sub

Private Sub TaskDialog1_Navigated()
Debug.Print "TaskDialog1_Navigated()"
If bRunMarquee2 Then
    TaskDialog1.ProgressStartMarquee
End If
End Sub

Private Sub TaskDialog1_SliderChange(ByVal lNewValue As Long)
Debug.Print "SliderChange=" & lNewValue
End Sub

Private Sub TaskDialog1_Timer(ByVal TimerValue As Long)

If lSecs > 60 Then
    Timer1.Enabled = False
    bRunProgress = False
Else
    TaskDialog1.ProgressSetValue lSecs
    TaskDialog1.Footer = "You've been thinking for " & lSecs & " seconds now..."
End If

End Sub

Private Sub TaskDialog1_VerificationClicked(ByVal Value As Long)
If Value = 1 Then
    Timer1.Enabled = False
    bRunProgress = False
Else
    bRunProgress = True
    Timer1.Enabled = True
End If
End Sub

Private Sub TaskDialog2_ButtonClick(ByVal ButtonID As Long)
Debug.Print "TaskDialog2_ButtonClick " & ButtonID

End Sub

Private Sub TaskDialog2_DialogConstucted(ByVal hWnd As LongPtr)
Debug.Print "TaskDialog2_DialogConstucted"

End Sub

Private Sub TaskDialog2_DialogCreated(ByVal hWnd As LongPtr)
Debug.Print "TaskDialog2_DialogCreated"
If bRunMarquee2 Then
    TaskDialog1.ProgressStartMarquee
End If

End Sub

Private Sub TaskDialog2_DropdownButtonClicked(ByVal hWnd As LongPtr)
Debug.Print "TD2 ButtonDropdown"
End Sub

Private Sub TaskDialog2_InputBoxChange(sText As String)
Debug.Print "TD2 Input=" & sText
End Sub

Private Sub TaskDialog3_DialogCreated(ByVal hWnd As LongPtr)
'Call SendMessageW(TaskDialog3.hWndCombo, CB_SETDROPPEDWIDTH, 900&, ByVal 0&)
End Sub

Private Sub TaskDialog3_InputBoxChange(sText As String)
Debug.Print "InputChange=" & sText

End Sub

Private Sub TaskDialogAC_DialogCreated(ByVal hWnd As LongPtr)
TaskDialogAC.ProgressSetRange 0, 15
TaskDialogAC.ProgressSetState ePBST_ERROR
End Sub

Private Sub TaskDialogAC_Timer(ByVal TimerValue As Long)
On Error Resume Next
TaskDialogAC.Footer = "Closing in " & TaskDialogAC.AutocloseTime & " seconds..."
TaskDialogAC.ProgressSetValue 15 - TaskDialogAC.AutocloseTime
On Error GoTo 0
End Sub

Private Sub TaskDialogMPX1_ButtonClick(ByVal ButtonID As Long)
Debug.Print "TaskDialogMPX1_ButtonClick id=" & ButtonID & ",page=" & TaskDialogMPX1.PageIndex
If bPageExampleEx Then
    'All button clicks for multi-page dialogs go through the dialog from .ShowDialog.
    'To more easily avoid collisions, and for other reasons, there's now a PageIndex
    'to keep track of the page that's actually showing.
    If TaskDialogMPX1.PageIndex = 1 Then
        If ButtonID = 201 Then
            TaskDialogMPX1.NavigatePage TaskDialogMPX2
        ElseIf ButtonID = 200 Then
            sMPLogin = "Anonymous"
            TaskDialogMPX1.NavigatePage TaskDialogMPX3
        End If
    End If
    If TaskDialogMPX1.PageIndex = 2 Then
        'Remember, even though all the items were set up on TaskDialogMPX2, they're
        'applied to the main dialog, which is TaskDialogMPX1, so we address that one.
        Dim sPW As String
        If ButtonID = TD_OK Then
            Select Case TaskDialogMPX1.ComboIndex
                Case 0: sPW = "password1"
                Case 1: sPW = "password2"
                Case 2: sPW = "password3"
            End Select
            If TaskDialogMPX1.InputText = sPW Then
                sMPLogin = "User " & (TaskDialogMPX1.ComboIndex + 1)
                TaskDialogMPX1.NavigatePage TaskDialogMPX3
            Else
                MessageBeep MB_ERROR
                Debug.Print TaskDialogMPX1.IconFooter
                TaskDialogMPX1.Footer = "Wrong password, try again."
                TaskDialogMPX1.IconFooter = TD_ERROR_ICON
            End If
        End If
    End If
    If TaskDialogMPX1.PageIndex = 3 Then
        If ButtonID = 310 Then 'Reset to page 1
            With TaskDialogMPX1
                .Init
                .PageIndex = 1
                .MainInstruction = "Mutli-page Testing"
                .Content = "Choose how you want to proceed."
                .Flags = TDF_USE_COMMAND_LINKS
                .AddCustomButton 200, "Proceed anonymously" & vbLf & "Click here to continue without logging in."
                .AddCustomButton 201, "Set log in information" & vbLf & "Select your username."
                .CommonButtons = TDCBF_CANCEL_BUTTON
                .IconMain = TD_SHIELD_ICON
                .ParenthWnd = Me.hWnd
                .SetButtonHold 200
                .SetButtonHold 201
                .Title = "cTaskDialog Project - Page 1"
            End With
            TaskDialogMPX1.NavigatePage TaskDialogMPX1
        End If
    End If
End If
End Sub

Private Sub TaskDialogMPX1_Navigated()
Debug.Print "TaskDialogMPX1_Navigated()"
If TaskDialogMPX1.PageIndex = 3 Then
    TaskDialogMPX1.ProgressStartMarquee
    TaskDialogMPX1.MainInstruction = "Logged in as " & sMPLogin
End If
End Sub

Private Sub TaskDialogPW_ButtonClick(ByVal ButtonID As Long)
If ButtonID = TD_OK Then
    If TaskDialogPW.InputText = "password" Then
        TaskDialogPW.CloseDialog
    Else
        MessageBeep MB_ERROR
        TaskDialogPW.Footer = "Wrong password, please try again."
        TaskDialogPW.IconFooter = TD_ERROR_ICON
    End If
End If
End Sub

Private Sub TaskDialogPW2_ButtonClick(ByVal ButtonID As Long)
Dim sPW As String
If ButtonID = TD_OK Then
    Select Case TaskDialogPW2.ComboIndex
        Case 0: sPW = "password1"
        Case 1: sPW = "password2"
        Case 2: sPW = "password3"
    End Select
    If TaskDialogPW2.InputText = sPW Then
        TaskDialogPW2.CloseDialog
    Else
        MessageBeep MB_ERROR
        TaskDialogPW2.Footer = "Wrong password, try again."
        TaskDialogPW2.IconFooter = TD_ERROR_ICON
    End If
End If
End Sub

Private Sub TaskDialogSC_DropdownButtonClicked(ByVal hWnd As LongPtr)
Debug.Print "Got DropDown Button!"
End Sub

Private Sub Timer1_Timer()
lSecs = lSecs + 1
End Sub
