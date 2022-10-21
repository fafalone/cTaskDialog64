VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaskDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************
'cTaskDialog v1.2 R2 X64 Version
'by Jon Johnson (aka fafalone)
'fafalone@gmail.com
'
'Originally Released 21 March 2020
'x64 Version released 21 October 2022
'*********************************
'
'
'Provides a full implementation of TaskDialogIndirect and extensive enhancements.
'
'Code may be modified, reused, republished.. anything as long as credit is given
'
'Some enums and declarations originally written by Randy Birch at VBNet and Julius Laslo at vb@rchiv
'
'This version is compatible with VB6, VBA6, VBA7 32/64bit, and twinBASIC 32/64bit.
'*********************************
'USAGE
'For full details on usage and support, see the development thread at VBForums:
'http://www.vbforums.com/showthread.php?777021
'The thread contains detailed samples complete with images for all features.
'Submit any bug reports in that thread as well.
'
'*********************************
'REQUIREMENTS
'--Common Controls 6.0 are required. Your application, and the VB IDE if running from there,
'  must include a manifest specifying 6.0. The cTaskDialog sample project includes one in its
'  resource file.
'
'--mTaskDialogHelper must be added to your project. While twinBASIC supports AddressOf on 
'  class members, I wanted to retain VBA support too.
'
'--The TaskDialogIndirect API is only available on Windows Vista and newer.
'
'*********************************
'VERSION HISTORY
'Version 1.2 R2 x64
'-Combos in DropDownList style now also report text.
'-(Bug fix) Multiple pages having their events directed to the first page was actually a bug,
'           not an unavoidable feature. Now, all events (including Navigated) fire in the page
'           itself. The multi-page demos have been updated to reflect this, please check your
'           code to adjust it if you're using this feature.
'
'Version 1.2 Release 2 (R2)
'--Bug fix: Quick bug fix for footer icons not redrawing when changed during runtime.
'
'NEW IN VERSION 1.2
'--Since it turns out, when you navigate to a new page, it doesn't activate the new TaskDialog
'  object, it just takes the settings and applies them to the original object. So the outcome
'  of this was that only the settings that are part of the native TaskDialog would be applied,
'  and see also the bugfix related to sending messages to active dialogs. Any of the custom
'  features wouldn't get applied. All settings are transferred now.
'  NOTE: Since this has to be done by overwriting the current configuration set, if you move
'    back to an earlier page, you'll need to reset the settings first. All of them, not just
'    customizations.
'--Added InputBoxTextAlign property to set left/right/center alignment.
'--Added User Options section below this changelong that includes a boolean for whether to print
'  debug messages in the immediate window, and whether to include the date/time if it does. Also
'  added an option for the minimum level of message to show, but so far no debug messages in the
'  project are set to a different level.
'--Bug fix: Flag for DateTime and Slider were assigned same value.
'--Bug fix: Input box should have had ES_AUTOHSCROLL set.
'--Bug fix: In ClickVerification, it used the message for a button click, which led to either
'           closing the dialog or doing nothing, instead of TDM_CLICK_VERIFICATION, the correct
'           message to send for modifying the checkbox through code.
'--Bug fix: Navigating to a new page zero'd the hWnd on the expectation it would change, but
'           this expectation was incorrect. It shouldn't be zero'd and any action applied should
'           be sent to the main TaskDialog object, the API messages will apply to the current
'           page. E.g. TaskDialog1 is the first page, which loads a second page with object
'           TaskDialog2. When navigating to TaskDialog2, TaskDialog1 receives a Navigated() event,
'           and if TaskDialog2 has a marquee progress bar, call TaskDialog1.ProgressStartMarquee.
'
'NEW IN VERSION 1.1
'--The cTaskDialog class is now entirely self-contained- using Paul Caton/LaVolpe's self-sub
'  self-callback routines to eliminate the need for a module or presence of the TaskDialog__Procs.
'--Additional left-right-center alignment for custom controls has been added. For the additions,
'  they are only relevent if the control is being used with a fixed width. Given this, automatic
'  prevention of only allowing 1 control per location has been disabled-- however, it is the on
'  the caller to ensure there's no overlap.
'--To further aid in this, you can now specify a manual offset with .(control)OffsetX that is
'  added to the default offset after alignment calculation. Note that for right alignment, to get
'  the control farther to the left, you'd want to make this a negative number.
'--ComboText now also a let that sets the text. This does not add an item and cannot set an image
'  if an imagelist is in use; for those features use ComboSetCurrentState
'--Added Get/Let for Combo dropdown width (.ComboDropWidth). Returns 0 unless previously set. Can
'  modify an open dialog.
'--Bug fix: TDF_NO_SET_FOREGROUND flag, which was added in Windows 8, was missing.
'--Bug fix: Alignment issues with certain flags with the DateTime control when placed in the footer.
'--Bug fix: Flag conflicts with built in ones and omitted flags. Fixed.
'
'NEW IN VERSION 1.0
'--You can now make a custom button into a Split Button (dropdown arrow). Use .SetSplitButton
'  with the ID of the custom button. This fires a DropdownButtonClicked event. Cannot be used
'  with Command Links. Note there's one more function prototype in mTaskDialogHelper you'll
'  need to copy if you're placing those in your own module (they've been condensed too so the
'  whole thing is only 3 lines now).
'--A logo type image can now be placed with .SetLogoImage. The current placement options are
'  the top right corner, or next to the buttons (only if no controls, either custom or the
'  expando/verify are present). The image is passed as an HBITMAP or HICON, so you can load
'  in whichever way you want. No copy is made, so you must destroy the image yourself, but
'  only after the dialog closes (or it will not appear).
'--Autoclose has been implemented. Set the .AutocloseTime property to a value in seconds; when
'  you get the property, it returns the current time remaining.
'--Events for custom controls: ComboItemChanged, ComboDropdown, InputBoxChange, DateTimeChange
'--Custom controls now have the option to manually set their width (and height for combo).
'--Custom controls in the Buttons position now have their width adjusted to fill the space.
'  Custom controls in the Footer position still use a standard size, however for InputBox,
'  ComboBox, and Slider if you specify a width of -1 the width will scale to the full width
'--For custom controls in the footer area, and the datetime control in the content area, there's
'  an additional alignment option to choose between left, center, and right. In the footer area
'  this could be used to show footer text and have the control off to the right.
'--Focus is no longer always on an InputBox: Focus is set on whatever custom control is in the
'  content area, if none button area, last footer area. Or, there's now a .DefaultCustomControl
'  option to manually specify which one, or not have focus set on any of them.
'--Bug fix: Several fixes relating to alignment and events for custom controls used on dialogs
'           that are loaded as a new page (NavigatePage/TDN_NAVIGATED)
'--Bug fix: .DefaultButton used TDBUTTONS instead of TDRESULT enum to identify IDs
'--Bug fix: .DefaultRadioButton was incorrectly associated with the TDBUTTONS enum
'--Alignment: -Adjusted Content area custom controls slightly higher when expanded
'              info is used.
'             -Custom controls in content area now have an appropriate X-position and width when
'              no icon appears, shifting everything to the left.
'             -Made sure that content text actually had a link before making the link adjustment,
'              since the flag would still be used if there's a link only in the footer area.
'             -When there's a link but no command links or radio buttons, content area controls
'              needed to be adjusted upwards a bit.
'--Alignment: -Custom controls size and position is now automatically adjust for the current DPI.
'             -You can also specify a scale factor manually with the DPIScaleX/DPIScaleY properties.
'KNOWN ISSUES: -If there's a custom control in the content area, when an expando control is expanded
'               and then collapsed, the excess whitespace isn't removed despite there not being any
'               extra vbCrLf's to account for it. No way to correct this has been discovered yet.
'              -If the dialog has to be resized in response to an expando control, the width is
'               reduced by several pixels for an unknown reason- but only on the first time. If the
'               numbers returned by getting the client RECT were wrong, you'd expect to lose more
'               pixels every time, but that's not the case. No solution has been discovered.
'              -Some alignment scenarios remain unsupported. If expando with the expand-footer w/
'               expand-default, have text with multiple lines, alignment will be off. This remains
'               a significant challenge to address as the DirectUIHWND reports incorrect information
'               to GetPixel, so figuring out where things are will require font height analysis and
'               reverse engineering line break length determination. Will work on in next version.
'              -Sliders with ticks on top and bottom simply don't fit on higher DPIs in the footer
'               position. Height available is too small no matter how it's positioned.
'----------------------------------------------------------------------------------------------------
'NEW IN VERSION 0.8
'--Added ability to use comboxes (normal edit dropdown and dropdown list), date/time pickers,
'  and sliders like the inputbox controls.
'--There's still the 3 alignment positions from the inputbox, so you can have controls in each
'  of these places (however, only 1 of each type is allowed)
'--When TDF_KILL_SHIELD_ICON is used, there's now an option to replace it with a different icon:
'   Set .IconReplaceGradient...
'    1) To an ID in your app's resource file (only works when compiled)
'    2) Add the TDF_USE_SHELL32_ICONID or TDF_USE_IMAGERES_ICONID flag and use an ID from those
'    3) Set .hInst to a custom module and use an icon ID from that
'   Using an hIcon is not supported at this time.
'--Bug fix: Using TDF_KILL_SHIELD_ICON and calling .Init crashed the app
'--Bug fix: Using an InputBox at the footer position, the bottom half was cut off if no
'           footer icon was set.
'--Bug fix: Input box position adjusted for expando only if in the content position, now adjusts
'           if in footer position too. Prevented this bug in all other controls.
'--Bug fix: If input box was in footer position, and expanded info was expanded automatically
'           and into the footer area, the input box would be in the wrong position. Fixed and
'           prevented for new controls.
'KNOWN ISSUE: If the slider control is placed in the footer position and an expando control is
'             set to expand into the footer area, the transparent background on the slider is
'             lost (turns white) until it's clicked. It's a very unusual and specific situation,
'             so not going to delay this version while I come up with a fix.
'TODO:
'The next version will be 1.0. There's a couple major features planned, including images and more
'events (for the custom controls). Maybe more controls; like dropdown menus on the buttons.
'I'm very aware this class needs substantial cleanup, that's also being saved for 1.0
' So please go easy on the sloppiness!!! :)
'
'---------------------------------------------------------------------------------------------------
'NEW IN VERSION 0.7
'--Input boxes! See details below for the TDF_INPUT_BOX custom flag.
'--Added flag TDF_KILL_SHIELD_ICON - kill it but keep that colored background.
'--Fixed bug where updating the footer icon in an open dialog instead updated the main icon
'NEW IN VERSION 0.6.1
'--Can't believe I forgot the TDF_SIZE_TO_CONTENT flag.
'NEW IN VERSION 0.6
'--Support for custom icons on the buttons themselves
'0.5.2: critical bugfixes
'NEW IN VERSION 0.5.1
'--Added a force case function for the enums; they'll be automatically cased like normal now
'  while in the IDE
'--Bugfix: If a dialog was called as a new page, setting holds/disabled/elevated automatically
'          did not work.
'--Bugfix: When navigating to a new page, the old hWnd was not cleared, leading to some calls
'          not working properly.
'--Customized the progressbarstate enum since it's likely to conflict; if you use a PBST_ var,
'  it needs to be changed to ePBST, if it wasn't publicly declared outside this class
'NEW IN VERSION 0.5
'--Added support for multiple pages
'NEW IN VERSION 0.4.1
'--Added queueing for EnableButton, EnableRadioButton, and SetButtonElevated: now you can
'  disable or elevate a button before the dialog is shown and it will be done automatically
'NEW IN VERSION 0.4
'--Added support for TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE; which adds the security shield
'  to a button indicated elevated user permission is required to perform the action
'--Added a SimpleDialog function to show the original TaskDialog() for very simple messages;
'  it can be a drop-in replacement for MsgBox, taking the common Message,Buttons,Title in
'  the same order
'--Set default dialog title to App.Title. If NULL is passed, it shows the exe name rather
'  than a blank string; and vb6.exe when in the IDE.
'--Misc minor bugfixes
'NEW IN VERSION 0.3.2/3
'Some internal improvements suggested by Bonnie West on VBForums. Also minor bugfixes.
'NEW IN VERSION 0.3.1
'Bugfixes; now use LoadLibrary instead of GetModuleHandle, which may fail in some circumstances
'NEW IN VERSION 0.3
'Automated the process of not closing the dialog when a button is clicked
'Uses .SetButtonHold and .ReleaseButtonHold
'NEW IN VERSION 0.2
'Schmidt gave me the brilliant idea to make using icons from shell32.dll and imageres.dll
'possible simply by setting .hInstance to their handle and specifying the index.
'See the new custom flags, and the readme for full details


'*******************************************************************************
'USER OPTIONS : Configuration for developers using this class.
Private Const useropt_dbg_PrintToImmediate As Boolean = False
Private Const useropt_dbg_MinLevel As Long = 0&
Private Const useropt_dbg_IncludeDateTime As Boolean = True
'********************************************************************************



Private Const WM_USER = &H400&
Private Const S_OK = 0
Private Const S_FALSE = 1&
Private Const CSIDL_SYSTEM = &H25
Private Const IDD_EDIT1 = 101&
Private Const IDD_DT1 = 102&
Private Const IDD_DT2 = 103&
Private Const IDD_COMBO = 104&
Private Const IDD_SLIDER = 105&

Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_LEFT As Long = &H0&
Private Const WS_EX_LTRREADING              As Long = &H0&
Private Const WS_EX_RIGHTSCROLLBAR          As Long = &H0&
Private Const WS_EX_TRANSPARENT = &H20
Private Const WS_EX_LAYERED = &H80000

Private Const DEFAULT_GUI_FONT = 17
Private Const WM_SETFONT                    As Long = &H30
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETTEXT = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COMMAND = &H111
Private Const WM_NOTIFY = &H4E
Private Const WM_HSCROLL = &H114
Private Const WM_DESTROY = &H2
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const BM_SETIMAGE = &HF7
Private Const EM_SETSEL As Long = &HB1
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const ES_LEFT = &H0
Private Const ES_CENTER = &H1
Private Const ES_RIGHT = &H2
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_PASSWORD As Long = &H20
Private Const BS_SPLITBUTTON As Long = &HC
Private Const CBN_SELCHANGE = 1
Private Const CBN_DROPDOWN = 7
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const EN_CHANGE = &H300
Private Const EN_UPDATE = &H400
Private Const BCN_FIRST As Long = -1250
Private Const BCN_DROPDOWN As Long = (BCN_FIRST + 2)
Private Const H_MAX As Long = &HFFFF + 1
Private Const DTN_FIRST = (H_MAX - 760&)
Private Const DTN_LAST = (H_MAX - 799&)
Private Const DTN_DATETIMECHANGE = (DTN_FIRST + 1)
Private Const TRBN_FIRST = -1501
Private Const TRBN_THUMBPOSCHANGING = TRBN_FIRST - 1
Private Const TBS_HORZ As Long = &H0
Private Const TBS_BOTH As Long = &H8
Private Const TBS_BOTTOM As Long = &H0
Private Const TBS_TOP As Long = &H4
Private Const TBS_NOTICKS As Long = &H10
Private Const TBS_AUTOTICKS As Long = &H1
Private Const TBS_TRANSPARENTBKGND = &H1000
Private Const TBM_SETRANGEMIN As Long = (WM_USER + 7)
Private Const TBM_SETRANGEMAX As Long = (WM_USER + 8)
Private Const TBM_SETTICFREQ As Long = (WM_USER + 20)
Private Const TBM_SETLINESIZE As Long = (WM_USER + 23)
Private Const TBM_SETPAGESIZE As Long = (WM_USER + 21)
Private Const TBM_GETPOS As Long = (WM_USER)
Private Const TBM_SETPOS As Long = (WM_USER + 5)
Private Const BCM_FIRST = &H1600
Private Const BCM_SETDROPDOWNSTATE = (BCM_FIRST + &H6)
Private Const DTM_FIRST                     As Long = &H1000
Private Const DTM_GETSYSTEMTIME             As Long = (DTM_FIRST + 1)
Private Const DTM_SETSYSTEMTIME             As Long = (DTM_FIRST + 2)
Private Const DTM_SETRANGE                  As Long = (DTM_FIRST + 4)
Private Const STM_SETICON = &H170
Private Const STM_GETICON = &H171
Private Const STM_SETIMAGE = &H172
Private Const STM_GETIMAGE = &H173
Private Const SS_LEFT = &H0
Private Const SS_CENTER = &H1
Private Const SS_RIGHT = &H2
Private Const SS_ICON = &H3
Private Const SS_BLACKRECT = &H4
Private Const SS_GRAYRECT = &H5
Private Const SS_WHITERECT = &H6
Private Const SS_BLACKFRAME = &H7
Private Const SS_GRAYFRAME = &H8
Private Const SS_WHITEFRAME = &H9
Private Const SS_SIMPLE = &HB
Private Const SS_LEFTNOWORDWRAP = &HC
Private Const SS_BITMAP = &HE
Private Const SS_ENHMETAFILE = &HF
Private Const SS_ETCHEDHORZ = &H10
Private Const SS_ETCHEDVERT = &H11
Private Const SS_ETCHEDFRAME = &H12
Private Const SS_REALSIZECONTROL = &H40
Private Const SS_NOPREFIX = &H80
Private Const SS_NOTIFY = &H100
Private Const SS_CENTERIMAGE = &H200
Private Const SS_RIGHTJUST = &H400
Private Const SS_SUNKEN = &H1000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type NMHDR
  hWndFrom As LongPtr   ' Window handle of control sending message
  idFrom As LongPtr      ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type
Private Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
  IMAGE_ENHMETAFILE = 3
End Enum
Private Type BITMAP
    BMType As Long
    BMWidth As Long
    BMHeight As Long
    BMWidthBytes As Long
    BMPlanes As Integer
    BMBitsPixel As Integer
    BMBits As LongPtr
End Type
Private Enum WinStyles
  WS_OVERLAPPED = &H0
  WS_TABSTOP = &H10000
  WS_MAXIMIZEBOX = &H10000
  WS_MINIMIZEBOX = &H20000
  WS_GROUP = &H20000
  WS_THICKFRAME = &H40000
  WS_SYSMENU = &H80000
  WS_HSCROLL = &H100000
  WS_VSCROLL = &H200000
  WS_DLGFRAME = &H400000
  WS_BORDER = &H800000
  WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
  WS_MAXIMIZE = &H1000000
  WS_CLIPCHILDREN = &H2000000
  WS_CLIPSIBLINGS = &H4000000
  WS_DISABLED = &H8000000
  WS_VISIBLE = &H10000000
  WS_MINIMIZE = &H20000000
  WS_CHILD = &H40000000
  WS_POPUP = &H80000000
  
  WS_TILED = WS_OVERLAPPED
  WS_ICONIC = WS_MINIMIZE
  WS_SIZEBOX = WS_THICKFRAME
  
  ' Common Window Styles
  WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
  WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
  WS_CHILDWINDOW = WS_CHILD
End Enum
Private Enum DTSTYLES
    DTS_SHORTDATEFORMAT = &H0
    DTS_UPDOWN = &H1
    DTS_SHOWNONE = &H2
    DTS_LONGDATEFORMAT = &H4
    DTS_TIMEFORMAT = &H9
    DTS_APPCANPARSE = &H10
    DTS_RIGHTALIGN = &H20
    DTS_SHORTDATECENTURYFORMAT = &HC
End Enum

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const GDT_ERROR                     As Long = -1
Private Const GDT_VALID                     As Long = 0
Private Const GDT_NONE                      As Long = 1
Private Const GDTR_MAX                      As Long = &H2
Private Const GDTR_MIN                      As Long = &H1


Public Enum SldTickStyle
    SldTickStyleBottom = 0
    SldTickStyleTop = 1
    SldTickStyleBoth = 2
    SldTickStyleNone = 3
End Enum
Private nSliderTickStyle As SldTickStyle


Public Enum eProgressBarStates
    ePBST_NORMAL = 1
    ePBST_ERROR = 2
    ePBST_PAUSED = 3
End Enum

Public Enum TDInputBoxAlign
    TDIBA_Content = 0
    TDIBA_Buttons = 1
    TDIBA_Footer = 2
End Enum
Public Enum TDInputBoxTextAlign
    TDIBTA_Left = 0
    TDIBTA_Right = 1
    TDIBTA_Center = 2
End Enum

Private Enum eShowWindowTypes
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
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20

Private Const WC_COMBOBOXEX = "ComboBoxEx32"
Private Const CBEM_SETIMAGELIST = (WM_USER + 2)
Private Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL = (WM_USER + 7)
Private Const CBEM_SETUNICODEFORMAT = 8192 + 5
Private Const CBEM_INSERTITEMW = (WM_USER + 11)
Private Const CBEM_SETITEMW = (WM_USER + 12)
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBS_AUTOHSCROLL = &H40
Private Const CB_GETCURSEL = &H147
Private Const CB_SETCURSEL = &H14E

Private Type COMBOBOXEXITEMW
    Mask As COMBOBOXEXITEM_Mask
    iItem As LongPtr
    pszText As LongPtr      '// LPCSTR
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    iOverlay As Long
    iIndent As Long
    lParam As LongPtr
End Type
Private Enum COMBOBOXEXITEM_Mask
    CBEIF_TEXT = &H1
    CBEIF_IMAGE = &H2
    CBEIF_SELECTEDIMAGE = &H4
    CBEIF_OVERLAY = &H8
    CBEIF_INDENT = &H10
    CBEIF_LPARAM = &H20
    CBEIF_DI_SETITEM = &H10000000
End Enum


Public Enum TDICONS
    TD_WARNING_ICON = -1 'exclamation point in a yellow 'yield' triangle (same image as IDI_EXCLAMATION)
    TD_ERROR_ICON = -2 'round red circle containg 'X' (same as IDI_HAND)
    TD_INFORMATION_ICON = -3 'round blue circle containing 'i' (same image as IDI_ASTERISK)
    TD_SHIELD_ICON = -4 'Vista's security shield
    IDI_APPLICATION = 32512&
    IDI_ERROR = 32513&
    'miniature picture of an application window
    IDI_QUESTION = 32514& 'round blue circle containing '?'
    IDI_WINLOGO = 32517&
    TD_SHIELD_GRADIENT_ICON = -5 'same image as TD_SHIELD_ICON; main message text on gradient blue background
    TD_SHIELD_WARNING_ICON = -6 'exclamation point in yellow Shield shape; main message text on gradient orange background
    TD_SHIELD_ERROR_ICON = -7 'X contained within Shield shape; main message text on gradient red background
    TD_SHIELD_OK_ICON = -8 'Shield shape containing green checkmark; main message text on gradient green background
    TD_SHIELD_GRAY_ICON = -9 'same image as TD_SHIELD_ICON; main message text on medium gray background
    TD_NO_ICON = 0 'no icon; text on white background
End Enum

'taskdialog common button flags
Public Enum TDBUTTONS
    TDCBF_OK_BUTTON = &H1&      'return value 1 (IDOK)
    TDCBF_YES_BUTTON = &H2&     'return value 6 (IDYES)
    TDCBF_NO_BUTTON = &H4&      'return value 7 (IDNO)
    TDCBF_CANCEL_BUTTON = &H8&  'return value 2 (IDCANCEL)
    TDCBF_RETRY_BUTTON = &H10&   'return value 4 (IDRETRY)
    TDCBF_CLOSE_BUTTON = &H20&   'return value 8 (IDCLOSE)
End Enum
Public Enum TDRESULT
    TD_OK = 1
    TD_YES = 6
    TD_NO = 7
    TD_CANCEL = 2
    TD_RETRY = 4
    TD_CLOSE = 8
End Enum
Public Enum TASKDIALOG_FLAGS
    TDF_ENABLE_HYPERLINKS = &H1
    TDF_USE_HICON_MAIN = &H2
    TDF_USE_HICON_FOOTER = &H4
    TDF_ALLOW_DIALOG_CANCELLATION = &H8
    TDF_USE_COMMAND_LINKS = &H10
    TDF_USE_COMMAND_LINKS_NO_ICON = &H20
    TDF_EXPAND_FOOTER_AREA = &H40
    TDF_EXPANDED_BY_DEFAULT = &H80
    TDF_VERIFICATION_FLAG_CHECKED = &H100
    TDF_SHOW_PROGRESS_BAR = &H200
    TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400
    TDF_CALLBACK_TIMER = &H800
    TDF_POSITION_RELATIVE_TO_WINDOW = &H1000
    TDF_RTL_LAYOUT = &H2000
    TDF_NO_DEFAULT_RADIO_BUTTON = &H4000
    TDF_CAN_BE_MINIMIZED = &H8000&
    TDF_NO_SET_FOREGROUND = &H10000
    TDF_SIZE_TO_CONTENT = &H1000000
    TDF_USE_SHELL32_ICONID = &H400000 'CUSTOM FLAG
    TDF_USE_IMAGERES_ICONID = &H800000 'CUSTOM FLAG
    TDF_DATETIME = &H80000         'CUSTOM FLAG: Add calendar control
    TDF_SLIDER = &H2000000            'CUSTOM FLAG: Add Slider control
    TDF_INPUT_BOX = &H4000000        'CUSTOM FLAG: Input box. See below.
    TDF_COMBO_BOX = &H8000000               'CUSTOM FLAG: Combo box; .ComboType controls edit or list
    TDF_KILL_SHIELD_ICON = &H10000000 'CUSTOM FLAG: Allows the colored background from the TD_SHIELD_x icons to be used with just the text and no shield icon
    TDF_EXEC_HYPERLINKS = &H20000000           'CUSTOM FLAG
    TDF_USE_SHELL32_ICONID_BUTTON = &H40000000 'CUSTOM FLAG
    TDF_USE_IMAGERES_ICONID_BUTTON = &H80000000 'CUSTOM FLAG
End Enum
'TDF_INPUT_BOX:
'Creates a textbox for input. There are multiple ways to set its position:
'1) Default, at the bottom of the content in the whitespace before the gray area
'2) To the left of the buttons. Not compatible with the expanded-info or verify-checkbox styles.
'3) As a footer. Aligned to retain the footer icon, but have an inputbox next to it.
'
'Additional inputbox items:
'-Initial text can be set, and current text from an open dialog, with .InputText
'-Cue banner can be set, .InputCueBanner
'-Input box can be set as a password input with * masks, .InputIsPassword


Public Enum TASKDIALOG_MESSAGES
    TDM_NAVIGATE_PAGE = WM_USER + 101&
    TDM_CLICK_BUTTON = WM_USER + 102&                  '// wParam = Button ID
    TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103&      '// wParam = 0 (nonMarque) wParam != 0 (Marquee)
    TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104&        '// wParam = new progress state
    TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105&        '// lParam = tdMAKELPARAM(nMinRange& nMaxRange)
    TDM_SET_PROGRESS_BAR_POS = WM_USER + 106&          '// wParam = new position
    TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107&      '// wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds between repaints)
    TDM_SET_ELEMENT_TEXT = WM_USER + 108&              '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    TDM_CLICK_RADIO_BUTTON = WM_USER + 110&            '// wParam = Radio Button ID
    TDM_ENABLE_BUTTON = WM_USER + 111&                 '// lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
    TDM_ENABLE_RADIO_BUTTON = WM_USER + 112&           '// lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
    TDM_CLICK_VERIFICATION = WM_USER + 113&            '// wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
    TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114&           '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115& '// wParam = Button ID, lParam = 0 (elevation not required)& lParam != 0 (elevation required)
    TDM_UPDATE_ICON = WM_USER + 116&                   '// wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
End Enum 'TASKDIALOG_MESSAGES;

Public Enum TASKDIALOG_NOTIFICATIONS
    TDN_CREATED = 0
    TDN_NAVIGATED = 1
    TDN_BUTTON_CLICKED = 2           '// wParam = Button ID
    TDN_HYPERLINK_CLICKED = 3        '// lParam = (LPCWSTR)pszHREF
    TDN_TIMER = 4                    '// wParam = Milliseconds since dialog created or timer reset
    TDN_DESTROYED = 5
    TDN_RADIO_BUTTON_CLICKED = 6     '// wParam = Radio Button ID
    TDN_DIALOG_CONSTRUCTED = 7
    TDN_VERIFICATION_CLICKED = 8     '// wParam = 1 if checkbox checked, 0 if not, lParam is unused and always 0
    TDN_HELP = 9
    TDN_EXPANDO_BUTTON_CLICKED = 10  '// wParam = 0 (dialog is now collapsed), wParam != 0 (dialog is now expanded)
End Enum 'TASKDIALOG_NOTIFICATIONS;

Public Enum TASKDIALOG_ELEMENTS
    TDE_CONTENT = 0
    TDE_EXPANDED_INFORMATION = 1
    TDE_FOOTER = 2
    TDE_MAIN_INSTRUCTION = 3
End Enum

Public Enum TASKDIALOG_ICON_ELEMENTS
    TDIE_ICON_MAIN = 0
    TDIE_ICON_FOOTER = 1
End Enum
[ PackingAlignment (1) ]
Private Type TASKDIALOG_BUTTON
    nButtonID As Long
    pszButtonText As LongPtr
End Type
[ PackingAlignment (1) ]
Private Type TASKDIALOGCONFIG
    cbSize As Long
    hWndParent As LongPtr
    hInstance As LongPtr
    dwFlags As TASKDIALOG_FLAGS
    dwCommonButtons As TDBUTTONS
    pszWindowTitle As LongPtr
    'hMainIcon As Long
    pszMainIcon As LongPtr 'TDICONS
    pszMainInstruction As LongPtr
    pszContent As LongPtr
    cButtons As Long
    pButtons As LongPtr
    nDefaultButton As Long
    cRadioButtons As Long
    pRadioButtons As LongPtr
    nDefaultRadioButton As Long
    pszVerificationText As LongPtr
    pszExpandedInformation As LongPtr
    pszExpandedControlText As LongPtr
    pszCollapsedControlText As LongPtr
    pszFooterIcon As LongPtr 'TDICONS
    'hFooterIcon As Long
    pszFooter As LongPtr
    pfCallback As LongPtr
    lpCallbackData As LongPtr
    CXWidth As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function TaskDialog Lib "comctl32.dll" _
                                                    (ByVal hWndParent As LongPtr, _
                                                     ByVal hInstance As LongPtr, _
                                                     ByVal pszWindowTitle As LongPtr, _
                                                     ByVal pszMainInstruction As LongPtr, _
                                                     ByVal pszContent As LongPtr, _
                                                     ByVal dwCommonButtons As Long, _
                                                     ByVal pszIcon As LongPtr, _
                                                     pnButton As Long) As Long
                                                     
Private Declare PtrSafe Function TaskDialogIndirect Lib "comctl32.dll" (pTaskConfig As TASKDIALOGCONFIG, _
                                                                pnButton As Long, _
                                                                pnRadioButton As Long, _
                                                                pfVerificationFlagChecked As Long) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As LongPtr, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal lpOperation As LongPtr, ByVal lpFile As LongPtr, ByVal lpParameters As LongPtr, ByVal lpDirectory As LongPtr, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hwndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As Any) As Long  ' lpPoint As POINT) As Long
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
                                                                                                                                                                                     
#Else
Private Declare Function TaskDialog Lib "comctl32.dll" _
                                                    (ByVal hWndParent As Long, _
                                                     ByVal hInstance As Long, _
                                                     ByVal pszWindowTitle As Long, _
                                                     ByVal pszMainInstruction As Long, _
                                                     ByVal pszContent As Long, _
                                                     ByVal dwCommonButtons As Long, _
                                                     ByVal pszIcon As Long, _
                                                     pnButton As Long) As Long
'WINCOMMCTRLAPI HRESULT WINAPI TaskDialogIndirect(const TASKDIALOGCONFIG *pTaskConfig, __out_opt int *pnButton, __out_opt int *pnRadioButton, __out_opt BOOL *pfVerificationFlagChecked);
Private Declare Function TaskDialogIndirect Lib "comctl32.dll" (pTaskConfig As TASKDIALOGCONFIG, _
                                                                pnButton As Long, _
                                                                pnRadioButton As Long, _
                                                                pfVerificationFlagChecked As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   lParam As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
' Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                        ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
#End If

Private Const CP_ACP = 0        ' ANSI code page

Private Enum GWL_nIndex
  GWL_WNDPROC = (-4)
  GWL_HINSTANCE = (-6)
  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
  GWL_USERDATA = (-21)
End Enum


Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private m_ScaleX As Single, m_ScaleY As Single
Public lngTimerID As Long


Public Event DialogCreated(ByVal hWnd As LongPtr)
Public Event ButtonClick(ByVal ButtonID As Long)
Public Event HyperlinkClick(ByVal lPtr As LongPtr)
Public Event Timer(ByVal TimerValue As Long)
Public Event DialogDestroyed()
Public Event RadioButtonClick(ByVal ButtonID As Long)
Public Event DialogConstucted(ByVal hWnd As LongPtr)
Public Event VerificationClicked(ByVal Value As Long)
Public Event ExpandButtonClicked(ByVal Value As Long)
Public Event Navigated()
Public Event Help()
'----------------------
'Custom Events:
Public Event DropdownButtonClicked(ByVal hWnd As LongPtr)
Public Event ComboItemChanged(ByVal iNewItem As Long)
Public Event ComboDropdown()
Public Event InputBoxChange(sText As String)
Public Event DateTimeChange(ByVal dtNew As Date, ByVal lCheckStatus As Long)
Public Event SliderChange(ByVal lNewValue As Long)
Public Event AutoClose()

Private m_TDhWnd As LongPtr
Private m_TDhWndTmp As LongPtr
Private mPageIndex As Long
Private bResetTimer As Boolean
Private lBtnHold() As Long
Private lBtnElv() As Long
Private bElv As Boolean
Private lBtnDis() As Long
Private bDis As Boolean
Private lRadDis() As Long
Private bRadDis As Boolean
Private btnCaps() As String 'weird things happen if these are set but not stored
Private rbtnCaps() As String
Private icoDataC() As LongPtr 'icon data for custom buttons
Private icoDataN(5) As LongPtr 'icon data for normal buttons
Private bAutoclose As Boolean
Private lAutoclose As Long 'in seconds
Private lAutocloseTrack As Long
Private dtAC1 As Date, dtAC2 As Date
Private bACTmrSup As Boolean
Private hEditBox As LongPtr
Private hEditCombo As LongPtr
Private hComboCombo As LongPtr
Private hStatic As LongPtr
Private hLogoImage As LongPtr
Private nLogoType As LogoImageType
Private nLogoPos As LogoPosition
Private lLogoX As Long, lLogoY As Long
Private hSlider As LongPtr
Private lSliderMin As Long
Private lSliderMax As Long
Private lSliderValue As Long
Private lSliderNTicks As Long
Private lSliderChangeS As Long
Private lSliderChangeL As Long
Private hUI As LongPtr
Private sEditText As String
Private sComboText As String
Private nComboIdx As Long
Private lButtonY As Long
Private lButtonX As Long
Private lButtonY2 As Long
Private lButtonX2 As Long
Private lButtonY3 As Long
Private lButtonX3 As Long
Private lButtonX4 As Long
Private lButtonY4 As Long
Private arBtnX() As Long
Private bLinkAdjust As Boolean
Private bEditFooter As Boolean
Private bEditPassword As Boolean
Private bReposTrigger As Boolean
Private bPublicTimer As Boolean
Private tDTCur As SYSTEMTIME
Private tDTCur2 As SYSTEMTIME
Private bDTCheck As Boolean
Private bDTCheck2 As Boolean
Private lLastCY As Long
Private nTmrCnt As Long
Private himlCombo As LongPtr
Private nSplitButtonID As Long
Private m_sTitle As String
Private m_sMainInstr As String
Private m_sContent As String
Private m_sFooter As String
Private m_sExpandedInfo As String
Private m_sExpControl As String
Private m_sColControl As String
Private m_sVerify As String
Private m_sInput As String
Private m_sComboCue As String
Private m_uButtons() As TASKDIALOG_BUTTON
Private m_uRadioButtons() As TASKDIALOG_BUTTON
Private m_Icon As LongPtr, m_IconFooter As LongPtr
Private m_HasCustomButtons As Boolean, m_HasRadioButtons As Boolean
Private m_hButton() As LongPtr
Private m_BtnLRSize As Long
Private nInputAlign As TDInputBoxAlign
Private nComboAlign As TDInputBoxAlign
Private mInputTextAlign As TDInputBoxTextAlign
Private nDateTimeAlign As TDInputBoxAlign
Private nSliderAlign As TDInputBoxAlign
Private cxInputBox As Long
Private cxSlider As Long
Private cxCombo As Long
Private cyCombo As Long
Private cxCBDrop As Long
Private DTOffX As Long
Private CBOffX As Long
Private SLOffX As Long
Private IBOffX As Long
Private hCombo As LongPtr
Private hDateTime As LongPtr, hDateTime2 As LongPtr
Private bPosUsed(2) As Boolean
Private m_HasCustomControl As Boolean
Private m_lNewIcon As LongPtr
Private nBreaksAdded As Integer
Private uTDC As TASKDIALOGCONFIG

Private hShell32 As LongPtr, hImageRes As LongPtr
Private lgpX As Long, lgpY As Long

Public Enum TDCustomControlType
    tdccInputBox = 1
    tdccComboBox = 2
    tdccDateTime = 3
    tdccSlider = 4
    tdccNothing = 99
End Enum
Private nDefaultCustCtrl As TDCustomControlType

Public Enum TDControlAlign
    tdcaLeft = 0
    tdcaCenter = 1
    tdcaRight = 2
End Enum
Private nDTContentAlign As TDControlAlign
Private nDTFooterAlign As TDControlAlign
Private nDTButtonAlign As TDControlAlign
Private nCBFooterAlign As TDControlAlign
Private nIBFooterAlign As TDControlAlign
Private nSLFooterAlign As TDControlAlign

Public Enum LogoImageType
    LogoBitmap = 0
    LogoIcon = 1
End Enum
Public Enum LogoPosition
    LogoTopRight = 0
    LogoButtons = 1
End Enum
Public Enum DateTimeType
    dttDate = 0
    dttDateWithCheck = 1
    dttTime = 2
    dttTimeWithCheck = 3
    dttDateTime = 4
    dttDateTimeWithCheck = 5
    dttDateTimeWithCheckTimeOnly = 6
    dttDateTimeWithCheckDateOnly = 7
End Enum
Private nDTType As DateTimeType

Private tDTInitial As SYSTEMTIME
Private bDTSetInit As Boolean

Private tDTRange(1) As SYSTEMTIME
Private bDTSetMin As Boolean
Private bDTSetMax As Boolean

Public Enum ComboType
    cbtDropdownEdit = 0
    cbtDropdownList = 1
End Enum
Private nCbType As ComboType

Private Type ctdComboItem
    sText As String
    iImage As Long
    iOverlay As Long
End Type
Private aComboItems() As ctdComboItem
Private tComboInit As ctdComboItem
Private nComboInit As Long
Private bCBInitValid As Boolean
Public Enum eWinIcoSize
    ICO_16 = 16
    ICO_24 = 24
    ICO_32 = 32
    ICO_48 = 48
    ICO_64 = 64
End Enum

Private m_ResMain As Long, m_ResRad As Long, m_ResVerify As Long


Private m_cbMain As Long
Private m_cbEnum As Long

'***************************************************************************
'
'----------------------------END DECLARE SECTION---------------------------
'
'***************************************************************************


Private Function GetStrFromPtrW(lpszW As LongPtr) As String
  Dim sRtn As String
  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
'  sRtn = String$(WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, 0, 0, 0, 0), 0)
  Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
  GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function
Private Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function


Private Sub DebugAppend(sMsg As String, Optional iLvl As Long)
If useropt_dbg_PrintToImmediate = False Then Exit Sub
If useropt_dbg_IncludeDateTime Then sMsg = "[" & Format$(Now, "yyyy-mm-dd Hh:nn:Ss") & "] " & sMsg
If iLvl >= useropt_dbg_MinLevel Then Debug.Print sMsg
End Sub



Public Property Let PageIndex(nPage As Long): mPageIndex = nPage: End Property
Public Property Get PageIndex() As Long: PageIndex = mPageIndex: End Property

Public Property Let ParenthWnd(hWnd As LongPtr): uTDC.hWndParent = hWnd: End Property
Public Property Get ParenthWnd() As LongPtr: ParenthWnd = uTDC.hWndParent: End Property

Public Property Let DPIScaleX(Value As Single): m_ScaleX = Value: End Property
Public Property Get DPIScaleX() As Single: DPIScaleX = m_ScaleX: End Property

Public Property Let DPIScaleY(Value As Single): m_ScaleY = Value: End Property
Public Property Get DPIScaleY() As Single: DPIScaleY = m_ScaleY: End Property

Public Property Let AutocloseTime(Value As Long): lAutocloseTrack = Value: End Property
Public Property Get AutocloseTime() As Long: AutocloseTime = lAutocloseTrack: End Property

Public Property Let hinst(hInstance As LongPtr): uTDC.hInstance = hInstance: End Property
Public Property Get hinst() As LongPtr: hinst = uTDC.hInstance: End Property

Public Property Get Flags() As TASKDIALOG_FLAGS: Flags = uTDC.dwFlags: End Property
Public Property Let Flags(dwFlags As TASKDIALOG_FLAGS): uTDC.dwFlags = dwFlags: End Property

Public Property Get CommonButtons() As TDBUTTONS: CommonButtons = uTDC.dwCommonButtons: End Property
Public Property Let CommonButtons(dwBtns As TDBUTTONS): uTDC.dwCommonButtons = dwBtns: End Property

Public Property Get Title() As String: Title = m_sTitle: End Property
Public Property Let Title(sTitle As String)
m_sTitle = sTitle
uTDC.pszWindowTitle = StrPtr(m_sTitle)
End Property

Public Property Get IconMain() As LongPtr: IconMain = m_Icon: End Property
Public Property Let IconMain(lIcon As LongPtr)
m_Icon = lIcon

If m_TDhWnd <> 0& Then
    If (uTDC.dwFlags And TDF_USE_HICON_MAIN) Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_Icon)
    Else
        Dim dwIcon As Long
        dwIcon = tdMakeIntResource(CLng(m_Icon))
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal dwIcon)
    End If
    'Sometimes the icon wasn't redrawing, repainting the window fixed it
    Dim hFrame As LongPtr
    hFrame = FindWindowEx(m_TDhWnd, 0&, "DirectUIHWND", vbNullString)
    RepaintWindow hFrame
End If
End Property

Public Property Get IconReplaceGradient() As LongPtr: IconReplaceGradient = m_lNewIcon: End Property
Public Property Let IconReplaceGradient(nIndex As LongPtr)
m_lNewIcon = nIndex
If m_TDhWnd <> 0& Then
    If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) Then
        SendMessageW m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_lNewIcon
    End If
End If
End Property

Public Property Get MainInstruction() As String: MainInstruction = m_sMainInstr: End Property
Public Property Let MainInstruction(sText As String)
m_sMainInstr = sText
uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_MAIN_INSTRUCTION, ByVal StrPtr(m_sMainInstr))
End If
End Property

Public Property Get Content() As String: Content = m_sContent: End Property
Public Property Let Content(sText As String)
m_sContent = sText
If (uTDC.dwFlags And TDF_INPUT_BOX) Then
    If nInputAlign = TDIBA_Content Then
        If Right$(m_sContent, 2) <> vbCrLf Then
            m_sContent = m_sContent & vbCrLf
        End If
    End If
End If
uTDC.pszContent = StrPtr(m_sContent)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_CONTENT, ByVal StrPtr(m_sContent))
End If
End Property

Public Property Get DefaultButton() As TDRESULT: DefaultButton = uTDC.nDefaultButton: End Property
Public Property Let DefaultButton(ButtonID As TDRESULT): uTDC.nDefaultButton = ButtonID: End Property

Public Property Get DefaultRadioButton() As Long: DefaultButton = uTDC.nDefaultRadioButton: End Property
Public Property Let DefaultRadioButton(ButtonID As Long): uTDC.nDefaultRadioButton = ButtonID: End Property

Public Property Get DefaultCustomControl() As TDCustomControlType: DefaultCustomControl = nDefaultCustCtrl: End Property
Public Property Let DefaultCustomControl(iControl As TDCustomControlType): nDefaultCustCtrl = iControl: End Property

Public Property Get VerifyText() As String: VerifyText = m_sVerify: End Property
Public Property Let VerifyText(sText As String)
m_sVerify = sText
uTDC.pszVerificationText = StrPtr(m_sVerify)
End Property

Public Property Get InputText() As String
    If hEditBox Then
            Dim lLen As Long, sText As String
            lLen = CLng(SendMessageW(hEditBox, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2)
            If lLen Then
                sText = Space$(lLen)
                Call SendMessageW(hEditBox, WM_GETTEXT, lLen, ByVal sText)
                sEditText = StrConv(sText, vbFromUnicode)
            End If
    End If
    InputText = sEditText
End Property
Public Property Let InputText(sText As String)
sEditText = sText
If hEditBox Then
    SendMessageW hEditBox, WM_SETTEXT, 0&, ByVal StrPtr(sEditText)
End If
End Property
Public Property Get InputBoxTextAlign() As TDInputBoxTextAlign: InputBoxTextAlign = mInputTextAlign: End Property
Public Property Let InputBoxTextAlign(nAlign As TDInputBoxTextAlign)
If mInputTextAlign <> nAlign Then
    mInputTextAlign = nAlign
    If hEditBox Then
        Dim dwStyle As Long
        dwStyle = CLng(GetWindowLong(hEditBox, GWL_STYLE))
        Select Case mInputTextAlign
            Case TDIBTA_Left
                dwStyle = dwStyle And Not ES_CENTER
                dwStyle = dwStyle And Not ES_RIGHT
                dwStyle = dwStyle Or ES_LEFT
            Case TDIBTA_Right
                dwStyle = dwStyle And Not ES_CENTER
                dwStyle = dwStyle And Not ES_LEFT
                dwStyle = dwStyle Or ES_RIGHT
            Case TDIBTA_Center
                dwStyle = dwStyle And Not ES_LEFT
                dwStyle = dwStyle And Not ES_RIGHT
                dwStyle = dwStyle Or ES_CENTER
        End Select
        SetWindowLong hEditBox, GWL_STYLE, dwStyle
    End If
End If
End Property

Public Property Get ExpandedInfo() As String: ExpandedInfo = m_sExpandedInfo: End Property

Public Property Let InputCueBanner(sText As String)
m_sInput = sText
If hEditBox Then
    SendMessageW hEditBox, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sInput)
End If
End Property
Public Property Get InputCueBanner() As String: InputCueBanner = m_sInput: End Property
Public Property Let ComboCueBanner(sText As String)
m_sComboCue = sText
If hEditCombo Then
    SendMessageW hEditCombo, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sComboCue)
End If
End Property
Public Property Get ComboCueBanner() As String: ComboCueBanner = m_sComboCue: End Property

Public Property Let InputIsPassword(bVal As Boolean)
If hEditBox Then
    If bEditPassword <> bVal Then
        Dim dwStyle As Long
        dwStyle = CLng(GetWindowLong(hEditBox, GWL_STYLE))
        If bVal Then
            dwStyle = dwStyle Or ES_PASSWORD
        Else
            dwStyle = dwStyle And Not ES_PASSWORD
        End If
        Call SetWindowLong(hEditBox, GWL_STYLE, dwStyle)
        
    End If
End If
bEditPassword = bVal
End Property
Public Property Get InputIsPassword() As Boolean: InputIsPassword = bEditPassword: End Property

Public Property Let InputAlign(nAlign As TDInputBoxAlign): nInputAlign = nAlign: End Property
Public Property Get InputAlign() As TDInputBoxAlign: InputAlign = nInputAlign: End Property

Public Property Let InputAlignInFooter(nAlign As TDControlAlign): nIBFooterAlign = nAlign: End Property
Public Property Get InputAlignInFooter() As TDControlAlign: InputAlignInFooter = nIBFooterAlign: End Property

Public Property Let InputOffsetX(lValue As Long): IBOffX = lValue: End Property
Public Property Get InputOffsetX() As Long: InputOffsetX = IBOffX: End Property

Public Property Let InputWidth(CX As Long): cxInputBox = CX: End Property
Public Property Get InputWidth() As Long: InputWidth = cxInputBox: End Property

Public Property Let ComboAlign(nAlign As TDInputBoxAlign): nComboAlign = nAlign: End Property
Public Property Get ComboAlign() As TDInputBoxAlign: ComboAlign = nComboAlign: End Property

Public Property Let ComboAlignInFooter(nAlign As TDControlAlign): nCBFooterAlign = nAlign: End Property
Public Property Get ComboAlignInFooter() As TDControlAlign: ComboAlignInFooter = nCBFooterAlign: End Property

Public Property Let ComboOffsetX(lValue As Long): CBOffX = lValue: End Property
Public Property Get ComboOffsetX() As Long: ComboOffsetX = CBOffX: End Property

Public Property Let ComboWidth(CX As Long): cxCombo = CX: End Property
Public Property Get ComboWidth() As Long: ComboWidth = cxCombo: End Property

Public Property Let ComboDropWidth(CX As Long)
cxCBDrop = CX
If hCombo Then
    Call SendMessageW(hCombo, CB_SETDROPPEDWIDTH, CX, ByVal 0&)
End If
End Property
Public Property Get ComboDropWidth() As Long: ComboDropWidth = cxCBDrop: End Property

Public Property Let ComboHeight(CY As Long): cyCombo = CY: End Property
Public Property Get ComboHeight() As Long: ComboHeight = cyCombo: End Property

Public Property Let SliderAlign(nAlign As TDInputBoxAlign): nSliderAlign = nAlign: End Property
Public Property Get SliderAlign() As TDInputBoxAlign: SliderAlign = nSliderAlign: End Property

Public Property Let SliderAlignInFooter(nAlign As TDControlAlign): nSLFooterAlign = nAlign: End Property
Public Property Get SliderAlignInFooter() As TDControlAlign: SliderAlignInFooter = nSLFooterAlign: End Property

Public Property Let SliderOffsetX(lValue As Long): SLOffX = lValue: End Property
Public Property Get SliderOffsetX() As Long: SliderOffsetX = SLOffX: End Property

Public Property Let SliderWidth(CX As Long): cxSlider = CX: End Property
Public Property Get SliderWidth() As Long: SliderWidth = cxSlider: End Property

Public Property Let DateTimeAlign(nAlign As TDInputBoxAlign): nDateTimeAlign = nAlign: End Property
Public Property Get DateTimeAlign() As TDInputBoxAlign: ComboAlign = nDateTimeAlign: End Property

Public Property Let DateTimeAlignInContent(nAlign As TDControlAlign): nDTContentAlign = nAlign: End Property
Public Property Get DateTimeAlignInContent() As TDControlAlign: DateTimeAlignInContent = nDTContentAlign: End Property

Public Property Let DateTimeAlignInFooter(nAlign As TDControlAlign): nDTFooterAlign = nAlign: End Property
Public Property Get DateTimeAlignInFooter() As TDControlAlign: DateTimeAlignInFooter = nDTFooterAlign: End Property

Public Property Let DateTimeOffsetX(lValue As Long): DTOffX = lValue: End Property
Public Property Get DateTimeOffsetX() As Long: DateTimeOffsetX = DTOffX: End Property

Public Property Let DateTimeAlignInButtons(nAlign As TDControlAlign): nDTButtonAlign = nAlign: End Property
Public Property Get DateTimeAlignInButtons() As TDControlAlign: DateTimeAlignInFooter = nDTButtonAlign: End Property

Public Property Let DateTimeType(nType As DateTimeType): nDTType = nType: End Property
Public Property Get DateTimeType() As DateTimeType: DateTimeType = nDTType: End Property

Public Property Let DateTimeValue(dt As Date)
    With tDTCur
        .wDay = VBA.DateTime.Day(dt)
        .wMonth = VBA.DateTime.Month(dt)
        .wYear = VBA.DateTime.Year(dt)
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
    End With
    With tDTCur2
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
    End With
    If hDateTime2 Then
        Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur2)
    End If
    If hDateTime Then
        Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
    End If
End Property
Public Property Get DateTimeValue() As Date
Call SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
Call SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
If hDateTime2 Then
    DateTimeValue = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur2.wHour, tDTCur2.wMinute, tDTCur2.wSecond)
Else
    DateTimeValue = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur.wHour, tDTCur.wMinute, tDTCur.wSecond)
End If
End Property

Public Property Let DateTimeChecked(newValue As Long)
If hDateTime Then
    Call SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
End If
If hDateTime2 Then
    Call SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
End If
If nDTType = dttDateTimeWithCheck Then
    Select Case newValue
        Case 4
            bDTCheck = True
            bDTCheck2 = True
        Case 3
            bDTCheck = True
            bDTCheck2 = False
        Case 2
            bDTCheck = False
            bDTCheck2 = True
        Case 0
            bDTCheck = False
            bDTCheck2 = False
    End Select
    If hDateTime Then
        If bDTCheck Then
            Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
        Else
            Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
        End If
        If bDTCheck2 Then
            Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur2)
        Else
            Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur2)
        End If
    End If

Else
    If newValue = 1 Then
        bDTCheck = True
    ElseIf newValue = 0 Then
        bDTCheck = False
    End If
    
    If nDTType = dttDateTimeWithCheckTimeOnly Then
        If bDTCheck Then
            Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
        Else
            Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
        End If
    Else 'only checkbox is the first control
        If bDTCheck Then
            Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
        Else
            Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
        End If
    End If
End If

End Property
Public Property Get DateTimeChecked() As Long
Dim rs As Long, rs2 As Long
rs = CLng(SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur))
If hDateTime2 Then
    rs2 = CLng(SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2))
End If
If rs = GDT_VALID Then
    bDTCheck = True
Else
    bDTCheck = False
End If
If rs2 = GDT_VALID Then
    bDTCheck2 = True
Else
    bDTCheck2 = False
End If
If (nDTType = dttDateTimeWithCheck) Then 'don't return these values if there's only 1 check
    If bDTCheck And bDTCheck2 Then
        DateTimeChecked = 4
    ElseIf (bDTCheck = True) And (bDTCheck2 = False) Then
        DateTimeChecked = 3
    ElseIf (bDTCheck = False) And (bDTCheck2 = True) Then
        DateTimeChecked = 2
    ElseIf (bDTCheck = False) And (bDTCheck2 = False) Then
        DateTimeChecked = 0
    End If
Else
    If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
        If bDTCheck = True Then
            DateTimeChecked = 1
        Else
            DateTimeChecked = 0
        End If
    Else 'date time with a single check, on the time
        If bDTCheck2 = True Then
            DateTimeChecked = 1
        Else
            DateTimeChecked = 0
        End If
    End If
End If
End Property

Public Sub DateTimeSetInitial(dt As Date)
bDTSetInit = True
With tDTInitial
        .wDay = VBA.DateTime.Day(dt)
        .wMonth = VBA.DateTime.Month(dt)
        .wYear = VBA.DateTime.Year(dt)
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
End With
End Sub
Public Sub zzDateTimeGetInitial(bs As Boolean, tdt() As Integer)
bs = bDTSetInit
ReDim tdt(5)
With tDTInitial
    tdt(0) = .wDay
    tdt(1) = .wMonth
    tdt(2) = .wYear
    tdt(3) = .wHour
    tdt(4) = .wMinute
    tdt(5) = .wSecond
End With
End Sub

Public Sub DateTimeSetRange(bEnableMin As Boolean, bEnableMax As Boolean, dtMin As Date, dtMax As Date)
bDTSetMin = bEnableMin
bDTSetMax = bEnableMax
With tDTRange(0)
        .wDay = VBA.DateTime.Day(dtMin)
        .wMonth = VBA.DateTime.Month(dtMin)
        .wYear = VBA.DateTime.Year(dtMin)
        .wHour = VBA.DateTime.Hour(dtMin)
        .wMinute = VBA.DateTime.Minute(dtMin)
        .wSecond = VBA.DateTime.Second(dtMin)
End With
With tDTRange(1)
        .wDay = VBA.DateTime.Day(dtMax)
        .wMonth = VBA.DateTime.Month(dtMax)
        .wYear = VBA.DateTime.Year(dtMax)
        .wHour = VBA.DateTime.Hour(dtMax)
        .wMinute = VBA.DateTime.Minute(dtMax)
        .wSecond = VBA.DateTime.Second(dtMax)
End With
SetDTRange
End Sub

Public Sub zzDateTimeGetRange(bn As Boolean, bm As Boolean, tdt() As Integer, tdt2() As Integer)
bn = bDTSetMin
bm = bDTSetMax
ReDim tdt(5)
ReDim tdt2(5)
With tDTRange(0)
    tdt(0) = .wDay
    tdt(1) = .wMonth
    tdt(2) = .wYear
    tdt(3) = .wHour
    tdt(4) = .wMinute
    tdt(5) = .wSecond
End With
With tDTRange(1)
    tdt2(0) = .wDay
    tdt2(1) = .wMonth
    tdt2(2) = .wYear
    tdt2(3) = .wHour
    tdt2(4) = .wMinute
    tdt2(5) = .wSecond
End With
End Sub
Private Sub SetDTRange()
Dim dwValid As Long
If bDTSetMin Then dwValid = GDTR_MIN
If bDTSetMax Then dwValid = dwValid Or GDTR_MAX
If hDateTime Then
    If (nDTType = dttTime) Or (nDTType = dttTimeWithCheck) Then
        tDTRange(1).wDay = tDTRange(0).wDay
        tDTRange(1).wMonth = tDTRange(0).wMonth
        tDTRange(1).wYear = tDTRange(0).wYear
    End If
    Call SendMessageW(hDateTime, DTM_SETRANGE, dwValid, ByVal VarPtr(tDTRange(0)))
End If
If hDateTime2 Then
    tDTRange(1).wDay = tDTRange(0).wDay
    tDTRange(1).wMonth = tDTRange(0).wMonth
    tDTRange(1).wYear = tDTRange(0).wYear
    'fixes a bug.. if the date isn't set like that, applying it to a time control will make the max ignored
     Call SendMessageW(hDateTime2, DTM_SETRANGE, dwValid, ByVal VarPtr(tDTRange(0)))
End If
End Sub

Public Property Get SliderTickStyle() As SldTickStyle: SliderTickStyle = nSliderTickStyle: End Property
Public Property Let SliderTickStyle(nStyle As SldTickStyle): nSliderTickStyle = nStyle: End Property

Public Property Get SliderValue() As Long
If hSlider Then
    lSliderValue = CLng(SendMessageW(hSlider, TBM_GETPOS, 0, ByVal 0&))
End If
SliderValue = lSliderValue
End Property
Public Property Let SliderValue(nValue As Long)
lSliderValue = nValue
If hSlider Then
    Call SendMessageW(hSlider, TBM_SETPOS, 1&, ByVal lSliderValue)
End If
    
End Property
Public Sub SliderSetRange(lMin As Long, lMax As Long, Optional lTickFreq As Long)
lSliderMin = lMin
lSliderMax = lMax
lSliderNTicks = lTickFreq
If hSlider Then
    SendMessageW hSlider, TBM_SETRANGEMIN, 0, ByVal lMin
    SendMessageW hSlider, TBM_SETRANGEMAX, 1, ByVal lMax
    If lTickFreq Then
        SendMessageW hSlider, TBM_SETTICFREQ, lSliderNTicks, ByVal 0&
    End If
End If
End Sub
Public Sub zzSliderGetRange(n1 As Long, n2 As Long, n3 As Long)
n1 = lSliderMin: n2 = lSliderMax: n3 = lSliderNTicks
End Sub

Public Sub SliderSetChangeValues(lSmallChange As Long, lLargeChange As Long)
lSliderChangeS = lSmallChange
lSliderChangeL = lLargeChange
If hSlider Then
    SendMessageW hSlider, TBM_SETLINESIZE, 0, ByVal lSliderChangeS
    SendMessageW hSlider, TBM_SETPAGESIZE, 0, ByVal lSliderChangeL
End If
End Sub
Public Sub zzSliderGetChangeValues(n1 As Long, n2 As Long)
n1 = lSliderChangeS: n2 = lSliderChangeL
End Sub

Public Property Let ComboIndex(nIndex As Long)
nComboIdx = nIndex
If hCombo Then
    Call SendMessageW(hCombo, CB_SETCURSEL, nIndex, ByVal 0&)
End If
End Property
Public Property Get ComboIndex() As Long
nComboIdx = -1
If hCombo Then
    nComboIdx = CLng(SendMessageW(hCombo, CB_GETCURSEL, 0&, ByVal 0&))
End If
ComboIndex = nComboIdx
End Property
Public Property Get ComboText() As String
If hCombo Then
    ComboText = GetComboTextW(hEditCombo)
End If
End Property
Public Property Let ComboText(sText As String)
If hCombo Then
    SendMessageW hEditCombo, WM_SETTEXT, 0&, ByVal StrPtr(sText)
    sComboText = sText
End If
End Property
Public Property Let ComboImageList(himl As LongPtr): himlCombo = himl: End Property
Public Property Get ComboImageList() As LongPtr: ComboImageList = himlCombo: End Property

Public Property Let ComboStyle(nStyle As ComboType): nCbType = nStyle: End Property
Public Property Get ComboStyle() As ComboType: ComboStyle = nCbType: End Property

Public Property Let ExpandedInfo(sText As String)
m_sExpandedInfo = sText
uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_EXPANDED_INFORMATION, ByVal StrPtr(m_sExpandedInfo))
End If
End Property

Public Property Get ExpandedControlText() As String: ExpandedControlText = m_sExpControl: End Property
Public Property Let ExpandedControlText(sText As String)
m_sExpControl = sText
uTDC.pszExpandedControlText = StrPtr(m_sExpControl)
End Property

Public Property Get CollapsedControlText() As String: CollapsedControlText = m_sColControl: End Property
Public Property Let CollapsedControlText(sText As String)
m_sColControl = sText
uTDC.pszCollapsedControlText = StrPtr(m_sColControl)
End Property

Public Property Get IconFooter() As LongPtr: IconFooter = m_IconFooter: End Property
Public Property Let IconFooter(lIcon As LongPtr)
m_IconFooter = lIcon

If m_TDhWnd <> 0& Then
    If (uTDC.dwFlags And TDF_USE_HICON_FOOTER) Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_FOOTER, ByVal m_IconFooter)
    Else
        Dim dwIcon As Long
        dwIcon = tdMakeIntResource(CLng(m_IconFooter))
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_FOOTER, ByVal dwIcon)
    End If
    'Sometimes the icon wasn't redrawing, repainting the window fixed it
    Dim hFrame As LongPtr
    hFrame = FindWindowEx(m_TDhWnd, 0&, "DirectUIHWND", vbNullString)
    RepaintWindow hFrame
End If
End Property

Public Property Get Footer() As String: Footer = m_sFooter: End Property
Public Property Let Footer(sText As String)
m_sFooter = sText
uTDC.pszFooter = StrPtr(m_sFooter)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
End If
End Property

Public Property Get Width() As Long: Width = uTDC.CXWidth: End Property
Public Property Let Width(Value As Long): uTDC.CXWidth = Value: End Property

Public Property Get ResultMain() As TDRESULT: ResultMain = m_ResMain: End Property
Public Property Get ResultRad() As Long: ResultRad = m_ResRad: End Property
Public Property Get ResultVerify() As Long: ResultVerify = m_ResVerify: End Property
Public Property Get ResultInput() As String: ResultInput = sEditText: End Property
Public Property Get ResultComboText() As String: ResultComboText = sComboText: End Property
Public Property Get ResultComboIndex() As Long: ResultComboIndex = nComboIdx: End Property
Public Property Get ResultSlider() As Long: ResultSlider = lSliderValue: End Property
Public Property Get ResultDateTime() As Date
    If (nDTType = dttDate) Or (nDTType = dttDateWithCheck) Then
        ResultDateTime = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay)
    End If
    If (nDTType = dttTime) Or (nDTType = dttTimeWithCheck) Then
        ResultDateTime = TimeSerial(tDTCur.wHour, tDTCur.wMinute, tDTCur.wSecond)
    End If
    If (nDTType = dttDateTime) Or (nDTType = dttDateTimeWithCheck) Then
        ResultDateTime = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur2.wHour, tDTCur2.wMinute, tDTCur2.wSecond)
    End If
End Property
Public Property Get ResultDateTimeChecked() As Long
If (nDTType = dttDateTimeWithCheck) Then 'don't return these values if there's only 1 check
    If bDTCheck And bDTCheck2 Then
        ResultDateTimeChecked = 4
    ElseIf (bDTCheck = True) And (bDTCheck2 = False) Then
        ResultDateTimeChecked = 3
    ElseIf (bDTCheck = False) And (bDTCheck2 = True) Then
        ResultDateTimeChecked = 2
    ElseIf (bDTCheck = False) And (bDTCheck2 = False) Then
        ResultDateTimeChecked = 0
    End If
Else
    If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
        If bDTCheck = True Then
            ResultDateTimeChecked = 1
        Else
            ResultDateTimeChecked = 0
        End If
    Else 'date time with a single check, on the time
        If bDTCheck2 = True Then
            ResultDateTimeChecked = 1
        Else
            ResultDateTimeChecked = 0
        End If
    End If
End If

End Property

Public Property Get hWndDlg() As LongPtr: hWndDlg = m_TDhWnd: End Property
Public Property Get hWndDUI() As LongPtr: hWndDUI = hUI: End Property 'the DirectUIHWND class that the controls are on; and what is used as a parent for added controls
Public Property Get hWndInput() As LongPtr: hWndInput = hEditBox: End Property
Public Property Get hWndCombo() As LongPtr: hWndCombo = hCombo: End Property
Public Property Get hWndComboEdit() As LongPtr: hWndComboEdit = hEditCombo: End Property
Public Property Get hWndDateTime() As LongPtr: hWndDateTime = hDateTime: End Property
Public Property Get hWndSlider() As LongPtr: hWndSlider = hSlider: End Property

Public Sub ComboAddItem(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
If aComboItems(0).sText = "" Then
    aComboItems(0).sText = sText
    aComboItems(0).iImage = iImage
    aComboItems(0).iOverlay = iOverlay
    Exit Sub
End If
ReDim Preserve aComboItems(UBound(aComboItems) + 1)
aComboItems(UBound(aComboItems)).sText = sText
aComboItems(UBound(aComboItems)).iImage = iImage
aComboItems(UBound(aComboItems)).iOverlay = iOverlay
If hCombo Then
    CBX_InsertItem hCombo, sText, iImage, iOverlay
End If
End Sub
Public Sub zzComboGetItemData(lMax As Long, sText() As String, iImage() As Long, iOverlay() As Long)
Attribute zzComboGetItemData.VB_MemberFlags = "40"
Dim nct As Long
nct = UBound(aComboItems)
lMax = nct
ReDim sText(nct)
ReDim iImage(nct)
ReDim iOverlay(nct)
Dim i As Long
For i = 0 To UBound(aComboItems)
    sText(i) = aComboItems(i).sText
    iImage(i) = aComboItems(i).iImage
    iOverlay(i) = aComboItems(i).iOverlay
Next i
End Sub


Public Sub ComboSetInitialState(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
bCBInitValid = True
tComboInit.sText = sText
tComboInit.iImage = iImage
tComboInit.iOverlay = iOverlay
End Sub
Public Sub zzComboGetInitialState(sText As String, iImage As Long, iOverlay As Long, bValid As Boolean)
Attribute zzComboGetInitialState.VB_MemberFlags = "40"
bValid = bCBInitValid
sText = tComboInit.sText
iImage = tComboInit.iImage
iOverlay = tComboInit.iOverlay
End Sub
Public Sub ComboSetInitialItem(nIndex As Long)
nComboInit = nIndex
End Sub
Public Function zzComboGetInitialItem() As Long
Attribute zzComboGetInitialItem.VB_MemberFlags = "40"
zzComboGetInitialItem = nComboInit
End Function
Public Sub ComboSetCurrentState(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
Dim cbxi As COMBOBOXEXITEMW
cbxi.Mask = CBEIF_IMAGE Or CBEIF_OVERLAY
cbxi.iItem = -1
cbxi.iImage = iImage
cbxi.iOverlay = iOverlay
If sText <> "" Then
    cbxi.Mask = cbxi.Mask Or CBEIF_TEXT
    cbxi.pszText = StrPtr(sText)
End If
Call SendMessageW(hCombo, CBEM_SETITEMW, 0, cbxi)

End Sub
Public Function SimpleDialog(sMessage As String, Optional dwBtn As TDBUTTONS = TDCBF_OK_BUTTON, Optional sTitle As String, Optional sMainText As String, Optional dwIco As TDICONS, Optional hWndOwner As LongPtr, Optional hinst As LongPtr) As TDBUTTONS
   Dim dwIcon As Long
   Dim pnButton As Long
   Dim Success As Long
    
Dim pszTitle As LongPtr
Dim pszMain As LongPtr
Dim pszContent As LongPtr

If sTitle = "" Then
    sTitle = App.Title
End If
pszTitle = StrPtr(sTitle)
If sMainText <> "" Then pszMain = StrPtr(sMainText)
If sMessage <> "" Then pszContent = StrPtr(sMessage)

If dwIco Then
   dwIcon = tdMakeIntResource(dwIco)
End If
  Call TaskDialog(hWndOwner, _
                        hinst, _
                        pszTitle, _
                        pszMain, _
                        pszContent, _
                        dwBtn, _
                        dwIcon, _
                        pnButton)

    SimpleDialog = pnButton
End Function

Public Function ShowDialog() As Boolean
'main function; returns True if an option was selected, False if cancelled
'then use .

Dim pnButton As Long, pnRadButton As Long, pfVerify As Long
Dim hr As Long
lAutoclose = lAutocloseTrack
'sEditText = ""
If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Footer)) Or _
   ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Footer)) Or _
   ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Footer)) Or _
   ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Footer)) Then

    'ensure footer text isn't blank, or it won't appear
    If m_sFooter = "" Then
        m_sFooter = "$input"
    End If
End If
bEditFooter = False
If m_sFooter = "$input" Then
    bEditFooter = True
    m_sFooter = " "
    uTDC.pszFooter = StrPtr(m_sFooter)
End If
If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Content)) Or _
   ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Content)) Or _
   ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Content)) Or _
   ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Content)) Then
        If m_sContent <> "" Then
            'If Right$(m_sContent, 3) <> (vbCrLf & " ") Then
'                m_sContent = m_sContent & vbCrLf & " "
'                nBreaksAdded = 1
                'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                    If Right$(m_sContent, 5) <> (vbCrLf & vbCrLf & " ") Then
                        m_sContent = m_sContent & vbCrLf & vbCrLf
                        If (uTDC.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                        Dim lnkPos As Long
                        lnkPos = (InStr(LCase$(m_sContent), "<a href"))
'                        DebugAppend "lnkPos=" & lnkPos
                         If lnkPos Then
                            bLinkAdjust = True
                            m_sContent = m_sContent & " " ' (link control requires a space or it ignores the double-break)
                                                           'Meanwhile, everything else gets thrown out of alignment, so ONLY add to link
                         End If
                        End If
                        nBreaksAdded = 2
                        uTDC.pszContent = StrPtr(m_sContent)
                    End If
                'End If
            'End If
        Else 'No content, main instruction only: add one break
            If Right$(m_sMainInstr, 2) <> vbCrLf Then
                m_sMainInstr = m_sMainInstr & vbCrLf
                uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
            End If

        End If
        If m_sExpandedInfo <> "" Then
            If Right$(m_sExpandedInfo, 2) <> vbCrLf Then
                m_sExpandedInfo = m_sExpandedInfo & vbCrLf
                'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                    If Right$(m_sExpandedInfo, 4) <> (vbCrLf & vbCrLf) Then
                        m_sExpandedInfo = m_sExpandedInfo & vbCrLf  '(it's already sure to have 1)
                    End If
                'End If
                uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
            End If

        End If
End If
If (uTDC.dwFlags And TDF_CALLBACK_TIMER) Then
    bPublicTimer = True
Else
    bPublicTimer = False
    If (uTDC.dwFlags And TDF_INPUT_BOX) Or (uTDC.dwFlags And TDF_COMBO_BOX) Or (uTDC.dwFlags And TDF_DATETIME) Or (uTDC.dwFlags And TDF_SLIDER) Then
        If m_sExpandedInfo <> "" Then
            If (uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                uTDC.dwFlags = uTDC.dwFlags Or TDF_CALLBACK_TIMER
            End If
        End If
    End If
    If (lAutoclose > 0) Then
        uTDC.dwFlags = uTDC.dwFlags Or TDF_CALLBACK_TIMER
    End If
End If
If (uTDC.dwFlags And TDF_USE_HICON_MAIN) Then
    uTDC.pszMainIcon = m_Icon
Else
    uTDC.pszMainIcon = tdMakeIntResource(CLng(m_Icon))
End If
If (uTDC.dwFlags And TDF_USE_HICON_FOOTER) Then
    uTDC.pszFooterIcon = m_IconFooter
Else
    uTDC.pszFooterIcon = tdMakeIntResource(CLng(m_IconFooter))
End If

If m_HasCustomButtons Then
    uTDC.cButtons = UBound(m_uButtons) + 1
Else
    uTDC.cButtons = 0
End If
uTDC.pButtons = VarPtr(m_uButtons(0).nButtonID)

If m_HasRadioButtons Then
    uTDC.cRadioButtons = UBound(m_uRadioButtons) + 1
Else
    uTDC.cRadioButtons = 0
End If
uTDC.pRadioButtons = VarPtr(m_uRadioButtons(0).nButtonID)

If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID) Then
    uTDC.hInstance = hShell32
    uTDC.pszMainIcon = tdMakeIntResource(CLng(m_Icon))
    uTDC.pszFooterIcon = tdMakeIntResource(CLng(m_IconFooter))
ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID) Then
    uTDC.hInstance = hImageRes
    uTDC.pszMainIcon = tdMakeIntResource(CLng(m_Icon))
    uTDC.pszFooterIcon = tdMakeIntResource(CLng(m_IconFooter))
End If

'If m_cbMain = 0& Then m_cbMain = AddressOf ProcessCallback ' scb_SetCallbackAddr(5, 2)
'uTDC.pfCallback = m_cbMain
uTDC.pfCallback = tdFARPROC(AddressOf TaskDialogCallbackProc)
uTDC.lpCallbackData = ObjPtr(Me)
sComboText = ""
nComboIdx = -1
'debugappend "ScaleX=" & m_ScaleX & ",Y=" & m_ScaleY
hr = TaskDialogIndirect(uTDC, pnButton, pnRadButton, pfVerify)
Dim sErr As String
'Call GetLastError_Msg(hr, , sErr)
DebugAppend "TaskDialogIndirect ret=0x" & Hex$(hr) & ": " & sErr
m_ResMain = pnButton
m_ResRad = pnRadButton
m_ResVerify = pfVerify
If pnButton <> 0 Then ShowDialog = True

End Function


Public Sub AddCustomButton(nID As Long, sText As String, Optional hIcon As LongPtr = 0&)
If m_HasCustomButtons Then
    ReDim Preserve m_uButtons(UBound(m_uButtons) + 1)
    ReDim Preserve btnCaps(UBound(btnCaps) + 1)
    btnCaps(UBound(btnCaps)) = sText
    ReDim Preserve icoDataC(UBound(m_uButtons))
    If hIcon Then icoDataC(UBound(icoDataC)) = hIcon
    m_uButtons(UBound(m_uButtons)).nButtonID = nID
    m_uButtons(UBound(m_uButtons)).pszButtonText = StrPtr(btnCaps(UBound(btnCaps)))
Else
    btnCaps(0) = sText
    m_uButtons(0).nButtonID = nID
    m_uButtons(0).pszButtonText = StrPtr(btnCaps(0))
    If hIcon Then icoDataC(0) = hIcon
    m_HasCustomButtons = True
End If
End Sub
Public Sub SetSplitButton(idCustomButton As Long)
nSplitButtonID = idCustomButton
End Sub
Public Function zzGetSplitButton() As Long
zzGetSplitButton = nSplitButtonID
End Function
Public Sub GetCustomButtons(nID() As Long, sText() As String)
Dim i As Long
Dim lID() As Long
Dim c As Long
c = UBound(m_uButtons)
ReDim lID(c)
For i = 0 To c
    lID(i) = m_uButtons(i).nButtonID
Next i
nID = lID
sText = btnCaps
    
End Sub
Public Sub AddRadioButton(nID As Long, sText As String)
If m_HasRadioButtons Then
    ReDim Preserve m_uRadioButtons(UBound(m_uRadioButtons) + 1)
    ReDim Preserve rbtnCaps(UBound(rbtnCaps) + 1)
    rbtnCaps(UBound(rbtnCaps)) = sText
    m_uRadioButtons(UBound(m_uRadioButtons)).nButtonID = nID
    m_uRadioButtons(UBound(m_uRadioButtons)).pszButtonText = StrPtr(rbtnCaps(UBound(rbtnCaps)))
Else
    rbtnCaps(0) = sText
    m_uRadioButtons(0).nButtonID = nID
    m_uRadioButtons(0).pszButtonText = StrPtr(rbtnCaps(0))
    m_HasRadioButtons = True
End If
End Sub

Public Sub GetRadioButtons(nID() As Long, sText() As String)
Dim i As Long
Dim lID() As Long
Dim c As Long
c = UBound(m_uRadioButtons)
ReDim lID(c)
For i = 0 To c
    lID(i) = m_uRadioButtons(i).nButtonID
Next i
nID = lID
sText = rbtnCaps
    
End Sub

Public Sub ClearCustomButtons()
ReDim m_uButtons(0)
ReDim btnCaps(0)
ReDim icoDataC(0)
nSplitButtonID = 0
uTDC.cButtons = 0
uTDC.pButtons = VarPtr(m_uButtons(0).nButtonID)
m_HasCustomButtons = False
End Sub
Public Sub ClearRadioButtons()
ReDim m_uRadioButtons(0)
ReDim rbtnCaps(0)
uTDC.cRadioButtons = 0
uTDC.pRadioButtons = VarPtr(m_uRadioButtons(0).nButtonID)
m_HasRadioButtons = False
End Sub

Public Sub Init()
'resets the entire module
mPageIndex = 0
hLogoImage = 0
nLogoPos = 0
nLogoType = 0
lLogoX = 0
lLogoY = 0
lAutoclose = 0
bLinkAdjust = False
ClearCustomButtons
ClearRadioButtons
m_HasCustomControl = False
ReDim m_uButtons(0)
ReDim m_uRadioButtons(0)
ReDim btnCaps(0)
ReDim rbtnCaps(0)
ReDim lBtnElv(0)
ReDim lBtnHold(0)
ReDim lBtnDis(0)
ReDim lRadDis(0)
ReDim icoDataC(0)
ReDim aComboItems(0)
tComboInit.sText = ""
tComboInit.iImage = 0
tComboInit.iOverlay = 0
nCbType = 0
bCBInitValid = False
nDTType = 0
bDTSetInit = False
bDTCheck = False
bDTSetMin = False
bDTSetMax = False
nComboInit = -1
himlCombo = 0
hDateTime = 0
hDateTime2 = 0
hCombo = 0
hSlider = 0
hEditBox = 0
icoDataN(5) = 0
icoDataN(4) = 0
icoDataN(3) = 0
icoDataN(2) = 0
icoDataN(1) = 0
icoDataN(0) = 0
m_lNewIcon = 0
m_BtnLRSize = 16
bDis = False
bRadDis = False
bElv = False
nInputAlign = TDIBA_Content
nComboAlign = TDIBA_Content
mInputTextAlign = TDIBTA_Left
nDateTimeAlign = TDIBA_Content
nDTContentAlign = 0
nDTFooterAlign = 0
nDTButtonAlign = 0
DTOffX = 0
CBOffX = 0
IBOffX = 0
SLOffX = 0
nCBFooterAlign = 0
nIBFooterAlign = 0
nSLFooterAlign = 0
nSliderAlign = TDIBA_Content
nDefaultCustCtrl = 0
bPosUsed(0) = False
bPosUsed(1) = False
bPosUsed(2) = False
cxInputBox = 0
cxCombo = 0
cyCombo = 0
cxCBDrop = 0
cxSlider = 0
lSliderMin = 0
lSliderMax = 10
lSliderNTicks = 1
lSliderValue = 0
lSliderChangeS = 1
lSliderChangeL = 2
nSliderTickStyle = 0
nBreaksAdded = 0
With uTDC
    .CXWidth = 0
    .dwCommonButtons = 0
    .dwFlags = 0
    .hInstance = 0 ' App.hInstance
    .hWndParent = 0
    .nDefaultButton = 0
    .nDefaultRadioButton = 0
    m_sExpandedInfo = ""
    .pszExpandedInformation = 0 'StrPtr(m_sExpandedInfo)
    
    m_sColControl = ""
    uTDC.pszCollapsedControlText = StrPtr(m_sColControl)
'
    m_sContent = ""
    .pszContent = StrPtr(m_sContent)
'
    m_sExpControl = ""
    .pszExpandedControlText = StrPtr(m_sExpControl)

'
    m_sFooter = ""
    .pszFooter = StrPtr(m_sFooter)
'
    m_IconFooter = 0
    .pszFooterIcon = 0
'
    m_Icon = 0
    .pszMainIcon = 0
'
    m_sMainInstr = ""
    .pszMainInstruction = StrPtr(m_sMainInstr)
'
    m_sVerify = ""
    .pszVerificationText = StrPtr(m_sVerify)
'
    m_sTitle = App.Title
    .pszWindowTitle = StrPtr(m_sTitle)
'
    m_sInput = ""
    sEditText = ""
    
End With
bEditPassword = False
bEditFooter = False
End Sub

Public Sub CloseDialog()
'closes the dialog if open
If m_TDhWnd <> 0& Then
    Call ClickButton(TDCBF_CANCEL_BUTTON)
End If

End Sub
Public Sub ResetTimer()
bResetTimer = True
End Sub

Public Sub ProgressStartMarquee(Optional dwSpeed As Long = 0)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_MARQUEE, 1, ByVal dwSpeed)
End If
End Sub
Public Sub ProgressStopMarquee()
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_MARQUEE, 0, ByVal 0&)
End If
End Sub

Public Sub ProgressSetType(lType As Long)
'0=regular
'1=marquee
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_MARQUEE_PROGRESS_BAR, lType, ByVal 0&)
End If
End Sub

Public Sub ProgressSetState(lState As eProgressBarStates)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_STATE, lState, ByVal 0&)
End If
End Sub

Public Sub ProgressSetRange(lMin As Long, lMax As Long)
Dim dwRange As Long
dwRange = tdMAKELPARAM(lMin, lMax)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_RANGE, 0, ByVal dwRange)
End If
End Sub

Public Sub ProgressSetValue(lValue As Long)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_POS, lValue, ByVal 0&)
End If
End Sub

Public Sub EnableButton(ButtonID As Long, lEnable As Long)
'lEnable=0 disable; <>0 enable
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_ENABLE_BUTTON, ButtonID, ByVal lEnable)
Else
    If lEnable = 0 Then
        If (lBtnDis(0) = 0) And (UBound(lBtnDis) = 0) Then
            lBtnDis(0) = ButtonID
        Else
            ReDim Preserve lBtnDis(UBound(lBtnDis) + 1)
            lBtnDis(UBound(lBtnDis)) = ButtonID
        End If
    Else
        Dim lnew() As Long
        ReDim lnew(0)
        Dim i As Long, k As Long
        For i = 0 To UBound(lBtnDis)
            If lBtnDis(i) <> ButtonID Then
                ReDim Preserve lnew(k)
                lnew(k) = lBtnDis(i)
            End If
        Next i
        lBtnDis = lnew
    End If
    If (UBound(lBtnDis) = 0) And (lBtnDis(0) = 0) Then
        bDis = False
    Else
        bDis = True
    End If
End If
End Sub
Public Sub EnableRadioButton(ButtonID As Long, lEnable As Long)
'lEnable=0 disable; <>0 enable
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_ENABLE_RADIO_BUTTON, ButtonID, ByVal lEnable)
Else
    If lEnable = 0 Then
        If (lRadDis(0) = 0) And (UBound(lRadDis) = 0) Then
            lRadDis(0) = ButtonID
        Else
            ReDim Preserve lRadDis(UBound(lRadDis) + 1)
            lRadDis(UBound(lRadDis)) = ButtonID
        End If
    Else
        Dim lnew() As Long
        ReDim lnew(0)
        Dim i As Long, k As Long
        For i = 0 To UBound(lRadDis)
            If lRadDis(i) <> ButtonID Then
                ReDim Preserve lnew(k)
                lnew(k) = lRadDis(i)
            End If
        Next i
        lRadDis = lnew
    End If
    If (UBound(lRadDis) = 0) And (lRadDis(0) = 0) Then
        bRadDis = False
    Else
        bRadDis = True
    End If
End If
End Sub

Public Sub ClickButton(ButtonID As Long)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_CLICK_BUTTON, ButtonID, ByVal 0&)
End If
End Sub
Public Sub ClickRadioButton(ButtonID As Long)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_CLICK_RADIO_BUTTON, ButtonID, ByVal 0&)
End If
End Sub
Public Sub ClickVerification(lChecked As Long, Optional lSetFocus As Long = 1)
'0=unchecked,1=checked
'lSetFocus=1, sets keyboard focus on box. 0 to not set focus
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_CLICK_VERIFICATION, lChecked, ByVal lSetFocus)
End If
End Sub
Public Sub SetButtonHold(ButtonID As Long)
If (lBtnHold(0) = 0) And (UBound(lBtnHold) = 0) Then
    lBtnHold(0) = ButtonID
Else
    ReDim Preserve lBtnHold(UBound(lBtnHold) + 1)
    lBtnHold(UBound(lBtnHold)) = ButtonID
End If
End Sub
Public Function zzGetButtonHolds() As Long()
zzGetButtonHolds = lBtnHold
End Function

Public Sub ReleaseButtonHold(ButtonID As Long)
Dim lnew() As Long
Dim i As Long, J As Long
ReDim lnew(0)
For i = 0 To UBound(lBtnHold)
    If lBtnHold(i) <> ButtonID Then
        ReDim Preserve lnew(J)
        lnew(J) = lBtnHold(i)
        J = J + 1
    End If
Next i
lBtnHold = lnew
End Sub
Private Function ButtonIsHeld(id As Long) As Boolean
Dim i As Long
For i = 0 To UBound(lBtnHold)
    If lBtnHold(i) = id Then
        ButtonIsHeld = True
        Exit Function
    End If
Next i
End Function

Public Sub SetButtonElevated(ButtonID As Long, lState As Long)
If m_TDhWnd <> 0& Then
    Call SendMessageW(m_TDhWnd, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, ButtonID, ByVal lState)
Else
    If lState Then
        If (lBtnElv(0) = 0) And (UBound(lBtnElv) = 0) Then
            lBtnElv(0) = ButtonID
        Else
            ReDim Preserve lBtnElv(UBound(lBtnElv) + 1)
            lBtnElv(UBound(lBtnElv)) = ButtonID
        End If
    Else
        Dim lnew() As Long
        ReDim lnew(0)
        Dim i As Long, k As Long
        For i = 0 To UBound(lBtnElv)
            If lBtnElv(i) <> ButtonID Then
                ReDim Preserve lnew(k)
                lnew(k) = lBtnElv(i)
            End If
        Next i
        lBtnElv = lnew
    End If
    If (UBound(lBtnElv) = 0) And (lBtnElv(0) = 0) Then
        bElv = False
    Else
        bElv = True
    End If
End If
End Sub
Public Function zzGetButtonElevated(be As Boolean) As Long()
be = bElv
zzGetButtonElevated = lBtnElv
End Function
Public Sub SetCommonButtonIcon(id As TDBUTTONS, hIcon As LongPtr)
Select Case id
    Case TDCBF_OK_BUTTON: icoDataN(0) = hIcon
    Case TDCBF_YES_BUTTON: icoDataN(1) = hIcon
    Case TDCBF_NO_BUTTON: icoDataN(2) = hIcon
    Case TDCBF_CANCEL_BUTTON: icoDataN(3) = hIcon
    Case TDCBF_RETRY_BUTTON: icoDataN(4) = hIcon
    Case TDCBF_CLOSE_BUTTON: icoDataN(5) = hIcon
End Select
End Sub
Public Function zzGetCommonButtonIcon() As Long()
zzGetCommonButtonIcon = icoDataN
End Function
Private Function tdMakeIntResource(ByVal dwVal As Long) As Long
   tdMakeIntResource = &HFFFF& And dwVal
End Function
Private Function tdFARPROC(pfn As LongPtr) As LongPtr
  tdFARPROC = pfn
End Function
Private Function tdMAKELPARAM(wLow As Long, wHigh As Long) As Long
  tdMAKELPARAM = tdLoWord(wLow) Or (&H10000 * tdLoWord(wHigh))
End Function
Private Function tdLoWord(dwValue As Long) As Integer
  CopyMemory tdLoWord, dwValue, 2
End Function
Private Function tdHiWord(ByVal dwValue As Long) As Integer
  CopyMemory tdHiWord, ByVal VarPtr(dwValue) + 2, 2
End Function

Private Sub Class_Initialize()
Dim hDC As LongPtr
hDC = GetDC(0&)
m_ScaleX = GetDeviceCaps(hDC, LOGPIXELSX) / 96
m_ScaleY = GetDeviceCaps(hDC, LOGPIXELSY) / 96
ReleaseDC 0&, hDC
ReDim m_uButtons(0)
ReDim m_uRadioButtons(0)
ReDim btnCaps(0)
ReDim rbtnCaps(0)
ReDim lBtnHold(0)
ReDim lBtnElv(0)
ReDim lBtnDis(0)
ReDim lRadDis(0)
ReDim icoDataC(0)
ReDim aComboItems(0)
lSliderMin = 0
lSliderMax = 10
lSliderNTicks = 1
lSliderValue = 0
lSliderChangeS = 1
lSliderChangeL = 2
nSliderTickStyle = 0
m_BtnLRSize = 16
hImageRes = LoadLibrary("imageres.dll")
hShell32 = LoadLibrary("shell32.dll")

uTDC.cbSize = LenB(uTDC)
uTDC.hInstance = 0 'App.hInstance

m_sTitle = App.Title
uTDC.pszWindowTitle = StrPtr(m_sTitle)

End Sub

Private Sub Class_Terminate()
If hShell32 Then FreeLibrary hShell32
If hImageRes Then FreeLibrary hImageRes
End Sub

Public Function NavigatePage(cTD As cTaskDialog) As Long
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
If m_TDhWnd <> 0& Then
Dim i As Long
Dim uTDC2 As TASKDIALOGCONFIG
Dim lBtn() As Long, lRBN() As Long
Dim sBTN() As String, sRBN() As String
Dim uBTN() As TASKDIALOG_BUTTON, uRBN() As TASKDIALOG_BUTTON
Dim sz() As String
ReDim sz(7)
cTD.GetCustomButtons lBtn, sBTN
cTD.GetRadioButtons lRBN, sRBN
ReDim uBTN(0)
ReDim uRBN(0)
With uTDC2
    .cbSize = Len(uTDC2)
    If (UBound(lBtn) = 0) And (lBtn(0) = 0) Then
        .cButtons = 0
    Else
        .cButtons = UBound(lBtn) + 1
        ReDim uBTN(UBound(lBtn))
        For i = 0 To UBound(lBtn)
            uBTN(i).nButtonID = lBtn(i)
            uBTN(i).pszButtonText = StrPtr(sBTN(i))
        Next i
    End If
    .pButtons = VarPtr(uBTN(0).nButtonID)
    If (UBound(lRBN) = 0) And (lRBN(0) = 0) Then
        .cRadioButtons = 0
    Else
        .cRadioButtons = UBound(lRBN) + 1
        ReDim uRBN(UBound(lRBN))
        For i = 0 To UBound(lRBN)
            uRBN(i).nButtonID = lRBN(i)
            uRBN(i).pszButtonText = StrPtr(sRBN(i))
        Next i
    End If
    
    .pRadioButtons = VarPtr(uRBN(0).nButtonID)
    .CXWidth = cTD.Width
    .dwCommonButtons = cTD.CommonButtons
    .dwFlags = cTD.Flags
    .hInstance = cTD.hinst
    .hWndParent = cTD.ParenthWnd
    .nDefaultButton = cTD.DefaultButton
    .nDefaultRadioButton = cTD.DefaultRadioButton
    
    sz(0) = cTD.CollapsedControlText
    sz(1) = cTD.Content
    sz(2) = cTD.ExpandedControlText
    sz(3) = cTD.ExpandedInfo
    sz(4) = cTD.Footer
    sz(5) = cTD.MainInstruction
    sz(6) = cTD.VerifyText
    sz(7) = cTD.Title
    
    
If ((.dwFlags And TDF_INPUT_BOX) And (cTD.InputAlign = TDIBA_Footer)) Or _
   ((.dwFlags And TDF_COMBO_BOX) And (cTD.ComboAlign = TDIBA_Footer)) Or _
   ((.dwFlags And TDF_DATETIME) And (cTD.DateTimeAlign = TDIBA_Footer)) Or _
   ((.dwFlags And TDF_SLIDER) And (cTD.SliderAlign = TDIBA_Footer)) Then

    'ensure footer text isn't blank, or it won't appear
    If sz(4) = "" Then
        sz(4) = "$input"
    End If
End If
bEditFooter = False
If sz(4) = "$input" Then
    bEditFooter = True
    sz(4) = " "
'    uTDC.pszFooter = StrPtr(m_sFooter)
End If
If ((.dwFlags And TDF_INPUT_BOX) And (cTD.InputAlign = TDIBA_Content)) Or _
   ((.dwFlags And TDF_COMBO_BOX) And (cTD.ComboAlign = TDIBA_Content)) Or _
   ((.dwFlags And TDF_DATETIME) And (cTD.DateTimeAlign = TDIBA_Content)) Or _
   ((.dwFlags And TDF_SLIDER) And (cTD.SliderAlign = TDIBA_Content)) Then
        If sz(1) <> "" Then
            'If Right$(m_sContent, 3) <> (vbCrLf & " ") Then
'                m_sContent = m_sContent & vbCrLf & " "
'                nBreaksAdded = 1
                'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                    If Right$(sz(1), 5) <> (vbCrLf & vbCrLf & " ") Then
                        sz(1) = sz(1) & vbCrLf & vbCrLf
                        If (.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                            sz(1) = sz(1) & " "  '(link control requires a space or it ignores the double-break)
                                                           'Meanwhile, everything else gets thrown out of alignment, so ONLY add to link
                        End If
                        nBreaksAdded = 2
'                        uTDC.pszContent = StrPtr(m_sContent)
                    End If
                'End If
            'End If
        Else 'No content, main instruction only: add one break
            If Right$(sz(5), 2) <> vbCrLf Then
                sz(5) = sz(5) & vbCrLf
'                uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
            End If

        End If
        If sz(3) <> "" Then
            If Right$(sz(3), 2) <> vbCrLf Then
                sz(3) = sz(3) & vbCrLf
                'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                    If Right$(sz(3), 4) <> (vbCrLf & vbCrLf) Then
                        sz(3) = sz(3) & vbCrLf  '(it's already sure to have 1)
                    End If
                'End If
'                uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
            End If

        End If
End If
    
If (.dwFlags And TDF_CALLBACK_TIMER) Then
    bPublicTimer = True
Else
    If (.dwFlags And TDF_INPUT_BOX) Or (.dwFlags And TDF_COMBO_BOX) Or (.dwFlags And TDF_DATETIME) Or (.dwFlags And TDF_SLIDER) Then
        If m_sExpandedInfo <> "" Then
            If (.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                .dwFlags = .dwFlags Or TDF_CALLBACK_TIMER
            End If
        End If
    End If
    If (lAutoclose > 0) Then
        .dwFlags = .dwFlags Or TDF_CALLBACK_TIMER
    End If
End If
    
    
    
    
    
    .pszCollapsedControlText = StrPtr(sz(0))
    .pszContent = StrPtr(sz(1))
    .pszExpandedControlText = StrPtr(sz(2))
    .pszExpandedInformation = StrPtr(sz(3))
    .pszFooter = StrPtr(sz(4))
    .pszMainInstruction = StrPtr(sz(5))
    .pszVerificationText = StrPtr(sz(6))
    .pszWindowTitle = StrPtr(sz(7))
    
    If (.dwFlags And TDF_USE_HICON_MAIN) Then
        .pszMainIcon = cTD.IconMain
    Else
        .pszMainIcon = tdMakeIntResource(CLng(cTD.IconMain))
    End If
    If (.dwFlags And TDF_USE_HICON_FOOTER) Then
        .pszFooterIcon = cTD.IconFooter
    Else
        .pszFooterIcon = tdMakeIntResource(CLng(cTD.IconFooter))
    End If
    'If m_cbMain = 0& Then m_cbMain = AddressOf ProcessCallback 'scb_SetCallbackAddr(5, 2)
    '.pfCallback = m_cbMain
    .pfCallback = tdFARPROC(AddressOf TaskDialogCallbackProc)
    .lpCallbackData = ObjPtr(cTD)
End With

'Copy all user-set data from the new TD into this one, since it takes over
    bPosUsed(0) = False
    bPosUsed(1) = False
    bPosUsed(2) = False
    Me.PageIndex = cTD.PageIndex
    Me.Flags = cTD.Flags
    Me.InputAlign = cTD.InputAlign
    Me.InputAlignInFooter = cTD.InputAlignInFooter
    Me.InputBoxTextAlign = cTD.InputBoxTextAlign
    Me.InputCueBanner = cTD.InputCueBanner
    Me.InputIsPassword = cTD.InputIsPassword
    Me.InputOffsetX = cTD.InputOffsetX
    Me.InputText = cTD.InputText
    Me.InputWidth = cTD.InputWidth
    Me.AutocloseTime = cTD.AutocloseTime
    Me.ComboAlign = cTD.ComboAlign
    Me.ComboAlignInFooter = cTD.ComboAlignInFooter
    Me.ComboCueBanner = cTD.ComboCueBanner
    Me.ComboDropWidth = cTD.ComboDropWidth
    Me.ComboHeight = cTD.ComboHeight
    Me.ComboImageList = cTD.ComboImageList
    Me.ComboIndex = cTD.ComboIndex
    Me.ComboOffsetX = cTD.ComboOffsetX
    Me.ComboStyle = cTD.ComboStyle
    Dim cbT() As String, cbI() As Long, cbO() As Long, nct As Long
    cTD.zzComboGetItemData nct, cbT, cbI, cbO
    ReDim aComboItems(nct)
    For i = 0 To nct
        aComboItems(i).sText = cbT(i)
        aComboItems(i).iImage = cbI(i)
        aComboItems(i).iOverlay = cbO(i)
    Next i
    nComboInit = cTD.zzComboGetInitialItem()
    Dim s1 As String, n1 As Long, n2 As Long, b1 As Boolean
    zzComboGetInitialState s1, n1, n2, b1
    bCBInitValid = b1
    tComboInit.sText = s1: tComboInit.iImage = n1: tComboInit.iOverlay = n2
    Me.DateTimeAlign = cTD.DateTimeAlign
    Me.DateTimeAlignInButtons = cTD.DateTimeAlignInButtons
    Me.DateTimeAlignInContent = cTD.DateTimeAlignInContent
    Me.DateTimeAlignInFooter = cTD.DateTimeAlignInFooter
    Me.DateTimeChecked = cTD.DateTimeChecked
    Me.DateTimeOffsetX = cTD.DateTimeOffsetX
    Dim dt1() As Integer
    cTD.zzDateTimeGetInitial b1, dt1
    bDTSetInit = b1
    With tDTInitial
            .wDay = dt1(0)
            .wMonth = dt1(1)
            .wYear = dt1(2)
            .wHour = dt1(3)
            .wMinute = dt1(4)
            .wSecond = dt1(5)
    End With
    Dim dt2() As Integer, b2 As Boolean
    cTD.zzDateTimeGetRange b1, b2, dt1, dt2
    bDTSetMin = b1: bDTSetMax = b2
    With tDTRange(0)
            .wDay = dt1(0)
            .wMonth = dt1(1)
            .wYear = dt1(2)
            .wHour = dt1(3)
            .wMinute = dt1(4)
            .wSecond = dt1(5)
    End With
    With tDTRange(1)
            .wDay = dt2(0)
            .wMonth = dt2(1)
            .wYear = dt2(2)
            .wHour = dt2(3)
            .wMinute = dt2(4)
            .wSecond = dt2(5)
    End With
    Me.DateTimeType = cTD.DateTimeType
    lBtnHold = cTD.zzGetButtonHolds()
    lBtnElv = cTD.zzGetButtonElevated(b1)
    bElv = b1
'    icoDataN = cTD.zzGetCommonButtonIcon()
    cTD.zzGetLogoImage hLogoImage, nLogoType, nLogoPos, lLogoX, lLogoY
    nSplitButtonID = cTD.zzGetSplitButton
    m_BtnLRSize = cTD.zzGetWindowsButtonIconSize
    Me.SliderAlign = cTD.SliderAlign
    Me.SliderAlignInFooter = cTD.SliderAlignInFooter
    Me.SliderOffsetX = cTD.SliderOffsetX
    cTD.zzSliderGetChangeValues lSliderChangeS, lSliderChangeL
    cTD.zzSliderGetRange lSliderMin, lSliderMax, lSliderNTicks
    Me.SliderTickStyle = cTD.SliderTickStyle
    Me.SliderValue = cTD.SliderValue
    Me.SliderWidth = cTD.SliderWidth
     
    
    
    Call SendMessageW(m_TDhWnd, TDM_NAVIGATE_PAGE, 0, uTDC2)
    
End If
'<EhFooter>
Exit Function

e0:
    DebugAppend "cTaskDialog.NavigatePage->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Sub SetInitialButtonStates()
Dim i As Long
If bElv Then
    For i = 0 To UBound(lBtnElv)
        If lBtnElv(i) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, lBtnElv(i), ByVal 1&)
    Next i
End If
If bDis Then
    For i = 0 To UBound(lBtnDis)
        If lBtnDis(i) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_ENABLE_BUTTON, lBtnDis(i), ByVal 0&)
    Next i
End If
If bRadDis Then
    For i = 0 To UBound(lRadDis)
        If lRadDis(i) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_ENABLE_RADIO_BUTTON, lRadDis(i), ByVal 0&)
    Next i
End If
End Sub

Private Sub AddComboBox(nType As ComboType)
Dim lEditX As Long, lEditY As Long, cxEdit As Long, cyEdit As Long
Dim lOffX As Long
Dim tRC As RECT
m_HasCustomControl = True
GetClientRect m_TDhWnd, tRC
'DebugAppend "Combo y=" & lButtonY2
If nComboAlign = TDIBA_Footer Then
    If bPosUsed(2) Then Exit Sub
    bPosUsed(2) = True
    If cxCombo = -1 Then
        cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
    Else
        cxEdit = 180 * m_ScaleX
    End If
    If cxCombo > 0 Then cxEdit = cxCombo
    Select Case nCBFooterAlign
        Case tdcaLeft
            lEditX = (16 * m_ScaleX) + 16 '32
        Case tdcaCenter
            lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
        Case tdcaRight
            lEditX = tRC.Right - cxEdit - 10
    End Select
    If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
        lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY2 + 55
    Else
        lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY2 + 37
    End If
    If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
        lEditY = tRC.Bottom - ((58 * m_ScaleY) + 4)
'        DebugAppend "lEditY=" & lEditY
    End If
    If (m_sFooter = "") Or (m_sFooter = "$input") Then
        m_sFooter = " "
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    End If
ElseIf nComboAlign = TDIBA_Content Then
    If bPosUsed(0) Then Exit Sub
    bPosUsed(0) = True
    If m_Icon Then
        lOffX = (32 * m_ScaleX) + 20 '52
    Else
        lOffX = 12
    End If
    If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
        lOffX = 12
    End If
    lEditX = lOffX
    If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
        lEditY = lButtonY2 - (32 * m_ScaleY)
    Else
        If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
            lEditY = lButtonY2 - (40 * m_ScaleY)
        Else
            lEditY = lButtonY2 - (34 * m_ScaleY)
        End If
        If bLinkAdjust Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
    End If
    If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
        lEditY = lEditY - (24 * m_ScaleY)
    End If
    If (m_sExpandedInfo <> "") Then
        lEditY = lEditY - (6 * m_ScaleY)
    End If
    cxEdit = tRC.Right - lOffX - 12
    If cxCombo > 0 Then cxEdit = cxCombo
Else
    If bPosUsed(1) Then Exit Sub
    bPosUsed(1) = True
    lEditX = 10
    lEditY = lButtonY2
'    cxEdit = 160
    cxEdit = GetLeftmostCommonButtonX()
'    DebugAppend "Got Leftmost x=" & cxEdit
    cxEdit = cxEdit - lEditX - 10
    If cxEdit < 10 Then cxEdit = 10
    If cxCombo > 0 Then cxEdit = cxCombo
End If
If cyCombo > 0 Then
    cyEdit = cyCombo
Else
    cyEdit = 115 * m_ScaleY
End If

Dim dwStyle As Long
Dim hFont As LongPtr
Dim i As Long
        dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL Or WS_TABSTOP
        If nType = cbtDropdownEdit Then
            dwStyle = dwStyle Or CBS_DROPDOWN
        Else
            dwStyle = dwStyle Or CBS_DROPDOWNLIST
        End If
        
        hCombo = CreateWindowEx(0&, StrPtr(WC_COMBOBOXEX), StrPtr(""), _
                      dwStyle, _
                      lEditX, lEditY, cxEdit, cyEdit, _
                     hUI, IDD_COMBO, App.hInstance, ByVal 0&)
'        DebugAppend "hCombo=" & hCombo & ",x=" & lEditX & ",y=" & lEditY & ",cx=" & cxEdit
        Call SendMessageW(hCombo, CBEM_SETUNICODEFORMAT, 1, ByVal 0&)

        hEditCombo = SendMessageW(hCombo, CBEM_GETEDITCONTROL, 0, ByVal 0&)
        hComboCombo = SendMessageW(hCombo, CBEM_GETCOMBOCONTROL, 0, ByVal 0&)
        
        If cxCBDrop <> 0& Then
            Call SendMessageW(hCombo, CB_SETDROPPEDWIDTH, cxCBDrop, ByVal 0&)
        End If

        If m_sComboCue <> "" Then
            SendMessageW hEditCombo, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sComboCue)
        End If
        If himlCombo Then
            Call SendMessageW(hCombo, CBEM_SETIMAGELIST, 0, ByVal himlCombo)
        End If
        For i = 0 To UBound(aComboItems)
            CBX_InsertItem hCombo, aComboItems(i).sText, aComboItems(i).iImage, aComboItems(i).iOverlay
        Next i
        If bCBInitValid Then
            Dim cbxi As COMBOBOXEXITEMW
            cbxi.Mask = CBEIF_IMAGE Or CBEIF_OVERLAY
            cbxi.iItem = -1
            cbxi.iImage = tComboInit.iImage
            cbxi.iOverlay = tComboInit.iOverlay
            If tComboInit.sText <> "" Then
                cbxi.Mask = cbxi.Mask Or CBEIF_TEXT
                cbxi.pszText = StrPtr(tComboInit.sText)
            End If
            Call SendMessageW(hCombo, CBEM_SETITEMW, 0, cbxi)
        End If
        If nComboInit <> -1 Then
            Call SendMessageW(hCombo, CB_SETCURSEL, nComboInit, ByVal 0&)
        End If
        Subclass2 hCombo, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)



End Sub

Private Sub AddSlider()
Dim lEditX As Long, lEditY As Long, cxEdit As Long
Dim tRC As RECT
Dim lOffX As Long
m_HasCustomControl = True
GetClientRect m_TDhWnd, tRC
If nSliderAlign = TDIBA_Footer Then
    If bPosUsed(2) Then Exit Sub
    bPosUsed(2) = True
    If cxSlider = -1 Then
        cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
    Else
        cxEdit = 180 * m_ScaleX
    End If
    If cxSlider > 0 Then cxEdit = cxSlider
    Select Case nSLFooterAlign
    
        Case tdcaLeft
            lEditX = (16 * m_ScaleX) + 16 '32
        Case tdcaCenter
            lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
        Case tdcaRight
            lEditX = tRC.Right - cxEdit - 10
    End Select
    If nSliderTickStyle = SldTickStyleBoth Then
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5) 'lButtonY4 + 47
        Else
            lEditY = tRC.Bottom - ((30 * m_ScaleY)) 'lButtonY4 + 29 tRC.Bottom - 38
        End If
    Else
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5) 'lButtonY4 + 55
        Else
            lEditY = tRC.Bottom - ((30 * m_ScaleY)) ' lButtonY4 + 37 'tRC.Bottom - 30
        End If
    End If
    If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
        lEditY = tRC.Bottom - ((64 * m_ScaleY) + 5)
'        DebugAppend "lEditY=" & lEditY
    End If
    If (m_sFooter = "") Or (m_sFooter = "$input") Then
        m_sFooter = " "
'        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    End If

    If (m_sExpandedInfo <> "") Then
        lEditY = lEditY - (6 * m_ScaleY)
    End If
    Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
ElseIf nSliderAlign = TDIBA_Content Then
    If bPosUsed(0) Then Exit Sub
    bPosUsed(0) = True
    If m_Icon Then
        lOffX = (32 * m_ScaleX) + 20 '52
    Else
        lOffX = 12
    End If
    If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
        lOffX = 12
    End If
    lEditX = lOffX
    If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
        lEditY = lButtonY4 - (42 * m_ScaleY)
    Else
        If nSliderTickStyle = SldTickStyleBoth Then
            lEditY = lButtonY4 - (54 * m_ScaleY)
        Else
            lEditY = lButtonY4 - (44 * m_ScaleY)
        End If
        If bLinkAdjust Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
    End If
    If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
        lEditY = lEditY - (26 * m_ScaleY)
    End If

    cxEdit = tRC.Right - lOffX - 12
    If cxSlider > 0 Then cxEdit = cxSlider
Else
    If bPosUsed(1) Then Exit Sub
    bPosUsed(1) = True
    lEditX = 10
    If nSliderTickStyle = SldTickStyleBoth Then
        lEditY = lButtonY4 - 10
    Else
        lEditY = lButtonY4
    End If
'    cxEdit = 160
    cxEdit = GetLeftmostCommonButtonX()
'    DebugAppend "Got Leftmost x=" & cxEdit
    cxEdit = cxEdit - lEditX - 10
    If cxSlider > 0 Then cxEdit = cxSlider
    If cxEdit < 10 Then cxEdit = 10
End If
Dim dwStyle As Long
Dim hFont As LongPtr
Dim i As Long

dwStyle = WS_CHILD Or WS_VISIBLE Or TBS_AUTOTICKS Or TBS_HORZ Or TBS_TRANSPARENTBKGND
Select Case nSliderTickStyle
    Case SldTickStyleBottom
            dwStyle = dwStyle Or TBS_BOTTOM
    Case SldTickStyleTop
            dwStyle = dwStyle Or TBS_TOP
    Case SldTickStyleBoth
        dwStyle = dwStyle Or TBS_BOTH
    Case SldTickStyleNone
        dwStyle = dwStyle Or TBS_NOTICKS
End Select

hSlider = CreateWindowEx(0&, StrPtr("msctls_trackbar32"), StrPtr("Slider"), dwStyle, _
                         lEditX, lEditY, cxEdit, 36 * m_ScaleY, _
                         hUI, IDD_SLIDER, App.hInstance, ByVal 0&)
If hSlider Then
    SendMessageW hSlider, TBM_SETRANGEMIN, 0, ByVal lSliderMin
    SendMessageW hSlider, TBM_SETRANGEMAX, 1, ByVal lSliderMax
    SendMessageW hSlider, TBM_SETTICFREQ, lSliderNTicks, ByVal 0&
    SendMessageW hSlider, TBM_SETLINESIZE, 0, ByVal lSliderChangeS
    SendMessageW hSlider, TBM_SETPAGESIZE, 0, ByVal lSliderChangeL
    SendMessageW hSlider, TBM_SETPOS, 1&, ByVal lSliderValue
End If
'Subclass hSlider, AddressOf TaskDialogSubclassProc, hSlider, ObjPtr(Me)

End Sub
Private Sub AddDateTime()
Dim lEditX As Long, lEditY As Long, cxEdit As Long, cxEdit2 As Long
Dim lOffX As Long
Dim tRC As RECT
m_HasCustomControl = True
Dim dwUnits As Long
dwUnits = GetDeviceCaps(GetDC(0), LOGPIXELSX)
Dim baseX As Long
baseX = tdLoWord(dwUnits)

GetClientRect m_TDhWnd, tRC

Select Case nDTType

    Case dttDate
        cxEdit = 100
    Case dttDateWithCheck
        cxEdit = 116
    Case dttTime
        cxEdit = 90
    Case dttTimeWithCheck
        cxEdit = 106
    Case dttDateTime
        cxEdit = 100
        cxEdit2 = 90
    Case dttDateTimeWithCheck
        cxEdit = 116
        cxEdit2 = 106
    Case dttDateTimeWithCheckDateOnly
        cxEdit = 116
        cxEdit2 = 90
    Case dttDateTimeWithCheckTimeOnly
        cxEdit = 100
        cxEdit2 = 106
End Select
'cxEdit = cxEdit * m_ScaleX
'cxEdit2 = cxEdit2 * m_ScaleX
If nDateTimeAlign = TDIBA_Footer Then
    If bPosUsed(2) Then Exit Sub
    bPosUsed(2) = True
    Select Case nDTFooterAlign
    
        Case tdcaLeft
            lEditX = ((16 * m_ScaleX) + 16) '32
        Case tdcaCenter
            lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
        Case tdcaRight
            lEditX = tRC.Right - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 10
    End Select
    If lEditX < ((16 * m_ScaleX) + 16) Then lEditX = ((16 * m_ScaleX) + 16)
    If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            'TODO: ADJUST FOR EXPAND-TO-FOOTER
        lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY3 + (23 * m_ScaleY) + 22 '58
    Else
'        lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'tRC.Bottom - (40 * m_ScaleY) '(tRC.Bottom - (4 * m_ScaleY)) - ((24 * m_ScaleY)) 'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) = TDF_KILL_SHIELD_ICON Then
            If (m_ScaleY > 1) Then
                lEditY = tRC.Bottom - ((40 * m_ScaleY) + 4)
            Else
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)
            End If
        Else
         
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)
        End If
    End If
    If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) = TDF_KILL_SHIELD_ICON Then
            If (m_ScaleY > 1) Then
                lEditY = tRC.Bottom - ((40 * m_ScaleY) + 4)
            Else
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)
            End If
        Else
         
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)
        End If
'        DebugAppend "lEditY=" & lEditY
    End If
    'NOTE: Certain flags on certain DPIs may not align right with the DateTime control in the footer area.
    
    If (m_sFooter = "") Or (m_sFooter = "$input") Then
        m_sFooter = " "
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    End If
ElseIf nDateTimeAlign = TDIBA_Content Then
    If bPosUsed(0) Then Exit Sub
    bPosUsed(0) = True
    If m_Icon Then
        lOffX = (32 * m_ScaleX) + 20 '52 + 16 '* m_ScaleX
    Else
        lOffX = 12
    End If
    If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
        lOffX = 12
    End If
    Select Case nDTContentAlign
    
        Case tdcaLeft
            lEditX = lOffX
        Case tdcaCenter
            lEditX = lOffX + (((tRC.Right - lOffX - 12) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
        Case tdcaRight
            lEditX = (tRC.Right) - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 12
    
    End Select
    If lEditX < lOffX Then lEditX = lOffX
    If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
        lEditY = lButtonY3 - (34 * m_ScaleY)
'        DebugAppend "Set DT Y=" & lEditY
    Else
        If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
            lEditY = lButtonY3 - (44 * m_ScaleY)
        Else
            lEditY = lButtonY3 - (34 * m_ScaleY)
        End If
        If bLinkAdjust Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
    End If
    If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
        lEditY = lEditY - (22 * m_ScaleY)
    End If
    If (m_sExpandedInfo <> "") Then
        lEditY = lEditY - (6 * m_ScaleY)
    End If
    
Else 'Button align
'    If bPosUsed(1) Then Exit Sub
    bPosUsed(1) = True
    Dim lBtn As Long
    lBtn = GetLeftmostCommonButtonX()
    lOffX = 10
    Select Case nDTButtonAlign
    
        Case tdcaLeft
            lEditX = lOffX
        Case tdcaCenter
            lEditX = lOffX + (((lBtn - lOffX - 10) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
        Case tdcaRight
            lEditX = (lBtn) - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 10
    
    End Select
    lEditY = lButtonY3
End If



Dim dwStyle As Long
Dim hFont As LongPtr
Dim i As Long
lEditX = lEditX + DTOffX

        dwStyle = WS_CHILD Or WS_OVERLAPPED Or WS_VISIBLE
        If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then dwStyle = dwStyle Or DTS_SHOWNONE
    
        If (nDTType = dttDate) Or (nDTType = dttDateWithCheck) Or (nDTType >= dttDateTime) Then
            dwStyle = dwStyle Or DTS_SHORTDATEFORMAT
        Else
            dwStyle = dwStyle Or DTS_TIMEFORMAT
        End If
        hDateTime = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, _
                                   StrPtr("SysDateTimePick32"), StrPtr(""), _
                                   dwStyle, _
                                   lEditX, lEditY, cxEdit * m_ScaleX, 24 * m_ScaleY, _
                                   hUI, IDD_DT1, App.hInstance, ByVal 0&)

        If (nDTType = dttDateTime) Or (nDTType >= dttDateTimeWithCheck) Then
            dwStyle = WS_CHILD Or WS_OVERLAPPED Or WS_VISIBLE Or DTS_TIMEFORMAT
            If (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckTimeOnly) Then dwStyle = dwStyle Or DTS_SHOWNONE
            
            hDateTime2 = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, _
                                       StrPtr("SysDateTimePick32"), StrPtr(""), _
                                       dwStyle, _
                                       (lEditX + ((cxEdit + 5) * m_ScaleX)), lEditY, cxEdit2 * m_ScaleX, 24 * m_ScaleY, _
                                       hUI, IDD_DT2, App.hInstance, ByVal 0&)
        End If
        
        hFont = GetStockObject(DEFAULT_GUI_FONT)
        If hFont Then
            SendMessageW hDateTime, WM_SETFONT, hFont, ByVal 0&
            If hDateTime2 Then
                SendMessageW hDateTime2, WM_SETFONT, hFont, ByVal 0&
            End If
        End If
        
        If (bDTSetMin = True) Or (bDTSetMax = True) Then
            SetDTRange
        End If
        If bDTSetInit Then
            SendMessageW hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTInitial
            If hDateTime2 Then
                SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTInitial
            End If
        End If
        If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
            If bDTCheck = False Then
                SendMessageW hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
            End If
        End If
        If (nDTType = dttDateTimeWithCheck) Then
            If bDTCheck2 = False Then
                SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
            End If
        End If
        If (nDTType = dttDateTimeWithCheckTimeOnly) Then
            If bDTCheck = False Then
                SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
            End If
        End If
End Sub
Private Sub AddInputBox()
Dim lEditX As Long, lEditY As Long, cxEdit As Long
Dim lOffX As Long
Dim tRC As RECT
m_HasCustomControl = True
GetClientRect m_TDhWnd, tRC


'DebugAppend "lButtonY=" & lButtonY & ",btm=" & tRC.Bottom
If nInputAlign = TDIBA_Footer Then
    If bPosUsed(2) Then Exit Sub
    bPosUsed(2) = True
    If cxInputBox = -1 Then
        cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
    Else
        cxEdit = 180 * m_ScaleX
    End If
    If cxInputBox > 0 Then cxEdit = cxInputBox
    Select Case nIBFooterAlign
    
        Case tdcaLeft
            lEditX = (16 * m_ScaleX) + 16 '32
        Case tdcaCenter
            lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
        Case tdcaRight
            lEditX = tRC.Right - (cxEdit) - 10
    End Select    'lEditY = tRC.Bottom - 27
    If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
      lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY3 + (23 * m_ScaleY) + 22 '58
    Else
        lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
    End If
    If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
        lEditY = tRC.Bottom - ((58 * m_ScaleY) + 4)
    End If
    If (m_sFooter = "") Or (m_sFooter = "$input") Then
        m_sFooter = " "
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    End If
ElseIf nInputAlign = TDIBA_Content Then
    If bPosUsed(0) Then Exit Sub
    bPosUsed(0) = True
    If m_Icon Then
        lOffX = (32 * m_ScaleX) + 20 '52
    Else
        lOffX = 12
    End If
    If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
        lOffX = 12
    End If
    lEditX = lOffX
    If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
        lEditY = lButtonY - (34 * m_ScaleY)
    Else
        If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
            lEditY = lButtonY - (44 * m_ScaleY)
        Else
            lEditY = lButtonY - (34 * m_ScaleY)
        End If
        If bLinkAdjust Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
    End If
    If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
        lEditY = lEditY - (22 * m_ScaleY)
    End If
    If (m_sExpandedInfo <> "") Then
        lEditY = lEditY - (6 * m_ScaleY)
    End If
    cxEdit = tRC.Right - lOffX - 12
    If cxInputBox > 0 Then cxEdit = cxInputBox
Else 'TDIBA_Buttons
    If bPosUsed(1) Then Exit Sub
    bPosUsed(1) = True
    lEditX = 10
    lEditY = lButtonY
'    cxEdit = 160
    cxEdit = GetLeftmostCommonButtonX()
'    DebugAppend "Got Leftmost x=" & cxEdit
    cxEdit = cxEdit - lEditX - 10
    If cxEdit < 10 Then cxEdit = 10
    If cxInputBox > 0 Then cxEdit = cxInputBox
End If
Dim dwStyle As Long
Dim hFont As LongPtr
Dim i As Long
        dwStyle = WS_VISIBLE Or WS_CHILD Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL
        Select Case mInputTextAlign
            Case TDIBTA_Left: dwStyle = dwStyle Or ES_LEFT
            Case TDIBTA_Right: dwStyle = dwStyle Or ES_RIGHT
            Case TDIBTA_Center: dwStyle = dwStyle Or ES_CENTER
        End Select
        If bEditPassword Then
            dwStyle = dwStyle Or ES_PASSWORD
        End If
    
        
        hEditBox = CreateWindowEx(WS_EX_CLIENTEDGE, StrPtr("Edit"), StrPtr(""), _
                      dwStyle, _
                      lEditX, lEditY, cxEdit, 22 * m_ScaleY, _
                     hUI, IDD_EDIT1, App.hInstance, ByVal 0&)
                     


        hFont = GetStockObject(DEFAULT_GUI_FONT)
        If hFont Then
            SendMessageW hEditBox, WM_SETFONT, hFont, ByVal 0&
        End If
        If m_sInput <> "" Then
            SendMessageW hEditBox, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sInput)
        End If
        If sEditText <> "" Then
            SendMessageW hEditBox, WM_SETTEXT, 0&, ByVal StrPtr(sEditText)
            SendMessageW hEditBox, EM_SETSEL, 0&, ByVal (-1&)
        End If
        

End Sub

Public Sub SetLogoImage(hImage As LongPtr, ImageType As LogoImageType, ImagePosition As LogoPosition, Optional OffsetX As Long = 0, Optional OffsetY As Long = 0)
hLogoImage = hImage
nLogoType = ImageType
nLogoPos = ImagePosition
lLogoX = OffsetX
lLogoY = OffsetY
End Sub
Public Sub zzGetLogoImage(a As LongPtr, b As Long, c As Long, d As Long, e As Long)
a = hLogoImage: b = nLogoType: c = nLogoPos: d = lLogoX: e = lLogoY
End Sub

Private Sub AddLogo()
Dim tBMP As BITMAP
Dim CX As Long, CY As Long
Dim px As Long, py As Long
Dim tRC As RECT

GetClientRect m_TDhWnd, tRC
Call GetObject(hLogoImage, LenB(tBMP), tBMP)
CX = tBMP.BMWidth
CY = tBMP.BMHeight
If nLogoPos = LogoTopRight Then
    
    px = tRC.Right - CX = lLogoX
    py = lLogoY
ElseIf nLogoPos = LogoButtons Then
    px = lLogoX
    py = (lButtonY - 6) + lLogoY
End If
lgpX = px
lgpY = py
hStatic = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(""), _
            WS_VISIBLE Or WS_CHILD Or SS_BITMAP, _
            px, py, CX, CY, hUI, 110, App.hInstance, ByVal 0&)
If hStatic Then
    SendMessageW hStatic, STM_SETIMAGE, nLogoType, ByVal hLogoImage
End If
End Sub
Public Sub SetWindowsButtonIconSize(SIZE As eWinIcoSize)
m_BtnLRSize = SIZE
End Sub
Public Function zzGetWindowsButtonIconSize() As Long
zzGetWindowsButtonIconSize = m_BtnLRSize
End Function
Private Function SetButtonIcons() As Boolean
Dim i As Long, J As Long, k As Long
Dim lOffset As Long
Dim hButton As LongPtr
Dim stay As Boolean
Dim hIco As LongPtr
Dim sID As String
sID = "23"
On Error Resume Next
If m_HasRadioButtons Then
    lOffset = UBound(m_uRadioButtons) + 1
End If

If m_HasCustomButtons Then
    For i = 0 To UBound(icoDataC)
        hButton = m_hButton(i + lOffset)
        If icoDataC(i) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataC(i), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataC(i), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataC(i)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If

        End If
        If m_uButtons(i).nButtonID = nSplitButtonID Then
            Call Subclass2(hButton, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me))
'                SetParent hButton, m_TDhWnd
            MakeSplitButton hButton
        End If
    Next i
    J = UBound(m_uButtons) + 1
End If

For i = 0 To 5
    If icoDataN(i) Then
        stay = True
    End If
Next i
If (stay = False) Or (uTDC.dwCommonButtons = 0) Then Exit Function
i = 0
If (uTDC.dwCommonButtons And TDCBF_OK_BUTTON) Then
    hButton = m_hButton(J + lOffset)
    If icoDataN(0) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(0), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(0), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
'        Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal icoDataN(0))
    End If
    k = k + 1
End If
If (uTDC.dwCommonButtons And TDCBF_YES_BUTTON) Then
    hButton = m_hButton(J + lOffset + k)
    If icoDataN(1) Then
        If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
            hIco = LoadImage(hShell32, icoDataN(1), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
            hIco = LoadImage(hImageRes, icoDataN(1), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        Else
            hIco = icoDataN(1)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
        End If
    End If
    k = k + 1
End If
If (uTDC.dwCommonButtons And TDCBF_NO_BUTTON) Then
    hButton = m_hButton(J + lOffset + k)
    If icoDataN(2) Then
        If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
            hIco = LoadImage(hShell32, icoDataN(2), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
            hIco = LoadImage(hImageRes, icoDataN(2), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        Else
            hIco = icoDataN(2)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
        End If
    End If
    k = k + 1
End If
If (uTDC.dwCommonButtons And TDCBF_CANCEL_BUTTON) Then
    hButton = m_hButton(J + lOffset + k)
    If icoDataN(3) Then
        If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
            hIco = LoadImage(hShell32, icoDataN(3), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
            hIco = LoadImage(hImageRes, icoDataN(3), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        Else
            hIco = icoDataN(3)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
        End If
    End If
    k = k + 1
End If
If (uTDC.dwCommonButtons And TDCBF_RETRY_BUTTON) Then
    hButton = m_hButton(J + lOffset + k)
    If icoDataN(4) Then
        If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
            hIco = LoadImage(hShell32, icoDataN(4), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
            hIco = LoadImage(hImageRes, icoDataN(4), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        Else
            hIco = icoDataN(4)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
        End If
    End If
    k = k + 1
End If
If (uTDC.dwCommonButtons And TDCBF_CLOSE_BUTTON) Then
    hButton = m_hButton(J + lOffset + k)
    If icoDataN(5) Then
        If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
            hIco = LoadImage(hShell32, icoDataN(5), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
            hIco = LoadImage(hImageRes, icoDataN(5), 1, m_BtnLRSize, m_BtnLRSize, 0)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            Call DestroyIcon(hIco)
        Else
            hIco = icoDataN(5)
            Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
        End If
    End If
    k = k + 1
End If
End Function
Private Sub InputReposition()
If hEditBox Then
        Dim tRC As RECT
        Dim lOffX As Long
        lButtonY = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(m_TDhWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        GetClientRect hUI, tRC
        Dim lEditX As Long, lEditY As Long, cxEdit As Long
    If nInputAlign = TDIBA_Content Then
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20 '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY - (34 * m_ScaleY)
        Else
            'If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY - (44 * m_ScaleY)
            'Else
            '    lEditY = lButtonY - (34 * m_ScaleY)
            'End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (22 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        cxEdit = tRC.Right - lOffX - 12
    ElseIf nInputAlign = TDIBA_Footer Then
'        lEditX = 32
'        lEditY = tRC.Bottom - 27
'        cxEdit = 180
        If cxInputBox = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxInputBox > 0 Then cxEdit = cxInputBox
        Select Case nIBFooterAlign
        
            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16 '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - (cxEdit) - 10
        End Select    'lEditY = tRC.Bottom - 27
         If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                  'TODO: ADJUST FOR EXPAND-TO-FOOTER
          lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY3 + (23 * m_ScaleY) + 22 '58
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
        End If
       
    End If
    If (nInputAlign = TDIBA_Content) Or (nInputAlign = TDIBA_Footer) Then
        SetWindowPos hEditBox, 0&, lEditX, lEditY, cxEdit, 22&, SWP_NOSIZE
        RepaintWindow hEditBox
    End If
    SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (28 * m_ScaleY), SWP_NOMOVE
End If
End Sub
Private Sub ComboReposition()
If hCombo Then
        Dim lOffX As Long
        Dim tRC As RECT
        lButtonY2 = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(m_TDhWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        GetClientRect hUI, tRC
        Dim lEditX As Long, lEditY As Long, cxEdit As Long
    If nComboAlign = TDIBA_Content Then
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20 '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY2 - (28 * m_ScaleY)
        Else
'            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY2 - (44 * m_ScaleY)
'            Else
'                lEditY = lButtonY2 - 34
'            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (24 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        cxEdit = tRC.Right - lOffX - 12
    ElseIf nComboAlign = TDIBA_Footer Then
'        lEditX = 32
'        lEditY = tRC.Bottom - 30
'        cxEdit = 180
        If cxCombo = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxCombo > 0 Then cxEdit = cxCombo
        Select Case nCBFooterAlign
            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16 '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - cxEdit - 10
        End Select
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY2 + 55
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4) 'lButtonY2 + 37
        End If

    End If
    If (nComboAlign = TDIBA_Content) Or (nComboAlign = TDIBA_Footer) Then
        SetWindowPos hCombo, 0&, lEditX, lEditY, cxEdit, 115&, SWP_NOSIZE
        RepaintWindow hCombo
    End If
    SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (28 * m_ScaleY), SWP_NOMOVE
End If
End Sub

Private Sub DateTimeReposition()
If hDateTime Then
        Dim tRC As RECT
        lButtonY3 = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(m_TDhWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        GetClientRect hUI, tRC
        Dim lEditX As Long, lEditY As Long, cxEdit As Long, cxEdit2 As Long
    
    Select Case nDTType
    
        Case dttDate
            cxEdit = 100
        Case dttDateWithCheck
            cxEdit = 116
        Case dttTime
            cxEdit = 90
        Case dttTimeWithCheck
            cxEdit = 106
        Case dttDateTime
            cxEdit = 100
            cxEdit2 = 90
        Case dttDateTimeWithCheck
            cxEdit = 116
            cxEdit2 = 106
        Case dttDateTimeWithCheckDateOnly
            cxEdit = 116
            cxEdit2 = 90
        Case dttDateTimeWithCheckTimeOnly
            cxEdit = 100
            cxEdit2 = 106
    End Select
    Dim lOffX As Long
    If nDateTimeAlign = TDIBA_Content Then
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20
        Else
            lOffX = 12
        End If
        Select Case nDTContentAlign
        
            Case tdcaLeft
                lEditX = lOffX
            Case tdcaCenter
'                DebugAppend "repos width=" & tRC.Right & ",cx1=" & cxEdit & ",cx2=" & cxEdit2
'                DebugAppend "repos x=" & (((tRC.Right - 52 - 12) - (cxEdit + cxEdit2 + 5)) \ 2)
                lEditX = lOffX + (((tRC.Right - lOffX - 12) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
            Case tdcaRight
                lEditX = (tRC.Right) - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 12
        
        End Select
        If lEditX < lOffX Then lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY3 - (34 * m_ScaleY)
'            DebugAppend "Set DT Y=" & lEditY
        Else
'            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY3 - (44 * m_ScaleY)
'            Else
'                lEditY = lButtonY3 - 34
'            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (22 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If

    ElseIf nDateTimeAlign = TDIBA_Footer Then
'        lEditX = 32
'        lEditY = tRC.Bottom - 30
'        m_sFooter = " "
        Select Case nDTFooterAlign
        
            Case tdcaLeft
                lEditX = ((16 * m_ScaleX) + 16) '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 10
        End Select
        If lEditX < ((16 * m_ScaleX) + 16) Then lEditX = ((16 * m_ScaleX) + 16)
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            'TODO: ADJUST FOR EXPAND-TO-FOOTER
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 6) 'lButtonY3 + (23 * m_ScaleY) + 22 '58
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 6) 'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
        End If
    
    End If
'If (nDTType = dttTime) Then
'    cxEdit = 90
'Else
'    cxEdit = 100
'End If


    If (nDateTimeAlign = TDIBA_Content) Or (nDateTimeAlign = TDIBA_Footer) Then
        SetWindowPos hDateTime, 0&, lEditX, lEditY, cxEdit * m_ScaleX, 24 * m_ScaleY, SWP_NOSIZE
        RepaintWindow hDateTime
        If hDateTime2 Then
            SetWindowPos hDateTime2, 0&, (lEditX + ((cxEdit + 5) * m_ScaleX)), lEditY, cxEdit2 * m_ScaleX, 24 * m_ScaleY, SWP_NOSIZE
            RepaintWindow hDateTime2
        End If
    End If
    SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + 30 * m_ScaleY, SWP_NOMOVE
End If
End Sub
Private Sub SliderReposition()
If hSlider Then
        Dim tRC As RECT
        Dim lOffX As Long
        lButtonY4 = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(m_TDhWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        GetClientRect hUI, tRC
        Dim lEditX As Long, lEditY As Long, cxEdit As Long
    If nSliderAlign = TDIBA_Content Then
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20 '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY4 - (42 * m_ScaleY)
        Else
            If nSliderTickStyle = SldTickStyleBoth Then
                lEditY = lButtonY4 - (54 * m_ScaleY)
            Else
                lEditY = lButtonY4 - (44 * m_ScaleY)
            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (22 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        cxEdit = tRC.Right - lOffX - 12
        If cxSlider > 0 Then cxEdit = cxSlider
    ElseIf nSliderAlign = TDIBA_Footer Then
'        lEditX = 32
'        If nSliderTickStyle = SldTickStyleBoth Then
'            lEditY = tRC.Bottom - 38
'        Else
'            lEditY = tRC.Bottom - 30
'        End If
'        cxEdit = 180
'        If cxSlider > 0 Then cxEdit = cxSlider
        If cxSlider = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32) '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxSlider > 0 Then cxEdit = cxSlider
        Select Case nSLFooterAlign
        
            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16 '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - cxEdit - 10
        End Select
        If nSliderTickStyle = SldTickStyleBoth Then
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5) 'lButtonY4 + 47
            Else
                lEditY = tRC.Bottom - ((30 * m_ScaleY)) 'lButtonY4 + 29 tRC.Bottom - 38
            End If
        Else
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5) 'lButtonY4 + 55
            Else
                lEditY = tRC.Bottom - ((30 * m_ScaleY)) ' lButtonY4 + 37 'tRC.Bottom - 30
            End If
        End If
        If (m_sFooter = "") Or (m_sFooter = "$input") Then
            m_sFooter = " "
    '        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
        End If
    
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If

    End If
    If (nSliderAlign = TDIBA_Content) Or (nSliderAlign = TDIBA_Footer) Then
        SetWindowPos hSlider, 0&, lEditX, lEditY, cxEdit, 36&, SWP_NOSIZE
        RepaintWindow hSlider

    End If
    SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (42 * m_ScaleY), SWP_NOMOVE

End If
End Sub
Private Sub RepaintWindow(hWnd As LongPtr)
Dim tR As RECT
Dim tP As POINTAPI
GetClientRect hWnd, tR
InvalidateRect hWnd, tR, 1
End Sub
Private Function CBX_InsertItem(hCB As LongPtr, sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1, Optional lParam As LongPtr = 0, Optional iItem As Long = -1) As Long
'DebugAppend "CBX_InsertItem " & sText & ", " & iImage
Dim cbxi As COMBOBOXEXITEMW

With cbxi
    .Mask = CBEIF_TEXT Or CBEIF_LPARAM Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_INDENT Or CBEIF_LPARAM
    .cchTextMax = Len(sText)
    .pszText = StrPtr(sText)
    .iItem = iItem
    .iImage = iImage
    .iSelectedImage = iImage
    .iOverlay = iOverlay
    .lParam = lParam
End With
CBX_InsertItem = CLng(SendMessageW(hCB, CBEM_INSERTITEMW, 0, cbxi))
End Function
Private Function GetComboTextW(hEdit As LongPtr) As String
Dim bytS() As Byte
Dim ch As Long
ch = CLng(SendMessageW(hEdit, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2 + 1)
ReDim Preserve bytS(ch)
ch = CLng(SendMessageW(hEdit, WM_GETTEXT, ch, ByVal VarPtr(bytS(0)))) * 2 - 1
If ch >= 0 Then
    ReDim Preserve bytS(ch)
End If
GetComboTextW = CStr(bytS)

End Function

Private Function MakeSplitButton(hWnd As LongPtr) As Long
'adds the drop down menu button to a command button
'Subclassing is required to detect a click
'However the PARENT received the BCN_DROPDOWN message

Dim dwStyle As Long

dwStyle = CLng(GetWindowLong(hWnd, GWL_STYLE))
dwStyle = dwStyle Or BS_SPLITBUTTON

Call SetWindowLong(hWnd, GWL_STYLE, dwStyle)
Call UpdateWindow(hWnd)

End Function
Private Sub SortLongArray(paintArray() As Long)
' This sub uses the Bubble Sort algorithm to sort an array of integers.
Dim lngY As Long
Dim intTemp As Long
Dim blnExchangeMade As Boolean

blnExchangeMade = True

Do While blnExchangeMade
    blnExchangeMade = False

    For lngY = LBound(paintArray) To (UBound(paintArray) - 1)
        If paintArray(lngY) > paintArray(lngY + 1) Then
            intTemp = paintArray(lngY)
            paintArray(lngY) = paintArray(lngY + 1)
            paintArray(lngY + 1) = intTemp
            blnExchangeMade = True
        End If
    Next
Loop
End Sub

Private Sub FocusMainControl()
'when custom controls are added, determines which one should get focus
'Priorities:
'Main=manually specified; else 1st=content, 2nd=buttons, 3rd=footer
If nDefaultCustCtrl > 0 Then
    Select Case nDefaultCustCtrl
        Case tdccInputBox: SetFocusAPI hEditBox
        Case tdccComboBox: SetFocusAPI hCombo
        Case tdccDateTime: SetFocusAPI hDateTime
        Case tdccSlider: SetFocusAPI hSlider
        Case tdccNothing: Exit Sub
    End Select
    Exit Sub
End If

If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Content)) Then
    SetFocusAPI hEditBox
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Content)) Then
    SetFocusAPI hCombo
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Content)) Then
    SetFocusAPI hDateTime
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Content)) Then
    SetFocusAPI hSlider
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Buttons)) Then
    SetFocusAPI hEditBox
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Buttons)) Then
    SetFocusAPI hCombo
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Buttons)) Then
    SetFocusAPI hDateTime
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Buttons)) Then
    SetFocusAPI hSlider
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Footer)) Then
    SetFocusAPI hEditBox
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Footer)) Then
    SetFocusAPI hCombo
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Footer)) Then
    SetFocusAPI hDateTime
    Exit Sub
End If
If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Footer)) Then
    SetFocusAPI hSlider
    Exit Sub
End If
End Sub

Private Function GetLeftmostCommonButtonX() As Long
'arBtnX stores all button X positions, but this includes command links
'and radio buttons. To size button-area controls, we need the furthest
'left position of only the buttons that appear in the button area.
Dim nButtonsNA As Long 'number of n/a buttons (radio+cmdlink)
Dim i As Long, J As Long
Dim arValid() As Long
Dim CX As Long

If m_HasRadioButtons Then
    nButtonsNA = UBound(m_uRadioButtons) + 1
End If
If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
    If m_HasCustomButtons Then
        nButtonsNA = nButtonsNA + UBound(m_uButtons) + 1
    Else
        'if there's NO buttons on the bottom, return -1 to indicate that
        'the caller should then use the dialog width or a standard width
        GetLeftmostCommonButtonX = -1
        Exit Function
    End If
End If
DebugAppend "Got nButtonsNA=" & nButtonsNA
If nButtonsNA > 0 Then
    SortLongArray arBtnX 'radio buttons and command links will be the smallest
                         'X values (for scenarios where there's even room in the
                         'button area for the control), so we can eliminate the
                         'nButtonsNA smallest values and this translates to
                         'eliminating the N/A buttons.
    ReDim arValid(UBound(arBtnX) - nButtonsNA)
    For i = nButtonsNA To UBound(arBtnX)
        arValid(J) = arBtnX(i)
        J = J + 1
    Next i
Else
    arValid = arBtnX
End If
If UBound(arValid) = 0 Then
    GetLeftmostCommonButtonX = arValid(0)
    Exit Function
End If
CX = arValid(0)
For i = 1 To UBound(arValid)
    If arValid(i) < CX Then
        CX = arValid(i)
    End If
Next i
GetLeftmostCommonButtonX = CX
End Function

Private Function GetBottomDividerPosA() As Long
'Dim tRC As RECT
'Dim pxBase As Long
'Dim pxTest As Long
'Dim yStart As Long
'Dim hDC As Long
'Dim i As Long
'GetClientRect hUI, tRC
'hDC = GetDC(hUI)
'
'yStart = tRC.Bottom - 4
'
'pxBase = GetPixel(hDC, 3, yStart + 1)
'
'For i = yStart To 0 Step -1
'    pxTest = GetPixel(hDC, 3, yStart - i)
'    If pxTest <> pxBase Then 'first color change is the divider line
'        GetBottomDividerPosA = i + 4
'        Exit For
'    End If
'Next i
'ReleaseDC hDC
'DebugAppend "Got div pos=" & GetBottomDividerPosA
End Function

Private Function Subclass2(hWnd As LongPtr, lpFN As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
If uId = 0 Then uId = hWnd
    Subclass2 = SetWindowSubclass(hWnd, lpFN, uId, dwRefData):      Debug.Assert Subclass2
End Function

Private Function UnSubclass2(hWnd As LongPtr, ByVal lpFN As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpFN, pid)
End Function

'@3
Public Function zz_ProcessEnumCallback(ByVal hWnd As LongPtr) As Long
Dim ub As Long
Dim sClass As String
Dim np As Long

sClass = Space$(64)
Call GetClassName(hWnd, sClass, 64)
np = InStr(sClass, Chr$(0))
If np > 1 Then
  sClass = Left(sClass, np - 1)
End If
If sClass = "DirectUIHWND" Then
    hUI = hWnd
End If
If sClass = "Button" Then
        Dim tRC As RECT, pt As POINTAPI
        GetWindowRect hWnd, tRC
        pt.X = tRC.Left
        pt.Y = tRC.Top
        ScreenToClient m_TDhWnd, pt
        If arBtnX(0) = 0 Then
            arBtnX(0) = pt.X
        Else
            ReDim Preserve arBtnX(UBound(arBtnX) + 1)
            arBtnX(UBound(arBtnX)) = pt.X
        End If
'        DebugAppend "Button top=" & pt.Y
        If True Then
        'If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
            If nInputAlign = TDIBA_Content Then
                If lButtonY = 0 Then
                    lButtonY = pt.Y
                    lButtonX = pt.X
                End If
                If (pt.Y > 10) And ((pt.Y < lButtonY) And (lButtonY <> 0)) Then
                    lButtonY = pt.Y
                    lButtonX = pt.X
                End If
            ElseIf (nInputAlign = TDIBA_Buttons) Or (nInputAlign = TDIBA_Footer) Then
                If (pt.Y > lButtonY) Then
                    lButtonX = pt.X
                    lButtonY = pt.Y
                End If
            End If
            If nComboAlign = TDIBA_Content Then
                If lButtonY2 = 0 Then
                    lButtonY2 = pt.Y
                    lButtonX2 = pt.X
                End If
                If (pt.Y > 10) And ((pt.Y < lButtonY2) And (lButtonY2 <> 0)) Then
                    lButtonY2 = pt.Y
                    lButtonX2 = pt.X
                End If
            ElseIf (nComboAlign = TDIBA_Buttons) Or (nComboAlign = TDIBA_Footer) Then
                If (pt.Y > lButtonY2) Then
                    lButtonX2 = pt.X
                    lButtonY2 = pt.Y
                End If
            End If
            If nDateTimeAlign = TDIBA_Content Then
                If lButtonY3 = 0 Then
                    lButtonY3 = pt.Y
                    lButtonX3 = pt.X
                End If
                If (pt.Y > 10) And ((pt.Y < lButtonY3) And (lButtonY3 <> 0)) Then
                    lButtonY3 = pt.Y
                    lButtonX3 = pt.X
                End If
            ElseIf (nDateTimeAlign = TDIBA_Buttons) Or (nDateTimeAlign = TDIBA_Footer) Then
                If (pt.Y > lButtonY3) Then
                    lButtonX3 = pt.X
                    lButtonY3 = pt.Y
                End If
            End If
            If nSliderAlign = TDIBA_Content Then
                If lButtonY4 = 0 Then
                    lButtonY4 = pt.Y
                    lButtonX4 = pt.X
                End If
                If (pt.Y > 10) And ((pt.Y < lButtonY4) And (lButtonY4 <> 0)) Then
                    lButtonY4 = pt.Y
                    lButtonX4 = pt.X
                End If
            ElseIf (nSliderAlign = TDIBA_Buttons) Or (nSliderAlign = TDIBA_Footer) Then
                If (pt.Y > lButtonY4) Then
                    lButtonX4 = pt.X
                    lButtonY4 = pt.Y
                End If
            End If
        Else
            If pt.Y > lButtonY Then
                lButtonX = pt.X
                lButtonY = pt.Y
            End If
            If pt.Y > lButtonY2 Then
                lButtonX2 = pt.X
                lButtonY2 = pt.Y
            End If
            If pt.Y > lButtonY3 Then
                lButtonX3 = pt.X
                lButtonY3 = pt.Y
            End If
            If pt.Y > lButtonY4 Then
                lButtonX4 = pt.X
                lButtonY4 = pt.Y
            End If
        End If
    ub = UBound(m_hButton)
    If (ub = 0) And (m_hButton(0) = 0) Then
        m_hButton(0) = hWnd
    Else
        ub = ub + 1
        ReDim Preserve m_hButton(ub)
        m_hButton(ub) = hWnd
    End If
End If

zz_ProcessEnumCallback = 1
End Function


'@2
Public Function zz_ProcessCallback(ByVal hWnd As LongPtr, ByVal uNotification As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
'***DO NOT CALL THIS FUNCTION***
Dim bSC As Boolean

Select Case uNotification
    Case TDN_DIALOG_CONSTRUCTED
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) Then
            SendMessageW hWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_lNewIcon
        End If
        RaiseEvent DialogConstucted(hWnd)
    Case TDN_CREATED
        m_TDhWnd = hWnd
        dtAC1 = Now
        ReDim m_hButton(0)
        lButtonX = 0
        lButtonY = 0
        lButtonX2 = 0
        lButtonY2 = 0
        lButtonX3 = 0
        lButtonY3 = 0
        lButtonX4 = 0
        lButtonY4 = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(hWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(hWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        SetButtonIcons
        SetInitialButtonStates
        If (m_sExpandedInfo <> "") Then bPosUsed(1) = True 'expando control taking up the button position, so we can't add ours
        If (m_sVerify <> "") Then bPosUsed(1) = True 'same for verify
        If (uTDC.dwFlags And TDF_INPUT_BOX) Then AddInputBox: bSC = True
        If (uTDC.dwFlags And TDF_COMBO_BOX) Then AddComboBox nCbType
        If (uTDC.dwFlags And TDF_DATETIME) Then AddDateTime: bSC = True
        If (uTDC.dwFlags And TDF_SLIDER) Then AddSlider: bSC = True
        If (hLogoImage) Then AddLogo: bSC = True
        If m_HasCustomControl = True Then FocusMainControl
        If bSC Then 'combo box subclasses different hWnds in the Add func
            Subclass2 hUI, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)
        End If
        RaiseEvent DialogCreated(hWnd)
    Case TDN_HYPERLINK_CLICKED
        If uTDC.dwFlags And TDF_EXEC_HYPERLINKS Then
            Call ShellExecuteW(0, 0, lParam, 0, 0, 1)
        End If
        zz_ProcessCallback = S_OK
        RaiseEvent HyperlinkClick(lParam)
        Exit Function
    Case TDN_BUTTON_CLICKED
        If ButtonIsHeld(CLng(wParam)) Then
            Debug.Print "BtnClick(Held) " & wParam
            zz_ProcessCallback = S_FALSE
        Else
            Debug.Print "BtnClick " & wParam
            zz_ProcessCallback = S_OK
        End If
        RaiseEvent ButtonClick(CLng(wParam))
        Exit Function
    Case TDN_TIMER
        dtAC2 = Now
        If (lAutoclose > 0) And (DateDiff("s", dtAC1, dtAC2) > lAutoclose) Then
            RaiseEvent AutoClose
            CloseDialog
        Else
            lAutocloseTrack = lAutoclose - DateDiff("s", dtAC1, dtAC2)
        End If
        If bReposTrigger Then
            If nTmrCnt = 5 Then
                nTmrCnt = 0
                Dim t3 As RECT
                GetClientRect m_TDhWnd, t3
                If t3.Bottom <> lLastCY Then
                    bReposTrigger = False
                    InputReposition
                    ComboReposition
                    DateTimeReposition
                    SliderReposition
                End If
            Else
                nTmrCnt = nTmrCnt + 1
            End If
        End If
        
        If bPublicTimer Then
            RaiseEvent Timer(CLng(wParam))
            If bResetTimer Then
                bResetTimer = False
                zz_ProcessCallback = S_FALSE
            Else
                zz_ProcessCallback = S_OK
            End If
            Exit Function
        End If
    Case TDN_RADIO_BUTTON_CLICKED
        m_ResRad = CLng(wParam) 'for when we need result before dialog is closed
        RaiseEvent RadioButtonClick(CLng(wParam))
        Exit Function
    Case TDN_VERIFICATION_CLICKED
        m_ResVerify = CLng(wParam)
        RaiseEvent VerificationClicked(CLng(wParam))
        Exit Function
    Case TDN_HELP
        RaiseEvent Help
        Exit Function
    Case TDN_EXPANDO_BUTTON_CLICKED
        If (uTDC.dwFlags And TDF_INPUT_BOX) Or (uTDC.dwFlags And TDF_COMBO_BOX) Or (uTDC.dwFlags And TDF_DATETIME) Or (uTDC.dwFlags And TDF_SLIDER) Then
            If (uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                Dim t2 As RECT
                GetClientRect m_TDhWnd, t2
                lLastCY = t2.Bottom
                bReposTrigger = True
                If wParam Then
'                    DebugAppend "Expando wParam 1; nba=" & nBreaksAdded
                    If nBreaksAdded = 1 Then
'                        DebugAppend "Remove single"
                        Me.Content = Left$(m_sContent, Len(m_sContent) - 2)
                    ElseIf nBreaksAdded = 2 Then
'                        DebugAppend "Remove double"
                        Me.Content = Left$(m_sContent, Len(m_sContent) - 4)
                    End If
                Else
'                    DebugAppend "Expando wParam 0; nba=" & nBreaksAdded
                    If nBreaksAdded = 1 Then
                        Me.Content = m_sContent & vbCrLf
                    ElseIf nBreaksAdded = 2 Then
'                        DebugAppend "add double"
                        If (uTDC.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                            Me.Content = m_sContent & vbCrLf & vbCrLf & " "  '(link control requires a space or it ignores the double-break)
                        Else
                            Me.Content = m_sContent & vbCrLf & vbCrLf
                        End If

    
                    End If
                End If
            End If
        End If

        RaiseEvent ExpandButtonClicked(CLng(wParam))
        Exit Function
    Case TDN_NAVIGATED
        m_TDhWnd = hWnd 'new dialog treated as page means no TDN_CREATED
        ReDim m_hButton(0)
        dtAC1 = Now
        lButtonX = 0
        lButtonY = 0
        lButtonX2 = 0
        lButtonY2 = 0
        lButtonX3 = 0
        lButtonY3 = 0
        lButtonX4 = 0
        lButtonY4 = 0
        ReDim arBtnX(0)
        'If m_cbEnum = 0& Then m_cbEnum = AddressOf ProcessEnumCallback ' scb_SetCallbackAddr(2, 3)
        'Call EnumChildWindows(hWnd, m_cbEnum, ObjPtr(Me))
        Call EnumChildWindows(hWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
        SetButtonIcons
        SetInitialButtonStates
        If (m_sExpandedInfo <> "") Then bPosUsed(1) = True 'expando control taking up the button position, so we can't add ours
        If (m_sVerify <> "") Then bPosUsed(1) = True 'same for verify
        If (uTDC.dwFlags And TDF_INPUT_BOX) Then AddInputBox: bSC = True
        If (uTDC.dwFlags And TDF_COMBO_BOX) Then AddComboBox nCbType
        If (uTDC.dwFlags And TDF_DATETIME) Then AddDateTime: bSC = True
        If (uTDC.dwFlags And TDF_SLIDER) Then AddSlider: bSC = True
        If (hLogoImage) Then AddLogo
        If m_HasCustomControl = True Then FocusMainControl
        If bSC Then 'combo box subclasses different hWnds in the Add func
            Subclass2 hUI, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)
        End If
        RaiseEvent Navigated
        Exit Function
    Case TDN_DESTROYED
        If hEditBox Then
            Dim lLen As Long, sText As String
            lLen = CLng(SendMessageW(hEditBox, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2)
            If lLen Then
                sText = Space$(lLen)
                Call SendMessageW(hEditBox, WM_GETTEXT, lLen, ByVal sText)
                sEditText = StrConv(sText, vbFromUnicode)
            End If
        End If
        If hCombo Then
            nComboIdx = CLng(SendMessageW(hCombo, CB_GETCURSEL, 0, ByVal 0&))
            If nCbType = cbtDropdownList Then
                If nComboIdx >= 0 Then
                    sComboText = aComboItems(nComboIdx).sText
                Else
                	sComboText = ""
                End If
            Else
            	sComboText = GetComboTextW(hEditCombo)
            End If
        End If
        Dim rs As Long
        If hDateTime Then
            rs = CLng(SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur))
            If rs = GDT_VALID Then
                bDTCheck = True
            Else
                bDTCheck = False
            End If
        End If
        If hDateTime2 Then
            rs = CLng(SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2))
            If rs = GDT_VALID Then
                bDTCheck2 = True
            Else
                bDTCheck2 = False
            End If
        End If
        If hSlider Then
            lSliderValue = CLng(SendMessageW(hSlider, TBM_GETPOS, 0, ByVal 0&))
        End If
        RaiseEvent DialogDestroyed
        m_TDhWnd = 0
End Select
End Function

'@1
Public Function zz_ProcessSubclass(ByVal lng_hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr
On Error GoTo e0
'DebugAppend "TDSC uMsg=" & CStr(uMsg)
Select Case uMsg
    Case WM_CTLCOLORSTATIC
        If (lParam = hStatic) And (hStatic <> 0&) Then
            Dim hdc1 As LongPtr
            hdc1 = GetDC(hUI)
            Dim pxl As Long
            If nLogoPos = LogoTopRight Then
                pxl = GetPixel(hdc1, lgpX - 1, 0)
            ElseIf nLogoPos = LogoButtons Then
                pxl = GetPixel(hdc1, lgpX, lgpY - 1) '109)
            End If
'            DebugAppend "pxl(" & lgpX & "," & lgpY & ")=0x" & Hex$(pxl) & ",y=" & lButtonY

            zz_ProcessSubclass = CreateSolidBrush(pxl)
            Exit Function
            
'----------------------------
'FOR FUTURE VERSION; IGNORE
'  Dim pxBase As Long
'Dim pxTest As Long
'Dim yStart As Long
'Dim i As Long
'Dim tRC As RECT
'GetClientRect hUI, tRC
'yStart = tRC.Bottom
'pxBase = GetPixel(hdc1, 3, yStart)
'DebugAppend "pxBase=0x" & Hex$(pxBase)
'For i = yStart To (yStart - 32) Step -1
'    pxTest = GetPixel(hdc1, 3, yStart - i)
'    DebugAppend "pxTest=0x" & Hex$(pxTest)
'    If pxTest <> pxBase Then 'first color change is the divider line
'        'GetBottomDividerPosA = i + 4
''        DebugAppend "Get pos=" & i + 4
''        Exit For
'    End If
'Next i
'-----------------------------
            
            Exit Function
        End If
            
    Case BCM_SETDROPDOWNSTATE
        Debug.Print "BCM_SETDROPDOWNSTATE"
        If wParam = 1 Then
            RaiseEvent DropdownButtonClicked(lng_hWnd)
        End If
    Case WM_HSCROLL
        If lParam = hSlider Then RaiseEvent SliderChange(SliderValue)
    Case WM_COMMAND
'        DebugAppend "WM_COMMAND " & tdHiWord(wParam)
        Select Case tdHiWord(CLng(wParam))
        
            Case CBN_SELCHANGE
                RaiseEvent ComboItemChanged(ComboIndex)
            Case CBN_DROPDOWN
                RaiseEvent ComboDropdown
            Case EN_CHANGE
                RaiseEvent InputBoxChange(InputText)
        End Select
    Case WM_NOTIFY
            Dim NM As NMHDR
            CopyMemory NM, ByVal lParam, LenB(NM)
'            DebugAppend "WM_NOTIFY code=" & NM.Code
            Select Case NM.code
                Case DTN_DATETIMECHANGE
                    RaiseEvent DateTimeChange(DateTimeValue, DateTimeChecked)
            End Select

    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call UnSubclass2(lng_hWnd, PtrTDWndProc, uIdSubclass)
     
  End Select
  
  zz_ProcessSubclass = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)

  
Exit Function
e0:
DebugAppend "ProcessSubclass Error=" & Err.Description

End Function
Private Function PtrTDWndProc() As LongPtr
    PtrTDWndProc = tdFARPROC(AddressOf TaskDialogSubclassProc)
End Function


#If False Then

Private Sub IDEForceCase()
'dont be changin my enums damnit. this section is not compiled.
    Const ePBST_NORMAL = 1
    Const ePBST_ERROR = 2
    Const ePBST_PAUSED = 3
    Const TD_WARNING_ICON = -1 'exclamation point in a yellow 'yield' triangle (same image as IDI_EXCLAMATION)
    Const TD_ERROR_ICON = -2 'round red circle containg 'X' (same as IDI_HAND)
    Const TD_INFORMATION_ICON = -3 'round blue circle containing 'i' (same image as IDI_ASTERISK)
    Const TD_SHIELD_ICON = -4 'Vista's security shield
    Const IDI_APPLICATION = 32512& 'miniature picture of an application window
    Const IDI_QUESTION = 32514& 'round blue circle containing '?'

    Const TD_SHIELD_GRADIENT_ICON = -5 'same image as TD_SHIELD_ICON; main message text on gradient blue background
    Const TD_SHIELD_WARNING_ICON = -6 'exclamation point in yellow Shield shape; main message text on gradient orange background
    Const TD_SHIELD_ERROR_ICON = -7 'X contained within Shield shape; main message text on gradient red background
    Const TD_SHIELD_OK_ICON = -8 'Shield shape containing green checkmark; main message text on gradient green background
    Const TD_SHIELD_GRAY_ICON = -9 'same image as TD_SHIELD_ICON; main message text on medium gray background
    Const TD_NO_ICON = 0 'no icon; text on white background
    Const TDCBF_OK_BUTTON = &H1&      'return value 1 (IDOK)
    Const TDCBF_YES_BUTTON = &H2&     'return value 6 (IDYES)
    Const TDCBF_NO_BUTTON = &H4&      'return value 7 (IDNO)
    Const TDCBF_CANCEL_BUTTON = &H8&  'return value 2 (IDCANCEL)
    Const TDCBF_RETRY_BUTTON = &H10&   'return value 4 (IDRETRY)
    Const TDCBF_CLOSE_BUTTON = &H20&   'return value 8 (IDCLOSE)
    Const TD_OK = 1
    Const TD_YES = 6
    Const TD_NO = 7
    Const TD_CANCEL = 2
    Const TD_RETRY = 4
    Const TD_CLOSE = 8
    Const TDF_ENABLE_HYPERLINKS = &H1
    Const TDF_USE_HICON_MAIN = &H2
    Const TDF_USE_HICON_FOOTER = &H4
    Const TDF_ALLOW_DIALOG_CANCELLATION = &H8
    Const TDF_USE_COMMAND_LINKS = &H10
    Const TDF_USE_COMMAND_LINKS_NO_ICON = &H20
    Const TDF_EXPAND_FOOTER_AREA = &H40
    Const TDF_EXPANDED_BY_DEFAULT = &H80
    Const TDF_VERIFICATION_FLAG_CHECKED = &H100
    Const TDF_SHOW_PROGRESS_BAR = &H200
    Const TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400
    Const TDF_CALLBACK_TIMER = &H800
    Const TDF_POSITION_RELATIVE_TO_WINDOW = &H1000
    Const TDF_RTL_LAYOUT = &H2000
    Const TDF_NO_DEFAULT_RADIO_BUTTON = &H4000
    Const TDF_CAN_BE_MINIMIZED = &H8000&
    Const TDF_USE_SHELL32_ICONID = &H10000 'CUSTOM FLAG
    Const TDF_USE_IMAGERES_ICONID = &H20000 'CUSTOM FLAG
    Const TDF_USE_SHELL32_ICONID_BUTTON = &H80000 'CUSTOM FLAG
    Const TDF_USE_IMAGERES_ICONID_BUTTON = &H100000 'CUSTOM FLAG
    Const TDF_EXEC_HYPERLINKS = &H40000 'CUSTOM FLAG
    Const TDF_KILL_SHIELD_ICON = &H200000 'CUSTOM FLAG: Allows the colored background from the TD_SHIELD_x icons to be used with just the text and no shield icon
    Const TDF_INPUT_BOX = &H400000        'CUSTOM FLAG: Input box. See below.
    Const TDF_COMBO_BOX = &H800000       'CUSTOM FLAG: Combo box; .ComboType controls edit or list
    Const TDF_DATETIME = &H1000000        'CUSTOM FLAG: Add calendar control
    Const TDF_SLIDER = &H2000000
    Const TDM_NAVIGATE_PAGE = WM_USER + 101&
    Const TDM_CLICK_BUTTON = WM_USER + 102&                  '// wParam = Button ID
    Const TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103&      '// wParam = 0 (nonMarque) wParam != 0 (Marquee)
    Const TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104&        '// wParam = new progress state
    Const TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105&        '// lParam = tdMAKELPARAM(nMinRange& nMaxRange)
    Const TDM_SET_PROGRESS_BAR_POS = WM_USER + 106&          '// wParam = new position
    Const TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107&      '// wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds between repaints)
    Const TDM_SET_ELEMENT_TEXT = WM_USER + 108&              '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    Const TDM_CLICK_RADIO_BUTTON = WM_USER + 110&            '// wParam = Radio Button ID
    Const TDM_ENABLE_BUTTON = WM_USER + 111&                 '// lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
    Const TDM_ENABLE_RADIO_BUTTON = WM_USER + 112&           '// lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
    Const TDM_CLICK_VERIFICATION = WM_USER + 113&            '// wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
    Const TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114&           '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    Const TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115& '// wParam = Button ID, lParam = 0 (elevation not required)& lParam != 0 (elevation required)
    Const TDM_UPDATE_ICON = WM_USER + 116&                   '// wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
    Const TDN_CREATED = 0
    Const TDN_NAVIGATED = 1
    Const TDN_BUTTON_CLICKED = 2           '// wParam = Button ID
    Const TDN_HYPERLINK_CLICKED = 3        '// lParam = (LPCWSTR)pszHREF
    Const TDN_TIMER = 4                    '// wParam = Milliseconds since dialog created or timer reset
    Const TDN_DESTROYED = 5
    Const TDN_RADIO_BUTTON_CLICKED = 6     '// wParam = Radio Button ID
    Const TDN_DIALOG_CONSTRUCTED = 7
    Const TDN_VERIFICATION_CLICKED = 8     '// wParam = 1 if checkbox checked, 0 if not, lParam is unused and always 0
    Const TDN_HELP = 9
    Const TDN_EXPANDO_BUTTON_CLICKED = 10  '// wParam = 0 (dialog is now collapsed), wParam != 0 (dialog is now expanded)
    Const TDE_CONTENT = 0
    Const TDE_EXPANDED_INFORMATION = 1
    Const TDE_FOOTER = 2
    Const TDE_MAIN_INSTRUCTION = 3
    Const TDIE_ICON_MAIN = 0
    Const TDIE_ICON_FOOTER = 1
    
End Sub

#End If

#If False Then
Dim WS_OVERLAPPED, WS_TABSTOP, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_GROUP, WS_THICKFRAME, _
WS_SYSMENU, WS_HSCROLL, WS_VSCROLL, WS_DLGFRAME, WS_BORDER, WS_CAPTION, WS_MAXIMIZE, _
WS_CLIPCHILDREN, WS_CLIPSIBLINGS, WS_DISABLED, WS_VISIBLE, WS_MINIMIZE, WS_CHILD, _
WS_POPUP, WS_TILED, WS_ICONIC, WS_SIZEBOX, WS_OVERLAPPEDWINDOW, WS_TILEDWINDOW, _
WS_POPUPWINDOW, WS_CHILDWINDOW
#End If

#If False Then
Dim TDIBA_Content, TDIBA_Buttons, TDIBA_Footer
#End If
#If False Then
Dim ctdInputText, ctdInputCombo, ctdInputList
#End If
#If False Then
Dim GWL_WNDPROC, GWL_HINSTANCE, GWL_HWNDPARENT, GWL_ID, GWL_STYLE, GWL_EXSTYLE, _
GWL_USERDATA
#End If
#If False Then
Dim TDIBTA_Left, TDIBTA_Right, TDIBTA_Center
#End If
#If False Then
Dim DTS_SHORTDATEFORMAT, DTS_UPDOWN, DTS_SHOWNONE, DTS_LONGDATEFORMAT, DTS_TIMEFORMAT, _
DTS_APPCANPARSE, DTS_RIGHTALIGN, DTS_SHORTDATECENTURYFORMAT
#End If

#If False Then
Dim cbtDropdownEdit, cbtDropdownList
#End If
#If False Then
Dim dttDate, dttDateWithCheck, dttTime, dttTimeWithCheck, dttDateTime, dttDateTimeWithCheck, _
dttDateTimeWithCheckTimeOnly, dttDateTimeWithCheckDateOnly
#End If

#If False Then
Dim SldTickStyleBottom, SldTickStyleTop, SldTickStyleBoth, SldTickStyleNone
#End If
#If False Then
Dim tdccInputBox, tdccComboBox, tdccDateTime, tdccSlider, tdccNothing
#End If
#If False Then
Dim LogoBitmap, LogoIcon
#End If
#If False Then
Dim LogoTopRight, LogoButtons
#End If
#If False Then
Dim tdcaLeft, tdcaCenter, tdcaRight
#End If
#If False Then
Dim SW_HIDE, SW_SHOWNORMAL, SW_NORMAL, SW_SHOWMINIMIZED, SW_SHOWMAXIMIZED, SW_MAXIMIZE, _
SW_SHOWNOACTIVATE, SW_SHOW, SW_MINIMIZE, SW_SHOWMINNOACTIVE, SW_SHOWNA, SW_RESTORE, _
SW_SHOWDEFAULT
#End If

