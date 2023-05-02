Attribute VB_Name = "mTDSample"
#If Win64 Then
Option Explicit
'mTDSample.bas
'Module for cTaskDialog Demo
'This module is only required for some actions performed by the demos
'It is not required to use cTaskDialog.cls, which now requires no external modules.



'Icon code was mostly written by Leandro Ascierto, from his clsMenuImage.
'I've simply modified the resource->hicon function to stand alone
Public Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As LongPtr) As Long

Private Const MAX_PATH = 260

Private Type IconHeader
    ihReserved      As Integer
    ihType          As Integer
    ihCount         As Integer
End Type

Private Type IconEntry
    ieWidth         As Byte
    ieHeight        As Byte
    ieColorCount    As Byte
    ieReserved      As Byte
    iePlanes        As Integer
    ieBitCount      As Integer
    ieBytesInRes    As Long
    ieImageOffset   As Long
End Type
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As LongPtr

Private Declare PtrSafe Function CreateIconFromResource Lib "user32.dll" (ByVal presbits As LongPtr, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As LongPtr
Private Declare PtrSafe Function LookupIconIdFromDirectoryEx Lib "user32.dll" (ByVal presbits As LongPtr, ByVal fIcon As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare PtrSafe Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As LongPtr
Public Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
Public Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As LongPtr, GpImage As LongPtr) As Long
Public Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As LongPtr, Width As Long) As Long
Public Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As LongPtr, Height As Long) As Long
Public Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
Public Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As Long
Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Public Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef lpInput As GdiplusStartupInput, ByRef lpOutput As Long) As Long
Public Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long

Public gdipInitToken As LongPtr
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long


 Public Function InitGDIPlus() As LongPtr
    Dim Token    As LongPtr
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As LongPtr)
    GdiplusShutdown Token
End Sub
 Public Function hBitmapFromFile(PicFile As String, Width As Long, Height As Long, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As LongPtr
    Dim hDC     As LongPtr
    Dim hBitmap As LongPtr
    Dim Img     As LongPtr
    
    If gdipInitToken = 0 Then
        gdipInitToken = InitGDIPlus()
    End If
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
'        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    Debug.Print "gdip himage=" & Img
    GdipCreateHBITMAPFromBitmap Img, hBitmap, &H0
    ' Calculate picture's width and height if not specified
'    If Width = -1 Or Height = -1 Then
'        GdipGetImageWidth Img, Width
'        GdipGetImageHeight Img, Height
'    End If
'
'    ' Initialise the hDC
'    InitDC hDC, hBitmap, BackColor, Width, Height
'
'    ' Resize the picture
'    'gdipResize Img, hDC, Width, Height, RetainRatio
'    gdipDrawCentered Img, hDC, Width, Height, True
    GdipDisposeImage Img
'
'    ' Get the bitmap back
'    GetBitmap hDC, hBitmap
    
    hBitmapFromFile = hBitmap
End Function




Public Function ResIconToHICON(id As String, Optional CX As Long = 24, Optional CY As Long = 24) As LongPtr
'returns an hIcon from an icon in the resource file
'Icons must be added as a custom resource

    Dim tIconHeader     As IconHeader
    Dim tIconEntry()    As IconEntry
    Dim MaxBitCount     As Long
    Dim MaxSize         As Long
    Dim Aproximate      As Long
    Dim IconID          As Long
    Dim hIcon           As LongPtr
    Dim i               As Long
    Dim bytIcoData() As Byte
    
On Error GoTo e0

    bytIcoData = LoadResData(id, "CUSTOM")
    Call CopyMemory(tIconHeader, bytIcoData(0), Len(tIconHeader))

    If tIconHeader.ihCount >= 1 Then
    
        ReDim tIconEntry(tIconHeader.ihCount - 1)
        
        Call CopyMemory(tIconEntry(0), bytIcoData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
        
        IconID = -1
           
        For i = 0 To tIconHeader.ihCount - 1
            If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
        Next

       
        For i = 0 To tIconHeader.ihCount - 1
            If MaxBitCount = tIconEntry(i).ieBitCount Then
                MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                If MaxSize > Aproximate And MaxSize <= (CX + CY) Then
                    Aproximate = MaxSize
                    IconID = i
                End If
            End If
        Next
                   
        If IconID = -1 Then Exit Function
       
        With tIconEntry(IconID)
            hIcon = CreateIconFromResourceEx(bytIcoData(.ieImageOffset), .ieBytesInRes, 1, &H30000, CX, CY, &H0)
            If hIcon <> 0 Then
                ResIconToHICON = hIcon
            End If
        End With
       
    End If
'Debug.Print "Res hIcon=" & hIcon

On Error GoTo 0
Exit Function

e0:
Debug.Print "modIcon.ResIconTohIcon.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Public Function IconToHICON(IcoData() As Byte, DesiredX As Long, DesiredY As Long) As LongPtr
    Dim lPtrSrc As Long, lPtrDst As Long, lID As Long
    Dim icDir() As Byte, LB As Long
    Dim tIconHeader As IconHeader
    Dim tIconEntry As IconEntry
    Dim ICRESVER As Long
    ICRESVER = &H30000
    LB = LBound(IcoData) ' just in case a non-zero LBound array passed
    ' convert 16 byte IconDir to 14 byte IconDir
    CopyMemory tIconHeader, IcoData(LB), Len(tIconHeader)
    ReDim icDir(0 To tIconHeader.ihCount * Len(tIconEntry) + Len(tIconHeader) - 1&)
    CopyMemory icDir(0), tIconHeader, Len(tIconHeader)
    lPtrDst = Len(tIconHeader)
    lPtrSrc = LB + lPtrDst
    For lID = 1& To tIconHeader.ihCount
        CopyMemory tIconEntry, IcoData(lPtrSrc), 12& ' size of standard tIconEntry less last 4 bytes
        tIconEntry.ieImageOffset = lID
        CopyMemory icDir(lPtrDst), tIconEntry, 14&     ' size of DLL tIconEntry
        lPtrDst = lPtrDst + 14&: lPtrSrc = lPtrSrc + Len(tIconEntry)
    Next
    lID = LookupIconIdFromDirectoryEx(VarPtr(icDir(0)), True, DesiredX, DesiredY, 0&)
    Erase icDir()
    If lID > 0& Then
        CopyMemory tIconEntry, IcoData(LB + (lID - 1&) * Len(tIconEntry) + Len(tIconHeader)), Len(tIconEntry)
        
        IconToHICON = CreateIconFromResource(VarPtr(IcoData(LB + tIconEntry.ieImageOffset)), tIconEntry.ieBytesInRes, True, ICRESVER)
    End If
End Function
Public Function LoadIcoFile(sFile As String) As Byte()
    Dim f As Long
    'Dim b() As Byte
     
    f = FreeFile()
    Open sFile For Binary As f
    ReDim LoadIcoFile(LOF(f))
    Get f, , LoadIcoFile
    Close f
End Function
Public Function GetSystemImagelist(uSize As Long) As LongPtr
  Dim sfi As SHFILEINFO
  Dim wd As String
  wd = Environ("WINDIR")
  wd = Left(wd, 3)
  ' Any valid file system path can be used to retrieve system image list handles.
  GetSystemImagelist = SHGetFileInfo(wd, 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uSize)
End Function

#If False Then
Dim SHGFI_LARGEICON, SHGFI_SMALLICON, SHGFI_OPENICON, SHGFI_SHELLICONSIZE, SHGFI_PIDL, _
SHGFI_USEFILEATTRIBUTES, SHGFI_ICON, SHGFI_DISPLAYNAME, SHGFI_TYPENAME, SHGFI_ATTRIBUTES, _
SHGFI_ICONLOCATION, SHGFI_EXETYPE, SHGFI_SYSICONINDEX, SHGFI_LINKOVERLAY, SHGFI_SELECTED, _
SHGFI_ATTR_SPECIFIED
#End If

#Else

Option Explicit
'mTDSample.bas
'Module for cTaskDialog Demo
'This module is only required for some actions performed by the demos
'It is not required to use cTaskDialog.cls, which now requires no external modules.



'Icon code was mostly written by Leandro Ascierto, from his clsMenuImage.
'I've simply modified the resource->hicon function to stand alone
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Const MAX_PATH = 260

Private Type IconHeader
    ihReserved      As Integer
    ihType          As Integer
    ihCount         As Integer
End Type

Private Type IconEntry
    ieWidth         As Byte
    ieHeight        As Byte
    ieColorCount    As Byte
    ieReserved      As Byte
    iePlanes        As Integer
    ieBitCount      As Integer
    ieBytesInRes    As Long
    ieImageOffset   As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, _
                                                                    ByVal dwResSize As Long, _
                                                                    ByVal fIcon As Long, _
                                                                    ByVal dwVer As Long, _
                                                                    ByVal cxDesired As Long, _
                                                                    ByVal cyDesired As Long, _
                                                                    ByVal Flags As Long) As Long
Private Declare Function CreateIconFromResource Lib "user32.dll" (ByVal presbits As Long, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Private Declare Function LookupIconIdFromDirectoryEx Lib "user32.dll" (ByVal presbits As Long, ByVal fIcon As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As Long
Public Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
Public Declare Function GdipLoadImageFromFile Lib "GDIplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Public Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal Image As Long, Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal Image As Long, Height As Long) As Long
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIplus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Public Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Public Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type
Public Declare Function GdiplusStartup Lib "GDIplus" (ByRef Token As Long, ByRef lpInput As GdiplusStartupInput, Optional ByRef lpOutput As Any) As Long
Public Declare Function GdiplusShutdown Lib "GDIplus" (ByVal Token As Long) As Long
Public gdipInitToken As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

 Public Function InitGDIPlus() As Long
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub
 Public Function hBitmapFromFile(PicFile As String, Width As Long, Height As Long, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As Long
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    
    If gdipInitToken = 0 Then
        gdipInitToken = InitGDIPlus()
    End If
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
'        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    Debug.Print "gdip himage=" & Img
    GdipCreateHBITMAPFromBitmap Img, hBitmap, &H0
    ' Calculate picture's width and height if not specified
'    If Width = -1 Or Height = -1 Then
'        GdipGetImageWidth Img, Width
'        GdipGetImageHeight Img, Height
'    End If
'
'    ' Initialise the hDC
'    InitDC hDC, hBitmap, BackColor, Width, Height
'
'    ' Resize the picture
'    'gdipResize Img, hDC, Width, Height, RetainRatio
'    gdipDrawCentered Img, hDC, Width, Height, True
    GdipDisposeImage Img
'
'    ' Get the bitmap back
'    GetBitmap hDC, hBitmap
    
    hBitmapFromFile = hBitmap
End Function




Public Function ResIconToHICON(id As String, Optional CX As Long = 24, Optional CY As Long = 24) As Long
'returns an hIcon from an icon in the resource file
'Icons must be added as a custom resource

    Dim tIconHeader     As IconHeader
    Dim tIconEntry()    As IconEntry
    Dim MaxBitCount     As Long
    Dim MaxSize         As Long
    Dim Aproximate      As Long
    Dim IconID          As Long
    Dim hIcon           As Long
    Dim i               As Long
    Dim bytIcoData() As Byte
    
On Error GoTo e0

    bytIcoData = LoadResData(id, "CUSTOM")
    Call CopyMemory(tIconHeader, bytIcoData(0), Len(tIconHeader))

    If tIconHeader.ihCount >= 1 Then
    
        ReDim tIconEntry(tIconHeader.ihCount - 1)
        
        Call CopyMemory(tIconEntry(0), bytIcoData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
        
        IconID = -1
           
        For i = 0 To tIconHeader.ihCount - 1
            If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
        Next

       
        For i = 0 To tIconHeader.ihCount - 1
            If MaxBitCount = tIconEntry(i).ieBitCount Then
                MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                If MaxSize > Aproximate And MaxSize <= (CX + CY) Then
                    Aproximate = MaxSize
                    IconID = i
                End If
            End If
        Next
                   
        If IconID = -1 Then Exit Function
       
        With tIconEntry(IconID)
            hIcon = CreateIconFromResourceEx(bytIcoData(.ieImageOffset), .ieBytesInRes, 1, &H30000, CX, CY, &H0)
            If hIcon <> 0 Then
                ResIconToHICON = hIcon
            End If
        End With
       
    End If
'Debug.Print "Res hIcon=" & hIcon

On Error GoTo 0
Exit Function

e0:
Debug.Print "modIcon.ResIconTohIcon.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Public Function IconToHICON(IcoData() As Byte, DesiredX As Long, DesiredY As Long) As Long
    Dim lPtrSrc As Long, lPtrDst As Long, lID As Long
    Dim icDir() As Byte, LB As Long
    Dim tIconHeader As IconHeader
    Dim tIconEntry As IconEntry
    Dim ICRESVER As Long
    ICRESVER = &H30000
    LB = LBound(IcoData) ' just in case a non-zero LBound array passed
    ' convert 16 byte IconDir to 14 byte IconDir
    CopyMemory tIconHeader, IcoData(LB), Len(tIconHeader)
    ReDim icDir(0 To tIconHeader.ihCount * Len(tIconEntry) + Len(tIconHeader) - 1&)
    CopyMemory icDir(0), tIconHeader, Len(tIconHeader)
    lPtrDst = Len(tIconHeader)
    lPtrSrc = LB + lPtrDst
    For lID = 1& To tIconHeader.ihCount
        CopyMemory tIconEntry, IcoData(lPtrSrc), 12& ' size of standard tIconEntry less last 4 bytes
        tIconEntry.ieImageOffset = lID
        CopyMemory icDir(lPtrDst), tIconEntry, 14&     ' size of DLL tIconEntry
        lPtrDst = lPtrDst + 14&: lPtrSrc = lPtrSrc + Len(tIconEntry)
    Next
    lID = LookupIconIdFromDirectoryEx(VarPtr(icDir(0)), True, DesiredX, DesiredY, 0&)
    Erase icDir()
    If lID > 0& Then
        CopyMemory tIconEntry, IcoData(LB + (lID - 1&) * Len(tIconEntry) + Len(tIconHeader)), Len(tIconEntry)
        
        IconToHICON = CreateIconFromResource(VarPtr(IcoData(LB + tIconEntry.ieImageOffset)), tIconEntry.ieBytesInRes, True, ICRESVER)
    End If
End Function
Public Function LoadIcoFile(sFile As String) As Byte()
    Dim f As Long
    'Dim b() As Byte
     
    f = FreeFile()
    Open sFile For Binary As f
    ReDim LoadIcoFile(LOF(f))
    Get f, , LoadIcoFile
    Close f
End Function
Public Function GetSystemImagelist(uSize As Long) As Long
  Dim sfi As SHFILEINFO
  Dim wd As String
  wd = Environ("WINDIR")
  wd = Left(wd, 3)
  ' Any valid file system path can be used to retrieve system image list handles.
  GetSystemImagelist = SHGetFileInfo(wd, 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uSize)
End Function

#If False Then
Dim SHGFI_LARGEICON, SHGFI_SMALLICON, SHGFI_OPENICON, SHGFI_SHELLICONSIZE, SHGFI_PIDL, _
SHGFI_USEFILEATTRIBUTES, SHGFI_ICON, SHGFI_DISPLAYNAME, SHGFI_TYPENAME, SHGFI_ATTRIBUTES, _
SHGFI_ICONLOCATION, SHGFI_EXETYPE, SHGFI_SYSICONINDEX, SHGFI_LINKOVERLAY, SHGFI_SELECTED, _
SHGFI_ATTR_SPECIFIED
#End If

#End If
