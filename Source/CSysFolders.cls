VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystemFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2004-8 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Enhanced API provided by shfolder (requires IE5 or VPC Virtual Machine Additions)
' Vista users see: http://msdn.microsoft.com/en-us/library/bb762181(VS.85).aspx
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
' Converts an item identifier list to a file system path.
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
   
' Used to determine if an API function is exported.
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
   
' The old reliables.
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const S_OK = 0
Private Const S_FALSE = 1
Private Const E_INVALIDARG = &H80070057
Private Const MAX_PATH = 260

' Enumerations
Private Enum SHGFP_TYPES
   SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
   SHGFP_TYPE_DEFAULT = &H1
End Enum

Public Enum CSIDL_VALUES
   CSIDL_DESKTOP = &H0
   CSIDL_INTERNET = &H1
   CSIDL_PROGRAMS = &H2
   CSIDL_CONTROLS = &H3
   CSIDL_PRINTERS = &H4
   CSIDL_PERSONAL = &H5
   CSIDL_FAVORITES = &H6
   CSIDL_STARTUP = &H7
   CSIDL_RECENT = &H8
   CSIDL_SENDTO = &H9
   CSIDL_BITBUCKET = &HA
   CSIDL_STARTMENU = &HB
   CSIDL_MYDOCUMENTS = &HC
   CSIDL_MYMUSIC = &HD
   CSIDL_MYVIDEO = &HE
   CSIDL_DESKTOPDIRECTORY = &H10
   CSIDL_DRIVES = &H11
   CSIDL_NETWORK = &H12
   CSIDL_NETHOOD = &H13
   CSIDL_FONTS = &H14
   CSIDL_TEMPLATES = &H15
   CSIDL_COMMON_STARTMENU = &H16
   CSIDL_COMMON_PROGRAMS = &H17
   CSIDL_COMMON_STARTUP = &H18
   CSIDL_COMMON_DESKTOPDIRECTORY = &H19
   CSIDL_APPDATA = &H1A
   CSIDL_PRINTHOOD = &H1B
   CSIDL_LOCAL_APPDATA = &H1C
   CSIDL_ALTSTARTUP = &H1D
   CSIDL_COMMON_ALTSTARTUP = &H1E
   CSIDL_COMMON_FAVORITES = &H1F
   CSIDL_INTERNET_CACHE = &H20
   CSIDL_COOKIES = &H21
   CSIDL_HISTORY = &H22
   CSIDL_COMMON_APPDATA = &H23
   CSIDL_WINDOWS = &H24
   CSIDL_SYSTEM = &H25
   CSIDL_PROGRAM_FILES = &H26
   CSIDL_MYPICTURES = &H27
   CSIDL_PROFILE = &H28
   CSIDL_SYSTEMX86 = &H29
   CSIDL_PROGRAM_FILESX86 = &H2A
   CSIDL_PROGRAM_FILES_COMMON = &H2B
   CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
   CSIDL_COMMON_TEMPLATES = &H2D
   CSIDL_COMMON_DOCUMENTS = &H2E
   CSIDL_COMMON_ADMINTOOLS = &H2F
   CSIDL_ADMINTOOLS = &H30
   CSIDL_CONNECTIONS = &H31
   CSIDL_COMMON_MUSIC = &H35
   CSIDL_COMMON_PICTURES = &H36
   CSIDL_COMMON_VIDEO = &H37
   CSIDL_RESOURCES = &H38
   CSIDL_RESOURCES_LOCALIZED = &H39
   CSIDL_COMMON_OEM_LINKS = &H3A
   CSIDL_CDBURN_AREA = &H3B
   CSIDL_COMPUTERSNEARME = &H3D
   CSIDL_FLAG_PER_USER_INIT = &H800
   CSIDL_FLAG_NO_ALIAS = &H1000
   CSIDL_FLAG_DONT_VERIFY = &H4000
   CSIDL_FLAG_CREATE = &H8000
   CSIDL_FLAG_MASK = &HFF00
End Enum

' Member variables.
Private m_Enhanced As Boolean

' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
   ' Make sure we have the function this class relies on!
   m_Enhanced = Exported("shfolder", "SHGetFolderPathA")
End Sub

' *************************************************
'  Public Properties (read-only)
' *************************************************
Public Property Get Enhanced() As Boolean
   ' Allow client to determine if shell is updated for full support.
   Enhanced = m_Enhanced
End Property

Public Property Get Exists(ByVal CSIDL As CSIDL_VALUES) As Boolean
   Dim Buffer As String
   ' Check for requested system folder. This may return false even
   ' after asking the Path method to create the folder, if the OS
   ' doesn't support this specific constant.  Developer needs to
   ' account for this possibility by checking Exists after trying
   ' to create with Path!
   Buffer = Space$(MAX_PATH)
   If m_Enhanced Then
      Exists = (SHGetFolderPath(0&, CSIDL, 0&, SHGFP_TYPE_CURRENT, Buffer) = S_OK)
   Else
      Exists = IsDirectory(Me.Path(CSIDL))
   End If
End Property

Public Property Get Path(ByVal CSIDL As CSIDL_VALUES, Optional Default As Boolean = True, Optional ForceCreation As Boolean = False) As String
   Dim Buffer As String
   Dim pidl As Long
   Dim nRtn As Long
   
   ' Create buffer for return results.
   Buffer = Space$(MAX_PATH)
   
   ' Branch based on what tools we have...
   If m_Enhanced Then
      ' Tell system to create, if necessary
      If ForceCreation Then
         CSIDL = CSIDL Or CSIDL_FLAG_CREATE
      End If
      
      ' Grab requested system folder.
      nRtn = SHGetFolderPath(0&, CSIDL, 0&, Abs(Default), Buffer)
      Select Case nRtn
         Case S_OK
            ' Return requested system folder.
            Path = TrimNull(Buffer)
         Case S_FALSE
            ' Folder either doesn't exist or isn't supported.
            Debug.Print "The CSIDL in nFolder is valid, but the folder does not exist."
         Case E_INVALIDARG
            ' OS doesn't support this folder. Unfortunately, this
            ' result isn't returned unless CSIDL_FLAG_CREATE is
            ' specified in SHGetFolderPath call.
            Debug.Print "The CSIDL in nFolder is not valid."
      End Select
   
   Else
      ' Use old-fashioned method to get system folder.
      nRtn = SHGetSpecialFolderLocation(0&, CSIDL, pidl)
      Select Case nRtn
         Case S_OK
            ' Retrieve actual path, given the pidl.
            If SHGetPathFromIDList(pidl, Buffer) Then
               Path = TrimNull(Buffer)
            End If
            Call CoTaskMemFree(pidl)
         Case E_INVALIDARG
            ' OS doesn't support this folder.
            Debug.Print "The CSIDL in nFolder is not valid."
         Case Else
            ' Some sort of OLE-defined error code
            Debug.Print "OLE error: &h" & Hex$(nRtn)
      End Select
   End If
End Property

Public Property Get SysDir() As String
   Dim Buffer As String
   Dim nRet As Long
   ' Very low tech, but effective.
   Buffer = String$(MAX_PATH, 0)
   nRet = GetSystemDirectory(Buffer, Len(Buffer))
   If nRet Then SysDir = Left$(Buffer, nRet)
End Property

Public Property Get TmpDir() As String
   Dim Buffer As String
   Dim nRet As Long
   ' Very low tech, but effective.
   Buffer = String$(MAX_PATH, 0)
   nRet = GetTempPath(Len(Buffer), Buffer)
   If nRet Then TmpDir = Left$(Buffer, nRet)
End Property

Public Property Get WinDir() As String
   Dim Buffer As String
   Dim nRet As Long
   ' Very low tech, but effective.
   Buffer = String$(MAX_PATH, 0)
   nRet = GetWindowsDirectory(Buffer, Len(Buffer))
   If nRet Then WinDir = Left$(Buffer, nRet)
End Property

Public Property Get Description(ByVal CSIDL As CSIDL_VALUES) As String
   Dim sRtn As String
   ' See also:
   ' http://msdn.microsoft.com/library/en-us/shellcc/platform/shell/reference/enums/csidl.asp
   Const au As String = " (all users)"
   Const cu As String = " (current user)"
   Const sy As String = " (system)"
   Const vi As String = " (virtual)"
   
   Select Case CSIDL
      ' System folders
      Case CSIDL_PROGRAM_FILES:           sRtn = "C:\Program Files" & sy
      Case CSIDL_PROGRAM_FILES_COMMON:    sRtn = "C:\Program Files\Common" & sy
      Case CSIDL_WINDOWS:                 sRtn = "C:\Windows" & sy
      Case CSIDL_FONTS:                   sRtn = "C:\Windows\Fonts" & sy
      Case CSIDL_SYSTEM:                  sRtn = "C:\Windows\System" & sy
      Case CSIDL_RESOURCES:               sRtn = "C:\Windows\Resources" & sy
      Case CSIDL_RESOURCES_LOCALIZED:     sRtn = "C:\Windows\Resources\{LCID}" & sy
      Case CSIDL_SYSTEMX86:               sRtn = "x86 system directory on RISC"
      Case CSIDL_PROGRAM_FILESX86:        sRtn = "x86 Program Files folder on RISC"
      Case CSIDL_PROGRAM_FILES_COMMONX86: sRtn = "x86 Program Files Common folder on RISC"
      
      ' All user folders
      Case CSIDL_COMMON_APPDATA:          sRtn = "\Application Data" & au
      Case CSIDL_COMMON_DESKTOPDIRECTORY: sRtn = "\Desktop" & au
      Case CSIDL_COMMON_DOCUMENTS:        sRtn = "\Documents" & au
      Case CSIDL_COMMON_MUSIC:            sRtn = "\Documents\My Music" & au
      Case CSIDL_COMMON_PICTURES:         sRtn = "\Documents\My Pictures" & au
      Case CSIDL_COMMON_VIDEO:            sRtn = "\Documents\My Video" & au
      Case CSIDL_COMMON_FAVORITES:        sRtn = "\Favorites" & au
      Case CSIDL_COMMON_STARTMENU:        sRtn = "\Start Menu" & au
      Case CSIDL_COMMON_PROGRAMS:         sRtn = "\Start Menu\Programs" & au
      Case CSIDL_COMMON_ADMINTOOLS:       sRtn = "\Start Menu\Programs\Administrative Tools" & au
      Case CSIDL_COMMON_STARTUP:          sRtn = "\Start Menu\Programs\Startup" & au
      Case CSIDL_COMMON_TEMPLATES:        sRtn = "\Templates" & au
      ' Need a better definition ----------------------------------
      Case CSIDL_COMMON_OEM_LINKS:        sRtn = "Links to OEM specific apps" & au
      Case CSIDL_COMMON_ALTSTARTUP:       sRtn = "Nonlocalized Startup group" & au
      
      ' Current user folders
      Case CSIDL_PROFILE:          sRtn = "\{username}" & cu
      Case CSIDL_APPDATA:          sRtn = "\Application Data" & cu
      Case CSIDL_COOKIES:          sRtn = "\Cookies" & cu
      Case CSIDL_DESKTOPDIRECTORY: sRtn = "\Desktop" & cu
      Case CSIDL_FAVORITES:        sRtn = "\Favorites" & cu
      Case CSIDL_LOCAL_APPDATA:    sRtn = "\Local Settings\Application Data" & cu
      Case CSIDL_CDBURN_AREA:      sRtn = "\Local Settings\Application Data\Microsoft\CD Burning" & cu
      Case CSIDL_HISTORY:          sRtn = "\Local Settings\History" & cu
      Case CSIDL_PERSONAL:         sRtn = "\My Documents" & cu
      Case CSIDL_MYMUSIC:          sRtn = "\My Documents\My Music" & cu
      Case CSIDL_MYPICTURES:       sRtn = "\My Documents\My Pictures" & cu
      Case CSIDL_MYVIDEO:          sRtn = "\My Documents\My Video" & cu
      Case CSIDL_NETHOOD:          sRtn = "\NetHood" & cu
      Case CSIDL_PRINTHOOD:        sRtn = "\PrintHood" & cu
      Case CSIDL_RECENT:           sRtn = "\Recent" & cu
      Case CSIDL_SENDTO:           sRtn = "\SendTo" & cu
      Case CSIDL_STARTMENU:        sRtn = "\Start Menu" & cu
      Case CSIDL_PROGRAMS:         sRtn = "\Start Menu\Programs" & cu
      Case CSIDL_ADMINTOOLS:       sRtn = "\Start Menu\Programs\Administrative Tools" & cu
      Case CSIDL_STARTUP:          sRtn = "\Start Menu\Programs\Startup" & cu
      Case CSIDL_TEMPLATES:        sRtn = "\Templates" & cu
      Case CSIDL_INTERNET_CACHE:   sRtn = "\Temporary Internet Files" & cu
      ' Need a better definition ----------------------------------
      Case CSIDL_ALTSTARTUP:       sRtn = "Nonlocalized Startup group" & cu
      
      ' Virtual folders - no filesystem equivalent!
      Case CSIDL_COMPUTERSNEARME:  sRtn = "Computers Near Me" & vi
      Case CSIDL_DESKTOP:          sRtn = "Desktop" & vi
      Case CSIDL_INTERNET:         sRtn = "Internet Explorer" & vi
      Case CSIDL_DRIVES:           sRtn = "My Computer" & vi
      Case CSIDL_CONTROLS:         sRtn = "My Computer\Control Panel" & vi
      Case CSIDL_PRINTERS:         sRtn = "My Computer\Printers" & vi
      Case CSIDL_MYDOCUMENTS:      sRtn = "My Documents" & vi
      Case CSIDL_CONNECTIONS:      sRtn = "Network and Dial-up Connections"
      Case CSIDL_NETWORK:          sRtn = "Network Neighborhood" & vi
      Case CSIDL_BITBUCKET:        sRtn = "Recycle Bin" & vi
      
      ' Flags
      Case CSIDL_FLAG_CREATE:      sRtn = "Combine with any other CSIDL to force " & _
                                          "the creation of the associated folder."
      
      ' Bad request or unknown constant
      Case Else:                   sRtn = "{unknown constant}"
   End Select
End Property

' ***********************************************
'  Private Methods
' ***********************************************
Private Function Exported(ByVal ModuleName As String, ByVal ProcName As String) As Boolean
   Dim hModule As Long
   Dim lpProc As Long
   Dim FreeLib As Boolean
   
   ' check first to see if the module is already
   ' mapped into this process.
   hModule = GetModuleHandle(ModuleName)
   If hModule = 0 Then
      ' need to load module into this process.
      hModule = LoadLibrary(ModuleName)
      FreeLib = True
   End If
   
   ' if the module is mapped, check procedure
   ' address to verify it's exported.
   If hModule Then
      lpProc = GetProcAddress(hModule, ProcName)
      Exported = (lpProc <> 0)
   End If
   
   ' unload library if we loaded it here.
   If FreeLib Then Call FreeLibrary(hModule)
End Function

Private Function IsDirectory(ByVal PathSpec As String) As Boolean
   On Error Resume Next
   IsDirectory = ((GetAttr(PathSpec) And vbDirectory) = vbDirectory)
End Function

Public Function TrimNull(ByVal StrZ As String) As String
   ' VB5 or VB6!
   TrimNull = Left$(StrZ, InStr(StrZ & vbNullChar, vbNullChar) - 1)
End Function
