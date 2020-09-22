VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DO NOT DELETE OR MODIFY - IDENTIFIER BY LINEID ADD-IN
'LINE NUMBER STATUS = 0
Option Explicit

'api declerations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                     (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
 
'Private Const VER_PLATFORM_WIN32s = 0
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'Private Const VER_PLATFORM_WIN32_NT = 2

Private lngOSPlatform As Long   '1 for 9X series and 2 for NT series
Private lngOSVersion As Long
'The lngOSVersion will contain one of these values according to the results obtained
'
'0 -> Unknown
'1 -> Windows 95
'2 -> Windows 95B
'3 -> Windows 98
'4 -> Windows 98SE
'5 -> Windows ME
'6 -> Windows NT 3.51
'7 -> Winodows NT 4
'8 -> Windows 2000
'
'=======================================================================================

'get the version number of the operating system
Private Function GetOSVersion() As Long
Dim version As OSVERSIONINFO

    version.dwOSVersionInfoSize = Len(version)
    GetVersionEx version

    If version.dwPlatformId = 1 And version.dwMinorVersion = 0 And LoWord(version.dwBuildNumber) = 950 Then
        lngOSVersion = 1: lngOSPlatform = 1
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 0 And LoWord(version.dwBuildNumber) = 1111 Then
        lngOSVersion = 2: lngOSPlatform = 1
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 10 And LoWord(version.dwBuildNumber) = 1998 Then
        lngOSVersion = 3: lngOSPlatform = 1
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 10 And LoWord(version.dwBuildNumber) = 2222 Then
        lngOSVersion = 4: lngOSPlatform = 1
    ElseIf version.dwPlatformId = 1 And version.dwMinorVersion = 90 And LoWord(version.dwBuildNumber) = 3000 Then
        lngOSVersion = 5: lngOSPlatform = 1
    End If
            
    If version.dwPlatformId = 2 And version.dwMajorVersion = 3 Then
        lngOSVersion = 6: lngOSPlatform = 2
    ElseIf version.dwPlatformId = 2 And version.dwMajorVersion = 4 Then
        lngOSVersion = 7: lngOSPlatform = 2
    ElseIf version.dwPlatformId = 2 And version.dwMajorVersion = 5 Then
        lngOSVersion = 8: lngOSPlatform = 2
    End If
    
End Function

Private Function LoWord(lngIn As Long) As Integer
   If (lngIn And &HFFFF&) > &H7FFF Then
      LoWord = (lngIn And &HFFFF&) - &H10000
   Else
      LoWord = lngIn And &HFFFF&
   End If
End Function

Public Property Get IsPlatformWin9X() As Boolean
    Call GetOSVersion
    If lngOSPlatform = 1 Then IsPlatformWin9X = True
End Property

Public Property Get IsPlatformWinNT() As Boolean
    Call GetOSVersion
    If lngOSPlatform = 2 Then IsPlatformWinNT = True
End Property

