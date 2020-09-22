VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'APIs to open/close the registry
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

'APIs to get/set values in the registry
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, _
        lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpData As String, ByVal cbData As Long) As Long
        
'APIs to create/remove keys
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long

'Error codes
Private Const ERROR_SUCCESS = 0

'Registry constants
Private Const REG_DWORD = 4
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1

'Access constants
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Enum for reg data types, for easy rememberence
Public Enum RegDataTypes
    eREG_BINARY = 3
    eREG_DWORD = 4
    eREG_DWORD_BIG_ENDIAN = 5
    eREG_DWORD_LITTLE_ENDIAN = 4
    eREG_EXPAND_SZ = 2
    eREG_MULTI_SZ = 7
    eREG_SZ = 1
End Enum


'This enum makes it easier than remembering constants
'It will help make the class function more like what
'the user is used to in RegEdit and RegEdt32.
Public Enum HKEYs
    eHKEY_CLASSES_ROOT = &H80000000
    eHKEY_CURRENT_USER = &H80000001
    eHKEY_LOCAL_MACHINE = &H80000002
    eHKEY_USERS = &H80000003
    eHKEY_PERFORMANCE_DATA = &H80000004
    eHKEY_CURRENT_CONFIG = &H80000005
    eHKEY_DYN_DATA = &H80000006
End Enum

Public Function CreateKey(PredefinedKey As HKEYs, KeyName As String) As Boolean
    Dim hNewKey As Long
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim rc As Long

    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If

    'Create the branch
    rc = RegCreateKeyEx(PredefinedKey, _
            KeyName, _
            0&, _
            vbNullString, _
            REG_OPTION_NON_VOLATILE, _
            KEY_ALL_ACCESS, _
            lpSecurityAttributes, _
            hNewKey, _
            rc)
    
    If rc = ERROR_SUCCESS Then
    
        'Close the registry
        rc = RegCloseKey(hNewKey)
    
        'Return the result code
        CreateKey = True
    Else
        CreateKey = False
    End If
    
    'Bypass the error handler
    Exit Function

handler:
    CreateKey = False
End Function

Public Function SetValue(PredefinedKey As HKEYs, KeyName As String, ValueName As String, _
        Value As Variant, Optional ValueType As RegDataTypes = 1) As Boolean
    Dim rc As Long
    Dim hKey As Long
    Dim lpType As Long
    Dim lpcbData As Long
    Dim lpData As String
    
    On Error GoTo handler
    
    'Make sure there is no backslash preceding the branch
    If Left$(KeyName, 1) = "\" Then
        KeyName = Right$(KeyName, Len(KeyName) - 1)
    End If
    
    'first call the create key so that if the key is not present, it will get created
    CreateKey PredefinedKey, KeyName
    
    'Open the registry
    rc = RegOpenKeyEx(PredefinedKey, _
            KeyName, _
            0, _
            KEY_ALL_ACCESS, _
            hKey)

    If rc = ERROR_SUCCESS Then
        'Create a buffer so we can retrieve the data type of
        'the key. We'll need this to determine which API
        'we should call.
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
        'Get the value type first.
        'It will be returned via lpType argument
        rc = RegQueryValueEx(hKey, _
            ValueName, _
            0, lpType, _
            ByVal lpData, _
            lpcbData)
        
        If rc = ERROR_SUCCESS Then
            Select Case lpType
                Case REG_SZ
                    'Use a string data type
                    rc = RegSetValueExString(hKey, _
                        ValueName, _
                        0, _
                        REG_SZ, _
                        CStr(Value), _
                        Len(Value) + 1)
                Case REG_DWORD
                    'Use a DWORD data type
                    rc = RegSetValueEx(hKey, _
                        ValueName, _
                        0, _
                        REG_DWORD, _
                        CLng(Value), _
                        lpcbData)
            End Select
        Else
            'if the key not present create a key depending upon the value type provided
            Select Case ValueType
                Case eREG_DWORD
                    'Use a DWORD data type
                    rc = RegSetValueEx(hKey, _
                        ValueName, _
                        0&, _
                        REG_DWORD, _
                        CLng(Value), _
                        4)
                Case eREG_SZ
                    'Use a string data type
                    rc = RegSetValueExString(hKey, _
                        ValueName, _
                        0, _
                        REG_SZ, _
                        ByVal CStr(Value), _
                        Len(CStr(Value)) + 1)
            End Select
        End If
        
        'Close the registry
        RegCloseKey hKey
    End If
    
    'Return the result code
    SetValue = True
    
    'Bypass the error handler
    Exit Function

handler:
    SetValue = False
End Function
