Attribute VB_Name = "modRegistry"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, hKeyHandle As Long) As Long
   
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const BUFFER_LENGTH = 255
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
    KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Function GetRegistryKey(sKeyName As String, sValue As String, _
   Optional bClasses_Root As Boolean = False, Optional bCurrent_User As Boolean = False, _
   Optional bLocal_Machine As Boolean = False, Optional bUsers As Boolean = False, _
   Optional bCurrent_Config As Boolean = False, Optional bDyn_Data As Boolean = False) As String

    Dim sReturnBuffer As String
    Dim iBufLen As Long
    Dim iReturn As Long
    Dim iTree As Long
    Dim hKeyHandle As Long
   
    If bClasses_Root Then
        iTree = HKEY_CLASSES_ROOT
    ElseIf bCurrent_User Then
        iTree = HKEY_CURRENT_USER
    ElseIf bLocal_Machine Then
        iTree = HKEY_LOCAL_MACHINE
    ElseIf bUsers Then
        iTree = HKEY_USERS
    ElseIf bCurrent_Config Then
        iTree = HKEY_CURRENT_CONFIG
    ElseIf bDyn_Data Then
        iTree = HKEY_DYN_DATA
    Else
        GetRegistryKey = ""
        Exit Function
    End If
    
    'Set up return buffer
    sReturnBuffer = Space(BUFFER_LENGTH)
    iBufLen = BUFFER_LENGTH
    iReturn = RegOpenKeyEx(iTree, sKeyName, 0, KEY_ALL_ACCESS, hKeyHandle)
    
    If iReturn = ERROR_SUCCESS Then
        iReturn = RegQueryValueExString(hKeyHandle, sValue, 0, 0, sReturnBuffer, iBufLen)
        If iReturn = ERROR_SUCCESS Then
            GetRegistryKey = Left$(sReturnBuffer, iBufLen - 1)
        Else
            GetRegistryKey = ""
        End If
    Else
        GetRegistryKey = ""
    End If
    
    RegCloseKey hKeyHandle
    
End Function
