Attribute VB_Name = "basRegistry"
'
' Created by E.Spencer (elliot@spnc.demon.co.uk) - This code is public domain.
'
Option Explicit
'Security Mask constants
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
   KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE _
   Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
' Possible registry data types
Public Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum
' Registry value type definitions
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
' Registry section definitions
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
' Codes returned by Reg API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
' Registry API functions used in this module (there are more of them)
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

' This routine allows you to get values from anywhere in the Registry, it currently
' only handles string, double word and binary values. Binary values are returned as
' hex strings.
'
' Example
' Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "DefaultUserName")
'
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
Dim TStr1 As String, TStr2 As String
Dim i As Integer
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
lValueLength = Len(sValue)
lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   If lDataTypeValue = REG_BINARY Then
       ' Return a binary field as a hex string (2 chars per byte)
       TStr2 = ""
       For i = 1 To lValueLength
          TStr1 = Hex(Asc(Mid(sValue, i, 1)))
          If Len(TStr1) = 1 Then TStr1 = "0" & TStr1
          TStr2 = TStr2 + TStr1
       Next
       sValue = TStr2
   Else
      sValue = Left$(sValue, lValueLength - 1)
   End If
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sValue
End Function

' This routine allows you to write values into the entire Registry, it currently
' only handles string and double word values.
'
' Example
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValString, "NewValueHere"
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValDWord, "31"
'
Public Sub WriteRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String, ByVal ValType As InTypes, ByVal Value As Variant)
Dim lResult As Long
Dim lKeyValue As Long
Dim InLen As Long
Dim lNewVal As Long
Dim sNewVal As String
On Error Resume Next
lResult = RegCreateKey(Group, Section, lKeyValue)
If ValType = ValDWord Then
   lNewVal = CLng(Value)
   InLen = 4
   lResult = RegSetValueExLong(lKeyValue, Key, 0&, ValType, lNewVal, InLen)
Else
   ' Fixes empty string bug - spotted by Marcus Jansson
   If ValType = ValString Then Value = Value + Chr(0)
   sNewVal = Value
   InLen = Len(sNewVal)
   lResult = RegSetValueExString(lKeyValue, Key, 0&, 1&, sNewVal, InLen)
End If
lResult = RegFlushKey(lKeyValue)
lResult = RegCloseKey(lKeyValue)
End Sub

' This routine enumerates the subkeys under any given key
' Call repeatedly until "Not Found" is returned - store values in array or something
'
' Example - this example just adds all the subkeys to a string - you will probably want to
' save then into an array or something.
'
' Dim Res, NewLine As String
' Dim i As Long
' Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
' NewLine = ""
' Do Until Res = "Not Found"
'   Text1.Text = Text1.Text & NewLine & Res
'   i = i + 1
'   Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
'   NewLine = Chr(13) & Chr(10)
' Loop

Public Function ReadRegistryGetSubkey(ByVal Group As Long, ByVal Section As String, Idx As Long) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
lValueLength = Len(sValue)
lResult = RegEnumKey(lKeyValue, Idx, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   sValue = Left$(sValue, InStr(sValue, Chr(0)) - 1)
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistryGetSubkey = sValue
End Function

' This routine allows you to get all the values from anywhere in the Registry under any
' given subkey, it currently only returns string and double word values.
'
' Example - returns list of names/values to multiline text box
' Dim Res As Variant
' Dim i As Long
' Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
' Do Until Res(2) = "Not Found"
'    Text1.Text = Text1.Text & Chr(13) & Chr(10) & Res(1) & " " & Res(2)
'    i = i + 1
'    Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
' Loop
'
Public Function ReadRegistryGetAll(ByVal Group As Long, ByVal Section As String, Idx As Long) As Variant
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long
Dim lValueLength As Long, lValueNameLength As Long
Dim sValueName As String, sValue As String
Dim td As Double
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
sValueName = Space$(2048)
lValueLength = Len(sValue)
lValueNameLength = Len(sValueName)
lResult = RegEnumValue(lKeyValue, Idx, sValueName, lValueNameLength, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   sValue = Left$(sValue, lValueLength - 1)
   sValueName = Left$(sValueName, lValueNameLength)
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
' Return the datatype, value name and value as an array
ReadRegistryGetAll = Array(lDataTypeValue, sValueName, sValue)
End Function

' This routine deletes a specified key (and all its subkeys and values if on Win95) from the registry.
' Be very careful using this function.
'
' Example
' DeleteSubkey HKEY_CURRENT_USER, "Software\My Name\My App"
'
Public Function DeleteSubkey(ByVal Group As Long, ByVal Section As String) As String
Dim lResult As Long, lKeyValue As Long
On Error Resume Next
lResult = RegOpenKeyEx(Group, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue)
lResult = RegDeleteKey(lKeyValue, Section)
lResult = RegCloseKey(lKeyValue)
End Function

' This routine deletes a specified value from below a specified subkey.
' Be very careful using this function.
'
' Example
' DeleteValue HKEY_CURRENT_USER, "Software\My Name\My App", "NewSubKey"
'
Public Function DeleteValue(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
Dim lResult As Long, lKeyValue As Long
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
lResult = RegDeleteValue(lKeyValue, Key)
lResult = RegCloseKey(lKeyValue)
End Function


