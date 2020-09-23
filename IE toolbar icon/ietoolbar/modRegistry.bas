Attribute VB_Name = "modRegistry"
' Module      : modRegistry
' Description : This module Implements routines for manipulating the registry.
' Source      : Total VB SourceBook 6
'

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey _
  Lib "advapi32.dll" _
  (ByVal lngHKey As Long) _
  As Long

Private Declare Function RegCreateKeyEx _
  Lib "advapi32.dll" _
  Alias "RegCreateKeyExA" _
  (ByVal lngHKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) _
  As Long

Private Declare Function RegOpenKeyEx _
  Lib "advapi32.dll" _
  Alias "RegOpenKeyExA" _
  (ByVal lngHKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) _
  As Long

Private Declare Function RegQueryValueExString _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long) _
  As Long

Private Declare Function RegQueryValueExLong _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long) _
  As Long

Private Declare Function RegQueryValueExBinary _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As Long, _
    lpcbData As Long) _
  As Long
  
Private Declare Function RegQueryValueExNULL _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As Long, _
    lpcbData As Long) _
  As Long

Private Declare Function RegSetValueExString _
  Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    ByVal lpValue As String, _
    ByVal cbData As Long) _
  As Long

Private Declare Function RegSetValueExLong _
  Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpValue As Long, _
    ByVal cbData As Long) _
  As Long

Private Declare Function RegSetValueExBinary _
  Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    ByVal lpValue As Long, _
    ByVal cbData As Long) _
  As Long
  
Private Declare Function RegEnumKey _
  Lib "advapi32.dll" _
  Alias "RegEnumKeyA" _
  (ByVal lngHKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    ByVal cbName As Long) _
  As Long

Private Declare Function RegQueryInfoKey _
  Lib "advapi32.dll" _
  Alias "RegQueryInfoKeyA" _
  (ByVal lngHKey As Long, _
    ByVal lpClass As String, _
    ByVal lpcbClass As Long, _
    ByVal lpReserved As Long, _
    lpcSubKeys As Long, _
    lpcbMaxSubKeyLen As Long, _
    ByVal lpcbMaxClassLen As Long, _
    lpcValues As Long, _
    lpcbMaxValueNameLen As Long, _
    ByVal lpcbMaxValueLen As Long, _
    ByVal lpcbSecurityDescriptor As Long, _
    lpftLastWriteTime As FILETIME) _
  As Long

Private Declare Function RegEnumValue _
  Lib "advapi32.dll" _
  Alias "RegEnumValueA" _
  (ByVal lngHKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    ByVal lpType As Long, _
    ByVal lpData As Byte, _
    ByVal lpcbData As Long) _
  As Long

Private Declare Function RegDeleteKey _
  Lib "advapi32.dll" _
  Alias "RegDeleteKeyA" _
  (ByVal lngHKey As Long, _
    ByVal lpSubKey As String) _
  As Long

Private Declare Function RegDeleteValue _
  Lib "advapi32.dll" _
  Alias "RegDeleteValueA" _
  (ByVal lngHKey As Long, _
    ByVal lpValueName As String) _
  As Long

Public Enum EnumRegistryRootKeys
  rrkHKeyClassesRoot = &H80000000
  rrkHKeyCurrentUser = &H80000001
  rrkHKeyLocalMachine = &H80000002
  rrkHKeyUsers = &H80000003
End Enum

Public Enum EnumRegistryValueType
  rrkRegSZ = 1
  rrkregbinary = 3
  rrkRegDWord = 4
End Enum

Private Const mcregOptionNonVolatile = 0

Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259

Private Const mcregKeyAllAccess = &H3F
Private Const mcregKeyQueryValue = &H1

Public Sub RegistryCreateNewKey( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String)
  ' Comments  : Creates a new key in the system registry
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key to create
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long
    
  On Error GoTo PROC_ERR
    
  ' Create the key
  lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
    mcregOptionNonVolatile, mcregKeyAllAccess, 0&, lngHKey, 0&)
    
  ' if the key was created, then close it
  If lngRetVal = mcregErrorNone Then
    RegCloseKey (lngHKey)
  End If
    
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryCreateNewKey"
  Resume PROC_EXIT
    
End Sub

Public Sub RegistryDeleteKey( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String)
  ' Comments  : Deletes a key from the system registry
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key to delete
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  
  On Error GoTo PROC_ERR
      
  ' Delete the key
  lngRetVal = RegDeleteKey(eRootKey, strKeyName)
    
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryDeleteKey"
  Resume PROC_EXIT
    
End Sub

Public Sub RegistryDeleteValue( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String, _
  strValueName As String)
  ' Comments  : Deletes a value from the system registry
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key to delete
  '             strValueName - The name of the value to delete
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long

  On Error GoTo PROC_ERR

  ' Open the key
  lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyAllAccess, _
    lngHKey)

  ' If the key was opened successfully, then delete it
  If lngRetVal = mcregErrorNone Then
    lngRetVal = RegDeleteValue(lngHKey, strValueName)
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryDeleteValue"
  Resume PROC_EXIT

End Sub

Public Sub RegistryEnumerateSubKeys( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String, _
  astrKeys() As String, _
  lngKeyCount As Long)
  ' Comments  : Enumerates the sub keys of the specified key
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key to enumerate
  '             astrKeys - An array of strings to fill with sub key names
  '             lngKeyCount - The number of sub keys returned in the parameter
  '             astrKeys
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long
  Dim lngKeyIndex As Long
  Dim strSubKeyName As String
  Dim lngSubkeyCount As Long
  Dim lngMaxKeyLen As Long
  Dim typFT As FILETIME
  
  On Error GoTo PROC_ERR
  
  ' Open the key
  lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyAllAccess, _
    lngHKey)
  
  If lngRetVal = mcregErrorNone Then
    'find the number of subkeys, and redim the return string array
    lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, lngSubkeyCount, _
      lngMaxKeyLen, 0, 0, 0, 0, 0, typFT)
    If mcregErrorNone = lngRetVal Then
      If lngSubkeyCount > 0 Then
        ReDim astrKeys(lngSubkeyCount - 1) As String
        
        'set up the while loop
        lngKeyIndex = 0
        ' Pad the string to the maximum length of a sub key, plus 1 for null
        ' termination
        lngMaxKeyLen = lngMaxKeyLen + 1
        strSubKeyName = Space$(lngMaxKeyLen)
        
        Do While RegEnumKey(lngHKey, lngKeyIndex, strSubKeyName, lngMaxKeyLen + 1) = 0
        
          ' Set the string array to the key name, removing null termination
          If InStr(1, strSubKeyName, vbNullChar) > 0 Then
            astrKeys(lngKeyIndex) = Left$(strSubKeyName, InStr(1, strSubKeyName, _
              vbNullChar) - 1)
          End If
          ' Increment the key index for the return string array
          lngKeyIndex = lngKeyIndex + 1
        
        Loop
      End If
      ' return the new dimension of the return string array
      lngKeyCount = lngSubkeyCount
    End If
    
    ' Close the key
    RegCloseKey (lngHKey)
  End If
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryEnumerateSubKeys"
  Resume PROC_EXIT

End Sub

Public Sub RegistryEnumerateValues( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String, _
  astrValues() As String, _
  lngValueCount As Long)
  ' Comments  : Enumerates the values of the specified key
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key to enumerate
  '             astrValues - An array of strings to fill with value names
  '             lngValueCount - The number of values returned in the parameter
  '             astrValues
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long
  Dim lngKeyIndex As Long
  Dim strValueName As String
  Dim lngTempValueCount As Long
  Dim lngMaxValueLen As Long
  Dim typFT As FILETIME
  
  On Error GoTo PROC_ERR
  
  ' Open the key
  lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyAllAccess, _
    lngHKey)
  
  If lngRetVal = mcregErrorNone Then
    'find the number of subkeys, and redim the return string array
    lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, 0, _
      0, 0, lngTempValueCount, lngMaxValueLen, 0, 0, typFT)
    If mcregErrorNone = lngRetVal Then
      If lngTempValueCount > 0 Then
        ReDim astrValues(lngTempValueCount - 1) As String
        
        'set up the while loop
        lngKeyIndex = 0
        ' Pad the string to the maximum length of a sub key, plus 1 for null
        ' termination
        lngMaxValueLen = lngMaxValueLen + 1
        strValueName = Space$(lngMaxValueLen)
        
        Do While RegEnumValue(lngHKey, lngKeyIndex, strValueName, _
          lngMaxValueLen + 1, 0, 0, 0, 0) = 0
        
          ' Set the string array to the key name, removing null termination
          If InStr(1, strValueName, vbNullChar) > 0 Then
            astrValues(lngKeyIndex) = Left$(strValueName, InStr(1, strValueName, _
              vbNullChar) - 1)
          End If
          ' Increment the key index for the return string array
          lngKeyIndex = lngKeyIndex + 1
        
        Loop
      End If
      ' return the new dimension of the return string array
      lngValueCount = lngTempValueCount
    End If
    
    ' Close the key
    RegCloseKey (lngHKey)
  End If
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryEnumerateValues"
  Resume PROC_EXIT

End Sub

Public Function RegistryGetKeyValue( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String, _
  strValueName As String) _
  As Variant
  ' Comments  : Returns a value from the system registry
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key
  '             strValueName - The name of the value
  ' Returns   : The data in the registry value
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long
  Dim varValue As Variant
  Dim strValueData As String
  Dim abytValueData() As Byte
  Dim lngValueData As Long
  Dim lngValueType As Long
  Dim lngDataSize As Long
  
  On Error GoTo PROC_ERR
  
  varValue = Empty
  
  lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0&, mcregKeyQueryValue, _
    lngHKey)
  
  If mcregErrorNone = lngRetVal Then
    
    lngRetVal = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, _
      0&, lngDataSize)
    
    If lngRetVal = mcregErrorNone Then
      
      Select Case lngValueType
      
      ' String type

        Case rrkRegSZ:
          If lngDataSize > 0 Then
            strValueData = String(lngDataSize, 0)
            lngRetVal = RegQueryValueExString(lngHKey, strValueName, 0&, _
              lngValueType, strValueData, lngDataSize)
            If InStr(strValueData, vbNullChar) > 0 Then
              strValueData = Mid$(strValueData, 1, InStr(strValueData, _
                vbNullChar) - 1)
            End If
          End If
          If mcregErrorNone = lngRetVal Then
            varValue = Left$(strValueData, lngDataSize)
          Else
            varValue = Empty
          End If
        
        ' Long type
        Case rrkRegDWord:
          lngRetVal = RegQueryValueExLong(lngHKey, strValueName, 0&, _
            lngValueType, lngValueData, lngDataSize)
          If mcregErrorNone = lngRetVal Then
            varValue = lngValueData
          End If
                
        ' Binary type
        Case rrkregbinary
          If lngDataSize > 0 Then
            ReDim abytValueData(lngDataSize) As Byte
            lngRetVal = RegQueryValueExBinary(lngHKey, strValueName, 0&, _
              lngValueType, VarPtr(abytValueData(0)), lngDataSize)
          End If
          If mcregErrorNone = lngRetVal Then
            varValue = abytValueData
          Else
            varValue = Empty
          End If
                
        Case Else
          'No other data types supported
          lngRetVal = -1
        
      End Select
      
    End If
    
    RegCloseKey (lngHKey)
  End If
  
  'Return varValue
  RegistryGetKeyValue = varValue
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryGetKeyValue"
  Resume PROC_EXIT
End Function

Public Sub RegistrySetKeyValue( _
  eRootKey As EnumRegistryRootKeys, _
  strKeyName As String, _
  strValueName As String, _
  varData As Variant, _
  eDataType As EnumRegistryValueType)
  ' Comments  : This procedure sets a key value
  ' Parameters: eRootKey - The root key
  '             strKeyName - The name of the key
  '             strValueName - The name of the value
  '             varData - The data to store in the value
  '             eDataType - The type of data to store in the value
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngRetVal As Long
  Dim lngHKey As Long
  Dim strData As String
  Dim lngData As Long
  Dim abytData() As Byte
    
  On Error GoTo PROC_ERR
  
  ' Open the specified key, if it does not exist then create it
  lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
    mcregOptionNonVolatile, mcregKeyAllAccess, 0&, lngHKey, 0&)
  
  ' Determine the data type of the key
  Select Case eDataType
  
  Case rrkRegSZ
    strData = varData & vbNullChar
    lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, eDataType, _
      strData, Len(strData))
    
  Case rrkRegDWord
    lngData = varData
    lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, eDataType, _
      lngData, Len(lngData))
  
  ' Binary type
  Case rrkregbinary
    abytData = varData
    lngRetVal = RegSetValueExBinary(lngHKey, strValueName, 0&, eDataType, _
      VarPtr(abytData(0)), UBound(abytData) + 1)
  
  End Select
  
  RegCloseKey (lngHKey)
    
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistrySetKeyValue"
  Resume PROC_EXIT
    
End Sub


