Attribute VB_Name = "DSNMODULE"
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_DWORD = 4 'if number is store
Private Const REG_SZ = 1    'if string is store
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_USERS = &H80000003 'for date style

Private Const DRIVERNAME As String = "Microsoft Dbase Driver (*.dbf)"
Private Const SUBPATH As String = "Software\ODBC\ODBC.INI\idbase" 'here idbase is dsn name


'constants for idbase folder in odbc
Private Const DRIVER As String = "C:\WINDOWS\SYSTEM\odbcjt32.dll"
Private Const DRIVERID As String = "533"
Private Const FILL As String = "dbase 5.0" 'dbase version
Private Const SAFETRANSCTION As String = "0"
Private Const UID As String = "" 'user id

'constant for Xbase folder in odbc
Private Const COLLATINGSEQUENCE   As String = "ASCII"
Private Const DELETED As String = "01"
Private Const IMPLICIT As String = ""
Private Const PAGETIMEOUT As String = "600"
Private Const STAT As String = "0"
Private Const THREADES As String = "3"
Private Const USERCOMMIT As String = "Yes"

  'function to check for dsn already exists or not
Private Function CheckDsn() As Boolean
Dim Result As Long, keyHand As Long, Hdata As Long, Sdata As Long
Result = RegOpenKey(HKEY_CURRENT_USER, SUBPATH, keyHand)
If keyHand = 0 Then CheckDsn = False Else CheckDsn = True
RegCloseKey (keyHand)
End Function


'function creates dsn ,if sucessfuly generated then returns true else false

Public Function Dsn(AppPath As String) As Boolean
If CheckDsn = False Then
  Dim Result As Long, keyHand As Long, SubKeyHand As Long
  'opening main key Hdbase
  Result = RegCreateKey(HKEY_CURRENT_USER, SUBPATH, keyHand)
  Result = RegSetValueEx(keyHand, "DefalutDir", 0&, REG_SZ, ByVal AppPath, Len(AppPath))
  Result = RegSetValueEx(keyHand, "Driver", 0&, REG_SZ, ByVal DRIVER, Len(DRIVER))
  Result = RegSetValueEx(keyHand, "FIL", 0&, REG_SZ, ByVal FILL, Len(FILL))
  Result = RegSetValueEx(keyHand, "DriverId", 0&, REG_DWORD, CLng(DRIVERID), 4)
  Result = RegSetValueEx(keyHand, "SafeTransction", 0&, REG_DWORD, CLng(SAFETRANSCTION), 4)
  Result = RegSetValueEx(keyHand, "UID", 0&, REG_SZ, ByVal UID, Len(UID))
  'opening SubKey Xbase
  Result = RegCreateKey(HKEY_CURRENT_USER, SUBPATH & "\Engines\Xbase", SubKeyHand)
  Result = RegSetValueEx(SubKeyHand, "CollatingSequence", 0&, REG_SZ, ByVal COLLATINGSEQUENCE, Len(COLLATINGSEQUENCE))
  Result = RegSetValueEx(SubKeyHand, "Deleted", 0&, REG_DWORD, CLng(DELETED), 4)
  Result = RegSetValueEx(SubKeyHand, "ImplicitCommitSync", 0&, REG_SZ, ByVal IMPLICIT, Len(IMPLICIT))
  Result = RegSetValueEx(SubKeyHand, "PageTimeout", 0&, REG_DWORD, CLng(PAGETIMEOUT), 4)
  Result = RegSetValueEx(SubKeyHand, "Statistics", 0&, REG_DWORD, CLng(STAT), 4)
  Result = RegSetValueEx(SubKeyHand, "Threads", 0&, REG_DWORD, CLng(THREADES), 4)
  Result = RegSetValueEx(SubKeyHand, "UserCommitSync", 0&, REG_SZ, ByVal USERCOMMIT, Len(USERCOMMIT))
  'closeing sub key and key
  Result = RegCloseKey(SubKeyHand)
  Result = RegCloseKey(keyHand)
  'listing our Dsn in odbc list
  Result = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", keyHand)
  Result = RegSetValueEx(keyHand, "idbase", 0&, REG_SZ, ByVal DRIVERNAME, Len(DRIVERNAME))
  Result = RegCloseKey(keyHand)
  If Result = 0 Then Dsn = True Else Dsn = False
Else
  Dsn = True
End If
End Function


  ' at form level

'Private Sub form_load()
'   If Dsn(App.PATH) = True Then
'    MsgBox "Ok"
'   Else
'        MsgBox "Failed"
'   End If
'End Sub



