Attribute VB_Name = "BasRegRudyz"
Option Explicit

Enum EnumRegHKey
     
    RegLocalMachine
    RegCurrentUser

End Enum

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hkey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hkey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hkey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long

Public Const REG_SZ = 1
Public Const REG_DWORD = 1

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const KEY_ALL_ACCESS = &HF003F
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_EXECUTE = &H20019
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_READ = &H20019
Public Const KEY_SET_VALUE = &H2
Public Const KEY_WRITE = &H20006

Public Sub RegCreatePath(TipoRegHKey As EnumRegHKey, strPath As String)

    On Error GoTo ErrorProcedure
    
    Dim H        As Long
    Dim hkey     As Long
    Dim KeyHand  As Long
    
    If TipoRegHKey = RegCurrentUser Then
        hkey = HKEY_CURRENT_USER
    Else
        hkey = HKEY_LOCAL_MACHINE
    End If
    
    If InStr(1, strPath, "Software\RudyZ", vbTextCompare) = 0 Then
        strPath = "Software\RudyZ\" + strPath
    End If
    
    H = RegCreateKey(hkey, strPath, KeyHand)
    
Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

End Sub

Public Sub RegWriteSetting(TipoRegHKey As EnumRegHKey, Applicazione As String, Sezione As String, Variabile As String, Valore As String)
      
    On Error GoTo ErrorProcedure
    
    Dim StrData  As String
    Dim H        As Long
    Dim hkey     As Long
    Dim KeyHand  As Long
    Dim strPath  As String
    
    If TipoRegHKey = RegCurrentUser Then
        hkey = HKEY_CURRENT_USER
    Else
        hkey = HKEY_LOCAL_MACHINE
    End If
    
    strPath = "Software\RudyZ"
    
    If Applicazione <> "" Then
        strPath = strPath + "\" + Applicazione
    End If
    
    If Sezione <> "" Then
        strPath = strPath + "\" + Sezione
    End If
    
    H = RegOpenKeyEx(hkey, strPath, 0, KEY_WRITE, KeyHand)
    
    If H <> 0 Then
         RegCreatePath TipoRegHKey, strPath
         H = RegOpenKeyEx(hkey, strPath, 0, KEY_WRITE, KeyHand)
    End If
    
    H = RegSetValueEx(KeyHand, Variabile, 0, REG_SZ, ByVal Valore, Len(Valore))
    H = RegCloseKey(KeyHand)
     
Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

End Sub

Public Function RegReadSetting(TipoRegHKey As EnumRegHKey, Applicazione As String, Sezione As String, Variabile As String, Default As String) As String

    On Error GoTo ErrorProcedure
    
    Dim StrData    As String
    Dim PosZero    As Integer
    Dim H          As Long
    Dim Esito      As Long
    Dim KeyHand    As Long
    Dim strPath    As String
    Dim hkey       As Long
    
    If TipoRegHKey = RegCurrentUser Then
        hkey = HKEY_CURRENT_USER
    Else
        hkey = HKEY_LOCAL_MACHINE
    End If
    
    strPath = "Software\RudyZ"
    
    If Applicazione <> "" Then
        strPath = strPath + "\" + Applicazione
    End If
    
    If Sezione <> "" Then
        strPath = strPath + "\" + Sezione
    End If
    
    H = RegOpenKeyEx(hkey, strPath, 0, KEY_READ, KeyHand)
    
    StrData = Space(255)
    Esito = RegQueryValueEx(KeyHand, Variabile, 0, REG_SZ, ByVal StrData, 255)
    
    H = RegCloseKey(KeyHand)
    
    If Esito = 0 Then
    
        PosZero = InStr(StrData, Chr(0))
        If PosZero > 0 Then
            StrData = Left(StrData, PosZero - 1)
        End If
    
    Else
    
        StrData = Default
    
    End If
    
    RegReadSetting = Trim(StrData)
     
Exit Function

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

    On Error Resume Next
    
    RegReadSetting = Default
    
End Function
