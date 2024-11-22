Attribute VB_Name = "ModNC"



'                                       ini file declarations
'******************************************************************************************************
Declare Function GetPrivateProfileString Lib "Kernel32.dll" Alias "GetPrivateProfileStringA" ( _
                                         ByVal lpApplicationName As String, _
                                         ByVal lpKeyName As Any, _
                                         ByVal lpDefault As String, _
                                         ByVal lpReturnedString As String, _
                                         ByVal nSize As Long, _
                                         ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32.dll" Alias "WritePrivateProfileStringA" ( _
                                           ByVal lpApplicationName As String, _
                                           ByVal lpKeyName As String, _
                                           ByVal lpString As String, _
                                           ByVal lpFileName As String) As Long


