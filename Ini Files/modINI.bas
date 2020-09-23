Attribute VB_Name = "modINI"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' Sub for writing a string in an ini file
Public Sub WriteString(Section As String, Key As String, Value As String, File As String)
    WritePrivateProfileString Section, Key, Value, File
End Sub

' Reads an integer from the ini file
Public Function ReadNumber(Section As String, Key As String, FileName As String) As Long
    ReadNumber = GetPrivateProfileInt(Section, Key, 0, FileName)
End Function

' Function for reading and returning a string from an ini file
Public Function ReadString(Section As String, Key As String, FileName As String) As String
    Dim s As String
    
    ' Create a buffer to hold enough data (enough = 255)
    s = String(255, vbNullChar)
    
    GetPrivateProfileString Section, Key, "", s, Len(s), FileName
    
    ' Get rid of trailing vbNullChars
    s = Left(s, InStr(s, vbNullChar) - 1)
    
    ReadString = s
End Function
