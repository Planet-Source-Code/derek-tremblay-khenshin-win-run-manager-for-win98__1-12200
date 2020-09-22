Attribute VB_Name = "Function_Module"
  Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
  Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
  Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
  Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Declare Function WritePrivateProfileByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
  Declare Function WritePrivateProfileToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
  Declare Function WritePrivateProfileToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)


Function GetPrivateStringValue(section$, Key$, File$) As String

  Dim KeyValue$
  Dim characters As Long
      
  KeyValue$ = String$(128, 0)
      
  characters = GetPrivateProfileStringByKeyName(section$, Key$, "", KeyValue$, 127, File$)
  
  If characters > 1 Then
     KeyValue$ = Left$(KeyValue$, characters)
  End If
      
  GetPrivateStringValue = KeyValue$

End Function

Function AppVertion(ObectToShow As Object)
  ObectToShow.Caption = "Vertion : " & App.Major & "." & App.Minor & " Build #" & App.Revision & ", Derek Tremblay, E-mail: khanjar@videotron.ca , ICQ: 23439375"
End Function

