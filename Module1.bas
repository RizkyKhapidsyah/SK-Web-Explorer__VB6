Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)


Public Const MAX_PATH = 260



Public Enum SpecialShellFolderIDs

  CSIDL_FAVORITES = &H6
  CSIDL_HISTORY = &H22
End Enum


Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" _
                              (ByVal hwndOwner As Long, _
                              ByVal pszPath As String, _
                              ByVal nFolder As SpecialShellFolderIDs, _
                              ByVal fCreate As Boolean) As Long



Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, _
                              ByVal nFolder As SpecialShellFolderIDs, _
                              pidl As Long) As Long


Public Const NOERROR = 0

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pidl As Long, _
                              ByVal pszPath As String) As Long


Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'




Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2
End Function


Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer
  cb = GetItemIDSize(pidl)
 
  If cb Then GetNextItemID = pidl + cb
End Function


Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer

  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop

    GetPIDLSize = cb + 2
  End If
  
Out:
End Function



Public Function GetSpecialFolderPath(hWnd As Long, _
                                                             nFolder As SpecialShellFolderIDs) As String
  Dim pidl As Long
  Dim sPath As String * MAX_PATH
  

  On Error GoTo NotExported

  Call SHGetSpecialFolderPath(hWnd, sPath, nFolder, 0)
 
  If InStr(sPath, vbNullChar) > 1 Then
    GetSpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    Exit Function
  End If
  
NotExported:

  
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
      
      If SHGetPathFromIDList(pidl, sPath) Then
      
        GetSpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      End If
      
      Call CoTaskMemFree(pidl)
    End If
  End If

End Function





