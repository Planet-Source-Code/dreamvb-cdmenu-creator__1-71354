Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public CtrlObj As Object
Public MyIni As New dINIFile

Public ButtonPress As VbMsgBoxResult
Public ProjPath As String
Public ProjResPath As String

'Picture Path location
Public mPictureLoc As String

'Browse for type
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'Browse for consts
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40

Public Function GetFolder(ByVal iHwnd As Long, ByVal sTitle As String, Optional FolRoot As Long = 0)
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim ppidl As Long
Dim h As String

    RetVal = SHGetSpecialFolderLocation(iHwnd, FolRoot, ppidl)

    With bInf
        .hOwner = iHwnd
        .lpszTitle = sTitle
        .ulFlags = (BIF_RETURNFSANCESTORS Or BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE)
        .pidlRoot = ppidl
        .lpfn = 0&
        .lParam = 0&
    End With
    
    'Get Path ID
    PathID = SHBrowseForFolder(bInf)
    
    If (PathID) Then
        'Create Buffer
        RetPath = Space$(512)
        'Get Folder Path
        If SHGetPathFromIDList(ByVal PathID, ByVal RetPath) Then
            'Strip nullchars
            GetFolder = Left$(RetPath, InStr(RetPath, Chr$(0)) - 1)
            CoTaskMemFree PathID
        End If
    End If
    If (ppidl) Then CoTaskMemFree ppidl
End Function

Public Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function FindFolder(ByVal FolderName As String) As Boolean
On Error Resume Next
    FindFolder = (GetAttr(FolderName) And vbDirectory) = vbDirectory
    Err.Clear
End Function

Public Function GetFilename(lFile As String) As String
Dim sPos As Integer
    
    If Len(lFile) > 0 Then
        'Find last slash.
        sPos = InStrRev(lFile, "\", Len(lFile), vbBinaryCompare)
        If (sPos > 0) Then
            'Return Filename.
            GetFilename = Mid$(lFile, sPos + 1)
        Else
            GetFilename = lFile
        End If
    End If
End Function

Public Function PathFromFile(ByVal FileName As String) As String
Dim sPos As Integer
    sPos = InStrRev(FileName, "\", Len(FileName), vbBinaryCompare)
    
    If (sPos > 0) Then
        PathFromFile = Mid(FileName, 1, sPos - 1)
    End If
End Function
