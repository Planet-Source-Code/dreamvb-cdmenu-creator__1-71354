VERSION 5.00
Begin VB.Form frmMenu 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBut 
      Caption         =   "Button"
      Height          =   390
      Index           =   0
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image ImgPic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   210
      Top             =   450
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyIni As dINIFile
Private mFilename As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Function ShortFile(ByVal Filename As String) As String
Dim iRet As Long
Dim sBuff As String
    'Create Buffer to hold File.
    sBuff = Space(256)
    
    iRet = GetShortPathName(Filename, sBuff, 256)
    If (iRet > 0) Then
        'Get short filename.
        ShortFile = Left$(sBuff, iRet)
    End If
    
    sBuff = vbNullString
End Function

Private Sub DoAction(CtrlName As String, Index As Integer)
Dim sName As String
Dim Action As Integer
Dim Ret As Long

    sName = CtrlName & Index
    'Get command action
    Action = Val(MyIni.ReadValue(sName, "Action"))

    'Check for execute program command
    If (Action = 0) Then
        sName = MyIni.ReadValue(sName, "Command")
        'Extract Command
        sName = Replace(sName, "$AppPath", FixPath(App.Path), , , vbTextCompare)
        If Len(sName) > 0 Then
            'Execute the command.
            Ret = RunApp(frmMenu.hwnd, "open", sName)
        End If
    End If
    
    'Check for exit command
    If (Action = 1) Then
        Unload frmMenu
    End If
    
    'Check for messagebox action
    If (Action = 2) Then
        MsgBox MyIni.ReadValue(sName, "Command"), vbInformation, frmMenu.Caption
    End If
    
    sName = vbNullString
    Ret = 0
    Action = 0
End Sub

Private Sub PlayMid(ByVal Filename As String)
Dim Ret As Long
    'Play Midi File
    mFilename = Filename
    Ret = mciSendString("Play " & mFilename, 0, 0, 0)
End Sub

Private Sub StopMid()
Dim Ret As Long
    'Stop current Mid File.
    Ret = mciSendString("Stop " & mFilename, 0, 0, 0)
End Sub

Private Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    'Returns true if a given filename is found.
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Private Function RunApp(iHwnd As Long, OpenOp As String, Filename As String)
Dim Ret As Long
    Ret = ShellExecute(iHwnd, OpenOp, Filename, "", "", 1)
End Function

Private Sub AddControl(CtrlTag As String)
Dim iCount As Integer
Dim Obj As Object

    Select Case UCase(CtrlTag)
        Case "BUTTON"
            iCount = cmdBut.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load cmdBut(iCount)
            End If
            'Set the object
            Set Obj = cmdBut(iCount - 1)
        Case "LABEL"
            iCount = lblA.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load lblA(iCount)
            End If
            'Set the object
            Set Obj = lblA(iCount - 1)
        Case "IMAGE"
            iCount = ImgPic.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load ImgPic(iCount)
            End If
            'Set the object
            Set Obj = ImgPic(iCount - 1)
    End Select
    
    Obj.Visible = True
    Set Obj = Nothing
    iCount = 0
End Sub

Private Sub LoadProject()
Dim ObjCount As Integer
Dim Count As Integer
Dim Obj As Object
Dim cName As String
Dim lFile As String

    'Load Main project Informaion
    frmMenu.Caption = MyIni.ReadValue("main", "Caption")
    frmMenu.BackColor = MyIni.ReadValue("main", "Backcolor")
    frmMenu.Height = Val(MyIni.ReadValue("main", "Height"))
    frmMenu.Width = Val(MyIni.ReadValue("main", "Width"))
    
    lFile = Replace(MyIni.ReadValue("main", "Sound"), "$AppPath", FixPath(App.Path), , , vbTextCompare)
    lFile = ShortFile(lFile)
    
    If Len(lFile) > 0 Then
        If FindFile(lFile) Then
            Call PlayMid(lFile)
        End If
    End If
    
    'Load Buttons
    ObjCount = Val(MyIni.ReadValue("main", "Buttons"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("BUTTON")
            'Set the object
            Set Obj = cmdBut(Count)
            cName = "Button" & Count
            Obj.Caption = MyIni.ReadValue(cName, "Caption")
            Obj.BackColor = MyIni.ReadValue(cName, "Backcolor")
            Obj.Top = MyIni.ReadValue(cName, "Top")
            Obj.Left = MyIni.ReadValue(cName, "Left")
            Obj.Width = MyIni.ReadValue(cName, "Width")
            Obj.Height = MyIni.ReadValue(cName, "Height")
            Obj.Tag = MyIni.ReadValue(cName, "ZOrder") & MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            
            'Obj.Tag = MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            Obj.FontName = MyIni.ReadValue(cName, "FontName")
            Obj.FontSize = MyIni.ReadValue(cName, "FontSize")
            Obj.FontBold = MyIni.ReadValue(cName, "FontBold")
            Obj.FontItalic = MyIni.ReadValue(cName, "FontItalic")
            Obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
        Next Count
    End If
    
    'Load Labels
    ObjCount = Val(MyIni.ReadValue("main", "Labels"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("LABEL")
            'Set the object
            Set Obj = lblA(Count)
            cName = "Label" & Count
            Obj.Caption = MyIni.ReadValue(cName, "Caption")
            Obj.ForeColor = MyIni.ReadValue(cName, "Forecolor")
            Obj.Top = MyIni.ReadValue(cName, "Top")
            Obj.Left = MyIni.ReadValue(cName, "Left")
            Obj.Width = MyIni.ReadValue(cName, "Width")
            Obj.Height = MyIni.ReadValue(cName, "Height")
            Obj.Tag = MyIni.ReadValue(cName, "ZOrder") & MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            'Obj.Tag = MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            Obj.FontName = MyIni.ReadValue(cName, "FontName")
            Obj.FontSize = MyIni.ReadValue(cName, "FontSize")
            Obj.FontBold = MyIni.ReadValue(cName, "FontBold")
            Obj.FontItalic = MyIni.ReadValue(cName, "FontItalic")
            Obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
        Next Count
    End If
    '
    'Load Images
    ObjCount = Val(MyIni.ReadValue("main", "Images"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("IMAGE")
            'Set the object
            Set Obj = ImgPic(Count)
            cName = "Image" & Count
            'Convert to the correct path
            lFile = Replace(MyIni.ReadValue(cName, "Picture"), "$AppPath", FixPath(App.Path), , , vbTextCompare)
            
            Obj.Top = MyIni.ReadValue(cName, "Top")
            Obj.Left = MyIni.ReadValue(cName, "Left")
            Obj.Height = MyIni.ReadValue(cName, "Height")
            Obj.Width = MyIni.ReadValue(cName, "Width")
            Obj.Stretch = Val(MyIni.ReadValue(cName, "Stretch"))

            'Check if the file is found
            If Not FindFile(lFile) Then
                Obj.BorderStyle = 1
            Else
                Obj.BorderStyle = MyIni.ReadValue(cName, "BorderStyle")
                Obj.Picture = LoadPicture(lFile)
            End If
            
            Obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
        Next Count
    End If
    
End Sub

Private Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Sub cmdBut_Click(Index As Integer)
    Call DoAction("Button", Index)
End Sub

Private Sub Form_Load()
    Set MyIni = New dINIFile
    MyIni.Filename = FixPath(App.Path) & App.EXEName & ".ini"
    
    'Check if the menu ini file is found.
    If (Not MyIni.IniFound) Then
        MsgBox "File Not Found:" & vbCrLf & vbCrLf & MyIni.Filename, vbCritical, "File Not Found"
        Unload frmMenu
        Exit Sub
    Else
        'Load the menu items
        Call LoadProject
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    frmMenu.Height = Val(MyIni.ReadValue("main", "Height", 3600))
    frmMenu.Width = Val(MyIni.ReadValue("main", "Width", 4800))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyIni = Nothing
    Set frmMenu = Nothing
End Sub

Private Sub ImgPic_Click(Index As Integer)
    Call DoAction("Image", Index)
End Sub

Private Sub lblA_Click(Index As Integer)
    Call DoAction("Label", Index)
End Sub
