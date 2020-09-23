VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000018&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.CommandButton cmdBut 
      Caption         =   "Button"
      Height          =   390
      Index           =   0
      Left            =   240
      MousePointer    =   15  'Size All
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "00"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image ImgSizeAll 
      Height          =   105
      Left            =   240
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmMain.frx":038A
      Top             =   1575
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image ImgPic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   240
      MousePointer    =   5  'Size
      Tag             =   "00"
      Top             =   1005
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   195
      Index           =   0
      Left            =   345
      MousePointer    =   5  'Size
      TabIndex        =   1
      Tag             =   "00"
      Top             =   705
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuA 
      Caption         =   "#"
      Visible         =   0   'False
      Begin VB.Menu mnuProp 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuFront 
         Caption         =   "&Bring to front"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Send to back"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldX As Integer
Private OldY As Integer
Private IsDown As Boolean

Public Sub NewProject()
    'Create a new blank project.
    Call DestroyControls
    frmMain.Move 0, 0, 4800, 3600
    frmMain.Caption = "Untitled"
    frmMain.BackColor = &H80000018
    
    cmdBut(0).Tag = "00"
    cmdBut(0).Caption = "Button"
    cmdBut(0).BackColor = vbButtonFace
    
    lblA(0).Tag = "00"
    lblA(0).ForeColor = vbBlack
    lblA(0).Caption = "Label"
    lblA(0).Width = 26
    lblA(0).Height = 13

    ImgPic(0).Width = 32
    ImgPic(0).Height = 32
    ImgPic(0).Stretch = False
    ImgPic(0).Tag = "00"
End Sub

Private Sub DestroyControls()
Dim c As Control
    
    'Unload the controls
    For Each c In frmMain.Controls
        If (c.Name = "cmdBut") Or (c.Name = "lblA") Or (c.Name = "ImgPic") Then
            c.Visible = False
            If (c.Index <> 0) Then
                Unload c
            End If
        End If
    Next c
End Sub

Public Sub LoadProject(ByVal FileName As String)
Dim ObjCount As Integer
Dim Count As Integer
Dim obj As Object
Dim cName As String
Dim lFile As String

    'Destory any exsiting controls
    Call DestroyControls
    'Main project Filename
    MyIni.FileName = FileName
    'Load Main project Informaion
    frmMain.Caption = MyIni.ReadValue("main", "Caption")
    frmMain.BackColor = MyIni.ReadValue("main", "Backcolor")
    frmMain.Tag = MyIni.ReadValue("main", "Sound")
    frmMain.Height = Val(MyIni.ReadValue("main", "Height"))
    frmMain.Width = Val(MyIni.ReadValue("main", "Width"))
    
    'Load Buttons
    ObjCount = Val(MyIni.ReadValue("main", "Buttons"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("BUTTON")
            'Set the object
            Set obj = cmdBut(Count)
            cName = "Button" & Count
            obj.Caption = MyIni.ReadValue(cName, "Caption")
            obj.BackColor = MyIni.ReadValue(cName, "Backcolor")
            obj.Top = MyIni.ReadValue(cName, "Top")
            obj.Left = MyIni.ReadValue(cName, "Left")
            obj.Width = MyIni.ReadValue(cName, "Width")
            obj.Height = MyIni.ReadValue(cName, "Height")
            obj.Tag = MyIni.ReadValue(cName, "ZOrder") & MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            '
            obj.FontName = MyIni.ReadValue(cName, "FontName")
            obj.FontSize = MyIni.ReadValue(cName, "FontSize")
            obj.FontBold = MyIni.ReadValue(cName, "FontBold")
            obj.FontItalic = MyIni.ReadValue(cName, "FontItalic")
            obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
        Next Count
    End If
    
    'Load Labels
    ObjCount = Val(MyIni.ReadValue("main", "Labels"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("LABEL")
            'Set the object
            Set obj = lblA(Count)
            cName = "Label" & Count
            obj.Caption = MyIni.ReadValue(cName, "Caption")
            obj.ForeColor = MyIni.ReadValue(cName, "Forecolor")
            obj.Top = MyIni.ReadValue(cName, "Top")
            obj.Left = MyIni.ReadValue(cName, "Left")
            obj.Width = MyIni.ReadValue(cName, "Width")
            obj.Height = MyIni.ReadValue(cName, "Height")
            obj.Tag = MyIni.ReadValue(cName, "ZOrder") & MyIni.ReadValue(cName, "Action") & MyIni.ReadValue(cName, "Command")
            '
            obj.FontName = MyIni.ReadValue(cName, "FontName")
            obj.FontSize = MyIni.ReadValue(cName, "FontSize")
            obj.FontBold = MyIni.ReadValue(cName, "FontBold")
            obj.FontItalic = MyIni.ReadValue(cName, "FontItalic")
            obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
        Next Count
    End If

    'Load Images
    ObjCount = Val(MyIni.ReadValue("main", "Images"))
    If (ObjCount > 0) Then
        For Count = 0 To (ObjCount - 1)
            Call AddControl("IMAGE")
            'Set the object
            Set obj = ImgPic(Count)
            cName = "Image" & Count
            'Convert to the correct path
            lFile = Replace(MyIni.ReadValue(cName, "Picture"), "$AppPath", ProjPath, , , vbTextCompare)
            
            obj.Top = MyIni.ReadValue(cName, "Top")
            obj.Left = MyIni.ReadValue(cName, "Left")
            obj.Height = MyIni.ReadValue(cName, "Height")
            obj.Width = MyIni.ReadValue(cName, "Width")
            obj.Stretch = Val(MyIni.ReadValue(cName, "Stretch"))
            obj.Tag = MyIni.ReadValue(cName, "ZOrder") & MyIni.ReadValue(cName, "Action") _
            & MyIni.ReadValue(cName, "Picture") & Chr(0) & MyIni.ReadValue(cName, "Command")
            
            'Check if the file is found
            If Not FindFile(lFile) Then
                obj.BorderStyle = 1
            Else
                obj.BorderStyle = MyIni.ReadValue(cName, "BorderStyle")
                obj.Picture = LoadPicture(lFile)
            End If

            obj.ZOrder Val(MyIni.ReadValue(cName, "ZOrder"))
            
        Next Count
    End If
End Sub

Public Sub SaveProject(ByVal FileName As String)
Dim Count As Integer
Dim sName As String
Dim obj As Object
Dim Tmp As String
Dim StrA As String
Dim sPos As Integer
    
    MyIni.FileName = FileName
    'Main Project Information
    MyIni.SetValue "main", "Caption", frmMain.Caption
    MyIni.SetValue "main", "Backcolor", frmMain.BackColor
    MyIni.SetValue "main", "Sound", frmMain.Tag
    MyIni.SetValue "main", "Width", frmMain.Width
    MyIni.SetValue "main", "Height", frmMain.Height
    MyIni.SetValue "main", "Labels", (lblA.Count - 1)
    MyIni.SetValue "main", "Buttons", (cmdBut.Count - 1)
    MyIni.SetValue "main", "Images", (ImgPic.Count - 1)
    
    'Save each label
    For Count = 0 To (lblA.Count - 1)
        If lblA(Count).Visible Then
            sName = "Label" & Count
            'Set the object
            Set obj = lblA(Count)
            MyIni.SetValue sName, "Caption", obj.Caption
            MyIni.SetValue sName, "Forecolor", obj.ForeColor
            
            MyIni.SetValue sName, "Action", Mid(obj.Tag, 2, 1)
            MyIni.SetValue sName, "Command", Mid(obj.Tag, 3)
            
            MyIni.SetValue sName, "Top", obj.Top
            MyIni.SetValue sName, "Left", obj.Left
            MyIni.SetValue sName, "Width", obj.Width
            MyIni.SetValue sName, "Height", obj.Height
            'Font Properties
            MyIni.SetValue sName, "FontName", obj.FontName
            MyIni.SetValue sName, "FontSize", obj.FontSize
            MyIni.SetValue sName, "FontBold", obj.FontBold
            MyIni.SetValue sName, "FontItalic", obj.FontItalic
            MyIni.SetValue sName, "ZOrder", Left(obj.Tag, 1)
        End If
    Next Count
    
    'Save each Button
    For Count = 0 To (cmdBut.Count - 1)
        If cmdBut(Count).Visible Then
            sName = "Button" & Count
            'Set the object
            Set obj = cmdBut(Count)
            MyIni.SetValue sName, "Caption", obj.Caption
            MyIni.SetValue sName, "BackColor", obj.BackColor
            MyIni.SetValue sName, "Action", Mid(obj.Tag, 2, 1)
            MyIni.SetValue sName, "Command", Mid(obj.Tag, 3)
            MyIni.SetValue sName, "Top", obj.Top
            MyIni.SetValue sName, "Left", obj.Left
            MyIni.SetValue sName, "Width", obj.Width
            MyIni.SetValue sName, "Height", obj.Height
            'Font Properties
            MyIni.SetValue sName, "FontName", obj.FontName
            MyIni.SetValue sName, "FontSize", obj.FontSize
            MyIni.SetValue sName, "FontBold", obj.FontBold
            MyIni.SetValue sName, "FontItalic", obj.FontItalic
            MyIni.SetValue sName, "ZOrder", Left(obj.Tag, 1)
        End If
    Next Count
    
     'Save each image
    For Count = 0 To (ImgPic.Count - 1)
        If ImgPic(Count).Visible Then
            StrA = vbNullString
            sName = "Image" & Count
            'Set the object
            Set obj = ImgPic(Count)
            'Get Object tag
            Tmp = obj.Tag
            'Get Chr 0 pos
            sPos = InStrRev(Tmp, Chr(0), Len(Tmp), vbBinaryCompare)
            '
            If (sPos > 0) Then
                'Extract Command
                StrA = Mid$(Tmp, sPos + 1)
                Tmp = Left$(Tmp, sPos - 1)
            End If
            
            MyIni.SetValue sName, "Picture", Mid(Tmp, 3)
            MyIni.SetValue sName, "Action", Mid(Tmp, 2, 1)
            MyIni.SetValue sName, "Command", StrA
            MyIni.SetValue sName, "Top", obj.Top
            MyIni.SetValue sName, "Left", obj.Left
            MyIni.SetValue sName, "Height", obj.Height
            MyIni.SetValue sName, "Width", obj.Width
            MyIni.SetValue sName, "BorderStyle", 0
            MyIni.SetValue sName, "Stretch", Abs(obj.Stretch)
            MyIni.SetValue sName, "ZOrder", Left(Tmp, 1)
            
        End If
    Next Count
    
    Set obj = Nothing
    sName = vbNullString
    Count = 0
    
End Sub

Private Sub DesignerAction(TheObj As Object, Action As Integer, Button As Integer, X As Single, Y As Single)
    Select Case Action
        Case 0
            'Mouse down
            OldX = (X \ Screen.TwipsPerPixelX)
            OldY = (Y \ Screen.TwipsPerPixelY)
            'Set the object
            Set CtrlObj = TheObj
            'Position the resizer
            ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
            ImgSizeAll.Visible = True
        Case 1
            'Mouse Move
            If (Button = vbLeftButton) Then
                CtrlObj.Left = (CtrlObj.Left + X \ Screen.TwipsPerPixelX) - OldX
                CtrlObj.Top = (CtrlObj.Top + Y \ Screen.TwipsPerPixelY) - OldY
                'Position the resizer
                ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
                ImgSizeAll.Visible = True
            End If
        Case 2
            'Mouse up
            Call FixObjectPos
            'Position the resizer
            ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
            ImgSizeAll.Visible = True
            
            If (Button = vbRightButton) Then
                Call PopupMenu(mnuA)
            End If
    End Select
End Sub

Private Sub FixObjectPos()
    'Check to see if the control has not gone of the form
    If (CtrlObj.Left < 0) Then CtrlObj.Left = 0
    If (CtrlObj.Top < 0) Then CtrlObj.Top = 0
            
    If (CtrlObj.Left > frmMain.ScaleWidth - CtrlObj.Width) Then
        CtrlObj.Left = (frmMain.ScaleWidth - CtrlObj.Width)
    End If
    If (CtrlObj.Top > frmMain.ScaleHeight - CtrlObj.Height) Then
        CtrlObj.Top = (frmMain.ScaleHeight - CtrlObj.Height)
    End If
End Sub

Private Sub CenterShowObject()
    CtrlObj.Left = (frmMain.ScaleWidth - CtrlObj.Width) \ 2
    CtrlObj.Top = (frmMain.ScaleHeight - CtrlObj.Height) \ 2
    
    CtrlObj.Tag = "00"
    CtrlObj.Visible = True
    CtrlObj.ZOrder vbBringToFront
End Sub

Public Sub AddControl(CtrlTag As String)
Dim iCount As Integer

    Select Case UCase(CtrlTag)
        Case "BUTTON"
            iCount = cmdBut.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load cmdBut(iCount)
            End If
            'Set the object
            Set CtrlObj = cmdBut(iCount - 1)
            CtrlObj.BackColor = vbButtonFace
            CtrlObj.Caption = "Button"
            Call ResetFont
        Case "LABEL"
            iCount = lblA.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load lblA(iCount)
            End If
            'Set the object
            Set CtrlObj = lblA(iCount - 1)
            CtrlObj.Width = 26
            CtrlObj.Height = 13
            CtrlObj.ForeColor = 0
            CtrlObj.Caption = "Label"
            Call ResetFont
        Case "IMAGE"
            iCount = ImgPic.Count
            If (iCount > 0) Then
                'Load control is more than 0
                Load ImgPic(iCount)
            End If
            'Set the object
            Set CtrlObj = ImgPic(iCount - 1)
            CtrlObj.Stretch = False
            CtrlObj.BorderStyle = 1
            CtrlObj.Height = 32
            CtrlObj.Width = 32
            Set CtrlObj.Picture = Nothing
    End Select
    
    Call CenterShowObject
    iCount = 0
End Sub

Private Sub ResetFont()
    CtrlObj.FontName = "MS Sans Serif"
    CtrlObj.FontSize = 8
    CtrlObj.FontBold = False
    CtrlObj.FontItalic = False
End Sub

Private Sub cmdBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(cmdBut(Index), 0, Button, X, Y)
End Sub

Private Sub cmdBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(cmdBut(Index), 1, Button, X, Y)
End Sub

Private Sub cmdBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(cmdBut(Index), 2, Button, X, Y)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If UCase(TypeName(CtrlObj)) = "FRMMAIN" Then
        Exit Sub
    Else
        'Check if Shift+vbKeyNum key is pressed
        If (Shift = 2) Then
            'Resize object
            Select Case KeyCode
                Case vbKeyNumpad8
                    CtrlObj.Height = (CtrlObj.Height - 1)
                Case vbKeyNumpad2
                    CtrlObj.Height = (CtrlObj.Height + 1)
                Case vbKeyNumpad4
                    CtrlObj.Width = (CtrlObj.Width - 1)
                Case vbKeyNumpad6
                    CtrlObj.Width = (CtrlObj.Width + 1)
            End Select
            KeyCode = 0
        End If
        
        'Move the object
        Select Case KeyCode
            Case vbKeyNumpad8
                CtrlObj.Top = (CtrlObj.Top - 1)
            Case vbKeyNumpad2
                CtrlObj.Top = (CtrlObj.Top + 1)
            Case vbKeyNumpad4
                CtrlObj.Left = (CtrlObj.Left - 1)
            Case vbKeyNumpad6
                CtrlObj.Left = (CtrlObj.Left + 1)
        End Select
        Call FixObjectPos
        'Position the resizer
        ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
    End If
End Sub

Private Sub Form_Load()
    Set CtrlObj = frmMain
    frmMain.Move 0, 0, 4800, 3600
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSizeAll.Visible = False
    Set CtrlObj = frmMain
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
        Call PopupMenu(mnuFile)
    End If
End Sub

Private Sub ImgPic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(ImgPic(Index), 0, Button, X, Y)
End Sub

Private Sub ImgPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(ImgPic(Index), 1, Button, X, Y)
End Sub

Private Sub ImgPic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(ImgPic(Index), 2, Button, X, Y)
End Sub

Private Sub ImgSizeAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        IsDown = True
        OldX = X
        OldY = Y
    End If
End Sub

Private Sub ImgSizeAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oWidth As Long
Dim oHeight As Long

    If (Button = vbLeftButton) And (IsDown) Then
        'Set Resizer position
        
        ImgSizeAll.Left = (ImgSizeAll.Left - (OldX - X) \ Screen.TwipsPerPixelX)
        ImgSizeAll.Top = (ImgSizeAll.Top - (OldY - Y) \ Screen.TwipsPerPixelY)
        'Resize Object
        oWidth = (ImgSizeAll.Left - CtrlObj.Left)
        oHeight = (ImgSizeAll.Top - CtrlObj.Top)
        
        If (oWidth < 16) Then oWidth = 16
        If (oHeight < 16) Then oHeight = 16
        
        CtrlObj.Width = oWidth
        CtrlObj.Height = oHeight
         
        ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
    End If
End Sub

Private Sub ImgSizeAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Resizer position
    ImgSizeAll.Move (CtrlObj.Left + CtrlObj.Width), (CtrlObj.Top + CtrlObj.Height), 7, 7
    IsDown = False
End Sub

Private Sub lblA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(lblA(Index), 0, Button, X, Y)
End Sub

Private Sub lblA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(lblA(Index), 1, Button, X, Y)
End Sub

Private Sub lblA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DesignerAction(lblA(Index), 2, Button, X, Y)
End Sub

Private Sub mnuAbout_Click()
    MsgBox "DM CD Menu Creator Version 1.0" & vbCrLf & vbTab & "By DreamVB", vbInformation, "About"
End Sub

Private Sub mnuBack_Click()
    Call MdiMain.CmdButton_Click(7)
End Sub

Private Sub mnuExit_Click()
    Unload MdiMain
End Sub

Private Sub mnuFront_Click()
    Call MdiMain.CmdButton_Click(6)
End Sub

Private Sub mnuProp_Click()
    Call MdiMain.CmdButton_Click(3)
End Sub
