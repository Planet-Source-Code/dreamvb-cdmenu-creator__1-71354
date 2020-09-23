VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "DM CDMenu Creator"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "MdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   75
      Top             =   615
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   571
      TabIndex        =   0
      Top             =   0
      Width           =   8565
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   8
         Left            =   2670
         TabIndex        =   11
         ToolTipText     =   "Image"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":0CCA
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   7
         Left            =   3555
         TabIndex        =   10
         ToolTipText     =   "Send to back"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":101C
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   6
         Left            =   3120
         TabIndex        =   9
         ToolTipText     =   "Bring to front"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":136E
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   5
         Left            =   2220
         TabIndex        =   8
         ToolTipText     =   "Label"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":16C0
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   4
         Left            =   1785
         TabIndex        =   7
         ToolTipText     =   "Button"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":1A12
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Tag             =   "1,Test"
         ToolTipText     =   "Properties"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":1D64
      End
      Begin Project1.Line3D Line3D2 
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   465
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   0
         Left            =   15
         TabIndex        =   2
         ToolTipText     =   "New"
         Top             =   45
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":20B6
      End
      Begin Project1.Line3D Line3D1 
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   1
         Left            =   435
         TabIndex        =   4
         ToolTipText     =   "Open"
         Top             =   45
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":21C8
      End
      Begin Project1.dFlatButton CmdButton 
         Height          =   390
         Index           =   2
         Left            =   855
         TabIndex        =   5
         ToolTipText     =   "Save"
         Top             =   45
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "MdiMain.frx":251A
      End
   End
End
Attribute VB_Name = "MdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FirstRun As Boolean

Private Sub ShowForm()

    If (Not frmMain.Visible) Then
        frmMain.Visible = True
    End If
    
    CmdButton(2).Visible = True
    CmdButton(3).Visible = True
    CmdButton(4).Visible = True
    CmdButton(5).Visible = True
    CmdButton(6).Visible = True
    CmdButton(7).Visible = True
    CmdButton(8).Visible = True
End Sub

Private Function GetDLGName(Title As String, Optional vInitDir As String, Optional vShowSave As Boolean = False) As String
On Error GoTo CanErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = "INI Files(*.*)|*.ini|"
        .InitDir = vInitDir
        
        If (vShowSave) Then
            Call .ShowSave
        Else
            Call .ShowOpen
        End If
        'Return Filename
        GetDLGName = .FileName
    End With
    
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Public Sub CmdButton_Click(Index As Integer)
Dim cName As String
Dim BinFolder As String
Dim Tmp As String
Dim StrA As String
Dim sPos As Integer

    Select Case Index
        Case 0
            'New Project
            frmNew.Show vbModal, MdiMain
            'Check if to start a new project
            If (ButtonPress = vbOK) Then
                'Check for project path
                If Not FindFolder(ProjPath) Then
                    'Create main project path.
                    Call MkDir(ProjPath)
                End If
                'Check for Resources path
                If Not FindFolder(ProjResPath) Then
                    'Create Project Resources path.
                    Call MkDir(ProjResPath)
                End If
                'Copy menu resources over to Project dest folder
                BinFolder = FixPath(App.Path) & "bin\"
                FileCopy BinFolder & "Autorun.inf", ProjPath & "Autorun.inf"
                FileCopy BinFolder & "icon.ico", ProjPath & "RES\icon.ico"
                FileCopy BinFolder & "menu.exe", ProjPath & "menu.exe"
                FileCopy BinFolder & "menu.ini", ProjPath & "menu.ini"
                'Load project
                Call frmMain.LoadProject(ProjPath & "menu.ini")
                'Show main designer window.
                Call ShowForm
            End If
        Case 1
            'Open Project.
            cName = GetDLGName("Open", ProjPath)
            'Check for project name.
            If Len(cName) > 0 Then
                'Store project path
                ProjPath = FixPath(PathFromFile(cName))
                ProjResPath = ProjPath & "RES\"
                'Open project file
                Call frmMain.LoadProject(cName)
                'Show main designer form.
                Call ShowForm
            End If
        Case 2
            'Save Project
            cName = GetDLGName("Save As", ProjPath, True)
            'Check for project name.
            If Len(cName) > 0 Then
                'Save the project
                Call frmMain.SaveProject(cName)
            End If
        Case 3
            'Get Control type
            cName = UCase(TypeName(CtrlObj))
            'Check what object we are dealing with
            Select Case cName
                Case "FRMMAIN"
                    'Form Object.
                    frmMnuProp.txtCaption.Text = frmMain.Caption
                    frmMnuProp.pBackColor.BackColor = frmMain.BackColor
                    frmMnuProp.txtSnd.Text = CtrlObj.Tag
                    frmMnuProp.Show vbModal, MdiMain
                Case "IMAGE"
                    'Image control
                    Tmp = CtrlObj.Tag
                    sPos = InStrRev(Tmp, Chr(0), Len(Tmp), vbBinaryCompare)
                    '
                    If (sPos > 0) Then
                        StrA = Mid$(Tmp, sPos + 1)
                        Tmp = Left$(Tmp, sPos - 1)
                    End If
   
                    frmImageProp.txtCmd.Text = Mid$(Tmp, 3)
                    frmImageProp.txtCmd2.Text = StrA
                    frmImageProp.ChkStretch.Value = Abs(CtrlObj.Stretch)
                    frmImageProp.cboAction.ListIndex = Val(Mid(Tmp, 2, 1))
                    frmImageProp.Show vbModal, MdiMain
                Case "COMMANDBUTTON", "LABEL"
                    'Get the objects font properties
                    frmButtonProp.CD1.FontName = CtrlObj.FontName
                    frmButtonProp.CD1.FontSize = CtrlObj.FontSize
                    frmButtonProp.CD1.FontBold = CtrlObj.FontBold
                    frmButtonProp.CD1.FontItalic = CtrlObj.FontItalic
                    frmButtonProp.CD1.FontUnderline = CtrlObj.FontUnderline
                    frmButtonProp.CD1.FontStrikethru = CtrlObj.FontStrikethru
                    'Control Object.
                    frmButtonProp.txtCaption.Text = CtrlObj.Caption
                    If (cName = "LABEL") Then
                        frmButtonProp.pBackColor.BackColor = CtrlObj.ForeColor
                    Else
                        frmButtonProp.pBackColor.BackColor = CtrlObj.BackColor
                    End If
                    'Extract the action value
                    frmButtonProp.cboAction.ListIndex = Val(Mid(CtrlObj.Tag, 2, 1))
                    'Extract the action command
                    frmButtonProp.txtCmd.Text = Mid$(CtrlObj.Tag, 3)
                    frmButtonProp.Show vbModal, MdiMain
                End Select
        Case 4
            'Add button
            Call frmMain.AddControl("BUTTON")
        Case 5
            'Add Label
            Call frmMain.AddControl("LABEL")
        Case 6
            'Send to back
            If UCase(TypeName(CtrlObj)) <> "FRMMAIN" Then
                Tmp = CtrlObj.Tag
                Mid(Tmp, 1, 1) = "0"
                CtrlObj.Tag = Tmp
                CtrlObj.ZOrder vbBringToFront
            End If
        Case 7
            'Bring to front
            If UCase(TypeName(CtrlObj)) <> "FRMMAIN" Then
                Tmp = CtrlObj.Tag
                Mid(Tmp, 1, 1) = "1"
                CtrlObj.Tag = Tmp
                CtrlObj.ZOrder vbSendToBack
            End If
        Case 8
            'Add Image
            Call frmMain.AddControl("IMAGE")
    End Select
End Sub

Private Sub MDIForm_Activate()
    If (FirstRun) Then
        frmMain.Visible = False
        FirstRun = False
    End If
End Sub

Private Sub MDIForm_Load()
   FirstRun = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set frmButtonProp = Nothing
    Set frmImageProp = Nothing
    Set frmMnuProp = Nothing
    Set frmMain = Nothing
    Set frmNew = Nothing
    Set MdiMain = Nothing
End Sub

Private Sub pTop_Resize()
    Line3D1.Width = pTop.ScaleWidth
    Line3D2.Width = pTop.ScaleWidth
End Sub
