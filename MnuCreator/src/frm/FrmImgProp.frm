VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImageProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCmd2 
      Height          =   350
      Left            =   165
      TabIndex        =   4
      Top             =   1830
      Width           =   4050
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      Left            =   165
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1170
      Width           =   1560
   End
   Begin VB.CheckBox ChkStretch 
      Caption         =   "Allow Image Stretching"
      Height          =   210
      Left            =   1845
      TabIndex        =   3
      Top             =   1215
      Width           =   3285
   End
   Begin VB.TextBox txtCmd 
      Height          =   350
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   4050
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4830
      Top             =   195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   5430
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2325
      Width           =   5430
      Begin Project1.dFlatButton cmdCancel 
         Height          =   345
         Left            =   4500
         TabIndex        =   7
         Top             =   105
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cancel"
      End
      Begin Project1.Line3D Line3D2 
         Height          =   30
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton cmdOk 
         Height          =   345
         Left            =   3585
         TabIndex        =   6
         Top             =   105
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OK"
         Enabled         =   0   'False
      End
   End
   Begin Project1.dFlatButton cmdOpen1 
      Height          =   345
      Left            =   4275
      TabIndex        =   1
      ToolTipText     =   "Open"
      Top             =   495
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ". . . ."
   End
   Begin Project1.dFlatButton cmdOpen2 
      Height          =   345
      Left            =   4275
      TabIndex        =   5
      ToolTipText     =   "Open"
      Top             =   1830
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ". . . ."
   End
   Begin VB.Label lblCmd 
      AutoSize        =   -1  'True
      Caption         =   "Command:"
      Height          =   195
      Left            =   165
      TabIndex        =   12
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      Caption         =   "Action:"
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   930
      Width           =   495
   End
   Begin VB.Label lblImgFile 
      AutoSize        =   -1  'True
      Caption         =   "Image Path:"
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   255
      Width           =   855
   End
End
Attribute VB_Name = "frmImageProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAction_Click()
    If (cboAction.ListIndex = 0) Then
        cmdOpen2.Enabled = True
        txtCmd2.Enabled = True
    End If
    If (cboAction.ListIndex = 1) Then
        cmdOpen2.Enabled = False
        txtCmd2.Enabled = False
    End If
    If (cboAction.ListIndex = 2) Then
        cmdOpen2.Enabled = False
        txtCmd2.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmImageProp
End Sub

Private Sub cmdOk_Click()
Dim lFile As String
On Error Resume Next
    'Get the Filename
    lFile = GetFilename(txtCmd.Text)
    'Copy the file over the the Project RES Folder
    FileCopy txtCmd.Text, ProjResPath & lFile
    lFile = ProjResPath & lFile
    'Load the image from the RES Path
    CtrlObj.Picture = LoadPicture(lFile)
    'Set the tag, Zrder, Action, Picture Filename
    
    txtCmd2.Text = Replace(txtCmd2.Text, PathFromFile(MyIni.FileName), "$AppPath")
    
    CtrlObj.Tag = Left(CtrlObj.Tag, 1) & cboAction.ListIndex & Replace(lFile, PathFromFile(MyIni.FileName), "$AppPath") & Chr(0) & txtCmd2.Text
    CtrlObj.Stretch = ChkStretch.Value
    CtrlObj.BorderStyle = 0
    Unload frmImageProp
End Sub

Private Sub cmdOpen1_Click()
On Error GoTo OpenErr:
Dim Tmp As String

    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Bitmap Files(*.bmp)|*.bmp|GIF Files(*.gif)|*.gif|Jpeg Files(*.jpg)|*.jpg|Icon Files(*.ico)|*.ico|"
        .FilterIndex = 2
        .ShowOpen
        txtCmd.Text = .FileName
        .FileName = vbNullString
    End With
    
    Exit Sub
OpenErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If

End Sub

Private Sub cmdOpen2_Click()
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "All Files(*.*)|*.*|"
        .ShowOpen
        txtCmd2.Text = .FileName
        .FileName = vbNullString
    End With
    
    Exit Sub
OpenErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If

End Sub

Private Sub Form_Load()
    Set frmImageProp.Icon = Nothing
    '
    cboAction.AddItem "Execute File"
    cboAction.AddItem "Exit"
    cboAction.AddItem "MessageBox"
    cboAction.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImageProp = Nothing
End Sub

Private Sub pBottom_Resize()
    Line3D2.Width = pBottom.ScaleWidth
End Sub

Private Sub txtCmd_Change()
    cmdOk.Enabled = True
End Sub
