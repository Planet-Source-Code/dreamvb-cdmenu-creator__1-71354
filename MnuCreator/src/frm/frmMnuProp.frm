VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMnuProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdOpen 
      Height          =   345
      Left            =   4515
      TabIndex        =   2
      Top             =   1740
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
   Begin VB.TextBox txtSnd 
      Height          =   350
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1740
      Width           =   4170
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4695
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.dFlatButton cmdOpen1 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Choose Color"
      Top             =   2490
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "..."
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   5250
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2985
      Width           =   5250
      Begin Project1.dFlatButton cmdCancel 
         Height          =   345
         Left            =   4365
         TabIndex        =   5
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton cmdOk 
         Height          =   345
         Left            =   3450
         TabIndex        =   4
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
      End
   End
   Begin VB.PictureBox pBackColor 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   300
      ScaleHeight     =   270
      ScaleWidth      =   1590
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1650
   End
   Begin VB.TextBox txtCaption 
      Height          =   350
      Left            =   300
      TabIndex        =   0
      Top             =   1035
      Width           =   4710
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5250
      TabIndex        =   6
      Top             =   0
      Width           =   5250
      Begin Project1.Line3D Line3D1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   540
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   53
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set your menu properties below:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2370
      End
      Begin VB.Image ImgTop 
         Height          =   480
         Left            =   75
         Picture         =   "frmMnuProp.frx":0000
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Background Sound:"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   1485
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Menu Back Color:"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   795
      Width           =   585
   End
End
Attribute VB_Name = "frmMnuProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmMnuProp
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
Dim lFile As String

    'Set Menu Prop type
    If Len(txtSnd.Text) > 0 Then
        lFile = GetFilename(txtSnd.Text)
        FileCopy txtSnd.Text, ProjResPath & lFile
        
        lFile = ProjResPath & lFile
        'Store the mids filename in the forms tag property
        CtrlObj.Tag = Replace(lFile, PathFromFile(MyIni.FileName), "$AppPath")
    End If
    
    CtrlObj.Caption = txtCaption.Text
    CtrlObj.BackColor = pBackColor.BackColor
    Unload frmMnuProp
End Sub

Private Sub cmdOpen_Click()
On Error GoTo CanErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Wav Files(*.wav)|*.wav|Midi Files(*.mid)|*.mid|"
        .ShowOpen
        'Return Filename
        txtSnd.Text = .FileName
    End With
    
    Exit Sub
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Sub

Private Sub cmdOpen1_Click()
On Error GoTo CancelErr:
    
    CD1.CancelError = True
    CD1.ShowColor
    pBackColor.BackColor = CD1.Color
    
    Exit Sub
CancelErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
    Set frmMnuProp.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMnuProp = Nothing
End Sub

Private Sub pBottom_Resize()
    Line3D2.Width = pBottom.ScaleWidth
End Sub

Private Sub pTop_Resize()
    Line3D1.Width = frmMnuProp.ScaleWidth
End Sub
