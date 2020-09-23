VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmButtonProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdFont 
      Height          =   345
      Left            =   3825
      TabIndex        =   18
      Top             =   1785
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Select Font"
      Picture         =   "frmPropA.frx":0000
   End
   Begin VB.TextBox txtCmd 
      Height          =   350
      Left            =   315
      TabIndex        =   3
      Top             =   2535
      Width           =   4050
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      Left            =   315
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   1560
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4695
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.dFlatButton cmdOpen1 
      Height          =   285
      Left            =   3330
      TabIndex        =   2
      ToolTipText     =   "Choose Color"
      Top             =   1815
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
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
      ScaleWidth      =   5430
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3135
      Width           =   5430
      Begin Project1.dFlatButton cmdCancel 
         Height          =   345
         Left            =   4500
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
         Caption         =   "Cancel"
      End
      Begin Project1.Line3D Line3D2 
         Height          =   30
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton cmdOk 
         Height          =   345
         Left            =   3585
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
         Caption         =   "OK"
      End
   End
   Begin VB.PictureBox pBackColor 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2070
      ScaleHeight     =   255
      ScaleWidth      =   1590
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1800
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
      ScaleWidth      =   5430
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   5430
      Begin Project1.Line3D Line3D1 
         Height          =   30
         Left            =   0
         TabIndex        =   9
         Top             =   540
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   53
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set your controls properties below:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   2565
      End
      Begin VB.Image ImgTop 
         Height          =   480
         Left            =   75
         Picture         =   "frmPropA.frx":0352
         Top             =   30
         Width           =   480
      End
   End
   Begin Project1.dFlatButton cmdOpen2 
      Height          =   345
      Left            =   4410
      TabIndex        =   4
      ToolTipText     =   "Open"
      Top             =   2535
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Font:"
      Height          =   195
      Left            =   3900
      TabIndex        =   17
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Command"
      Height          =   195
      Left            =   315
      TabIndex        =   16
      Top             =   2295
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Action"
      Height          =   195
      Left            =   315
      TabIndex        =   15
      Top             =   1530
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Text Color:"
      Height          =   195
      Left            =   2055
      TabIndex        =   12
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   795
      Width           =   585
   End
End
Attribute VB_Name = "frmButtonProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAction_Click()
    If (cboAction.ListIndex = 0) Then
        cmdOpen2.Enabled = True
        txtCmd.Enabled = True
    End If
    If (cboAction.ListIndex = 1) Then
        cmdOpen2.Enabled = False
        txtCmd.Enabled = False
    End If
    If (cboAction.ListIndex = 2) Then
        cmdOpen2.Enabled = False
        txtCmd.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmButtonProp
End Sub

Private Sub cmdFont_Click()
On Error GoTo CanErr:
    
    With CD1
        .CancelError = True
        .Flags = cdlCFBoth
        .ShowFont
    End With
    
    Exit Sub
CanErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If
    
End Sub

Private Sub cmdOk_Click()
Dim cName As String

    'Set Font Properties
    CtrlObj.FontName = CD1.FontName
    CtrlObj.FontSize = CD1.FontSize
    CtrlObj.FontBold = CD1.FontBold
    CtrlObj.FontItalic = CD1.FontItalic
    CtrlObj.FontUnderline = CD1.FontUnderline
    CtrlObj.FontStrikethru = CD1.FontStrikethru

    'Set Button properties
    CtrlObj.Caption = txtCaption.Text
    
    If UCase(TypeName(CtrlObj)) = "LABEL" Then
        'Set forecolor for label
        CtrlObj.ForeColor = pBackColor.BackColor
    Else
        'Set back color for command buttons
        CtrlObj.BackColor = pBackColor.BackColor
    End If
    '
    txtCmd.Text = Replace(txtCmd.Text, PathFromFile(MyIni.FileName), "$AppPath")
    CtrlObj.Tag = Left(CtrlObj.Tag, 1) & cboAction.ListIndex & txtCmd.Text
    Unload frmButtonProp
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

Private Sub cmdOpen2_Click()
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "All Files(*.*)|*.*|"
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

Private Sub Form_Load()
    Set frmButtonProp.Icon = Nothing
    
    cboAction.AddItem "Execute File"
    cboAction.AddItem "Exit"
    cboAction.AddItem "MessageBox"
    cboAction.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmButtonProp = Nothing
End Sub

Private Sub pBottom_Resize()
    Line3D2.Width = pBottom.ScaleWidth
End Sub

Private Sub pTop_Resize()
    Line3D1.Width = frmButtonProp.ScaleWidth
End Sub
