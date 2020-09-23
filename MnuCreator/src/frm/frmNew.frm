VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Project"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdOpen 
      Height          =   345
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "Choose Folder"
      Top             =   1170
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
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5175
      TabIndex        =   8
      Top             =   1755
      Width           =   5175
      Begin Project1.dFlatButton cmdCancel 
         Height          =   345
         Left            =   4230
         TabIndex        =   3
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
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   53
      End
      Begin Project1.dFlatButton cmdCreate 
         Height          =   345
         Left            =   3180
         TabIndex        =   2
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         Caption         =   "Create"
         Enabled         =   0   'False
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   345
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1170
      Width           =   4125
   End
   Begin VB.PictureBox pBar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5175
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin VB.Line lnTop 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   750
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a new location for your CDMenu resources"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CD Menu Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   75
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4515
         Picture         =   "frmNew.frx":0000
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label lblSaveProj 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Path"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   915
      Width           =   1080
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmNew
End Sub

Private Sub cmdCreate_Click()
    ProjPath = txtPath.Text
    ProjResPath = ProjPath & "RES\"
    ButtonPress = vbOK
    Unload frmNew
End Sub

Private Sub cmdOpen_Click()
Dim sPath As String
    sPath = FixPath(GetFolder(frmNew.hWnd, "Choose Folder:"))
    
    If (sPath = "\") Then
        Exit Sub
    Else
        txtPath.Text = sPath
    End If
End Sub

Private Sub Form_Load()
    Set frmNew.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNew = Nothing
End Sub

Private Sub pBar_Resize()
    lnTop.X2 = pBar.ScaleWidth
End Sub

Private Sub pBottom_Resize()
    Line3D2.Width = pBottom.ScaleWidth
End Sub

Private Sub txtPath_Change()
    cmdCreate.Enabled = Len(txtPath.Text) > 0
End Sub
