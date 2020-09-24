VERSION 5.00
Begin VB.Form Sats 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox SSVCSat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      TabIndex        =   10
      Text            =   "30"
      Top             =   1110
      Width           =   915
   End
   Begin VB.TextBox ISVCSat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      TabIndex        =   9
      Text            =   "34"
      Top             =   1470
      Width           =   915
   End
   Begin VB.TextBox RASat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      TabIndex        =   8
      Text            =   "26"
      Top             =   1800
      Width           =   915
   End
   Begin VB.TextBox SinusSat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      TabIndex        =   7
      Text            =   "15"
      Top             =   2130
      Width           =   915
   End
   Begin VB.TextBox RVSat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      TabIndex        =   6
      Text            =   "35"
      Top             =   2475
      Width           =   915
   End
   Begin VB.TextBox PASat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3705
      TabIndex        =   5
      Text            =   "32"
      Top             =   2820
      Width           =   915
   End
   Begin VB.TextBox PCSat 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "120"
      Top             =   3150
      Width           =   915
   End
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   5880
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label72 
      BackStyle       =   0  'Transparent
      Caption         =   "Superior SVC SO2%="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   17
      Top             =   1155
      Width           =   2010
   End
   Begin VB.Label Label73 
      BackStyle       =   0  'Transparent
      Caption         =   "Inferior SVC SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   16
      Top             =   1500
      Width           =   2010
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Right Atrial SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   15
      Top             =   1830
      Width           =   2010
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "Coronary Sinus SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   14
      Top             =   2175
      Width           =   2010
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "Right Vent SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   13
      Top             =   2505
      Width           =   2010
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulm Artery SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   12
      Top             =   2850
      Width           =   2010
   End
   Begin VB.Label Label85 
      BackStyle       =   0  'Transparent
      Caption         =   "PC Wedge SO2%= "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   11
      Top             =   3180
      Width           =   2010
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Sided O2 Saturations"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Sided O2 Saturations"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   2
      Top             =   210
      Width           =   4455
   End
   Begin VB.Label Label107 
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label107 
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   1
      Left            =   5970
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "Sats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
End Sub

Private Sub FlatButton5_Click()
sSVCSO2 = Val(SSVCSat.Text)
iSVCSO2 = Val(ISVCSat.Text)
RASO2 = Val(RASat.Text)
RVSO2 = Val(RVSat.Text)
PASO2 = Val(PASat.Text)
CSSO2 = Val(SinusSat.Text)
PCSO2 = Val(PCSat.Text)

Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
sSVCSO2 = 30

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub



Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label107_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label107(0).Visible = False
        Label107(1).Visible = True
End Sub

Private Sub Label107_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label107(1).Visible = False
        Label107(0).Visible = True
        Unload Me
End Sub


Private Sub Label11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
