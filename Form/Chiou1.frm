VERSION 5.00
Begin VB.Form Chiou1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   120
      Top             =   240
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox InfusVoltxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   22
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox DesiredTheophtxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Time2Txt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   20
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Time1Txt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox C2txt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox ChiouMaintxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox C1txt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox ChiouMgHrtxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox MgDay1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
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
      TabIndex        =   24
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chiou Kinetic Dosing Method"
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
      Height          =   345
      Index           =   0
      Left            =   840
      TabIndex        =   11
      Top             =   45
      Width           =   4725
   End
   Begin VB.Label OsmGaptxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Infusion Volume (ml):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   10
      Top             =   3720
      Width           =   2160
   End
   Begin VB.Label CalcSosm1txt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desired Theophylline Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   9
      Top             =   3360
      Width           =   2940
   End
   Begin VB.Label AGGtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3285
      TabIndex        =   8
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label AGtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   645
      TabIndex        =   7
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label FENAtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conc 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label RFItxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maint Dose (mg/kg/hr) ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   5
      Top             =   1560
      Width           =   2490
   End
   Begin VB.Label ActualCrCltxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conc 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   645
      TabIndex        =   4
      Top             =   2145
      Width           =   780
   End
   Begin VB.Label EstimatedCrCltxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Milligrams per Hour ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   3
      Top             =   1200
      Width           =   2220
   End
   Begin VB.Label BUN_Crtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Milligrams per day ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   2
      Top             =   840
      Width           =   2115
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
      Left            =   6015
      TabIndex        =   0
      Top             =   135
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
      Left            =   5985
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chiou Kinetic Dosing Method"
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
      Height          =   345
      Index           =   1
      Left            =   795
      TabIndex        =   12
      Top             =   60
      Width           =   4725
   End
End
Attribute VB_Name = "Chiou1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
 Unload Interpretive
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub ChiouMaintxt_GotFocus()
    On Error Resume Next
If Trim(ChiouMaintxt.Text) = "" Then _
    ChiouMaintxt.Text = Format((Val(ChiouMgHrtxt.Text) / Wtkg), "###.00")
If Trim(ChiouMaintxt.Text) = "" Then _
    ChiouMaintxt.Text = Format((Val(MgDay1.Text) / 24 / Wtkg), "###.00")

End Sub

Private Sub ChiouMgHrtxt_GotFocus()
    On Error Resume Next
If Trim(ChiouMgHrtxt.Text) = "" Then _
    ChiouMgHrtxt.Text = Format(Val(MgDay1.Text) / 24, "####.0")
If Trim(ChiouMgHrtxt.Text) = "" Then _
    ChiouMgHrtxt.Text = Format(Val(ChiouMaintxt.Text) * Wtkg, "####.0")
Exit Sub


End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub FlatButton5_Click()
MgDay = Val(MgDay1.Text)
RO = Val(ChiouMaintxt.Text)
C1 = Val(C1txt.Text)
C2 = Val(C2txt.Text)
TimeA = Val(Time1Txt.Text)
TimeB = Val(Time2Txt.Text)
C3 = Val(DesiredTheophtxt.Text)
VI1 = Val(InfusVoltxt.Text)
MGHR = Val(ChiouMgHrtxt.Text)
If C1 = 0 Or C2 = 0 Or TimeA = 0 Or TimeB = 0 Or C3 = 0 Or VI1 = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Complete all Fields Please."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
Else
        Load Chiou2
        Chiou2.Show
End If

End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
   
If Val(MgDay1.Text) <> 0 Then
    MgDay1.Text = Format(MgDay, "###.0")
    ChiouMgHrtxt.Text = Format(MgDay / 24, "###.0")
    ChiouMaintxt.Text = Format(MgDay / (24 * ABW), "###.0")
End If
If MGHR <> 0 Then
    ChiouMgHrtxt.Text = Format(MGHR, "###.0")
    MgDay1.Text = Format(MGHR * 24, "###.0")
    ChiouMaintxt.Text = Format(MGHR / ABW, "###.0")
End If
If RO <> 0 Then
    ChiouMaintxt.Text = Format(RO, "###.0")
    ChiouMgHrtxt.Text = Format(RO * ABW, "###.0")
    MgDay1.Text = Format(RO * ABW * 24, "###.0")
End If
If C1 <> 0 Then
    C1txt.Text = Format(C1, "###.0")
End If
If C2 <> 0 Then
    C2txt.Text = Format(C2, "###.0")
End If
If TimeA <> 0 Then
    Time1Txt.Text = Format(TimeA, "###.0")
End If
If TimeB <> 0 Then
    Time2Txt.Text = Format(TimeB, "###.0")
End If
If C3 <> 0 Then
    DesiredTheophtxt.Text = Format(C3, "###.0")
End If
If VI1 <> 0 Then
    InfusVoltxt.Text = Format(VI1, "###.0")
End If
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

Private Sub MgDay_GotFocus()
    On Error Resume Next
If Trim(MgDay1.Text) = "" Then _
    MgDay1.Text = Format(Val(ChiouMgHrtxt.Text) * 24, "####.0")
If Trim(MgDay1.Text) = "" Then _
    MgDay1.Text = Format(Val(ChiouMaintxt.Text) * 24 * Wtkg, "####.0")

End Sub

Private Sub Label11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
