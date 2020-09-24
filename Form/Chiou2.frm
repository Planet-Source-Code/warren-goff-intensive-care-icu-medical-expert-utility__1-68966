VERSION 5.00
Begin VB.Form Chiou2 
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
   Begin VB.TextBox DoseFactstxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
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
      Height          =   1095
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3000
      Width           =   5655
   End
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
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox NMaint 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox TBolus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox THalfLife 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Csstxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox CalClear 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
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
      TabIndex        =   16
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chiou Theophylline Results"
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
      Left            =   975
      TabIndex        =   7
      Top             =   45
      Width           =   4455
   End
   Begin VB.Label OsmGaptxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Maint Dose mg/hr="
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
      TabIndex        =   6
      Top             =   2280
      Width           =   2460
   End
   Begin VB.Label CalcSosm1txt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bolus Patient with mg ="
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
      TabIndex        =   5
      Top             =   1920
      Width           =   2385
   End
   Begin VB.Label RFItxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Theophylline Half Life (hr) ="
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
      Top             =   1560
      Width           =   2865
   End
   Begin VB.Label EstimatedCrCltxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desired Conc: Css (mg/L) = "
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label BUN_Crtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Theophylline Clearance L/hr ="
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
      TabIndex        =   2
      Top             =   840
      Width           =   3120
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
      Caption         =   "Chiou Theophylline Results"
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
      Left            =   930
      TabIndex        =   8
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "Chiou2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Initialize()
Dim Css As Single 'Desired Concentration


'x = (2 * MGHR * .8)
'y = (C1 + C2)
'w = ABW * (C1 - C2)
'z = (C1 + C2) * (TimeB - TimeA)
'CL1 = (x/y) + (w/z)
CL1 = ((2 * MGHR * 0.8) / (C1 + C2)) + (ABW * (C1 - C2)) / ((C1 + C2) * (TimeB - TimeA)) 'L/h


Css = MGHR * 0.8 / CL1 'mg/l
r2 = MGHR * C3 / Css  'mg/hr
'MsgBox CL1
L3 = 0.5 * ABW * (C3 - C2)
If CL1 > 0 Then
    Csstxt.Text = Format(Css, "####.0")
    CalClear.Text = Format(CL1, "####.0")
    THalfLife.Text = Format((0.693 * 0.5 * ABW) / CL1, "####.0") 'Format(500 * .693 / CL1, "####.0")
    If L3 > 0 Then
        TBolus.Text = Format(L3, "######.0")
    End If
    NMaint.Text = Format(r2, "######.0")
    DoseFactstxt.Text = "The Dosage is " + Format(r2 * 24, "#####.0") + " mg/day in " + Format(VI1, "######") + "cc at " + Format(VI1 / 24, "####.0") + "cc/hr. " & vbCrLf & vbCrLf + "If 1000 mg were placed in 250 cc it would run at " + Format(r2 / 4, "####.0") + " cc/hr"
Else
    CalClear.Text = " 0"
    DoseFactstxt.Text = "Since there is NO APPARENT Clearance of Theophylline, HOLD the infusion." & vbCrLf & vbCrLf + "A Lab Error must be excluded first." & vbCrLf & vbCrLf + "Recheck a set of 2 Theoph levels within 24 hours, 4 hours apart and recalculate kinetics if no lab error exists." & vbCrLf & vbCrLf + "If Theophylline Toxicity exists, consider Oral Activated Charcoal or Charcoal Hemoperfusion."
End If
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True


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

