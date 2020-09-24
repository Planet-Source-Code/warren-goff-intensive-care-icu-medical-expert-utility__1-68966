VERSION 5.00
Begin VB.Form Apachefrm 
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
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3015
      TabIndex        =   26
      Top             =   4650
      Width           =   1575
   End
   Begin VB.TextBox ApacheMortalitytxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4080
      TabIndex        =   24
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Apachetxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4080
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
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
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   480
      Width           =   1455
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
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
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF0000&
         Caption         =   "YES"
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
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF0000&
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
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   840
      Width           =   1455
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF0000&
         Caption         =   "YES"
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
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
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
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.ComboBox PDCpop 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Text            =   "Select PDC"
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   300
      Left            =   1695
      TabIndex        =   11
      Top             =   4650
      Width           =   1095
   End
   Begin VB.TextBox PDCtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   3885
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Emergency PostOp Patient"
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
         Height          =   435
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   3840
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Elective PostOp Patient"
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
         Height          =   435
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   3840
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "NonOperative Patient"
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
         Height          =   435
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   3840
      End
   End
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   15
      Left            =   0
      Top             =   3720
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Mortality Estimate ="
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
      Index           =   6
      Left            =   1920
      TabIndex        =   25
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Apache II Score ="
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
      Index           =   5
      Left            =   1920
      TabIndex        =   23
      Top             =   3840
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "PDC="
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
      Index           =   4
      Left            =   4320
      TabIndex        =   21
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Acute Renal Failure ?"
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
      Index           =   3
      Left            =   1560
      TabIndex        =   17
      Top             =   600
      Width           =   2235
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
      TabIndex        =   12
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Severe Organ System Insuff ?"
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
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "PDC Leading to ICU Admit:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apache II Mortality Score"
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
      Left            =   1095
      TabIndex        =   3
      Top             =   120
      Width           =   4185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apache II Mortality Score"
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
      Left            =   1065
      TabIndex        =   2
      Top             =   90
      Width           =   4185
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
Attribute VB_Name = "Apachefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CandyButton1_Click()

End Sub

Private Sub Command1_Click()
Dim T1 As Single
Dim Z7 As Single
Z7 = 0
T1 = 0
If Sex = 0 Or BSA = 0 Or Wtkg = 0 Or ABW = 0 Or Htcm = 0 Or Age = 0 Or Sex = 0 Or MAP = 0 Or AR = 0 Or RR = 0 Or Temp = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Please Enter all Demographic Information."
    SetTopMostWindow Me.hwnd, True
    Load DataInput
    DataInput.Show
    Exit Sub
End If
If pH = 0 Or PaCO2 = 0 Or PaO2 = 0 Or Patm = 0 Or SaO2 = 0 Or FiO2 = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Input pH, PaCO2, PaO2, Patm, SaO2 and FiO2"
    SetTopMostWindow Me.hwnd, True
    Load DataInput
    DataInput.Show
    Exit Sub
End If
If WBC = 0 Or Hematocrit = 0 Then     'or Hb = 0 or Hct = 0
    SetTopMostWindow Me.hwnd, False
    MsgBox "Input WBC and Hematocrit Please."
    SetTopMostWindow Me.hwnd, True
    Load DataInput
    DataInput.Show
    Exit Sub
End If
If Creatinine = 0 Or Sodium = 0 Or Potassium = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Input Serum Creatinine, Sodium and Potassium Please."
    SetTopMostWindow Me.hwnd, True
    Load DataInput
    DataInput.Show
    Exit Sub
End If
If Glasgow = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Input all Data for Glasgow Coma Score Please."
    SetTopMostWindow Me.hwnd, True
    Load GlasgowFrm
    GlasgowFrm.Show
    Exit Sub
End If

If ARF = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Please Select Renal Failure Status."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Please Select Surgical Status."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If PDC = 0 Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Please Select PDC."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

Apache2 = 0

T1 = ((Temp - 32) * 5 / 9)
     If T1 >= 41 Or T1 <= 29.9 Then
        Apache2 = Apache2 + 4
        GoTo ONE48
     End If
     If T1 >= 39 Or T1 <= 31.9 Then
       Apache2 = Apache2 + 3
       GoTo ONE48
     End If
     If T1 <= 33.9 Then
        Apache2 = Apache2 + 2
        GoTo ONE48
     End If
     If T1 >= 38.5 Or T1 <= 35.9 Then
        Apache2 = Apache2 + 1
     End If

ONE48:
     If MAP >= 160 Or MAP <= 49 Then
        Apache2 = Apache2 + 4
        GoTo ONE51
     End If
     If MAP >= 130 Then
        Apache2 = Apache2 + 3
        GoTo ONE51
     End If
     If MAP >= 110 Or MAP <= 69 Then
          Apache2 = Apache2 + 2
     End If

ONE51:
     If AR >= 180 Or AR <= 39 Then
        Apache2 = Apache2 + 4
        GoTo ONE54
     End If
     If AR >= 140 Or AR <= 54 Then
         Apache2 = Apache2 + 3
         GoTo ONE54
     End If
     If AR >= 110 Or AR <= 69 Then
        Apache2 = Apache2 + 2
     End If

ONE54:
     If RR >= 50 Or RR <= 5 Then
        Apache2 = Apache2 + 4
        GoTo ONE58
     End If
     If RR >= 35 Then
        Apache2 = Apache2 + 3
        GoTo ONE58
     End If
     If RR <= 9 Then
        Apache2 = Apache2 + 2
        GoTo ONE58
     End If
     If RR >= 25 Or RR <= 11 Then
         Apache2 = Apache2 + 1
     End If

ONE58:
     If FiO2 >= 50 And Aadif >= 500 Then
         Apache2 = Apache2 + 4
         GoTo ONE64
     End If
     If FiO2 >= 50 And Aadif >= 350 Then
        Apache2 = Apache2 + 3
        GoTo ONE64
     End If
     If FiO2 >= 50 And Aadif >= 200 Then
        Apache2 = Apache2 + 2
        GoTo ONE64
     End If
     If FiO2 < 50 And PaO2 < 55 Then
        Apache2 = Apache2 + 4
        GoTo ONE64
     End If
     If FiO2 < 50 And PaO2 <= 60 Then
        Apache2 = Apache2 + 3
        GoTo ONE64
     End If
     If FiO2 < 50 And PaO2 <= 70 Then
        Apache2 = Apache2 + 1
     End If

ONE64:
     If pH = 0 Then
      GoTo ONE90
     End If
     If pH >= 7.7 Or pH < 7.15 Then
         Apache2 = Apache2 + 4
         GoTo ONE69
     End If
     If pH >= 7.6 Or pH <= 7.24 Then
        Apache2 = Apache2 + 3
        GoTo ONE69
     End If
     If pH <= 7.32 Then
        Apache2 = Apache2 + 2
        GoTo ONE69
     End If
     If pH >= 7.5 Then
        Apache2 = Apache2 + 1
     End If

ONE69:
    If Sodium >= 180 Or Sodium <= 110 Then
       Apache2 = Apache2 + 4
       GoTo ONE73
    End If
    If Sodium >= 160 Or Sodium <= 119 Then
       Apache2 = Apache2 + 3
       GoTo ONE73
    End If
    If Sodium >= 155 Or Sodium <= 129 Then
       Apache2 = Apache2 + 2
       GoTo ONE73
    End If
    If Sodium >= 150 Then
       Apache2 = Apache2 + 1
    End If

ONE73:
    If Potassium >= 7 Or Potassium < 2.5 Then
        Apache2 = Apache2 + 4
        GoTo ONE77
    End If
    If Potassium >= 6 Then
       Apache2 = Apache2 + 3
       GoTo ONE77
    End If
    If Potassium <= 2.9 Then
       Apache2 = Apache2 + 2
       GoTo ONE77
    End If
    If Potassium >= 5.5 Or Potassium <= 3.4 Then
       Apache2 = Apache2 + 1
    End If

ONE77:
    If Creatinine >= 3.5 Then
       Apache2 = Apache2 + 4 * ARF
       GoTo ONE80
    End If
    If Creatinine >= 2 Then
       Apache2 = Apache2 + 3 * ARF
       GoTo ONE80
    End If
    If Creatinine >= 1.5 Or Creatinine < 0.6 Then
      Apache2 = Apache2 + 2 * ARF
    End If

ONE80:
    If Hematocrit >= 0.6 Or Hematocrit < 0.2 Then
      Apache2 = Apache2 + 4
      GoTo ONE83
    End If
    If Hematocrit >= 0.5 Or Hematocrit <= 0.299 Then
      Apache2 = Apache2 + 2
      GoTo ONE83
    End If
    If Hematocrit >= 0.46 Then
      Apache2 = Apache2 + 1
    End If

ONE83:
    If WBC >= 40 Or WBC < 1 Then
      Apache2 = Apache2 + 4
      GoTo ONE86
    End If
    If WBC >= 20 Or WBC <= 2.9 Then
      Apache2 = Apache2 + 2
      GoTo ONE86
    End If
    If WBC >= 15 Then
      Apache2 = Apache2 + 1
    End If

ONE86:

If Age >= 75 Then
      Apache2 = Apache2 + 6
      GoTo ONE95
    End If
    If Age >= 65 Then
      Apache2 = Apache2 + 5
      GoTo ONE95
    End If
    If Age >= 45 Then
      Apache2 = Apache2 + 2
      GoTo ONE95
    End If
    GoTo ONE95

ONE90:
 
    If HCO3 >= 52 Or HCO3 < 15 Then
      Apache2 = Apache2 + 4
      GoTo ONE69
    End If
    If HCO3 >= 41 Or HCO3 <= 17.9 Then
      Apache2 = Apache2 + 3
      GoTo ONE69
    End If
    If HCO3 <= 21.9 Then
      Apache2 = Apache2 + 2
      GoTo ONE69
    End If
    If HCO3 >= 32 Then
      Apache2 = Apache2 + 1
      GoTo ONE69
    End If
    GoTo ONE69

ONE95:
Apache2 = 15 - Glasgow + CHP + Apache2
    If Option1(2).Value = True Then 'Emergency.Status = 1 Then
      Z7 = Exp(-3.517 + (Apache2 * 0.146) + 0.603 + PDC)
    Else
      Z7 = Exp(-3.517 + (Apache2 * 0.146) + PDC)
    End If


ONE98:
 
Apachetxt.Text = Format(Apache2, "###.0")
Apache2 = Val(Apachetxt.Text)
ApacheMortalitytxt.Text = Format(Z7 * 100 / (1 + Z7), "####.0") + " %"
ApacheIIMortality = Z7 * 100 / (1 + Z7) 'Val(ApacheMortalitytxt.Text)
End Sub

Private Sub FlatButton5_Click()
Unload Me
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

Private Sub Form_Unload(Cancel As Integer)
Set Apachefrm = Nothing
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


Private Sub Text1_Change()
End Sub

Private Sub Label11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Option7(1).Value = False Then        ' OrganNotxt.status <> 1  Then
            CHP = 5
        End If
        NonOpFlag = 1
        OpFlag = 0
        Call object1492
        EmergencyOpFlag = 0
        ElectiveOpFlag = 0
    Case 1
        If Option7(1).Value = False Then        'If OrganNotxt.Status <> 1 Then
            CHP = 2
        End If
        ElectiveOpFlag = 1
        OpFlag = 1
        Call object1488
        EmergencyOpFlag = 0
        NonOpFlag = 0
    Case 2
        If Option7(1).Value = False Then        'If OrganNotxt.Status <> 1 Then
            CHP = 5
        End If
        EmergencyOpFlag = 1
        OpFlag = 1
        Call object1488
        NonOpFlag = 0
        ElectiveOpFlag = 0
End Select

End Sub
Sub object1488()
PDCpop.Clear
PDCpop.AddItem "Multiple Trauma"
PDCpop.AddItem "Admit w/ Chronic CV Ds"
PDCpop.AddItem "Peripheral Vascular Sx"
PDCpop.AddItem "Heart Valve Sx"
PDCpop.AddItem "Craniotomy for Neoplasm"
PDCpop.AddItem "Renal Sx of Neoplasm"
PDCpop.AddItem "Renal Transplant"
PDCpop.AddItem "Head Trauma"
PDCpop.AddItem "Thoracic Sx for Neoplasm"
PDCpop.AddItem "Craniotomy (SAH/SDH/ICH)"
PDCpop.AddItem "Lamin or Spinal Cord Sx"
PDCpop.AddItem "Hemorrhagic Shock"
PDCpop.AddItem "GI Bleeding"
PDCpop.AddItem "GI Sx for Neoplasm"
PDCpop.AddItem "Resp Insuf Postop"
PDCpop.AddItem "GI Perf/Obstr"
PDCpop.AddItem "Neurologic"
PDCpop.AddItem "Cardiovascular"
PDCpop.AddItem "Respiratory"
PDCpop.AddItem "Gastrointestinal"
PDCpop.AddItem "Metabolic/Renal"
PDCpop.AddItem "Sepsis"
PDCpop.AddItem "Postrespir Arrest"
PDCpop.Text = PDCpop.List(0)
PDC = -1.685
PDCFlag = 1
PDCtxt.Text = Str(PDC)
End Sub

Sub object1492()
PDCpop.Clear
PDCpop.AddItem "Resp Fail d/t Asthma/Allergy"
PDCpop.AddItem "Resp Fail d/t COPD"
PDCpop.AddItem "Resp Fail d/t ARDS"
PDCpop.AddItem "Resp Fail d/t Postresp Arrest"
PDCpop.AddItem "Resp Fail d/t Aspir/Poison/Toxic"
PDCpop.AddItem "Resp Fail d/t Pulmonary Embolism"
PDCpop.AddItem "Resp Fail d/t Infection"
PDCpop.AddItem "Resp Fail d/t Neoplasm"
PDCpop.AddItem "CV Fail d/t HTN"
PDCpop.AddItem "CV Fail d/t Arhythmia"
PDCpop.AddItem "CV Fail d/t CHF"
PDCpop.AddItem "CV Fail d/t Hemorr Shock/Hypovol"
PDCpop.AddItem "CV Fail d/t CAD"
PDCpop.AddItem "CV Fail d/t Sepsis"
PDCpop.AddItem "CV Fail d/t Postcardiac Arrest"
PDCpop.AddItem "CV Fail d/t Cardiogenic Shock"
PDCpop.AddItem "CV Fail d/t Dissecting Aneurysm"
PDCpop.AddItem "Multiple Trauma"
PDCpop.AddItem "Head Trauma"
PDCpop.AddItem "Seizure Disorder"
PDCpop.AddItem "ICH/SDH/SAH"
PDCpop.AddItem "Drug Overdose"
PDCpop.AddItem "DKA"
PDCpop.AddItem "GI Bleeding"
PDCpop.AddItem "Metabolic/Renal"
PDCpop.AddItem "Respiratory"
PDCpop.AddItem "Neurologic"
PDCpop.AddItem "Cardiovascular"
PDCpop.AddItem "Gastrointestinal"
PDCpop.Text = PDCpop.List(0)
PDC = -2.108
PDCFlag = 1
PDCtxt.Text = Str(PDC)
End Sub

Private Sub Option7_Click(Index As Integer)
    Select Case Index
        Case 0
            Organ = 2
        Case 1
            Organ = 1
            CHP = 0
    End Select
End Sub

Private Sub Option8_Click(Index As Integer)
    Select Case Index
        Case 0
            ARF = 2
        Case 1
            ARF = 1
    End Select

End Sub

Private Sub PDCpop_Click()
Select Case OpFlag
Case 1
      Select Case PDCpop.ListIndex + 1
         Case 1
            PDC = -1.685
            PDCFlag = 1
         Case 2
            PDC = -1.376
            PDCFlag = 2
         Case 3
            PDC = -1.315
            PDCFlag = 3
         Case 4
            PDC = -1.261
            PDCFlag = 4
         Case 5
            PDC = -1.245
            PDCFlag = 5
         Case 6
            PDC = -1.204
            PDCFlag = 6
         Case 7
            PDC = -1.042
            PDCFlag = 7
         Case 8
            PDC = -0.955
            PDCFlag = 8
         Case 9
            PDC = -0.802
            PDCFlag = 9
         Case 10
            PDC = -0.788
            PDCFlag = 10
         Case 11
            PDC = -0.699
            PDCFlag = 11
         Case 12
            PDC = -0.682
            PDCFlag = 12
         Case 13
            PDC = -0.617
            PDCFlag = 13
         Case 14
            PDC = -0.248
            PDCFlag = 14
         Case 15
            PDC = -0.14
            PDCFlag = 15
         Case 16
            PDC = 0.06
            PDCFlag = 16
         Case 17  'Met/Renal
            PDC = -0.196
            PDCFlag = 17
         Case 18   'Resp
            PDC = -0.61
            PDCFlag = 18
         Case 19   'Neuro
            PDC = -1.15
            PDCFlag = 19
         Case 20   'CV
            PDC = -0.797
            PDCFlag = 20
         Case 21    'GI
            PDC = -0.613
            PDCFlag = 21
         Case 22   'Sepsis
            PDC = 0.113
            PDCFlag = 22
         Case 23   'PostResparrest
            PDC = -0.168
            PDCFlag = 23
      End Select
Case 0
      Select Case PDCpop.ListIndex + 1
         Case 1
            PDC = -2.108
            PDCFlag = 1
         Case 2
            PDC = -0.367
            PDCFlag = 2
         Case 3
            PDC = -0.251
            PDCFlag = 3
         Case 4
            PDC = -0.168
            PDCFlag = 4
         Case 5
            PDC = -0.142
            PDCFlag = 5
         Case 6
            PDC = -0.128
            PDCFlag = 6
         Case 7
            PDC = 0
            PDCFlag = 7
         Case 8
            PDC = 0.891
            PDCFlag = 8
         Case 9
            PDC = -1.798
            PDCFlag = 9
         Case 10
            PDC = -1.368
            PDCFlag = 10
         Case 11
            PDC = -0.424
            PDCFlag = 11
         Case 12
            PDC = 0.493
            PDCFlag = 12
         Case 13
            PDC = -0.191
            PDCFlag = 13
         Case 14
            PDC = 0.113
            PDCFlag = 14
         Case 15
            PDC = 0.393
            PDCFlag = 15
         Case 16
            PDC = -2.59
            PDCFlag = 16
         Case 17
            PDC = 0.731
            PDCFlag = 17
         Case 18
            PDC = -1.228
            PDCFlag = 18
         Case 19
            PDC = -0.517
            PDCFlag = 19
         Case 20
            PDC = -0.584
            PDCFlag = 20
         Case 21
            PDC = 0.723
            PDCFlag = 21
         Case 22
            PDC = -3.353
            PDCFlag = 22
         Case 23
            PDC = -1.507
            PDCFlag = 23
         Case 24
            PDC = 0.334
            PDCFlag = 24
         Case 25
            PDC = -0.885
            PDCFlag = 25
         Case 26
            PDC = -0.89
            PDCFlag = 26
         Case 27
            PDC = -0.759
            PDCFlag = 27
         Case 28
            PDC = 0.47
            PDCFlag = 28
         Case 29
            PDC = 0.501
            PDCFlag = 29
      End Select
End Select
PDCtxt.Text = Str(PDC)
End Sub
