VERSION 5.00
Begin VB.Form Theophylline 
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
   Begin VB.ComboBox TheophCF 
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
      Left            =   360
      TabIndex        =   14
      Text            =   "No Comorbid Conditions"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox CFtxt 
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
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   2295
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   4125
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Chiou 2 Point Kinetic Dosing?"
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
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   3840
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Non-Kinetic Dosage Adjustment?"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   3840
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Initial Maintainance Dose ="
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   3840
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Initial Loading Dose ="
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
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
      TabIndex        =   13
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Condition Factor ="
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
      Left            =   4320
      TabIndex        =   11
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Condition Factors Adults"
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
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Theophylline Dosing"
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
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Theophylline Dosing"
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
      Top             =   330
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
Attribute VB_Name = "Theophylline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CandyButton1_Click()

End Sub

Private Sub FlatButton5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
Dim selectedname As Integer

If Sex = 0 Or BSA = 0 Or Wtkg = 0 Or Htcm = 0 Or Age = 0 Or ABW = 0 Then
    TheophFlag = 1
    MsgBox "Please Enter Age, Height & Weight"
    Exit Sub
End If

TheophCF.Clear
TheophCF.AddItem "No Comorbid Conditions"
TheophCF.AddItem "Congestive heart failure"
TheophCF.AddItem "Pneumonia"
TheophCF.AddItem "Pulmonary edema"
TheophCF.AddItem "Sev bronchial obstruction"
TheophCF.AddItem "Erythromycin"
TheophCF.AddItem "Furosemide"
TheophCF.AddItem "Phenobarbital"
TheophCF.AddItem "Propranolol"
TheophCF.AddItem "Verapamil"
TheophCF.AddItem "Young adult smoker"
TheophCF.AddItem "Older patients (>60 yrs)"
TheophCF.ListIndex = 0 ' Highlight 1
CFtxt.Text = Format("1", "###.000")
ConditionFactor = 1
Multi = ConditionFactor

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
        Load TheophyllineLoad
        TheophyllineLoad.Show
    Case 1
        MGHR = 0.5 * ABW * ConditionFactor
        MgDay = MGHR * 24
        Interpretive.Interpret.Text = vbCrLf & vbCrLf & vbCrLf & "The Maintainance Dose =" & _
        Format(MGHR, "######.0") + "mg/hr" & vbCrLf & vbCrLf + Format(MgDay, "######.0") + "mg in 250 cc at " + Format(250 / 24, "####") + " cc/hr)"
        Load Interpretive
        Interpretive.Show
    Case 2
        Load NonKinetic
        NonKinetic.Show
    Case 3
        Load Chiou1
        Chiou1.Show
End Select

End Sub

Private Sub TheophCF_Click()

     Select Case TheophCF.ListIndex + 1
         Case 1
            ConditionFactor = 1 * Multi
         Case 2
            ConditionFactor = 0.4 * Multi
         Case 3
            ConditionFactor = 0.4 * Multi
         Case 4
            ConditionFactor = 0.4 * Multi
         Case 5
            ConditionFactor = 0.8 * Multi
         Case 6
            ConditionFactor = 0.6 * Multi
         Case 7
            ConditionFactor = 0.7 * Multi
         Case 8
            ConditionFactor = 1.2 * Multi
         Case 9
            ConditionFactor = 0.7 * Multi
         Case 10
            ConditionFactor = 0.5 * Multi
         Case 11
            ConditionFactor = 1.6 * Multi
         Case 12
            ConditionFactor = 0.6 * Multi
      End Select
      
Multi = ConditionFactor
CFtxt.Text = Format(ConditionFactor, "###.000")
End Sub
