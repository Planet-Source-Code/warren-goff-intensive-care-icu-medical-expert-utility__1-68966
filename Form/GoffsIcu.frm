VERSION 5.00
Begin VB.Form Hemodynamics 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   6015
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "GoffsIcu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   2220
      Picture         =   "GoffsIcu.frx":08CA
      ScaleHeight     =   930
      ScaleWidth      =   1290
      TabIndex        =   5
      ToolTipText     =   "Data Input"
      Top             =   90
      Width           =   1290
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Input"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Input"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5490
      Picture         =   "GoffsIcu.frx":1756
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5550
      Picture         =   "GoffsIcu.frx":1850
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   225
      Visible         =   0   'False
      Width           =   495
   End
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   150
      Top             =   2400
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   1095
      Picture         =   "GoffsIcu.frx":194A
      ScaleHeight     =   4245
      ScaleWidth      =   4845
      TabIndex        =   3
      Top             =   1230
      Width           =   4845
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   1
      Left            =   2265
      Picture         =   "GoffsIcu.frx":F934
      ScaleHeight     =   930
      ScaleWidth      =   1290
      TabIndex        =   6
      Top             =   165
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Index           =   0
      Left            =   60
      Picture         =   "GoffsIcu.frx":107C0
      ScaleHeight     =   1050
      ScaleWidth      =   795
      TabIndex        =   7
      ToolTipText     =   "File Operations"
      Top             =   105
      Width           =   795
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   330
         TabIndex        =   16
         Top             =   15
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   17
         Top             =   30
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Index           =   1
      Left            =   135
      Picture         =   "GoffsIcu.frx":1149D
      ScaleHeight     =   1080
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   3915
      Picture         =   "GoffsIcu.frx":1217A
      ScaleHeight     =   930
      ScaleWidth      =   1320
      TabIndex        =   9
      ToolTipText     =   "Interpretations & Modules"
      Top             =   60
      Width           =   1320
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   15
         Top             =   15
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   1
      Left            =   3915
      Picture         =   "GoffsIcu.frx":16B3C
      ScaleHeight     =   930
      ScaleWidth      =   1080
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
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
      Left            =   6000
      TabIndex        =   11
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label107 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   0
      Left            =   6285
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   105
      Width           =   375
   End
   Begin VB.Label Label107 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Index           =   1
      Left            =   6255
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu ksfgs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuCalc 
      Caption         =   "Calculations"
      Visible         =   0   'False
      Begin VB.Menu mnuHemodynamics 
         Caption         =   "Hemodynamics"
      End
      Begin VB.Menu mnuABGs 
         Caption         =   "Arterial Blood Gases"
         Begin VB.Menu mnuInterpretABGs 
            Caption         =   "Interpretation ABGs"
         End
         Begin VB.Menu mnuOxyHb 
            Caption         =   "Oxyhemoglobin Dissociation"
         End
         Begin VB.Menu mnuO2Dose 
            Caption         =   "Oxygen Dosage"
         End
      End
      Begin VB.Menu mnuRenal 
         Caption         =   "Renal"
         Begin VB.Menu mnurCalc 
            Caption         =   "Calculations"
         End
         Begin VB.Menu mnuFluidChal1 
            Caption         =   "Fluid Challenge"
         End
      End
      Begin VB.Menu mnuNutrition 
         Caption         =   "Nutrition"
      End
      Begin VB.Menu mnuDrugs 
         Caption         =   "Drug Dosing"
         Begin VB.Menu mnuAmino 
            Caption         =   "Aminoglycisides"
            Begin VB.Menu mnuKinetics 
               Caption         =   "Kinetic Dosing"
            End
            Begin VB.Menu mnuOncee 
               Caption         =   "Once Daily Dosing"
            End
            Begin VB.Menu mnuCaveats 
               Caption         =   "Caveats & Instructions"
            End
         End
         Begin VB.Menu mnuTheoph 
            Caption         =   "Theophyline"
         End
         Begin VB.Menu mnuVanco 
            Caption         =   "Vancomycin"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDig 
            Caption         =   "Digoxin"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHeparin 
            Caption         =   "Heparin"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDilantin 
            Caption         =   "Dilantin"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSteroids 
            Caption         =   "Corticosteroids"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuVents 
         Caption         =   "Mechanical Ventilation"
      End
      Begin VB.Menu mnuMortality 
         Caption         =   "Mortality"
         Begin VB.Menu mnuApache 
            Caption         =   "Apache II"
         End
         Begin VB.Menu mnuGlasgow 
            Caption         =   "Glasgow Coma Score"
         End
         Begin VB.Menu mnuARDS 
            Caption         =   "ARDS Scoring"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFluidChal2 
         Caption         =   "Fluid Challenge"
      End
      Begin VB.Menu mnuPFT 
         Caption         =   "Pulmonary Function Testing"
      End
      Begin VB.Menu mnuTransfusions 
         Caption         =   "Transfusion Medicine"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPE 
         Caption         =   "Pulmonary Embolism"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPneumonia 
         Caption         =   "Pneumonia"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHIV 
         Caption         =   "HIV"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Hemodynamics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CandyButton1_Click()
    PopupMenu mnuCalc
End Sub

Private Sub CandyButton3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataInput.Left = Hemodynamics.Left + 100
DataInput.Top = Hemodynamics.Top + 900
If CandyButton3.Checked = False Then
    CandyButton3.Checked = True
    Load DataInput
    DataInput.Show
Else
    CandyButton3.Checked = False
    DataInput.Visible = False
End If

End Sub

Sub OxyHb()
Interpretive.Interpret.Text = "              Secondary Analysis of Blood Gas Values" & vbCrLf & vbCrLf
'If X <> 0 And ABGFlag = 1 Then
'ABGFlag = 0
Dim SO2 As Single
Dim PO2 As Single
Dim PO2C As Single
Dim PO2S As Single
Dim p50 As Single
Dim a As Single
Dim a1 As Single
Dim a2 As Single
Dim a3 As Single
Dim a4 As Single
Dim a5 As Single
Dim a6 As Single
Dim a7 As Single
Dim b As Single
Dim c As Single
Dim i As Integer
Dim i1 As Single
Dim DPO2 As Single
'Oxyhemoglobin and p50
a1 = -8532.2289
a2 = 2121.401
a3 = -67.073989
a4 = 935960.87
a5 = -31346.258
a6 = 2396.1674
a7 = -67.104406
DPO2 = 99
PO2 = 100
'& vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & _
"Fitting Data to Normal OHD Curve:" & vbCrLf & vbCrLf & vbCrLf
One:
X = (i - 27.5) / (i + 27.5)
SO2 = 100 * (PO2 * (PO2 * (PO2 * (PO2 + a3) + a2) + a1)) / (PO2 * (PO2 * (PO2 * (PO2 + a7) + a6) + a5) + a4)
'Moose.Label = Str(SaO2) + " " + Str(Int(SO2 * 10) / 10)
If DPO2 < 0.01 Then
'If int(SaO2*10)/10 = int(SO2*10)/10 then
    PO2S = PO2
    GoTo Here
End If
If SO2 > SaO2 Then
    PO2 = PO2 - DPO2
    DPO2 = DPO2 / 2
    GoTo One
End If
PO2 = PO2 + DPO2
DPO2 = DPO2 / 2
GoTo One
Here:

Y = (Temp - 32) * (5 / 9)
X = 0.024 * (37 - Y) + 0.4 * (pH - 7.4) + 0.6 * (Log(40) / (Log(10#)) - Log(PaCO2) / (Log(10#))) '(log10(40) - log10(PaCO2))
z = 10 ^ X 'pow(10, X)
PO2C = PaO2 * z
If PO2S <> 0 Then
    p50 = 26.6 * (PO2C / PO2S)
Else
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
    "--The PaO2-SaO2 do not reside on a Physiologic OHD Curve! Please recheck the Data." & vbCrLf & vbCrLf
    Load Interpretive
    Interpretive.Show
    Exit Sub
End If

Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"--The PaO2 corrected to pH, PaCO2 and Temperature = " + Format(PO2C, "###.0") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"--The PaO2 predicted from the SaO2 on a normal OHD Curve = " + Format(PO2S, "###.0") & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"--The P50 = " + Format(p50, "###.0") + " (Normallly = 27 +/- 1.3mmHg)" & vbCrLf & vbCrLf & vbCrLf & vbCrLf


If p50 >= 28.3 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
    "The Oxyhemoglobin Dissociation Curve is Shifted to the Right" & vbCrLf + Chr(10)
Else
    If p50 <= 25.7 Then
       Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
       "The Oxyhemoglobin Dissociation Curve is Shifted to the Left" & vbCrLf + Chr(10)
    Else
        MoreTxt.Label = MoreTxt.Label + _
        "The Oxyhemoglobin Dissociation Curve is not Shifted" & vbCrLf + Chr(10)
    End If
End If

Load Interpretive
Interpretive.Show
End Sub

Private Sub Form_Load()
    sSVCSO2 = 30
    iSVCSO2 = 34
    RASO2 = 26
    RVSO2 = 35
    PASO2 = 32
    CSSO2 = 15
    PCSO2 = 120
    DataInput.SSTab1.Tab = 0
    DataInput.SStab2.Tab = 0

    Load DataInput
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 SetTopMostWindow Me.hwnd, True

    IntSave = MsgBox("Do you want to Exit This Program?", _
                     vbYesNoCancel + vbQuestion)
    Select Case IntSave
        Case vbYes
            CloseAll
        Case vbNo
            Cancel = 1
            Exit Sub

        Case vbCancel
            Cancel = 1
            Exit Sub
    End Select

End Sub

Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count
    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub




Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2(0).Visible = False
    Picture2(1).Visible = True

End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DataInput.SSTab1.Tab = 0
    DataInput.SStab2.Tab = 0
    Picture2(0).Visible = True
    Picture2(1).Visible = False
    DataInput.Left = Hemodynamics.Left + 100
    DataInput.Top = Hemodynamics.Top + 900
    Load DataInput
    DataInput.Show

End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture5(0).Visible = False
    Picture5(1).Visible = True

End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture5(0).Visible = True
    Picture5(1).Visible = False
    PopupMenu mnuFiles

End Sub

Private Sub Label31_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture6(0).Visible = False
    Picture6(1).Visible = True

End Sub

Private Sub Label31_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture6(0).Visible = True
    Picture6(1).Visible = False
    PopupMenu mnuCalc

End Sub

Private Sub mnuApache_Click()
On Error Resume Next
    Load Apachefrm
    Apachefrm.Show
End Sub

Private Sub mnuCaveats_Click()
Interpretive.Interpret.Text = "Exclusions For Single Daily Aminoglycoside Dosing:" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Elderly patients." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Neutropenic patients (ANC < 1000)" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Endocarditis." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Patients treated For suspected Or documented endocarditis" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Infants." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Patients with Cystic Fibrosis, Cirrhosis, Ascites, Or" & vbCrLf & _
"   Myasthenia Gravis" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   > 20% BSA burns." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Pregnant patients" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   History of hearing loss or vestibular dysfunction." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Dialysis patients" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Patients treated For staphylococcal Or enterococcal" & vbCrLf & _
"   infections" & vbCrLf & _
"   when aminoglycoside is used for Synergy and Mycobacterial" & vbCrLf & _
"   Infections." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Patients in acute renal failure" & vbCrLf & vbCrLf & vbCrLf


Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"General Information for Kinetic Dosing:" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Check Levels 24 hours after initiating Tx (Pk/Tr) after" & vbCrLf & _
"   24 hours (4 x T1/2) and every 2-3 days." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Bun/Creatinine <= qod." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Weight q 2-7 days" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Daily I&O" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Baseline & Weekly Audiograms and check for tinnitis/vertigo" & vbCrLf & _
"   daily." & vbCrLf + Chr(10)
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Gent/Tobra/Netil => Serious Peak 6-8 Trough .5 - 1.5" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Gent/Tobra/Netil => Life Threatening Peak 8-10 Trough 1-< 2." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Amikacin/Kanamycin => Serious Peak 20-25 Trough 1 - 4" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Amikacin/Kanamycin => Life Threatening Peak 25-30 Trough 4-8" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   Timing of Levels must be precise.  Peak MUST be 15-30 min" & vbCrLf & _
"   after infusion and trough 30 min prior to next dose.  Documentation" & vbCrLf & _
"   is essential and if timing is inaccurate, dosing will not be correct!" & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   A Correction Factor is applied for Obesity.  The Ideal Body Weight" & vbCrLf & _
"   is used for those underweight." & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
"   This algorithm is based on a Single Compartment Model of" & vbCrLf & _
"   Sarrubi & Hull which has been shown to result a 90% accuracy." & vbCrLf

Load Interpretive
Interpretive.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me

End Sub

Private Sub mnuFluidChal_Click()
    Load FluidChal
    FluidChal.Show
End Sub

Private Sub mnuFluidChal1_Click()
    Load FluidChal
    FluidChal.Show

End Sub

Private Sub mnuFluidChal2_Click()
    Load FluidChal
    FluidChal.Show

End Sub

Private Sub mnuGlasgow_Click()
    Load GlasgowFrm
    GlasgowFrm.Show
End Sub

Private Sub mnuHemodynamics_Click()
    DataInput.Show
    DataInput.SSTab1.Tab = 1
    DataInput.FlatButton5_Click
End Sub

Private Sub mnuInterpretABGs_Click()
    DataInput.ABGInterpret
End Sub

Private Sub mnuKinetics_Click()
    Load AminoGlyc
    AminoGlyc.Show

End Sub

Private Sub mnuNutrition_Click()
Load Nutritional
Nutritional.Show
End Sub

Private Sub mnuO2Dose_Click()
    Load O2Dose
    O2Dose.Show
End Sub

Private Sub mnuOncee_Click()
    Load OnceDaily
    OnceDaily.Show

End Sub

Private Sub mnuOxyHb_Click()
OxyHb
End Sub

Private Sub mnurCalc_Click()
    Load Renal
    Renal.Show

End Sub

Private Sub mnuTheoph_Click()
    Load Theophylline
    Theophylline.Show
End Sub

Private Sub mnuVents_Click()
Load Ventilator
Ventilator.Show

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub


Private Sub Picture2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2(0).Visible = False
    Picture2(1).Visible = True
End Sub



Private Sub Picture2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DataInput.SSTab1.Tab = 0
    DataInput.SStab2.Tab = 0
    Picture2(0).Visible = True
    Picture2(1).Visible = False
    DataInput.Left = Hemodynamics.Left + 100
    DataInput.Top = Hemodynamics.Top + 900
    Load DataInput
    DataInput.Show

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
Picture4.Visible = True
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
Picture4.Visible = False
'OpenBrowser App.Path & "\help.htm", Hemodynamics.hWnd
End Sub

Private Sub Picture5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture5(0).Visible = False
    Picture5(1).Visible = True
End Sub

Private Sub Picture5_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture5(0).Visible = True
    Picture5(1).Visible = False
    PopupMenu mnuFiles
End Sub

Private Sub Picture6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture6(0).Visible = False
    Picture6(1).Visible = True
End Sub

Private Sub Picture6_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture6(0).Visible = True
    Picture6(1).Visible = False
    PopupMenu mnuCalc

End Sub
