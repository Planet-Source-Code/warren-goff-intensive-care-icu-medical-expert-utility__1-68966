VERSION 5.00
Begin VB.Form Renal 
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
      TabIndex        =   26
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Lipidstext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Osmgaptext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox CalcSosm1text 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox AGGtext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox AGtext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox FeNatext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox RFItext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox ActualCrCltext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox EstimatedCrCltext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox BUN_Crtext 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Interpret"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1065
      TabIndex        =   2
      Text            =   "Identify the patient's Volume Status"
      Top             =   4200
      Width           =   4095
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
      TabIndex        =   27
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Renal and Electrolytes"
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
      Left            =   975
      TabIndex        =   13
      Top             =   45
      Width           =   4455
   End
   Begin VB.Label Lipidz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Triglycerides=  "
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
      TabIndex        =   12
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label OsmGaptxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Osmolal Gap=  "
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
      TabIndex        =   11
      Top             =   3480
      Width           =   1605
   End
   Begin VB.Label CalcSosm1txt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculated Sosm=  "
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
      Top             =   3120
      Width           =   2010
   End
   Begin VB.Label AGGtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AG Gap=  "
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
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label AGtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anion Gap=  "
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
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label FENAtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fractional Excretion of Na+=   "
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
      TabIndex        =   7
      Top             =   2040
      Width           =   3105
   End
   Begin VB.Label RFItxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Renal Failure Index =  "
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2325
   End
   Begin VB.Label ActualCrCltxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Creatinine Clearance=   "
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
      Top             =   1305
      Width           =   3195
   End
   Begin VB.Label EstimatedCrCltxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Creatinine Clearance=  "
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
      TabIndex        =   4
      Top             =   960
      Width           =   3525
   End
   Begin VB.Label BUN_Crtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUN/Creatinine=  "
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
      Top             =   600
      Width           =   1860
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
      BackStyle       =   0  'Transparent
      Caption         =   "Renal and Electrolytes"
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
      Left            =   930
      TabIndex        =   14
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "Renal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
 Unload Interpretive
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub FlatButton5_Click()
Dim wNa As Single
Dim Wdef As Single
Dim Ndef As Single
Dim NaDegree As String
If Combo1.Text = "Identify the patient's Volume Status." Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Identify the patient's Volume Status, Please."
    SetTopMostWindow Me.hwnd, True

    Exit Sub
End If
Interpretive.Interpret.Text = "               Analysis of Renal & Electrolyte Values" & vbCrLf & vbCrLf

    
        wNa = Sodium
        If TP <> 0 Then
            wNa = wNa + ((TP - 8) * 0.025)
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
            "The Serum Sodium Corrected for the Plasma Protein Level (Na + (TP - 8) x .025 ) is " _
            + Format(Sodium + (TP - 8) * 0.025, "####") & vbCrLf & vbCrLf
            'MsgBox str(wNa)
        End If
        If Lipids <> 0 Then
            wNa = wNa + (Lipids * 0.002)
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
            "The Serum Sodium Corrected for the Plasma Triglycerides Level (Na + (Trig) x .002 ) is " _
            + Format(Sodium + (Lipids * 0.002), "####") & vbCrLf & vbCrLf
            'MsgBox str(wNa)
        End If
        If Glucose <> 0 Then
            wNa = wNa + (((Glucose - 100) / 100) * 1.8)
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
            "The Serum Sodium Corrected for the Blood Sugar (Na decreases 1.6 - 2 mEq/L for each 100 mg/dL rise in BS) is " _
            + Format(Sodium + (((Glucose - 100) / 100) * 1.8), "####") & vbCrLf & vbCrLf
            'MsgBox str(wNa)
        End If
        If wNa <> Sodium Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
            "The Cumulative Corrected Serum Sodium is " + Format(wNa, "####") _
             & vbCrLf & vbCrLf
        End If
        'Sodium = wNa
        If wNa < 140 And wNa > 130 Then
            NaDegree = "Mild "
        Else
            If wNa <= 130 And wNa > 125 Then
                NaDegree = "Moderate "
            Else
                If wNa <= 125 And wNa > 120 Then
                    NaDegree = "Severe "
                Else
                    If wNa <= 120 Then
                        NaDegree = "Life Threatening "
                    End If
                End If
            End If
        End If
        If Sex = 1 Then
            Wdef = (0.6 * IBW) * (1 - (wNa / 140))
            Ndef = (140 - wNa) * (0.6 * IBW)
            If Age >= 65 And BMI <= 20 Then
                Wdef = (0.5 * IBW) * ((wNa / 140) - 1)
                Ndef = (140 - wNa) * (0.5 * IBW)
            End If
        Else
            Wdef = (0.5 * IBW) * ((wNa / 140) - 1)
            Ndef = (140 - wNa) * (0.5 * IBW)
            If Age >= 65 And BMI <= 20 Then
                Wdef = (0.4 * IBW) * ((wNa / 140) - 1)
                Ndef = (140 - wNa) * (0.4 * IBW)
            End If
        End If
        Wdef = Wdef * 1000

    If wNa < 136 Then
        If Sosm < 280 Then
            If Combo1.Text = "Decreased Intravascular Volume" Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & NaDegree + "Hypotonic, Hypovolemic Hyponatremia Exists." + Chr(10) + Chr(10)
                If BUN / Creatinine > 20 And UricAcid > 8 And Uosm > 300 And uSodium < 20 Then
                    Interpretive.Interpret.Text = Interpretive.Interpret.Text & _
                    "Based upon an elevated BUN/Creatinine Ratio, an Elevated Uric Acid and Urine osmolality ( > 300) with a decreased Urine Sodium (< 20) " _
                    & "Strongly consider GI, Respiratory or Skin Losses.   Third spacing of fluids might also exist. " _
                    & "The Estimated Sodium Deficit is " + Format(Ndef, "#####.n") + " mEq " _
                    & "Replace no more than half that amount or " + Format(Ndef / 2, "#####.n") + " mEq with " + Format(Ndef / 308, "###.###") + " L Isotonic Saline " _
                    & "or " + Format(Ndef / 1026, "###.###") + " L 3% Hypertonic saline or " + Format(Ndef / 1710, "###.###") + " L 5% Hypertonic saline " _
                    & ". If Third Spaced Fluids are suspected clinically, the estimated Plasma (Free Water) Deficit is " + Format(Wdef, "#####") + " ml. " _
                    & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                    & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                    & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                    & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                    & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                    & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                    & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                Else
                    If BUN / Creatinine > 20 And UricAcid > 8 And Uosm >= Sosm - 100 And Uosm <= Sosm + 100 And uSodium >= 40 Then
                        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon an elevated BUN/Creatinine Ratio, an Elevated Uric Acid and Urine osmolality (= Sosm +/- 100 ) with an Increased Urine Sodium (> = 40) " _
                        & "Strongly Consider Renal Losses from Diuretics, Intrinsic Renal Disease or Partial Urinary Tract Obstruction. " _
                        & "The Estimated Sodium Deficit is " + Format(Ndef, "#####") + " mEq " _
                        & "Replace no more than half that amount or " + Format(Ndef / 2, "#####") + " mEq with " + Format(Ndef / 308, "####.###") + " L isotonic saline " _
                        & "or " + Format(Ndef / 1026, "####.###") + " L 3% Hypertonic saline or " + Format(Ndef / 1710, "####.###") + " L 5% Hypertonic saline " _
                        & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                        & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                        & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                        & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                        & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                        & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                        & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                    Else
                        If BUN >= 25 And Creatinine >= 2 And BUN / Creatinine >= 5 And BUN / Creatinine < 20 Then
                            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon an elevated BUN (>=25), Creatinine >=2, and BUN/Creatinine = 5-20, " _
                            & "Strongly Consider Renal Damage. " _
                            & "The Estimated Sodium Deficit is " + Format(Ndef, "#####") + " mEq " _
                            & "Replace no more than half that amount or " + Format(Ndef / 2, "#####") + " mEq with " + Format(Ndef / 308, "####.###") + " L isotonic saline " _
                            & "or " + Format(Ndef / 1026, "####.###") + " L 3% Hypertonic saline or " + Format(Ndef / 1710, "####.###") + " L 5% Hypertonic saline " _
                            & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                            & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                            & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                            & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                            & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                            & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                            & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                        Else
                            If Uosm > Sosm + 150 And UricAcid > 8 And BUN / Creatinine >= 20 And uSodium > 40 Then
                                Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon an elevated BUN/Creatinine Ratio (>=20), Uric Acid > 8, uSodium > 40 with Uosm > Sosm + 150, " _
                                & "Adrenal Insufficiency must be considered." _
                                & "The Estimated Sodium Deficit is " + Format(Ndef, "#####") + " mEq " _
                                & "Replace no more than half that amount or " + Format(Ndef / 2, "#####") + " mEq with " + Format(Ndef / 308, "####.###") + " L isotonic saline " _
                                & "or " + Format(Ndef / 1026, "####.###") + " L 3% Hypertonic saline or " + Format(Ndef / 1710, "####.###") + " L 5% Hypertonic saline " _
                                & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                                & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                                & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                                & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                                & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                                & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                                & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). "
                            Else
            '***Alternative diagnostic
             
                                Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon the provided Indices, no further Differential Diagnosis is possible between " _
                                & "Extra-Renal, Renal and Adrenal Insufficiency.   Correlate Clinically as Hyponatremia may be Multifactorial! " _
                                & "The Estimated Sodium Deficit is " + Format(Ndef, "#####") + " mEq " _
                                & "Replace no more than half that amount or " + Format(Ndef / 2, "#####") + " mEq with " + Format(Ndef / 308, "####.###") + " L isotonic saline " _
                                & "or " + Format(Ndef / 1026, "####.###") + " L 3% Hypertonic saline or " + Format(Ndef / 1710, "####.###") + " L 5% Hypertonic saline " _
                                & ". If Third Spaced Fluids are suspected clinically, the estimated Plasma (Free Water) Deficit is " + Format(1000 * ((Wtkg * 0.7) - (((Wtkg * 0.7 * 0.45) / Hematocrit))), "###") + " ml. " _
                                & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                                & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                                & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                                & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                                & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                                & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                                & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                            End If
                        End If
                    End If
                End If
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Reassess serum sodium concentrations every 2-4 hours during active intervention. "
            End If
            If Combo1.Text = "Increased Intravascular Volume" Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & NaDegree + "Hypotonic, Hypervolemic Hyponatremia Exists." + Chr(10) + Chr(10)
                If BUN / Creatinine > 20 And UricAcid > 8 And Uosm > 300 And uSodium < 20 Then
                    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon a BUN/Creatinine Ratio > 20, Uric Acid > 8, Urine osmolality > 300 and Urine Sodium < 20, " _
                    & "Strongly consider Congestive Heart Failure (Cardiac Edema), Hormonal Edema (Cushing's) or Cirrhosis (Hepatic Edema or other Circulatory Edematous State). " _
                    & "Despite the presence of Hyponatremia, Total Body Exchangeable Sodium is Increased in all edematous states. " _
                    & "In General, Fluid Restriction and Loop Diuretics are prescribed. " & vbCrLf & vbCrLf
                Else
                    If BUN > 25 And Creatinine >= 2 And BUN / Creatinine >= 5 And BUN / Creatinine < 20 Then
                        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon a BUN/Creatinine Ratio >=5 and < 20 and Creatinine > 2, Nephrosis Must be Strongly Considered " _
                        & "Despite the presence of Hyponatremia, Total Body Exchangeable Sodium is Increased in all edematous states. " _
                        & "In General, Fluid Restriction and Loop Diuretics are prescribed. " & vbCrLf & vbCrLf
                    Else
                        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon the Indices provided, a Differentiation between Cardiac, Circulatory, Hormonal, Hepatic and Renal Edematous States CannOT be made.   Correlate Clinically as Hyponatremia may be Multifactorial! " _
                        & "Despite the presence of Hyponatremia, Total Body Exchangeable Sodium is Increased in all edematous states. " _
                        & "In General, Fluid Restriction and Loop Diuretics are prescribed. " & vbCrLf & vbCrLf
                    End If
                End If
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Reassess serum sodium concentrations every 2-4 hours during active intervention. "
            End If
            If Combo1.Text = "Normal Intravascular Volume" Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & NaDegree + "Hypotonic, Isovolemic Hyponatremia Exists." + Chr(10) + Chr(10)
                'Deficit = (.4*WtKg)*(142-Sodium)
                'Cellular = (.4*WtKg)
                'Cellular = Cellular *(300-(2*(142-Sodium)))
                'MsgBox "Cellular " + str(Cellular)
                'Volume = Cellular/300
                'MsgBox "Volume Cell Water " + str(volume)
                'Volume = ((.4*WtKg) - Volume)
                'MsgBox "Volume Flows Out " + str(volume)
                'NaClMeq = (((Volume * 300) + Deficit))/2
                'MsgBox "Total Nacl needed " + str(NaClMeq)
                If BUN < 15 And Creatinine < 1 And UricAcid < 8 And Uosm < 350 And uSodium < 20 Then
                    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon a BUN < 15, Uric Acid < 8, Urine osmolality < 350 and Urine Sodium < 20, " _
                    & "Strongly consider Water Intoxication.   This may accompany the intake of any hypotonic fluids (4 L of beer/day)." _
                    & "In general, Fluid Restriction alone is all that is needed and will result in over a 1000 ml insensible loss/day." _
                    & "If the person is seizing, comatose or with acute neurologic deficit secondary to the severe hyponatremia, supplimentation with Hypertonic Saline should be initiated." _
                    & "Thus, there is a " + Format(Ndef, "#####") + " mEq Intracelluar and Extracellular Sodium Deficit. " _
                    & Format(Ndef / 1026, "####.###") + " L of 3% Hypertonic Saline Or " + Format(Ndef / 1710, "###.###") + " L of 5% Hypertonic Saline will increase the Serum Sodium by 10 meq/L." _
                    & "On Occassion, Normal Saline is used in these situations with a Loop Diuretic.   " + Format(Ndef / 308, "###.###") + " L of Normal Saline would be needed to increase the Serum Sodium by 10 meq/L.   Diuretics have a variable effect on serum Na+." _
                    & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause." _
                    & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly!" _
                    & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L." _
                    & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by" _
                    & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk." _
                    & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved." _
                    & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline.  If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr)." & vbCrLf & vbCrLf
                Else
                    If BUN > 30 And Creatinine >= 2 And (BUN / Creatinine) >= 5 And BUN / Creatinine < 20 And UricAcid > 8 And Uosm >= (Sosm - 100) And Uosm < (Sosm + 100) And uSodium > 40 Then
                        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon BUN > 30, Creatinine >= 2 And BUN/Creatinine >= 5, BUN/Creatinine < 20, Uric Acid > 8, Uosm >+ Sosm - 100, Uosm < Sosm + 100 and uSodium > 40, Renal Failure is the most Likely Etiology. " _
                        & "This entity represents Renal Salt Wasting or a Salt-Losing Nephritis in the " + "setting of Chronic Intrinsic Renal Disease (Chronic Pyelonephritis, Polycystic or Medullary Cystic Disease)." + " This is a Sodium Loss Syndrome and is aggravated by a Low Salt Diet and Excess Free Water. " _
                        & "In General, Fluid Restriction and Sodium Supplimentation (Shol's Solution sometimes" + " up to 40 gm/day) are prescribed.   High doses of Loop Diuretics and Zaroxylyn may be of assistance if edema occurs. " _
                        & "If Severe, Acute Hyponatremia Results in Coma or Neurologic Deficit, Hypertonic Saline may be needed. " _
                        & "Thus, there is a " + Format(Ndef, "####") + " meq Intracelluar and Extracellular Sodium Deficit. " _
                        & Format(Ndef / 1026, "###.###") + " L of 3% Hypertonic Saline Or " + Format(Ndef / 1710, "###.###") + " L of 5% Hypertonic Saline will increase the Serum Sodium by 10 meq/L. " _
                        & "On Occassion, Normal Saline is used in these situations with a Loop Diuretic.   " + Format(Ndef / 308, "###.###") + " L of Normal Saline would be needed to increase the Serum Sodium by 10 meq/L. Diuretics have a variable effect on serum Na+. " _
                        & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                        & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                        & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                        & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                        & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                        & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                        & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline. If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                    Else
                        If UricAcid > 8 And Uosm > 350 And uSodium < 20 Then
                            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon Uric Acid > 8, Uosm > 350 And uSodium < 20, consider the Effect of Profound Potassium Loss. " _
                            & "This is a rare Clinical Entity.   In general, treatment of the underlying cause of hypokalemia is needed as is supplimentation. " _
                            & "Fluid Restriction and Sodium Supplimentation are also prescribed. " _
                            & "If Severe, Acute Hyponatremia Results in Coma or Neurologic Deficit, Hypertonic Saline may be needed. " _
                            & "Thus, there is a " + Format(Ndef, "####") + " meq Intracelluar and Extracellular Sodium Deficit. " _
                            & Format(Ndef / 1026, "###.###") + " L of 3% Hypertonic Saline Or " + Format(Ndef / 1710, "###.###") + " L of 5% Hypertonic Saline will increase the Serum Sodium by 10 meq/L. " _
                            & "On Occassion, Normal Saline is used in these situations with a Loop Diuretic.   " + Format(Ndef / 308, "###.###") + " L of Normal Saline would be needed to increase the Serum Sodium by 10 meq/L. The effect of the diuretic on Serum Na+ is variable. " _
                            & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                            & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                            & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                            & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                            & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                            & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                            & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline. If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                        Else
                            If BUN < 15 And Creatinine < 1 And UricAcid < 6 And Uosm > Sosm And Sosm < 275 And uSodium > 20 Then
                                Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon BUN < 15, Creatinine < 1, Uric Acid < 6, Uosm > Sosm, Sosm < 275 and uSodium > 20, " _
                                & "The Syndrome of Inappropriate Secretion of Antidiuretic Hormone (SIADH) must be considered.   There must be normal " _
                                & "Renal, Adrenal and Thyroid function to make this diagnosis! " _
                                & "Fluid Restriction is often all that is required in an assymptomatic person. Normal Saline will have no effect on Serum Sodium as it will simply be excreted. " _
                                & "If Severe, Acute Hyponatremia Results in Coma or Neurologic Deficit, Hypertonic Saline may be needed. " _
                                & "Thus, there is a " + Format(Ndef, "####") + " meq Intracelluar and Extracellular Sodium Deficit. " _
                                & Format(Ndef / 1026, "###.###") + " L of 3% Hypertonic Saline Or " + Format(Ndef / 1710, "###.###") + " L of 5% Hypertonic Saline will increase the Serum Sodium by 10 meq/L. " _
                                & "On Occassion, Normal Saline is used in these situations with a Loop Diuretic.   " + Format(Ndef / 308, "###.###") + " L of Normal Saline would be needed to increase the Serum Sodium by 10 meq/L.   Diuretics have a variable effect on serum Na+. " _
                                & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                                & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly! " _
                                & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                                & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                                & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                                & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                                & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline. If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                            Else
            '***Alternative diagnostic
            
                                Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Based upon the indices provided a specific Diagnosis of Water Intoxication, Renal Salt Wasting or SIADH could not be made.   Correlate Clinically as Hyponatremia may be multifactorial. " _
                                & "There is and entity characterized by variable Urinary Indices with Normal Renal Function.   This is a Reset Osmostat. " _
                                & "There must be Normal Renal, Adrenal and Thyroid function to make this diagnosis! " _
                                & "Fluid Restriction is often all that is required in an assymptomatic person. Normal Saline will have no effect on Serum Sodium as it will simply be excreted. " _
                                & "If Severe, Acute Hyponatremia Results in Coma or Neurologic Deficit, Hypertonic Saline may be needed. " _
                                & "Thus, there is a " + Format(Ndef, "####") + " meq Intracelluar and Extracellular Sodium Deficit. " _
                                & Format(Ndef / 1026, "###.###") + " L of 3% Hypertonic Saline Or " + Format(Ndef / 1710, "###.###") + " L of 5% Hypertonic Saline will increase the Serum Sodium by 10 meq/L. " _
                                & "On Occassion, Normal Saline is used in these situations with a Loop Diuretic.   " + Format(Ndef / 308, "###.###") + " L of Normal Saline would be needed to increase the Serum Sodium by 10 meq/L.   Diuretics have a variable effect on serum Na+. " _
                                & Chr(10) + Chr(10) + "If the patient is NOT symptomatic (e.g. seizures, altered LOC, coma, focal neuro signs, lethargy, apathy or resp arrest), do NOT treat the hyponatremia but rather treat the underlying cause. " _
                                & Chr(10) + Chr(10) + "Severe hyponatremia (< 120) in an assymptomatic person usually reflects chronicity, e.g. development in over 48 hrs. These people are at high risk of CPM if treated rapidly!" _
                                & "To limit the risk of a demyelinating encephalopathy (Central Pontine Myelinolysis) the rate rise in Plasma Sodium should not exceed .5 mEq/L/Hr and the final Sodium concentration should not exceed 130 mEq/L. " _
                                & "CPM can lead to death, mutism, dysphasia, quadriparesis, pseudobulbar palsy, delirium and coma.   It appears to be caused by aggressive Tx of Hyponatremia present for longer than 24-48 hrs and raising the Na+ by " _
                                & "25 mEq/L or to normal or above in the first 48 hours.  Alcoholics, elderly women on Thiazides, malnourished, hypokalemia and burn patients are at increased risk. " _
                                & "Aggressive therapy should be discontinued when the serum sodium is raised 10% or when symptoms are relieved. " _
                                & "Only patients with Severe and Acute Hyponatremia (< 120, < 48 hrs) and (Severe) Symptoms should receive Hypertonic Saline. If actively seizing give 2-3 cc/kg/hr for 30-60 minutes and re-check Na+ (should increase by 4-6 meq/L/hr). " & vbCrLf & vbCrLf
                            End If
                        End If
                    End If
                End If
            End If
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Reassess serum sodium concentrations every 2-4 hours during active intervention. "
        End If
         If Sosm >= 280 And Sosm <= 300 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Isotonic Hyponatremia Exists!   This is also called Factitious or Pseudo- Hyponatremia. " _
                & "Isotonic infusions of Glucose, Mannitol or Glycine may cause this, as may Hyperlipidemia or Hyperproteinemia. "
         End If
         If Sosm > 300 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Hypertonic Hyponatremia Exists!   This is also called Factitious or Pseudo- Hyponatremia. " _
                & "Hyperglycemia must be suspected and corrected for.   Hypertonic infusions of Glucose, Mannitol or Glycine may cause this. "
         End If
    End If
    
    If wNa >= 136 And wNa <= 145 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Serum Sodium (Corrected) concentration is NORMAL (136-145). " & vbCrLf & vbCrLf
    End If
    If wNa > 145 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Hypernatremia exists ( > 145) which is always a hyperosmolar state.  Hypernatremia is almost never found in an alert person with access to water!  The elderly (>60) have decreased thirst osmotic stimulation and decreased maximal urinary concentration. " & vbCrLf & vbCrLf
        If Combo1.Text = "Decreased Intravascular Volume" Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Hypovolemic Hypernatremia exists. "
            If uSodium > 20 And Uosm > 800 Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & "Based on a urine Sodium in excress of 20, Renal losses of water by osmotic diuresis, hyperglycemia, urea, or mannitol should be considered. "
            End If
            If uSodium < 10 And Uosm > 800 Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & "Insensible losses of water via sweating, fever or respiration must be considered.   Likewise GI losses such as diarrhea may exist with urine Sodium < 10. "
            End If
            If uSodium >= 10 And uSodium <= 20 And Uosm > 800 Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & "It is not possible to differentiate renal from non-renal sources of water loss. "
            End If
            If Uosm <= 800 Then
                Interpretive.Interpret.Text = Interpretive.Interpret.Text & "With a submaximal urinary concentration (uosm <=800) considers Central Diabetes Insipidus with superimposed Volume Depletion. " _
                    & "Central DI commonly accompanies head trauma, hypoxic/ischemic encephalopathy and idiopathic. Nephrogenic DI may be caused by Lithium, hypercalcemia, hypokalemia, osmotic diuresis and sickle cell anemia. " & vbCrLf & vbCrLf
            End If
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The total free Water Deficit is " + Format(Wdef, "######") + "ml (Female = .5*IBW*([Sodium/140]-1) Male = .6*IBW*([Sodium/140]-1). " _
                & "Rapid correction should be avoided with the risk of cerebral edema.   Attempt to lower the serum Sodium concentration by about .5 mEq/L per hour and replace no more than 50% of the water deficit in the first 24 hours. " _
                    & "Use Normal Saline initially to correct the intravascular volume deficit.   Afterwards Half Normal Saline is fine. " & vbCrLf & vbCrLf
        End If
        If Combo1.Text = "Normal Intravascular Volume" Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Euvolemic Hypernatremia exists. " _
                & "Urine Sodium values are variable and of little use in this setting. " & vbCrLf & vbCrLf
            If Uosm < 200 Then
                   Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Renal losses of water by Central Or Nephrogenic Diabetes Insipidus is highly suggested with a uosmolity of < 200 mOsm/kg. " _
                   & "Central DI commonly accompanies head trauma, hypoxic/ischemic encephalopathy and idiopathic.   Nephrogenic DI may be caused by Lithium, hypercalcemia, hypokalemia, osmotic diuresis and sickle cell anemia. " & vbCrLf & vbCrLf
            Else
                If Uosm > 800 Then
                    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & "Hypothalamic disorders such as Primary Hypodipsia or a Reset Osmostat should be considered with normal maximum urinary concentration (uosm >800). " & vbCrLf & vbCrLf
                Else
                    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & "Partial Diabetes Insipidus or osmotic diuresis should be considered with uosm = 200 - 800 mOsm/kg. " & vbCrLf & vbCrLf
                End If
            End If
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The total free Water Deficit is " + Format(Wdef, "######") + "ml (Female = .5*IBW*([Sodium/140]-1) Male = .6*IBW*([Sodium/140]-1). " _
            & "Rapid correction should be avoided with the risk of cerebral edema.   Attempt to lower the serum Sodium concentration by about .5mEq/L per hour and replace no more than 50% of the water deficit in the first 24 hours. " _
            & "Use water or D5w is appropriate with a normal Volume Status. " & vbCrLf & vbCrLf
        End If
        If Combo1.Text = "Increased Intravascular Volume" Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Hypervolemic Hypernatremia exists.   And is secondary to Sodium Retention" + "(eg. xs Hypertonic NaHCO3 during Code). Hypertonic saline administration should also be considered.   Primary Hyperaldosteronism may be present" + "as well.   Urine sodium will exceed 20 in each of these cases and uosm will exceed 800. " & "The total free Water Deficit is " + Format(Wdef, "######") + "ml (Female = .5*IBW*([Sodium/140]-1) Male = .6*IBW*([Sodium/140]-1)." & "Rapid correction should be avoided with the risk of cerebral edema.   Attempt to lower the serum Sodium concentration by about .5mEq/L per hour and replace no more than 50% of the water deficit in the first 24 hours. " & "Remove the source of Salt Excess, administer diuretics and use water or D5W. " & vbCrLf & vbCrLf
        End If
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & vbCrLf & _
    "Reassess serum sodium concentrations every 2-4 hours during active intervention. " & vbCrLf & vbCrLf
    End If
    ATN
    Load Interpretive
    Interpretive.Show
End Sub
Sub CalcIt()

    If BUN <> 0 And Creatinine <> 0 Then
        BUN_Crtext.Text = Format(BUN / Creatinine, "####.0")
        BUN_Cr = Val(BUN_Crtext.Text)
    End If
    If EstimatedCrCl <> 0 Then
        EstimatedCrCltext.Text = Format(EstimatedCrCl, "####.0")
        EstimatedCrCl = Val(EstimatedCrCltext.Text)
    End If
    If ActualCrCl <> 0 Then
        ActualCrCltext.Text = Format(ActualCrCl, "####.0")
        ActualCrCl = Val(ActualCrCltext.Text)
    End If
    If RFI <> 0 Then
        RFItext.Text = Format(RFI, "##.000")
        RFI = Val(RFItext.Text)
    End If
    If FeNa <> 0 Then
        FeNatext.Text = Format(FeNa, "##.000")
        FeNa = Val(FeNatext.Text)
    End If
    If Sodium <> 0 And Chloride <> 0 And CO2 <> 0 Then
        AGtext.Text = Format(Sodium - (Chloride + CO2), "###.0")
        AG = Val(AGtext.Text)
        AGGtext.Text = Format(24 - (AG - 12), "###.0")
        AGG = Val(AGGtext.Text)
    End If
    If CalcSosm <> 0 Then
        CalcSosm1text.Text = Format(CalcSosm, "####.0")
    End If
    If Sosm <> 0 And CalcSosm <> 0 Then
        Osmgaptext.Text = Format(Sosm - CalcSosm, "####.0")
        Osmgap = Val(Osmgaptext.Text)
    End If
    If Lipids <> 0 Then
        Lipidstext.Text = Format(Lipids, "####")
    End If
End Sub
Sub Calc1()
    Creatinine = Val(DataInput.Creatininetxt.Text)
    Sodium = Val(DataInput.Natxt.Text)
    Potassium = Val(DataInput.Ktxt.Text)
    uPotassium = Val(DataInput.uPotassiumtxt.Text)
    Chloride = Val(DataInput.CLtxt.Text)
    uChloride = Val(DataInput.uChloridetxt.Text)
    CO2 = Val(DataInput.CO2txt.Text)
    BUN = Val(DataInput.BUNtxt.Text)
    VolumeUrine = Val(DataInput.VolumeUrinetxt.Text)
    TimedUrine = Val(DataInput.TimedUrinetxt.Text)
    UCreatinine = Val(DataInput.UCreatininetxt.Text)
    uSodium = Val(DataInput.Usodiumtxt.Text)
    Sosm = Val(DataInput.Sosmtxt.Text)
    Glucose = Val(DataInput.BStxt.Text)
    UricAcid = Val(DataInput.Urate.Text)
    Uosm = Val(DataInput.Uosml.Text)
    If Magnesium <> 0 Then
        DataInput.Magn.Text = Format(Magnesium, "###.##")
        Magnesium = Val(DataInput.Magn.Text)
    End If
    If Calcium <> 0 Then
        DataInput.Calc.Text = Format(Calcium, "###.##")
        Calcium = Val(DataInput.Calc.Text)
    End If
    If Phosphorus <> 0 Then
        DataInput.Phos.Text = Format(Phosphorus, "###.##")
        Phosphorus = Val(DataInput.Phos.Text)
    End If
    If UricAcid <> 0 Then
        DataInput.Urate.Text = Format(UricAcid, "###.##")
        UricAcid = Val(DataInput.Urate.Text)
    End If
    
    '- M : Clcr (ml/min) = ( 140 - age ) x (tbw or ibw)(kg) / (72 x serum creatinine(mg/dl))
    '- F : Clcr (ml/min) = 0.85 x ( 140 - age ) x (tbw or ibw)(kg) / (72 x serum creatinine(mg/dl))
    '- In bedside tools, estimation of Ccr is performed with tbw and ibw ( Peck formula ).
    
    If Wtkg = 0 Or Sex = 0 Or Htcm = 0 Or ABW = 0 Then
        RenalFlag = 1
        MsgBox "Please Enter Age, Sex, Height & Weight"
        Exit Sub
    End If
    
    If Creatinine = 0 Then
        MsgBox "Please Input Serum Creatinine"
        Exit Sub
    End If
    
        If Sex = 1 Then
            'EstimatedCrCltxt.label = Format((((140 - Age) * ABW) / (Creatinine * 72)),"####.0")
            'EstimatedCrCl = val(EstimatedCrCltxt.label)
             EstimatedCrCl = (((140 - Age) * ABW) / (Creatinine * 72))
        Else
            'EstimatedCrCltxt.label = Format((((140 - Age) * ABW) / (Creatinine * 72))*.85,"####.0")
            'EstimatedCrCl = val(EstimatedCrCltxt.label)
             EstimatedCrCl = (((140 - Age) * ABW) / (Creatinine * 72)) * 0.85
        End If
            
                                    
    
    If uSodium <> 0 And UCreatinine <> 0 Then
        'RFItxt.label = Format(USodium / UCreatinine / Creatinine,"####.0")
        'RFI = val(RFItxt.label)
        RFI = uSodium / UCreatinine / Creatinine
    End If
    
    If uSodium <> 0 And UCreatinine <> 0 And Sodium <> 0 Then
        'FENAtxt.label =  Format(USodium / Sodium / UCreatinine * Creatinine * 100,"####.0")
        'FENa  = val(FENAtxt.label)
        FeNa = uSodium / Sodium / UCreatinine * Creatinine * 100
    End If
    
    If UCreatinine <> 0 And VolumeUrine <> 0 And TimedUrine <> 0 Then
        'ActualCrCLtxt.label = Format(UCreatinine * VolumeUrine / TimedUrine / 60 / Creatinine,"####.0")
        ActualCrCl = UCreatinine * VolumeUrine / TimedUrine / 60 / Creatinine
    End If
    If Glucose <> 0 And BUN <> 0 And Sodium <> 0 Then
        DataInput.CalcSosmtxt.Text = Format(2 * Sodium + Glucose / 18 + BUN / 2.8, "####.0")
        CalcSosm = Val(DataInput.CalcSosmtxt.Text)
    End If
    If UricAcid <> 0 Then
        DataInput.Urate.Text = Str(UricAcid)
    End If
    If Uosm <> 0 Then
        DataInput.Uosml.Text = Str(Uosm)
    End If

End Sub
Sub ATN()
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & ""
If Uosm > 500 And uSodium < 20 And UCreatinine / Creatinine > 40 And FeNa < 1 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Based upon a Urine Osmolality > 500 mOsm/kg, Urine Sodium < 20 meq/L, " _
    & "Urine/Serum Creatinine > 40 and Fractional Excretion of Sodium < 1, " _
    & "Prerenal Azotemia is most likely. The UA will usually be normal with occasional " _
    & "hyaline or finely granular casts. " & vbCrLf & vbCrLf
Else
    If Uosm < 400 And uSodium > 40 And UCreatinine / Creatinine < 20 And FeNa > 1 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
        "Based upon a Urine Osmolality < 400 mOsm/kg, Urine Sodium > 40 meq/L, " _
        & "Urine/Serum Creatinine < 20 and Fractional Excretion of Sodium > 1, " _
        & "Acute Tubular Necrosis is most likely. The UA will usually show Renal Tubular Epithelial Cells, " _
        & "Granulars and Muddy Brown granular casts. " & vbCrLf & vbCrLf
    Else
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
        "Based upon Urine Osmolality, Urine Sodium, Urine/Serum Creatinine and Fractional Excretion of Sodium, " _
        & "Prerenal Azotemia and Acute Tubular Necrosis CANNOT be differentiated. Check the urine sediment. " _
        & "In ATN expect Renal Tubular Epithelial Cells, Granulars and Muddy Brown granular casts, while Prerenal " _
        & "Azotemia will be normal with occasional hyaline or finely granular casts. " & vbCrLf & vbCrLf
    End If
End If
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & Chr(10) + Chr(10) _
& "First, clinically exclude prerenal causes (volume depletion, chf, cirrhosis, nsaids, ace-inhibitors.) " _
& "Next, clinically exclude postrenal causes (renal ultrasound & post-void residuals.) " _
& "Note RBC casts implies Glomerulonephritis or Vasculitis while Pyuria accompanies Interstitial Nephritis. " _
& "Consult Nephrology when serum Creatinine exceeds 2.0.  Dialysis may be needed in 85% of Oliguric (< 400 cc/24 hr) " _
& "ATN and 30-40% non-oliguric ATN pts. " & vbCrLf & vbCrLf

Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
"Since increased tubular secretion of Creatinine typically occurs at low Glomerular Filtration Rates (GFR), " _
& "Creatinine Clearance often OVER-ESTIMATES GFR!  Drugs (Cimetidine & Trimethoprim) may inhibit tubular creatinine " _
& "secretion resulting in an increase in serum Creatinine with no change in GFR.  One must evaluate a pt. For Acute Renal Failure (ARF) when " _
& "serum Creatinine increases by > 0.5 mg/dl. " & vbCrLf & vbCrLf _
& "Advanced Chronic Renal Failure and recent diuretic tx may alter the above indices.  ATN in the setting of rhabdomyolysis/myoglobinuria, hemolysis, " _
& "Sepsis, cirrhosis, CHF and Radiocontrast Nephropathy may present with urine Sodium < 10 and FeNa+ < 1%.  There is no Gold Standard.  The mortality " _
& "in ATN requiring dialysis is 50-80% and has not been improved despite technological innovations.  It is underestimated by APACHE II & III scores (See Renal Mortality). " & vbCrLf & vbCrLf _
& "Using Non-ionic Contrast in high risk patients with CRF and DM may be preventative.  Likewise, Acetylcysteine (600 mg BID)decreased the risk of ATN from 21% to 2% in one study. " _
& "Mannitol and Calcium Channel Blockers may maintain GFR in people post-cadaveric renal transplant." & vbCrLf & vbCrLf _
& "Maintaining renal perfusion (MAP) is essential for pts in ATN.  Alpha-Adrenergic pressors decrease renal perfusion and should be avoided.  Renal dose Dopamine is useless. " _
& "Vasopressin may increase SVR without decreasing renal perfusion.  Hypertension need not be treated aggresively unless Crisis (Organ Damage) exists." & vbCrLf & vbCrLf _
& "ATN is a Catabolic State with Protein Energy Malnutrition frequently associated.  Enteral nutrition is preferred.  Lipids may be detrimental.  One should provide protein and non-protein " _
& "caloric energy expenditures not to exceed 1.5 g protein/kg/day in ATN." + Chr(10) + Chr(10) _
& "Dialysis using biocompatible membranes may be delivered intermittently or continuously.  The continuous method is better tolerated in hypotension.  Continuous Renal Replacement Methods " _
& "may be attempted as well.  At present, no intervention clearly effects morbidity, mortality, LOS or cost."
End Sub
Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
Combo1.AddItem "Normal Intravascular Volume"
Combo1.AddItem "Decreased Intravascular Volume"
Combo1.AddItem "Increased Intravascular Volume"
Combo1.Text = "Identify the patient's Volume Status."
Calc1
CalcIt
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
