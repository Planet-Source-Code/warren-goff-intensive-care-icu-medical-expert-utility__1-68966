VERSION 5.00
Begin VB.Form OnceDaily 
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
   Begin VB.CommandButton Command1 
      Caption         =   "Caveats"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox estBwag 
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
      Height          =   345
      Left            =   4200
      TabIndex        =   19
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox ActBwag 
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
      Height          =   345
      Left            =   4200
      TabIndex        =   18
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox eCrClag 
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
      Height          =   345
      Left            =   4200
      TabIndex        =   17
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox aCrClag 
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
      Height          =   345
      Left            =   4200
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton CandyButton1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2130
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Kanamycin"
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   1920
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Amikacin"
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
         Width           =   1920
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Netilimicin"
         Enabled         =   0   'False
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
         Width           =   1920
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Tobramycin"
         Enabled         =   0   'False
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
         Width           =   1920
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Gentamycin"
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
         Width           =   1920
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
      Index           =   4
      Left            =   5880
      TabIndex        =   21
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Body Weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   15
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Body Weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Creatinine Clearance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Creatinine Clearance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aminoglycosides Dosing q Day"
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
      Left            =   585
      TabIndex        =   3
      Top             =   360
      Width           =   4965
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aminoglycosides Dosing q Day"
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
      Left            =   555
      TabIndex        =   2
      Top             =   330
      Width           =   4965
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
Attribute VB_Name = "OnceDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
If AminoRx = "" Then
    SetTopMostWindow Me.hwnd, False
    MsgBox "Please Select a Drug!"
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If
Interpretive.Interpret.Text = "               Once a Day Aminoglycoside Dosing." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
If ActualCrCl <> 0 Then
    CrCl = ActualCrCl
Else
    If EstimatedCrCl <> 0 Then
        CrCl = EstimatedCrCl
    Else
        Exit Sub
    End If
End If
Select Case AminoRx

Case "Gentamycin"
If CrCl >= 50 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 5mg/kg/24hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 5 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "nn####") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Else
        w = 5 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
 End If
If CrCl >= 30 And CrCl < 50 Then  '= 5mg/kg/36hrs
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 5mg/kg/36hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 5 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Else
        w = 5 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
If CrCl >= 20 And CrCl < 30 Then  ' = 5mg/kg/48hrs
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 5mg/kg/48hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 5 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 5 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
If CrCl < 20 Then '= 2mg/kg X 1 & CONSULT KINETICS
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 2mg/kg ONE DOSAGE. With ARF Use Kinetic Dosing!" & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 2 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 2 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If


Case "Amikacin"
If CrCl >= 40 Then '= 15mg/kg/24hrs
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 15mg/kg/24hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 15 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 15 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
If CrCl >= 30 And CrCl < 40 Then    'ml/min = 15mg/kg/36hrs
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 15mg/kg/36hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 15 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 15 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
If CrCl >= 20 And CrCl < 30 Then    ' = 15mg/kg/48hrs
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 15mg/kg/48hrs." & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 15 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 15 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
If CrCl < 20 Then '= 15mg/kg X 1 & CONSULT KINETICS
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug dosage is 15mg/kg ONE DOSAGE. With ARF Use Kinetic Dosing!" & vbCrLf & vbCrLf
    If ((ABW - IBW) / IBW) >= 0.2 Then
        w = 0.4 * (ABW - IBW)
        w = w + IBW
        w = 15 * w
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The person is Obese (Wt > 20% above IBW).   Therefore 40% of Excess Weight is added to IBW in the calculation." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
    Else
        w = 15 * ABW
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Drug Dosage is: " + Format(w, "######") + "  Round off to the nearest 50 mg." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Check BUN/Creatinine every 2 days and a Random Drug Level 8-14 hours before the next dose." & vbCrLf & vbCrLf
    End If
End If
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The MAXIMUM DOSE of Drug = 1500 MG" & vbCrLf & vbCrLf

End Select
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Standard Initial Administration time is 3pm" & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Patients with high creatinine clearance (> 70ml/min), 2nd dose may be given > 12 hours from first dose, And Then 3pm the Next day." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Patients with low creatinine clearance (< 70ml/min), 2nd dose must be given 24 hours from the first dose."
Load Interpretive
Interpretive.Show
End Sub

Private Sub Command1_Click()
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

Load Interpretive
Interpretive.Show
End Sub

Private Sub FlatButton5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
AminoRx = ""
aCrClag.Text = Format(ActualCrCl, "####.0")
eCrClag.Text = Format(EstimatedCrCl, "####.0")
ActBwag.Text = Format(ABW, "####.0")
estBwag.Text = Format(IBW, "####.0")
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

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        AminoRx = "Gentamycin"
        CF = 0.43
        AminogLD = 2
    Case 1
        CF = 0.58
        AminoRx = "Tobramycin"
        AminogLD = 2
    Case 2
        AminoRx = "Netilmicin"
        CF = 0.5
        AminogLD = 2
    Case 3
        AminoRx = "Amikacin"
        CF = 0.38
        AminogLD = 7.5
    Case 4
        AminoRx = "Kanamycin"
        CF = 1
        AminogLD = 7.5
End Select
End Sub



