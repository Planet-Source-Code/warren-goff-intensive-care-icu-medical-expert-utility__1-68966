VERSION 5.00
Begin VB.Form AminoGlyc 
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
      Left            =   4200
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CandyButton1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Left            =   3960
      TabIndex        =   26
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF0000&
      Caption         =   "Neutropenia"
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
      Height          =   855
      Left            =   4560
      TabIndex        =   22
      Top             =   2880
      Width           =   1695
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF0000&
      Caption         =   "Severity"
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
      Height          =   855
      Left            =   4560
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF0000&
         Caption         =   "SE"
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
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF0000&
         Caption         =   "MO"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF0000&
         Caption         =   "MI"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "Hydration"
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
      Height          =   855
      Left            =   4560
      TabIndex        =   14
      Top             =   960
      Width           =   1695
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "I"
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
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "N"
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
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "D"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Height          =   1815
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Adjust Dosage"
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1920
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Estimate Dosage"
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
         TabIndex        =   12
         Top             =   720
         Width           =   1920
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Initial Dosage"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1920
      End
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
      Index           =   2
      Left            =   5880
      TabIndex        =   31
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "(.5 - 1hr):"
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
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Infusion Time:"
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
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aminoglycosides Dosing"
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
      Caption         =   "Aminoglycosides Dosing"
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
Attribute VB_Name = "AminoGlyc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
X = 0 'Peak
Y = 0 'trough
w = 0
z = 0
KDAminog = 0
DIAminog = 0
MDAminog = 0
PE = 0
TR = 0
t = 0
Pk = 0
TA = 0
TB = 0
TH = 0



Dw = IBW + ((ABW - IBW) * CF)
If Dw < 0 Then Dw = IBW
If Age >= 65 Then
    vd = 0.3 * Dw
Else
    vd = 0.27 * Dw
End If
SetTopMostWindow Me.hwnd, False

If Trim(AminoRx) = "" Then 'Gentamycin.Status = 0 And Netilmicin.Status = 0 And Kanamycin.Status = 0 And Tobramycin.Status = 0 And Amikacin.Status = 0 Then
    MsgBox "Please Select an Aminoglycoside Drug."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If
If Option2(1).Value = False And Option2(2).Value = False And Option2(0).Value = False Then
    MsgBox "Please Select a Dosing Protocol."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If AminoVolDistribution = 0 Then    'Decreased.Status = 0 And Normal.Status = 0 And Increased.Status = 0 Then
    MsgBox "Please Select a State of Hydration."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If Trim(AminoDiseaseSev) = "" Then 'Mild.Status = 0 And Moderate.Status = 0 And Severe.Status = 0 Then
    MsgBox "Please Select Disease Severity."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If Trim(Neut) = "" Then 'NeutropeniaYes.Status = 0 And NeutropeniaNo.Status = 0 Then
    MsgBox "Please Specify if Patient is Neutropenic."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
End If

If Val(Text1.Text) = 0 Then     'InfusionTimetxt.Label = "" Then
    MsgBox "Please Select Either a .5 or 1 hr Infusion time."
    SetTopMostWindow Me.hwnd, True
    Exit Sub
Else
    t = Val(Text1.Text)
End If
SetTopMostWindow Me.hwnd, True

    If Option4(0).Value = True And Option1(3).Value = False And Option1(4).Value = False Then
      X = 5
      GoTo NINE60
    End If
    If Option4(0).Value = True Then
       X = 20
    End If
    If Option4(1).Value = True And Option1(3).Value = False And Option1(4).Value = False Then
      X = 7
      GoTo NINE60
    End If
    If Option4(1).Value = True Then
      X = 25
    End If
    If Option4(2).Value = True And Option1(3).Value = False And Option1(4).Value = False Then
      X = 8
      GoTo NINE60
    End If
    If Option4(2).Value = True Then
      X = 30
    End If
NINE60:

    If Option7(0).Value = True And Option1(3).Value = False And Option1(4).Value = False Then
      Y = 1.5
      GoTo NINE68
    End If
    If Option7(0).Value = True Then
      Y = 8
      GoTo NINE68
    End If
    If Option1(3).Value = True Or Option1(4).Value = True Then
      Y = 5
      GoTo NINE68
    End If

Y = 1

NINE68:

    'If Gentamycin.status = 1 Then
    '  KDAminog = (.00285 * crcl) + .015
    'End If
    'If Tobramycin.status = 1 Then
    '  KDAminog = (.0031 * crcl) + .01
    'End If
    'If Amikacin.status = 1 Then
    '   KDAminog = (.0024 * crcl) + .01
    'End If

'w =  LOG(y / x) 'LOG10(y / x)
'DIAminog = ((-1 / KDAminog) * w) + t
'w = 1 - EXP(-1 * KDAminog * DIAminog)
'z = 1 - EXP(-1 * KDAminog * t)
'MDAminog = t* x * AminoVolDistribution *KDAminog * w / z
'PE = (MDAminog / AminoVolDistribution) * z / KDAminog / w
'MDAminog = MDAminog * ABW
'w =  EXP(-1 * KDAminog * (DIAminog - t))
'TR = PE * w
SetTopMostWindow Me.hwnd, False
If ActualCrCl <> 0 Then
    CrCl = ActualCrCl
Else
    If EstimatedCrCl <> 0 Then
        CrCl = EstimatedCrCl
    Else
        MsgBox "Please Input Serum Creatinine and/or Urinary data for a Creatinine Clearance."
        If Trim(DataInput.Creatininetxt) = "" Then
            Unload Me
            DataInput.Show
            DataInput.SSTab1.Tab = 0
            DataInput.SStab2.Tab = 2
        Else
            Load Renal
            Renal.Show
            Exit Sub
        End If
    End If
End If
SetTopMostWindow Me.hwnd, True

kel = 0.01 + (CrCl * 0.0024)
AminogLD = Dw * AminogLD
If Option2(0).Value = True Then
    Load InitialAmino
    InitialAmino.Show
    Exit Sub
End If
If Option2(1).Value = True Then
    Load Estimated
    Estimated.Show
    Exit Sub
End If
If Option2(2).Value = True Then
    Load Adjust
    Adjust.Show
    Exit Sub
End If





End Sub

Private Sub Command1_Click()
Interpretive.Interpret.Text = "General Information for Kinetic Dosing:" & vbCrLf
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

Private Sub FlatButton5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
AminoRx = ""
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

Private Sub Option3_Click(Index As Integer)
Select Case Index
    Case 0
        AminoVolDistribution = 0.2
    Case 1
        AminoVolDistribution = 0.25
    Case 2
        AminoVolDistribution = 0.3
End Select
End Sub

Private Sub Option4_Click(Index As Integer)
Select Case Index
    Case 0
        AminoDiseaseSev = "MI"
    Case 1
        AminoDiseaseSev = "MO"
    Case 2
        AminoDiseaseSev = "SE"
End Select
End Sub

Private Sub Option7_Click(Index As Integer)
Select Case Index
    Case 0
        Neut = "Y"
    Case 1
        Neut = "N"
End Select
End Sub

Private Sub Text1_Change()
    t = Val(Text1.Text)
End Sub
