VERSION 5.00
Begin VB.Form FluidChal 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Baseline Interpretation"
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CVP Delta"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox PCWP1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox PCWP2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox CVP3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox CVP2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox CVP1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox PCWP3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
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
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "PCWP Delta"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
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
      Index           =   7
      Left            =   5880
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   ''   After 10 minutes:"
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
      Index           =   9
      Left            =   840
      TabIndex        =   15
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CVP During:"
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
      Index           =   5
      Left            =   840
      TabIndex        =   14
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   ''   After 10 minutes:"
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
      Index           =   4
      Left            =   840
      TabIndex        =   13
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   ''   Immediately after:"
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
      Left            =   840
      TabIndex        =   12
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   ''   Immediately after:"
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
      Left            =   840
      TabIndex        =   11
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCWP or PADP During:"
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
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   960
      Width           =   2805
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fluid Challenge Calculations"
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
      Left            =   750
      TabIndex        =   3
      Top             =   120
      Width           =   4635
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fluid Challenge Calculations"
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
      Left            =   720
      TabIndex        =   2
      Top             =   90
      Width           =   4635
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
Attribute VB_Name = "FluidChal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Interpretive.Interpret.Text = ""
    If CVP < 8 Then
       If Val(CVP1.Text) - CVP > 5 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Fluid Challenge of 100 ml over 10 minutes increased the CVP by more then 5 cmH20 (3.7 mmHg) during infusion, STOP THE Infusion Immediately." & vbCrLf & vbCrLf
            Exit Sub
       End If
       If Val(CVP2.Text) - CVP >= 2 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg) and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
       If Val(CVP3.Text) - CVP > 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
               Exit Sub
       End If
       If Val(CVP3.Text) - CVP <= 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
               Exit Sub
        End If
     End If
    If CVP < 14 Then
       If Val(CVP1.Text) - CVP > 5 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Fluid Challenge of 200 ml over 10 minutes increased the CVP by more then 5 cmH20 (3.7 mmHg) during infusion, STOP THE Infusion Immediately." & vbCrLf & vbCrLf
            Exit Sub
       End If
       If Val(CVP2.Text) - CVP >= 2 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg) and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
       If Val(CVP3.Text) - CVP > 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
               Exit Sub
       End If
       If Val(CVP3.Text) - CVP <= 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
               Exit Sub
        End If
    End If
    
    If CVP >= 14 Then
       If CVP >= 18 Then Fluids.Text = "Use caution at this elevated CVP (>= 18 cm H2O)."
       If Val(CVP1.Text) - CVP > 5 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "Fluid Challenge of 50 ml over 10 minutes increased the CVP by more then 5 cmH20 (3.7 mmHg) during infusion, STOP THE Infusion Immediately." & vbCrLf & vbCrLf
            Exit Sub
       End If
       If Val(CVP2.Text) - CVP >= 2 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg) and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
       If Val(CVP3.Text) - CVP > 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
               Exit Sub
       End If
       If Val(CVP3.Text) - CVP <= 2 Then
               Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
               Exit Sub
        End If
    End If
End Sub

Private Sub Command3_Click()
Interpretive.Interpret.Text = ""
If PCWP <> 0 Or PAPdias <> 0 Then
    If PCWP < 12 Or PAPdias < 12 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 200 ml over 10 minutes.   If PCWP or PADP increases > 7 mmHg during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
    If PCWP < 16 Or PAPdias < 16 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 100 ml over 10 minutes.   If PCWP or PADP increases > 7 mmHg during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
    If PCWP >= 16 Or PAPdias >= 16 Then
        If PCWP >= 20 Or PAPdias >= 20 Then Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Use caution at this elevated PCWP."
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 50 ml over 10 minutes.   If PCWP or PADP increases > 7 mmHg during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
Else
    If CVP < 8 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 200 ml over 10 minutes.   If CVP increases > 5 cmH20 (3.7 mmHg) during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg)  and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
    If CVP < 14 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 100 ml over 10 minutes.   If CVP increases > 5 cmH20 (3.7 mmHg) during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg)  and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
End If
    If CVP >= 14 Then
        If CVP >= 18 Then Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Use caution at this elevated CVP."
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "Administer a Fluid Challenge of 50 ml over 10 minutes.   If CVP increases > 5 cmH20 (3.7 mmHg) during infusion, STOP it." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If IMMEDIATELY after the infusion, the change in CVP is greater than or equal to 2 cmH20 (1.45 mmHg)  and less than or equal to 5 cmH20 (3.7 mmHg), wait 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is greater than 2 cmH20 (1.45 mmHg), STOP & wait (reassess)." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If 10 minutes after the infusion, the change in CVP is less than or equal to 2 cmH20 (1.45 mmHg), repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "If the SV has increased, it is probably safe to rechallenge the patient as you are on a favorable Pressure/Volume Curve." & vbCrLf & vbCrLf
        Load Interpretive
        Interpretive.Show
        Exit Sub
    End If
End Sub

Private Sub FlatButton5_Click()
Interpretive.Interpret.Text = ""
    If PCWP < 12 Or PAPdias < 12 Then
        If Val(PCWP1.Text) - PCWP > 7 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "STOP Infusion (200ml) Immediately and observe as preload has increased more than 7 mmHg during the infusion." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP2.Text) - PCWP >= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP > 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP <= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
    End If
    
    If PCWP < 16 Or PAPdias < 16 Then
        If Val(PCWP1.Text) - PCWP > 7 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "STOP Infusion (100ml) Immediately and observe as preload has increased more than 7 mmHg during the infusion." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP2.Text) - PCWP >= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP > 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP <= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & "10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
    End If
    If PCWP >= 16 Or PAPdias >= 16 Then
        If PCWP >= 20 Or PAPdias >= 20 Then _
            Fluids.Text = "Use caution at this elevated PCWP (>= 20mmHg)."
        If Val(PCWP1.Text) - PCWP > 7 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "STOP Infusion (50ml) Immediately and observe as preload has increased more than 7 mmHg during the infusion." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP2.Text) - PCWP >= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "The change in PCWP or PADP is greater than or equal to 3 mmHg and less than or equal to 7, wait 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP > 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "10 minutes after the infusion, the change in PCWP or PADP is greater than 3 mmHg, STOP & wait (reassess)." & vbCrLf & vbCrLf
            Exit Sub
        End If
        If Val(PCWP3.Text) - PCWP <= 3 Then
            Interpretive.Interpret.Text = Interpretive.Interpret.Text & Fluids2.Text + "10 minutes after the infusion, the change in PCWP or PADP is less than or equal to 3 mmHg, repeat the infusion over 10 minutes and reassess." & vbCrLf & vbCrLf
            Exit Sub
        End If
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


Private Sub Label11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
