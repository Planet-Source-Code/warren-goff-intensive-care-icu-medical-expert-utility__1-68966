VERSION 5.00
Begin VB.Form Nutritional 
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
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   255
      Left            =   600
      TabIndex        =   51
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Ftxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   50
      Top             =   3720
      Width           =   765
   End
   Begin VB.TextBox kNtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   48
      Top             =   4080
      Width           =   765
   End
   Begin VB.TextBox nPtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   47
      Top             =   3720
      Width           =   765
   End
   Begin VB.TextBox gNtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   46
      Top             =   3360
      Width           =   765
   End
   Begin VB.TextBox Wtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   45
      Top             =   4200
      Width           =   765
   End
   Begin VB.TextBox Ptxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   44
      Top             =   3960
      Width           =   765
   End
   Begin VB.TextBox Dtxt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   43
      Top             =   3480
      Width           =   765
   End
   Begin VB.CommandButton CandyButton1 
      Caption         =   "Rx"
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   4680
      Width           =   855
   End
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.TextBox REEcalctxt 
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
      Height          =   360
      Left            =   1200
      TabIndex        =   36
      Top             =   2220
      Width           =   1005
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080FF80&
      Caption         =   "REE"
      Height          =   255
      Index           =   1
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3120
      Width           =   690
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080FF80&
      Caption         =   "BEE"
      Height          =   255
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3120
      Value           =   -1  'True
      Width           =   690
   End
   Begin VB.TextBox Voltxt 
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
      Height          =   360
      Left            =   3810
      TabIndex        =   23
      Text            =   "2500"
      Top             =   2415
      Width           =   1005
   End
   Begin VB.TextBox CHOkcal 
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
      Height          =   360
      Left            =   5205
      TabIndex        =   21
      Text            =   "60"
      Top             =   1965
      Width           =   1005
   End
   Begin VB.TextBox Fatkcal 
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
      Height          =   360
      Left            =   5205
      TabIndex        =   19
      Text            =   "40"
      Top             =   1500
      Width           =   1005
   End
   Begin VB.TextBox REEmeasured 
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
      Height          =   360
      Left            =   1200
      TabIndex        =   17
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "% Protein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   705
      Left            =   2430
      TabIndex        =   10
      Top             =   600
      Width           =   2115
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "10"
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
         Height          =   390
         Index           =   3
         Left            =   1515
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "8.5"
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
         Height          =   390
         Index           =   2
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   255
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "5"
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
         Height          =   390
         Index           =   1
         Left            =   555
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "3.5"
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
         Height          =   390
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   255
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Caption         =   "% Fat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   2430
      TabIndex        =   6
      Top             =   1380
      Width           =   1440
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "30"
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
         Height          =   390
         Index           =   2
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   420
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "20"
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
         Height          =   390
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   420
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "10"
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
         Height          =   390
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Stress Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1155
      Left            =   600
      TabIndex        =   1
      Top             =   585
      Width           =   1635
      Begin VB.TextBox Btxt 
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
         Height          =   360
         Left            =   720
         TabIndex        =   42
         Top             =   720
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "4"
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
         Height          =   405
         Index           =   3
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "3"
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
         Height          =   405
         Index           =   2
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "2"
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
         Height          =   405
         Index           =   1
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   300
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "1"
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
         Height          =   405
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Value           =   -1  'True
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BEE="
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
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   810
         Width           =   585
      End
   End
   Begin VB.CommandButton CandyButton2 
      Caption         =   "Information"
      Height          =   255
      Left            =   3720
      TabIndex        =   40
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
      TabIndex        =   49
      Top             =   4440
      Width           =   495
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
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   6030
      TabIndex        =   41
      Top             =   90
      Width           =   345
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Parenteral Nutrition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   0
      Left            =   735
      TabIndex        =   37
      Top             =   135
      Width           =   4980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cREE="
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
      Height          =   240
      Index           =   4
      Left            =   465
      TabIndex        =   35
      Top             =   2295
      Width           =   720
   End
   Begin VB.Label kcal_N2txt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kcal/N2="
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
      Height          =   240
      Left            =   4395
      TabIndex        =   32
      Top             =   4065
      Width           =   930
   End
   Begin VB.Label gmtxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gm Protein="
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
      Height          =   240
      Left            =   4395
      TabIndex        =   31
      Top             =   3735
      Width           =   1245
   End
   Begin VB.Label gmN2txt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gm N2="
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
      Height          =   240
      Left            =   4395
      TabIndex        =   30
      Top             =   3405
      Width           =   795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rx:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Index           =   1
      Left            =   180
      TabIndex        =   29
      Top             =   3375
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rx:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   180
      TabIndex        =   28
      Top             =   3420
      Width           =   840
   End
   Begin VB.Label Watertxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Water (cc) = "
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
      Left            =   1230
      TabIndex        =   27
      Top             =   4200
      Width           =   1320
   End
   Begin VB.Label Proteintxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protein (cc) = "
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
      Left            =   1200
      TabIndex        =   26
      Top             =   3960
      Width           =   1440
   End
   Begin VB.Label Fattxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fat (cc) = "
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
      Left            =   1200
      TabIndex        =   25
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Dextrosetxt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dextrose (gm) = "
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
      Left            =   1230
      TabIndex        =   24
      Top             =   3480
      Width           =   1710
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Height          =   60
      Left            =   0
      Top             =   2895
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vol(cc)/24h="
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
      Height          =   240
      Index           =   3
      Left            =   2475
      TabIndex        =   22
      Top             =   2505
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHO kcal%="
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
      Height          =   240
      Index           =   2
      Left            =   3870
      TabIndex        =   20
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fat kcal%="
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
      Height          =   240
      Index           =   1
      Left            =   4005
      TabIndex        =   18
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mREE="
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
      Height          =   240
      Index           =   0
      Left            =   465
      TabIndex        =   16
      Top             =   1875
      Width           =   780
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
      Left            =   6000
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Parenteral Nutrition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   750
      TabIndex        =   38
      Top             =   120
      Width           =   4980
   End
End
Attribute VB_Name = "Nutritional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub TPN()
Dim result As Integer
Dim flagler As Boolean
    Stress = 1.2
    ProtFact = 0.45
    Protein = 3.5
REE = Val(Trim(REEmeasured.Text))
Interpretive.Interpret.Text = ""
Interpretive.Interpret.Text = "Nutritional Information: " & vbCrLf & vbCrLf



If REE <> 0 Then
End If

If Sex = 1 Then
    BEE = (Stress * 66.473 + 13.7516 * Wtkg + 5.0033 * Htcm - 6.755 * Age)
    'Males REE = 66.473 + (13.7516 x weight in kg) + (5.0033 x height in cm) – (6.755 x age in years)
Else
    BEE = (Stress * 655.0955 + 9.5634 * Wtkg + 1.8496 * Htcm - 4.6756 * Age)
    'Females REE = 655.0955 + (9.5634 x weight in kg) + (1.8496 x height in cm) – (4.6756 x age in years)
End If
Btxt.Text = Format(BEE, "#####.0") 'BEE
BEE = REE
'Dextrosetxt.Caption = Format(REE * (Val(CHOkcal.Text) / 100) / 4, "#####.0")
Dextrose = BEE * (Val(CHOkcal.Text) / 100)
Dtxt.Text = Format(Dextrose, "#####.0")
'Fattxt.Caption = Format(REE * (Val(Fatkcal.Text) / 100) / (Fat * 9 / 100), "#####.0")
Fat = 10
ccFat = BEE * (Val(Fatkcal.Text) / 100) / (Fat * 9 / 100)
Ftxt.Text = Format(ccFat, "#####.0")

gmProtein = ProtFact * Wtkg
Ptxt.Text = Format(gmProtein, "####.0")
gmN2 = gmProtein * 0.16
gNtxt.Text = Format(gmN2, "###.0")
kcal_N2 = BEE / gmN2
kNtxt.Text = Format(kcal_N2, "####.0")
'kcal_N2txt.Caption = Format(REE / gmN2, "####.0")

'If Stress1.Status = 1 Then
    If kcal_N2 > 100 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "For this Level of Stress the Calorie/N2 Ratio is Elevated!" & vbCrLf & _
         "Be cautious of Overfeeding and it's complications." & vbCrLf & _
         "Correlate Clinically." & vbCrLf & vbCrLf
    End If
'End If
'If Stress2.Status = 1 Or Stress3.Status = 1 Then
    'If kcal_N2 > 150 Then
        'Interpretive.Interpret.Text =  vbCrLf & Interpretive.Interpret.Text &  "For this Level of Stress the Calorie/N2 Ratio is Elevated.   Be cautious of Overfeeding and it's complications.  Correlate Clinically."
    'End If
'End If
'If Stress4.Status = 1 Then
    'If kcal_N2 > 150 Then
        'Interpretive.Interpret.Text =  vbCrLf & Interpretive.Interpret.Text &  "For this Level of Stress the Calorie/N2 Ratio is Elevated but may be needed.   Be cautious of Overfeeding and it's complications.  Correlate Clinically."
    'End If
'End If
        
ccProtein = gmProtein * 100 / Protein

Ptxt.Text = Format(ccProtein, "#####")

If ccFat + ccProtein > TPNvol Then
    flagler = True
    'MsgBox "The 24 hour total Volume was increased to " + Str(ccFat + ccProtein) + vbCrLf & _
           "You may adjust this by changing the %`Fat and/or the %`Protein!"
    Voltxt.Text = Format(ccFat + ccProtein, "#####")
    TPNvol = ccFat + ccProtein  '0
Else
    Wtxt.Text = Format(TPNvol - (ccFat + ccProtein), "#####")
End If

'Fat = 30
'Interpretive.Interpret.Text = vbCrLf & Interpretive.Interpret.Text & "Stress Level 1)Factors Harris-Benedict by 1.2 and Protein of .45 gm/kg/d; 2)1.5 & 1; 3)2 & 1.75 ; 4)2.5 & 2.5"

    Stress = 1.2
    ProtFact = 0.45

    'Stress = 1.5
    'ProtFact = 1

    'Stress = 2
    'ProtFact = 1.75
    'Stress = 2.5
    'ProtFact = 2.5

    'Prot35.Status = 0
    'Prot5.Status = 0
    'Prot85.Status = 0
    'Prot10.Status = 0
    'Protein = Val(PProttxt.Caption)

'Fatkcal = Val(Fatkcaltxt.Caption)
'CHOkcaltxt.Caption = Str(100 - Fatkcal)
'CHOkcal = Val(CHOkcaltxt.Caption)

'CHOkcal = Val(CHOkcaltxt.Caption)
'Fatkcaltxt.Caption = Str(100 - CHOkcal)
'Fatkcal = Val(Fatkcaltxt.Caption)

Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Basal Energy Expenditure is calculated from the Harris Benedict Equation based on sex, height and weight (Adjusted) and a Stress Factor." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The typical 24 Hour Volume of Fluids approximates 1 cc/ kcal.  Correlate Clinically." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "A minimum of 100g/d (400 kcal/day) of Dextrose (CHO) is required for Protein Sparing" & vbCrLf & vbCrLf

Interpretive.Interpret.Text = Interpretive.Interpret.Text & "IV Fat Solutions at 10% & 20% (1 and 2 kcal/cc each) may be given peripherally by vein. 30% (3 kcal/cc) Fat must be given centrally and mixed with TPN.  A minimum of 100-200 kcal/d is needed in fat." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "A Positive Protein (Nitrogen) Balance is related to the level of Stress.   Choose the standard %'s or input one." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Trim(Interpretive.Interpret.Text)

End Sub
Sub TPN1()
Dim result As Integer
Dim flagler As Boolean
REE = Val(Trim(REEmeasured.Text))
Interpretive.Interpret.Text = ""
Interpretive.Interpret.Text = "Nutritional Information: " & vbCrLf & vbCrLf

If Sex = 1 Then
    BEE = (Stress * 66.473 + 13.7516 * Wtkg + 5.0033 * Htcm - 6.755 * Age)
    'Males REE = 66.473 + (13.7516 x weight in kg) + (5.0033 x height in cm) – (6.755 x age in years)
Else
    BEE = (Stress * 655.0955 + 9.5634 * Wtkg + 1.8496 * Htcm - 4.6756 * Age)
    'Females REE = 655.0955 + (9.5634 x weight in kg) + (1.8496 x height in cm) – (4.6756 x age in years)
End If
If REE <> 0 And Option4(0).Value = False Then BEE = REE
Btxt.Text = Format(BEE, "#####.0")  'BEE
'Dextrosetxt.Caption = Format(REE * (Val(CHOkcal.Text) / 100) / 4, "#####.0")
Dextrose = BEE * (Val(CHOkcal.Text) / 100)
Dtxt.Text = Format(Dextrose, "#####.0")
'Fattxt.Caption = Format(REE * (Val(Fatkcal.Text) / 100) / (Fat * 9 / 100), "#####.0")
ccFat = BEE * (Val(Fatkcal.Text) / 100) / (Fat * 9 / 100)
Ftxt.Text = Format(ccFat, "#####.0")

gmProtein = ProtFact * Wtkg
Ptxt.Text = Format(gmProtein, "####.0")
gmN2 = gmProtein * 0.16
gNtxt.Text = Format(gmN2, "###.0")
kcal_N2 = BEE / gmN2
kNtxt.Text = Format(kcal_N2, "####.0")
'kcal_N2txt.Caption = Format(REE / gmN2, "####.0")

If Option1(0).Value = True Then
    If kcal_N2 > 100 Then
        Interpretive.Interpret.Text = Interpretive.Interpret.Text & "For this Level of Stress the Calorie/N2 Ratio is Elevated!" & vbCrLf & _
         "Be cautious of Overfeeding and it's complications." & vbCrLf & _
         "Correlate Clinically." & vbCrLf & vbCrLf
    End If
End If
If Option1(1).Value = True Or Option1(2).Value = True Then
    If kcal_N2 > 150 Then
        Interpretive.Interpret.Text = vbCrLf & Interpretive.Interpret.Text & "For this Level of Stress the Calorie/N2 Ratio is Elevated.   Be cautious of Overfeeding and it's complications.  Correlate Clinically."
    End If
End If
If Option1(3).Value = True Then
    If kcal_N2 > 150 Then
        Interpretive.Interpret.Text = vbCrLf & Interpretive.Interpret.Text & "For this Level of Stress the Calorie/N2 Ratio is Elevated but may be needed.   Be cautious of Overfeeding and it's complications.  Correlate Clinically."
    End If
End If
        
ccProtein = gmProtein * 100 / Protein

Ptxt.Text = Format(ccProtein, "#####")

If ccFat + ccProtein > TPNvol Then
    flagler = True
        SetTopMostWindow Me.hwnd, False

        MsgBox "The 24 hour total Volume was increased to " + Str(ccFat + ccProtein) + vbCrLf & _
           "You may adjust this by changing the %`Fat and/or the %`Protein!"
           
        SetTopMostWindow Me.hwnd, True

    Voltxt.Text = Format(ccFat + ccProtein, "#####")
    TPNvol = ccFat + ccProtein  '0
Else
    Wtxt.Text = Format(TPNvol - (ccFat + ccProtein), "#####")
End If

Interpretive.Interpret.Text = vbCrLf & Interpretive.Interpret.Text & "Stress Level: " & vbCrLf & vbCrLf _
    & "    1)Factors Harris-Benedict by 1.2 and Protein of .45 gm/kg/d" _
        & vbCrLf & vbCrLf _
        & vbCrLf & "     2) 1.5 1" _
        & vbCrLf & "     3) 2  1.75" _
        & vbCrLf & "     4) 2.5 & 2.5"




Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The Basal Energy Expenditure is calculated from the Harris Benedict Equation based on sex, height and weight (Adjusted) and a Stress Factor." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "The typical 24 Hour Volume of Fluids approximates 1 cc/ kcal.  Correlate Clinically." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "A minimum of 100g/d (400 kcal/day) of Dextrose (CHO) is required for Protein Sparing" & vbCrLf & vbCrLf

Interpretive.Interpret.Text = Interpretive.Interpret.Text & "IV Fat Solutions at 10% & 20% (1 and 2 kcal/cc each) may be given peripherally by vein. 30% (3 kcal/cc) Fat must be given centrally and mixed with TPN.  A minimum of 100-200 kcal/d is needed in fat." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & "A Positive Protein (Nitrogen) Balance is related to the level of Stress.   Choose the standard %'s or input one." & vbCrLf & vbCrLf
Interpretive.Interpret.Text = Trim(Interpretive.Interpret.Text)

End Sub

Private Sub CandyButton1_Click()
REE = Val(Trim(REEmeasured.Text))
 TPN1
End Sub

Private Sub CandyButton2_Click()
Load Interpretive
Interpretive.Show
End Sub

Private Sub FlatButton5_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
TPN

End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
REEmeasured.Text = Format(REE, "#####.0")
REEcalctxt.Text = Format(REEcalcu, "#####.0")
If Sex = 1 Then
    BEE = (1.2) * (66.473 + 13.7516 * Wtkg + 5.0033 * Htcm - 6.755 * Age)
    'Males REE = 66.473 + (13.7516 x weight in kg) + (5.0033 x height in cm) – (6.755 x age in years)
Else
    BEE = (1.2) * (655.0955 + 9.5634 * Wtkg + 1.8496 * Htcm - 4.6756 * Age)
    'Females REE = 655.0955 + (9.5634 x weight in kg) + (1.8496 x height in cm) – (4.6756 x age in years)
End If
Btxt = Format(BEE, "#####.0")
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

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        Stress = 1.2
        ProtFact = 0.45
    Case 1
        Stress = 1.5
        ProtFact = 1
    Case 2
        Stress = 2
        ProtFact = 1.75
    Case 3
        Stress = 2.5
        ProtFact = 2.5
End Select

If Sex = 1 Then
    BEE = (Stress) * (66.473 + 13.7516 * Wtkg + 5.0033 * Htcm - 6.755 * Age)
    'Males REE = 66.473 + (13.7516 x weight in kg) + (5.0033 x height in cm) – (6.755 x age in years)
Else
    BEE = (Stress) * (655.0955 + 9.5634 * Wtkg + 1.8496 * Htcm - 4.6756 * Age)
    'Females REE = 655.0955 + (9.5634 x weight in kg) + (1.8496 x height in cm) – (4.6756 x age in years)
End If
TPN1
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
    Case 0
        Fat = 10
    Case 1
        Fat = 20
    Case 2
        Fat = 30
End Select
TPN1

End Sub

Private Sub Option3_Click(Index As Integer)
Select Case Index
    Case 0
        Protein = 3.5
        'PProttxt.Caption = Protein
    Case 1
        Protein = 5
        'PProttxt.Caption = Protein
    Case 2
        Protein = 8.5
        'PProttxt.Caption = Protein
    Case 3
        Protein = 10
        'PProttxt.Caption = Protein
End Select
TPN1
End Sub

Private Sub Option4_Click(Index As Integer)
TPN1
End Sub

Private Sub REEmeasured_Change()
  REE = Val(REEmeasured.Text)
End Sub

Private Sub Text5_Change()
    If Val(REEmeasured.Text) = 0 Then
        REE = Val(Text5.Text)
    End If

End Sub

