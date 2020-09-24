VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DataInput 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   360
      Top             =   240
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4890
      Left            =   60
      TabIndex        =   0
      Top             =   315
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8625
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   8388608
      TabCaption(0)   =   "Data"
      TabPicture(0)   =   "DataInput.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SStab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Hemodynamics Interpretation"
      TabPicture(1)   =   "DataInput.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label108(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FlatButton6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FlatButton7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FlatButton5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "InterpretHemo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox InterpretHemo 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1410
         Left            =   420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2610
         Width           =   5445
      End
      Begin VB.Frame Frame5 
         Caption         =   "Derived Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   420
         TabIndex        =   1
         Top             =   450
         Width           =   5460
         Begin VB.TextBox LCWtxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   675
            Width           =   800
         End
         Begin VB.TextBox SVRtxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   675
            Width           =   800
         End
         Begin VB.TextBox VDVTtxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   330
            Width           =   800
         End
         Begin VB.TextBox PVRItxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1350
            Width           =   800
         End
         Begin VB.TextBox O2exttxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1005
            Width           =   800
         End
         Begin VB.TextBox DO2txt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1695
            Width           =   800
         End
         Begin VB.TextBox REEtxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1005
            Width           =   800
         End
         Begin VB.TextBox RVSWItxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1695
            Width           =   800
         End
         Begin VB.TextBox RCWtxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1350
            Width           =   800
         End
         Begin VB.TextBox LVSWitxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1005
            Width           =   800
         End
         Begin VB.TextBox VO2txt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   675
            Width           =   800
         End
         Begin VB.TextBox AVO2dif2txt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   330
            Width           =   800
         End
         Begin VB.TextBox SItxt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   330
            Width           =   800
         End
         Begin VB.TextBox RQ1txt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1350
            Width           =   800
         End
         Begin VB.Label Label94 
            Caption         =   "SVRi"
            Height          =   285
            Left            =   3795
            TabIndex        =   29
            Top             =   690
            Width           =   960
         End
         Begin VB.Label Label93 
            Caption         =   "Vd/Vt c"
            Height          =   285
            Left            =   3795
            TabIndex        =   28
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label91 
            Caption         =   "PVRi"
            Height          =   285
            Left            =   3825
            TabIndex        =   27
            Top             =   1335
            Width           =   960
         End
         Begin VB.Label Label90 
            Caption         =   "O2ext"
            Height          =   285
            Left            =   3825
            TabIndex        =   26
            Top             =   1005
            Width           =   960
         End
         Begin VB.Label Label89 
            Caption         =   "DO2i"
            Height          =   285
            Left            =   2235
            TabIndex        =   25
            Top             =   1710
            Width           =   750
         End
         Begin VB.Label Label88 
            Caption         =   "REE"
            Height          =   285
            Left            =   2235
            TabIndex        =   24
            Top             =   1020
            Width           =   750
         End
         Begin VB.Label Label86 
            Caption         =   "RVSWi"
            Height          =   285
            Left            =   135
            TabIndex        =   23
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label Label82 
            Caption         =   "RCWi"
            Height          =   285
            Left            =   135
            TabIndex        =   22
            Top             =   1365
            Width           =   960
         End
         Begin VB.Label Label81 
            Caption         =   "LVSWi"
            Height          =   285
            Left            =   135
            TabIndex        =   21
            Top             =   1020
            Width           =   750
         End
         Begin VB.Label Label80 
            Caption         =   "VO2i"
            Height          =   285
            Left            =   2235
            TabIndex        =   20
            Top             =   690
            Width           =   750
         End
         Begin VB.Label Label79 
            Caption         =   "AVO2Dif"
            Height          =   285
            Left            =   2235
            TabIndex        =   19
            Top             =   345
            Width           =   750
         End
         Begin VB.Label Label75 
            Caption         =   "LCWi"
            Height          =   285
            Left            =   135
            TabIndex        =   18
            Top             =   690
            Width           =   960
         End
         Begin VB.Label Label71 
            Caption         =   "Stroke Index"
            Height          =   285
            Left            =   135
            TabIndex        =   17
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label66 
            Caption         =   "RQ"
            Height          =   285
            Left            =   2235
            TabIndex        =   16
            Top             =   1365
            Width           =   750
         End
      End
      Begin TabDlg.SSTab SStab2 
         Height          =   4560
         Left            =   -74955
         TabIndex        =   31
         Top             =   315
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   8043
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   -2147483641
         TabCaption(0)   =   "Demographics"
         TabPicture(0)   =   "DataInput.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label108(5)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Image1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Image1(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "FlatButton1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Respiratory"
         TabPicture(1)   =   "DataInput.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label108(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Image2(0)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Image2(1)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "FlatButton2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame2"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Renal"
         TabPicture(2)   =   "DataInput.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label108(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "FlatButton3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame3"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Vent-Hemo"
         TabPicture(3)   =   "DataInput.frx":008C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label108(3)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "FlatButton4"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Frame4"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         Begin VB.Frame Frame1 
            Caption         =   "                      Demographics"
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
            Height          =   2985
            Left            =   555
            TabIndex        =   177
            Top             =   210
            Width           =   5685
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4590
               TabIndex        =   202
               Text            =   "6.8"
               Top             =   1920
               Width           =   600
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4590
               TabIndex        =   201
               Text            =   "80"
               Top             =   1575
               Width           =   600
            End
            Begin VB.TextBox Patmtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4590
               TabIndex        =   200
               Text            =   "760"
               Top             =   2580
               Width           =   600
            End
            Begin VB.TextBox FiO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4590
               TabIndex        =   199
               Text            =   "21"
               Top             =   2235
               Width           =   600
            End
            Begin VB.TextBox WtKgg 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   198
               Top             =   1935
               Width           =   600
            End
            Begin VB.TextBox BSAtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   197
               Top             =   2295
               Width           =   600
            End
            Begin VB.TextBox cTemptxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3570
               Locked          =   -1  'True
               TabIndex        =   196
               Top             =   1260
               Width           =   600
            End
            Begin VB.TextBox ABWtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   195
               Top             =   2595
               Width           =   600
            End
            Begin VB.TextBox IBWtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   2265
               Width           =   600
            End
            Begin VB.TextBox BMItxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2775
               Locked          =   -1  'True
               TabIndex        =   193
               Top             =   2610
               Width           =   600
            End
            Begin VB.TextBox ARtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   192
               Text            =   "100"
               Top             =   1950
               Width           =   600
            End
            Begin VB.TextBox RRtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2745
               TabIndex        =   191
               Text            =   "12"
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox Temptxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   190
               Text            =   "98.6"
               Top             =   1245
               Width           =   600
            End
            Begin VB.TextBox MAPtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4710
               Locked          =   -1  'True
               TabIndex        =   189
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox Diastolictxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3555
               TabIndex        =   188
               Text            =   "80"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox Systolictxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   187
               Text            =   "140"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox Weighttxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   186
               Text            =   "250"
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox Heighttxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   185
               Text            =   "72"
               Top             =   1260
               Width           =   600
            End
            Begin VB.TextBox Agetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   184
               Text            =   "51"
               Top             =   915
               Width           =   600
            End
            Begin VB.OptionButton Female_Status 
               Caption         =   "Female"
               Height          =   225
               Left            =   3270
               TabIndex        =   183
               Top             =   300
               Width           =   900
            End
            Begin VB.OptionButton Male_Status 
               Caption         =   "Male"
               Height          =   225
               Left            =   2550
               TabIndex        =   182
               Top             =   285
               Value           =   -1  'True
               Width           =   765
            End
            Begin VB.TextBox Middletxt 
               Height          =   285
               Left            =   4695
               TabIndex        =   181
               Text            =   "S"
               Top             =   615
               Width           =   240
            End
            Begin VB.TextBox Firsttxt 
               Height          =   285
               Left            =   3375
               TabIndex        =   180
               Text            =   "Warren"
               Top             =   615
               Width           =   1260
            End
            Begin VB.TextBox Lasttxt 
               Height          =   285
               Left            =   2010
               TabIndex        =   179
               Text            =   "Goff"
               Top             =   585
               Width           =   1260
            End
            Begin VB.TextBox Datetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   600
               TabIndex        =   178
               Top             =   255
               Width           =   1260
            End
            Begin VB.Label Label104 
               Caption         =   "ET-CO ppm"
               Height          =   285
               Left            =   3570
               TabIndex        =   225
               Top             =   1950
               Width           =   960
            End
            Begin VB.Label Label103 
               Caption         =   "SpO2"
               Height          =   285
               Left            =   3570
               TabIndex        =   224
               Top             =   1605
               Width           =   750
            End
            Begin VB.Label Label36 
               Caption         =   "pAtmosphere:"
               Height          =   285
               Left            =   3570
               TabIndex        =   223
               Top             =   2610
               Width           =   960
            End
            Begin VB.Label Label37 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "FiO2 (%):"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   3570
               TabIndex        =   222
               ToolTipText     =   "Click to Change FiO2"
               Top             =   2295
               Width           =   750
            End
            Begin VB.Label Label92 
               Caption         =   "            (kg):"
               Height          =   285
               Left            =   45
               TabIndex        =   221
               Top             =   1935
               Width           =   960
            End
            Begin VB.Label Label18 
               Caption         =   "BSA:"
               Height          =   285
               Left            =   1875
               TabIndex        =   220
               Top             =   2325
               Width           =   540
            End
            Begin VB.Label Label17 
               Caption         =   "C"
               Height          =   285
               Left            =   4215
               TabIndex        =   219
               Top             =   1260
               Width           =   150
            End
            Begin VB.Label Label16 
               Caption         =   "F"
               Height          =   285
               Left            =   3390
               TabIndex        =   218
               Top             =   1275
               Width           =   150
            End
            Begin VB.Label Label15 
               Caption         =   "ABW (kg):"
               Height          =   285
               Left            =   75
               TabIndex        =   217
               Top             =   2625
               Width           =   960
            End
            Begin VB.Label Label14 
               Caption         =   "IBW (kg):"
               Height          =   285
               Left            =   75
               TabIndex        =   216
               Top             =   2295
               Width           =   960
            End
            Begin VB.Label Label13 
               Caption         =   "BMI:"
               Height          =   285
               Left            =   1875
               TabIndex        =   215
               Top             =   2640
               Width           =   750
            End
            Begin VB.Label Label12 
               Caption         =   "Pulse:"
               Height          =   285
               Left            =   1875
               TabIndex        =   214
               Top             =   1980
               Width           =   960
            End
            Begin VB.Label Label11 
               Caption         =   "Resp. Rate:"
               Height          =   285
               Left            =   1875
               TabIndex        =   213
               Top             =   1620
               Width           =   960
            End
            Begin VB.Label Label10 
               Caption         =   "Temp (F):"
               Height          =   285
               Left            =   1875
               TabIndex        =   212
               Top             =   1275
               Width           =   750
            End
            Begin VB.Label Label9 
               Caption         =   "MAP:"
               Height          =   285
               Left            =   4245
               TabIndex        =   211
               Top             =   945
               Width           =   420
            End
            Begin VB.Label Label8 
               Caption         =   "/"
               Height          =   285
               Left            =   3390
               TabIndex        =   210
               Top             =   945
               Width           =   150
            End
            Begin VB.Label Label7 
               Caption         =   "BP (S/D):"
               Height          =   285
               Left            =   1875
               TabIndex        =   209
               Top             =   945
               Width           =   750
            End
            Begin VB.Label Label6 
               Caption         =   "Weight (lb):"
               Height          =   285
               Left            =   45
               TabIndex        =   208
               Top             =   1620
               Width           =   960
            End
            Begin VB.Label Label5 
               Caption         =   "Height (in):"
               Height          =   285
               Left            =   45
               TabIndex        =   207
               Top             =   1290
               Width           =   960
            End
            Begin VB.Label Label4 
               Caption         =   "Age (yrs):"
               Height          =   285
               Left            =   45
               TabIndex        =   206
               Top             =   945
               Width           =   750
            End
            Begin VB.Label Label3 
               Caption         =   "Sex:"
               Height          =   285
               Left            =   2025
               TabIndex        =   205
               Top             =   270
               Width           =   420
            End
            Begin VB.Label Label1 
               Caption         =   "Name (Last, First, Middle): "
               Height          =   285
               Left            =   75
               TabIndex        =   204
               Top             =   615
               Width           =   1830
            End
            Begin VB.Label Label2 
               Caption         =   "Date:"
               Height          =   285
               Left            =   75
               TabIndex        =   203
               Top             =   285
               Width           =   465
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00C0C0C0&
               X1              =   30
               X2              =   1440
               Y1              =   112
               Y2              =   112
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "                      Respiratory"
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
            Height          =   2985
            Left            =   -74655
            TabIndex        =   132
            Top             =   225
            Width           =   5475
            Begin VB.TextBox HctTxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4575
               TabIndex        =   247
               Text            =   "38"
               Top             =   2280
               Width           =   600
            End
            Begin VB.TextBox WBCtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4560
               TabIndex        =   245
               Text            =   "15"
               Top             =   2595
               Width           =   600
            End
            Begin VB.TextBox rAG 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2730
               Locked          =   -1  'True
               TabIndex        =   154
               Top             =   1635
               Width           =   600
            End
            Begin VB.TextBox rCO2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   153
               Text            =   "15"
               Top             =   1320
               Width           =   600
            End
            Begin VB.TextBox CaO2txt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4560
               TabIndex        =   152
               Top             =   225
               Width           =   600
            End
            Begin VB.TextBox CvO2txt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4575
               TabIndex        =   151
               Top             =   555
               Width           =   600
            End
            Begin VB.TextBox AVO2diftxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   150
               Top             =   900
               Width           =   600
            End
            Begin VB.TextBox rK 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2715
               TabIndex        =   149
               Text            =   "4"
               Top             =   645
               Width           =   600
            End
            Begin VB.TextBox aAdiftxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   148
               Top             =   1230
               Width           =   600
            End
            Begin VB.TextBox PO2_FiO2txt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4575
               Locked          =   -1  'True
               TabIndex        =   147
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox QsQttxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4575
               Locked          =   -1  'True
               TabIndex        =   146
               Top             =   1935
               Width           =   600
            End
            Begin VB.TextBox SvO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   145
               Text            =   "40"
               Top             =   2640
               Width           =   600
            End
            Begin VB.TextBox HBtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   144
               Text            =   "12"
               Top             =   2325
               Width           =   600
            End
            Begin VB.TextBox rBS 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   143
               Text            =   "120"
               Top             =   2655
               Width           =   600
            End
            Begin VB.TextBox pHtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   142
               Text            =   "7.22"
               Top             =   300
               Width           =   600
            End
            Begin VB.TextBox PaCO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   141
               Text            =   "52"
               Top             =   645
               Width           =   600
            End
            Begin VB.TextBox PaO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   140
               Text            =   "48"
               Top             =   975
               Width           =   600
            End
            Begin VB.TextBox rNa 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2715
               TabIndex        =   139
               Text            =   "140"
               Top             =   315
               Width           =   600
            End
            Begin VB.TextBox rCL 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2715
               TabIndex        =   138
               Text            =   "100"
               Top             =   975
               Width           =   600
            End
            Begin VB.TextBox PvO2txt 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2730
               TabIndex        =   137
               Text            =   "34"
               Top             =   1965
               Width           =   600
            End
            Begin VB.TextBox PvCO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   136
               Text            =   "68"
               Top             =   2325
               Width           =   600
            End
            Begin VB.TextBox EtCO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   135
               Text            =   "48"
               Top             =   1305
               Width           =   600
            End
            Begin VB.TextBox SaO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   134
               Text            =   "84"
               Top             =   1650
               Width           =   600
            End
            Begin VB.TextBox HCO3txt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   133
               Top             =   1980
               Width           =   600
            End
            Begin VB.Label Label35 
               Caption         =   "Hct"
               Height          =   285
               Index           =   2
               Left            =   3735
               TabIndex        =   248
               Top             =   2310
               Width           =   630
            End
            Begin VB.Label Label35 
               Caption         =   "WBC"
               Height          =   285
               Index           =   1
               Left            =   3720
               TabIndex        =   246
               Top             =   2625
               Width           =   630
            End
            Begin VB.Label Label106 
               Caption         =   "AG"
               Height          =   285
               Left            =   1845
               TabIndex        =   176
               Top             =   1665
               Width           =   540
            End
            Begin VB.Label Label105 
               Caption         =   "CO2"
               Height          =   285
               Left            =   1845
               TabIndex        =   175
               Top             =   1365
               Width           =   960
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00C0C0C0&
               X1              =   105
               X2              =   1380
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Label Label41 
               Caption         =   "CaO2"
               Height          =   285
               Left            =   3705
               TabIndex        =   174
               Top             =   255
               Width           =   750
            End
            Begin VB.Label Label40 
               Caption         =   "CvO2"
               Height          =   285
               Left            =   3705
               TabIndex        =   173
               Top             =   585
               Width           =   750
            End
            Begin VB.Label Label39 
               Caption         =   "AV O2 dif"
               Height          =   285
               Left            =   3705
               TabIndex        =   172
               Top             =   930
               Width           =   750
            End
            Begin VB.Label Label38 
               Caption         =   "Potassium"
               Height          =   285
               Left            =   1830
               TabIndex        =   171
               Top             =   675
               Width           =   750
            End
            Begin VB.Label Label29 
               Caption         =   "a-A grad"
               Height          =   285
               Left            =   3690
               TabIndex        =   170
               Top             =   1260
               Width           =   960
            End
            Begin VB.Label Label28 
               Caption         =   "PaO2/FiO2"
               Height          =   285
               Left            =   3690
               TabIndex        =   169
               Top             =   1620
               Width           =   960
            End
            Begin VB.Label Label21 
               Caption         =   "Qs/Qt"
               Height          =   285
               Left            =   3705
               TabIndex        =   168
               Top             =   1965
               Width           =   540
            End
            Begin VB.Label Label20 
               Caption         =   "SvO2"
               Height          =   285
               Left            =   1890
               TabIndex        =   167
               Top             =   2670
               Width           =   540
            End
            Begin VB.Label Label35 
               Caption         =   "Hb"
               Height          =   285
               Index           =   0
               Left            =   45
               TabIndex        =   166
               Top             =   2355
               Width           =   750
            End
            Begin VB.Label Label34 
               Caption         =   "Glucose"
               Height          =   285
               Left            =   45
               TabIndex        =   165
               Top             =   2640
               Width           =   750
            End
            Begin VB.Label Label33 
               Caption         =   "pH"
               Height          =   285
               Left            =   45
               TabIndex        =   164
               Top             =   330
               Width           =   750
            End
            Begin VB.Label Pa 
               Caption         =   "PaCO2"
               Height          =   285
               Left            =   45
               TabIndex        =   163
               Top             =   675
               Width           =   960
            End
            Begin VB.Label Label31 
               Caption         =   "PaO2"
               Height          =   285
               Left            =   45
               TabIndex        =   162
               Top             =   1005
               Width           =   960
            End
            Begin VB.Label Label30 
               Caption         =   "Sodium"
               Height          =   285
               Left            =   1860
               TabIndex        =   161
               Top             =   315
               Width           =   750
            End
            Begin VB.Label Label27 
               Caption         =   "Choride"
               Height          =   285
               Left            =   1845
               TabIndex        =   160
               Top             =   1020
               Width           =   750
            End
            Begin VB.Label Label26 
               Caption         =   "PvO2"
               Height          =   285
               Left            =   1860
               TabIndex        =   159
               Top             =   1995
               Width           =   960
            End
            Begin VB.Label Label25 
               Caption         =   "PvCO2"
               Height          =   285
               Left            =   1860
               TabIndex        =   158
               Top             =   2355
               Width           =   960
            End
            Begin VB.Label Label24 
               Caption         =   "EtCO2"
               Height          =   285
               Left            =   75
               TabIndex        =   157
               Top             =   1335
               Width           =   750
            End
            Begin VB.Label Label23 
               Caption         =   "SaO2"
               Height          =   285
               Left            =   75
               TabIndex        =   156
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label22 
               Caption         =   "HCO3"
               Height          =   285
               Left            =   45
               TabIndex        =   155
               Top             =   2010
               Width           =   960
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "                             Renal"
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
            Height          =   3045
            Left            =   -74685
            TabIndex        =   83
            Top             =   195
            Width           =   5475
            Begin VB.TextBox COPtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4620
               Locked          =   -1  'True
               TabIndex        =   107
               Top             =   2655
               Width           =   600
            End
            Begin VB.TextBox TPtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4620
               TabIndex        =   106
               Text            =   "6.5"
               Top             =   2325
               Width           =   600
            End
            Begin VB.TextBox Lipidstxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4635
               TabIndex        =   105
               Text            =   "200"
               Top             =   1635
               Width           =   600
            End
            Begin VB.TextBox uChloridetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4635
               TabIndex        =   104
               Text            =   "134"
               Top             =   1275
               Width           =   600
            End
            Begin VB.TextBox Usodiumtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2745
               Locked          =   -1  'True
               TabIndex        =   103
               Text            =   "150"
               Top             =   1950
               Width           =   600
            End
            Begin VB.TextBox VolumeUrinetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   102
               Text            =   "250"
               Top             =   2595
               Width           =   600
            End
            Begin VB.TextBox Creatininetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   101
               Text            =   "2.2"
               Top             =   2265
               Width           =   600
            End
            Begin VB.TextBox BStxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   100
               Text            =   "120"
               Top             =   1575
               Width           =   600
            End
            Begin VB.TextBox UCreatininetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2745
               TabIndex        =   99
               Text            =   "23"
               Top             =   1605
               Width           =   600
            End
            Begin VB.TextBox Urate 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   98
               Text            =   "10"
               Top             =   1245
               Width           =   600
            End
            Begin VB.TextBox Albumintxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4635
               TabIndex        =   97
               Text            =   "4.5"
               Top             =   1995
               Width           =   600
            End
            Begin VB.TextBox Magn 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   96
               Text            =   "1.8"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox BUNtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   95
               Text            =   "32"
               Top             =   1935
               Width           =   600
            End
            Begin VB.TextBox CO2txt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   94
               Text            =   "15"
               Top             =   1245
               Width           =   600
            End
            Begin VB.TextBox CLtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   93
               Text            =   "100"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox Phos 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2745
               TabIndex        =   92
               Text            =   "2.4"
               Top             =   570
               Width           =   600
            End
            Begin VB.TextBox Calc 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2730
               TabIndex        =   91
               Text            =   "7.2"
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox Ktxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   90
               Text            =   "4"
               Top             =   585
               Width           =   600
            End
            Begin VB.TextBox Natxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1065
               TabIndex        =   89
               Text            =   "140"
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox TimedUrinetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   88
               Text            =   "2"
               Top             =   2640
               Width           =   600
            End
            Begin VB.TextBox Uosml 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4620
               TabIndex        =   87
               Text            =   "723"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox CalcSosmtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4605
               Locked          =   -1  'True
               TabIndex        =   86
               Top             =   585
               Width           =   600
            End
            Begin VB.TextBox Sosmtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4620
               TabIndex        =   85
               Text            =   "285"
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox uPotassiumtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2745
               TabIndex        =   84
               Text            =   "32"
               Top             =   2295
               Width           =   600
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00C0C0C0&
               X1              =   105
               X2              =   1830
               Y1              =   90
               Y2              =   90
            End
            Begin VB.Label Label52 
               Caption         =   "Ca"
               Height          =   285
               Left            =   1875
               TabIndex        =   131
               Top             =   270
               Width           =   750
            End
            Begin VB.Label Label102 
               Caption         =   "COP"
               Height          =   285
               Left            =   3660
               TabIndex        =   130
               Top             =   2685
               Width           =   945
            End
            Begin VB.Label Label101 
               Caption         =   "Total Prot."
               Height          =   285
               Left            =   3660
               TabIndex        =   129
               Top             =   2355
               Width           =   945
            End
            Begin VB.Label Label47 
               Caption         =   "Triglycerides"
               Height          =   285
               Left            =   3675
               TabIndex        =   128
               Top             =   1665
               Width           =   945
            End
            Begin VB.Label Label46 
               Caption         =   "uCL"
               Height          =   285
               Left            =   3765
               TabIndex        =   127
               Top             =   1305
               Width           =   540
            End
            Begin VB.Label Label64 
               Caption         =   "uNa"
               Height          =   285
               Left            =   1875
               TabIndex        =   126
               Top             =   1980
               Width           =   540
            End
            Begin VB.Label Label63 
               Caption         =   "Urine Vol cc"
               Height          =   285
               Left            =   75
               TabIndex        =   125
               Top             =   2625
               Width           =   960
            End
            Begin VB.Label Label62 
               Caption         =   "Creatinine"
               Height          =   285
               Left            =   75
               TabIndex        =   124
               Top             =   2310
               Width           =   960
            End
            Begin VB.Label Label61 
               Caption         =   "Glucose"
               Height          =   285
               Left            =   75
               TabIndex        =   123
               Top             =   1605
               Width           =   750
            End
            Begin VB.Label Label60 
               Caption         =   "uCreat"
               Height          =   285
               Left            =   1860
               TabIndex        =   122
               Top             =   1635
               Width           =   960
            End
            Begin VB.Label Label59 
               Caption         =   "Uric Acid"
               Height          =   285
               Left            =   1860
               TabIndex        =   121
               Top             =   1275
               Width           =   960
            End
            Begin VB.Label Label58 
               Caption         =   "Albumin"
               Height          =   285
               Left            =   3660
               TabIndex        =   120
               Top             =   2010
               Width           =   750
            End
            Begin VB.Label Label57 
               Caption         =   "Mg"
               Height          =   285
               Left            =   1875
               TabIndex        =   119
               Top             =   945
               Width           =   750
            End
            Begin VB.Label Lbl 
               Caption         =   "BUN"
               Height          =   285
               Left            =   45
               TabIndex        =   118
               Top             =   1965
               Width           =   960
            End
            Begin VB.Label Label55 
               Caption         =   "CO2"
               Height          =   285
               Left            =   45
               TabIndex        =   117
               Top             =   1290
               Width           =   960
            End
            Begin VB.Label Label54 
               Caption         =   "CL"
               Height          =   285
               Left            =   45
               TabIndex        =   116
               Top             =   930
               Width           =   750
            End
            Begin VB.Label Label53 
               Caption         =   "PO4"
               Height          =   285
               Left            =   1875
               TabIndex        =   115
               Top             =   600
               Width           =   750
            End
            Begin VB.Label Label51 
               Caption         =   "K"
               Height          =   285
               Left            =   45
               TabIndex        =   114
               Top             =   615
               Width           =   960
            End
            Begin VB.Label Label50 
               Caption         =   "Na"
               Height          =   285
               Left            =   45
               TabIndex        =   113
               Top             =   270
               Width           =   750
            End
            Begin VB.Label Label49 
               Caption         =   "Time Urine"
               Height          =   285
               Left            =   1890
               TabIndex        =   112
               Top             =   2670
               Width           =   855
            End
            Begin VB.Label Label45 
               Caption         =   "uOsm"
               Height          =   285
               Left            =   3750
               TabIndex        =   111
               Top             =   945
               Width           =   750
            End
            Begin VB.Label Label44 
               Caption         =   "calc Sosm"
               Height          =   285
               Left            =   3750
               TabIndex        =   110
               Top             =   615
               Width           =   750
            End
            Begin VB.Label Label43 
               Caption         =   "Sosm"
               Height          =   285
               Left            =   3750
               TabIndex        =   109
               Top             =   270
               Width           =   750
            End
            Begin VB.Label Label42 
               Caption         =   "uK"
               Height          =   285
               Left            =   1890
               TabIndex        =   108
               Top             =   2325
               Width           =   750
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ventilator                                  Hemodynamics"
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
            Height          =   3300
            Left            =   -74745
            TabIndex        =   32
            Top             =   0
            Width           =   5910
            Begin VB.TextBox REEcalctxt 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4785
               TabIndex        =   243
               Text            =   "2000"
               Top             =   2760
               Width           =   600
            End
            Begin VB.TextBox REEmeasured 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4785
               TabIndex        =   226
               Text            =   "2000"
               Top             =   2475
               Width           =   600
            End
            Begin VB.TextBox RQtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4785
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   2160
               Width           =   600
            End
            Begin VB.TextBox CItxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4785
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   225
               Width           =   600
            End
            Begin VB.TextBox FickCItxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4785
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   585
               Width           =   600
            End
            Begin VB.TextBox COP_PCWPtxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   4785
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   1815
               Width           =   600
            End
            Begin VB.TextBox RAs 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4785
               TabIndex        =   53
               Text            =   "12"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox RAd 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4785
               TabIndex        =   52
               Text            =   "2"
               Top             =   1245
               Width           =   600
            End
            Begin VB.TextBox VO2itxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0E0FF&
               Height          =   285
               Left            =   4785
               TabIndex        =   51
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox SV 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   50
               Text            =   "70"
               Top             =   2895
               Width           =   600
            End
            Begin VB.TextBox COtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   49
               Tag             =   "5.5"
               Text            =   "8"
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox PCWPtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   48
               Tag             =   "22"
               Text            =   "16"
               Top             =   600
               Width           =   600
            End
            Begin VB.TextBox RVdias 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   47
               Text            =   "0"
               Top             =   2625
               Width           =   600
            End
            Begin VB.TextBox CVPtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   46
               Text            =   "15"
               Top             =   930
               Width           =   600
            End
            Begin VB.TextBox PAPsystxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   45
               Text            =   "35"
               Top             =   1260
               Width           =   600
            End
            Begin VB.TextBox PAPdiastxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               TabIndex        =   44
               Text            =   "14"
               Top             =   1605
               Width           =   600
            End
            Begin VB.TextBox PAPmeantxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   1950
               Width           =   600
            End
            Begin VB.TextBox RVs 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2985
               Locked          =   -1  'True
               TabIndex        =   42
               Text            =   "33"
               Top             =   2280
               Width           =   600
            End
            Begin VB.TextBox PCWPitxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               Locked          =   -1  'True
               TabIndex        =   41
               Text            =   "26"
               Top             =   2265
               Width           =   600
            End
            Begin VB.TextBox PEEPitxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1125
               TabIndex        =   40
               Text            =   "10"
               Top             =   1935
               Width           =   600
            End
            Begin VB.TextBox PEEPetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   39
               Text            =   "10"
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox IFRtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   38
               Text            =   "150"
               Top             =   1245
               Width           =   600
            End
            Begin VB.TextBox RespRateTxt 
               Height          =   285
               Left            =   1095
               TabIndex        =   37
               Text            =   "12"
               Top             =   2850
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.TextBox VTtxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   36
               Text            =   "700"
               Top             =   915
               Width           =   600
            End
            Begin VB.TextBox PCWPetxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   35
               Text            =   "16"
               Top             =   2610
               Width           =   600
            End
            Begin VB.TextBox PlatPrestxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   34
               Text            =   "40"
               Top             =   585
               Width           =   600
            End
            Begin VB.TextBox PAPrestxt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1110
               TabIndex        =   33
               Text            =   "80"
               Top             =   225
               Width           =   600
            End
            Begin VB.Label Label107 
               Caption         =   "REE calc"
               Height          =   285
               Index           =   3
               Left            =   3720
               TabIndex        =   244
               Top             =   2790
               Width           =   960
            End
            Begin VB.Label Label107 
               Caption         =   "REE meas"
               Height          =   285
               Index           =   1
               Left            =   3720
               TabIndex        =   227
               Top             =   2505
               Width           =   960
            End
            Begin VB.Label Label70 
               Caption         =   "RQ"
               Height          =   285
               Left            =   3720
               TabIndex        =   82
               Top             =   2190
               Width           =   960
            End
            Begin VB.Label Label98 
               Caption         =   "COP-PCWPe"
               Height          =   285
               Left            =   3720
               TabIndex        =   81
               Top             =   1830
               Width           =   960
            End
            Begin VB.Label Label100 
               Caption         =   "CI"
               Height          =   285
               Left            =   3720
               TabIndex        =   80
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label Label99 
               Caption         =   "Fick CI"
               Height          =   285
               Left            =   3720
               TabIndex        =   79
               Top             =   615
               Width           =   960
            End
            Begin VB.Label Label97 
               Caption         =   "Ra syst"
               Height          =   285
               Left            =   3720
               TabIndex        =   78
               Top             =   945
               Width           =   990
            End
            Begin VB.Label Label96 
               Caption         =   "RA diast"
               Height          =   285
               Left            =   3720
               TabIndex        =   77
               Top             =   1290
               Width           =   960
            End
            Begin VB.Label Label95 
               Caption         =   "VO2i"
               Height          =   285
               Left            =   3720
               TabIndex        =   76
               Top             =   1620
               Width           =   750
            End
            Begin VB.Label Label87 
               Caption         =   "SV"
               Height          =   285
               Left            =   1920
               TabIndex        =   75
               Top             =   2925
               Width           =   960
            End
            Begin VB.Label Label69 
               Caption         =   "Cardiac Out"
               Height          =   285
               Left            =   1920
               TabIndex        =   74
               Top             =   285
               Width           =   1065
            End
            Begin VB.Label Label68 
               Caption         =   "PCWPe"
               Height          =   285
               Left            =   1920
               TabIndex        =   73
               Top             =   630
               Width           =   960
            End
            Begin VB.Label Label67 
               Caption         =   "RV diast"
               Height          =   285
               Left            =   1920
               TabIndex        =   72
               Top             =   2640
               Width           =   825
            End
            Begin VB.Label Label65 
               Caption         =   "CVP"
               Height          =   285
               Left            =   1920
               TabIndex        =   71
               Top             =   960
               Width           =   990
            End
            Begin VB.Label Label56 
               Caption         =   "PAPsyst"
               Height          =   285
               Left            =   1920
               TabIndex        =   70
               Top             =   1305
               Width           =   960
            End
            Begin VB.Label Label48 
               Caption         =   "PAPdiast"
               Height          =   285
               Left            =   1920
               TabIndex        =   69
               Top             =   1635
               Width           =   750
            End
            Begin VB.Label Label32 
               Caption         =   "PAPmean"
               Height          =   285
               Left            =   1920
               TabIndex        =   68
               Top             =   1980
               Width           =   960
            End
            Begin VB.Label Label19 
               Caption         =   "RVsyst"
               Height          =   285
               Left            =   1920
               TabIndex        =   67
               Top             =   2310
               Width           =   960
            End
            Begin VB.Line Line3 
               X1              =   1125
               X2              =   3060
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               X1              =   1830
               X2              =   1830
               Y1              =   105
               Y2              =   3255
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   1830
               X2              =   1830
               Y1              =   120
               Y2              =   3285
            End
            Begin VB.Label Label85 
               Caption         =   "Insp PCWP"
               Height          =   285
               Left            =   45
               TabIndex        =   66
               Top             =   2295
               Width           =   960
            End
            Begin VB.Label Label84 
               Caption         =   "Intr PEEP"
               Height          =   285
               Left            =   45
               TabIndex        =   65
               Top             =   1965
               Width           =   960
            End
            Begin VB.Label Label83 
               Caption         =   "Ext PEEP"
               Height          =   285
               Left            =   45
               TabIndex        =   64
               Top             =   1620
               Width           =   750
            End
            Begin VB.Label Label78 
               Caption         =   "IFR"
               Height          =   285
               Left            =   45
               TabIndex        =   63
               Top             =   1290
               Width           =   960
            End
            Begin VB.Label Label77 
               Caption         =   "RR"
               Height          =   285
               Left            =   45
               TabIndex        =   62
               Top             =   2880
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label Label76 
               Caption         =   "Tidal Volume"
               Height          =   285
               Left            =   45
               TabIndex        =   61
               Top             =   945
               Width           =   990
            End
            Begin VB.Label Label74 
               Caption         =   "Exp PCWP"
               Height          =   285
               Left            =   45
               TabIndex        =   60
               Top             =   2625
               Width           =   825
            End
            Begin VB.Label Label73 
               Caption         =   "Plateau Pres"
               Height          =   285
               Left            =   45
               TabIndex        =   59
               Top             =   615
               Width           =   960
            End
            Begin VB.Label Label72 
               Caption         =   "Peak Pressure"
               Height          =   285
               Left            =   45
               TabIndex        =   58
               Top             =   270
               Width           =   1065
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00C0C0C0&
               X1              =   1110
               X2              =   3105
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.CommandButton FlatButton2 
            Caption         =   "Re-Calculate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -72600
            TabIndex        =   231
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CommandButton FlatButton1 
            Caption         =   "Re-Calculate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   232
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CommandButton FlatButton3 
            Caption         =   "Re-Calculate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -72600
            TabIndex        =   233
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CommandButton FlatButton4 
            Caption         =   "Re-Calculate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -72600
            TabIndex        =   234
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Image Image1 
            Height          =   495
            Index           =   1
            Left            =   5385
            Picture         =   "DataInput.frx":00A8
            Stretch         =   -1  'True
            Top             =   3270
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   495
            Index           =   0
            Left            =   5400
            Picture         =   "DataInput.frx":0972
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   375
         End
         Begin VB.Image Image2 
            Height          =   480
            Index           =   1
            Left            =   -69525
            Picture         =   "DataInput.frx":123C
            Top             =   3210
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Index           =   0
            Left            =   -69480
            Picture         =   "DataInput.frx":1B06
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label Label108 
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
            Height          =   495
            Index           =   5
            Left            =   5760
            TabIndex        =   242
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label108 
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
            Height          =   495
            Index           =   3
            Left            =   -69240
            TabIndex        =   240
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label108 
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
            Height          =   495
            Index           =   2
            Left            =   -69240
            TabIndex        =   239
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label108 
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
            Height          =   495
            Index           =   1
            Left            =   -69240
            TabIndex        =   238
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label108 
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
            Height          =   495
            Index           =   0
            Left            =   -69240
            TabIndex        =   237
            Top             =   3720
            Width           =   495
         End
      End
      Begin VB.CommandButton FlatButton5 
         Caption         =   "Interpetation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   235
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CommandButton FlatButton7 
         Caption         =   "Caveats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   230
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton FlatButton6 
         Caption         =   "Technique"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   236
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label108 
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
         Height          =   495
         Index           =   4
         Left            =   5880
         TabIndex        =   241
         Top             =   4320
         Width           =   495
      End
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
      Left            =   6030
      TabIndex        =   228
      Top             =   30
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
      Index           =   2
      Left            =   6000
      TabIndex        =   229
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "DataInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim Picker As Integer



Private Sub Command1_Click()

End Sub

Private Sub CandyButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Interpretive



Exit Sub

    If CandyButton1.Checked = True Then
        CandyButton1.Checked = False
        SSTab1.Visible = False
    Else
        CandyButton1.Checked = True
        SSTab1.Visible = True
        Unload Interpretive
    End If
End Sub

Private Sub CandyButton2_Click()

End Sub

Private Sub CandyButton3_Click()

'Interpretive.Show
'ABGInterpret
End Sub

Private Sub FlatButton1_Click()
'On Error Resume Next
BSAtxt.Text = ""
Dim X As Single
Dim Y As Single
Dim z As Single
Wtkg = Val(Weighttxt.Text) / 2.2046
WtKgg.Text = Wtkg
Weight = Val(Weighttxt.Text)
X = (Wtkg ^ 0.425)
Htcm = Val(Heighttxt.Text) * 2.54
Heightt = Val(Heighttxt.Text)
Y = (Htcm ^ 0.725)
z = ((Htcm / 100) ^ 2)

If Weight <> 0 And Heightt <> 0 Then
    BSAtxt.Text = Format((X * Y * 7.184) / 1000, "#.00000")
    BSA = Val(BSAtxt.Text)
    BMItxt.Text = Format(Wtkg / z, "###.0")
    BMI = Val(BMItxt.Text)
    If Sex = 1 Then
        If Heightt <= 60 Then
            IBWtxt.Text = Format(106 / 2.2046, "####.0")
            IBW = Val(IBWtxt.Text)
        Else
            IBWtxt.Text = Format((106 + (6 * (Heightt - 60))) / 2.2046, "####.0")
            IBW = Val(IBWtxt.Text)
        End If
    Else
        If Heightt <= 60 Then
            IBWtxt.Text = Format(100 / 2.2046, "####.0")
            IBW = Val(IBWtxt.Text)
        Else
            IBWtxt.Text = Format((100 + (5 * (Heightt - 60))) / 2.2046, "####.0")
            IBW = Val(IBWtxt.Text)
        End If
    End If
    ABWtxt.Text = Format((Wtkg - IBW) * 0.25 + IBW, "###.0")
    ABW = Val(ABWtxt.Text)
End If
If Val(Diastolictxt.Text) >= Val(Systolictxt.Text) Then
    Diastolictxt.Text = ""
    Systolictxt.Text = ""
End If
If Val(Diastolictxt.Text) <> 0 And Val(Systolictxt.Text) <> 0 Then
    MAPtxt.Text = Format(Val(Diastolictxt.Text) + (Val(Systolictxt.Text) - Val(Diastolictxt.Text)) / 3, "###.0")
End If
cTemptxt.Text = Format((Val(Temptxt.Text) - 32) * (5 / 9), "###.0")
End Sub

Private Sub FlatButton2_Click()
w = Val(Patmtxt.Text) - 47

X = Val(FiO2txt.Text) / 100

Y = Val(PaCO2txt.Text) / 0.8

z = Val(PaO2txt.Text)

If w > 0 And X <> 0 And Y <> 0 And z <> 0 Then
    aAdiftxt.Text = Format((w * X - Y) - z, "###.0")
    Aadif = Val(aAdiftxt.Text)   '(w * x -y) - z
    If Val(SvO2txt.Text) <> 0 And Val(PvO2txt.Text) <> 0 Then
        CcO2 = Val(HBtxt.Text) * 0.0137 + 0.003 * (w * X - Y)
        CaO2 = Val(HBtxt.Text) * 0.0137 * Val(SaO2txt.Text) + 0.003 * Val(PaO2txt.Text)
        CvO2 = Val(HBtxt.Text) * 0.0137 * Val(SvO2txt.Text) + 0.003 * Val(PvO2txt.Text)
        w1 = (CcO2 - CaO2) / (CcO2 - CvO2)
        QsQttxt.Text = Format(w1, "###.00")
        'QsQttxt.Text= Format(val(aAdiftxt.Text)/20,"###.00")
        Qs_Qt = Val(QsQttxt.Text)
    Else
        QsQttxt.Text = Format(Val(aAdiftxt.Text) / 20, "###.00")
        Qs_Qt = Val(QsQttxt.Text)
    End If
End If

If X = 1 And aAdiftxt.Text <> "" And HBtxt.Text <> "" And SaO2txt.Text <> "" And PvO2txt.Text <> "" And SvO2txt.Text <> "" Then
Y = Val(aAdiftxt.Text) + (Val(HBtxt.Text) * 432 * (1 - Val(SaO2txt.Text) / 100))
z = z - Val(PvO2txt.Text)
z = z + Val(HBtxt.Text) * 432 * (1 - Val(SvO2txt.Text) / 100)
QsQttxt.Text = Format((Y / z) * 100, "###.0")
Qs_Qt = Val(QsQttxt.Text)
End If

If Val(PaO2txt.Text) <> 0 And Val(FiO2txt.Text) <> 0 Then
  PO2_FiO2txt.Text = Format(Val(PaO2txt.Text) * 100 / Val(FiO2txt.Text), "###.0")
End If

If Val(HBtxt.Text) <> 0 And Val(SaO2txt.Text) <> 0 And Val(PaO2txt.Text) <> 0 Then
  CaO2txt.Text = Format(Val(HBtxt.Text) * 0.0137 * Val(SaO2txt.Text) + 0.003 * Val(PaO2txt.Text), "##.00")
End If

'HCO3
'pH=6.1+log([HCO3-]/pCO2*(3.01*10^-2))

Y = Val(pHtxt.Text) - 6.1
X = (10 ^ Y) * Val(PaCO2txt.Text) * 0.0301

If X <> 0 Then
    HCO3txt.Text = Format(X, "##.0")
End If

'CvO2 and AVO2
If Val(HBtxt.Text) <> 0 And Val(SvO2txt.Text) <> 0 And Val(PvO2txt.Text) <> 0 Then
  CvO2txt.Text = Format(Val(HBtxt.Text) * 0.0137 * Val(SvO2txt.Text) + 0.003 * Val(PvO2txt.Text), "##.00")
End If

If Val(SvO2txt.Text) <> 0 And Val(CvO2txt.Text) <> 0 Then
    AVO2diftxt.Text = Format(Val(CaO2txt.Text) - Val(CvO2txt.Text), "###.00")
End If
rAG.Text = Format(Val(rNa) - (Val(rCO2) + Val(rCL)), "##.0")
FiO2 = Val(FiO2txt.Text)
Patm = Val(Patmtxt.Text)
Hb = Val(HBtxt.Text)
PaO2 = Val(PaO2txt.Text)
PaCO2 = Val(PaCO2txt.Text)
SaO2 = Val(SaO2txt.Text)
PaO2FiO2 = Val(Format(Val(PaO2txt.Text) * 100 / Val(FiO2txt.Text), "###.0"))
Qs_Qt = Val(QsQttxt.Text)
HCO3 = Val(HCO3txt.Text)
AVO2dif = Val(AVO2diftxt.Text)
CaO2 = Val(CaO2txt.Text)
CvO2 = Val(CvO2txt.Text)
PvO2 = Val(PvO2txt.Text)
PvCO2 = Val(PvCO2txt.Text)
SvO2 = Val(SvO2txt.Text)
pH = Val(pHtxt.Text)
End Sub

Private Sub FlatButton3_Click()
If Val(BStxt.Text) <> 0 And Val(BUNtxt.Text) <> 0 And Val(Natxt.Text) <> 0 Then
    CalcSosmtxt.Text = Format(2 * Val(Natxt.Text) + Val(BStxt.Text) / 18 + Val(BUNtxt.Text) / 2.8, "####.0")
    CalcSosm = Val(CalcSosmtxt.Text)
End If
TP = Val(TPtxt.Text)
If TP <> 0 Then
    X = (TP ^ 2)
    Y = (TP ^ 3)
    COPtxt.Text = Format(2.1 * TP + 0.16 * X + 0.009 * Y, "###.00")
    COP = Val(COPtxt.Text)
End If
End Sub

Private Sub FlatButton4_Click()

If Val(COtxt.Text) <> 0 And Val(BSAtxt.Text) <> 0 Then
    CO = Val(COtxt.Text)
    BSA = Val(BSAtxt.Text)
    X = CO / BSA
    CItxt.Text = Int(X * 1000) / 1000
    CI = X
End If
If Val(COPtxt.Text) <> 0 And Val(PCWPtxt.Text) <> 0 Then
    COP_PCWPtxt.Text = Val(COPtxt.Text) - Val(PCWPtxt.Text)
End If
If VO2itxt.Text <> "" Then
    VO2iMeasured = Val(VO2itxt.Text)
    VO2i = Val(Format(Val(AVO2diftxt.Text) * Val(CItxt.Text) * 10, "####.0"))
    RQtxt.Text = Format((Val(PvCO2txt.Text) - Val(PaCO2txt.Text)) / Val(AVO2diftxt.Text), "##.00")
    RQ = Val(RQtxt.Text)
    REEcalcu = VO2i * BSA * 1.44 * 5
    REEcalctxt.Text = Format(REEcalcu, "#####.0")
Else
    VO2iMeasured = 0
    VO2itxt.Text = Format(Val(AVO2diftxt.Text) * Val(CItxt.Text) * 10, "####.0")
    VO2i = Val(VO2itxt.Text)
    RQtxt.Text = Format((Val(PvCO2txt.Text) - Val(PaCO2txt.Text)) / Val(AVO2diftxt.Text), "##.00")
    RQ = Val(RQtxt.Text)
End If
Wtkg = Val(Weighttxt.Text) / 2.2046
FickCI = (Val(VO2itxt.Text) / Val(AVO2diftxt.Text)) / 10
FickCItxt.Text = FickCI
PAPmeantxt.Text = Format(Val(PAPdiastxt.Text) + (Val(PAPsystxt.Text) - Val(PAPdiastxt.Text)) / 3, "###.0")
PAPmean = Val(PAPmeantxt.Text)
SItxt.Text = Format(Val(CItxt.Text) / Val(ARtxt.Text) * 1000, "###.00")
SI = Val(SItxt.Text)
LCWtxt.Text = Format(Val(CItxt.Text) * Val(MAPtxt.Text) * 1.44 / 100, "###.00")
LCW = Val(LCWtxt.Text)
LVSWitxt.Text = Format(SI * Val(MAPtxt.Text) * 1.44 / 100, "###.00")
LVSWI = Val(LVSWitxt.Text)
RCWtxt.Text = Format(Val(CItxt.Text) * Val(PAPmeantxt.Text) * 1.44 / 100, "###.00")
RCW = Val(RCWtxt.Text)
RVSWItxt.Text = Format(SI * Val(PAPmeantxt.Text) * 1.44 / 100, "###.00")
RVSWI = Val(RVSWItxt.Text)
'msgbox  str(79.92 * (val(MAPtxt.text) - val(CVptxt.text))/ val(CItxt.text))
SVRtxt.Text = Format(79.92 * (Val(MAPtxt.Text) - Val(CVPtxt.Text)) / Val(CItxt.Text), "######.0")
SVRI = Val(SVRtxt.Text)
PVRItxt.Text = Format(79.92 * (Val(PAPmeantxt.Text) - Val(PCWPtxt.Text)) / Val(CItxt.Text), "######.0")
PVRI = Val(PVRItxt.Text)
VDVTtxt.Text = Format(0.76 - 0.02 * Val(RVSWItxt.Text), "###.00")
Vd_Vt = Val(VDVTtxt.Text)
AVO2dif2txt.Text = Str(Val(AVO2diftxt.Text))
VO2txt.Text = Str(VO2i)
REEtxt.Text = Format(Val(VO2itxt.Text) * 1.44 * 5, "####.0")
REE = Val(REEtxt.Text)
RQ1txt.Text = Format((Val(PvCO2txt.Text) - Val(PaCO2txt.Text)) / Val(AVO2diftxt.Text), "###.000")
DO2txt.Text = Format(Val(CItxt.Text) * CaO2 * 10, "###.00")
DO2 = Val(DO2txt.Text)
O2exttxt.Text = Format(100 * Val(AVO2diftxt.Text) / Val(CaO2txt.Text), "###.00")


SvO2 = Val(SvO2txt.Text)
Systolic = Val(Systolictxt.Text)
Temp = Val(Temptxt.Text)
TimedUrine = Val(TimedUrinetxt.Text)
uChloride = Val(uChloridetxt.Text)
UCreatinine = Val(UCreatininetxt.Text)
Uosml = Val(Uosml.Text)
uPotassium = Val(uPotassiumtxt.Text)
Urate = Val(Urate.Text)
uSodium = Val(Usodiumtxt.Text)
VolumeUrine = Val(VolumeUrinetxt.Text)
VT = Val(VTtxt.Text)
Weight = Val(Weighttxt.Text)
Datetxt.Text = Date
Hb = Val(HBtxt.Text)
HCO3 = Val(HCO3txt.Text)
Heightt = Val(Heighttxt.Text)
IBW = Val(IBWtxt.Text)
IFR = Val(IFRtxt.Text)
Potassium = Val(Ktxt.Text)
LastName = Lasttxt.Text
Lipids = Val(Lipidstxt.Text)
Magn = Val(Magn.Text)
MAP = Val(MAPtxt.Text)
MiddleName = Middletxt.Text
Sodium = Natxt.Text
PaCO2 = Val(PaCO2txt.Text)
PaO2 = Val(PaO2txt.Text)
PeakAirwayPressure = Val(PAPrestxt.Text)
Patm = Val(Patmtxt.Text)
PCWPi = Val(PCWPitxt.Text)
PCWP = Val(PCWPtxt.Text)
PCWPe = PCWP
PEEPe = Val(PEEPetxt.Text)
PEEPi = Val(PEEPitxt.Text)
Phos = Val(Phos.Text)
pH = Val(pHtxt.Text)
PlateauPressure = Val(PlatPrestxt.Text)
PvCO2 = Val(PvCO2txt.Text)
PvO2 = Val(PvO2txt.Text)
rBS = Val(rBS.Text)
BStxt.Text = rBS.Text
Chloride = Val(rCL.Text)
CLtxt.Text = rCL.Text
RR = Val(RespRateTxt.Text)
RespRateTxt.Text = RRtxt.Text
Potassium = Val(rK.Text)
Ktxt.Text = rK.Text
Sodium = Val(rNa.Text)
Natxt.Text = rNa.Text
RR = Val(RRtxt.Text)
RespRateTxt = RRtxt.Text
SaO2 = Val(SaO2txt.Text)
Sosm = Val(Sosmtxt.Text)
ABW = Val(ABWtxt.Text)
Age = Val(Agetxt.Text)
AR = Val(ARtxt.Text)
rBS = Val(BStxt.Text)
BUN = Val(BUNtxt.Text)
Calc = Val(Calc.Text)
Chloride = Val(CLtxt.Text)
CO2 = Val(CO2txt.Text)
Creatinine = Val(Creatininetxt.Text)
Diastolic = Val(Diastolictxt.Text)
EtCO2 = Val(EtCO2txt.Text)
Sex = 1
FiO2 = Val(FiO2txt.Text)
FirstName = Firsttxt.Text
ABW = Val(ABWtxt.Text)
Age = Val(Agetxt.Text)
AR = Val(ARtxt.Text)
rBS = Val(BStxt.Text)
BUN = Val(BUNtxt.Text)
Calc = Val(Calc.Text)
Chloride = Val(CLtxt.Text)
CO2 = Val(CO2txt.Text)
Creatinine = Val(Creatininetxt.Text)
Diastolic = Val(Diastolictxt.Text)
EtCO2 = Val(EtCO2txt.Text)
FiO2 = Val(FiO2txt.Text)
FirstName = Firsttxt.Text
CVP = Val(CVPtxt.Text)
CO = Val(COtxt.Text)
CI = Val(CItxt.Text)
FickCI = Val(FickCItxt.Text)
PAPsys = Val(PAPsystxt.Text)
PAPdias = Val(PAPdiastxt.Text)
PAPmean = Val(PAPmeantxt.Text)
RVsyst = Val(RVs.Text)
RVdiast = Val(RVdias.Text)
RAsyst = Val(RAs.Text)
RAdiast = Val(RAd.Text)
VO2i = Val(VO2itxt.Text)
SV = Val(SV.Text)
COP = Val(COPtxt.Text)
COP_PCWP = Val(COP_PCWPtxt.Text)
CalcSosm = Val(CalcSosmtxt.Text)
RQ = Val(RQtxt.Text)
WBC = Val(WBCtxt.Text)
Hematocrit = Val(HctTxt.Text)
End Sub

Public Sub FlatButton5_Click()
Dim PHTNFlag As Single
PHTNFlag = 0
InterpretHemo.Text = " "
If CI < 2.8 Then
    InterpretHemo.Text = "The Cardiac Index is DECREASED.   A hypoperfusion state exists. " + Chr(10) + Chr(10)
Else
    If CI > 4.2 Then
        InterpretHemo.Text = InterpretHemo.Text + "The Cardiac Index is INCREASED. A high output state exists. " + Chr(10) + Chr(10)
    Else
        InterpretHemo.Text = InterpretHemo.Text + "The Cardiac Index is NORMAL. " + Chr(10) + Chr(10)
    End If
End If
If CVP < 2 Then
    InterpretHemo.Text = "The CVP is DECREASED. " + Chr(10) + Chr(10)
Else
    If CVP > 10 Then
        InterpretHemo.Text = InterpretHemo.Text + "The CVP is INCREASED. " + Chr(10) + Chr(10)
    Else
        InterpretHemo.Text = InterpretHemo.Text + "The CVP is NORMAL. " + Chr(10) + Chr(10)
    End If
End If


If PCWP < 5 Then
    InterpretHemo.Text = InterpretHemo.Text + "The Pulmonary Capillary Wedge Pressure (End Expiration) is DECREASED. "
    If CI < 2.8 Then
        InterpretHemo.Text = InterpretHemo.Text + "If this patient is post-MI, then Forrester Class 3 exists.  Hydration might be "
        InterpretHemo.Text = InterpretHemo.Text + "considered as Hypovolemia is likely. "
        If CVP > PCWP Then InterpretHemo.Text = InterpretHemo.Text + "Right Ventricular Infarction should be considered as CVP > PCWP. "
        If MAP < 70 And PAPmean < 9 And CVP < 8 And SVRI > 2500 And PVRI > 240 And SvO2 < 60 And SI < 30 Then
            InterpretHemo.Text = InterpretHemo.Text + "Hypovolemic Shock is highly suggested by low BP, PAP, PAWP, CVP, CI and "
            InterpretHemo.Text = InterpretHemo.Text + " SVO2; with elevated SVRI and PVRI. "
        End If
        If MAP < 70 And PAPmean < 9 And CVP < 8 And SVRI > 2500 And PVRI > 240 And SI < 30 Then
            InterpretHemo.Text = InterpretHemo.Text + "Hypovolemic Shock is suggested by low BP, PAP, PAWP, CVP and CI; "
            InterpretHemo.Text = InterpretHemo.Text + " with elevated SVRI and PVRI. "
        End If
    End If
    If CI > 4.2 And MAP < 105 And PAPmean < 9 And CVP < 12 And SVRI < 1200 And PVRI > 80 And SvO2 >= 60 And SI < 30 Then
        InterpretHemo.Text = InterpretHemo.Text + "Hyperdynamic Shock is highly suggested by low BP(or n), PAP, PAWP, SVRI, "
        InterpretHemo.Text = InterpretHemo.Text + "SI and CVP(or n); with elevated CI, SVO2(or n), PVR(or n).  Highly consider Septic Shock "
    End If
    If CI > 4.2 And MAP < 105 And PAPmean < 9 And CVP < 12 And SVRI < 1200 And PVRI > 80 And SvO2 < 60 And SI < 30 Then
        InterpretHemo.Text = InterpretHemo.Text + "Hyperdynamic Shock is highly suggested by low BP(or n), PAP, PAWP, SVRI, "
        InterpretHemo.Text = InterpretHemo.Text + "SI and CVP(or n); with elevated CI and PVR(or n).  The low SVO2 suggests maximal oxygen extraction (not "
        InterpretHemo.Text = InterpretHemo.Text + "characteristic of the shunt in septic shock), consider anemia, thyrotoxicosis or pregnancy. "
    End If
Else
    If PCWP > 12 Then
        InterpretHemo.Text = InterpretHemo.Text + "The Pulmonary Capillary Wedge Pressure (End Expiration) is INCREASED "
        InterpretHemo.Text = InterpretHemo.Text + "(presuming Normal Ventricular Compliance).  Consider Left Atrial Hypertension from LV MI, Valvular Disease, "
        InterpretHemo.Text = InterpretHemo.Text + " Cardiomyopathy.   Also Consider Pulmonary Venous Obstruction (rare), tumors, vasculitis and atria myxoma. "
            If PCWP > 18 Then
                InterpretHemo.Text = InterpretHemo.Text + "The Wedge Pressure exceeds 18 and is definitely elevated even with decreased LV compliance as in MI. "
                InterpretHemo.Text = InterpretHemo.Text + "The same considerations exist. "
                    If CI < 2.8 Then
                        InterpretHemo.Text = InterpretHemo.Text + "If this patient is post-MI, then a Forrester Class 4 exists.   Prognosis is poor. "
                        InterpretHemo.Text = InterpretHemo.Text + "Consider revascularization, ionotropes, afterload reduction or IA ballon pump. "
                    End If
                    If CI < 1.8 And MAP < 80 And PAPmean > 16 And CVP >= 12 And SVRI > 2500 And PVRI > 240 And SvO2 < 60 And SI < 30 Then
                        InterpretHemo.Text = InterpretHemo.Text + "Cardiogenic Shock is present consider Acute  MI, Cardiomyopathy, etc.   Also, "
                        InterpretHemo.Text = InterpretHemo.Text + "Hypodynamic Shock associated with septic or non-septic Multiorgan Failure may exist.   Prognosis is very poor. "
                    End If
            End If
            If CI > 4.2 Then
                InterpretHemo.Text = InterpretHemo.Text + "If the patient is post-MI then a Forrester 2 exists.   Prognosis is guarded. "
                InterpretHemo.Text = InterpretHemo.Text + "Diuresis for hypervolemia and CHF consistent with diastolic dysfunction. "
            End If
    Else
        InterpretHemo.Text = InterpretHemo.Text + "The Pulmonary Capillary Wedge Pressure (End Expiration) is Normal. "
        If CI > 4.2 Then InterpretHemo.Text = InterpretHemo.Text + "If the patient is post-MI then a Forrester 1 exists.   Prognosis is good. "
    End If
End If
If PAPmean > 20 Then
    PHTNFlag = 1
    InterpretHemo.Text = InterpretHemo.Text + Chr(10) + Chr(10) + "Pulmonary Arterial Hypertension is PRESENT (PAPmean > 20 mmHg). "
    InterpretHemo.Text = InterpretHemo.Text + " Clinical correlation to differentiate LV/LA hypertension, Cor-Pulmonale, Veno-occlusive disease, Thrombotic Pulmonary Hypertension or Primary Pulmonary Hypertension. " + Chr(10) + Chr(10)
End If
If PAPdias - PCWP > 5 Then
    InterpretHemo.Text = InterpretHemo.Text + "Pulmonary Arterial Hypertension is Suggested by the PADP - PCWP > 5 mmHg. "
    If PHTNFlag = 1 Then
        PHTNFlag = 0
        InterpretHemo.Text = InterpretHemo.Text + " Clinical correlation to differentiate LV/LA hypertension, Cor-Pulmonale, Veno-occlusive disease, Thrombotic Pulmonary Hypertension or Primary Pulmonary Hypertension. "
    End If
    InterpretHemo.Text = InterpretHemo.Text + "The PADP can not be substituted for the PCWP as a measure of Left Atrial Function! " + Chr(10) + Chr(10)
Else
    If PAPdias - PCWP < 1 Then
        InterpretHemo.Text = InterpretHemo.Text + "The PADP is less than the PCWP.   This may reflect a non-Zone 3 placement of the Catheter tip. "
        InterpretHemo.Text = InterpretHemo.Text + "Severe Mitral Regurgitation with Giant V-Waves can also cause this "
        InterpretHemo.Text = InterpretHemo.Text + "The PADP can not be substituted for the PCWP as a measure of Left Atrial Function! " + Chr(10) + Chr(10)
    End If
End If
If CVP > 12 Or CVP > PCWP Then
    If PAPmean < 16 And PAPsys < 30 And PCWP <= 18 Then
        InterpretHemo.Text = InterpretHemo.Text + "Right Atrial and Central Venous Pressures are elevated with Normal Pulmonary Vascular Pressures and Normal PCWP. "
        InterpretHemo.Text = InterpretHemo.Text + "If the patient is post-MI (Inferior with hypotension from Nitrates) strongly consider Right Ventricular Infarction. "
        InterpretHemo.Text = InterpretHemo.Text + "Also consider Massive Pulmonary Embolism, Pericardial Tamponade, Acute MI and Constrictive Pericarditis. " + Chr(10) + Chr(10)
    End If
End If
If AVO2dif > 6 And CvO2 < 12 And PvO2 <= 20 Then  'check CVO2
        InterpretHemo.Text = InterpretHemo.Text + "Arterial to Venous Oxygen Content is Increased because Venous Oxygen Content is REDUCED. "
        InterpretHemo.Text = InterpretHemo.Text + "This usually reflects increased tissue oxygen extraction.   Suspect Hypoxic Hypoxia at this level of PvO2. " + Chr(10) + Chr(10)
Else
    If AVO2dif > 6 And CvO2 < 12 Then
        InterpretHemo.Text = InterpretHemo.Text + "Arterial to Venous Oxygen Content is Increased because Venous Oxygen Content is REDUCED. "
        InterpretHemo.Text = InterpretHemo.Text + "This usually reflects increased tissue oxygen extraction.   Suspect Anemic Hypoxia at this level of PvO2. " + Chr(10) + Chr(10)
    End If
End If
If AVO2dif < 4 And CvO2 > 12 Then
    InterpretHemo.Text = InterpretHemo.Text + "AV O2 content difference is Narrowed reflecting an INCREASED CvO2.  This suggests that tissue utilization "
    InterpretHemo.Text = InterpretHemo.Text + "of O2 is decreased or being bypassed by a shunt as in Sepsis.   This may also only reflect an alteration "
    InterpretHemo.Text = InterpretHemo.Text + "in blood flow distribution as MvO2 reflects all tissue vascular beds.   Clinical correlation is advised. "
    InterpretHemo.Text = InterpretHemo.Text + "A decreasing PvO2 (SvO2) and an Elevated AvO2 Difference would reflect improvement in O2 extraction in a Septic "
    InterpretHemo.Text = InterpretHemo.Text + "situation and a better prognosis while in most studies a low PvO2 co##otes a poor outcome.   In Sepsis and the "
    InterpretHemo.Text = InterpretHemo.Text + "MODS, there is no convincing data supporting an improvement in survival by attempting to optimize hemodynamics or Oxygen Transport. " + Chr(10) + Chr(10)
End If
If RAdiast < (PAPdias + 2) And RAdiast > (PAPdias - 2) And RVdiast < (PAPdias + 2) And RVdiast > (PAPdias - 2) Then
    InterpretHemo.Text = InterpretHemo.Text + Chr(10) + Chr(10) + "There appears to be Equalization of RA, RV and PA diastolic pressures.  This suggest a restriction of diastolic filling as in "
    InterpretHemo.Text = InterpretHemo.Text + "Pericardial Tamponade.  Rule out Constrictive Pericarditis or Restrictive Cardiomyopathy as well.  A marked Pulsus Paradoxicus "
    InterpretHemo.Text = InterpretHemo.Text + "would suggest Tamponade in this setting as well. " + Chr(10) + Chr(10)
End If
If RASO2 > sSVCSO2 Or RASO2 > iSVCSO2 Then
    InterpretHemo.Text = InterpretHemo.Text + "There is a Step-Up in Oxygen Saturation from the Superior Vena Cava into the Right Atrium. "
    InterpretHemo.Text = InterpretHemo.Text + "Rule out an Atrial Septal Defect with Left to Right Shunting. "
    InterpretHemo.Text = InterpretHemo.Text + "There is NO acceptable site for sampling Mixed Venous Blood! " + Chr(10) + Chr(10)
End If
If RVSO2 > RASO2 Then
    InterpretHemo.Text = InterpretHemo.Text + "There is a Step-Up in Oxygen Saturation from the Right Atrium into the Right Ventricle. "
    InterpretHemo.Text = InterpretHemo.Text + "Rule out a Ventricular Septal Defect with Left to Right Shunting.  This may complicate an extensive Anterior MI. "
    InterpretHemo.Text = InterpretHemo.Text + "The Right Atrium is the most appropriate site to sample Mixed Venous Blood in this setting. " + Chr(10) + Chr(10)
End If
End Sub

Private Sub FlatButton6_Click()
InterpretHemo.Text = ""
InterpretHemo.Text = InterpretHemo.Text + "Strict Aseptic Technique is essential.  Remove catheter ASAP as increased morbidity, mortality, LOS and cost have been associated. "
InterpretHemo.Text = InterpretHemo.Text + "Before insertion, check patency, balloon symmetry, calibration (LA = 0), damping & frequency response to whipping.  Record all measurements on paper. "
InterpretHemo.Text = InterpretHemo.Text + "The Right Internal Jugular is the most direct insertion site followed by the Left Subclavian.  The catheter should be inserted to 20 cm. "
InterpretHemo.Text = InterpretHemo.Text + "at which time the balloon should be inflated.  The balloon must be inflated whenever advancing the catheter and deflated when retracting it.  Observe "
InterpretHemo.Text = InterpretHemo.Text + "the characteristic Wave Forms should be identified.  The RA is usually at 24 cm, RV at 28, PA at 32 and PCWP at 34.  Use caution not to knot the catheter. "
InterpretHemo.Text = InterpretHemo.Text + "Oxygen Saturation may be obtained during insertion.  If a CLBBB is present, use extreme caution as a CRBBB may be induced by insertion and thus a Complete "
InterpretHemo.Text = InterpretHemo.Text + "Heart Block.  Consider a Pacer-Swan in this setting.  Transient Cardiac Arrhythmias are the most common complication.  The PCWP position is verified by the "
InterpretHemo.Text = InterpretHemo.Text + "Wave Form Changes but also one typically sees a fall in the Pulmonary Artery Mean Pressure upon wedging.  When balloon occlusion occurs, repeatedly deflate "
InterpretHemo.Text = InterpretHemo.Text + "and retract and then inflate until occlusion is lost, then advance to just re-occlude.  This takes the slack out of the catheter especially as it softens at "
InterpretHemo.Text = InterpretHemo.Text + "body temperature.  Insure a Zone 3 position of the catheter by the absence of marked respiratory variation in PCWP tracing.  If the PADP > PCWP, the catheter "
InterpretHemo.Text = InterpretHemo.Text + "tip may not be in a Zone 3 (MR with giant V-waves may also cause this.)  If the catheter tip is below the level of the LA on a cross table lateral CXR, a Zone 3 "
InterpretHemo.Text = InterpretHemo.Text + "state probably exists.  Finally, one must realize that PCWP is presumed to reflect Left Ventricular Preload.  This is predicated by PCWP  RAP  LVEDP  LVEDV  "
InterpretHemo.Text = InterpretHemo.Text + "Myocardial Fiber Length.  Ventricular compliance changes may alter these relationships.  Valvular pathology may also interfer as may Pulmonary-Veno-occlusive ds. "
InterpretHemo.Text = InterpretHemo.Text + "In ARDS (and MODS), in situ thrombosis and vasospasm may interupt the vascular co##ection.  Consider a Pulmonary Capillary Wedge Angiogram to insure Left Atrial "
InterpretHemo.Text = InterpretHemo.Text + "co##ection.  If patency isn't present, direct measurement of LAP may be neccessary."

End Sub

Private Sub FlatButton7_Click()
InterpretHemo.Text = ""
InterpretHemo.Text = "Hemodynamically measured and derived (calculated) variables are generally not reliable in assessing the Oxygen "
InterpretHemo.Text = InterpretHemo.Text + "Transport System (Oxygen Supply Dependency of Demand controversy).   This is felt to be related to a mathematical "
InterpretHemo.Text = InterpretHemo.Text + "coupling of the values which is not present when supply and consumption are measured separately. Therefore, "
InterpretHemo.Text = InterpretHemo.Text + "direct measurements of Oxygen Consumption should be performed (Metabolic Cart or Mass Spec) when these values are critical. "
InterpretHemo.Text = InterpretHemo.Text + "In addition, O2 Transport which is easily measured (CI*CaO2) is not the same as "
InterpretHemo.Text = InterpretHemo.Text + "Oxygen Delivery (which is complex, the most vital factor in tissue respiration and cannot be measured or estimated.) "
InterpretHemo.Text = InterpretHemo.Text + "This lack of correlation is very frequent in Sepsis, ARDS and MODS. " + Chr(10) + Chr(10)
InterpretHemo.Text = InterpretHemo.Text + "Finally, the Resting Energy Expenditure (REE) in this program is derived using these coupled variables. Its accuracy must be questioned therefore clinically. "
InterpretHemo.Text = InterpretHemo.Text + "If it correlates with a value verified by another method, then it may have utility. Likewise if the BEE (Harris-Benedict) correlates, it may be used. "

End Sub
Public Sub ABGInterpret()
Dim X As Long
Interpretive.Show

Interpretive.Interpret.Text = "                          Arterial Blood Gases & Gas Exchange" & vbCrLf & vbCrLf
nPaO2 = 0.9 * (104.2 - 0.27 * Age)
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
"Normal PaO2 = " + Format(nPaO2, "###.0") & vbCrLf
naAdif = 1.1 * (2.5 + 0.21 * Age)
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
"Normal aA Difference = " + Format(naAdif, "###.0")
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf
X = PaO2
If FiO2 <> 21 Then
   corPaO2 = 21 * PaO2FiO2 / 100
   PaO2 = corPaO2
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "You should assess Oxygentation on 21%. It is corrected to " + Format(PaO2, "###.0") & vbCrLf
End If

If PaO2 >= nPaO2 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "Normoxia is Present (over the 90 %tile)" & vbCrLf & vbCrLf
   GoTo Shunto
End If

If PaO2 >= 0.8 * nPaO2 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "Mild Hypoxia is Present (over the 80 %tile)" & vbCrLf & vbCrLf
   GoTo Shunto
End If

If PaO2 > 55 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "Moderate Hypoxia is Present (over 55 mmHg)" & vbCrLf & vbCrLf
   GoTo Shunto
End If

If PaO2 <= 55 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "Severe Hypoxia is Present (<= 55 mmHg)" & vbCrLf & vbCrLf
End If

Shunto:
If Qs_Qt = 0 Then GoTo aAdif1:

If Qs_Qt < 10 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "The Calculated Shunt is less than 10%. This is clinically compatible with Normal Lungs and Gas Exchange." & vbCrLf & vbCrLf
   GoTo aAdif1
End If

If Qs_Qt > 10 And Qs_Qt < 19 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "A Shunt of 10-19% denotes a degree of pathology that seldom requires significant support" & vbCrLf & vbCrLf
   GoTo aAdif1
End If

If Qs_Qt >= 19 And Qs_Qt < 29 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "A Shunt in this range may be life threatening in a patient with limited Cardiovascular or CNS Reserve." & vbCrLf & vbCrLf
   GoTo aAdif1
End If

If Qs_Qt >= 30 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "A Shunt greater than 30 is potentially life-threatening and usually requires significant Cardiopulmonary Support." & vbCrLf & vbCrLf
End If

aAdif1:
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
"When significant shunt effect is present (as in Low V/Q), the shunt calculation will increase significantly as FiO2 is lowered from 50%"
PaO2 = X
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf
Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf
If Aadif > naAdif Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "An Increased a-A Difference Exists" & vbCrLf
Else
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Normal a-A Difference Exists" & vbCrLf
End If

If pH >= 7.38 And pH <= 7.42 And PaCO2 >= 38 And PaCO2 <= 42 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Normal Acid Base Balance Exists" & vbCrLf
    Exit Sub
End If

If pH <= 7.4 And PaCO2 > 42 Then GoTo AcidBase1

If pH > 7.42 And PaCO2 > 38 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Primary Metabolic Alkalosis Exists with normal compensation" & vbCrLf
    Exit Sub
End If

If pH <= 7.42 And PaCO2 <= 42 Then GoTo AcidBase2

If pH > 7.4 And PaCO2 <= 38 Then GoTo AcidBase3



AcidBase1:
j8 = (740 - pH * 100) / (PaCO2 - 42)

If j8 < 0.2 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Mixed Respiratory Acidosis and Metabolic Alkalosis Exists" & vbCrLf
    Exit Sub
End If

If j8 <= 0.4 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Chronic Respiratory Acidosis is Present" & vbCrLf
    Exit Sub
End If

If j8 <= 0.69 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Acute on Chronic Respiratory Acidosis with a Mixed Respiratory and Metabolic Acidosis is Present" & vbCrLf
    Exit Sub
End If

If j8 <= 0.9 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "An Acute Respiratory Acidosis Exists" & vbCrLf
    Exit Sub
End If

If j8 < 0.2 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Mixed Respiratory and Metabolic Alkalosis Exists" & vbCrLf
    Exit Sub
End If

                                      
AcidBase2:
j8 = (42 - PaCO2) / (25 - HCO3)

If j8 < 1 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Mixed Metabolic + Respiratory Acidosis Exists" & vbCrLf
    Exit Sub
End If

If j8 <= 1.2 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Compensated Metabolic Acidosis Exists" & vbCrLf
    Exit Sub
End If
 
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Mixed Metabolic Acidosis + Respiratory Alkalosis Exists" & vbCrLf
    Exit Sub

                      
AcidBase3:
j8 = (pH * 100 - 740) / (38.1 - PaCO2)

If j8 < 0.2 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "A Mixed Respiratory Alkalosis and Metabolic Acidosis Exists" & vbCrLf
    Exit Sub
End If


If j8 <= 0.6 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "Chronic Respiratory Alkalosis Exists" & vbCrLf
    Exit Sub
End If
  
If j8 <= 0.69 Then
   Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
   "An Acute on Chronic Respiratory Alkalosis or a Mixed Respiratory + Metabolic Alkalosis Exists" & vbCrLf
    Exit Sub
End If


If j8 <= 0.9 Then
    Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
    "An Acute Respiratory Alkalosis Exists" & vbCrLf
    Exit Sub
End If

Interpretive.Interpret.Text = Interpretive.Interpret.Text & vbCrLf & _
"A Mixed Respiratory + Metabolic Alkalosis Exists" & vbCrLf
End Sub

Private Sub Form_Activate()
DataInput.Left = Hemodynamics.Left + 100
DataInput.Top = Hemodynamics.Top + 875
FlatButton1_Click
FlatButton2_Click
FlatButton3_Click
FlatButton4_Click
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True
Datetxt.Text = Date
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
'Interpretive.Left = Hemodynamics.Left + 100
'Interpretive.Top = Hemodynamics.Top + 900

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Interpretive.Left = Hemodynamics.Left + 100
'Interpretive.Top = Hemodynamics.Top + 900

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Interpretive.Left = Hemodynamics.Left + 100
'Interpretive.Top = Hemodynamics.Top + 900

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Picture1_Click()

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


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(0).Visible = False
Image1(1).Visible = True
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(0).Visible = True
Image1(1).Visible = False
Load O2Dose
O2Dose.Show
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2(0).Visible = False
    Image2(1).Visible = True
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2(0).Visible = False
    Image2(1).Visible = True
    Load Sats
    Sats.Show
End Sub

Private Sub Label107_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label107(0).Visible = False
        Label107(2).Visible = True
End Sub

Private Sub Label107_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label107(0).Visible = True
        Label107(2).Visible = False
        Me.Visible = False
End Sub







Private Sub Label108_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label37_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Load O2Dose
    O2Dose.Show
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub SStab2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

