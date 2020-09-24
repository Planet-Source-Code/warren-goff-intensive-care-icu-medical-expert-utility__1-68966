VERSION 5.00
Begin VB.Form GlasgowFrm 
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
   Begin VB.ComboBox Motor 
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
      Left            =   1305
      TabIndex        =   14
      Text            =   "Select PDC"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.ComboBox Eyespop 
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
      Left            =   1305
      TabIndex        =   13
      Text            =   "Select PDC"
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Verbal 
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
      Left            =   1305
      TabIndex        =   9
      Text            =   "Select PDC"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton FlatButton5 
      Caption         =   "Back"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Glasgowtext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin GoffsICU.ocxFormShape ocxFormShape1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This scale is used to the assess Coma and Impaired consciousness: < 8 are in Coma; Mild is 13-15; Moderate is 9-12; Severe 3-8"
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
      Height          =   735
      Left            =   720
      TabIndex        =   15
      Top             =   3600
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Eyes:"
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
      Left            =   2745
      TabIndex        =   11
      Top             =   585
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Verbal Response"
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
      Left            =   2145
      TabIndex        =   10
      Top             =   1320
      Width           =   1830
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
      TabIndex        =   8
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Motor Response"
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
      Left            =   2145
      TabIndex        =   6
      Top             =   2040
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Glasgow Coma Score ="
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
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   2820
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glasgow Coma Score"
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
      Left            =   1485
      TabIndex        =   3
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glasgow Coma Score"
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
      Left            =   1455
      TabIndex        =   2
      Top             =   90
      Width           =   3405
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
Attribute VB_Name = "GlasgowFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CandyButton1_Click()

End Sub

Private Sub Command1_Click()
Glasgow = 0
Select Case Eyespop.ListIndex + 1
Case 1    '"Opens Eyes"
    Glasgow = 4
Case 2    '"Opens Eyes on Request"
    Glasgow = 3
Case 3    '"Opens Eyes on Pain"
    Glasgow = 2
Case 4    '"Fails to Open Eyes"
    Glasgow = 1
End Select

Select Case Verbal.ListIndex + 1
Case 1    '"Appropriate Conversation, Oriented to Month + Yea"
    Glasgow = Glasgow + 5
Case 2    '"Confused and Disoriented"
    Glasgow = Glasgow + 4
Case 3   ' "Inapprorpiate Conversation"
    Glasgow = Glasgow + 3
Case 4    '"Incomprehensible Sounds"
    Glasgow = Glasgow + 2
Case 5     '"No Sounds"
    Glasgow = Glasgow + 1
End Select

Select Case Motor.ListIndex + 1
Case 1     '"Follow Simple Directions"
    Glasgow = Glasgow + 6
Case 2     '"Removes Pain Source"
    Glasgow = Glasgow + 5
Case 3     '"Withdraws from Pain Source"
    Glasgow = Glasgow + 4
Case 4     '"Decortication (non-purposeful Flexion)"
    Glasgow = Glasgow + 3
Case 5     '"Decerebration (non-purposeful Extension)"
    Glasgow = Glasgow + 2
Case 6     '"No Motor Response"
    Glasgow = Glasgow + 1
End Select

Glasgowtext.Text = Format(Glasgow, "##")
Glasgow = Val(Glasgowtext.Text)

End Sub



Private Sub Eyespop_Click()
Command1_Click
End Sub

Private Sub FlatButton5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Hemodynamics.Left + 100
Me.Top = Hemodynamics.Top + 900
SetTopMostWindow Me.hwnd, True

Dim ItemsOnList As Integer
Dim selectedname As String

'Eyes
Eyespop.Clear
Eyespop.AddItem "Opens Eyes"
Eyespop.AddItem "Opens Eyes on Request"
Eyespop.AddItem "Opens Eyes on Pain"
Eyespop.AddItem "Fails to Open Eyes"
'ItemsOnList = Eyespop.NoItems 'get items on list
Eyespop.Text = Eyespop.List(0)
Eyespop.ListIndex = 0


'Verbal
Verbal.Clear
Verbal.AddItem "Appropriate Conversation/Oriented"     ' to Month + Year"
Verbal.AddItem "Confused and Disoriented"
Verbal.AddItem "Inapprorpiate Conversation"
Verbal.AddItem "Incomprehensible Sounds"
Verbal.AddItem "No Sounds"
'ItemsOnList=Verbal.NoItems 'get items on list
'selectedName= Verbal.text(Verbal.listindex)
Verbal.Text = Verbal.List(0)
Verbal.ListIndex = 0


' Motor
Motor.Clear
Motor.AddItem "Follows Simple Directions"
Motor.AddItem "Removes Pain Source"
Motor.AddItem "Withdraws from Pain"     ' Source"
Motor.AddItem "Decortication" '(non-purposeful Flexion)"
Motor.AddItem "Decerebration"   ' (non-purposeful Extension)"
Motor.AddItem "No Motor Response"
'ItemsOnList=Motor.NoItems 'get items on list
'selectedName= Motor.text(Motor.listindex)
Motor.Text = Motor.List(0)
Motor.ListIndex = 0
Glasgow = 15
Glasgowtext.Text = 15
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



Private Sub Motor_Click()
Command1_Click
End Sub

Private Sub Verbal_Click()
Command1_Click
End Sub
