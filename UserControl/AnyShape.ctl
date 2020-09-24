VERSION 5.00
Begin VB.UserControl AnyShape 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
   MaskColor       =   &H00808080&
   ScaleHeight     =   630
   ScaleWidth      =   810
   Begin VB.PictureBox pic2 
      BackColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   3555
      ScaleHeight     =   495
      ScaleWidth      =   630
      TabIndex        =   1
      Top             =   720
      Width           =   690
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   3735
      ScaleHeight     =   855
      ScaleWidth      =   1890
      TabIndex        =   0
      Top             =   1620
      Width           =   1950
   End
End
Attribute VB_Name = "AnyShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 

'   CONTROL IS DRAWN OR DBL CLICKED ON FORM FOR FIRST TIME
'          -UserControl_Initialize
'          -UserControl_InitProperties
'          -UserControl_Resize
'          -UserControl_Show
'          -UserControl_Paint
'
'   CONTROL IS RESIZED ON FORM AT DESIGN TIME
'          -UserControl_Resize
'          -UserControl_Paint
'
'   THE "RUN" OR "START" BUTTON IS CLICKED PUTTING CONTROL IN RUN MODE
'          -UserControl_Hide
'          -UserControl_Terminate
'          -UserControl_Initialize
'          -UserControl_Resize
'          -UserControl_ReadProperties
'          -UserControl_Show
'          -UserControl_EnterFocus
'          -UserControl_GotFocus
'          -UserControl_Paint
'
'    THE FORM IS CLOSED OR TEMINATED, PUTTING CONTROL BACK IN DESIGN MODE
'          -UserControl_LostFocus
'          -UserControl_ExitFocus
'          -UserControl_Hide
'          -UserControl_Terminate
'          -UserControl_Initialize
'          -UserControl_Resize
'          -UserControl_ReadProperties
'          -UserControl_Show
'          -UserControl_Paint
'
'    YOU CHANGE YOUR PROJECT FROM CODE VIEW TO DESIGN (FORM) VIEW
'          -UserControl_Paint
'
'    THE CONTROL IS REMOVED FROM THE FORM
'          -UserControl_WriteProperties
'          -UserControl_Hide
'          -UserControl_Terminate
'=======================================================================

'[EVENTS]
Event MouseEnter()
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'[CONSTANTS]
Private Const SRCCOPY = &HCC0020

'[ENUMS]
Private Enum enBS
   enNull = 0
   enDown = 1
   enEntered = 2
End Enum

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
 
Dim BtnState As enBS
Dim m_maskColor As Long
'Default Property Values:
Const m_def_CaptionY = 30
Const m_def_CaptionX = 10
Const m_def_MouseOverCaptionColor = &HFF0000
Const m_def_CaptionColor = 0
Const m_def_Caption = ""
'Property Variables:
Dim m_CaptionY As Long
Dim m_CaptionX As Long
Dim m_MouseOverCaptionColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_Caption As String



 
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BtnState = enDown
    Call UserControl_Paint
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------
' create the mouseenter
' and exits
'---------------------
     If (x < 0) Or (y < 0) Or (x > ScaleWidth) Or (y > ScaleHeight) Then
              ReleaseCapture
              BtnState = enNull
              Call UserControl_Paint
              'mouseexit event
              RaiseEvent MouseExit
                           
     ElseIf GetCapture() <> hwnd Then
              SetCapture hwnd
              BtnState = enEntered
              Call UserControl_Paint
              'mouseenter event
              RaiseEvent MouseEnter
     End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BtnState = enNull
    Call UserControl_Paint
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()

        pic2.Cls

        'the mouse is not down
        If BtnState = enNull Or BtnState = enEntered Then
            'paint the picture(button) to pic2, since were painting from
            'pic1 (the button image) to pic2, 2 pixels less wide and high
            'than than the actual button picture, 2 pixels of pic2's backcolor
            'shows through to the right and bottom creating the shadow effect
            TransparentBlt pic2.hdc, 0, 0, (ScaleWidth - 2), (ScaleHeight - 2), _
                        pic1.hdc, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, _
                        RGB(255, 0, 255)
        'mouse is down
        ElseIf BtnState = enDown Then
             'paint the picture(button) to pic2
             'now that were painting the same dimensions from source
             'to destination, the shadow effect is eliminated, enhancing
             'the mouse down effect on the image
             TransparentBlt pic2.hdc, 0, 0, ScaleWidth, ScaleHeight, _
                        pic1.hdc, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, _
                        RGB(255, 0, 255)
        End If

       'print caption the right color based on whether
       'mouse is over control or not
        If BtnState = enEntered Or BtnState = enDown Then
            Call PrintCaption(True)
        ElseIf BtnState = enNull Then
            Call PrintCaption
        End If
        
        'now paint what is drawn on picture2 to the usercontrol
        'pic2 acts as a paint buffer so we get no flicker on the UC
        BitBlt hdc, 0, 0, ScaleWidth, ScaleHeight, pic2.hdc, 0, 0, SRCCOPY
End Sub

Private Sub PrintCaption(Optional bMouseEnter As Boolean = False)
'----------------------
Dim txtWid&, txtHei&
Dim lPos&, tPos&
'----------------------
' code for the controls
' caption
'----------------------
             'print caption the right color based on whether
             'mouse is over control or not
             If bMouseEnter = True Then
                  SetTextColor pic2.hdc, m_MouseOverCaptionColor
             Else
                  SetTextColor pic2.hdc, m_CaptionColor
             End If
             
             'prints the text
             TextOut pic2.hdc, m_CaptionX, m_CaptionY, m_Caption, Len(m_Caption)
End Sub
Private Sub UserControl_Resize()
          'constrict uc size to the pic1 dimensions
          'remember, pic1 holds the actual image we
          'will be drawing to the UC
          If pic1.Picture <> 0 Then
              'UserControl.Size pic1.ScaleWidth, pic1.ScaleHeight
          End If
End Sub

Private Sub UserControl_Show()
    Call ReadyControl
End Sub
Private Sub ReadyControl()
   With pic1
      .ScaleMode = 3 'pixels
      .AutoRedraw = True
      .AutoSize = True
      .BorderStyle = 0 'none
      .Appearance = 0 'flat
   End With
   With pic2
      .ScaleMode = 3 'pixels
      .AutoRedraw = True
      .AutoSize = False
      .BorderStyle = 0 'none
      .Appearance = 0 'flat
      .BackColor = &H808080
   End With
   With UserControl
      .ScaleMode = 3 'pixels
      .BackStyle = 0 'transparent
      .AutoRedraw = False 'were doing the painting
   End With
End Sub











 

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
        m_Caption = m_def_Caption
        m_CaptionColor = m_def_CaptionColor
        m_MouseOverCaptionColor = m_def_MouseOverCaptionColor
        m_CaptionX = m_def_CaptionX
        m_CaptionY = m_def_CaptionY
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        Set Picture = PropBag.ReadProperty("Picture", Nothing)
        UserControl.MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
        m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
        m_CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
        m_MouseOverCaptionColor = PropBag.ReadProperty("MouseOverCaptionColor", m_def_MouseOverCaptionColor)
        m_CaptionX = PropBag.ReadProperty("CaptionX", m_def_CaptionX)
        m_CaptionY = PropBag.ReadProperty("CaptionY", m_def_CaptionY)
End Sub
 
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        Call PropBag.WriteProperty("Picture", Picture, Nothing)
        Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &HFF00FF)
        Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
        Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)
        Call PropBag.WriteProperty("MouseOverCaptionColor", m_MouseOverCaptionColor, m_def_MouseOverCaptionColor)
        Call PropBag.WriteProperty("CaptionX", m_CaptionX, m_def_CaptionX)
        Call PropBag.WriteProperty("CaptionY", m_CaptionY, m_def_CaptionY)
End Sub
 
'[PICTURE]
Public Property Get Picture() As Picture
        Set Picture = pic1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
        Set pic1.Picture = New_Picture
        PropertyChanged "Picture"
        
        If pic1.Picture <> 0 Then
            UserControl.Size pic1.Width, pic1.Height + 200
            pic2.Width = pic1.Width: pic2.Height = pic1.Height
            UserControl.MaskPicture = pic1.Picture
        End If
        
        Call UserControl_Paint
End Property

'[MASKCOLOR]
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
        MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
        m_maskColor = New_MaskColor
        UserControl.MaskColor() = New_MaskColor
        PropertyChanged "MaskColor"
        Call UserControl_Paint
End Property

'[CAPTION]
Public Property Get Caption() As String
        Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
        m_Caption = New_Caption
        PropertyChanged "Caption"
        Call UserControl_Paint
End Property

'[CAPTIONCOLOR]
Public Property Get CaptionColor() As OLE_COLOR
        CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
        m_CaptionColor = New_CaptionColor
        PropertyChanged "CaptionColor"
        Call UserControl_Paint
End Property

'[MOUSEOVERCAPTIONCOLOR]
Public Property Get MouseOverCaptionColor() As OLE_COLOR
        MouseOverCaptionColor = m_MouseOverCaptionColor
End Property

Public Property Let MouseOverCaptionColor(ByVal New_MouseOverCaptionColor As OLE_COLOR)
        m_MouseOverCaptionColor = New_MouseOverCaptionColor
        PropertyChanged "MouseOverCaptionColor"
End Property

'[CAPTIONX]
Public Property Get CaptionX() As Long
Attribute CaptionX.VB_Description = "the x position, in pixels of the start of captioin printing"
        CaptionX = m_CaptionX
End Property

Public Property Let CaptionX(ByVal New_CaptionX As Long)
        m_CaptionX = New_CaptionX
        PropertyChanged "CaptionX"
        Call UserControl_Paint
End Property

'[CAPTIONY
Public Property Get CaptionY() As Long
Attribute CaptionY.VB_Description = "the y position, in pixels of the start of captioin printing"
        CaptionY = m_CaptionY
End Property

Public Property Let CaptionY(ByVal New_CaptionY As Long)
        m_CaptionY = New_CaptionY
        PropertyChanged "CaptionY"
        Call UserControl_Paint
End Property

