VERSION 5.00
Begin VB.UserControl lsaFrame 
   Alignable       =   -1  'True
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   ControlContainer=   -1  'True
   PropertyPages   =   "lsaFrameControl.ctx":0000
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   ToolboxBitmap   =   "lsaFrameControl.ctx":0023
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1080
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000015&
         Caption         =   "vb3DDKShadow"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "vb3DShadow"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "vb3DLight"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "vb3DHighlight"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000F&
      X1              =   32
      X2              =   64
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000F&
      X1              =   64
      X2              =   144
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000010&
      X1              =   64
      X2              =   64
      Y1              =   40
      Y2              =   24
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   24
      X2              =   72
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000015&
      X1              =   72
      X2              =   72
      Y1              =   16
      Y2              =   32
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   8
      X2              =   8
      Y1              =   32
      Y2              =   88
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000014&
      X1              =   72
      X2              =   152
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000010&
      X1              =   144
      X2              =   144
      Y1              =   40
      Y2              =   80
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   144
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000015&
      X1              =   8
      X2              =   152
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000015&
      X1              =   152
      X2              =   152
      Y1              =   32
      Y2              =   88
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000F&
      X1              =   32
      X2              =   32
      Y1              =   24
      Y2              =   40
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000F&
      X1              =   32
      X2              =   16
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000F&
      X1              =   16
      X2              =   16
      Y1              =   80
      Y2              =   40
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   24
      X2              =   24
      Y1              =   32
      Y2              =   16
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   8
      X2              =   24
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "lsaFrame"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "lsaFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Event Declarations:
Event Click() 'MappingInfo=Picture1,Picture1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Picture1,Picture1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Picture1,Picture1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Picture1,Picture1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Picture1,Picture1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Dim LightCol, SLightCol, DarkCol, SDarkCol
'Default Property Values:
'Const m_def_Light1 = 0
'Const m_def_Light2 = 0
'Const m_def_Dark1 = 0
'Const m_def_Dark2 = 0
'Property Variables:
'Dim m_Light1 As Variant
'Dim m_Light2 As Variant
'Dim m_Dark1 As Variant
'Dim m_Dark2 As Variant




Private Sub Label1_Change()
    UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
    Caption = "lsaFrame"
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
        On Error Resume Next
        'Set Colors:
        Picture3.BackColor = Picture2.BackColor
        Label1.BackColor = BackColor
        LightCol = Label3.BackColor
        SLightCol = Label4.BackColor
        DarkCol = Label5.BackColor
        SDarkCol = Label6.BackColor
        Line1.BorderColor = LightCol
        Line2.BorderColor = LightCol
        Line3.BorderColor = LightCol
        Line4.BorderColor = LightCol
        Line11.BorderColor = LightCol
        Line6.BorderColor = SLightCol
        Line7.BorderColor = SLightCol
        Line8.BorderColor = SLightCol
        Line10.BorderColor = SLightCol
        Line12.BorderColor = SLightCol
        Line9.BorderColor = DarkCol
        Line15.BorderColor = DarkCol
        Line16.BorderColor = DarkCol
        Line5.BorderColor = SDarkCol
        Line13.BorderColor = SDarkCol
        Line14.BorderColor = SDarkCol
        If Width <= 150 Then Width = 150
        If Height <= 360 Then Height = 360
        Line4.X1 = (Label1.Width / 4)
        Label1.Top = 2
        Label1.Left = (Label1.Width / 4) + 5
        Line4.X2 = Label1.Width + Label1.Left + 4
        Line4.Y1 = 0
        Line4.Y2 = 0
        Line3.X1 = (Label1.Width / 4)
        Line3.X2 = (Label1.Width / 4)
        Line3.Y1 = 0
        Line3.Y2 = (Label1.Height / 1.5)
        Line2.X1 = 0
        Line2.X2 = Line3.X2 + 1
        Line2.Y1 = (Label1.Height / 1.5)
        Line2.Y2 = (Label1.Height / 1.5)
        Line11.X1 = Line4.X2
        Line11.X2 = ScaleWidth
        Line11.Y1 = Line2.Y1
        Line11.Y2 = Line2.Y1
        Line5.X1 = Line4.X2
        Line5.X2 = Line4.X2
        Line5.Y1 = Line4.Y2
        Line5.Y2 = Line11.Y1
        Line1.X1 = 0
        Line1.X2 = 0
        Line1.Y1 = Line2.Y1
        Line1.Y2 = ScaleHeight
        Line14.X1 = 0
        Line14.X2 = ScaleWidth
        Line14.Y1 = ScaleHeight - 1
        Line14.Y2 = ScaleHeight - 1
        Line13.X1 = ScaleWidth - 1
        Line13.X2 = ScaleWidth - 1
        Line13.Y1 = Line11.Y2
        Line13.Y2 = ScaleHeight
        Line16.X1 = ScaleWidth - 2
        Line16.X2 = ScaleWidth - 2
        Line16.Y1 = Line11.Y2 + 2
        Line16.Y2 = ScaleHeight - 1
        Line15.X1 = 2
        Line15.X2 = ScaleWidth - 2
        Line15.Y1 = ScaleHeight - 2
        Line15.Y2 = ScaleHeight - 2
        Line9.X1 = Line4.X2 - 1
        Line9.X2 = Line4.X2 - 1
        Line9.Y1 = Line4.Y2 + 2
        Line9.Y2 = Line11.Y1 + 1
        Line12.X1 = Line4.X2
        Line12.X2 = ScaleWidth - 1
        Line12.Y1 = Line2.Y1 + 1
        Line12.Y2 = Line2.Y1 + 1
        Line10.X1 = (Label1.Width / 4) + 2
        Line10.X2 = Label1.Width + Label1.Left + 4
        Line10.Y1 = 1
        Line10.Y2 = 1
        Line8.X1 = Line3.X1 + 1
        Line8.X2 = Line3.X2 + 1
        Line8.Y1 = 1
        Line8.Y2 = Line3.Y2 + 1
        Line6.X1 = 1
        Line6.X2 = 1
        Line6.Y1 = Line2.Y1 + 1
        Line6.Y2 = ScaleHeight - 1
        Line7.X1 = 1
        Line7.X2 = Line3.X2 + 2
        Line7.Y1 = (Label1.Height / 1.5) + 1
        Line7.Y2 = (Label1.Height / 1.5) + 1
        Picture2.Top = 0
        Picture2.Left = 0
        Picture2.Width = Line4.X1
        Picture2.Height = Line2.Y1
        Picture3.Top = 0
        Picture3.Left = Line5.X1 + 1
        Picture3.Width = (ScaleWidth - Line5.X1)
        Picture3.Height = Line2.Y1
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BorderStyle






Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub Picture1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Picture1,Picture1,-1,ScaleWidth
'Public Property Get FrameWidth() As Single
'    FrameWidth = Picture1.ScaleWidth
'End Property
'
'Public Property Let FrameWidth(ByVal New_FrameWidth As Single)
'    Picture1.ScaleWidth() = New_FrameWidth
'    PropertyChanged "FrameWidth"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Picture1,Picture1,-1,ScaleHeight
'Public Property Get FrameHeight() As Single
'    FrameHeight = Picture1.ScaleHeight
'End Property
'
'Public Property Let FrameHeight(ByVal New_FrameHeight As Single)
'    Picture1.ScaleHeight() = New_FrameHeight
'    PropertyChanged "FrameHeight"
'End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Picture1.ScaleWidth = PropBag.ReadProperty("FrameWidth", 1755)
    Picture1.ScaleHeight = PropBag.ReadProperty("FrameHeight", 915)
    Label1.Caption = PropBag.ReadProperty("Caption", "CoolMan")
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    m_Light1 = PropBag.ReadProperty("Light1", m_def_Light1)
'    m_Light2 = PropBag.ReadProperty("Light2", m_def_Light2)
'    m_Dark1 = PropBag.ReadProperty("Dark1", m_def_Dark1)
'    m_Dark2 = PropBag.ReadProperty("Dark2", m_def_Dark2)
    Label3.BackColor = PropBag.ReadProperty("Light1", &H8000000F)
    Label4.BackColor = PropBag.ReadProperty("Light2", &H8000000F)
    Label5.BackColor = PropBag.ReadProperty("Dark1", &H8000000F)
    Label6.BackColor = PropBag.ReadProperty("Dark2", &H8000000F)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Picture2.BackColor = PropBag.ReadProperty("TopColor", &H8000000F)
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
    Call PropBag.WriteProperty("FrameWidth", Picture1.ScaleWidth, 1755)
    Call PropBag.WriteProperty("FrameHeight", Picture1.ScaleHeight, 915)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "CoolMan")
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("Light1", m_Light1, m_def_Light1)
'    Call PropBag.WriteProperty("Light2", m_Light2, m_def_Light2)
'    Call PropBag.WriteProperty("Dark1", m_Dark1, m_def_Dark1)
'    Call PropBag.WriteProperty("Dark2", m_Dark2, m_def_Dark2)
    Call PropBag.WriteProperty("Light1", Label3.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Light2", Label4.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Dark1", Label5.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Dark2", Label6.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("TopColor", Picture2.BackColor, &H8000000F)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=Picture1,Picture1,-1,Appearance
''Public Property Get Appearance() As Integer
''    Appearance = Picture1.Appearance
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Picture1,Picture1,-1,Appearance
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=Picture1,Picture1,-1,BackColor
''Public Property Get BackColor() As OLE_COLOR
''    BackColor = Picture1.BackColor
''End Property
''
''Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
''    Picture1.BackColor() = New_BackColor
''    PropertyChanged "BackColor"
''End Property
''
'Public Property Get Light1() As Variant
'    Light1 = m_Light1
'End Property
'
'Public Property Let Light1(ByVal New_Light1 As Variant)
'    m_Light1 = New_Light1
'    PropertyChanged "Light1"
'End Property
'
'Public Property Get Light2() As Variant
'    Light2 = m_Light2
'End Property
'
'Public Property Let Light2(ByVal New_Light2 As Variant)
'    m_Light2 = New_Light2
'    PropertyChanged "Light2"
'End Property
'
'Public Property Get Dark1() As Variant
'    Dark1 = m_Dark1
'End Property
'
'Public Property Let Dark1(ByVal New_Dark1 As Variant)
'    m_Dark1 = New_Dark1
'    PropertyChanged "Dark1"
'End Property
'
'Public Property Get Dark2() As Variant
'    Dark2 = m_Dark2
'End Property
'
'Public Property Let Dark2(ByVal New_Dark2 As Variant)
'    m_Dark2 = New_Dark2
'    PropertyChanged "Dark2"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Light1 = m_def_Light1
'    m_Light2 = m_def_Light2
'    m_Dark1 = m_def_Dark1
'    m_Dark2 = m_def_Dark2
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,BackColor
Public Property Get Light1() As OLE_COLOR
Attribute Light1.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Light1 = Label3.BackColor
End Property

Public Property Let Light1(ByVal New_Light1 As OLE_COLOR)
    Label3.BackColor() = New_Light1
    UserControl_Resize
    PropertyChanged "Light1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label4,Label4,-1,BackColor
Public Property Get Light2() As OLE_COLOR
Attribute Light2.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Light2 = Label4.BackColor
End Property

Public Property Let Light2(ByVal New_Light2 As OLE_COLOR)
    Label4.BackColor() = New_Light2
    UserControl_Resize
    PropertyChanged "Light2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label5,Label5,-1,BackColor
Public Property Get Dark1() As OLE_COLOR
Attribute Dark1.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Dark1 = Label5.BackColor
End Property

Public Property Let Dark1(ByVal New_Dark1 As OLE_COLOR)
    Label5.BackColor() = New_Dark1
    UserControl_Resize
    PropertyChanged "Dark1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label6,Label6,-1,BackColor
Public Property Get Dark2() As OLE_COLOR
Attribute Dark2.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Dark2 = Label6.BackColor
End Property

Public Property Let Dark2(ByVal New_Dark2 As OLE_COLOR)
    Label6.BackColor() = New_Dark2
    UserControl_Resize
    PropertyChanged "Dark2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Label1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture2,Picture2,-1,BackColor
Public Property Get TopColor() As OLE_COLOR
Attribute TopColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    TopColor = Picture2.BackColor
End Property

Public Property Let TopColor(ByVal New_TopColor As OLE_COLOR)
    Picture2.BackColor() = New_TopColor
    Picture3.BackColor() = New_TopColor
    PropertyChanged "TopColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get FrameWidth() As Single
Attribute FrameWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    FrameWidth = UserControl.ScaleWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get FrameHeight() As Single
Attribute FrameHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    FrameHeight = UserControl.ScaleHeight
End Property

