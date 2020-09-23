VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Continue"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin lsaFrameControl.lsaFrame lsaFrame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      FrameWidth      =   101
      FrameHeight     =   69
      Caption         =   "About lsaFrame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483624
      Light1          =   -2147483632
      Light2          =   -2147483627
      Dark2           =   -2147483628
      BackColor       =   -2147483624
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2002 Lewis Software Australia."
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   2160
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub


