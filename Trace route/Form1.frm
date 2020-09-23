VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3945
      Top             =   375
   End
   Begin VB.Label Label16 
      Caption         =   "Label7"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "That's Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3225
      TabIndex        =   7
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "That's Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3240
      TabIndex        =   6
      Top             =   1245
      Width           =   1185
   End
   Begin VB.Shape Shape8 
      Height          =   300
      Left            =   3210
      Top             =   1185
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   4425
      X2              =   4425
      Y1              =   1200
      Y2              =   1470
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   3225
      X2              =   4440
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   3225
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "#2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   900
      TabIndex        =   5
      Top             =   690
      Width           =   3525
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "#2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   915
      TabIndex        =   4
      Top             =   705
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   885
      TabIndex        =   3
      Top             =   480
      Width           =   3630
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   900
      TabIndex        =   2
      Top             =   495
      Width           =   3630
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   150
      Picture         =   "Form1.frx":0000
      Top             =   405
      Width           =   510
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   -15
      X2              =   4575
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   0
      Top             =   255
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TraceThisGuy - Notification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   15
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TraceThisGuy - Notification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   -930
      Picture         =   "Form1.frx":090A
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Private Sub Form_Load()
 Label16.Caption = 20
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 1
 X_Initial = X
 Y_Initial = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + X - X_Initial
  Me.Top = Me.Top + Y - Y_Initial
 Else
  Call remap
 End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 0
 Dist_Am = 100
 
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape3.BorderColor
 Line9.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape8.Visible = False Then
  Call remap
  Shape8.Visible = True
 End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape3.BorderColor
 Line9.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value

 ReadyToShow = True
 Unload Me
End Sub

Sub remap()
 If Shape8.Visible = True Then Shape8.Visible = False
End Sub

Private Sub Timer1_Timer()
 If Val(Label16.Caption) <> 0 Then
  SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
  Label16.Caption = Val(Label16.Caption - 1)
 End If
End Sub
