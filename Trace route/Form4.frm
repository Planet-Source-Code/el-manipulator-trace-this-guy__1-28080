VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4365
      Top             =   345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   30
      Picture         =   "Form4.frx":08CA
      ScaleHeight     =   1635
      ScaleWidth      =   4830
      TabIndex        =   5
      Top             =   285
      Width           =   4860
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   15
         TabIndex        =   11
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   15
         TabIndex        =   10
         Top             =   690
         Width           =   630
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   15
         TabIndex        =   9
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   8
         Top             =   1095
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   7
         Top             =   705
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   6
         Top             =   285
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Update Mode"
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
      Left            =   375
      TabIndex        =   13
      Top             =   2010
      Width           =   3105
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Update Mode"
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
      Left            =   390
      TabIndex        =   12
      Top             =   2025
      Width           =   3105
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   60
      Picture         =   "Form4.frx":4E7C
      Top             =   2745
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   105
      Picture         =   "Form4.frx":539E
      Top             =   1995
      Width           =   210
   End
   Begin VB.Label Label16 
      Caption         =   "Label7"
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   2400
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
      Left            =   3645
      TabIndex        =   3
      Top             =   2010
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
      Left            =   3660
      TabIndex        =   2
      Top             =   2025
      Width           =   1185
   End
   Begin VB.Shape Shape8 
      Height          =   300
      Left            =   3630
      Top             =   1965
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   4845
      X2              =   4845
      Y1              =   1980
      Y2              =   2250
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   3645
      X2              =   4860
      Y1              =   2235
      Y2              =   2235
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   3645
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   4920
      X2              =   4920
      Y1              =   -15
      Y2              =   2295
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
      Height          =   2040
      Left            =   0
      Top             =   255
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TraceThisGuy - Datas'n'Graph plugin"
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
      Caption         =   "TraceThisGuy - Datas'n'Graph plugin"
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
      Picture         =   "Form4.frx":58C0
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Dim quad As Boolean

Private Sub Form_Load()
 Me.Top = Form2.Top + Form2.Height
 Me.Left = Form2.Left
 Ancre = True
 
 Label16.Caption = 20
 AutoUpdate = True

 Call redo
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Image1_Click()
 Y = Image1.Top
 X = Image1.Left
 Image1.Visible = False
 Image1.Top = Image4.Top
 Image1.Left = Image4.Left
 Image1.Visible = True
 
 Image4.Visible = False
 Image4.Top = Y
 Image4.Left = X
 Image4.Visible = True
 
 AutoUpdate = False
End Sub

Private Sub Image4_Click()
 Y = Image4.Top
 X = Image4.Left
 Image4.Visible = False
 Image4.Top = Image1.Top
 Image4.Left = Image1.Left
 Image4.Visible = True
 
 Image1.Visible = False
 Image1.Top = Y
 Image1.Left = X
 Image1.Visible = True
 
 AutoUpdate = True
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
  Ancre = False
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

Private Sub Label10_Click()
 If AutoUpdate = True Then
  Y = Image1.Top
  X = Image1.Left
  Image1.Visible = False
  Image1.Top = Image4.Top
  Image1.Left = Image4.Left
  Image1.Visible = True
  Image4.Visible = False
  Image4.Top = Y
  Image4.Left = X
  Image4.Visible = True
  AutoUpdate = False
 Else
  Y = Image4.Top
  X = Image4.Left
  Image4.Visible = False
  Image4.Top = Image1.Top
  Image4.Left = Image1.Left
  Image4.Visible = True
  Image1.Visible = False
  Image1.Top = Y
  Image1.Left = X
  Image1.Visible = True
  AutoUpdate = True
 End If
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

 AutoUpdate = False
 Ancre = False

 Unload Me
End Sub

Sub remap()
 If Shape8.Visible = True Then Shape8.Visible = False
 If Label3.Visible = True Then Label3.Visible = False
 If Label4.Visible = True Then Label4.Visible = False
 If Label5.Visible = True Then Label5.Visible = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label3.Visible = False Then
  Call remap
  Label3.Visible = True
 End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label4.Visible = False Then
  Call remap
  Label4.Visible = True
 End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label5.Visible = False Then
  Call remap
  Label5.Visible = True
 End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If quad = False Then
  quad = True
 Else
  quad = False
 End If
 
 Call redo
End Sub

Private Sub Timer1_Timer()
 If Val(Label16.Caption) <> 0 Then
  SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
  Label16.Caption = Val(Label16.Caption - 1)
 End If
End Sub

Public Sub redo()
 Dim WorkFlow(200)
 
 w = 0
 For i = 1 To Form2.ListView1.ListItems.Count
  For t = 2 To 4
   If Val(Form2.ListView1.ListItems(i).ListSubItems(t)) <> 0 Then
    If w < 200 Then WorkFlow(w) = Val(Form2.ListView1.ListItems(i).ListSubItems(t))
    If Val(Form2.ListView1.ListItems(i).ListSubItems(t)) > Imax Then Imax = Val(Form2.ListView1.ListItems(i).ListSubItems(t))
    w = w + 1
   End If
  Next t
 Next i
 If w > 200 Then w = 200
 If w <= 1 Then Exit Sub

 Picture1.Cls
 
 Label3.Caption = Imax / 2 + Imax / 4
 Label4.Caption = Imax / 2
 Label5.Caption = Imax / 4
 
 Imax = Imax + (Imax / 10)
 Xstep = Picture1.Width / (w - 1)
 YScale = Picture1.Height / Imax
 
 Moy = 0
 For i = 0 To w - 1
  Moy = Moy + WorkFlow(i)
 Next i
 Moy = Moy / w
 Picture1.Line (0, Picture1.Height - (Moy * YScale))-(Picture1.Width, Picture1.Height - (Moy * YScale)), RGB(170, 0, 0)
 
 For i = 0 To w - 2
  If i <> 0 And quad = True Then Picture1.Line (i * Xstep, 0)-(i * Xstep, Picture1.Height), RGB(120, 120, 120)
  Picture1.Line (i * Xstep, Picture1.Height - (WorkFlow(i) * YScale))-(i * Xstep + Xstep, Picture1.Height - (WorkFlow(i + 1) * YScale)), RGB(0, 0, 170)
  Picture1.Line (i * Xstep - 15, Picture1.Height - (WorkFlow(i) * YScale) - 15)-(i * Xstep + 15, Picture1.Height - (WorkFlow(i) * YScale) + 15), RGB(0, 0, 0), BF
 Next i
End Sub
