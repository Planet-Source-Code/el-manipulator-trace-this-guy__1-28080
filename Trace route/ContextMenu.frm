VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   1575
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1095
      Top             =   120
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Data"
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
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   1245
      Width           =   1260
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Data"
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
      Height          =   240
      Left            =   195
      TabIndex        =   8
      Top             =   1260
      Width           =   1260
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Plugin"
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
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   945
      Width           =   1260
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Plugin"
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
      Height          =   240
      Left            =   195
      TabIndex        =   6
      Top             =   960
      Width           =   1260
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00755433&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   60
      Top             =   1215
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00755433&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   60
      Top             =   915
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Ports"
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
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   645
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Ports"
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
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   660
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00755433&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   60
      Top             =   615
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QOS watcher"
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
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   345
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "QOS watcher"
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
      Height          =   240
      Left            =   195
      TabIndex        =   2
      Top             =   360
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00755433&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   45
      Top             =   315
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   1560
      X2              =   1560
      Y1              =   0
      Y2              =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   -15
      X2              =   1560
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PingThisGuy"
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
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PingThisGuy"
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
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00755433&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   30
      Top             =   15
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape1.Visible = False Then
  Call remap
  Shape1.Visible = True
 End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Comm$ = "PingThisGuy.exe /PING " + SelectedIp$
 If Dir$("PingThisGuy.exe") <> "" Then
  Shell Comm$, vbNormalFocus
 Else
  Call NotFound
 End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form2.ListView1.ListItems.Clear
 If AutoRedraw = True Then
  Call Form4.redo
 End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape3.Visible = False Then
  Call remap
  Shape3.Visible = True
 End If
End Sub

Sub remap()
 If Shape1.Visible = True Then Shape1.Visible = False
 If Shape3.Visible = True Then Shape3.Visible = False
 If Shape4.Visible = True Then Shape4.Visible = False
 If Shape5.Visible = True Then Shape5.Visible = False
 If Shape6.Visible = True Then Shape6.Visible = False
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Comm$ = "QOS TimeWatcher.exe /HOST " + SelectedIp$
 If Dir$("QOS TimeWatcher.exe") <> "" Then
  Shell Comm$, vbNormalFocus
 Else
  Call NotFound
 End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Comm$ = "Local Port Sniffer.exe"
 If Dir$("Local Port Sniffer.exe") <> "" Then
  Shell Comm$, vbNormalFocus
 Else
  Call NotFound
 End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape5.Visible = False Then
  Call remap
  Shape5.Visible = True
 End If
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form4.Show
End Sub

Private Sub Timer1_Timer()
 Dim Cursor As POINTAPI
 Call GetCursorPos(Cursor)
 If (Cursor.X * 15) < Me.Left Or (Cursor.X * 15) > Me.Left + Me.Width Or (Cursor.Y * 15) < Me.Top Or (Cursor.Y * 15) > Me.Top + Me.Height Then
  Unload Me
 End If
End Sub

Sub NotFound()
 Unload Me
 Form1.Label3.Caption = "Sorry but i'm not able to find required appli."
 Form1.Label4.Caption = "Sorry but i'm not able to find required appli."
 Form1.Label5.Caption = "on your computer..."
 Form1.Label6.Caption = "on your computer..."
 Form2.Show
End Sub
