VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{570928AD-1209-11D3-967B-B4129805661E}#5.0#0"; "CSTRAY.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4140
      Top             =   1965
   End
   Begin csTrayOCX.csTray csTray1 
      Left            =   4110
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "Form2.frx":08CA
      ToolTip         =   "Manipulator - TraceThisGuy"
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2055
      TabIndex        =   13
      Text            =   "2000"
      Top             =   765
      Width           =   510
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4185
      TabIndex        =   10
      Text            =   "255"
      Top             =   360
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1470
      TabIndex        =   7
      Text            =   "213.36.119.66"
      Top             =   360
      Width           =   1650
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   3810
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hop"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Router/Gateway/Host"
         Object.Width           =   3828
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "#1"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "#2"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "#3"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Label Label16 
      Caption         =   "2"
      Height          =   285
      Left            =   540
      TabIndex        =   18
      Top             =   3855
      Width           =   1605
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   45
      Picture         =   "Form2.frx":11A4
      Top             =   3780
      Width           =   210
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore host pinging ( if the host is behind a firewall )"
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
      Left            =   345
      TabIndex        =   17
      Top             =   3420
      Width           =   4365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore host pinging ( if the host is behind a firewall )"
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
      Left            =   360
      TabIndex        =   16
      Top             =   3435
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   75
      Picture         =   "Form2.frx":16C6
      Top             =   3420
      Width           =   210
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   1455
      Top             =   345
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   4170
      Top             =   345
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   2040
      Top             =   750
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4950
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   4920
      X2              =   4920
      Y1              =   0
      Y2              =   3705
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trace this guy!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2835
      TabIndex        =   15
      Top             =   795
      Width           =   1890
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trace this guy!"
      Height          =   255
      Left            =   2850
      TabIndex        =   14
      Top             =   810
      Width           =   1890
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00755433&
      X1              =   4725
      X2              =   4725
      Y1              =   765
      Y2              =   1035
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00755433&
      X1              =   2850
      X2              =   4740
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Shape Shape3 
      Height          =   300
      Left            =   2835
      Top             =   750
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   2850
      Top             =   765
      Width           =   1890
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle Timeout ( in ms ) :"
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
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle Timeout ( in ms ) :"
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
      Left            =   135
      TabIndex        =   11
      Top             =   795
      Width           =   1920
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MaxHops :"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   405
      Width           =   915
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MaxHops :"
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
      Left            =   3255
      TabIndex        =   8
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Host IP :"
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
      Left            =   90
      TabIndex        =   6
      Top             =   390
      Width           =   1995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Host IP :"
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
      Left            =   105
      TabIndex        =   5
      Top             =   405
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   30
      Width           =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   1
      Top             =   30
      Width           =   210
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   4845
      X2              =   4845
      Y1              =   60
      Y2              =   210
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   4695
      X2              =   4860
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   4620
      X2              =   4620
      Y1              =   45
      Y2              =   195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   4470
      X2              =   4635
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   4710
      Picture         =   "Form2.frx":1BE8
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   4485
      Picture         =   "Form2.frx":1D26
      Top             =   60
      Width           =   135
   End
   Begin VB.Shape Shape4 
      Height          =   195
      Left            =   4680
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape6 
      Height          =   195
      Left            =   4455
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   4470
      Top             =   45
      Width           =   165
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   4695
      Top             =   45
      Width           =   165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - TraceThisGuy"
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
      Left            =   60
      TabIndex        =   3
      Top             =   15
      Width           =   4365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - TraceThisGuy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   30
      Width           =   4365
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
      Left            =   -915
      Picture         =   "Form2.frx":1E64
      Top             =   15
      Width           =   5850
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   3450
      Left            =   0
      Top             =   255
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------
' -                        TRACE ROUTE                       -
' ------------------------------------------------------------
' - Auteur : RazorBack                                       -
' - Date   : 05/01/2001                                      -
' - URL    : http://manipulator.citeweb.net                  -
' - @Email : manipulator@citeweb.net                         -
' ------------------------------------------------------------
' -    Permet de tracer la route  entre votre  ordinateur et -
' - n'importe qu'elle machine d'un réseau joignable. Chaques -
' - sauts de routeurs sont analysés.                         -
' ------------------------------------------------------------

Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Private Sub csTray1_MouseUp(Button As Integer)
 ' On retire l'icone de la barre des taches et on restaure la fenêtre
 csTray1.Visible = False
 Me.WindowState = 0
 Me.Visible = True
 Label16.Caption = 2
 
 If AutoUpdate = True Then Form4.Show
End Sub

Private Sub Form_Load()
 IgnoreFirstPing = True
 AutoUpdate = False
 Ancre = False
 
 If Left$(Command$, 7) = "/TRACE " Then
  IP$ = Right$(Command$, Len(Command$) - 7)
  Text1.Text = IP$
  Form2.Refresh
  DoEvents
  If Text1.Text = "0.0.0.0" Then
   Form1.Label3.Caption = "Sorry but the IP you have entered"
   Form1.Label4.Caption = "Sorry but the IP you have entered"
   Form1.Label5.Caption = "seems to be invalid."
   Form1.Label6.Caption = "seems to be invalid."
   Form1.Show
  Else
   Call TraceThisGuy
  End If
 Else
  Me.Top = Int((Screen.Height - Me.Height) / 2)
  Me.Left = Int((Screen.Width - Me.Width) / 2)
 End If
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
 
 IgnoreFirstPing = False
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
 
 IgnoreFirstPing = True
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape8.BorderColor
 Line7.BorderColor = Shape8.BorderColor
 Shape8.BorderColor = Value
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape3.Visible = False Then
  Call remap
  Shape3.Visible = True
 End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape8.BorderColor
 Line7.BorderColor = Shape8.BorderColor
 Shape8.BorderColor = Value
 
 If Text1.Text = "0.0.0.0" Then
  Form1.Label3.Caption = "Sorry but the IP you have entered"
  Form1.Label4.Caption = "Sorry but the IP you have entered"
  Form1.Label5.Caption = "seems to be invalid."
  Form1.Label6.Caption = "seems to be invalid."
  Form1.Show
 Else
  Call TraceThisGuy
 End If
End Sub
Sub TraceThisGuy()
 ' Procédure de trace.
 ' Envois de paquets icmp avec incrément du TTL
 ' La trace est terminée quand l'Host est atteind
 
 
 If IgnoreFirstPing = False Then
  ' Test si l'host est déterminable avec un TimeOut de 4s au cas de grosses latences
  Value = PingIp(Text1.Text, 255, 4000)
  If Value = -1 Then
   ' Il ne peut pas être pinger, ça craint...
   Form1.Label3.Caption = "Sorry but the host you want to trace"
   Form1.Label4.Caption = "Sorry but the host you want to trace"
   Form1.Label5.Caption = "seems to be unreachable. (>Tout)"
   Form1.Label6.Caption = "seems to be unreachable. (>Tout)"
   Form1.Show
   Exit Sub
  End If
 End If
 
 ' Vide la liste
 ListView1.ListItems.Clear
 
 ' Effectue la résolution
 Good = True
 For TTL = 1 To Val(Text2.Text)
  ' Glande pour ne pas bloquer l'interface
  DoEvents
  Value = PingIp(Text1.Text, Val(TTL), Val(Text3.Text))
  If Value = 0 Or Value = 1 Then
   ' Ajout à la liste de l'élément
   Set lvItem = Form2.ListView1.ListItems.Add(, , Str$(TTL))
   lvItem.SubItems(1) = ReturnedIP$
   ' Analyse du temps de réponse
   If RoundTime(0) <> -1 Then
    lvItem.SubItems(2) = Str$(RoundTime(0))
   Else
    lvItem.SubItems(2) = "-"
   End If
   If RoundTime(1) <> -1 Then
    lvItem.SubItems(3) = Str$(RoundTime(1))
   Else
    lvItem.SubItems(3) = "-"
   End If
   If RoundTime(2) <> -1 Then
    lvItem.SubItems(4) = Str$(RoundTime(2))
   Else
    lvItem.SubItems(4) = "-"
   End If
   If AutoUpdate = True Then Call Form4.redo
  End If
  If Value = -1 And TTL <> 1 Then
   ' Probablement un Firewall qui bloque les réponse icmp
   ' Ajout à la liste de l'élément inconnu
   Set lvItem = Form2.ListView1.ListItems.Add(, , Str$(TTL))
   lvItem.SubItems(1) = "Unknow"
   ' Temps de réponse inconnu
   lvItem.SubItems(2) = "?"
   lvItem.SubItems(3) = "?"
   lvItem.SubItems(4) = "?"
  End If
  If Value = -1 And TTL = 1 Then
   ' Impossible d'envoyer des paquets à la destination...
   Good = False
   Exit For
  End If
  If Value = 1 Then Exit For
 Next TTL
 If Good = True Then
  ' Sa roule
  Form1.Label3.Caption = "Trace complete! [" + Str$(TTL - 1) + " hops ]"
  Form1.Label4.Caption = "Trace complete! [" + Str$(TTL - 1) + " hops ]"
  Form1.Label5.Caption = ""
  Form1.Label6.Caption = ""
 Else
  ' Problème avec les trames ICMP
  Form1.Label3.Caption = "Your system seems to don't be able to"
  Form1.Label4.Caption = "Your system seems to don't be able to"
  Form1.Label5.Caption = "send ICMP packets on target host. (>Tout)"
  Form1.Label6.Caption = "send ICMP packets on target host. (>Tout)"
 End If
 Form1.Show
End Sub

Private Sub Label14_Click()
 If IgnoreFirstPing = True Then
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
  IgnoreFirstPing = False
 Else
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
  IgnoreFirstPing = True
 End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 1
 X_Initial = X
 Y_Initial = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + X - X_Initial
  Me.Top = Me.Top + Y - Y_Initial
  If Ancre = True Then
   Form4.Left = Form4.Left + X - X_Initial
   Form4.Top = Form4.Top + Y - Y_Initial
  End If
 Else
  Call remap
 End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 0
 Dist_Am = 100
 
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height

 If Ancre = True Then
  Form4.Top = Form2.Top + Form2.Height
  Form4.Left = Form2.Left
 End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value

 End
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Sub remap()
 If Shape3.Visible = True Then Shape3.Visible = False
 If Shape4.Visible = True Then Shape4.Visible = False
 If Shape6.Visible = True Then Shape6.Visible = False
 If Shape9.Visible = True Then Shape9.Visible = False
 If Shape10.Visible = True Then Shape10.Visible = False
 If Shape11.Visible = True Then Shape11.Visible = False
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
 
 Me.WindowState = 1
 Me.Visible = False
 If AutoUpdate = True Then Unload Form4
  
 csTray1.Visible = True
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 And ListView1.ListItems.Count <> 0 Then
  SelectedIp$ = ListView1.SelectedItem.ListSubItems(1)
  Dim Cursor As POINTAPI
  Call GetCursorPos(Cursor)
  Form5.Top = (Cursor.Y * 15) - 50
  Form5.Left = (Cursor.X * 15) - 200
  Form5.Show
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call TraceThisGuy
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape11.Visible = False Then
  Call remap
  Shape11.Visible = True
 End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape10.Visible = False Then
  Call remap
  Shape10.Visible = True
 End If
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape9.Visible = False Then
  Call remap
  Shape9.Visible = True
 End If
End Sub

Private Sub Timer1_Timer()
 If Val(Label16.Caption) <> 0 Then
  SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
  Label16.Caption = Val(Label16.Caption - 1)
 End If
End Sub
