VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   -30
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   -30
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu HiddenMenu 
      Caption         =   "HiddenMenu"
      Begin VB.Menu ShowGraph 
         Caption         =   "Show Graph"
      End
      Begin VB.Menu ClearAll 
         Caption         =   "Clear Datas"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearAll_Click()
 Form2.ListView1.ListItems.Clear
 If AutoRedraw = True Then
  Call Form4.redo
 End If
End Sub

Private Sub ShowGraph_Click()
 Form4.Show
End Sub
