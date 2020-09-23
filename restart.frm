VERSION 5.00
Begin VB.Form restart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   465
   ClientLeft      =   -5880
   ClientTop       =   3075
   ClientWidth     =   435
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   15
   End
End
Attribute VB_Name = "restart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    tedred.Show
    Timer1.Enabled = False
    Unload Me
End Sub
