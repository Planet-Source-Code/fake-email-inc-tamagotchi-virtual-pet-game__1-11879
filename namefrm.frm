VERSION 5.00
Begin VB.Form namefrm 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   Icon            =   "namefrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   705
      TabIndex        =   2
      Top             =   780
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   60
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "Ted Red"
      Top             =   420
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose a name for Ted Red then press Ok"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   1875
   End
End
Attribute VB_Name = "namefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
        Beep
    Else
        tedredname = Text1.Text
        Me.Hide
        tedred.Show
    End If
End Sub
