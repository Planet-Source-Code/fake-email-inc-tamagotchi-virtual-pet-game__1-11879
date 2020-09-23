VERSION 5.00
Begin VB.Form tedred 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Ted Red Virtual Pet"
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   ControlBox      =   0   'False
   Icon            =   "tedred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer cake2 
      Enabled         =   0   'False
      Left            =   2565
      Top             =   4725
   End
   Begin VB.Timer cake 
      Enabled         =   0   'False
      Left            =   2010
      Top             =   4725
   End
   Begin VB.Timer feel 
      Interval        =   1000
      Left            =   1530
      Top             =   4725
   End
   Begin VB.Timer eat2 
      Enabled         =   0   'False
      Left            =   1050
      Top             =   4710
   End
   Begin VB.Timer eat 
      Enabled         =   0   'False
      Left            =   525
      Top             =   4755
   End
   Begin VB.Timer healthtimer 
      Interval        =   25000
      Left            =   45
      Top             =   4770
   End
   Begin VB.Timer drink 
      Enabled         =   0   'False
      Left            =   30
      Top             =   5220
   End
   Begin VB.Timer healthdrop 
      Interval        =   8000
      Left            =   525
      Top             =   5205
   End
   Begin VB.Timer beat2 
      Enabled         =   0   'False
      Left            =   960
      Top             =   5190
   End
   Begin VB.Timer beat 
      Enabled         =   0   'False
      Left            =   1425
      Top             =   5190
   End
   Begin VB.Timer bodymove 
      Interval        =   400
      Left            =   1875
      Top             =   5130
   End
   Begin VB.Timer shroomt3 
      Enabled         =   0   'False
      Left            =   2370
      Top             =   5190
   End
   Begin VB.Timer shroomt2 
      Enabled         =   0   'False
      Left            =   75
      Top             =   5775
   End
   Begin VB.Timer shroomt1 
      Enabled         =   0   'False
      Left            =   600
      Top             =   5700
   End
   Begin VB.Timer needlet2 
      Enabled         =   0   'False
      Left            =   1020
      Top             =   5655
   End
   Begin VB.Timer needlet1 
      Enabled         =   0   'False
      Left            =   1455
      Top             =   5640
   End
   Begin VB.Timer jointt3 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   5670
   End
   Begin VB.Timer jointt2 
      Enabled         =   0   'False
      Left            =   2385
      Top             =   5655
   End
   Begin VB.Timer jointt1 
      Enabled         =   0   'False
      Left            =   150
      Top             =   6180
   End
   Begin VB.Timer mast3 
      Enabled         =   0   'False
      Left            =   615
      Top             =   6225
   End
   Begin VB.Timer mast 
      Enabled         =   0   'False
      Left            =   1110
      Top             =   6255
   End
   Begin VB.ListBox List1 
      BackColor       =   &H008080FF&
      Height          =   1425
      ItemData        =   "tedred.frx":0CCA
      Left            =   60
      List            =   "tedred.frx":0CF5
      TabIndex        =   3
      Top             =   2625
      Width           =   2940
   End
   Begin VB.Label gameover 
      BackStyle       =   0  'Transparent
      Caption         =   "Game over.. Click here to play again, or click on the X in the top left corner to exit."
      Height          =   615
      Left            =   195
      TabIndex        =   7
      Top             =   330
      Width           =   2355
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1995
      Picture         =   "tedred.frx":0E29
      Top             =   1260
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   1320
      Left            =   450
      Picture         =   "tedred.frx":1133
      Top             =   975
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.fakeemail.org"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   555
      TabIndex        =   6
      Top             =   4350
      Width           =   2190
   End
   Begin VB.Image slice 
      Height          =   480
      Left            =   1875
      Picture         =   "tedred.frx":3444
      Top             =   1275
      Width           =   480
   End
   Begin VB.Label feeling 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   4095
      Width           =   3105
   End
   Begin VB.Image burger 
      Height          =   480
      Left            =   1965
      Picture         =   "tedred.frx":374E
      Top             =   1275
      Width           =   480
   End
   Begin VB.Image water 
      Height          =   480
      Left            =   1815
      Picture         =   "tedred.frx":4018
      Top             =   1110
      Width           =   480
   End
   Begin VB.Label healthbar 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   4
      Top             =   2370
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2490
      Picture         =   "tedred.frx":4322
      Top             =   285
      Width           =   480
   End
   Begin VB.Image glove1 
      Height          =   480
      Left            =   90
      Picture         =   "tedred.frx":462C
      Top             =   705
      Width           =   480
   End
   Begin VB.Image glove2 
      Height          =   480
      Left            =   495
      Picture         =   "tedred.frx":4936
      Top             =   720
      Width           =   480
   End
   Begin VB.Image body2 
      Height          =   480
      Left            =   1380
      Picture         =   "tedred.frx":4C40
      Top             =   1830
      Width           =   480
   End
   Begin VB.Image shroom 
      Height          =   480
      Left            =   1980
      Picture         =   "tedred.frx":4F4A
      Top             =   1305
      Width           =   480
   End
   Begin VB.Image needle2 
      Height          =   480
      Left            =   1725
      Picture         =   "tedred.frx":5254
      Top             =   1815
      Width           =   480
   End
   Begin VB.Image needle1 
      Height          =   480
      Left            =   1725
      Picture         =   "tedred.frx":555E
      Top             =   1815
      Width           =   480
   End
   Begin VB.Image redeye2 
      Height          =   480
      Left            =   1575
      Picture         =   "tedred.frx":5868
      Top             =   795
      Width           =   480
   End
   Begin VB.Image redeye1 
      Height          =   480
      Left            =   1140
      Picture         =   "tedred.frx":6532
      Top             =   810
      Width           =   480
   End
   Begin VB.Image joint3 
      Height          =   480
      Left            =   1815
      Picture         =   "tedred.frx":71FC
      Top             =   1275
      Width           =   480
   End
   Begin VB.Image joint2 
      Height          =   480
      Left            =   1815
      Picture         =   "tedred.frx":7506
      Top             =   1275
      Width           =   480
   End
   Begin VB.Image joint1 
      Height          =   480
      Left            =   1815
      Picture         =   "tedred.frx":7810
      Top             =   1275
      Width           =   480
   End
   Begin VB.Image mast2 
      Height          =   480
      Left            =   1380
      Picture         =   "tedred.frx":7B1A
      Top             =   1830
      Width           =   480
   End
   Begin VB.Image mast1 
      Height          =   480
      Left            =   1380
      Picture         =   "tedred.frx":7E24
      Top             =   1830
      Width           =   480
   End
   Begin VB.Image body 
      Height          =   480
      Left            =   1380
      Picture         =   "tedred.frx":812E
      Top             =   1830
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   840
      Picture         =   "tedred.frx":8438
      ToolTipText     =   "Ted Red"
      Top             =   315
      Width           =   1320
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3030
      X2              =   2895
      Y1              =   30
      Y2              =   165
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2910
      X2              =   3030
      Y1              =   45
      Y2              =   165
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2715
      X2              =   2835
      Y1              =   165
      Y2              =   165
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Height          =   180
      Left            =   2895
      TabIndex        =   0
      Top             =   15
      Width           =   150
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "_"
      Height          =   180
      Left            =   2700
      TabIndex        =   1
      Top             =   15
      Width           =   165
   End
   Begin VB.Label myTitle 
      BackColor       =   &H000000C0&
      Caption         =   " Ted Red Virtual Pet"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   1560
      Left            =   840
      Picture         =   "tedred.frx":A858
      Top             =   315
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2070
      Left            =   75
      Top             =   270
      Width           =   2925
   End
End
Attribute VB_Name = "tedred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MouseStartX As Long
Public MouseStartY As Long

Private Sub beat_Timer()
    If glove1.Visible = True Then
        glove1.Visible = False
        glove2.Visible = True
        Image2.Visible = True
        Image1.Visible = False
        Exit Sub
    End If
    If glove2.Visible = True Then
        glove2.Visible = False
        glove1.Visible = True
        Image2.Visible = False
        Image1.Visible = True
    End If
End Sub

Private Sub beat2_Timer()
    glove1.Visible = False
    glove2.Visible = False
    Image2.Visible = False
    Image1.Visible = True
    beat.Enabled = False
    beat2.Enabled = False
End Sub

Private Sub bodymove_Timer()
    If body.Visible = True Then
        body2.Visible = True
        body.Visible = False
        Exit Sub
    End If
    If body2.Visible = True Then
        body.Visible = True
        body2.Visible = False
        Exit Sub
    End If
End Sub

Private Sub cake_Timer()
    If Image1.Visible = True Then
        Image2.Visible = True
        Image1.Visible = False
        Exit Sub
    End If
    If Image2.Visible = True Then
        Image1.Visible = True
        Image2.Visible = False
    End If
End Sub

Private Sub cake2_Timer()
    Image2.Visible = False
    Image1.Visible = True
    cake.Enabled = False
    slice.Visible = False
    cake2.Enabled = False
End Sub

Private Sub drink_Timer()
    Image2.Visible = False
    Image1.Visible = True
    water.Visible = False
    drink.Enabled = False
End Sub

Private Sub eat_Timer()
    If Image1.Visible = True Then
        Image2.Visible = True
        Image1.Visible = False
        Exit Sub
    End If
    If Image2.Visible = True Then
        Image1.Visible = True
        Image2.Visible = False
    End If
End Sub

Private Sub eat2_Timer()
    eat.Enabled = False
    Image1.Visible = True
    Image2.Visible = False
    burger.Visible = False
    eat2.Enabled = False
End Sub

Private Sub feel_Timer()
    If hungry = 1 Then
        If thirsty = 1 Then
            feeling.Caption = " Ted is feeling hungry and thirsty!"
            Exit Sub
        Else
            feeling.Caption = " Ted is feeling hungry"
        End If
    End If
    If thirsty = 1 Then
        feeling.Caption = " Ted is thirsty!"
    End If
    If health < 50 Then
        feeling.Caption = " Ted is dying!!!"
        Exit Sub
    End If
    If hungry = 0 Then
        If thirsty = 0 Then
            feeling.Caption = " Ted is happy as a pig in shit"
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    mast1.Visible = False
    mast2.Visible = False
    joint1.Visible = False
    joint2.Visible = False
    joint3.Visible = False
    redeye1.Visible = False
    redeye2.Visible = False
    needle1.Visible = False
    needle2.Visible = False
    shroom.Visible = False
    body2.Visible = False
    glove1.Visible = False
    glove2.Visible = False
    water.Visible = False
    burger.Visible = False
    slice.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    gameover.Visible = False
    health = 100
    healthbar.Caption = " Ted's health: " & health & "%"
    feeling.Caption = " Ted is feeling: Great!"
End Sub

Private Sub gameover_Click()
    If gameover.Visible = False Then
        Exit Sub
    Else
        restart.Show
        Unload Me
    End If
End Sub

Private Sub healthdrop_Timer()
    health = health - 3
    If health < 1 Then
        healthtimer.Enabled = False
        eat.Enabled = False
        eat2.Enabled = False
        feel.Enabled = False
        bodymove.Enabled = False
        cake.Enabled = False
        cake2.Enabled = False
        drink.Enabled = False
        healthdrop.Enabled = False
        beat2.Enabled = False
        beat.Enabled = False
        shroomt3.Enabled = False
        shroomt2.Enabled = False
        shroomt1.Enabled = False
        needlet2.Enabled = False
        needlet1.Enabled = False
        jointt3.Enabled = False
        jointt2.Enabled = False
        jointt1.Enabled = False
        mast3.Enabled = False
        mast.Enabled = False
        List1.Enabled = False
        mast1.Visible = False
        mast2.Visible = False
        joint1.Visible = False
        joint2.Visible = False
        joint3.Visible = False
        redeye1.Visible = False
        redeye2.Visible = False
        needle1.Visible = False
        needle2.Visible = False
        shroom.Visible = False
        body2.Visible = False
        glove1.Visible = False
        glove2.Visible = False
        water.Visible = False
        burger.Visible = False
        slice.Visible = False
        Image4.Visible = True
        Image5.Visible = True
        Image1.Visible = False
        Image2.Visible = False
        body.Visible = False
        gameover.Visible = True
        feeling.Caption = "You killed him!!!!!!!"
        healthbar.Caption = " Ted's health: Deceased"
    Else
        healthbar.Caption = " Ted's health: " & health & "%"
    End If
End Sub

Private Sub healthtimer_Timer()
    thirsty = 1
    hungry = 1
End Sub

Private Sub jointt1_Timer()
    If joint1.Visible = True Then
        joint2.Visible = True
        joint1.Visible = False
        Exit Sub
    End If
    If joint2.Visible = True Then
        joint3.Visible = True
        joint2.Visible = False
    End If
End Sub

Private Sub jointt2_Timer()
    jointt1.Enabled = False
    joint3.Visible = False
    redeye1.Visible = True
    redeye2.Visible = True
    jointt3.Interval = 10000
    jointt3.Enabled = True
    jointt2.Enabled = False
End Sub

Private Sub jointt3_Timer()
    redeye1.Visible = False
    redeye2.Visible = False
    jointt3.Enabled = False
End Sub

Private Sub Label1_Click()
    End
End Sub

Private Sub Label2_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub List1_Click()
    If List1.ListIndex = 3 Then
        mast.Interval = 100
        mast.Enabled = True
        mast3.Interval = 5000
        mast3.Enabled = True
        bodymove.Enabled = False
        body2.Visible = False
        body.Visible = True
        If health < 100 Then
            health = health + 10
            If health > 100 Then
                health = 100
            End If
        End If
    End If
    If List1.ListIndex = 4 Then
        needlet1.Interval = 400
        needlet1.Enabled = True
        needlet2.Interval = 10000
        needlet2.Enabled = True
        needle1.Visible = True
        health = health - 7
    End If
    If List1.ListIndex = 5 Then
        shroomt1.Interval = 200
        shroomt1.Enabled = True
        shroomt2.Interval = 2000
        shroomt2.Enabled = True
        shroom.Visible = True
        health = health - 5
    End If
    If List1.ListIndex = 6 Then
        jointt1.Interval = 500
        jointt2.Interval = 1500
        jointt1.Enabled = True
        jointt2.Enabled = True
        joint1.Visible = True
        health = health - 3
    End If
    If List1.ListIndex = 8 Then
        beat.Interval = 700
        glove1.Visible = True
        beat.Enabled = True
        beat2.Interval = 6700
        beat2.Enabled = True
        health = health - 4
    End If
    If List1.ListIndex = 10 Then
        drink.Interval = 2000
        drink.Enabled = True
        Image2.Visible = True
        Image1.Visible = False
        water.Visible = True
        If thirsty = 1 Then
            thirsty = 0
            If health < 100 Then
                health = health + 9
                If health > 100 Then
                    health = 100
                End If
            End If
        Else
            health = health - 6
        End If
    End If
    If List1.ListIndex = 11 Then
        eat.Interval = 220
        eat.Enabled = True
        burger.Visible = True
        eat2.Interval = 3000
        eat2.Enabled = True
        If hungry = 1 Then
            hungry = 0
            If health < 100 Then
                health = health + 6
                If health > 100 Then
                    health = 100
                End If
            End If
        Else
            health = health - 5
        End If
    End If
    If List1.ListIndex = 12 Then
        cake.Interval = 220
        cake.Enabled = True
        slice.Visible = True
        cake2.Interval = 3000
        cake2.Enabled = True
        If hungry = 0 Then
            health = health - 2
        End If
    End If
    If health < 1 Then
        healthbar.Caption = " Ted's health: Terminally Ill"
    Else
        healthbar.Caption = " Ted's health: " & health & "%"
    End If
End Sub

Private Sub mast_Timer()
    If body.Visible = True Then
        mast1.Visible = True
        body.Visible = False
    End If
    If mast1.Visible = True Then
        mast2.Visible = True
        mast1.Visible = False
        Exit Sub
    End If
    If mast2.Visible = True Then
        mast1.Visible = True
        mast2.Visible = False
        Exit Sub
    End If
End Sub

Private Sub mast3_Timer()
    mast.Enabled = False
    mast1.Visible = False
    mast2.Visible = False
    body.Visible = True
    mast3.Enabled = False
    bodymove.Enabled = True
End Sub

Private Sub myTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStartX = X
    MouseStartY = Y
End Sub

Private Sub myTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        tedred.Left = tedred.Left - (MouseStartX - X)
        tedred.Top = tedred.Top - (MouseStartY - Y)
    End If
End Sub

Private Sub needlet1_Timer()
    If needle1.Visible = True Then
        needle2.Visible = True
        needle1.Visible = False
        Exit Sub
    End If
    If needle2.Visible = True Then
        needle2.Visible = False
        redeye1.Visible = True
        redeye2.Visible = True
        needlet1.Enabled = False
    End If
End Sub

Private Sub needlet2_Timer()
    redeye1.Visible = False
    redeye2.Visible = False
    needlet2.Enabled = False
End Sub

Private Sub shroomt1_Timer()
    If Image1.Visible = True Then
        Image2.Visible = True
        Image1.Visible = False
        Exit Sub
    End If
    If Image2.Visible = True Then
        Image1.Visible = True
        Image2.Visible = False
    End If
End Sub

Private Sub shroomt2_Timer()
    shroomt1.Enabled = False
    Image1.Visible = True
    Image2.Visible = False
    shroom.Visible = False
    redeye1.Visible = True
    redeye2.Visible = True
    shroomt3.Interval = 5000
    shroomt3.Enabled = True
    shroomt2.Enabled = False
End Sub

Private Sub shroomt3_Timer()
    redeye1.Visible = False
    redeye2.Visible = False
    shroomt3.Enabled = False
End Sub
