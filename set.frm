VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Set"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton nextdeck 
      Caption         =   "Reshuffle deck"
      Height          =   735
      Left            =   6960
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton endgame 
      Caption         =   "End game"
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Player 
      Caption         =   "Player 4"
      Height          =   735
      Index           =   4
      Left            =   5280
      TabIndex        =   18
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Player 
      Caption         =   "Player 3"
      Height          =   735
      Index           =   3
      Left            =   3600
      TabIndex        =   17
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Player 
      Caption         =   "Player 2"
      Height          =   735
      Index           =   2
      Left            =   1920
      TabIndex        =   16
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Player 
      Caption         =   "Player 1"
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   12
      Left            =   5040
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   12
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   12
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   12
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   11
      Left            =   5040
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   11
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   11
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   11
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   10
      Left            =   5040
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   10
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   10
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   10
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   9
      Left            =   3600
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   9
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   9
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   9
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   8
      Left            =   3600
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   8
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   8
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   8
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   7
      Left            =   3600
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   9
      Top             =   240
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   7
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   7
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   7
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   6
      Left            =   2160
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   6
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   6
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   6
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   5
      Left            =   2160
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   5
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   5
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   5
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   4
      Left            =   2160
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   4
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   4
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   4
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   3
      Left            =   720
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   3
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   3
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   3
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   2
      Left            =   720
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   2
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   2
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   2
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deal 1st 6 cards"
      Height          =   735
      Left            =   6840
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Put back 6 cards"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deal 3 more cards"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   1
      Left            =   720
      ScaleHeight     =   1785
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   1
         Left            =   120
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   3  'Vertical Line
         Height          =   375
         Index           =   1
         Left            =   120
         Shape           =   2  'Oval
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillStyle       =   4  'Upward Diagonal
         Height          =   375
         Index           =   1
         Left            =   120
         Shape           =   2  'Oval
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label Lblcardsindeck 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin VB.Menu Game 
      Caption         =   "Game"
      Index           =   0
      Begin VB.Menu New 
         Caption         =   "New"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu options 
         Caption         =   "Options..."
         Index           =   3
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Index           =   2
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Index           =   0
      Begin VB.Menu blah 
         Caption         =   "How to Play"
         Index           =   1
      End
      Begin VB.Menu showamountofsets 
         Caption         =   "Show amount of sets"
         Index           =   3
      End
      Begin VB.Menu hint 
         Caption         =   "Show all sets"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim card(3, 3, 3, 3)        'card(shape,color,shading,number)
Dim table(12) As String
Dim c(4)    'c means character
Dim cc As Integer 'highlighted cards
Dim deck
Dim cardsontable As Integer
Dim cardcheck(3)
Dim gamer(4)
Public gamer1 As Integer
Public gamer2 As Integer
Public gamer3 As Integer
Public gamer4 As Integer
Dim sets As Integer
Dim picked As Boolean
Dim mistakes(4) As Variant
Public mistake1 As Integer
Public mistake2 As Integer
Public mistake3 As Integer
Public mistake4 As Integer


Public gamestyle As String
Public displaypoints As Boolean
Public countmistakes As Boolean
Public displaymistakes As Boolean
Public displaysets As Boolean

Private Sub blah_Click(Index As Integer)
x = MsgBox("The object of the game is to find three cards whose four characteristics (number, color, shape, and shading) are each the same or different. For example, a set would be all cards green, all diffent shapes, all the same shading and all different amounts. Select the three cards that make up a set and then click on the button corresponding to the player who had found it. Enjoy! -Bumy Goldson, February 18, 2009.", 0, "How to Play")
End Sub



Private Sub Command2_Click()
For x = 7 To 12
card(Mid(table(x), 4, 1), Mid(table(x), 3, 1), Mid(table(x), 2, 1), Mid(table(x), 1, 1)) = 0
table(x) = ""
deck = deck + 1
Lblcardsindeck = "Cards in deck: " & deck

cardsontable = cardsontable - 1
Next
Call phil
Command1.Enabled = True
Command2.Enabled = False

If displaysets = True Then Call showset

End Sub

Private Sub Command3_Click()

'DEAL OUT SIX CARDS
For y = 1 To 6          'FOR SIX CARDS
    For x = 1 To 4      'PICK A RANDOM CARD
        Randomize Timer
        c(x) = Int((Rnd * 3) + 1)
    Next
    
    If card(c(1), c(2), c(3), c(4)) = 0 Then
       card(c(1), c(2), c(3), c(4)) = 1
       deck = deck - 1
       Lblcardsindeck = "Cards in deck: " & deck
       
        table(y) = c(1) & c(2) & c(3) & c(4)
        cardsontable = cardsontable + 1
    Else: y = y - 1
    End If
Next

Call phil
Command3.Visible = False
Command1.Enabled = True

If displaysets = True Then Call countsets

End Sub

Private Sub Command1_Click()

For y = 1 To 3          'FOR three CARDS
    For x = 1 To 4      'PICK RANDOM CHARACTERISTICS
        Randomize Timer
        c(x) = Int((Rnd * 3) + 1)
    Next
    
    If card(c(1), c(2), c(3), c(4)) = 0 Then
       card(c(1), c(2), c(3), c(4)) = 1
       deck = deck - 1
       Lblcardsindeck = "Cards in deck: " & deck
       
       For L = 1 To 12
       If table(L) = "" Then
            table(L) = c(1) & c(2) & c(3) & c(4)
            cardsontable = cardsontable + 1
            Picture1(L).Visible = True
            Exit For
        End If
        Next
    Else: y = y - 1
    End If
Next

Call phil
If cardsontable = 12 Then
    Command1.Enabled = False
    Command2.Enabled = True
End If

If deck = 0 Then
    Print "No more cards"
    Command1.Enabled = False
    nextdeck.Visible = True
End If

If displaysets = True Then Call countsets

End Sub



Private Sub Command5_Click()
Call showset
End Sub



Private Sub endgame_Click()
    For p = 1 To 4
        Player(p).Caption = "Player " & p & vbNewLine & gamer(p) & " point(s)"
        Player(p).Enabled = True
        If countmistakes = True Then Player(p).Caption = Player(p).Caption & " / " & vbNewLine & mistakes(p) & " mistake(s)"
    Next
    
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    endgame.Visible = False
End Sub

Private Sub exit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

For x = 1 To 3
    If Form2.GSO(x) = True Then gamestyle = x
Next

If Form2.Checkdp = 1 Then displaypoints = True Else: displaypoints = False
If Form2.Checkcm = 1 Then countmistakes = True Else: countmistakes = False
If Form2.Checkdm = 1 Then displaymistakes = True Else: displaymistakes = False
If Form2.Checkds = 1 Then displaysets = True Else: displaysets = False


'   MAKE ALL CARDS BE IN DECK
For w = 1 To 3
For x = 1 To 3
For y = 1 To 3
For z = 1 To 3
    card(w, x, y, z) = 0
Next
Next
Next
Next

For x = 1 To 12
table(x) = ""
Picture1(x).Visible = False
Next

For p = 1 To 4
    gamer(p) = 0
    Player(p).Enabled = False
    Player(p).Caption = "Player " & (p)
Next

deck = 81
Lblcardsindeck = "Cards in deck: " & deck

For clearit = 1 To 12
    Picture1(clearit).BackColor = &H80000005
Next

cardsontable = 0
sets = 0

cc = 1

If gamestyle = 0 Then
    For gop = 1 To 4
        mistakes(gop) = 0
    Next
    
    mistake1 = 0
    mistake2 = 0
    mistake3 = 0
    mistake4 = 0
    
End If

Command3.Visible = True
Command3.Enabled = True
Command1.Visible = True
Command1.Enabled = False
Command2.Visible = True
Command2.Enabled = False
endgame.Visible = True
endgame.Enabled = True

End Sub

Private Sub phil()
For x = 1 To 12     'FOR TABLE

'MAKE CARD VISIBLE
If table(x) <> "" Then Picture1(x).Visible = True
If table(x) = "" Then Picture1(x).Visible = False

'DETERMINE SHAPE
If Mid(table(x), 4, 1) = 1 Then Shape1(x).Shape = 0 'rectanlge
If Mid(table(x), 4, 1) = 2 Then Shape1(x).Shape = 1 'square
If Mid(table(x), 4, 1) = 3 Then Shape1(x).Shape = 2 'oval

'DETERMINE COLOR
If Mid(table(x), 3, 1) = 1 Then Shape1(x).BorderColor = &HFF&       'RED
If Mid(table(x), 3, 1) = 2 Then Shape1(x).BorderColor = &HFF00&     'Green
If Mid(table(x), 3, 1) = 3 Then Shape1(x).BorderColor = &HC000C0    'Purple

'DETERMINE SHADING
If Mid(table(x), 2, 1) = 1 Then Shape1(x).FillStyle = 0  'solid
If Mid(table(x), 2, 1) = 2 Then Shape1(x).FillStyle = 1  'transparent
If Mid(table(x), 2, 1) = 3 Then Shape1(x).FillStyle = 4  'diagnol lines

'MAKE OTHER SHAPES ON CARD THE SAME
If Shape1(x).FillStyle = 0 Then Shape1(x).FillColor = Shape1(x).BorderColor
If Shape1(x).FillStyle = 4 Then Shape1(x).FillColor = vbBlack

Shape2(x).Shape = Shape1(x).Shape
Shape2(x).BorderColor = Shape1(x).BorderColor
Shape2(x).FillStyle = Shape1(x).FillStyle
If Shape2(x).FillStyle = 0 Then
    Shape2(x).FillColor = Shape2(x).BorderColor
    Else: Shape2(x).FillColor = vbBlack
End If

Shape3(x).Shape = Shape1(x).Shape
Shape3(x).BorderColor = Shape1(x).BorderColor
Shape3(x).FillStyle = Shape1(x).FillStyle
If Shape3(x).FillStyle = 0 Then
    Shape3(x).FillColor = Shape3(x).BorderColor
    Else: Shape3(x).FillColor = vbBlack
End If

'DETERMINE NUMBER
If Mid(table(x), 1, 1) = 1 Then
    Shape1(x).Visible = True
    Shape2(x).Visible = False
    Shape3(x).Visible = False
End If

If Mid(table(x), 1, 1) = 2 Then
    Shape1(x).Visible = True
    Shape2(x).Visible = True
    Shape3(x).Visible = False
End If

If Mid(table(x), 1, 1) = 3 Then
    Shape1(x).Visible = True
    Shape2(x).Visible = True
    Shape3(x).Visible = True
End If

Next             'FOR NEXT TABLE

End Sub

Private Sub hint_Click(Index As Integer)
Call showset
End Sub

Private Sub New_Click(Index As Integer)

Call Form_Load
Call phil

End Sub

Private Sub nextdeck_Click()

For w = 1 To 3
For x = 1 To 3
For y = 1 To 3
For z = 1 To 3
card(w, x, y, z) = 0
Next
Next
Next
Next

deck = 81

For goomba = 1 To 12
    If Picture1(goomba).Visible = True Then
        card((Mid(table(goomba), 4, 1)), (Mid(table(goomba), 3, 1)), (Mid(table(goomba), 2, 1)), (Mid(table(goomba), 1, 1))) = 1
        deck = deck - 1
    End If
Next

Lblcardsindeck.Caption = "Cards in deck: " & deck

End Sub

Private Sub options_Click(Index As Integer)
    Form2.Visible = True
    Form2.Command2.SetFocus
    
    '       FORM2 STUFF
For x = 1 To 3
    If gamestyle = x Then Form2.GSO(x) = True Else: Form2.GSO(x) = False
Next

If displaypoints = False Then Form2.Checkdp = 0
If displaypoints = True Then Form2.Checkdp = 1

If countmistakes = True Then Form2.Checkcm = 1
If countmistakes = False Then Form2.Checkcm = 0

If displaymistakes = True Then Form2.Checkdm = 1
If displaymistakes = False Then Form2.Checkdm = 0

If displaysets = True Then Form2.Checkds = 1
If displaysets = False Then Form2.Checkds = 0


    
End Sub

Private Sub Picture1_Click(Index As Integer)


If Picture1(Index).BackColor = &HFFFFFF Then picked = False
If Picture1(Index).BackColor = &HC0C0C0 Then picked = True

If picked = False Then
    If cc < 4 Then
    Picture1(Index).BackColor = &HC0C0C0
    cc = cc + 1
    End If
End If

If picked = True Then
    If cc > 1 Then
    Picture1(Index).BackColor = &HFFFFFF
     cc = cc - 1
    End If
End If

If cc = 4 Then
    For p = 1 To 4
        Player(p).Enabled = True
    Next
End If

End Sub

Private Sub Player_Click(Index As Integer)
check = 1
For find = 1 To 12

    If Picture1(find).BackColor = &HC0C0C0 Then
        cardcheck(check) = table(find)
        check = check + 1
    End If
Next

'                   CHECK 'EM
If (Mid(cardcheck(1), 4, 1) = Mid(cardcheck(2), 4, 1) And Mid(cardcheck(1), 4, 1) = Mid(cardcheck(3), 4, 1)) Or (Mid(cardcheck(1), 4, 1) <> Mid(cardcheck(2), 4, 1) And Mid(cardcheck(2), 4, 1) <> Mid(cardcheck(3), 4, 1)) Then
    If (Mid(cardcheck(1), 3, 1) = Mid(cardcheck(2), 3, 1) And Mid(cardcheck(1), 3, 1) = Mid(cardcheck(3), 3, 1)) Or (Mid(cardcheck(1), 3, 1) <> Mid(cardcheck(2), 3, 1) And Mid(cardcheck(2), 3, 1) <> Mid(cardcheck(3), 3, 1)) Then
        If (Mid(cardcheck(1), 2, 1) = Mid(cardcheck(2), 2, 1) And Mid(cardcheck(1), 2, 1) = Mid(cardcheck(3), 2, 1)) Or (Mid(cardcheck(1), 2, 1) <> Mid(cardcheck(2), 2, 1) And Mid(cardcheck(2), 2, 1) <> Mid(cardcheck(3), 2, 1)) Then
            If (Mid(cardcheck(1), 1, 1) = Mid(cardcheck(2), 1, 1) And Mid(cardcheck(1), 1, 1) = Mid(cardcheck(3), 1, 1)) Or (Mid(cardcheck(1), 1, 1) <> Mid(cardcheck(2), 1, 1) And Mid(cardcheck(2), 1, 1) <> Mid(cardcheck(3), 1, 1)) Then
                
                
                gamer(Index) = gamer(Index) + 1
                For dumdum = 1 To 12
                    If Picture1(dumdum).BackColor = &HC0C0C0 Then
                        table(dumdum) = ""
                        Picture1(dumdum).Visible = False
                        cardsontable = cardsontable - 1
                        For hmm = 1 To 4
                            Player(hmm).Enabled = False
                        Next
                        
                        If gamestyle = 3 Then
                            card(Mid(cardcheck(1), 4, 1), Mid(cardcheck(1), 3, 1), Mid(cardcheck(1), 2, 1), Mid(cardcheck(1), 1, 1)) = 0
                            card(Mid(cardcheck(2), 4, 1), Mid(cardcheck(2), 3, 1), Mid(cardcheck(2), 2, 1), Mid(cardcheck(2), 1, 1)) = 0
                            card(Mid(cardcheck(3), 4, 1), Mid(cardcheck(3), 3, 1), Mid(cardcheck(3), 2, 1), Mid(cardcheck(3), 1, 1)) = 0
                            deck = deck + 3
                        End If
                        
                        Command1.Enabled = True
                        Command2.Enabled = False
                        
                    End If
                Next
                
                For find = 1 To 12
                    If Picture1(find).BackColor = &HC0C0C0 Then Picture1(find).BackColor = &HFFFFFF
                Next
                
                cc = 1
                If gamestyle = 3 Then Call nextdeck_Click
                
            Else: mistakes(Index) = mistakes(Index) + 1
            End If
        Else: mistakes(Index) = mistakes(Index) + 1
        End If
    Else: mistakes(Index) = mistakes(Index) + 1
    End If
Else: mistakes(Index) = mistakes(Index) + 1
End If

gamer1 = gamer(1)
gamer2 = gamer(2)
gamer3 = gamer(3)
gamer4 = gamer(4)

mistake1 = mistakes(1)
mistake2 = mistakes(2)
mistake3 = mistakes(3)
mistake4 = mistakes(4)


If displaypoints = True Then
    Player(1).Caption = "Player " & "1" & vbNewLine & gamer1 & " point(s)"
    Player(2).Caption = "Player " & "2" & vbNewLine & gamer2 & " point(s)"
    Player(3).Caption = "Player " & "3" & vbNewLine & gamer3 & " point(s)"
    Player(4).Caption = "Player " & "4" & vbNewLine & gamer4 & " point(s)"
End If

If displaymistakes = True Then
    Player(1).Caption = Player(1).Caption & vbNewLine & mistakes(1) & " mistakes"
    Player(2).Caption = Player(2).Caption & vbNewLine & mistakes(2) & " mistakes"
    Player(3).Caption = Player(3).Caption & vbNewLine & mistakes(3) & " mistakes"
    Player(4).Caption = Player(4).Caption & vbNewLine & mistakes(4) & " mistakes"
End If

Call phil



If displaysets = True Then Call countsets
                
End Sub
Private Sub countsets()

Cls

sets = 0

For aa = 1 To 10
    If Picture1(aa).Visible = True Then
        For bb = (aa + 1) To 11
            If Picture1(bb).Visible = True Then
                 For ccc = (bb + 1) To 12
                    If Picture1(ccc).Visible = True Then

'                   CHECK 'EM
                          If (Mid(table(aa), 4, 1) = Mid(table(bb), 4, 1) And Mid(table(aa), 4, 1) = Mid(table(ccc), 4, 1)) Or (Mid(table(aa), 4, 1) <> Mid(table(bb), 4, 1) And Mid(table(bb), 4, 1) <> Mid(table(ccc), 4, 1) And (Mid(table(aa), 4, 1) <> Mid(table(ccc), 4, 1))) Then
                               If (Mid(table(aa), 3, 1) = Mid(table(bb), 3, 1) And Mid(table(aa), 3, 1) = Mid(table(ccc), 3, 1)) Or (Mid(table(aa), 3, 1) <> Mid(table(bb), 3, 1) And Mid(table(bb), 3, 1) <> Mid(table(ccc), 3, 1) And (Mid(table(aa), 3, 1) <> Mid(table(ccc), 3, 1))) Then
                                   If (Mid(table(aa), 2, 1) = Mid(table(bb), 2, 1) And Mid(table(aa), 2, 1) = Mid(table(ccc), 2, 1)) Or (Mid(table(aa), 2, 1) <> Mid(table(bb), 2, 1) And Mid(table(bb), 2, 1) <> Mid(table(ccc), 2, 1) And (Mid(table(aa), 2, 1) <> Mid(table(ccc), 2, 1))) Then
                                     If (Mid(table(aa), 1, 1) = Mid(table(bb), 1, 1) And Mid(table(aa), 1, 1) = Mid(table(ccc), 1, 1)) Or (Mid(table(aa), 1, 1) <> Mid(table(bb), 1, 1) And Mid(table(bb), 1, 1) <> Mid(table(ccc), 1, 1) And (Mid(table(aa), 1, 1) <> Mid(table(ccc), 1, 1))) Then
                
                
                                        sets = sets + 1
                
                
                                      End If
                                   End If
                                End If
                            End If
                    End If
                 Next
            End If
        Next
    End If
Next

Print sets

End Sub
Private Sub showset()

Cls

tonsofsets = 0

For aa = 1 To 10
    If Picture1(aa).Visible = True Then
        For bb = (aa + 1) To 11
            If Picture1(bb).Visible = True Then
                 For ccc = (bb + 1) To 12
                    If Picture1(ccc).Visible = True Then

                                        

'                   CHECK 'EM
                          If ((Mid(table(aa), 4, 1)) = (Mid(table(bb), 4, 1)) And (Mid(table(aa), 4, 1)) = (Mid(table(ccc), 4, 1))) Or ((Mid(table(aa), 4, 1)) <> (Mid(table(bb), 4, 1)) And (Mid(table(bb), 4, 1) <> Mid(table(ccc), 4, 1)) And (Mid(table(aa), 4, 1) <> Mid(table(ccc), 4, 1))) Then
                               If ((Mid(table(aa), 3, 1)) = (Mid(table(bb), 3, 1)) And (Mid(table(aa), 3, 1)) = (Mid(table(ccc), 3, 1))) Or ((Mid(table(aa), 3, 1)) <> (Mid(table(bb), 3, 1)) And (Mid(table(bb), 3, 1) <> Mid(table(ccc), 3, 1)) And (Mid(table(aa), 3, 1) <> Mid(table(ccc), 3, 1))) Then
                                   If ((Mid(table(aa), 2, 1)) = (Mid(table(bb), 2, 1)) And (Mid(table(aa), 2, 1)) = (Mid(table(ccc), 2, 1))) Or ((Mid(table(aa), 2, 1)) <> (Mid(table(bb), 2, 1)) And (Mid(table(bb), 2, 1) <> Mid(table(ccc), 2, 1)) And (Mid(table(aa), 2, 1) <> Mid(table(ccc), 2, 1))) Then
                                     If ((Mid(table(aa), 1, 1)) = (Mid(table(bb), 1, 1)) And (Mid(table(aa), 1, 1)) = (Mid(table(ccc), 1, 1))) Or ((Mid(table(aa), 1, 1)) <> (Mid(table(bb), 1, 1)) And (Mid(table(bb), 1, 1) <> Mid(table(ccc), 1, 1)) And (Mid(table(aa), 1, 1) <> Mid(table(ccc), 1, 1))) Then
                
                                        Print ""
                                        Print Picture1(aa).Index
                                        Print Picture1(bb).Index
                                        Print Picture1(ccc).Index
                                        Print ""
                                        
                                        'Print table(aa)
                                        'Print table(bb)
                                        'Print table(ccc)
                                        
                                        tonsofsets = tonsofsets + 1
                
                
                                      End If
                                   End If
                                End If
                            End If
                            
                    End If
                 Next
            End If
        Next
    End If
Next

If tonsofsets = 0 Then Print "No sets"

End Sub

Private Sub showamountofsets_Click(Index As Integer)
Call countsets
End Sub
