VERSION 5.00
Begin VB.Form RndObjectDemo 
   Caption         =   "clsRandomObject Demo"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleMode       =   0  'User
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowProperUse 
      Caption         =   "Show Proper Use of Coins/Dice Form"
      Height          =   975
      Left            =   4440
      TabIndex        =   56
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame fraRange 
      Caption         =   "Range (-200 to 200)"
      Height          =   2535
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "Class copes if Min < Max"
      Top             =   2520
      Width           =   1815
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   2205
         Index           =   2
         Left            =   100
         ScaleHeight     =   2205
         ScaleWidth      =   1500
         TabIndex        =   26
         Top             =   270
         Width           =   1500
         Begin VB.CommandButton cmdRange 
            Caption         =   "Array"
            Height          =   375
            Index           =   1
            Left            =   20
            TabIndex        =   30
            Top             =   1782
            Width           =   1455
         End
         Begin VB.CommandButton cmdRange 
            Caption         =   "Member"
            Height          =   375
            Index           =   0
            Left            =   20
            TabIndex        =   29
            Top             =   1422
            Width           =   1455
         End
         Begin VB.HScrollBar hscRange 
            Height          =   255
            Index           =   1
            Left            =   20
            Max             =   200
            Min             =   -200
            TabIndex        =   28
            Top             =   822
            Value           =   200
            Width           =   1455
         End
         Begin VB.HScrollBar hscRange 
            Height          =   255
            Index           =   0
            Left            =   20
            Max             =   200
            Min             =   -200
            TabIndex        =   27
            Top             =   222
            Width           =   1455
         End
         Begin VB.Label lblRange 
            Caption         =   "Max"
            Height          =   255
            Index           =   1
            Left            =   20
            TabIndex        =   34
            Top             =   582
            Width           =   375
         End
         Begin VB.Label lblRange 
            Caption         =   "Min"
            Height          =   255
            Index           =   0
            Left            =   20
            TabIndex        =   33
            Top             =   -18
            Width           =   375
         End
         Begin VB.Label lblRangeNums 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "200"
            Height          =   255
            Index           =   1
            Left            =   500
            TabIndex        =   32
            Top             =   582
            Width           =   495
         End
         Begin VB.Label lblRangeNums 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   500
            TabIndex        =   31
            Top             =   -18
            Width           =   495
         End
      End
   End
   Begin VB.Frame fraSimplePoker 
      Caption         =   "Simple Poker (No AI just shows how to read cards off the array)"
      Height          =   1815
      Left            =   5760
      TabIndex        =   24
      Top             =   3240
      Width           =   4695
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   1545
         Index           =   3
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   4500
         TabIndex        =   35
         Top             =   240
         Width           =   4500
         Begin VB.CommandButton cmdShuffle 
            Caption         =   "Shuffle"
            Height          =   375
            Left            =   3000
            TabIndex        =   55
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeal5Cards 
            Caption         =   "Deal 5 cards"
            Height          =   375
            Left            =   140
            TabIndex        =   36
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblHand 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   1
            Left            =   135
            TabIndex        =   38
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblHand 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   0
            Left            =   140
            TabIndex        =   37
            Top             =   102
            Width           =   2775
         End
      End
   End
   Begin VB.CheckBox chkLetterUcase 
      Caption         =   "UCase"
      Height          =   195
      Left            =   2400
      TabIndex        =   23
      ToolTipText     =   "applies to Letter/Vowel/Consonent"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame fraWordSettings 
      Caption         =   "Word Settings"
      Height          =   3855
      Left            =   2160
      TabIndex        =   21
      Top             =   1200
      Width           =   1935
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   3585
         Index           =   4
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   1740
         TabIndex        =   39
         Top             =   240
         Width           =   1735
         Begin VB.CommandButton cmdArray 
            Caption         =   "Word Array"
            Height          =   375
            Index           =   7
            Left            =   20
            TabIndex        =   47
            Top             =   3105
            Width           =   1695
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "WORD"
            Height          =   375
            Index           =   7
            Left            =   20
            TabIndex        =   46
            Top             =   2625
            Width           =   1695
         End
         Begin VB.HScrollBar hscCharacters 
            Height          =   255
            Index           =   2
            Left            =   980
            Max             =   100
            Min             =   2
            TabIndex        =   45
            Top             =   2265
            Value           =   10
            Width           =   495
         End
         Begin VB.CheckBox chkRandomSize 
            Caption         =   "Random size"
            Height          =   255
            Left            =   140
            TabIndex        =   44
            Top             =   702
            Width           =   1455
         End
         Begin VB.CheckBox chkCVC 
            Caption         =   "CVC (Off - random)"
            Height          =   255
            Left            =   20
            TabIndex        =   43
            Top             =   1665
            Width           =   1695
         End
         Begin VB.ListBox lstCase 
            Height          =   645
            Left            =   15
            TabIndex        =   42
            Top             =   945
            Width           =   1575
         End
         Begin VB.HScrollBar hscCharacters 
            Height          =   255
            Index           =   1
            Left            =   980
            Max             =   20
            Min             =   2
            TabIndex        =   41
            Top             =   462
            Value           =   10
            Width           =   495
         End
         Begin VB.HScrollBar hscCharacters 
            Height          =   255
            Index           =   0
            Left            =   980
            Max             =   19
            Min             =   1
            TabIndex        =   40
            Top             =   222
            Value           =   5
            Width           =   495
         End
         Begin VB.Label lblNumberOfCharacters 
            Caption         =   "Number of characters"
            Height          =   255
            Left            =   20
            TabIndex        =   54
            Top             =   -18
            Width           =   1575
         End
         Begin VB.Label lblChar 
            Caption         =   "Array Members"
            Height          =   255
            Index           =   2
            Left            =   15
            TabIndex        =   53
            Top             =   2025
            Width           =   1095
         End
         Begin VB.Label lblCharNums 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   1215
            TabIndex        =   52
            Top             =   2025
            Width           =   255
         End
         Begin VB.Label lblChar 
            Caption         =   "Max"
            Height          =   255
            Index           =   1
            Left            =   140
            TabIndex        =   51
            Top             =   462
            Width           =   375
         End
         Begin VB.Label lblChar 
            Caption         =   "Min"
            Height          =   255
            Index           =   0
            Left            =   140
            TabIndex        =   50
            Top             =   222
            Width           =   375
         End
         Begin VB.Label lblCharNums 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   615
            TabIndex        =   49
            Top             =   465
            Width           =   375
         End
         Begin VB.Label lblCharNums 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   620
            TabIndex        =   48
            Top             =   222
            Width           =   375
         End
      End
   End
   Begin VB.CheckBox chkJoker 
      Caption         =   "Joker in the Pack"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "include a joker in cards"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox lstDisplayArray 
      Columns         =   5
      Height          =   2595
      ItemData        =   "RndObjectDemo.frx":0000
      Left            =   5760
      List            =   "RndObjectDemo.frx":0007
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame fraFrame2 
      Caption         =   "Random Arrays"
      Height          =   3015
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   2685
         Index           =   0
         Left            =   100
         ScaleHeight     =   2685
         ScaleWidth      =   1260
         TabIndex        =   5
         Top             =   276
         Width           =   1260
         Begin VB.CommandButton cmdArray 
            Caption         =   "Consonent"
            Height          =   375
            Index           =   6
            Left            =   15
            TabIndex        =   20
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "Vowel"
            Height          =   375
            Index           =   5
            Left            =   15
            TabIndex        =   19
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "Card"
            Height          =   375
            Index           =   4
            Left            =   15
            TabIndex        =   10
            Top             =   1422
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "Die"
            Height          =   375
            Index           =   3
            Left            =   15
            TabIndex        =   9
            Top             =   1062
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "coin"
            Height          =   375
            Index           =   2
            Left            =   15
            TabIndex        =   8
            Top             =   702
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "Letter"
            Height          =   375
            Index           =   1
            Left            =   15
            TabIndex        =   7
            Top             =   342
            Width           =   1215
         End
         Begin VB.CommandButton cmdArray 
            Caption         =   "Numeral"
            Height          =   375
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Top             =   -18
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraSingleMember 
      Caption         =   "Single Member"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   1838
         Index           =   1
         Left            =   100
         ScaleHeight     =   1845
         ScaleWidth      =   1620
         TabIndex        =   11
         Top             =   276
         Width           =   1620
         Begin VB.CommandButton cmdMember 
            Caption         =   "Consonent"
            Height          =   375
            Index           =   6
            Left            =   20
            TabIndex        =   18
            Top             =   2385
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "vowel"
            Height          =   375
            Index           =   5
            Left            =   20
            TabIndex        =   17
            Top             =   2025
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "Card"
            Height          =   375
            Index           =   4
            Left            =   20
            TabIndex        =   16
            Top             =   1485
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "Die"
            Height          =   375
            Index           =   3
            Left            =   20
            TabIndex        =   15
            Top             =   1110
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "Numeral"
            Height          =   375
            Index           =   0
            Left            =   20
            TabIndex        =   14
            Top             =   -18
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "Letter"
            Height          =   375
            Index           =   1
            Left            =   20
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdMember 
            Caption         =   "Coin"
            Height          =   375
            Index           =   2
            Left            =   20
            TabIndex        =   12
            Top             =   735
            Width           =   1455
         End
      End
   End
   Begin VB.CheckBox chkReturnAString 
      Caption         =   "Return a String"
      Height          =   195
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "applies to coin/card"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblListSize 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Members: "
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   2760
      Width           =   4695
   End
End
Attribute VB_Name = "RndObjectDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RO        As New clsRandomObject
Private Deck      As Variant
Private dealt     As Long
Private Sub chkReturnAString_Click()
  RO.StringReturn = chkReturnAString.Value = vbChecked
End Sub
Private Sub cmdArray_Click(Index As Integer)
  Dim arr As Variant
  Select Case Index
   Case 0
    arr = RO.NumeralArray
   Case 1
    arr = RO.LetterArray(chkLetterUcase.Value = vbChecked)
   Case 2
    arr = RO.CoinArray
   Case 3
    arr = RO.DieArray
   Case 4
    Deck = RO.CardArray(chkJoker.Value = vbChecked)
    arr = Deck
   Case 5
    arr = RO.VowelArray(chkLetterUcase.Value = vbChecked)
   Case 6
    arr = RO.ConsonentArray(chkLetterUcase.Value = vbChecked)
   Case 7
    arr = RO.RndWordArray(hscCharacters(2).Value, hscCharacters(0).Value, hscCharacters(1).Value, , chkCVC.Value = vbChecked, lstCase.ListIndex + 1)
  End Select
  FillList arr
End Sub
Private Sub cmdDeal5Cards_Click()
  Dim I        As Long
  lblHand(0) = ""
  lblHand(1) = ""
  If IsEmpty(Deck) Then
'create a pack if necessary
    shuffleTheDeck
  End If
  If dealt + 10 >= UBound(Deck) Then
' if you have reached the end of the pack
' or there are not enough members left for a hand
' restart, don't forget resetting the tracking variable to 0
    shuffleTheDeck
    dealt = 0
  End If
  lblHand(0) = ""
  lblHand(1) = ""
  For I = 0 To 9 Step 2
    lblHand(0) = lblHand(0) & Deck(dealt + I) & "  "
    lblHand(1) = lblHand(1) & Deck(dealt + I + 1) & "  "
  Next I
  dealt = dealt + 10
End Sub
Private Sub cmdMember_Click(Index As Integer)
  Dim strCom As String
  Select Case Index
   Case 0
    strCom = "Numeral " & RO.NumeralMember
   Case 1
    strCom = "letter " & RO.LetterMember(chkLetterUcase.Value = vbChecked)
   Case 2
    strCom = "Coin " & RO.CoinMember
   Case 3
    strCom = "Die " & RO.DieMember
   Case 4
    strCom = "Card " & RO.CardMember(chkJoker.Value = vbChecked)
   Case 5
    strCom = "vowel " & RO.VowelMember(chkLetterUcase.Value = vbChecked)
   Case 6
    strCom = "Consonent " & RO.ConsonentMember(chkLetterUcase.Value = vbChecked)
   Case 7
    strCom = "WORD: " & RO.RndWord(hscCharacters(0).Value, hscCharacters(1).Value, chkRandomSize.Value = vbChecked, chkCVC.Value = vbChecked, lstCase.ListIndex + 1)
  End Select
  cmdMember(Index).Caption = strCom
End Sub
Private Sub cmdRange_Click(Index As Integer)
  Dim arr As Variant
  Select Case Index
   Case 0
    cmdRange(0).Caption = "Member: " & RO.RangeMember(hscRange(0).Value, hscRange(1).Value)
   Case 1
    arr = RO.RangeArray(hscRange(0).Value, hscRange(1).Value)
    FillList arr
  End Select
End Sub
Private Sub cmdShowProperUse_Click()
  ProperUse.Show
End Sub
Private Sub cmdShuffle_Click()
  shuffleTheDeck
End Sub
Private Sub FillList(arr As Variant)
  Dim I As Long
  lstDisplayArray.Clear
  For I = LBound(arr) To UBound(arr)
    lstDisplayArray.AddItem arr(I)
  Next I
  lblListSize = "Members: " & lstDisplayArray.ListCount
End Sub
Private Sub Form_Load()
  hscCharacters_Change 0
  hscCharacters_Change 1
  hscCharacters_Change 2
  hscRange_Change 0
  hscRange_Change 1
  With lstCase
    .AddItem "Ucase"
    .AddItem "Lcase"
    .AddItem "Pcase"
    .ListIndex = 1
  End With 'List1
End Sub
Private Sub hscCharacters_Change(Index As Integer)
  lblCharNums(Index) = hscCharacters(Index).Value
End Sub
Private Sub hscRange_Change(Index As Integer)
  lblRangeNums(Index) = hscRange(Index).Value
End Sub
Private Sub shuffleTheDeck()
  Deck = RO.CardArray(chkJoker.Value = vbChecked)
  FillList Deck
End Sub
':)Code Fixer V2.5.3 (1/09/2004 3:11:11 PM) 4 + 162 = 166 Lines Thanks Ulli for inspiration and lots of code.
