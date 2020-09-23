VERSION 5.00
Begin VB.Form ProperUse 
   Caption         =   "Proper Use Of Coins/Dice"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   6030
   ScaleMode       =   0  'User
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame fraUsingCoins 
      Caption         =   "Using Coins"
      Height          =   1335
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
      Begin VB.PictureBox picCFXPBugFixForm2 
         BorderStyle     =   0  'None
         Height          =   998
         Index           =   1
         Left            =   100
         ScaleHeight     =   1005
         ScaleWidth      =   1860
         TabIndex        =   10
         Top             =   276
         Width           =   1860
         Begin VB.CommandButton cmdCoins 
            Caption         =   "2 Coins (WRONG)"
            Height          =   315
            Index           =   1
            Left            =   20
            TabIndex        =   13
            Top             =   297
            Width           =   1815
         End
         Begin VB.CommandButton cmdCoins 
            Caption         =   "2 Coins  (CORRECT)"
            Height          =   375
            Index           =   2
            Left            =   20
            TabIndex        =   12
            Top             =   612
            Width           =   1815
         End
         Begin VB.CommandButton cmdCoins 
            Caption         =   "1 Coin"
            Height          =   315
            Index           =   0
            Left            =   20
            TabIndex        =   11
            Top             =   -18
            Width           =   1815
         End
      End
   End
   Begin VB.Frame fraUsingDice 
      Caption         =   "Using Dice"
      Height          =   1335
      Left            =   7800
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
      Begin VB.PictureBox picCFXPBugFixForm2 
         BorderStyle     =   0  'None
         Height          =   1005
         Index           =   0
         Left            =   100
         ScaleHeight     =   1005
         ScaleWidth      =   1740
         TabIndex        =   6
         Top             =   276
         Width           =   1740
         Begin VB.CommandButton cmdDice 
            Caption         =   "2 dice =  (WRONG)"
            Height          =   315
            Index           =   1
            Left            =   20
            TabIndex        =   9
            Top             =   297
            Width           =   1695
         End
         Begin VB.CommandButton cmdDice 
            Caption         =   "2 dice= (CORRECT)"
            Height          =   375
            Index           =   2
            Left            =   20
            TabIndex        =   8
            Top             =   612
            Width           =   1695
         End
         Begin VB.CommandButton cmdDice 
            Caption         =   "1 die "
            Height          =   315
            Index           =   0
            Left            =   20
            TabIndex        =   7
            Top             =   -18
            Width           =   1695
         End
      End
   End
   Begin VB.ListBox lstCounts 
      Height          =   2595
      Left            =   7680
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.PictureBox picGraph 
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Line linChart 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   20
         Index           =   0
         Visible         =   0   'False
         X1              =   120
         X2              =   135
         Y1              =   5280
         Y2              =   5295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2295
      Left            =   5880
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblCodeUsed 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label lblGraph 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   5655
   End
End
Attribute VB_Name = "ProperUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RO    As New clsRandomObject
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdCoins_Click(Index As Integer)

  Dim val As Long
  Dim I   As Long
  ZeroGraph
'lblGraph = " TT TH/HT HH"
  lblGraph = "   0      1      2  HEADS"
  lblGraph.Refresh
  Select Case Index
   Case 0
    lblCodeUsed = "Code: RO.CoinMember"
    ListSetup 1
   Case 1
    lblCodeUsed = "Code: RO.RangeMember(0,2)"
    ListSetup 2
   Case 2
    lblCodeUsed = "Code: RO.CoinMember(2)"
    ListSetup 2
  End Select
  lblCodeUsed.Refresh
  For I = 1 To 10000
    Select Case Index
     Case 0
      val = RO.CoinMember
     Case 1
      val = RO.RangeMember(0, 2)
     Case 2
      val = RO.CoinMember(2)
'val = RO.RangeMember(0, 1, 2)
'val = RO.CoinMember + RO.CoinMember
    End Select
    doDrawing val, I
  Next I
End Sub
Private Sub cmdDice_Click(Index As Integer)
  Dim I   As Long
  Dim val As Long
  ZeroGraph
  lblGraph = "   0      1      2       3      4      5      6      7      8      9     10    11     12"
  lblGraph.Refresh
  ListSetup 6 * IIf(Index = 0, 1, 2)
  Select Case Index
   Case 0
    lblCodeUsed = "Code: RO.DieMember"
   Case 1
    lblCodeUsed = "Code: RO.RangeMember(2, 12)"
   Case 2
    lblCodeUsed = "Code: RO.DieMember(2)"
  End Select
  lblCodeUsed.Refresh
  For I = 1 To 30000
    Select Case Index
     Case 0
      val = RO.DieMember
     Case 1
      val = RO.RangeMember(2, 12)
     Case 2
      val = RO.DieMember(2)
    End Select
    doDrawing val, I
  Next I
End Sub
Private Sub doDrawing(ByVal lngVal As Long, _
                      ByVal trigger As Long)
  lstCounts.List(lngVal) = lstCounts.List(lngVal) + 1
  linChart(lngVal).Y2 = linChart(lngVal).Y2 - 1
  If trigger Mod 100 = 0 Then
    linChart(lngVal).Refresh
    lstCounts.Refresh
  End If
End Sub
Private Sub Form_Load()
  Dim I As Long
  On Error Resume Next
  If App.PrevInstance Then
    MsgBox "Program already running !"
    End
  End If
'  btnFlat Command1
  With linChart(0)
    .BorderColor = QBColor(0)
    .X1 = 180
    .X2 = 180
    .Y1 = 5280
    .Y2 = 5280
    .Visible = True
  End With 'Line1(0)
  For I = 1 To 12
    Load linChart(I)
    With linChart(I)
      .BorderColor = QBColor(I)
      .X1 = linChart(I - 1).X1 + 360
      .X2 = .X1
      .Visible = True
    End With 'Line1(I)
  Next I
  Label1 = "You need to be careful when combining random objects." & vbNewLine & _
           "It is not safe to simply call a random value that is in the range or possible answers." & vbNewLine & _
           "You have to call multiples of the actual event."
  On Error GoTo 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub ListSetup(ByVal lsize As Long)
  Dim I As Long
  lstCounts.Clear
  For I = 0 To lsize
    lstCounts.AddItem "0"
  Next I
End Sub
Private Sub ZeroGraph()
  Dim I As Long
  For I = 0 To 12
    With linChart(I)
      .Y2 = 5280
      .Refresh
      .Visible = True
    End With 'Line1(I)
  Next I
End Sub
':)Code Fixer V2.5.3 (1/09/2004 3:11:11 PM) 2 + 163 = 165 Lines Thanks Ulli for inspiration and lots of code.


