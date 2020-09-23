VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRandomObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'by roger gilchrist
'this class allows you to generate 2 types of random object
'RangeMember(Min, Max)
' returns a random member in the range specified
'(NOTE: If you get min and Max the wrong way around the class takes care of it)
'RangeArray(Min, Max)
' returns a 0-based array of all the values between Min and Max
' Once you have the array you can use it to get a non-repeating set of values
' See the Demo Simple Poker for how to deal with using up existing array members
'
'There are also a number of named pre-defined versions of these
'Numeral returns 0 - 9
'Coin returns 0/1 OR T/H (if StringReturn =True)
'Die  returns 1 - 6
'Card returns 1 - 52 (53 if bJoker = True) OR simple string names (if StringReturn =True)
'Vowel returns a/e/i/o/u
'Consonent returns rest of letters
'RndWord returns a word with Min-Max range of characters
'                         Optionally in CVC (Consonent-Vowel-Consonent which is usually speakable)or totally random
'                         lower/UPPER/Proper case supported
'                      NOTE 'as is' non-standard English letters only are supported
'                            let me know if you extend it to deal with other languages
'
'NOTE
'CoinMember, DieMember, CardMember and RangeMember have an optional count parameter which
'allow you to correctly simulate multiple rolls/flips/draws or whatever range you are interested in
'See Proper Use form in demo with die and coins
Private bStringReturn     As Boolean
Private CardNames         As Variant
Private CoinNames         As Variant
Private vowels            As Variant
Private consonents        As Variant
Public Function CardArray(Optional ByVal bJoker As Boolean = False) As Variant
  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTmp2 As Variant
  arrTmp = RangeArray(1, 52 + IIf(bJoker, 1, 0))
  If bStringReturn Then
    ReDim arrTmp2(UBound(arrTmp)) As Variant
    For I = LBound(arrTmp) To UBound(arrTmp)
      arrTmp2(I) = CardNames(arrTmp(I) - 1)
    Next I
    arrTmp = arrTmp2
  End If
  CardArray = arrTmp
End Function
Public Function CardMember(Optional bJoker As Boolean = False, _
                           Optional Packs As Long = 1) As Variant
  CardMember = RangeMember(1, 52 + IIf(bJoker, 1, 0), Packs)
  If bStringReturn Then
    CardMember = CardNames(CardMember)
  End If
End Function
Private Sub Class_Initialize()
'ensure VB's best random is available
  Randomize Timer
'to create an specialist string arrays simply create the basic array here
  CardNames = Array("AH", "2H", "3H", "4H", "5H", "6H", "7H", "8H", "9H", "10H", "JH", "QH", "KH", "AC", "2C", "3C", "4C", _
                    "5C", "6C", "7C", "8C", "9C", "10C", "JC", "QC", "KC", "AS", "2S", "3S", "4S", "5S", "6S", "7S", "8S", _
                    "9S", "10S", "JS", "QS", "KS", "AD", "2D", "3D", "4D", "5D", "6D", "7D", "8D", "9D", "10D", "JD", "QD", _
                    "KD", "JOKER")
  CoinNames = Array("T", "H")
  vowels = Array("a", "e", "i", "o", "u")
  consonents = Array("b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "y", _
                     "z")
End Sub
Public Function CoinArray() As Variant
  Dim I      As Long
  Dim arrTmp As Variant
  arrTmp = RangeArray(0, 1)
  If bStringReturn Then
    ReDim arrTmp2(UBound(arrTmp)) As Variant
    For I = LBound(arrTmp) To UBound(arrTmp)
      arrTmp2(I) = CoinNames(arrTmp(I))
    Next I
    arrTmp = arrTmp2
  End If
  CoinArray = arrTmp
End Function
Public Function CoinMember(Optional Coins As Long = 1) As Variant
  CoinMember = RangeMember(0, 1, Coins)
  If bStringReturn Then
    CoinMember = CoinNames(CoinMember)
  End If
End Function
''
Public Function ConsonentArray(Optional ByVal bUcase As Boolean = False) As Variant
  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTmp2 As Variant
  arrTmp = RangeArray(LBound(consonents), UBound(consonents))
  ReDim arrTmp2(UBound(arrTmp)) As Variant
  For I = LBound(arrTmp) To UBound(arrTmp)
    arrTmp2(I) = consonents(arrTmp(I))
    If bUcase Then
      arrTmp2(I) = UCase$(arrTmp2(I))
    End If
  Next I
  ConsonentArray = arrTmp2
End Function
Public Function ConsonentMember(Optional ByVal bUcase As Boolean = False) As String
  ConsonentMember = consonents(RangeMember(LBound(consonents), UBound(consonents)))
  If bUcase Then
    ConsonentMember = UCase$(ConsonentMember)
  End If
End Function
Public Function DieArray() As Long()
  DieArray = RangeArray(1, 6)
End Function
Public Function DieMember(Optional dice As Long = 1) As Long
  DieMember = RangeMember(1, 6, dice)
End Function
Public Function LetterArray(Optional ByVal bUcase As Boolean = False) As Variant
  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTmp2 As Variant
  arrTmp = RangeArray(0, 25)
  ReDim arrTmp2(UBound(arrTmp)) As Variant
  For I = LBound(arrTmp) To UBound(arrTmp)
    arrTmp2(I) = Chr$(arrTmp(I) + IIf(bUcase, 65, 97))
  Next I
  LetterArray = arrTmp2
End Function
Public Function LetterMember(Optional bUcase As Boolean = False) As String
  LetterMember = Chr$(RangeMember(0, 25) + IIf(bUcase, 65, 97))
End Function
Public Function NumeralArray() As Long()
  NumeralArray = RangeArray(0, 9)
End Function
Public Function NumeralMember() As Long
  NumeralMember = RangeMember(0, 9)
End Function
Private Function RandMember(ByVal lngMin As Long, _
                           ByVal lngMax As Long) As Long
'returns a single member of the range
'use if you don't need the whole set
  If lngMax < lngMin Then
    SwapLong lngMin, lngMax
  End If
  RandMember = Int((lngMax - lngMin + 1) * Rnd + lngMin)
End Function
Public Function RangeArray(lngMin As Long, _
                           lngMax As Long) As Long()
  RangeArray = RandArray(lngMin, lngMax)
End Function
Public Function RangeMember(lngMin As Long, _
                            lngMax As Long, _
                            Optional ByVal Multi As Long = 1) As Long
  Dim I As Long
  For I = 1 To Multi
    RangeMember = RangeMember + RandMember(lngMin, lngMax)
  Next I
End Function
Private Function RandArray(ByVal lngMin As Long, _
                         ByVal lngMax As Long) As Variant
'Returns a shuffled array of the values between Min and Max
  Dim I    As Long
  Dim LMax As Long
  If lngMax < lngMin Then
    SwapLong lngMin, lngMax
  End If
  LMax = lngMax - lngMin
  ReDim Arraytmp(0 To LMax) As Long
'fill array with members
  For I = 0 To LMax
    Arraytmp(I) = lngMin + I
  Next I
'shuffle it
'For I = 0 To LMax ' for some reason this biases it to the low numbers
  For I = LMax To 0 Step -1
    SwapLong Arraytmp(I), Arraytmp(Int((LMax + 1) * Rnd))
  Next I
  RandArray = Arraytmp
End Function
Public Function RndWord(Optional MinLen As Long = 1, _
                        Optional MaxLen As Long = -1, _
                        Optional ByVal bRndSize As Boolean = False, _
                        Optional ByVal CVC As Boolean = True, _
                        Optional Kase As VbStrConv = vbLowerCase) As String
  Dim I  As Long
  Dim VC As Boolean
  If MinLen = -1 Then
    MinLen = RangeMember(1, 5)
  End If
  If MaxLen = -1 Then
    MaxLen = RangeMember(MinLen + 1, MinLen + 5)
  End If
  If bRndSize Then
    MaxLen = RangeMember(MinLen, MaxLen)
    MinLen = RangeMember(MinLen, MaxLen)
    If MaxLen < MinLen Then
      SwapLong MaxLen, MinLen
    End If
  End If
  VC = True
  For I = 1 To MaxLen
    If CVC Then
      VC = Not VC
     Else
      VC = Rnd > (26 / 5)
    End If
    If VC Then
      RndWord = RndWord & VowelMember
     Else
      RndWord = RndWord & ConsonentMember
    End If
    If I >= MinLen Then
      If Rnd > 0.5 Then
        Exit For
      End If
    End If
  Next I
  RndWord = StrConv(RndWord, Kase)
End Function
Public Function RndWordArray(ByVal lngSize As Long, _
                             Optional MinLen As Long = -1, _
                             Optional MaxLen As Long = -1, _
                             Optional bRndSize As Boolean = True, _
                             Optional CVC As Boolean = True, _
                             Optional Kase As VbStrConv = vbLowerCase) As Variant
  Dim I As Long
  ReDim arrTmp(lngSize - 1) As Variant
  For I = 0 To lngSize - 1
    arrTmp(I) = RndWord(MaxLen, MinLen, bRndSize, CVC, Kase)
  Next I
  RndWordArray = arrTmp
End Function
Public Property Get StringReturn() As Boolean
  StringReturn = bStringReturn
End Property
Public Property Let StringReturn(ByVal bStrLng As Boolean)
  bStringReturn = bStrLng
End Property
Public Sub SwapLong(element1 As Long, _
                    element2 As Long)
'Swap any two Long items
  Dim TEMP As Long
  TEMP = element1
  element1 = element2
  element2 = TEMP
End Sub
''
''
Public Function VowelArray(Optional ByVal bUcase As Boolean = False) As Variant
  Dim I       As Long
  Dim arrTmp  As Variant
  Dim arrTmp2 As Variant
  arrTmp = RangeArray(LBound(vowels), UBound(vowels))
  ReDim arrTmp2(UBound(arrTmp)) As Variant
  For I = LBound(arrTmp) To UBound(arrTmp)
    arrTmp2(I) = vowels(arrTmp(I))
    If bUcase Then
      arrTmp2(I) = UCase$(arrTmp2(I))
    End If
  Next I
  VowelArray = arrTmp2
End Function
''
Public Function VowelMember(Optional ByVal bUcase As Boolean = False) As String
  VowelMember = vowels(RangeMember(LBound(vowels), UBound(vowels)))
  If bUcase Then
    VowelMember = UCase$(VowelMember)
  End If
End Function
':)Code Fixer V2.5.3 (1/09/2004 3:11:10 PM) 33 + 304 = 337 Lines Thanks Ulli for inspiration and lots of code.


