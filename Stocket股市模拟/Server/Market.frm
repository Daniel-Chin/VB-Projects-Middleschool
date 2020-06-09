VERSION 5.00
Begin VB.Form Market 
   Caption         =   "Market"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Debuffer 
      Interval        =   19
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer BankShrinker 
      Interval        =   1000
      Left            =   1560
      Top             =   720
   End
End
Attribute VB_Name = "Market"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typBuff
    S As Integer
    MS As Integer
    ID As Integer
    BuyIndex As Integer
End Type

Private Type typOwner
    ID As Integer
    Share As Single
End Type

Private Type typStock
    OwnerNum As Integer
    Price As Double
    Owner(1 To 99) As typOwner
    Share As Single
End Type

Private Type typPlayer
    Dead As Boolean
    HaveWhich As Integer
End Type

Const StartBank As Single = 1000

Dim Bank() As Single
Dim PlayerNum As Integer
Dim StockNum As Integer
Dim Stock() As typStock
Dim Buff(1 To 999) As typBuff
Dim BuffNum As Integer
Dim Player() As typPlayer

Public Sub TheMarketIsOpen(passStockNum As Integer, passPlayerNum As Integer)
PlayerNum = passPlayerNum
StockNum = passStockNum
ReDim Bank(1 To PlayerNum)
Dim k As Integer
For k = 1 To PlayerNum
    Bank(k) = StartBank
Next k
ReDim Stock(1 To StockNum)
ReDim Player(1 To PlayerNum)
For k = 1 To StockNum
    Stock(k).Price = 1
Next k
End Sub

Private Sub BankShrinker_Timer()
Me.Visible = False
Dim SomeOneHold As Boolean
Dim k As Integer
For k = 1 To PlayerNum
    If Player(k).HaveWhich = 0 Then
        Bank(k) = Bank(k) - 1
        If Bank(k) <= 0 Then
            Bank(k) = 0
            Die k
        End If
        Web.Send k, "balance", Bank(k)
        SomeOneHold = True
    End If
Next k
If Not SomeOneHold Then
    BankShrinker = False
End If
End Sub

Private Sub Die(ID As Integer)
Player(ID).Dead = True
Dim k As Integer
Dim Death As Integer
For k = 1 To PlayerNum
    If Player(k).Dead Then
        Death = Death + 1
    End If
Next k
If PlayerNum - Death <= 2 Then
    GameOver
End If
Web.Send ID, "msgbox", "你死了。"
End Sub

Public Sub AddBuff(S As Integer, MS As Integer, ID As Integer, BuyIndex As Integer)
If Not Player(ID).Dead Then
    BuffNum = BuffNum + 1
    With Buff(BuffNum)
        .BuyIndex = BuyIndex
        .ID = ID
        .MS = MS
        .S = S
    End With
End If
End Sub

Private Sub Debuffer_Timer()
Dim k As Integer
Dim Delta As Long
Dim AllClear As Boolean
Do Until AllClear
    AllClear = True
    For k = 1 To BuffNum
        With Buff(k)
            Delta = (Second(Now) - .S) * 1000 + Web.NowMs - .MS
            Delta = (Delta + 60000) Mod 60000
            If Delta > 1000 Then
                Debug.Print "有问题！Buff from " & .ID
            End If
            If Delta >= 500 Then
                Buy .ID, .BuyIndex
                .BuyIndex = Buff(BuffNum).BuyIndex
                .ID = Buff(BuffNum).ID
                .MS = Buff(BuffNum).MS
                .S = Buff(BuffNum).S
                BuffNum = BuffNum - 1
                AllClear = False
                Exit For
            End If
        End With
    Next k
Loop
End Sub

Private Sub Buy(ID As Integer, BuyIndex As Integer)
Dim TempFund As Single
Dim PreFund As Single, NowFund As Single
Dim OwnerID As Integer
Dim k As Integer
If Player(ID).HaveWhich = 0 Then
    'First Buy
    TempFund = Bank(ID)
    Bank(ID) = 0
    Web.Send ID, "balance", 0
Else
    With Stock(Player(ID).HaveWhich)
        For OwnerID = 1 To .OwnerNum
            If .Owner(OwnerID).ID = ID Then
                Exit For
            End If
        Next OwnerID
        Debug.Assert .Owner(OwnerID).ID = ID
        PreFund = .Share * .Price
        .Share = .Share - .Owner(OwnerID).Share
        .Price = Price(.Share)
        NowFund = .Share * .Price
        TempFund = PreFund - NowFund
        If OwnerID <> .OwnerNum Then
            For k = OwnerID To .OwnerNum - 1
                .Owner(k) = .Owner(k + 1)
            Next k
        End If
        .Owner(.OwnerNum).ID = 0
        .Owner(.OwnerNum).Share = 0
        .OwnerNum = .OwnerNum - 1
    End With
End If
'以上是卖
If TempFund < 10 Then ''''''''
    Die ID
End If
'以下是买
Dim Assume As Single, sTep As Single
Player(ID).HaveWhich = BuyIndex
With Stock(BuyIndex)
    Assume = .Share
    sTep = TempFund / 3 / .Price ''''''
    PreFund = .Price * .Share
    NowFund = Assume * Price(Assume)
    Do Until NowFund - PreFund > TempFund
        Assume = Assume + sTep
        NowFund = Assume * Price(Assume)
    Loop
    Do Until Abs(NowFund - PreFund - TempFund) <= 0.001 '''''
        Assume = Assume - sTep
        NowFund = Assume * Price(Assume)
        If NowFund - PreFund > TempFund Then
            sTep = Abs(sTep * 0.501)
        Else
            sTep = -Abs(sTep * 0.501)
        End If
        DoEvents
    Loop
    .OwnerNum = .OwnerNum + 1
    .Owner(.OwnerNum).ID = ID
    .Owner(.OwnerNum).Share = Assume - .Share
    .Share = Assume
    .Price = Price(Assume)
End With
TransmitMarket
End Sub

Private Function Price(Share As Single) As Double
''''''''''''
Price = 1 + Share * 0.0019
End Function

Private Sub TransmitMarket()
Dim strData As String
Dim k As Integer, kk As Integer
Dim SuoFang As Single
SuoFang = Sqr(2 / PlayerNum)
For k = 1 To StockNum
    With Stock(k)
        strData = strData & Int(.Price * 500 * SuoFang) & " " & .OwnerNum & " " ''''''
        For kk = 1 To .OwnerNum
            strData = strData & .Owner(kk).ID & " " & Int(.Owner(kk).Share * 5 * SuoFang) & " "
        Next kk
    End With
Next k
For k = 1 To PlayerNum
    Web.Send k, "market", strData
Next k
End Sub

Public Sub GameOver()
Dim k As Integer
For k = 1 To PlayerNum
    Web.Send k, "gameover", ""
Next k
MsgBox "Sever:Over."
End Sub

