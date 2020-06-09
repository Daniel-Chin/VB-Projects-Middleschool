Attribute VB_Name = "Market"
Option Explicit

Private Type typOwner
    ID As Integer
    Share As Single
End Type

Private Type typStock
    OwnerNum As Integer
    Price As Double
    Owner() As typOwner
End Type

Const StartBank As Single = 1000

Dim Bank() As Single
Dim PlayerNum As Integer
Dim StockNum As Integer
Dim Stock() As typStock

Public Sub TheMarketIsOpen(passStockNum As Integer, passPlayerNum As Integer)
PlayerNum = passPlayerNum
StockNum = passStockNum
ReDim Bank(1 To PlayerNum)
Dim k As Integer
For k = 1 To PlayerNum
    Bank(k) = StartBank
Next k
ReDim Stock(1 To StockNum)
End Sub
