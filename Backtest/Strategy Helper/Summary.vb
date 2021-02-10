﻿Namespace StrategyHelper
    Public Class Summary
        Public AllTrades As List(Of Trade)

        Public ReadOnly Property Instrument As String
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.FirstOrDefault.TradingSymbol
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property StartDate As Date
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.Min(Function(x)
                                                If x.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                                                    Return x.EntryTime.Date
                                                Else
                                                    Return Date.MaxValue
                                                End If
                                            End Function)
                Else
                    Return Date.MaxValue
                End If
            End Get
        End Property

        Public ReadOnly Property EndDate As Date
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.Max(Function(x)
                                                If x.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                                                    If x.ExitTime = Date.MinValue Then
                                                        Return Now.Date
                                                    Else
                                                        Return x.ExitTime.Date
                                                    End If
                                                Else
                                                    Return Date.MinValue
                                                End If
                                            End Function)
                Else
                    Return Date.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property NumberOfDays As Integer
            Get
                Return Me.EndDate.Subtract(Me.StartDate).Days + 1
            End Get
        End Property

        Public ReadOnly Property TradeCount As Integer
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.FindAll(Function(x)
                                                    Return x.TradeCurrentStatus <> Trade.TradeStatus.Cancel
                                                End Function).Count
                Else
                    Return Integer.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property ContractRolloverTradeCount As Integer
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.FindAll(Function(x)
                                                    Return x.ExitType = Trade.TypeOfExit.ContractRollover
                                                End Function).Count
                Else
                    Return Integer.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property ReverseTradeCount As Integer
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.FindAll(Function(x)
                                                    Return x.ExitType = Trade.TypeOfExit.Reversal
                                                End Function).Count
                Else
                    Return Integer.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property OverallPL As Decimal
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.Sum(Function(x)
                                                If x.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                                                    Return x.PLAfterBrokerage
                                                Else
                                                    Return 0
                                                End If
                                            End Function)
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property MaxCapital As Decimal
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Dim maxCapitalUsed As Decimal = 0
                    For Each runningTrade In Me.AllTrades.OrderBy(Function(x)
                                                                      Return x.EntryTime
                                                                  End Function)
                        maxCapitalUsed += runningTrade.CapitalRequiredWithMargin
                        If runningTrade.TradeCurrentStatus <> Trade.TradeStatus.Inprogress Then
                            If runningTrade.ExitType <> Trade.TypeOfExit.Target Then
                                maxCapitalUsed -= runningTrade.CapitalRequiredWithMargin + runningTrade.PLAfterBrokerage
                            End If
                        End If
                    Next
                    Return maxCapitalUsed
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property AbsoluteReturnOfInvestment As Decimal
            Get
                If Me.MaxCapital <> 0 Then
                    Return (Me.OverallPL / Me.MaxCapital) * 100
                Else
                    Return 0
                End If
            End Get
        End Property

        Public ReadOnly Property AnnualReturnOfInvestment As Decimal
            Get
                Return Me.AbsoluteReturnOfInvestment / 365
            End Get
        End Property
    End Class
End Namespace