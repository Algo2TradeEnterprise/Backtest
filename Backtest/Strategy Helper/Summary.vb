Namespace StrategyHelper
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
                                                If x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
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
                                                If x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                    If x.ExitTime = Date.MinValue Then
                                                        Return Date.MaxValue
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

        Public ReadOnly Property OverallPL As Decimal
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.Sum(Function(x)
                                                If x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
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

        Public ReadOnly Property TradeCount As Integer
            Get
                If Me.AllTrades IsNot Nothing AndAlso Me.AllTrades.Count > 0 Then
                    Return Me.AllTrades.FindAll(Function(x)
                                                    Return x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                End Function).Count
                Else
                    Return Integer.MinValue
                End If
            End Get
        End Property
    End Class
End Namespace