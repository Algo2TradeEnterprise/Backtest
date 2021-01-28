Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class BelowPortfolioValueStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
        Public MaxIteration As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentDayCandlePayload IsNot Nothing AndAlso currentDayCandlePayload.PreviousCandlePayload IsNot Nothing Then
            Dim signals As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)) = GetSignalForEntry(currentDayCandlePayload)
            If signals IsNot Nothing AndAlso signals.Count > 0 Then
                For Each runningSignal In signals
                    Dim entryPrice As Decimal = runningSignal.Item2
                    Dim quantity As Integer = runningSignal.Item3
                    Dim iteration As Integer = runningSignal.Item4
                    Dim remarks As String = runningSignal.Item5

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = entryPrice,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = iteration,
                                                                .Supporting2 = remarks,
                                                                .Supporting3 = runningSignal.Item6
                                                            }

                    If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                    parameters.Add(parameter)
                Next
            End If
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function GetSignalForEntry(ByVal candle As Payload) As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal))
        Dim ret As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, candle.Open)
        Dim entryPrice As Decimal = candle.Close
        If lastTrade Is Nothing Then
            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal))
            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)(True, entryPrice, quantity, iteration, "First Trade", quantity * entryPrice))
        Else
            Dim allTrades As List(Of Trade) = Me._parentStrategy.GetSpecificTrades(candle, Me._parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
            Dim totalQuantity As Double = allTrades.Sum(Function(x)
                                                            Return x.Quantity
                                                        End Function)

            Dim highestValue As Decimal = candle.High * totalQuantity
            Dim currentValue As Decimal = candle.Close * totalQuantity
            If highestValue > Val(lastTrade.Supporting3) Then
                lastTrade.UpdateTrade(Supporting3:=highestValue)
            End If

            If currentValue <= highestValue - highestValue * 5 / 100 Then
                If entryPrice < lastTrade.EntryPrice Then
                    If lastTrade.EntryPrice - entryPrice >= atr Then
                        If Val(lastTrade.Supporting1) < _userInputs.MaxIteration Then
                            iteration = Val(lastTrade.Supporting1) + 1
                            quantity = GetQuantity(iteration, entryPrice)

                            totalQuantity += quantity
                            currentValue = candle.Close * totalQuantity

                            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal))
                            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)(True, entryPrice, quantity, iteration, "Below last entry", currentValue))
                        Else
                            totalQuantity += quantity
                            currentValue = candle.Close * totalQuantity

                            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal))
                            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)(True, entryPrice, quantity, iteration, "(Reset) Max Iteration", currentValue))
                        End If
                        'Else
                        '    Console.WriteLine(String.Format("Trade Neglected for ATR on {0}", candle.PayloadDate.ToString("dd-MMM-yyyy")))
                    End If
                Else
                    totalQuantity += quantity
                    currentValue = candle.Close * totalQuantity

                    If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal))
                    ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal)(True, entryPrice, quantity, iteration, "(Reset) Above last entry", currentValue))
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital * Math.Pow(2, iterationNumber - 1)
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function

    Private Function GetAveragePrice(ByVal currentPayload As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If currentPayload IsNot Nothing Then
            Dim allTrades As List(Of Trade) = Me._parentStrategy.GetSpecificTrades(currentPayload, Me._parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                Dim totalTurnover As Double = allTrades.Sum(Function(x)
                                                                Return x.EntryPrice * x.Quantity
                                                            End Function)
                Dim totalQuantity As Double = allTrades.Sum(Function(x)
                                                                Return x.Quantity
                                                            End Function)

                ret = totalTurnover / totalQuantity
            End If
        End If
        Return ret
    End Function

    Private Function GetTotalQu(ByVal currentPayload As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If currentPayload IsNot Nothing Then
            Dim allTrades As List(Of Trade) = Me._parentStrategy.GetSpecificTrades(currentPayload, Me._parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                Dim totalTurnover As Double = allTrades.Sum(Function(x)
                                                                Return x.EntryPrice * x.Quantity
                                                            End Function)
                Dim totalQuantity As Double = allTrades.Sum(Function(x)
                                                                Return x.Quantity
                                                            End Function)

                ret = totalTurnover / totalQuantity
            End If
        End If
        Return ret
    End Function
End Class