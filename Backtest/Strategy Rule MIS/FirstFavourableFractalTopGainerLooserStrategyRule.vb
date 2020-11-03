Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FirstFavourableFractalTopGainerLooserStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public MaxProfitPerTrade As Decimal
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _dayATR As Decimal
    Private ReadOnly _direction As Trade.TradeExecutionDirection
    Private ReadOnly _favourableFractalTime As Date
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal dayATR As Decimal,
                   ByVal direction As Integer,
                   ByVal fractalTime As Date)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _dayATR = dayATR
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        Else
            _direction = Trade.TradeExecutionDirection.Sell
        End If
        _favourableFractalTime = fractalTime
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item4
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item3

                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - Math.Abs(entryPrice - stoploss), _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, signal.Item5, Trade.TypeOfStock.Cash)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = signal.Item5,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                                        }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private _lastSignalHighLevel As Payload
    Private _lastSignalLowLevel As Payload
    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentCandle.PreviousCandlePayload.Close, RoundOfType.Floor)
            If currentCandle.PreviousCandlePayload.PayloadDate = _favourableFractalTime Then
                If _direction = Trade.TradeExecutionDirection.Buy Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) _
                        (True, _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) + buffer, GetLowestPointOfTheDay(currentCandle.PreviousCandlePayload) - buffer, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) _
                        (True, _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) - buffer, GetHighestPointOfTheDay(currentCandle.PreviousCandlePayload) + buffer, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestPointOfTheDay(ByVal signalCandle As Payload) As Decimal
        Return _signalPayload.Max(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                          Return x.Value.High
                                      Else
                                          Return Decimal.MinValue
                                      End If
                                  End Function)
    End Function

    Private Function GetLowestPointOfTheDay(ByVal signalCandle As Payload) As Decimal
        Return _signalPayload.Min(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                          Return x.Value.Low
                                      Else
                                          Return Decimal.MaxValue
                                      End If
                                  End Function)
    End Function
End Class