Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class OptionPairHLStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _quantity As Integer
    Private _supertrendPayload As Dictionary(Of Date, Decimal)
    Private _supertrendColorPayload As Dictionary(Of Date, Color)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _quantity = Me.LotSize
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.Supertrend.CalculateSupertrend(8, 2, _signalPayload, _supertrendPayload, _supertrendColorPayload)
    End Sub
    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Payload, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = signalCandleSatisfied.Item3
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim buffer As Decimal = 1
                If signalCandleSatisfied.Item4 = Trade.TypeOfOrder.Market Then buffer = 0
                Dim entryPrice As Decimal = signalCandleSatisfied.Item2 + buffer
                Dim quantity As Decimal = _quantity
                Dim slPoint As Decimal = entryPrice - signalCandle.Low
                Dim targetPoint As Decimal = slPoint

                parameter1 = New PlaceOrderParameters With {
                        .entryPrice = entryPrice,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .quantity = quantity,
                        .Stoploss = .EntryPrice - slPoint,
                        .Target = .EntryPrice + targetPoint,
                        .buffer = buffer,
                        .signalCandle = signalCandle,
                        .OrderType = signalCandleSatisfied.Item4,
                        .Supporting1 = "Trade 1",
                        .Supporting2 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                        .Supporting3 = signalCandleSatisfied.Item4.ToString
                    }

                targetPoint = slPoint * _userInputs.TargetMultiplier
                parameter2 = New PlaceOrderParameters With {
                        .entryPrice = entryPrice,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .quantity = quantity,
                        .Stoploss = .EntryPrice - slPoint,
                        .Target = .EntryPrice + targetPoint,
                        .buffer = buffer,
                        .signalCandle = signalCandle,
                        .OrderType = signalCandleSatisfied.Item4,
                        .Supporting1 = "Trade 2",
                        .Supporting2 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                        .Supporting3 = signalCandleSatisfied.Item4.ToString
                    }
            End If
        End If
        If parameter1 IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter1, parameter2})
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TypeOfOrder)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TypeOfOrder) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _supertrendColorPayload IsNot Nothing AndAlso _supertrendColorPayload.Count > 0 AndAlso
                _supertrendColorPayload.ContainsKey(candle.PayloadDate) AndAlso _supertrendColorPayload(candle.PayloadDate) = Color.Green Then
                If IsEngulfingCandle(candle) Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TypeOfOrder)(True, candle.High, candle, Trade.TypeOfOrder.SL)
                Else
                    Dim startTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                                     _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)
                    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then startTime = lastExecutedTrade.ExitTime
                    Dim validCandle As Payload = GetPreviousValidEngulfingCandle(candle.PayloadDate, startTime)
                    If validCandle IsNot Nothing Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TypeOfOrder)(True, currentTick.Open, validCandle, Trade.TypeOfOrder.Market)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsEngulfingCandle(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If candle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                If candle.CandleColor = Color.Green AndAlso candle.PreviousCandlePayload.CandleColor = Color.Red Then
                    If candle.High >= candle.PreviousCandlePayload.High AndAlso candle.Low <= candle.PreviousCandlePayload.Low Then
                        ret = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetPreviousValidEngulfingCandle(ByVal beforeThisTime As Date, ByVal afterThisTime As Date) As Payload
        Dim ret As Payload = Nothing
        Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Payload)) = _signalPayload.Where(Function(x)
                                                                                                         Return x.Key > afterThisTime AndAlso
                                                                                                         x.Key <= beforeThisTime
                                                                                                     End Function)
        If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
            For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                             Return x.Key
                                                                         End Function)
                If IsEngulfingCandle(runningPayload.Value) Then
                    Dim startTime As Date = runningPayload.Value.PayloadDate.AddMinutes(Me._parentStrategy.SignalTimeFrame)
                    Dim endTime As Date = beforeThisTime.AddMinutes(Me._parentStrategy.SignalTimeFrame)
                    If Not IsSignalTriggered(runningPayload.Value.Low, Trade.TradeExecutionDirection.Sell, startTime, endTime) Then
                        ret = runningPayload.Value
                    End If
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
End Class