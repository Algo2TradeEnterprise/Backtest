Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class NikhilKumarStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPoint As Decimal
        Public FirstTargetPoint As Decimal
        Public TrailingPoint As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _quantity As Integer

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _quantity = 5
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                Dim lastTrade As Trade = GetLastCancelTrade(currentMinuteCandlePayload)
                If lastTrade Is Nothing OrElse
                    lastTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    If lastTrade Is Nothing Then
                        signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    Else
                        If lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Cancel Then
                            signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                        ElseIf lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            If currentMinuteCandlePayload.PayloadDate > lastTrade.ExitTime Then
                                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                            End If
                        End If
                    End If
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then


                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                    Dim buffer As Decimal = 1
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item3 + buffer
                    Dim quantity As Decimal = _quantity * Me.LotSize
                    Dim slPoint As Decimal = _userInputs.StoplossPoint
                    Dim targetPoint As Decimal = _userInputs.FirstTargetPoint

                    If currentMinuteCandlePayload.Open < entryPrice Then
                        parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = "Trade 1"
                            }

                        'parameter2 = New PlaceOrderParameters With {
                        '        .EntryPrice = entryPrice,
                        '        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        '        .Quantity = quantity,
                        '        .Stoploss = .EntryPrice - slPoint,
                        '        .Target = .EntryPrice + 1000000,
                        '        .Buffer = buffer,
                        '        .SignalCandle = signalCandle,
                        '        .OrderType = Trade.TypeOfOrder.SL,
                        '        .Supporting1 = "Trade 2"
                        '    }
                    End If
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                    Dim buffer As Decimal = 1
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item3 - buffer
                    Dim quantity As Decimal = _quantity * Me.LotSize
                    Dim slPoint As Decimal = _userInputs.StoplossPoint
                    Dim targetPoint As Decimal = _userInputs.FirstTargetPoint

                    If currentMinuteCandlePayload.Open > entryPrice Then
                        parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = "Trade 1"
                            }

                        'parameter2 = New PlaceOrderParameters With {
                        '        .EntryPrice = entryPrice,
                        '        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        '        .Quantity = quantity,
                        '        .Stoploss = .EntryPrice + slPoint,
                        '        .Target = .EntryPrice - 1000000,
                        '        .Buffer = buffer,
                        '        .SignalCandle = signalCandle,
                        '        .OrderType = Trade.TypeOfOrder.SL,
                        '        .Supporting1 = "Trade 2"
                        '    }
                    End If
                End If
            End If
        End If
        If parameter1 IsNot Nothing AndAlso parameter2 IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter1, parameter2})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If signalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                If plPoint > _userInputs.TrailingPoint Then
                    Dim multiplier As Decimal = Math.Floor(plPoint / _userInputs.TrailingPoint)
                    Dim initialStoploss As Decimal = currentTrade.EntryPrice - _userInputs.StoplossPoint
                    Dim potentialSL As Decimal = initialStoploss + multiplier * _userInputs.TrailingPoint
                    If potentialSL > currentTrade.PotentialStopLoss Then
                        triggerPrice = potentialSL
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                If plPoint > _userInputs.TrailingPoint Then
                    Dim multiplier As Decimal = Math.Floor(plPoint / _userInputs.TrailingPoint)
                    Dim initialStoploss As Decimal = currentTrade.EntryPrice + _userInputs.StoplossPoint
                    Dim potentialSL As Decimal = initialStoploss - multiplier * _userInputs.TrailingPoint
                    If potentialSL < currentTrade.PotentialStopLoss Then
                        triggerPrice = potentialSL
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("SL Moved at {0} on {1}",
                                                                                               currentTick.Open,
                                                                                               currentTick.PayloadDate.ToString("HH:mm:ss")))
            End If
        End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If candle.CandleColor = Color.Green AndAlso candle.PreviousCandlePayload.CandleColor = Color.Green Then
                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal)(True, Trade.TradeExecutionDirection.Buy, Math.Max(candle.High, candle.PreviousCandlePayload.High))
            ElseIf candle.CandleColor = Color.Red AndAlso candle.PreviousCandlePayload.CandleColor = Color.Red Then
                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Decimal)(True, Trade.TradeExecutionDirection.Sell, Math.Min(candle.Low, candle.PreviousCandlePayload.Low))
            End If
        End If
        Return ret
    End Function

    Private Function GetLastCancelTrade(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        Dim potentialCancelTrades As List(Of Trade) = Me._parentStrategy.GetSpecificTrades(currentPayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Cancel)
        Dim potentialCloseTrades As List(Of Trade) = Me._parentStrategy.GetSpecificTrades(currentPayload, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim potentialTrades As List(Of Trade) = New List(Of Trade)
        If potentialCancelTrades IsNot Nothing AndAlso potentialCancelTrades.Count > 0 Then potentialTrades.AddRange(potentialCancelTrades)
        If potentialCloseTrades IsNot Nothing AndAlso potentialCloseTrades.Count > 0 Then potentialTrades.AddRange(potentialCloseTrades)
        If potentialTrades IsNot Nothing AndAlso potentialTrades.Count > 0 Then
            ret = potentialTrades.OrderBy(Function(x)
                                              Return x.EntryTime
                                          End Function).LastOrDefault
        End If
        Return ret
    End Function
End Class
