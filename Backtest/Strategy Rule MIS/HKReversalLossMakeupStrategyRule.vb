Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKReversalLossMakeupStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signal.Item4,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item4.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item5,
                                .Supporting3 = "Original"
                            }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload, entryPrice)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, target, Math.Abs(pl), Trade.TypeOfStock.Cash)

                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signal.Item4,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item4.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = "Loss Makeup"
                                }
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signal.Item4,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item4.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = signal.Item5,
                                .Supporting3 = "Original"
                            }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload, entryPrice)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, target, entryPrice, Math.Abs(pl), Trade.TypeOfStock.Cash)

                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signal.Item4,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signal.Item4.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = "Loss Makeup"
                                }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters) From {parameter1}
        End If
        If parameter2 IsNot Nothing Then
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, String) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If

            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS)
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If currentTrade.EntryTime < lastExecutedTrade.EntryTime Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, Decimal, Payload, String)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, String) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
            If buyLevel = Math.Round(hkCandle.High, 2) Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.High, RoundOfType.Floor)
                buyLevel = buyLevel + buffer
            End If
            Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
            If sellLevel = Math.Round(hkCandle.Low, 2) Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.Low, RoundOfType.Floor)
                sellLevel = sellLevel - buffer
            End If
            Dim atr As Decimal = ConvertFloorCeling(GetHighestATR(hkCandle), _parentStrategy.TickSize, RoundOfType.Floor)

            If direction = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                If buyLevel - atr > sellLevel Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, String)(True, buyLevel, buyLevel - atr, hkCandle, "ATR")
                Else
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, String)(True, buyLevel, sellLevel, hkCandle, "CR")
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                If sellLevel + atr < buyLevel Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, String)(True, sellLevel, sellLevel + atr, hkCandle, "ATR")
                Else
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, String)(True, sellLevel, buyLevel, hkCandle, "CR")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            ret = _atrPayload.Max(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                          Return x.Value
                                      Else
                                          Return Decimal.MinValue
                                      End If
                                  End Function)
        End If
        Return ret
    End Function

    Private Function GetStockPotentialPL(ByVal currentMinutePayload As Payload, ByVal potentialExitPrice As Decimal) As Decimal
        Dim ret As Decimal = 0
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            ret = completeTrades.Sum(Function(x)
                                         Return x.PLAfterBrokerage
                                     End Function)
        End If
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            For Each runningTrade In inprogressTrades
                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, runningTrade.EntryPrice, potentialExitPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, potentialExitPrice, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                End If
            Next
        End If
        Return ret
    End Function
End Class