Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayHKTrendBollingerStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetPrice As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrHKPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighHKPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowHKPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaHKPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _firstCandleOfTheDay As Payload = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    _firstCandleOfTheDay = runningPayload.Value
                    Exit For
                End If
            End If
        Next

        If _firstCandleOfTheDay IsNot Nothing Then
            Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
            Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)

            Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
            Indicator.ATR.CalculateATR(14, _hkPayload, _atrHKPayload)
            Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _hkPayload, _bollingerHighHKPayload, _bollingerLowHKPayload, _smaHKPayload)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                Dim targetPoint As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim quantity As Integer = GetQuantityToTrade(currentMinuteCandle, entryPrice, targetPoint, signal.Item4)
                If quantity > 0 Then
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - 10000000,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + 10000000,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryPrice <> signal.Item2 OrElse
                    currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
            If lastExecutedOrder IsNot Nothing Then
                If currentTrade.PotentialTarget <> lastExecutedOrder.PotentialTarget Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, lastExecutedOrder.PotentialTarget, Math.Abs(lastExecutedOrder.PotentialTarget - currentTrade.EntryPrice))
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim tradeCount As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol)
            If _direction = Trade.TradeExecutionDirection.Buy Then
                If tradeCount < 1 AndAlso _firstCandleOfTheDay.Open < _bollingerLowPayload(_firstCandleOfTheDay.PayloadDate) Then   'Gap Down
                    Dim candle As Payload = currentCandle.PreviousCandlePayload
                    Dim bolliengerLow As Decimal = _bollingerLowPayload(candle.PayloadDate)
                    If candle.Open < bolliengerLow OrElse candle.Close < bolliengerLow Then
                        Dim entryPrice As Decimal = GetEntryPrice(candle, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, candle, Trade.TradeExecutionDirection.Buy)
                    End If
                Else
                    Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                    Dim bollingerLow As Decimal = _bollingerLowHKPayload(hkCandle.PayloadDate)
                    If Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.High, 2) AndAlso hkCandle.Close < bollingerLow Then
                        Dim entryPrice As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Buy)
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                        If lastExecutedOrder IsNot Nothing Then
                            Dim firstExecutedTrade As Trade = _parentStrategy.GetFirstExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                            Dim minDistance As Decimal = GetHighestATR(_atrHKPayload, firstExecutedTrade.SignalCandle.PayloadDate, hkCandle.PayloadDate)
                            If entryPrice <= lastExecutedOrder.EntryPrice - minDistance Then
                                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Buy)
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Buy)
                        End If
                    End If
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                If tradeCount < 1 AndAlso _firstCandleOfTheDay.Open > _bollingerHighPayload(_firstCandleOfTheDay.PayloadDate) Then   'Gap Up
                    Dim candle As Payload = currentCandle.PreviousCandlePayload
                    Dim bolliengerHigh As Decimal = _bollingerHighPayload(candle.PayloadDate)
                    If candle.Open > bolliengerHigh OrElse candle.Close > bolliengerHigh Then
                        Dim entryPrice As Decimal = GetEntryPrice(candle, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, candle, Trade.TradeExecutionDirection.Sell)
                    End If
                Else
                    Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                    Dim bollingerHigh As Decimal = _bollingerHighHKPayload(hkCandle.PayloadDate)
                    If Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.Low, 2) AndAlso hkCandle.Close > bollingerHigh Then
                        Dim entryPrice As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Sell)
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                        If lastExecutedOrder IsNot Nothing Then
                            Dim firstExecutedTrade As Trade = _parentStrategy.GetFirstExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                            Dim minDistance As Decimal = GetHighestATR(_atrHKPayload, firstExecutedTrade.SignalCandle.PayloadDate, hkCandle.PayloadDate)
                            If entryPrice >= lastExecutedOrder.EntryPrice + minDistance Then
                                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Sell)
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Sell)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal atrPayload As Dictionary(Of Date, Decimal), ByVal fromTime As Date, ByVal toTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If atrPayload IsNot Nothing AndAlso atrPayload.Count > 0 Then
            ret = atrPayload.Max(Function(x)
                                     If x.Key >= fromTime AndAlso x.Key <= toTime Then
                                         Return x.Value
                                     Else
                                         Return Decimal.MinValue
                                     End If
                                 End Function)
        End If
        Return ret
    End Function

    Private Function GetEntryPrice(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If candle IsNot Nothing AndAlso direction <> Trade.TradeExecutionDirection.None Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                ret = ConvertFloorCeling(candle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If ret = Math.Round(candle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                    ret = ret + buffer
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                ret = ConvertFloorCeling(candle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If ret = Math.Round(candle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                    ret = ret - buffer
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantityToTrade(ByVal candle As Payload, ByVal entryPrice As Decimal, ByVal targetPoint As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Integer
        Dim ret As Integer = 0
        Dim exitPrice As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            exitPrice = entryPrice + targetPoint
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            exitPrice = entryPrice - targetPoint
        End If
        If exitPrice <> Decimal.MinValue Then
            Dim overallPL As Decimal = 0
            Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
            If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
                For Each runningTrade In inprogressTrades
                    Dim buyPrice As Decimal = Decimal.MinValue
                    Dim sellPrice As Decimal = Decimal.MinValue
                    If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        buyPrice = runningTrade.EntryPrice
                        sellPrice = exitPrice
                    ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        sellPrice = runningTrade.EntryPrice
                        buyPrice = exitPrice
                    End If

                    overallPL += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, buyPrice, sellPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                Next
            End If

            Dim plToAchive As Decimal = _userInputs.TargetPrice - overallPL
            ret = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, plToAchive, _parentStrategy.StockType)
        End If
        Return ret
    End Function
End Class