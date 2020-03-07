Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighLowEMAStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _atrPayloads As Dictionary(Of Date, Decimal)
    Private _highEMAPayloads As Dictionary(Of Date, Decimal)
    Private _lowEMAPayloads As Dictionary(Of Date, Decimal)
    Private _eligibleToTakeTrade As Boolean = False

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayloads)
        Indicator.EMA.CalculateEMA(5, Payload.PayloadFields.High, _signalPayload, _highEMAPayloads)
        Indicator.EMA.CalculateEMA(5, Payload.PayloadFields.Low, _signalPayload, _lowEMAPayloads)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
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
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If Not (lastExecutedOrder IsNot Nothing AndAlso
                    _parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload) > signal.Item4.PayloadDate) Then
                    If IsEligibleToTakeTrade(currentMinuteCandlePayload) Then signalCandle = signal.Item4
                ElseIf lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.ExitCondition = Trade.TradeExitCondition.ForceExit Then
                    If IsEligibleToTakeTrade(currentMinuteCandlePayload) Then signalCandle = signal.Item4
                End If
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim slPoint As Decimal = ConvertFloorCeling(_atrPayloads(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = signal.Item2
                    Dim stoploss As Decimal = entry - slPoint
                    Dim target As Decimal = entry + targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = signal.Item2
                    Dim stoploss As Decimal = entry + slPoint
                    Dim target As Decimal = entry - targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
                                }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If _lastSignalCandleWithDirection IsNot Nothing Then
                If currentTrade.EntryDirection <> _lastSignalCandleWithDirection.Item2 Then
                    ret = New Tuple(Of Boolean, String)(True, "Opposite direction signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private _lastSignalCandleWithDirection As Tuple(Of Payload, Trade.TradeExecutionDirection) = Nothing
    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If Not _eligibleToTakeTrade Then
                If (candle.Open >= _lowEMAPayloads(candle.PayloadDate) AndAlso candle.Open <= _highEMAPayloads(candle.PayloadDate)) OrElse
                    (candle.Close >= _lowEMAPayloads(candle.PayloadDate) AndAlso candle.Close <= _highEMAPayloads(candle.PayloadDate)) Then
                    _eligibleToTakeTrade = True
                End If
            ElseIf _eligibleToTakeTrade AndAlso currentTick.PreviousCandlePayload Is Nothing Then
                If candle.Close > _highEMAPayloads(candle.PayloadDate) AndAlso
                    candle.PreviousCandlePayload.Close > _highEMAPayloads(candle.PreviousCandlePayload.PayloadDate) Then
                    _lastSignalCandleWithDirection = New Tuple(Of Payload, Trade.TradeExecutionDirection)(candle, Trade.TradeExecutionDirection.Buy)
                ElseIf candle.Close < _lowEMAPayloads(candle.PayloadDate) AndAlso
                    candle.PreviousCandlePayload.Close < _lowEMAPayloads(candle.PreviousCandlePayload.PayloadDate) Then
                    _lastSignalCandleWithDirection = New Tuple(Of Payload, Trade.TradeExecutionDirection)(candle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
            If _lastSignalCandleWithDirection IsNot Nothing Then
                If _lastSignalCandleWithDirection.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buyPrice As Decimal = ConvertFloorCeling(_lowEMAPayloads(candle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    If currentTick.Open <= buyPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, buyPrice, Trade.TradeExecutionDirection.Buy, _lastSignalCandleWithDirection.Item1)
                    End If
                ElseIf _lastSignalCandleWithDirection.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim sellPrice As Decimal = ConvertFloorCeling(_highEMAPayloads(candle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                    If currentTick.Open >= sellPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, sellPrice, Trade.TradeExecutionDirection.Sell, _lastSignalCandleWithDirection.Item1)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsEligibleToTakeTrade(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = True
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then
            Dim lossCount As Integer = 0
            Dim profitCount As Integer = 0
            For Each runningTrade In closeTrades.OrderByDescending(Function(x)
                                                                       Return x.ExitTime
                                                                   End Function)
                If runningTrade.PLPoint < 0 Then
                    lossCount += 1
                    If lossCount = 2 Then
                        ret = False
                        Exit For
                    End If
                Else
                    lossCount = 0
                End If

                'If runningTrade.PLPoint > 0 Then
                '    profitCount += 1
                '    If profitCount = 2 Then
                '        ret = False
                '        Exit For
                '    End If
                'Else
                '    profitCount = 0
                'End If
            Next
        End If
        Return ret
    End Function
End Class