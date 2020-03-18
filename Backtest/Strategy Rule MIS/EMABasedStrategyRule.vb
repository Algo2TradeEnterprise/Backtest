Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class EMABasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public ImmediateBreakout As Boolean
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _emaPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _eligibleToTakeTrade As Boolean

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.EMA.CalculateEMA(3, Payload.PayloadFields.Close, _signalPayload, _emaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
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
                If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.SignalCandle.PayloadDate = signal.Item4.PayloadDate) Then
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim slPoint As Decimal = signalCandle.CandleRange + 2 * buffer
                Dim atr As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                Dim targetPoint As Decimal = Decimal.MinValue
                Dim targetRemark As String = Nothing
                If atr / slPoint >= _userInputs.TargetMultiplier Then
                    targetPoint = atr
                    targetRemark = "ATR Target"
                Else
                    targetPoint = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    targetRemark = "SL Target"
                End If
                Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signal.Item2, signal.Item2 + targetPoint, _userInputs.MaxProfitPerTrade, _parentStrategy.StockType)

                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2 + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate,
                                .Supporting2 = targetRemark,
                                .Supporting3 = atr
                            }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2 - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate,
                                .Supporting2 = targetRemark,
                                .Supporting3 = atr
                            }
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
        If currentTrade IsNot Nothing Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.EntryDirection <> signal.Item3 Then
                        ret = New Tuple(Of Boolean, String)(True, "Reverse Signal Exit")
                    End If
                End If
            ElseIf currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                If _userInputs.ImmediateBreakout Then
                    If currentTrade.SignalCandle.PayloadDate <= currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
                If ret Is Nothing Then
                    Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        If currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
                    End If
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
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            For Each runningTrades In inprogressTrades
                If runningTrades.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If runningTrades.EntryTime >= runningTrades.SignalCandle.PayloadDate.AddMinutes(2) Then
                        runningTrades.UpdateTrade(Supporting4:="Delay breakout")
                    End If
                End If
            Next
        End If
    End Function

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            If Not _eligibleToTakeTrade Then
                If candle.High >= _emaPayload(candle.PayloadDate) AndAlso candle.Low <= _emaPayload(candle.PayloadDate) Then
                    _eligibleToTakeTrade = True
                End If
            ElseIf _eligibleToTakeTrade Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Open, RoundOfType.Floor)
                If candle.CandleRange >= buffer Then
                    If candle.Low > _emaPayload(candle.PayloadDate) Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.Low, Trade.TradeExecutionDirection.Sell, candle)
                    ElseIf candle.High < _emaPayload(candle.PayloadDate) Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.High, Trade.TradeExecutionDirection.Buy, candle)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class