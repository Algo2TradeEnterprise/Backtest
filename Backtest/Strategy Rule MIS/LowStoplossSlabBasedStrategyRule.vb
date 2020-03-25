Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowStoplossSlabBasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum SlabMTMType
        Individual = 1
        Net
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetSlabMultiplier As Decimal
        Public StoplossSlabMultiplier As Decimal
        Public ExitAtStockSlabMTM As Boolean
        Public TypeOfSlabMTM As SlabMTMType
        Public SlabMTMTarget As Decimal
        Public SlabMTMStoploss As Decimal
    End Class
#End Region

    Private _hkPayloads As Dictionary(Of Date, Payload)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _slab As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _slab = slab
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder Is Nothing OrElse lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    Dim quantity As Decimal = Me.LotSize
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim slPoint As Decimal = _slab * _userInputs.StoplossSlabMultiplier
                    Dim target As Decimal = _slab * _userInputs.TargetSlabMultiplier
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                            If currentTick.Open < signal.Item2 + buffer Then
                                parameter = New PlaceOrderParameters With {
                                            .EntryPrice = signal.Item2 + buffer,
                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                            .Quantity = quantity,
                                            .Stoploss = signal.Item2 - slPoint - buffer,
                                            .Target = .EntryPrice + target,
                                            .Buffer = buffer,
                                            .SignalCandle = signal.Item3,
                                            .OrderType = Trade.TypeOfOrder.SL,
                                            .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = _slab
                                        }
                            End If
                        End If
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                            If currentTick.Open > signal.Item2 - buffer Then
                                parameter = New PlaceOrderParameters With {
                                            .EntryPrice = signal.Item2 - buffer,
                                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                            .Quantity = quantity,
                                            .Stoploss = signal.Item2 + slPoint + buffer,
                                            .Target = .EntryPrice - target,
                                            .Buffer = buffer,
                                            .SignalCandle = signal.Item3,
                                            .OrderType = Trade.TypeOfOrder.SL,
                                            .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = _slab
                                        }
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            If _userInputs.ExitAtStockSlabMTM AndAlso Me.MaxLossOfThisStock = Decimal.MinValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                If _userInputs.TypeOfSlabMTM = SlabMTMType.Net Then
                    Dim slPoint As Decimal = _slab * _userInputs.SlabMTMStoploss + parameter.Buffer * 2 * _userInputs.SlabMTMStoploss
                    Dim targetPoint As Decimal = _slab * _userInputs.SlabMTMTarget
                    Dim projectedLoss As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice - slPoint, parameter.Quantity, Me.LotSize, _parentStrategy.StockType)
                    Dim projectedProfit As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice + targetPoint, parameter.Quantity, Me.LotSize, _parentStrategy.StockType)
                    Me.MaxLossOfThisStock = projectedLoss
                    Me.MaxProfitOfThisStock = projectedProfit
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection = signal.Item4 AndAlso currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
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
            Dim slPoint As Decimal = _slab * _userInputs.StoplossSlabMultiplier + 2 * currentTrade.StoplossBuffer
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                triggerPrice = currentTrade.EntryPrice - slPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                triggerPrice = currentTrade.EntryPrice + slPoint
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, slPoint)
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

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayloads(candle.PayloadDate)
            If hkCandle.PreviousCandlePayload IsNot Nothing Then
                If hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    Dim buyLevel As Decimal = GetSlabBasedLevel(hkCandle.PreviousCandlePayload.High, Trade.TradeExecutionDirection.Buy)
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                ElseIf hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    Dim sellLevel As Decimal = GetSlabBasedLevel(hkCandle.PreviousCandlePayload.Low, Trade.TradeExecutionDirection.Sell)
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class