Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowSLPinbarStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumInvestmentPerStock As Decimal
        Public MaxLossPerTrade As Decimal
        Public PinbarTailPercentage As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public StopAtFirstTarget As Boolean
        Public AllowMomentumReversal As Boolean
    End Class
#End Region

    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _stockATR As Decimal
    Private ReadOnly _dayATR As Decimal
    Private ReadOnly _slPoint As Decimal
    Private ReadOnly _quantity As Integer

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockATR As Decimal,
                   ByVal dayATR As Decimal,
                   ByVal slPoint As Decimal,
                   ByVal quantity As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _stockATR = stockATR
        _dayATR = dayATR
        _slPoint = slPoint
        _quantity = quantity
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _ATRPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            (Not _userInputs.StopAtFirstTarget OrElse Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS)) AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If Not IsLastTradeExitedAtCurrentCandle(currentMinuteCandlePayload) Then
                    signalCandle = signalCandleSatisfied.Item5
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = _quantity,
                                .Stoploss = .EntryPrice - _slPoint,
                                .Target = .EntryPrice + _slPoint * _userInputs.TargetMultiplier,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                            }
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = _quantity,
                                .Stoploss = .EntryPrice + _slPoint,
                                .Target = .EntryPrice - _slPoint * _userInputs.TargetMultiplier,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                            }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            'If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
            '    Dim stockMaxProfitPoint As Decimal = ConvertFloorCeling(GetCandleATR(parameter.SignalCandle), Me._parentStrategy.TickSize, RoundOfType.Celing) * _parentStrategy.StockMaxProfitPercentagePerDay
            '    Me.MaxProfitOfThisStock = Me._parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice + stockMaxProfitPoint, parameter.Quantity, _lotSize, Me._parentStrategy.StockType)
            'End If
            'If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
            '    Dim stockMaxLossPoint As Decimal = ConvertFloorCeling(GetCandleATR(parameter.SignalCandle), Me._parentStrategy.TickSize, RoundOfType.Floor) * _parentStrategy.StockMaxLossPercentagePerDay
            '    Me.MaxLossOfThisStock = Me._parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice - stockMaxLossPoint, parameter.Quantity, _lotSize, Me._parentStrategy.StockType)
            'End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing Then
            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                Dim signalCandle As Payload = currentTrade.SignalCandle
                If signalCandle IsNot Nothing Then
                    Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                    If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                        If signalCandle.PayloadDate <> signalCandleSatisfied.Item5.PayloadDate OrElse
                            currentTrade.EntryDirection <> signalCandleSatisfied.Item4 Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
                    End If
                    If signalCandle.PayloadDate <= currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            If currentMinuteCandlePayload.PreviousCandlePayload.Low < signalCandle.Low Then
                                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                            End If
                        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            If currentMinuteCandlePayload.PreviousCandlePayload.High > signalCandle.High Then
                                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                            End If
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim reason As String = Nothing
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                If _userInputs.BreakevenMovement Then
                    Dim targetPoint As Decimal = _slPoint * _userInputs.TargetMultiplier
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim potentialTarget As Decimal = currentTrade.EntryPrice + ConvertFloorCeling(targetPoint / 2, Me._parentStrategy.TickSize, RoundOfType.Celing)
                        If currentTick.Open > potentialTarget Then
                            triggerPrice = ConvertFloorCeling(currentTrade.EntryPrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                            reason = "Breakeven movement"
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim potentialTarget As Decimal = currentTrade.EntryPrice - ConvertFloorCeling(targetPoint / 2, Me._parentStrategy.TickSize, RoundOfType.Celing)
                        If currentTick.Open < potentialTarget Then
                            triggerPrice = ConvertFloorCeling(currentTrade.EntryPrice, Me._parentStrategy.TickSize, RoundOfType.Celing)
                            reason = "Breakeven movement"
                        End If
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, reason)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            Dim lowBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
            Dim highBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
            'If candle.CandleRange > _ATRPayload(candle.PayloadDate) / 3 AndAlso candle.CandleRange < _ATRPayload(candle.PayloadDate) Then
            If candle.CandleRange > _ATRPayload(candle.PayloadDate) / 3 Then
                If candle.CandleWicks.Top <> 0 AndAlso candle.CandleWicks.Top + lowBuffer >= candle.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                    direction = Trade.TradeExecutionDirection.Sell
                ElseIf candle.CandleWicks.Bottom <> 0 AndAlso candle.CandleWicks.Bottom + highBuffer >= candle.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                    direction = Trade.TradeExecutionDirection.Buy
                End If
            End If
            If direction <> Trade.TradeExecutionDirection.None Then
                If _userInputs.AllowMomentumReversal Then
                    Dim middlePoint As Decimal = (candle.High + candle.Low) / 2
                    Dim range As Decimal = candle.High + highBuffer - middlePoint
                    If currentTick.Open >= middlePoint + range * 30 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.High + highBuffer, candle.High + highBuffer - _slPoint, Trade.TradeExecutionDirection.Buy, candle)
                    ElseIf currentTick.Open <= middlePoint - range * 30 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.Low - lowBuffer, candle.Low - lowBuffer + _slPoint, Trade.TradeExecutionDirection.Sell, candle)
                    End If
                Else
                    If direction = Trade.TradeExecutionDirection.Buy Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.High + highBuffer, candle.High + highBuffer - _slPoint, Trade.TradeExecutionDirection.Buy, candle)
                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.Low - lowBuffer, candle.Low - lowBuffer + _slPoint, Trade.TradeExecutionDirection.Sell, candle)
                    End If
                End If
            ElseIf direction = Trade.TradeExecutionDirection.None Then
                If candle.High <= candle.PreviousCandlePayload.High AndAlso
                    candle.Low >= candle.PreviousCandlePayload.Low Then
                    'If candle.PreviousCandlePayload.CandleRange > _ATRPayload(candle.PreviousCandlePayload.PayloadDate) / 3 AndAlso
                    '    candle.PreviousCandlePayload.CandleRange < _ATRPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    If candle.PreviousCandlePayload.CandleRange > _ATRPayload(candle.PreviousCandlePayload.PayloadDate) / 3 Then
                        Dim previousLowBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.PreviousCandlePayload.Low, RoundOfType.Floor)
                        Dim previousHighBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.PreviousCandlePayload.High, RoundOfType.Floor)
                        If candle.PreviousCandlePayload.CandleWicks.Top <> 0 AndAlso
                            candle.PreviousCandlePayload.CandleWicks.Top + previousLowBuffer >= candle.PreviousCandlePayload.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.PreviousCandlePayload.Low - previousLowBuffer, candle.PreviousCandlePayload.Low - previousLowBuffer + _slPoint, Trade.TradeExecutionDirection.Sell, candle.PreviousCandlePayload)
                        ElseIf candle.PreviousCandlePayload.CandleWicks.Bottom <> 0 AndAlso
                            candle.PreviousCandlePayload.CandleWicks.Bottom + previousHighBuffer >= candle.PreviousCandlePayload.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection, Payload)(True, candle.PreviousCandlePayload.High + previousHighBuffer, candle.PreviousCandlePayload.High + previousHighBuffer - _slPoint, Trade.TradeExecutionDirection.Buy, candle.PreviousCandlePayload)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastOrderExitTime(ByVal currentCandle As Payload) As Date
        Dim ret As Date = Date.MinValue
        Dim lastExecutedOrder As Trade = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Me._parentStrategy.TradeType)
        If lastExecutedOrder IsNot Nothing Then
            ret = lastExecutedOrder.ExitTime
        End If
        Return ret
    End Function

    Private Function IsLastTradeExitedAtCurrentCandle(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim lastTradeExitTime As Date = GetLastOrderExitTime(currentCandle)
        If lastTradeExitTime <> Date.MinValue Then
            Dim blockDateInThisTimeframe As Date = Date.MinValue
            Dim timeframe As Integer = Me._parentStrategy.SignalTimeFrame
            If Me._parentStrategy.ExchangeStartTime.Minutes Mod timeframe = 0 Then
                blockDateInThisTimeframe = New Date(lastTradeExitTime.Year,
                                                    lastTradeExitTime.Month,
                                                    lastTradeExitTime.Day,
                                                    lastTradeExitTime.Hour,
                                                    Math.Floor(lastTradeExitTime.Minute / timeframe) * timeframe, 0)
            Else
                Dim exchangeStartTime As Date = New Date(lastTradeExitTime.Year, lastTradeExitTime.Month, lastTradeExitTime.Day, Me._parentStrategy.ExchangeStartTime.Hours, Me._parentStrategy.ExchangeStartTime.Minutes, 0)
                Dim currentTime As Date = New Date(lastTradeExitTime.Year, lastTradeExitTime.Month, lastTradeExitTime.Day, lastTradeExitTime.Hour, lastTradeExitTime.Minute, 0)
                Dim timeDifference As Double = currentTime.Subtract(exchangeStartTime).TotalMinutes
                Dim adjustedTimeDifference As Integer = Math.Floor(timeDifference / timeframe) * timeframe
                Dim currentMinute As Date = exchangeStartTime.AddMinutes(adjustedTimeDifference)
                blockDateInThisTimeframe = New Date(lastTradeExitTime.Year, lastTradeExitTime.Month, lastTradeExitTime.Day, currentMinute.Hour, currentMinute.Minute, 0)
            End If
            If blockDateInThisTimeframe <> Date.MinValue Then
                ret = Utilities.Time.IsDateTimeEqualTillMinutes(blockDateInThisTimeframe, currentCandle.PayloadDate)
            End If
        End If
        Return ret
    End Function
End Class
