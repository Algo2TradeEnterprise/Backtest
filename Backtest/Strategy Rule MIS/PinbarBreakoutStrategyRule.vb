Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PinbarBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumInvestmentPerStock As Decimal
        Public TargetMultiplier As Decimal
        Public PinbarTailPercentage As Decimal
        Public MaxLossPerTradeMultiplier As Decimal
        Public MinLossPercentagePerTrade As Decimal
    End Class
#End Region

    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _firstTradedQuantity As Integer = Integer.MinValue
    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _dayATR As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal dayATR As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _dayATR = dayATR
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
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If IsLastTradeForceExitForCandleClose(currentMinuteCandlePayload) OrElse Not IsLastTradeExitedAtCurrentCandle(currentMinuteCandlePayload) Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    If _firstTradedQuantity = Integer.MinValue Then
                        _firstTradedQuantity = Me._parentStrategy.CalculateQuantityFromInvestment(_lotSize, _userInputs.MinimumInvestmentPerStock, signalCandleSatisfied.Item2, Me._parentStrategy.StockType, True)
                    End If
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                Dim targetPoint As Decimal = ConvertFloorCeling(GetCandleATR(signalCandle) * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Celing)
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = _firstTradedQuantity,
                                .Stoploss = signalCandleSatisfied.Item3,
                                .Target = .EntryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = GetSignalCandleType(signalCandle, Trade.TradeExecutionDirection.Buy),
                                .Supporting3 = GetCandleATR(signalCandle)
                            }
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = _firstTradedQuantity,
                                .Stoploss = signalCandleSatisfied.Item3,
                                .Target = .EntryPrice - targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = GetSignalCandleType(signalCandle, Trade.TradeExecutionDirection.Sell),
                                .Supporting3 = GetCandleATR(signalCandle)
                            }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Dim stockMaxProfitPoint As Decimal = ConvertFloorCeling(GetCandleATR(parameter.SignalCandle), Me._parentStrategy.TickSize, RoundOfType.Celing) * _parentStrategy.StockMaxProfitPercentagePerDay
                Me.MaxProfitOfThisStock = Me._parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice + stockMaxProfitPoint, parameter.Quantity, _lotSize, Me._parentStrategy.StockType)
            End If
            If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
                Dim stockMaxLossPoint As Decimal = ConvertFloorCeling(GetCandleATR(parameter.SignalCandle), Me._parentStrategy.TickSize, RoundOfType.Floor) * _parentStrategy.StockMaxLossPercentagePerDay
                Me.MaxLossOfThisStock = Me._parentStrategy.CalculatePL(_tradingSymbol, parameter.EntryPrice, parameter.EntryPrice - stockMaxLossPoint, parameter.Quantity, _lotSize, Me._parentStrategy.StockType)
            End If
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
                    If signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Signal not triggered")
                    End If
                End If
            ElseIf currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                Dim signalCandle As Payload = currentTrade.SignalCandle
                If signalCandle IsNot Nothing Then
                    Dim triggerPrice As Decimal = Decimal.MinValue
                    Dim buffer As Decimal = currentTrade.StoplossBuffer
                    If signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            Dim potentialSLPrice As Decimal = Decimal.MinValue
                            If signalCandle.CandleColor = Color.Red Then
                                potentialSLPrice = signalCandle.Close - buffer
                            Else
                                potentialSLPrice = signalCandle.Open - buffer
                            End If
                            Dim minimusSL As Decimal = currentTrade.EntryPrice * _userInputs.MinLossPercentagePerTrade / 100
                            triggerPrice = Math.Min(potentialSLPrice, ConvertFloorCeling(currentTrade.EntryPrice - minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor))
                            If currentMinuteCandlePayload.PreviousCandlePayload.Close <= triggerPrice Then
                                ret = New Tuple(Of Boolean, String)(True, "Candle close beyond body")
                            End If
                        ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            Dim potentialSLPrice As Decimal = Decimal.MinValue
                            If signalCandle.CandleColor = Color.Red Then
                                potentialSLPrice = signalCandle.Open + buffer
                            Else
                                potentialSLPrice = signalCandle.Close + buffer
                            End If
                            Dim minimusSL As Decimal = currentTrade.EntryPrice * _userInputs.MinLossPercentagePerTrade / 100
                            triggerPrice = Math.Max(potentialSLPrice, ConvertFloorCeling(currentTrade.EntryPrice + minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor))
                            If currentMinuteCandlePayload.PreviousCandlePayload.Close >= triggerPrice Then
                                ret = New Tuple(Of Boolean, String)(True, "Candle close beyond body")
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
                Dim buffer As Decimal = currentTrade.StoplossBuffer
                If signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim potentialSLPrice As Decimal = Decimal.MinValue
                        If signalCandle.CandleColor = Color.Red Then
                            potentialSLPrice = signalCandle.Close - buffer
                        Else
                            potentialSLPrice = signalCandle.Open - buffer
                        End If
                        Dim minimusSL As Decimal = currentTrade.EntryPrice * _userInputs.MinLossPercentagePerTrade / 100
                        If potentialSLPrice <= ConvertFloorCeling(currentTrade.EntryPrice - minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                            triggerPrice = potentialSLPrice
                            reason = "Move to candle body"
                        Else
                            triggerPrice = ConvertFloorCeling(currentTrade.EntryPrice - minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor)
                            reason = "Minimum loss % per trade"
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim potentialSLPrice As Decimal = Decimal.MinValue
                        If signalCandle.CandleColor = Color.Red Then
                            potentialSLPrice = signalCandle.Open + buffer
                        Else
                            potentialSLPrice = signalCandle.Close + buffer
                        End If
                        Dim minimusSL As Decimal = currentTrade.EntryPrice * _userInputs.MinLossPercentagePerTrade / 100
                        If potentialSLPrice >= ConvertFloorCeling(currentTrade.EntryPrice + minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                            triggerPrice = potentialSLPrice
                            reason = "Move to candle body"
                        Else
                            triggerPrice = ConvertFloorCeling(currentTrade.EntryPrice + minimusSL, Me._parentStrategy.TickSize, RoundOfType.Floor)
                            reason = "Minimum loss % per trade"
                        End If
                    End If
                End If
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    If currentMinuteCandlePayload.PreviousCandlePayload.Low > currentTrade.EntryPrice Then
                        Dim potentialPrice As Decimal = currentTrade.EntryPrice + GetMovementPoint(currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection)
                        triggerPrice = ConvertFloorCeling(potentialPrice, Me._parentStrategy.TickSize, RoundOfType.Celing)
                        reason = "Breakeven movement"
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    If currentMinuteCandlePayload.PreviousCandlePayload.High < currentTrade.EntryPrice Then
                        Dim potentialPrice As Decimal = currentTrade.EntryPrice - GetMovementPoint(currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection)
                        triggerPrice = ConvertFloorCeling(potentialPrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        reason = "Breakeven movement"
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim slPoint As Decimal = candle.CandleRange
            Dim lowBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
            Dim highBuffer As Decimal = Me._parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
            If candle.CandleWicks.Top + lowBuffer >= candle.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                slPoint = slPoint + 2 * lowBuffer
                Dim potentialSLPrice As Decimal = Decimal.MinValue
                If candle.CandleColor = Color.Red Then
                    potentialSLPrice = candle.Open
                Else
                    potentialSLPrice = candle.Close
                End If
                If Math.Abs(potentialSLPrice - candle.Low) <= (GetCandleATR(candle) + lowBuffer) * _userInputs.MaxLossPerTradeMultiplier Then
                    If slPoint < candle.Low * _userInputs.MinLossPercentagePerTrade / 100 Then
                        slPoint = ConvertFloorCeling(candle.Low * _userInputs.MinLossPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, candle.Low - lowBuffer, candle.Low - lowBuffer + slPoint, Trade.TradeExecutionDirection.Sell)
                End If
            ElseIf candle.CandleWicks.Bottom + highBuffer >= candle.CandleRange * _userInputs.PinbarTailPercentage / 100 Then
                slPoint = slPoint + 2 * highBuffer
                Dim potentialSLPrice As Decimal = Decimal.MinValue
                If candle.CandleColor = Color.Red Then
                    potentialSLPrice = candle.Close
                Else
                    potentialSLPrice = candle.Open
                End If
                If Math.Abs(candle.High - potentialSLPrice) <= (GetCandleATR(candle) + highBuffer) * _userInputs.MaxLossPerTradeMultiplier Then
                    If slPoint < candle.High * _userInputs.MinLossPercentagePerTrade / 100 Then
                        slPoint = ConvertFloorCeling(candle.High * _userInputs.MinLossPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, candle.High + highBuffer, candle.High + highBuffer - slPoint, Trade.TradeExecutionDirection.Buy)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetCandleBody(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If candle IsNot Nothing Then
            If candle.CandleColor = Color.Red Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = candle.High - candle.Close
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    ret = candle.Open - candle.Low
                End If
            Else
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = candle.High - candle.Open
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    ret = candle.Close - candle.Low
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetCandleATR(ByVal candle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If candle IsNot Nothing Then
            If _ATRPayload IsNot Nothing AndAlso _ATRPayload.Count > 0 AndAlso
                _ATRPayload.ContainsKey(candle.PayloadDate) Then
                ret = Math.Round(_ATRPayload(candle.PayloadDate), 2)
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

    Private Function IsLastTradeForceExitForCandleClose(ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim lastExecutedOrder As Trade = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Me._parentStrategy.TradeType)
        If lastExecutedOrder IsNot Nothing Then
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
                Dim signalCandle As Payload = lastExecutedOrder.SignalCandle
                If blockDateInThisTimeframe = signalCandle.PayloadDate.AddMinutes(2 * timeframe) Then
                    Dim blockCandle As Payload = _signalPayload(blockDateInThisTimeframe)
                    If lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                        Dim potentialExitPrice As Decimal = signalCandle.High - GetCandleBody(signalCandle, Trade.TradeExecutionDirection.Buy) - buffer
                        If blockCandle.Close > signalCandle.Low AndAlso
                            blockCandle.Close <= potentialExitPrice Then
                            ret = True
                        End If
                    ElseIf lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                        Dim potentialExitPrice As Decimal = signalCandle.Low + GetCandleBody(signalCandle, Trade.TradeExecutionDirection.Sell) + buffer
                        If blockCandle.Close < signalCandle.High AndAlso
                            blockCandle.Close > potentialExitPrice Then
                            ret = True
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetMovementPoint(ByVal entryPrice As Decimal, ByVal quantity As Integer, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Me._parentStrategy.TickSize
        Dim potentialExitPoint As Decimal = ConvertFloorCeling(entryPrice * 0.1 / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
        If direction = Trade.TradeExecutionDirection.Buy Then
            For exitPrice As Decimal = (entryPrice - potentialExitPoint) To Decimal.MaxValue Step ret
                Dim pl As Decimal = Me._parentStrategy.CalculatePL(_tradingSymbol, entryPrice, exitPrice, quantity, _lotSize, Me._parentStrategy.StockType)
                If pl >= -100 Then
                    ret = ConvertFloorCeling(exitPrice - entryPrice, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    Exit For
                End If
            Next
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            For exitPrice As Decimal = (entryPrice + potentialExitPoint) To Decimal.MinValue Step ret * -1
                Dim pl As Decimal = Me._parentStrategy.CalculatePL(_tradingSymbol, exitPrice, entryPrice, quantity, _lotSize, Me._parentStrategy.StockType)
                If pl >= -100 Then
                    ret = ConvertFloorCeling(entryPrice - exitPrice, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetSignalCandleType(ByVal signalCandle As Payload, ByVal signalDirection As Trade.TradeExecutionDirection) As String
        Dim ret As String = ""
        If signalCandle.High > signalCandle.PreviousCandlePayload.High AndAlso
            signalCandle.Low > signalCandle.PreviousCandlePayload.Low Then
            If signalDirection = Trade.TradeExecutionDirection.Buy Then
                ret = "RHHLL"
            ElseIf signalDirection = Trade.TradeExecutionDirection.Sell Then
                ret = "HHLL"
            End If
        ElseIf signalCandle.High < signalCandle.PreviousCandlePayload.High AndAlso
            signalCandle.Low < signalCandle.PreviousCandlePayload.Low Then
            If signalDirection = Trade.TradeExecutionDirection.Buy Then
                ret = "HHLL"
            ElseIf signalDirection = Trade.TradeExecutionDirection.Sell Then
                ret = "RHHLL"
            End If
        ElseIf signalCandle.High <= signalCandle.PreviousCandlePayload.High AndAlso
            signalCandle.Low >= signalCandle.PreviousCandlePayload.Low Then
            ret = "Inside Bar"
        ElseIf signalCandle.High > signalCandle.PreviousCandlePayload.High AndAlso
            signalCandle.Low < signalCandle.PreviousCandlePayload.Low Then
            ret = "Outside Bar"
        End If

        Dim result As Decimal = signalCandle.PreviousCandlePayload.Volume * signalCandle.CandleRange / signalCandle.PreviousCandlePayload.CandleRange
        If result < signalCandle.Volume Then
            ret = String.Format("{0} - Good", ret)
        ElseIf result <= signalCandle.Volume * 90 / 100 Then
            ret = String.Format("{0} - Ok", ret)
        Else
            ret = String.Format("{0} - Poor", ret)
        End If
        Return ret
    End Function

End Class
