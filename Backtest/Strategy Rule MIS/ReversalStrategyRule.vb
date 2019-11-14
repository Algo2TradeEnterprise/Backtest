Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ReversalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public NumberOfTradeOnNewSignal As Integer
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _highSignalCandle As Payload = Nothing
    Private _lowSignalCandle As Payload = Nothing
    Private _confirmationCandle As Payload = Nothing
    Private _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
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

        If currentMinuteCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing Then
            If _highSignalCandle IsNot Nothing Then
                Dim highestPoint As Decimal = _highSignalCandle.High
                If _confirmationCandle IsNot Nothing Then highestPoint = Math.Max(highestPoint, _confirmationCandle.High)
                If currentMinuteCandlePayload.PreviousCandlePayload.Close > highestPoint Then
                    _highSignalCandle = Nothing
                End If
            End If
            If _lowSignalCandle IsNot Nothing Then
                Dim lowestPoint As Decimal = _lowSignalCandle.Low
                If _confirmationCandle IsNot Nothing Then lowestPoint = Math.Min(lowestPoint, _confirmationCandle.Low)
                If currentMinuteCandlePayload.PreviousCandlePayload.Close < lowestPoint Then
                    _lowSignalCandle = Nothing
                End If
            End If
        End If

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) AndAlso
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
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedTrade Is Nothing Then
                    signalCandle = _confirmationCandle
                Else
                    If lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                        If lastExecutedTrade.SignalCandle.PayloadDate <> _confirmationCandle.PayloadDate Then
                            If Me._parentStrategy.GetNumberOfUniqueSignalTradeOfTheStock(currentMinuteCandlePayload, Me._parentStrategy.TradeType) < _userInputs.NumberOfTradeOnNewSignal Then
                                signalCandle = _confirmationCandle
                            End If
                        Else
                            signalCandle = _confirmationCandle
                        End If
                    End If
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim targetMul As Decimal = _userInputs.TargetMultiplier
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2 + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = LotSize,
                                .Stoploss = signalCandleSatisfied.Item3 - buffer,
                                .Target = .EntryPrice + (.EntryPrice - .Stoploss) * targetMul,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _lowSignalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting3 = _dayATR
                            }
                    Dim slPoint As Decimal = parameter.EntryPrice - parameter.Stoploss
                    If _dayATR / slPoint > 10 Then targetMul = _userInputs.TargetMultiplier + 1
                    parameter.Target = parameter.EntryPrice + slPoint * targetMul
                    parameter.Supporting4 = _dayATR / slPoint
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim targetMul As Decimal = _userInputs.TargetMultiplier
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandleSatisfied.Item2 - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = LotSize,
                                .Stoploss = signalCandleSatisfied.Item3 + buffer,
                                .Target = .EntryPrice - (.Stoploss - .EntryPrice) * targetMul,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _highSignalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting3 = _dayATR
                            }
                    Dim slPoint As Decimal = parameter.Stoploss - parameter.EntryPrice
                    If _dayATR / slPoint > 10 Then targetMul = _userInputs.TargetMultiplier + 1
                    parameter.Target = parameter.EntryPrice - slPoint * targetMul
                    parameter.Supporting4 = _dayATR / slPoint
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Me.MaxProfitOfThisStock = (parameter.EntryPrice * parameter.Quantity / Me._parentStrategy.MarginMultiplier) * _parentStrategy.StockMaxProfitPercentagePerDay / 100
            End If
            If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
                Me.MaxLossOfThisStock = (parameter.EntryPrice * parameter.Quantity / Me._parentStrategy.MarginMultiplier) * _parentStrategy.StockMaxLossPercentagePerDay / 100
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentMinuteCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing Then
            If _highSignalCandle IsNot Nothing Then
                Dim highestPoint As Decimal = _highSignalCandle.High
                If _confirmationCandle IsNot Nothing Then highestPoint = Math.Max(highestPoint, _confirmationCandle.High)
                If currentMinuteCandlePayload.PreviousCandlePayload.Close > highestPoint Then
                    _highSignalCandle = Nothing
                End If
            End If
            If _lowSignalCandle IsNot Nothing Then
                Dim lowestPoint As Decimal = _lowSignalCandle.Low
                If _confirmationCandle IsNot Nothing Then lowestPoint = Math.Min(lowestPoint, _confirmationCandle.Low)
                If currentMinuteCandlePayload.PreviousCandlePayload.Close < lowestPoint Then
                    _lowSignalCandle = Nothing
                End If
            End If
        End If

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If _confirmationCandle IsNot Nothing Then
                    If currentTrade.SignalCandle.PayloadDate <> _confirmationCandle.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                Else
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim potentialTargetMultiplier As Decimal = 1.5
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim targetPoint As Decimal = currentTrade.PotentialTarget - currentTrade.EntryPrice
                Dim slPoint As Decimal = targetPoint / _userInputs.TargetMultiplier
                If _dayATR / slPoint > 10 Then potentialTargetMultiplier = 2
                If currentTrade.MaximumDrawUp - currentTrade.EntryPrice >= slPoint * potentialTargetMultiplier Then
                    triggerPrice = currentTrade.EntryPrice + Me._parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, LotSize, Me._parentStrategy.StockType)
                    _lowSignalCandle = Nothing
                    _confirmationCandle = Nothing
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim targetPoint As Decimal = currentTrade.EntryPrice - currentTrade.PotentialTarget
                Dim slPoint As Decimal = targetPoint / _userInputs.TargetMultiplier
                If _dayATR / slPoint > 10 Then potentialTargetMultiplier = 2
                If currentTrade.EntryPrice - currentTrade.MaximumDrawUp >= slPoint * potentialTargetMultiplier Then
                    triggerPrice = currentTrade.EntryPrice - Me._parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Sell, LotSize, Me._parentStrategy.StockType)
                    _highSignalCandle = Nothing
                    _confirmationCandle = Nothing
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move to breakeven: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
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
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _highSignalCandle Is Nothing AndAlso _confirmationCandle IsNot Nothing Then
                If (_confirmationCandle.CandleColor = Color.Red AndAlso _confirmationCandle.Supporting Is Nothing) OrElse
                    (_confirmationCandle.Supporting IsNot Nothing AndAlso _confirmationCandle.Supporting = Color.Red.ToString) Then
                    _confirmationCandle = Nothing
                End If
            End If
            If _lowSignalCandle Is Nothing AndAlso _confirmationCandle IsNot Nothing Then
                If (_confirmationCandle.CandleColor = Color.Green AndAlso _confirmationCandle.Supporting Is Nothing) OrElse
                    (_confirmationCandle.Supporting IsNot Nothing AndAlso _confirmationCandle.Supporting = Color.Green.ToString) Then
                    _confirmationCandle = Nothing
                End If
            End If

            Dim dayHighestCandle As Payload = GetHighestCandleOfTheDay(candle.PayloadDate)
            Dim dayLowestCandle As Payload = GetLowestCandleOfTheDay(candle.PayloadDate)
            If dayHighestCandle.PayloadDate = candle.PreviousCandlePayload.PayloadDate AndAlso
                candle.PreviousCandlePayload.CandleColor = Color.Green Then
                _highSignalCandle = candle.PreviousCandlePayload
            End If
            If dayLowestCandle.PayloadDate = candle.PreviousCandlePayload.PayloadDate AndAlso
                candle.PreviousCandlePayload.CandleColor = Color.Red Then
                _lowSignalCandle = candle.PreviousCandlePayload
            End If

            If _highSignalCandle IsNot Nothing AndAlso
                _highSignalCandle.CandleRange < _ATRPayload(_highSignalCandle.PayloadDate) * 2 Then
                If candle.CandleColor = Color.Red AndAlso IsStrongClose(candle) AndAlso
                    ((candle.CandleRange >= _highSignalCandle.CandleRange AndAlso candle.Volume < _highSignalCandle.Volume) OrElse
                    (candle.CandleRange < _highSignalCandle.CandleRange AndAlso candle.Volume <= _highSignalCandle.Volume / 2)) Then
                    _confirmationCandle = candle
                Else
                    If _confirmationCandle Is Nothing AndAlso
                        Not IsInsideBar(_highSignalCandle, candle) Then
                        _highSignalCandle = Nothing
                    End If
                End If
            End If
            If _lowSignalCandle IsNot Nothing AndAlso
                _lowSignalCandle.CandleRange < _ATRPayload(_lowSignalCandle.PayloadDate) * 2 Then
                If candle.CandleColor = Color.Green AndAlso IsStrongClose(candle) AndAlso
                    ((candle.CandleRange >= _lowSignalCandle.CandleRange AndAlso candle.Volume < _lowSignalCandle.Volume) OrElse
                    (candle.CandleRange < _lowSignalCandle.CandleRange AndAlso candle.Volume <= _lowSignalCandle.Volume / 2)) Then
                    _confirmationCandle = candle
                Else
                    If _confirmationCandle Is Nothing AndAlso
                        Not IsInsideBar(_lowSignalCandle, candle) Then
                        _lowSignalCandle = Nothing
                    End If
                End If
            End If

            If _confirmationCandle IsNot Nothing Then
                If (_confirmationCandle.CandleColor = Color.Green AndAlso _confirmationCandle.Supporting Is Nothing) OrElse
                    (_confirmationCandle.Supporting IsNot Nothing AndAlso _confirmationCandle.Supporting = Color.Green.ToString) Then
                    Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(_confirmationCandle.High, RoundOfType.Floor)
                    'If (_confirmationCandle.High - _lowSignalCandle.Low) + 2 * buffer > _dayATR / 2 Then
                    If candle.High < _confirmationCandle.High AndAlso candle.Low > Math.Min(_lowSignalCandle.Low, _confirmationCandle.Low) Then
                        _confirmationCandle = candle
                        _confirmationCandle.Supporting = Color.Green.ToString
                    End If
                    'End If
                    If (_confirmationCandle.High - _lowSignalCandle.Low) + 2 * buffer <= _dayATR / 2 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _confirmationCandle.High, Math.Min(_lowSignalCandle.Low, _confirmationCandle.Low), Trade.TradeExecutionDirection.Buy)
                    End If
                ElseIf (_confirmationCandle.CandleColor = Color.Red AndAlso _confirmationCandle.Supporting Is Nothing) OrElse
                    (_confirmationCandle.Supporting IsNot Nothing AndAlso _confirmationCandle.Supporting = Color.Red.ToString) Then
                    Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(_confirmationCandle.Low, RoundOfType.Floor)
                    'If (_highSignalCandle.High - _confirmationCandle.Low) + 2 * buffer > _dayATR / 2 Then
                    If candle.Low > _confirmationCandle.Low AndAlso candle.High < Math.Max(_highSignalCandle.High, _confirmationCandle.High) Then
                        _confirmationCandle = candle
                        _confirmationCandle.Supporting = Color.Red.ToString
                    End If
                    'End If
                    If (_highSignalCandle.High - _confirmationCandle.Low) + 2 * buffer <= _dayATR / 2 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _confirmationCandle.Low, Math.Max(_highSignalCandle.High, _confirmationCandle.High), Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsInsideBar(ByVal signalCandle As Payload, ByVal currentCandle As Payload) As Boolean
        Dim ret As Boolean = False
        If currentCandle.CandleColor = Color.Green Then
            If currentCandle.Open >= signalCandle.Low AndAlso currentCandle.Close <= signalCandle.High Then
                ret = True
            End If
        Else
            If currentCandle.Open <= signalCandle.High AndAlso currentCandle.Close >= signalCandle.Low Then
                ret = True
            End If
        End If
        Return ret
    End Function

    Private Function IsStrongClose(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        If candle IsNot Nothing Then
            If candle.CandleColor = Color.Green Then
                If candle.Close >= candle.Low + candle.CandleRange * 75 / 100 Then
                    ret = True
                End If
            ElseIf candle.CandleColor = Color.Red Then
                If candle.Close <= candle.High - candle.CandleRange * 75 / 100 Then
                    ret = True
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestCandleOfTheDay(ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = currentTime.Date AndAlso runningPayload.Key < currentTime Then
                If ret Is Nothing Then
                    ret = runningPayload.Value
                Else
                    If runningPayload.Value.High >= ret.High Then
                        ret = runningPayload.Value
                    End If
                End If
            End If
        Next
        Return ret
    End Function

    Private Function GetLowestCandleOfTheDay(ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = currentTime.Date AndAlso runningPayload.Key < currentTime Then
                If ret Is Nothing Then
                    ret = runningPayload.Value
                Else
                    If runningPayload.Value.Low <= ret.Low Then
                        ret = runningPayload.Value
                    End If
                End If
            End If
        Next
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function
End Class
