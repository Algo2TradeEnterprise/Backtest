Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PivotsPointsStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _pivotPayload As Dictionary(Of Date, PivotPoints)
    Private _atrPayload As Dictionary(Of Date, Decimal)

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
        _quantity = Me.LotSize
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.Pivots.CalculatePivots(_signalPayload, _pivotPayload)
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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = signalCandleSatisfied.Item2
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim buffer As Decimal = Me._parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                If _tradingSymbol.Contains("BANKNIFTY") Then buffer = 1

                Dim numberOfTrade As Integer = _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol)
                Dim quantity As Integer = _quantity * (numberOfTrade + 1)
                If Me._parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                    quantity = Me._parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.High + buffer, signalCandle.Low - buffer, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Me._parentStrategy.StockType)
                    If numberOfTrade = 1 Then
                        quantity = Me._parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.High + buffer, signalCandle.High + signalCandle.CandleRange + 2 * buffer, Math.Abs(_userInputs.MaxLossPerTrade) * 2, Me._parentStrategy.StockType)
                    ElseIf numberOfTrade = 2 Then
                        Dim usedQuantity As Integer = Me._parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.High + buffer, signalCandle.High + signalCandle.CandleRange + 2 * buffer, Math.Abs(_userInputs.MaxLossPerTrade) * 2, Me._parentStrategy.StockType)
                        Dim pl As Decimal = Me._parentStrategy.CalculatePL(_tradingSymbol, signalCandle.High + buffer, signalCandle.Low - buffer, usedQuantity, Me.LotSize, Me._parentStrategy.StockType)
                        quantity = Me._parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.High + buffer, signalCandle.High + signalCandle.CandleRange + 2 * buffer, Math.Abs(_userInputs.MaxLossPerTrade) + Math.Abs(pl), Me._parentStrategy.StockType)
                    End If
                End If

                If signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Buy Then
                    If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = signalCandle.High + buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = signalCandle.Low - buffer,
                            .Target = .EntryPrice + ((.EntryPrice - .Stoploss) * _userInputs.TargetMultiplier),
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                        }
                    End If
                End If
                If signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Sell Then
                    If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = signalCandle.Low - buffer,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = signalCandle.High + buffer,
                            .Target = .EntryPrice - ((.Stoploss - .EntryPrice) * _userInputs.TargetMultiplier),
                            .Buffer = buffer,
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandleSatisfied As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If Me._parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            ElseIf signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signalCandleSatisfied.Item2.PayloadDate Then
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            Dim lastExecutedTrade As Trade = Me._parentStrategy.GetLastExecutedTradeOfTheStock(candle, Me._parentStrategy.TradeType)
            If lastExecutedTrade Is Nothing Then
                Dim previousDayPayload As Payload = GetPreviousDay(candle)
                If previousDayPayload IsNot Nothing Then
                    Dim currentDayPivot As PivotPoints = _pivotPayload(candle.PayloadDate)
                    Dim previousDayPivot As PivotPoints = _pivotPayload(previousDayPayload.PayloadDate)
                    If currentDayPivot.Pivot > previousDayPivot.Pivot Then
                        If candle.High < candle.PreviousCandlePayload.High Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
                        End If
                        'Dim lowestHighPayload As Payload = Nothing
                        'For Each runningPayload In _signalPayload.Values
                        '    If runningPayload.PayloadDate.Date = _tradingDate.Date AndAlso
                        '        runningPayload.PayloadDate <= candle.PayloadDate Then
                        '        If lowestHighPayload Is Nothing Then
                        '            lowestHighPayload = runningPayload
                        '        Else
                        '            If runningPayload.High < lowestHighPayload.High Then lowestHighPayload = runningPayload
                        '        End If
                        '    End If
                        'Next
                        'If lowestHighPayload IsNot Nothing Then
                        '    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, lowestHighPayload, Trade.TradeExecutionDirection.Buy)
                        'End If
                    ElseIf currentDayPivot.Pivot < previousDayPivot.Pivot Then
                        If candle.Low > candle.PreviousCandlePayload.Low Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
                        End If
                        'Dim highestLowPayload As Payload = Nothing
                        'For Each runningPayload In _signalPayload.Values
                        '    If runningPayload.PayloadDate.Date = _tradingDate.Date AndAlso
                        '        runningPayload.PayloadDate <= candle.PayloadDate AndAlso
                        '        runningPayload.PayloadDate.Date = runningPayload.PreviousCandlePayload.PayloadDate.Date Then
                        '        If highestLowPayload Is Nothing Then
                        '            highestLowPayload = runningPayload
                        '        Else
                        '            If runningPayload.Low > highestLowPayload.Low Then highestLowPayload = runningPayload
                        '        End If
                        '    End If
                        'Next
                        'If highestLowPayload IsNot Nothing Then
                        '    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, highestLowPayload, Trade.TradeExecutionDirection.Sell)
                        'End If
                    End If
                End If
            Else
                Dim direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                If lastExecutedTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    direction = Trade.TradeExecutionDirection.Sell
                ElseIf lastExecutedTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    direction = Trade.TradeExecutionDirection.Buy
                End If
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, lastExecutedTrade.SignalCandle, direction)
            End If
        End If
        Return ret
    End Function

    Private Function GetPreviousDay(ByVal currentCandle As Payload) As Payload
        Dim ret As Payload = currentCandle.PreviousCandlePayload
        While True
            If currentCandle.PayloadDate.Date = ret.PayloadDate.Date Then
                Dim temp As Payload = ret.PreviousCandlePayload
                ret = temp
            Else
                Exit While
            End If
        End While
        If ret.PayloadDate.Date = currentCandle.PayloadDate.Date Then ret = Nothing
        Return ret
    End Function
End Class