Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SmallestCandleBreakoutStrategyRule
    Inherits StrategyRule

    Private ReadOnly _firstSignalCandleTime As Date
    Private _firstEntryQuantity As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)

        _firstSignalCandleTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 16, 0)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrder(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        Dim parameter As PlaceOrderParameters = Nothing
        Dim setMTM As Boolean = False
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TradeType.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TradeType.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TradeType.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate >= _firstSignalCandleTime Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
            Dim signalCandle As Payload = Nothing
            'Dim previousLoss As Decimal = 0
            If lastExecutedTrade Is Nothing Then
                If currentMinuteCandlePayload.PreviousCandlePayload.CandleRange < currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleRange Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    setMTM = True
                    _firstEntryQuantity = _parentStrategy.CalculateQuantityFromInvestment(_lotSize, 25000, signalCandle.High, _parentStrategy.StockType, True)
                End If
            ElseIf lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                Dim slCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                If slCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate AndAlso
                    currentMinuteCandlePayload.PreviousCandlePayload.CandleRange < currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.CandleRange Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
                'previousLoss = Math.Abs(lastExecutedTrade.EntryPrice - lastExecutedTrade.ExitPrice)
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim middlePoint As Decimal = (signalCandle.High + signalCandle.Low) / 2
                Dim buyTriggerPrice As Decimal = middlePoint + ConvertFloorCeling((signalCandle.CandleRange * 30 / 100), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim sellTriggerPrice As Decimal = middlePoint - ConvertFloorCeling((signalCandle.CandleRange * 30 / 100), _parentStrategy.TickSize, RoundOfType.Celing)
                If currentTick.Open >= buyTriggerPrice Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    'Dim targetPoint As Decimal = ConvertFloorCeling((signalCandle.High + buffer) * 1 / 100, _parentStrategy.TickSize, RoundOfType.Celing) + previousLoss
                    Dim potentialStoploss As Decimal = (signalCandle.High + buffer) - ConvertFloorCeling((signalCandle.High + buffer) * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Celing)

                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.High + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _firstEntryQuantity,
                        .Target = .EntryPrice + 100,
                        .Stoploss = Math.Max(signalCandle.Low - buffer, potentialStoploss),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                    }
                ElseIf currentTick.Open <= sellTriggerPrice Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    'Dim targetPoint As Decimal = ConvertFloorCeling((signalCandle.Low - buffer) * 1 / 100, _parentStrategy.TickSize, RoundOfType.Celing) + previousLoss
                    Dim potentialStoploss As Decimal = (signalCandle.Low - buffer) + ConvertFloorCeling((signalCandle.Low - buffer) * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Celing)

                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.Low - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _firstEntryQuantity,
                        .Target = .EntryPrice - 100,
                        .Stoploss = Math.Min(signalCandle.High + buffer, potentialStoploss),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString
                    }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            If setmtm Then
                Dim pl As Decimal = Decimal.MaxValue
                Dim targetPoint As Decimal = ConvertFloorCeling(parameter.EntryPrice * 2 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                If parameter.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    pl = _parentStrategy.CalculatePL(Me._tradingSymbol, parameter.EntryPrice, parameter.EntryPrice + targetPoint, parameter.Quantity, Me._lotSize, _parentStrategy.StockType)
                ElseIf parameter.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    pl = _parentStrategy.CalculatePL(Me._tradingSymbol, parameter.EntryPrice - targetPoint, parameter.EntryPrice, parameter.Quantity, Me._lotSize, _parentStrategy.StockType)
                End If
                Me.MaxProfitOfThisStock = pl
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrder(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                If signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    Dim middlePoint As Decimal = (signalCandle.High + signalCandle.Low) / 2
                    Dim buyTriggerPrice As Decimal = middlePoint + ConvertFloorCeling((signalCandle.CandleRange * 30 / 100), _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim sellTriggerPrice As Decimal = middlePoint - ConvertFloorCeling((signalCandle.CandleRange * 30 / 100), _parentStrategy.TickSize, RoundOfType.Celing)
                    If currentTick.Open >= buyTriggerPrice Then
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            ret = New Tuple(Of Boolean, String)(True, "LTP near to opposite direction")
                        End If
                    ElseIf currentTick.Open <= sellTriggerPrice Then
                        If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            ret = New Tuple(Of Boolean, String)(True, "LTP near to opposite direction")
                        End If
                    End If
                ElseIf signalCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrder(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim breakoutCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If breakoutCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                signalCandle = breakoutCandle
            End If
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTrade.EntryPrice, RoundOfType.Floor)
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim potentialStoploss As Decimal = currentTrade.EntryPrice - ConvertFloorCeling(currentTrade.EntryPrice * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                triggerPrice = Math.Max(signalCandle.Low - buffer, potentialStoploss)
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim potentialStoploss As Decimal = currentTrade.EntryPrice + ConvertFloorCeling(currentTrade.EntryPrice * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                triggerPrice = Math.Min(signalCandle.High + buffer, potentialStoploss)
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move for breakout candle: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrder(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function
End Class