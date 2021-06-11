Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class ORBStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPercentage As Decimal
        Public CapitalForEachTrade As Decimal
    End Class
#End Region

    Private _firstCandleOfTheDay As Payload = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
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

        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload.Values
                If runningPayload.PayloadDate.Date = _tradingDate.Date AndAlso runningPayload.PreviousCandlePayload IsNot Nothing AndAlso
                    runningPayload.PayloadDate.Date <> runningPayload.PreviousCandlePayload.PayloadDate.Date Then
                    _firstCandleOfTheDay = runningPayload
                    If _firstCandleOfTheDay.Open > _userInputs.CapitalForEachTrade Then Me.EligibleToTakeTrade = False
                    Exit For
                End If
            Next
        End If
        If _firstCandleOfTheDay Is Nothing Then Me.EligibleToTakeTrade = False
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso Me.EligibleToTakeTrade AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsTradeActive(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item2
            End If
            If signalCandle IsNot Nothing Then
                Dim quantity As Decimal = Math.Floor(_userInputs.CapitalForEachTrade / signal.Item3)
                If quantity > 0 Then
                    Dim slPoint As Decimal = ConvertFloorCeling(signal.Item3 * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = signal.Item3,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - slPoint,
                            .Target = .EntryPrice + 10000000,
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = slPoint
                        }
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = signal.Item3,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + slPoint,
                            .Target = .EntryPrice - 10000000,
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = slPoint
                        }
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing And currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso
            ((signal.Item2.PayloadDate <> currentTrade.SignalCandle.PayloadDate) OrElse (signal.Item4 <> currentTrade.EntryDirection)) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting2
            Dim potentialSL As Decimal = Decimal.MinValue
            Dim reason As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                Dim mul As Integer = Math.Floor(plPoint / slPoint)
                If currentTrade.EntryPrice + slPoint * mul > currentTrade.PotentialStopLoss Then
                    potentialSL = currentTrade.EntryPrice + slPoint * mul
                    reason = String.Format("Move {0}", mul)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                Dim mul As Integer = Math.Floor(plPoint / slPoint)
                If currentTrade.EntryPrice - slPoint * mul < currentTrade.PotentialStopLoss Then
                    potentialSL = currentTrade.EntryPrice - slPoint * mul
                    reason = String.Format("Move {0}", mul)
                End If
            End If
            If potentialSL <> Decimal.MinValue AndAlso potentialSL <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, potentialSL, reason)
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

    Private Function GetSignalForEntry(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection) = Nothing
        Dim signalCandle As Payload = Nothing
        Dim potentialBuyEntryPrice As Decimal = Decimal.MaxValue
        Dim potentialSellEntryPrice As Decimal = Decimal.MinValue
        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
        If lastExecutedOrder Is Nothing Then
            If currentCandle.PayloadDate = _firstCandleOfTheDay.PayloadDate Then
                potentialBuyEntryPrice = ConvertFloorCeling(_firstCandleOfTheDay.Open + _firstCandleOfTheDay.Open * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                potentialSellEntryPrice = ConvertFloorCeling(_firstCandleOfTheDay.Open - _firstCandleOfTheDay.Open * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                signalCandle = currentCandle.PreviousCandlePayload
            Else
                potentialBuyEntryPrice = _firstCandleOfTheDay.High
                potentialSellEntryPrice = _firstCandleOfTheDay.Low
                signalCandle = _firstCandleOfTheDay
            End If
        Else
            If lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                Dim exitPrice As Decimal = lastExecutedOrder.ExitPrice
                potentialBuyEntryPrice = ConvertFloorCeling(exitPrice + exitPrice * 0.25 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                potentialSellEntryPrice = ConvertFloorCeling(exitPrice - exitPrice * 0.25 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                signalCandle = currentCandle.PreviousCandlePayload
            End If
        End If
        If signalCandle IsNot Nothing AndAlso potentialBuyEntryPrice <> Decimal.MaxValue AndAlso potentialSellEntryPrice <> Decimal.MinValue Then
            Dim middle As Decimal = (potentialBuyEntryPrice + potentialSellEntryPrice) / 2
            Dim range As Decimal = potentialBuyEntryPrice - middle
            If currentTick.Open >= middle + range * 70 / 100 Then
                ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection)(True, signalCandle, potentialBuyEntryPrice, Trade.TradeExecutionDirection.Buy)
            ElseIf currentTick.Open <= middle - range * 70 / 100 Then
                ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TradeExecutionDirection)(True, signalCandle, potentialSellEntryPrice, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function
End Class