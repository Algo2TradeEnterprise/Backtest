Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class RetracementSellOnlyStrategy
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public Capital As Decimal
        Public SignalStartTime As Date
        Public SignalEndTime As Date
        Public CandleBobyPercentage As Decimal
        Public StoplossTrailingPercentage As Decimal
    End Class
#End Region

    Private ReadOnly _signalStartTime As Date
    Private ReadOnly _signalEndTime As Date
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

        _signalStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _userInputs.SignalStartTime.Hour, _userInputs.SignalStartTime.Minute, _userInputs.SignalStartTime.Second)
        _signalEndTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _userInputs.SignalEndTime.Hour, _userInputs.SignalEndTime.Minute, _userInputs.SignalEndTime.Second)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso Not _tradingDate.Date.DayOfWeek = DayOfWeek.Monday Then
            Dim signalCandle As Payload = Nothing
            If currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate >= _signalStartTime AndAlso
                currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate <= _signalEndTime Then
                If currentMinuteCandlePayload.PreviousCandlePayload.CandleColor = Color.Red AndAlso
                    (currentMinuteCandlePayload.PreviousCandlePayload.CandleBody / currentMinuteCandlePayload.PreviousCandlePayload.Close) * 100 >= _userInputs.CandleBobyPercentage Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)
                Dim entryPrice As Decimal = currentTick.Open
                Dim stoploss As Decimal = signalCandle.High + buffer
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(1, _userInputs.Capital, entryPrice, Trade.TypeOfStock.Cash, False)
                Dim modifiedQuantity As Integer = Math.Ceiling(quantity / 2)

                parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = modifiedQuantity,
                                .Stoploss = stoploss,
                                .Target = entryPrice - 100000000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = (currentMinuteCandlePayload.PreviousCandlePayload.CandleBody / currentMinuteCandlePayload.PreviousCandlePayload.Close) * 100,
                                .Supporting3 = "Order 1"
                            }

                If (quantity - modifiedQuantity) > 0 Then
                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity - modifiedQuantity,
                                .Stoploss = stoploss,
                                .Target = entryPrice - 100000000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = (currentMinuteCandlePayload.PreviousCandlePayload.CandleBody / currentMinuteCandlePayload.PreviousCandlePayload.Close) * 100,
                                .Supporting3 = "Order 2"
                            }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then parameters = New List(Of PlaceOrderParameters) From {parameter1}
        If parameter2 IsNot Nothing Then parameters.Add(parameter2)
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting3.ToUpper = "ORDER 1" Then
            Dim exitTrade As Boolean = False
            Dim reason As String = ""
            Dim entryPrice As Decimal = currentTrade.EntryPrice
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim gainLoss As Decimal = ((currentTick.Open - entryPrice) / entryPrice) * 100
                If gainLoss >= _userInputs.StoplossTrailingPercentage Then
                    exitTrade = True
                    reason = String.Format("Gain:{0}%. So hard close.", Math.Round(gainLoss, 2))
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim gainLoss As Decimal = ((entryPrice - currentTick.Open) / entryPrice) * 100
                If gainLoss >= _userInputs.StoplossTrailingPercentage Then
                    exitTrade = True
                    reason = String.Format("Gain:{0}%. So hard close.", Math.Round(gainLoss, 2))
                End If
            End If
            If exitTrade Then
                ret = New Tuple(Of Boolean, String)(True, reason)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim reason As String = ""
            Dim entryPrice As Decimal = currentTrade.EntryPrice
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim gainLoss As Decimal = ((currentTick.Open - entryPrice) / entryPrice) * 100
                Dim multiplier As Decimal = Math.Floor(gainLoss / _userInputs.StoplossTrailingPercentage)
                If multiplier > 1 Then
                    Dim stoploss As Decimal = ConvertFloorCeling(entryPrice + entryPrice * 0.1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    triggerPrice = stoploss + ConvertFloorCeling(stoploss * _userInputs.StoplossTrailingPercentage / 100 * (multiplier - 1), _parentStrategy.TickSize, RoundOfType.Floor)
                    reason = String.Format("Gain:{0}%, So move to {1}%", Math.Round(gainLoss, 2), _userInputs.StoplossTrailingPercentage * (multiplier - 1))
                ElseIf multiplier > 0 Then
                    triggerPrice = ConvertFloorCeling(entryPrice + entryPrice * 0.1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    reason = String.Format("Gain:{0}%, So cost to cost movement", Math.Round(gainLoss, 2))
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <= currentTrade.PotentialStopLoss Then
                    triggerPrice = Decimal.MinValue
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim gainLoss As Decimal = ((entryPrice - currentTick.Open) / entryPrice) * 100
                Dim multiplier As Decimal = Math.Floor(gainLoss / _userInputs.StoplossTrailingPercentage)
                If multiplier > 1 Then
                    Dim stoploss As Decimal = ConvertFloorCeling(entryPrice - entryPrice * 0.1 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    triggerPrice = stoploss - ConvertFloorCeling(stoploss * _userInputs.StoplossTrailingPercentage / 100 * (multiplier - 1), _parentStrategy.TickSize, RoundOfType.Floor)
                    reason = String.Format("Gain:{0}%, So move to {1}%", Math.Round(gainLoss, 2), _userInputs.StoplossTrailingPercentage * (multiplier - 1))
                ElseIf multiplier > 0 Then
                    triggerPrice = ConvertFloorCeling(entryPrice - entryPrice * 0.1 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    reason = String.Format("Gain:{0}%, So cost to cost movement", Math.Round(gainLoss, 2))
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice >= currentTrade.PotentialStopLoss Then
                    triggerPrice = Decimal.MinValue
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function
End Class