Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PriceDropContinuesPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Forward = 1
        Reverse
    End Enum

    Enum TypeOfEntry
        PriceDrop = 1
        PriceUp
        Both
    End Enum

    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
        Public EntryType As TypeOfEntry
        Public BuyAtEveryPriceDropPercentage As Decimal
        Public BuyTillPriceDropPercentage As Decimal
        Public BuyAtEveryPriceUpPercentage As Decimal
        Public BuyTillPriceUpPercentage As Decimal
    End Class
#End Region

    Private _weeklyPayloads As Dictionary(Of Date, Payload)

    Private ReadOnly _userInputs As StrategyRuleEntities
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
        _weeklyPayloads = Common.ConvertDayPayloadsToWeek(_signalPayload)

        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then

            If _userInputs.EntryType = TypeOfEntry.Both OrElse _userInputs.EntryType = TypeOfEntry.PriceDrop Then
                Dim lastQuantity As Integer = Integer.MinValue
                For i = Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) To Math.Abs(_userInputs.BuyTillPriceDropPercentage) Step Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage)
                    Dim parameter As PlaceOrderParameters = Nothing
                    Dim signalCandle As Payload = Nothing
                    Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForDrop(currentTick, i)
                    If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                        signalCandle = signalReceivedForEntry.Item3

                        If lastQuantity = Integer.MinValue Then
                            Select Case _userInputs.QuantityType
                                Case TypeOfQuantity.Forward
                                    lastQuantity = 1
                                Case TypeOfQuantity.Reverse
                                    lastQuantity = Math.Floor(Math.Abs(_userInputs.BuyTillPriceDropPercentage) / Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage))
                            End Select
                        Else
                            Select Case _userInputs.QuantityType
                                Case TypeOfQuantity.Forward
                                    lastQuantity = lastQuantity + 1
                                Case TypeOfQuantity.Reverse
                                    lastQuantity = lastQuantity - 1
                            End Select
                        End If
                        If lastQuantity = 0 Then Throw New ApplicationException("Check Quantity")
                        If signalCandle IsNot Nothing AndAlso lastQuantity <> 0 Then
                            parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signalReceivedForEntry.Item2,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = lastQuantity,
                                    .Stoploss = .EntryPrice - 1000000000,
                                    .Target = .EntryPrice + 1000000000,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yy HH:mm:ss"),
                                    .Supporting2 = i
                                }

                        End If
                    End If
                    If parameter IsNot Nothing Then
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                Next
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

#Region "Entry Rule"
    Private Function GetSignalForDrop(ByVal currentTick As Payload, ByVal drpPer As Decimal) As Tuple(Of Boolean, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Payload) = Nothing
        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim weeklyPayload As Payload = _weeklyPayloads.Where(Function(x)
                                                                 Return x.Value.PayloadDate <= currentDayPayload.PayloadDate
                                                             End Function).LastOrDefault.Value
        If weeklyPayload.PreviousCandlePayload IsNot Nothing Then
            Dim lowChange As Decimal = ((currentDayPayload.Low / weeklyPayload.PreviousCandlePayload.Close) - 1) * 100
            If lowChange <= drpPer * -1 Then
                Dim lastTrade As Trade = GetLastOrder(currentDayPayload)
                If Not (lastTrade IsNot Nothing AndAlso lastTrade.SignalCandle.PayloadDate = weeklyPayload.PreviousCandlePayload.PayloadDate AndAlso
                    CDec(lastTrade.Supporting2) >= drpPer) Then
                    Dim potentialEntry As Decimal = ConvertFloorCeling(weeklyPayload.PreviousCandlePayload.Close * (100 - drpPer) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    If potentialEntry <= weeklyPayload.Open Then
                        ret = New Tuple(Of Boolean, Decimal, Payload)(True, potentialEntry, weeklyPayload.PreviousCandlePayload)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
#End Region

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            Dim lastEntryOrder As Trade = Me._parentStrategy.GetLastEntryTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
            Dim lastClosedOrder As Trade = Me._parentStrategy.GetLastExitTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
            If lastEntryOrder IsNot Nothing Then ret = lastEntryOrder
            If lastClosedOrder IsNot Nothing AndAlso lastClosedOrder.ExitTime >= lastEntryOrder.EntryTime Then
                ret = lastClosedOrder
            End If
        End If
        Return ret
    End Function
End Class
