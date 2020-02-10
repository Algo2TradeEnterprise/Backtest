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
        MyBase.CompletePreProcessing()

        _weeklyPayloads = Common.ConvertDayPayloadsToWeek(_signalPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then

            If _userInputs.EntryType = TypeOfEntry.Both OrElse _userInputs.EntryType = TypeOfEntry.PriceDrop Then
                Dim initialQuantity As Integer = Integer.MinValue
                Select Case _userInputs.QuantityType
                    Case TypeOfQuantity.Forward
                        initialQuantity = 1
                    Case TypeOfQuantity.Reverse
                        initialQuantity = Math.Floor(Math.Abs(_userInputs.BuyTillPriceDropPercentage) / Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage))
                End Select
                Dim ctr As Integer = 0
                For i = Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) To Math.Abs(_userInputs.BuyTillPriceDropPercentage) Step Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage)
                    Dim signalCandle As Payload = Nothing
                    Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForDrop(currentTick, i)
                    If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                        signalCandle = signalReceivedForEntry.Item3

                        Dim quantity As Integer = initialQuantity
                        Select Case _userInputs.QuantityType
                            Case TypeOfQuantity.Forward
                                quantity = initialQuantity + ctr
                            Case TypeOfQuantity.Reverse
                                quantity = initialQuantity - ctr
                        End Select

                        If signalCandle IsNot Nothing Then
                            Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                        .EntryPrice = signalReceivedForEntry.Item2,
                                                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                        .Quantity = quantity,
                                                                        .Stoploss = .EntryPrice - 1000000000,
                                                                        .Target = .EntryPrice + 1000000000,
                                                                        .Buffer = 0,
                                                                        .SignalCandle = signalCandle,
                                                                        .OrderType = Trade.TypeOfOrder.Market,
                                                                        .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yy HH:mm:ss"),
                                                                        .Supporting2 = i * -1
                                                                    }

                            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                            parameters.Add(parameter)
                        End If
                    End If
                    ctr += 1
                Next
            End If

            If _userInputs.EntryType = TypeOfEntry.Both OrElse _userInputs.EntryType = TypeOfEntry.PriceUp Then
                Dim initialQuantity As Integer = Integer.MinValue
                Select Case _userInputs.QuantityType
                    Case TypeOfQuantity.Forward
                        initialQuantity = 1
                    Case TypeOfQuantity.Reverse
                        initialQuantity = Math.Floor(Math.Abs(_userInputs.BuyTillPriceUpPercentage) / Math.Abs(_userInputs.BuyAtEveryPriceUpPercentage))
                End Select
                Dim ctr As Integer = 0
                For i = Math.Abs(_userInputs.BuyAtEveryPriceUpPercentage) To Math.Abs(_userInputs.BuyTillPriceUpPercentage) Step Math.Abs(_userInputs.BuyAtEveryPriceUpPercentage)
                    Dim signalCandle As Payload = Nothing
                    Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForUp(currentTick, i)
                    If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                        signalCandle = signalReceivedForEntry.Item3

                        Dim quantity As Integer = initialQuantity
                        Select Case _userInputs.QuantityType
                            Case TypeOfQuantity.Forward
                                quantity = initialQuantity + ctr
                            Case TypeOfQuantity.Reverse
                                quantity = initialQuantity - ctr
                        End Select

                        If signalCandle IsNot Nothing Then
                            Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                        .EntryPrice = signalReceivedForEntry.Item2,
                                                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                        .Quantity = quantity,
                                                                        .Stoploss = .EntryPrice - 1000000000,
                                                                        .Target = .EntryPrice + 1000000000,
                                                                        .Buffer = 0,
                                                                        .SignalCandle = signalCandle,
                                                                        .OrderType = Trade.TypeOfOrder.Market,
                                                                        .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yy HH:mm:ss"),
                                                                        .Supporting2 = i
                                                                    }

                            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                            parameters.Add(parameter)
                        End If
                    End If
                    ctr += 1
                Next
            End If
        End If
        If parameters IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
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
                    CDec(lastTrade.Supporting2) <= drpPer * -1) Then
                    Dim potentialEntry As Decimal = ConvertFloorCeling(weeklyPayload.PreviousCandlePayload.Close * (100 - drpPer) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    If potentialEntry <= weeklyPayload.Open Then
                        ret = New Tuple(Of Boolean, Decimal, Payload)(True, potentialEntry, weeklyPayload.PreviousCandlePayload)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSignalForUp(ByVal currentTick As Payload, ByVal upPer As Decimal) As Tuple(Of Boolean, Decimal, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Payload) = Nothing
        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim weeklyPayload As Payload = _weeklyPayloads.Where(Function(x)
                                                                 Return x.Value.PayloadDate <= currentDayPayload.PayloadDate
                                                             End Function).LastOrDefault.Value
        If weeklyPayload.PreviousCandlePayload IsNot Nothing Then
            Dim highChange As Decimal = ((currentDayPayload.High / weeklyPayload.PreviousCandlePayload.Close) - 1) * 100
            If highChange >= upPer Then
                Dim lastTrade As Trade = GetLastOrder(currentDayPayload)
                If Not (lastTrade IsNot Nothing AndAlso lastTrade.SignalCandle.PayloadDate = weeklyPayload.PreviousCandlePayload.PayloadDate AndAlso
                    CDec(lastTrade.Supporting2) >= upPer) Then
                    Dim potentialEntry As Decimal = ConvertFloorCeling(weeklyPayload.PreviousCandlePayload.Close * (100 + upPer) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    If potentialEntry >= weeklyPayload.Open Then
                        ret = New Tuple(Of Boolean, Decimal, Payload)(True, potentialEntry, weeklyPayload.PreviousCandlePayload)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

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
