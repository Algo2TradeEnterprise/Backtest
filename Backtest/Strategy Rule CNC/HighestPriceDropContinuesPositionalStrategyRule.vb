Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HighestPriceDropContinuesPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Linear = 1
        AP
        GP
    End Enum

    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
        Public BuyAtEveryPriceDropPercentage As Decimal
    End Class
#End Region

    Private ReadOnly _highestPrice As Decimal

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal highestPrice As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _highestPrice = highestPrice
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            Dim dropPer As Decimal = ((currentDayPayload.Low / _highestPrice) - 1) * 100
            If dropPer <= Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) * -1 Then
                Dim initialQuantity As Integer = 1
                Dim ctr As Integer = 1
                For i = Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) To Math.Abs(dropPer) Step Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage)
                    Dim signalCandle As Payload = Nothing
                    Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload) = GetSignalForDrop(currentTick, i)
                    If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                        signalCandle = signalReceivedForEntry.Item3

                        Dim quantity As Integer = initialQuantity
                        Select Case _userInputs.QuantityType
                            Case TypeOfQuantity.AP
                                quantity = initialQuantity * ctr
                            Case TypeOfQuantity.GP
                                quantity = initialQuantity * Math.Pow(2, ctr - 1)
                            Case TypeOfQuantity.Linear
                                quantity = initialQuantity
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
                                                                        .Supporting1 = i * -1,
                                                                        .Supporting2 = _highestPrice
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
        Dim lowChange As Decimal = ((currentDayPayload.Low / _highestPrice) - 1) * 100
        If lowChange <= drpPer * -1 Then
            Dim lastTrade As Trade = GetLastOrder(currentDayPayload)
            Dim potentialEntry As Decimal = ConvertFloorCeling(_highestPrice * (100 - drpPer) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
            If potentialEntry <= currentDayPayload.Open Then
                ret = New Tuple(Of Boolean, Decimal, Payload)(True, potentialEntry, currentDayPayload)
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
