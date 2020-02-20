Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AveragePriceDropContinuesStrategyRule
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

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _highestPrice As Decimal

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
            For runningTick As Decimal = currentDayPayload.Open To currentDayPayload.Low Step _parentStrategy.TickSize
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Decimal, Integer, Payload, Integer) = GetSignalForDrop(currentTick, runningTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item4
                    Dim quantity As Integer = signal.Item3
                    If signalCandle IsNot Nothing Then
                        Dim totalCapitalUsedWithoutMargin As Decimal = 0
                        Dim totalQuantity As Decimal = 0
                        Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                        If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                            For Each runningTrade In openActiveTrades
                                If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                    totalCapitalUsedWithoutMargin += runningTrade.EntryPrice * runningTrade.Quantity
                                    totalQuantity += runningTrade.Quantity
                                End If
                            Next
                        End If
                        totalCapitalUsedWithoutMargin += signal.Item2 * quantity
                        totalQuantity += quantity
                        Dim averageTradePrice As Decimal = totalCapitalUsedWithoutMargin / totalQuantity
                        If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                            For Each runningTrade In openActiveTrades
                                runningTrade.UpdateTrade(Supporting1:=ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor))
                            Next
                        End If

                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signal.Item2,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = signalCandle,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor),
                                                                .Supporting2 = signal.Item5,
                                                                .Supporting3 = _highestPrice
                                                            }

                        If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                        parameters.Add(parameter)
                    End If
                End If
            Next
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

    Private Function GetSignalForDrop(ByVal currentTick As Payload, ByVal runningTick As Decimal) As Tuple(Of Boolean, Decimal, Integer, Payload, Integer)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Payload, Integer) = Nothing
        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim initialQuantity As Integer = 1
        Dim lastTrade As Trade = GetLastOrder(currentDayPayload)
        If lastTrade IsNot Nothing Then
            Dim averagePrice As Decimal = lastTrade.Supporting1
            Dim changePer As Decimal = ((runningTick / averagePrice) - 1) * 100
            If changePer <= Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) * -1 Then
                Dim potentialEntry As Decimal = ConvertFloorCeling(averagePrice * (100 - Math.Floor(Math.Abs(changePer))) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                If CInt(lastTrade.Supporting2) = 0 Then
                    ret = New Tuple(Of Boolean, Decimal, Integer, Payload, Integer)(True, potentialEntry, initialQuantity, currentDayPayload, Math.Floor(Math.Abs(changePer)))
                Else
                    Dim quantity As Integer = initialQuantity
                    Select Case _userInputs.QuantityType
                        Case TypeOfQuantity.AP
                            quantity = lastTrade.Quantity + initialQuantity
                        Case TypeOfQuantity.GP
                            quantity = lastTrade.Quantity * 2
                        Case TypeOfQuantity.Linear
                            quantity = initialQuantity
                    End Select
                    ret = New Tuple(Of Boolean, Decimal, Integer, Payload, Integer)(True, potentialEntry, quantity, currentDayPayload, Math.Floor(Math.Abs(changePer)))
                End If
            ElseIf runningTick > averagePrice Then
                Dim drpPer As Decimal = ((runningTick / _highestPrice) - 1) * 100
                If drpPer <= Math.Abs(_userInputs.BuyAtEveryPriceDropPercentage) * -1 Then
                    Dim potentialEntry As Decimal = ConvertFloorCeling(_highestPrice * (100 - Math.Floor(Math.Abs(drpPer))) / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, Integer, Payload, Integer)(True, potentialEntry, initialQuantity, currentDayPayload, 0)
                End If
            End If
        Else
            ret = New Tuple(Of Boolean, Decimal, Integer, Payload, Integer)(True, currentDayPayload.Open, initialQuantity, currentDayPayload, 0)
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
