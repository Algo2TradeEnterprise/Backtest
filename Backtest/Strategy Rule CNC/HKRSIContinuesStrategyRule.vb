Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKRSIContinuesStrategyRule
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
        Public BuyAtBelowRSI As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _hkPayloads As Dictionary(Of Date, Payload)
    Private _rsiPayloads As Dictionary(Of Date, Decimal)

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)
        Indicator.RSI.CalculateRSI(2, _hkPayloads, _rsiPayloads)
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
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Payload) = GetSignalForEntry(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim quantity As Integer = signal.Item3
                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = signal.Item2,
                                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                            .Quantity = quantity,
                                                            .Stoploss = .EntryPrice - 1000000000,
                                                            .Target = .EntryPrice + 1000000000,
                                                            .Buffer = 0,
                                                            .SignalCandle = signal.Item4,
                                                            .OrderType = Trade.TypeOfOrder.Market,
                                                            .Supporting1 = signal.Item4.PayloadDate.ToString("dd-MM-yyyy")
                                                        }

                If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                parameters.Add(parameter)
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

    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Payload) = Nothing
        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim currentDayHKPayload As Payload = _hkPayloads(currentTick.PayloadDate.Date)
        Dim initialQuantity As Integer = 1
        If currentDayPayload.High >= currentDayHKPayload.PreviousCandlePayload.High Then
            Dim rsiSatisfiedPayload As Payload = GetLastRSISatisfiedCandle(currentDayPayload.PayloadDate)
            If rsiSatisfiedPayload IsNot Nothing Then
                Dim lastOrder As Trade = GetLastOrder(currentDayPayload)
                If lastOrder Is Nothing OrElse lastOrder.SignalCandle.PayloadDate <> rsiSatisfiedPayload.PayloadDate Then
                    Dim potentialEntry As Decimal = ConvertFloorCeling(currentDayHKPayload.PreviousCandlePayload.High, _parentStrategy.TickSize, RoundOfType.Floor)
                    If lastOrder Is Nothing OrElse potentialEntry >= lastOrder.EntryPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Integer, Payload)(True, potentialEntry, initialQuantity, rsiSatisfiedPayload)
                    ElseIf lastOrder IsNot Nothing AndAlso potentialEntry < lastOrder.EntryPrice Then
                        Dim quantity As Integer = initialQuantity
                        Select Case _userInputs.QuantityType
                            Case TypeOfQuantity.AP
                                quantity = lastOrder.Quantity + initialQuantity
                            Case TypeOfQuantity.GP
                                quantity = lastOrder.Quantity * 2
                            Case TypeOfQuantity.Linear
                                quantity = initialQuantity
                        End Select
                        ret = New Tuple(Of Boolean, Decimal, Integer, Payload)(True, potentialEntry, quantity, rsiSatisfiedPayload)
                    Else
                        Throw New NotImplementedException
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastRSISatisfiedCandle(ByVal beforeThisTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _hkPayloads IsNot Nothing AndAlso _hkPayloads.Count > 0 Then
            For Each runningPayload In _hkPayloads.OrderByDescending(Function(x)
                                                                         Return x.Key
                                                                     End Function)
                If runningPayload.Key < beforeThisTime Then
                    Dim rsi As Decimal = _rsiPayloads(runningPayload.Key)
                    If rsi <= _userInputs.BuyAtBelowRSI Then
                        ret = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
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
