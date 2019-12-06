﻿Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ATRPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Linear = 1
        AP
        GP
    End Enum
    Enum ExitType
        None = 1
        CompoundingToNextEntry
        CompoundingToMonthlyATR
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
        Public QuntityForLinear As Integer
        Public TypeOfExit As ExitType
        Public TargetMultiplier As Decimal
        Public EntryATRMultiplier As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _stockPrice As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockPrice As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _stockPrice = stockPrice
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim monthlyPayload As Dictionary(Of Date, Payload) = Common.ConvertDayPayloadsToMonth(_signalPayload)
        Indicator.ATR.CalculateATR(14, monthlyPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim signalCandle As Payload = Nothing

            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload, Decimal) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                signalCandle = signalReceivedForEntry.Item3

                If signalCandle IsNot Nothing Then
                    Dim highestEntryPrice As Decimal = Decimal.MinValue
                    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                    If lastExecutedTrade IsNot Nothing Then highestEntryPrice = lastExecutedTrade.Supporting1
                    Dim quantity As Integer = 1
                    If lastExecutedTrade IsNot Nothing AndAlso
                        lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                        highestEntryPrice > signalReceivedForEntry.Item2 Then
                        Select Case _userInputs.QuantityType
                            Case TypeOfQuantity.AP
                                quantity = lastExecutedTrade.Quantity + 1
                            Case TypeOfQuantity.GP
                                quantity = lastExecutedTrade.Quantity * 2
                            Case TypeOfQuantity.Linear
                                quantity = _userInputs.QuntityForLinear
                        End Select
                    End If
                    highestEntryPrice = Math.Max(highestEntryPrice, signalReceivedForEntry.Item2)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalReceivedForEntry.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 1000000,
                                .Target = .EntryPrice + 1000000,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = highestEntryPrice,
                                .Supporting3 = signalReceivedForEntry.Item4
                            }

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
                    totalCapitalUsedWithoutMargin += parameter.EntryPrice * parameter.Quantity
                    totalQuantity += parameter.Quantity
                    Dim averageTradePrice As Decimal = totalCapitalUsedWithoutMargin / totalQuantity
                    If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                        For Each runningTrade In openActiveTrades
                            runningTrade.UpdateTrade(Supporting2:=ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor))
                        Next
                    End If
                    parameter.Supporting2 = ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _userInputs.TypeOfExit = ExitType.CompoundingToNextEntry Then
                Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload, Decimal) = GetSignalForEntry(currentTick)
                If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                    If signalReceivedForEntry.Item3 IsNot Nothing Then
                        Dim highestEntryPrice As Decimal = Decimal.MinValue
                        Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                        If lastExecutedTrade IsNot Nothing Then highestEntryPrice = lastExecutedTrade.Supporting1
                        Dim quantity As Integer = 1
                        If highestEntryPrice > signalReceivedForEntry.Item2 Then
                            Select Case _userInputs.QuantityType
                                Case TypeOfQuantity.AP
                                    quantity = lastExecutedTrade.Quantity + 1
                                Case TypeOfQuantity.GP
                                    quantity = lastExecutedTrade.Quantity * 2
                                Case TypeOfQuantity.Linear
                                    quantity = _userInputs.QuntityForLinear
                            End Select
                        End If
                        If quantity = 1 Then
                            ret = New Tuple(Of Boolean, Decimal, String)(True, signalReceivedForEntry.Item2, "Compounding Exit on Next Entry")
                        End If
                    End If
                End If
            ElseIf _userInputs.TypeOfExit = ExitType.CompoundingToMonthlyATR Then
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                'Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                'Dim atr As Decimal = ConvertFloorCeling(_atrPayload(previousMonth) * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                Dim atr As Decimal = currentTrade.Supporting3
                Dim averagePrice As Decimal = currentTrade.Supporting2
                If currentDayPayload.High >= averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                    Dim price As Decimal = averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, String)(True, price, "Compunding Exit on Monthly ATR")
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

#Region "Entry Rule"
    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Decimal) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            If currentDayPayload.PreviousCandlePayload IsNot Nothing Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                If lastExecutedTrade Is Nothing Then
                    Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                    Dim atr As Decimal = ConvertFloorCeling(_atrPayload(previousMonth), Me._parentStrategy.TickSize, RoundOfType.Floor)
                    Dim entryPrice As Decimal = _stockPrice - ConvertFloorCeling(atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    If currentDayPayload.Low <= entryPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Decimal)(True, entryPrice, currentDayPayload, atr)
                    End If
                Else
                    Dim atr As Decimal = lastExecutedTrade.Supporting3
                    Dim entryPrice As Decimal = lastExecutedTrade.EntryPrice - ConvertFloorCeling(atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                        Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                        atr = ConvertFloorCeling(_atrPayload(previousMonth), Me._parentStrategy.TickSize, RoundOfType.Floor)
                        entryPrice = lastExecutedTrade.ExitPrice - ConvertFloorCeling(atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    If currentDayPayload.Low <= entryPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Decimal)(True, entryPrice, currentDayPayload, atr)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
#End Region

End Class
