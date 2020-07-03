﻿Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class PreviousSwingLowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentDayCandlePayload IsNot Nothing AndAlso currentDayCandlePayload.PreviousCandlePayload IsNot Nothing Then
            Dim signals As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)) = GetSignalForEntry(currentDayCandlePayload)
            If signals IsNot Nothing AndAlso signals.Count > 0 Then
                For Each runningSignal In signals
                    Dim entryPrice As Decimal = runningSignal.Item2
                    Dim quantity As Integer = runningSignal.Item3
                    Dim iteration As Integer = runningSignal.Item4
                    Dim remarks As String = runningSignal.Item5

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = entryPrice,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = iteration,
                                                                .Supporting2 = remarks,
                                                                .Supporting3 = runningSignal.Item6.ToString("dd-MMM-yyyy HH:mm:ss")
                                                            }

                    If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                    parameters.Add(parameter)
                Next
            End If
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Private Function GetSignalForEntry(ByVal candle As Payload) As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
        Dim ret As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, candle.Open)
        Dim swing As Indicator.Swing = _swingPayload(candle.PreviousCandlePayload.PayloadDate)
        If candle.Low < swing.SwingLow Then
            Dim entryPrice As Decimal = swing.SwingLow
            If candle.Open < entryPrice Then entryPrice = candle.Open
            If lastTrade Is Nothing Then
                If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "First Trade", swing.SwingLowTime))
            Else
                Dim lastSignalTime As Date = Date.ParseExact(lastTrade.Supporting3, "dd-MMM-yyyy HH:mm:ss", Nothing)
                If swing.SwingLowTime <> lastSignalTime Then
                    If entryPrice < lastTrade.EntryPrice Then
                        If lastTrade.EntryPrice - entryPrice >= atr Then
                            iteration = Val(lastTrade.Supporting1) + 1
                            quantity = GetQuantity(iteration, entryPrice)
                            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "Below last entry", swing.SwingLowTime))
                        End If
                    Else
                        If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                        ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "(Reset) Above last entry", swing.SwingLowTime))
                    End If
                End If
            End If
            'If candle.Close < candle.PreviousCandlePayload.Low Then
            '    entryPrice = candle.Close
            '    iteration = 1
            '    quantity = GetQuantity(iteration, entryPrice)
            '    Dim lastEntryPrice As Decimal = Decimal.MinValue
            '    Dim lastIteration As Integer = Integer.MinValue
            '    If ret IsNot Nothing AndAlso ret.Count > 0 Then
            '        lastEntryPrice = ret.FirstOrDefault.Item2
            '        lastIteration = ret.FirstOrDefault.Item4
            '    ElseIf lastTrade IsNot Nothing Then
            '        lastEntryPrice = lastTrade.EntryPrice
            '        lastIteration = lastTrade.Supporting1
            '    End If
            '    If lastEntryPrice <> Decimal.MinValue AndAlso lastIteration <> Integer.MinValue Then
            '        If entryPrice < lastEntryPrice Then
            '            If lastEntryPrice - entryPrice >= atr Then
            '                iteration = lastIteration + 1
            '                quantity = GetQuantity(iteration, entryPrice)
            '                If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String))
            '                ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String)(True, entryPrice, quantity, iteration, "(Close) Below last entry"))
            '            End If
            '        Else
            '            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String))
            '            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String)(True, entryPrice, quantity, iteration, "(Close) (Reset) Above last entry"))
            '        End If
            '    End If
            'End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital * Math.Pow(2, iterationNumber - 1)
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function
End Class
