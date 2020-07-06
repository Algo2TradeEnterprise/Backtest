﻿Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class PreviousCandleSwingHighStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
    End Class
#End Region

    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date) = GetSignalForEntry(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim entryPrice As Decimal = signal.Item2
                Dim quantity As Integer = signal.Item3
                Dim iteration As Integer = signal.Item4
                Dim remarks As String = signal.Item5
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(entryPrice, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)

                If currentTick.Open < entryPrice Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 1000000000,
                                .Target = .EntryPrice + 1000000000,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = iteration,
                                .Supporting2 = remarks,
                                .Supporting3 = signal.Item6.ToString("dd-MMM-yyyy HH:mm:ss"),
                                .Supporting4 = signal.Item7.ToString("dd-MMM-yyyy HH:mm:ss")
                            }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date) = GetSignalForEntry(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso currentTrade.EntryPrice <> signal.Item2 Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        End If
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, candle.Open)
        Dim swing As Indicator.Swing = _swingPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim preSwing As Indicator.Swing = _swingPayload(swing.SwingHighTime)
        If swing.SwingHigh < preSwing.SwingHigh Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(swing.SwingHigh, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
            Dim entryPrice As Decimal = swing.SwingHigh + buffer
            If lastTrade Is Nothing Then
                ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date)(True, entryPrice, quantity, iteration, "First Trade", swing.SwingHighTime, preSwing.SwingHighTime)
            Else
                Dim lastSignalTime As Date = Date.ParseExact(lastTrade.Supporting3, "dd-MMM-yyyy HH:mm:ss", Nothing)
                If swing.SwingHighTime <> lastSignalTime Then
                    If entryPrice < lastTrade.EntryPrice Then
                        If lastTrade.EntryPrice - entryPrice >= atr Then
                            iteration = Val(lastTrade.Supporting1) + 1
                            quantity = GetQuantity(iteration, entryPrice)
                            ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date)(True, entryPrice, quantity, iteration, "Below last entry", swing.SwingHighTime, preSwing.SwingHighTime)
                        End If
                    Else
                        ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date, Date)(True, entryPrice, quantity, iteration, "(Reset) Above last entry", swing.SwingHighTime, preSwing.SwingHighTime)
                    End If
                End If
            End If
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
