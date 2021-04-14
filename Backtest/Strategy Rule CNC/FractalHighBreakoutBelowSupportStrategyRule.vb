Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class FractalHighBreakoutBelowSupportStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Integer
        Public TargetMultiplier As Integer
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _pivotPayload As Dictionary(Of Date, PivotPoints) = Nothing

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

        Indicator.Pivots.CalculatePivots(_signalPayload, _pivotPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Decimal, Decimal) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.CNC)
                If lastTrade Is Nothing OrElse lastTrade.ExitTime < currentMinuteCandle.PayloadDate Then
                    signalCandle = signal.Item2
                End If
            End If
            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item3
                Dim stoploss As Decimal = signal.Item4
                Dim buffer As Decimal = signal.Item5
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)

                parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = stoploss,
                            .Target = target,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                            .Supporting2 = _fractalHighPayload(signalCandle.PayloadDate),
                            .Supporting3 = _fractalHighPayload(signalCandle.PayloadDate),
                            .Supporting4 = _pivotPayload(signalCandle.PayloadDate).Support3
                        }
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
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Decimal, Decimal) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item2.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If currentTick.Open >= currentTrade.PotentialTarget Then
                ret = New Tuple(Of Boolean, String)(True, "Target")
            ElseIf currentTick.Open <= currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, String)(True, "Stoploss")
            Else
                Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If currentMinuteCandle.PreviousCandlePayload.Close >= _pivotPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate).Resistance1 Then
                    ret = New Tuple(Of Boolean, String)(True, "Resistance 1 Target")
                End If
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

    Private Function GetSignalForEntry(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Decimal, Decimal, Decimal)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, Decimal, Decimal) = Nothing
        If _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _pivotPayload(currentCandle.PreviousCandlePayload.PayloadDate).Support3 AndAlso
            _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _pivotPayload(currentCandle.PreviousCandlePayload.PayloadDate).Support3 Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)
            Dim entryPrice As Decimal = _fractalHighPayload(signalCandle.PayloadDate) + buffer
            Dim stoploss As Decimal = _fractalLowPayload(signalCandle.PayloadDate) - buffer

            If entryPrice - stoploss >= 0.1 Then
                ret = New Tuple(Of Boolean, Payload, Decimal, Decimal, Decimal)(True, signalCandle, entryPrice, stoploss, buffer)
            End If
        End If
        Return ret
    End Function
End Class