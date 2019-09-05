Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class VijayCNCStrategyRule
    Inherits StrategyRule

    Private ReadOnly _StartingQuantity As Integer = 1
    Private ReadOnly _QuantityMultiplier As Decimal = 1
    Private ReadOnly _TargetPerecentage As Decimal = 3
    Private ReadOnly _DropPercentage As Decimal = 5

    Private _FirstLTP As Decimal = Decimal.MinValue
    Private _CurrentTarget As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TradeType.CNC) AndAlso currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim quantity As Integer = _StartingQuantity

            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TradeType.MIS)
            If lastExecutedTrade IsNot Nothing Then
                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If (lastExecutedTrade.EntryPrice - currentTick.Open) <= lastExecutedTrade.EntryPrice * _DropPercentage / 100 Then
                        signalCandle = currentMinuteCandlePayload
                        quantity = lastExecutedTrade.Quantity * _QuantityMultiplier
                    End If
                ElseIf lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                    signalCandle = currentMinuteCandlePayload
                End If
            Else
                'If _FirstLTP = Decimal.MinValue Then
                '    _FirstLTP = currentTick.Open
                'Else
                '    If (_FirstLTP - currentTick.Open) <= _FirstLTP * _DropPercentage / 100 Then
                signalCandle = currentMinuteCandlePayload
                '    End If
                'End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 100000,
                            .Target = .EntryPrice + ConvertFloorCeling(.EntryPrice * _TargetPerecentage / 100, _parentStrategy.TickSize, RoundOfType.Celing),
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.Market
                        }
            End If
        End If
        If parameter IsNot Nothing Then
            _CurrentTarget = parameter.Target
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If currentTrade.PotentialTarget <> _CurrentTarget Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, _CurrentTarget, String.Format("{0}. Time:{1}", _CurrentTarget, currentTick.PayloadDate))
            End If
        End If
        Return ret
    End Function
End Class
