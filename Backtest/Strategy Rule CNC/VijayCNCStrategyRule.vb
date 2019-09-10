Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class VijayCNCStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public RefreshQuantityAtDayStart As Boolean
    End Class
#End Region

    Private ReadOnly _StartingQuantity As Integer = 1
    Private ReadOnly _QuantityMultiplier As Decimal = 1.5
    Private ReadOnly _TargetPerecentage As Decimal = 1
    Private ReadOnly _DropPercentage As Decimal = 0.5

    Private _FirstLTP As Decimal = Decimal.MinValue
    Private _CurrentTarget As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing
            Dim quantity As Integer = _StartingQuantity
            Dim remark As String = ""
            Dim averageTradePrice As Decimal = currentTick.Open
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
            If lastExecutedTrade IsNot Nothing Then
                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If (lastExecutedTrade.EntryPrice - currentTick.Open) >= lastExecutedTrade.EntryPrice * _DropPercentage / 100 Then
                        signalCandle = currentMinuteCandlePayload
                        If CType(_entities, StrategyRuleEntities).RefreshQuantityAtDayStart Then         'Refresh Quantity at day start
                            If lastExecutedTrade.EntryTime.Date = _tradingDate.Date Then
                                quantity = Math.Ceiling(lastExecutedTrade.Quantity * _QuantityMultiplier)
                            End If
                        Else
                            quantity = Math.Ceiling(lastExecutedTrade.Quantity * _QuantityMultiplier)
                        End If

                        Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                        If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                            Dim totalCapitalUsedWithoutMargin As Decimal = 0
                            Dim totalQuantity As Decimal = 0
                            For Each runningTrade In openActiveTrades
                                If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                    totalCapitalUsedWithoutMargin += runningTrade.EntryPrice * runningTrade.Quantity
                                    totalQuantity += runningTrade.Quantity
                                End If
                            Next
                            totalCapitalUsedWithoutMargin += currentTick.Open * quantity
                            totalQuantity += quantity
                            averageTradePrice = totalCapitalUsedWithoutMargin / totalQuantity
                        End If
                    End If
                ElseIf lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                    signalCandle = currentMinuteCandlePayload
                    remark = "Fresh start"
                End If
            Else
                'If _FirstLTP = Decimal.MinValue Then
                '    _FirstLTP = currentTick.Open
                'Else
                '    If (_FirstLTP - currentTick.Open) <= _FirstLTP * _DropPercentage / 100 Then
                signalCandle = currentMinuteCandlePayload
                remark = "Fresh start"
                '    End If
                'End If
            End If

            If signalCandle IsNot Nothing Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 100000,
                            .Target = ConvertFloorCeling(averageTradePrice + (averageTradePrice * _TargetPerecentage / 100), _parentStrategy.TickSize, RoundOfType.Celing),
                            .Buffer = 0,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = remark
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
            If _CurrentTarget <> Decimal.MinValue AndAlso currentTrade.PotentialTarget <> _CurrentTarget Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, _CurrentTarget, String.Format("{0}. Time:{1}", _CurrentTarget, currentTick.PayloadDate))
            End If
        End If
        Return ret
    End Function
End Class
