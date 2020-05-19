Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPercentage As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _lastSignal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = Nothing
    Private _lastTrade As Trade = Nothing

    Private ReadOnly _buffer As Decimal
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal buffer As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _buffer = buffer
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        _hkPayload = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinutePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinutePayload IsNot Nothing AndAlso currentMinutePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso Not ForceCancellationDone Then
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinutePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim quantity As Integer = Me.LotSize
                'Dim slPoint As Decimal = ConvertFloorCeling(signal.Item2 * _userInputs.StoplossPercentage / 100, _parentStrategy.GetTickSize(_tradingSymbol), RoundOfType.Floor)
                Dim slPoint As Decimal = 100000000000000
                Dim tgtPoint As Decimal = 100000000000000

                If signal.Item3 = Trade.TradeExecutionDirection.Buy AndAlso Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy) Then
                    'Dim enrtyPrice As Decimal = signal.Item2 + _buffer
                    Dim enrtyPrice As Decimal = signal.Item2
                    Dim orderType As Trade.TypeOfOrder = Trade.TypeOfOrder.SL
                    If currentTick.Open >= enrtyPrice Then
                        orderType = Trade.TypeOfOrder.Market
                        enrtyPrice = currentTick.Open
                    End If

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = enrtyPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - slPoint,
                                .Target = .EntryPrice + tgtPoint,
                                .Buffer = _buffer,
                                .SignalCandle = currentMinutePayload.PreviousCandlePayload,
                                .OrderType = orderType,
                                .Supporting1 = currentMinutePayload.PreviousCandlePayload.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                                .Supporting2 = orderType.ToString
                            }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell AndAlso Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell) Then
                    'Dim enrtyPrice As Decimal = signal.Item2 - _buffer
                    Dim enrtyPrice As Decimal = signal.Item2
                    Dim orderType As Trade.TypeOfOrder = Trade.TypeOfOrder.SL
                    If currentTick.Open <= enrtyPrice Then
                        orderType = Trade.TypeOfOrder.Market
                        enrtyPrice = currentTick.Open
                    End If

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = enrtyPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - tgtPoint,
                                .Buffer = _buffer,
                                .SignalCandle = currentMinutePayload.PreviousCandlePayload,
                                .OrderType = orderType,
                                .Supporting1 = currentMinutePayload.PreviousCandlePayload.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                                .Supporting2 = orderType.ToString
                            }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            Me.ContractRolloverForceEntry = False
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinutePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentMinutePayload.PayloadDate = _signalPayload.LastOrDefault.Key AndAlso (Me.ContractRollover OrElse Me.BlankDayExit) Then
                ret = New Tuple(Of Boolean, String)(True, If(Me.ContractRollover, "Contract Rollover Exit", "Blank Day Exit"))
                ForceCancellationDone = True
            Else
                Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalForEntry(currentMinutePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.SignalCandle.PayloadDate <> currentMinutePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Me.ContractRollover OrElse Me.BlankDayExit Then
                Dim currentMinutePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                If currentMinutePayload.PayloadDate = _signalPayload.LastOrDefault.Key Then
                    ret = New Tuple(Of Boolean, String)(True, If(Me.ContractRollover, "Contract Rollover Exit", "Blank Day Exit"))
                    ForceCancellationDone = True
                    _lastTrade = currentTrade
                    _lastSignal = GetSignalForEntry(currentMinutePayload.PreviousCandlePayload, currentTick)
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = Nothing
        If Not Me.ContractRolloverForceEntry Then
            If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
                Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
                If hkCandle IsNot Nothing AndAlso hkCandle.PreviousCandlePayload IsNot Nothing Then
                    If hkCandle.CandleColor <> hkCandle.PreviousCandlePayload.CandleColor Then
                        If hkCandle.CandleColor = Color.Green Then
                            Dim price As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.GetTickSize(_tradingSymbol), RoundOfType.Celing)
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, price, Trade.TradeExecutionDirection.Buy)
                        ElseIf hkCandle.CandleColor = Color.Red Then
                            Dim price As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.GetTickSize(_tradingSymbol), RoundOfType.Floor)
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, price, Trade.TradeExecutionDirection.Sell)
                        End If
                    End If
                End If
            End If
        Else        'Rollover entry
            If _lastSignal IsNot Nothing AndAlso _lastSignal.Item1 Then
                If _lastSignal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    If currentTick.Open > _lastSignal.Item2 Then
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open - _buffer, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Buy)
                    Else
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open + _buffer, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Sell)
                    End If
                ElseIf _lastSignal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    If currentTick.Open < _lastSignal.Item2 Then
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open + _buffer, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Sell)
                    Else
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open - _buffer, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Buy)
                    End If
                End If
            Else
                If _lastTrade IsNot Nothing Then
                    If _lastTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open - _buffer, Trade.TradeExecutionDirection.Buy)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Buy)
                    ElseIf _lastTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        'ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open + _buffer, Trade.TradeExecutionDirection.Sell)
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, currentTick.Open, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
