Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TopGainerTopLooserOptionsBuyOnlyStrategy2
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public EntryBelowBuyLevel As Boolean
    End Class
#End Region

    Private _remarks As String
    Private _buyPrice As Decimal
    Private _sellPrice As Decimal

    Private _breakoutCandle As Payload = Nothing
    Private _targetReached As Boolean = False
    Private _lastLowestHigh As Payload = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _siganlTime As Date
    Private ReadOnly _stkNmbr As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal signalTime As Date,
                   ByVal stkNmbr As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _siganlTime = signalTime
        _userInputs = entities
        _stkNmbr = stkNmbr
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim nifty50Payload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Cash, "NIFTY 50", _tradingDate, _tradingDate)
        If nifty50Payload IsNot Nothing AndAlso nifty50Payload.ContainsKey(_siganlTime) Then
            Dim open As Decimal = nifty50Payload.FirstOrDefault.Value.Open
            Dim close As Decimal = nifty50Payload(_siganlTime).Close
            If close > open AndAlso _stkNmbr <= 10 Then
                _remarks = "Relevant"
            ElseIf close < open AndAlso _stkNmbr >= 11 Then
                _remarks = "Relevant"
            Else
                _remarks = "Contradicory"
            End If
        End If

        Dim first15MinHigh As Decimal = _signalPayload.Max(Function(x)
                                                               If x.Key.Date = _tradingDate.Date AndAlso x.Key <= _siganlTime Then
                                                                   Return x.Value.High
                                                               Else
                                                                   Return Decimal.MinValue
                                                               End If
                                                           End Function)
        Dim first15MinLow As Decimal = _signalPayload.Min(Function(x)
                                                              If x.Key.Date = _tradingDate.Date AndAlso x.Key <= _siganlTime Then
                                                                  Return x.Value.Low
                                                              Else
                                                                  Return Decimal.MaxValue
                                                              End If
                                                          End Function)

        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(first15MinHigh, RoundOfType.Floor)
        _buyPrice = first15MinHigh + buffer
        _sellPrice = first15MinLow - buffer
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso _buyPrice <> Decimal.MinValue AndAlso _buyPrice <> Decimal.MaxValue Then
            Dim signalCandle As Payload = Nothing
            If currentMinuteCandlePayload.PayloadDate > _siganlTime Then
                If _breakoutCandle Is Nothing Then
                    If currentMinuteCandlePayload.PreviousCandlePayload.High >= _buyPrice Then
                        _breakoutCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    End If
                End If
                If Not _targetReached Then
                    If currentTick.High >= _buyPrice + (_buyPrice - _sellPrice) * _userInputs.TargetMultiplier Then
                        _targetReached = True
                    End If
                End If
                If _breakoutCandle IsNot Nothing AndAlso Not _targetReached Then
                    If currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High < currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High Then
                        If currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PayloadDate >= _breakoutCandle.PayloadDate Then
                            If _userInputs.EntryBelowBuyLevel Then
                                If currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High < _buyPrice Then
                                    _lastLowestHigh = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload
                                End If
                            Else
                                _lastLowestHigh = currentMinuteCandlePayload.PreviousCandlePayload.PreviousCandlePayload
                            End If
                        End If
                    End If
                    If _lastLowestHigh IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload.Close > _lastLowestHigh.High Then
                        signalCandle = currentMinuteCandlePayload
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)
                Dim entryPrice As Decimal = signalCandle.High
                Dim quantity As Integer = Me.LotSize
                Dim stoploss As Decimal = _sellPrice
                Dim target As Decimal = entryPrice + ((entryPrice - stoploss) * _userInputs.TargetMultiplier)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = entryPrice - stoploss,
                                                            .Supporting3 = _remarks
                                                        }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
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
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function
End Class