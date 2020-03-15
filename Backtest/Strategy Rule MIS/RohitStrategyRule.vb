Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class RohitStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TraillingStoploss As Decimal
        Public TraillingStoplossPercentage As Decimal
        Public RSIUpperCircuit As Decimal
        Public RSILowerCircuit As Decimal
        Public CCIUpperCircuit As Decimal
        Public CCILowerCircuit As Decimal
        Public MACDUpperCircuit As Decimal
        Public MACDLowerCircuit As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _buffer As Decimal
    Private ReadOnly _target As Decimal
    Private ReadOnly _stoploss As Decimal
    Private ReadOnly _bb As Decimal
    Private ReadOnly _bbTarget As Decimal
    Private ReadOnly _bbStoploss As Decimal

    Private _hkPayloads As Dictionary(Of Date, Payload)
    Private _rsiPayloads As Dictionary(Of Date, Decimal)
    Private _cciPayloads As Dictionary(Of Date, Decimal)
    Private _macdPayloads As Dictionary(Of Date, Decimal)
    Private _macdSignalPayloads As Dictionary(Of Date, Decimal)
    Private _macdHistogramPayloads As Dictionary(Of Date, Decimal)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal buffer As Decimal,
                   ByVal target As Decimal,
                   ByVal stoploss As Decimal,
                   ByVal bb As Decimal,
                   ByVal bbTarget As Decimal,
                   ByVal bbStoploss As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _buffer = buffer
        _target = target
        _stoploss = stoploss
        _bb = bb
        _bbTarget = bbTarget
        _bbStoploss = bbStoploss
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)
        Indicator.RSI.CalculateRSI(4, _hkPayloads, _rsiPayloads)
        Indicator.CCI.CalculateCCI(20, _hkPayloads, _cciPayloads)
        Indicator.MACD.CalculateMACD(12, 26, 9, _hkPayloads, _macdPayloads, _macdSignalPayloads, _macdHistogramPayloads)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        'If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
        '    _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
        '    _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
        '    _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
        '    _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
        '    _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
        '    _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
        '    _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
        '    currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) Then
                Dim signalCandle As Payload = Nothing
                Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                    If Not (lastExecutedOrder IsNot Nothing AndAlso
                        lastExecutedOrder.SignalCandle.PayloadDate = signal.Item4.PayloadDate) Then
                        signalCandle = signal.Item4
                    End If
                End If
                If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                    If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                        Dim entry As Decimal = signal.Item2 + _buffer
                        Dim stoploss As Decimal = entry - _stoploss
                        Dim target As Decimal = entry + _target
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = _buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
                                }
                    ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                        Dim entry As Decimal = signal.Item2 - _buffer
                        Dim stoploss As Decimal = entry + _stoploss
                        Dim target As Decimal = entry - _target
                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = _buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
                                }
                    End If
                End If
            Else
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso lastExecutedTrade.Supporting1 <> "BB" Then
                    If lastExecutedTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If currentTick.Open <= lastExecutedTrade.EntryPrice - _bb Then
                            Dim entry As Decimal = currentTick.Open
                            Dim stoploss As Decimal = entry - _bbStoploss
                            Dim target As Decimal = entry + _bbTarget
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entry,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = Me.LotSize,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .Buffer = _buffer,
                                        .SignalCandle = currentMinuteCandlePayload,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "BB"
                                    }
                        End If
                    ElseIf lastExecutedTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If currentTick.Open >= lastExecutedTrade.EntryPrice + _bb Then
                            Dim entry As Decimal = currentTick.Open
                            Dim stoploss As Decimal = entry + _bbStoploss
                            Dim target As Decimal = entry - _bbTarget
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entry,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = Me.LotSize,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .Buffer = _buffer,
                                        .SignalCandle = currentMinuteCandlePayload,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = "BB"
                                    }
                        End If
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso currentTrade.EntryDirection <> signal.Item3 Then
                ret = New Tuple(Of Boolean, String)(True, "Opposite direction signal")
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso (currentTrade.EntryDirection <> signal.Item3 OrElse currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid signal")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.TraillingStoploss Then
            If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim currentMinuteHKCandle As Payload = _hkPayloads(currentMinuteCandle.PayloadDate)
                If currentMinuteHKCandle.PreviousCandlePayload IsNot Nothing Then
                    Dim triggerPrice As Decimal = Decimal.MinValue
                    Dim pointsToMoveSL As Decimal = Math.Abs(currentTrade.PotentialTarget - currentTrade.EntryPrice) * _userInputs.TraillingStoplossPercentage / 100
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        Dim price As Decimal = Decimal.MinValue
                        If currentTrade.SLRemark.Contains("Moved") Then
                            price = currentMinuteHKCandle.PreviousCandlePayload.Low
                        Else
                            If currentMinuteHKCandle.PreviousCandlePayload.High >= currentTrade.EntryPrice + pointsToMoveSL Then
                                price = currentMinuteHKCandle.PreviousCandlePayload.Low
                            End If
                        End If
                        If price <> Decimal.MinValue AndAlso price > currentTrade.PotentialStopLoss Then
                            triggerPrice = price
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        Dim price As Decimal = Decimal.MinValue
                        If currentTrade.SLRemark.Contains("Moved") Then
                            price = currentMinuteHKCandle.PreviousCandlePayload.High
                        Else
                            If currentMinuteHKCandle.PreviousCandlePayload.Low <= currentTrade.EntryPrice - pointsToMoveSL Then
                                price = currentMinuteHKCandle.PreviousCandlePayload.High
                            End If
                        End If
                        If price <> Decimal.MinValue AndAlso price < currentTrade.PotentialStopLoss Then
                            triggerPrice = price
                        End If
                    End If
                    If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Moved at {0} on {1}", triggerPrice, currentTick.PayloadDate.ToString("HH:mm:ss")))
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayloads(candle.PayloadDate)
            If _rsiPayloads(hkCandle.PayloadDate) < _userInputs.RSILowerCircuit AndAlso
                _rsiPayloads(hkCandle.PreviousCandlePayload.PayloadDate) > _userInputs.RSILowerCircuit Then
                If _macdPayloads(hkCandle.PayloadDate) < _macdSignalPayloads(hkCandle.PayloadDate) AndAlso
                    _macdPayloads(hkCandle.PayloadDate) < _userInputs.MACDLowerCircuit AndAlso
                    _macdSignalPayloads(hkCandle.PayloadDate) < _userInputs.MACDLowerCircuit Then
                    If _cciPayloads(hkCandle.PayloadDate) < _userInputs.CCILowerCircuit Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, hkCandle.Low, Trade.TradeExecutionDirection.Sell, hkCandle)
                    End If
                End If
            ElseIf _rsiPayloads(hkCandle.PayloadDate) > _userInputs.RSIUpperCircuit AndAlso
                _rsiPayloads(hkCandle.PreviousCandlePayload.PayloadDate) < _userInputs.RSIUpperCircuit Then
                If _macdPayloads(hkCandle.PayloadDate) > _macdSignalPayloads(hkCandle.PayloadDate) AndAlso
                    _macdPayloads(hkCandle.PayloadDate) > _userInputs.MACDUpperCircuit AndAlso
                    _macdSignalPayloads(hkCandle.PayloadDate) > _userInputs.MACDUpperCircuit Then
                    If _cciPayloads(hkCandle.PayloadDate) > _userInputs.CCIUpperCircuit Then
                        ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, hkCandle.High, Trade.TradeExecutionDirection.Buy, hkCandle)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class