Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowStoplossFractalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinStoplossPerTrade As Decimal
        Public MaxStoplossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    'Private _hkPayloads As Dictionary(Of Date, Payload)
    Private _fractalHighPayloads As Dictionary(Of Date, Decimal)
    Private _fractalLowPayloads As Dictionary(Of Date, Decimal)

    Private _buyLevel As Decimal = Decimal.MinValue
    Private _sellLevel As Decimal = Decimal.MinValue
    Private _signalCandle As Payload = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        'Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayloads)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayloads, _fractalLowPayloads)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If Not IsSignalTriggered(_buyLevel, Trade.TradeExecutionDirection.Buy, _signalCandle.PayloadDate.AddMinutes(1), currentMinuteCandlePayload.PayloadDate) AndAlso
                    Not IsSignalTriggered(_sellLevel, Trade.TradeExecutionDirection.Sell, _signalCandle.PayloadDate.AddMinutes(1), currentMinuteCandlePayload.PayloadDate) Then
                    Dim quantity As Decimal = Me.LotSize
                    Dim buffer As Decimal = 0
                    Dim slPoint As Decimal = Math.Abs(_buyLevel - _sellLevel)
                    Dim target As Decimal = slPoint * _userInputs.TargetMultiplier
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                        If currentTick.Open < signal.Item2 + buffer Then
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = signal.Item2,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - slPoint,
                                        .Target = .EntryPrice + target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item3,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = (.EntryPrice - .Stoploss)
                                    }
                        End If
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                        If currentTick.Open > signal.Item2 - buffer Then
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = signal.Item2,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice + slPoint,
                                        .Target = .EntryPrice - target,
                                        .Buffer = buffer,
                                        .SignalCandle = signal.Item3,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signal.Item3.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = (.Stoploss - .EntryPrice)
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

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item4 Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting2
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                triggerPrice = currentTrade.EntryPrice - slPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                triggerPrice = currentTrade.EntryPrice + slPoint
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, slPoint)
            End If
        End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _buyLevel = Decimal.MinValue AndAlso _sellLevel = Decimal.MinValue Then
                'Dim hkCandle As Payload = _hkPayloads(candle.PayloadDate)
                Dim fractalHigh As Decimal = _fractalHighPayloads(candle.PayloadDate)
                Dim fractalLow As Decimal = _fractalLowPayloads(candle.PayloadDate)
                If candle.Close > fractalLow AndAlso candle.Close < fractalHigh Then
                    Dim slPoint As Decimal = Math.Abs(fractalHigh - fractalLow)
                    Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, candle.High, candle.High - slPoint, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                    If Math.Abs(pl) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso Math.Abs(pl) <= Math.Abs(_userInputs.MaxStoplossPerTrade) Then
                        _buyLevel = ConvertFloorCeling(fractalHigh, _parentStrategy.TickSize, RoundOfType.Floor)
                        _sellLevel = ConvertFloorCeling(fractalLow, _parentStrategy.TickSize, RoundOfType.Floor)
                        _signalCandle = candle
                    End If
                End If
            End If
            If _buyLevel <> Decimal.MinValue AndAlso _sellLevel <> Decimal.MinValue Then
                Dim middleLevel As Decimal = (_buyLevel + _sellLevel) / 2
                Dim range As Decimal = (_buyLevel - middleLevel) * 50 / 100
                If currentTick.Open >= middleLevel + range Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, _buyLevel, _signalCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf currentTick.Open <= middleLevel - range Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, _sellLevel, _signalCandle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class