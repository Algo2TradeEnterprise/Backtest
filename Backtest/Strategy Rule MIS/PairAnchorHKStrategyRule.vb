Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PairAnchorHKStrategyRule
    Inherits StrategyRule

    Public DummyCandle As Payload = Nothing
    Public DirectionToTrade As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public DifferencePercentage As Decimal = Decimal.MinValue

    Private ReadOnly _controller As Boolean

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _firstCandleOfTheDay As Payload = Nothing
    Private _highestATR As Decimal = Decimal.MinValue

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If controller = 1 Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousTradingDay As Date = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, _tradingDate)

            Dim previousDayFirstCandleClose As Decimal = _signalPayload.Where(Function(x)
                                                                                  Return x.Key.Date = previousTradingDay.Date
                                                                              End Function).OrderBy(Function(y)
                                                                                                        Return y.Key
                                                                                                    End Function).FirstOrDefault.Value.Close

            Dim currentDayFirstCandleClose As Decimal = _signalPayload.Where(Function(x)
                                                                                 Return x.Key.Date = _tradingDate.Date
                                                                             End Function).OrderBy(Function(y)
                                                                                                       Return y.Key
                                                                                                   End Function).FirstOrDefault.Value.Close

            Me.DifferencePercentage = ((currentDayFirstCandleClose - previousDayFirstCandleClose) / currentDayFirstCandleClose) * 100
        End If

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()

        Dim myPair As PairAnchorHKStrategyRule = Me.AnotherPairInstrument
        If Me.DifferencePercentage <> Decimal.MinValue AndAlso myPair.DifferencePercentage <> Decimal.MinValue Then
            If Me.DifferencePercentage > myPair.DifferencePercentage Then
                Me.DirectionToTrade = Trade.TradeExecutionDirection.Sell
                myPair.DirectionToTrade = Trade.TradeExecutionDirection.Buy
            Else
                Me.DirectionToTrade = Trade.TradeExecutionDirection.Buy
                myPair.DirectionToTrade = Trade.TradeExecutionDirection.Sell
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If DirectionToTrade = Trade.TradeExecutionDirection.Buy AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy, True)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(1, 5000, entryPrice, Trade.TypeOfStock.Cash, True)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 1000000000000,
                                .Target = .EntryPrice + 1000000000000,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = signal.Item4,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = If(_highestATR = Decimal.MinValue, "∞", _highestATR)
                            }
                End If
            End If
            If DirectionToTrade = Trade.TradeExecutionDirection.Sell AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell, True)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim entryPrice As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(1, 5000, entryPrice, Trade.TypeOfStock.Cash, True)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + 1000000000000,
                                .Target = .EntryPrice - 1000000000000,
                                .Buffer = 0,
                                .SignalCandle = signal.Item2,
                                .OrderType = signal.Item4,
                                .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = If(_highestATR = Decimal.MinValue, "∞", _highestATR)
                            }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            Me.AnotherPairInstrument.ForceTakeTrade = True
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection, False)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso currentTrade.EntryPrice <> signal.Item3 Then
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection, ByVal forcePrint As Boolean) As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)
        Dim ret As Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder) = Nothing
        If candle IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(candle, Trade.TypeOfTrade.MIS, direction)
            If lastExecutedTrade IsNot Nothing Then
                If _highestATR = Decimal.MinValue Then
                    _highestATR = ConvertFloorCeling(GetHighestATR(lastExecutedTrade.SignalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                End If
                If direction = Trade.TradeExecutionDirection.Buy Then
                    If hkCandle.High <= lastExecutedTrade.EntryPrice - _highestATR Then
                        If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                            Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                            If entryPrice = hkCandle.High Then
                                entryPrice = hkCandle.High + _parentStrategy.CalculateBuffer(hkCandle.High, RoundOfType.Floor)
                            End If
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, hkCandle, entryPrice, Trade.TypeOfOrder.SL)
                        End If
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    If hkCandle.Low >= lastExecutedTrade.EntryPrice + _highestATR Then
                        If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                            Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                            If entryPrice = hkCandle.Low Then
                                entryPrice = hkCandle.Low - _parentStrategy.CalculateBuffer(hkCandle.Low, RoundOfType.Floor)
                            End If
                            ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, hkCandle, entryPrice, Trade.TypeOfOrder.SL)
                        End If
                    End If
                End If
            Else
                If direction = Trade.TradeExecutionDirection.Buy Then
                    If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                        Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        If entryPrice = hkCandle.High Then
                            entryPrice = hkCandle.High + _parentStrategy.CalculateBuffer(hkCandle.High, RoundOfType.Floor)
                        End If
                        ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, hkCandle, entryPrice, Trade.TypeOfOrder.SL)
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                        Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        If entryPrice = hkCandle.Low Then
                            entryPrice = hkCandle.Low - _parentStrategy.CalculateBuffer(hkCandle.Low, RoundOfType.Floor)
                        End If
                        ret = New Tuple(Of Boolean, Payload, Decimal, Trade.TypeOfOrder)(True, hkCandle, entryPrice, Trade.TypeOfOrder.SL)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            ret = _atrPayload.Max(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso
                                        x.Key <= signalCandle.PayloadDate Then
                                          Return x.Value
                                      Else
                                          Return Decimal.MinValue
                                      End If
                                  End Function)
        End If
        Return ret
    End Function
End Class