Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKRSIColorRetracementEntryStrategyRule
    Inherits StrategyRule

    '#Region "Entity"
    '    Public Class StrategyRuleEntities
    '        Inherits RuleEntities

    '        Public MaxLossPerTrade As Decimal
    '        Public MaxProfitPerTrade As Decimal
    '    End Class
    '#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _rsiPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousTradingDay As Date

    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    'Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _previousLoss As Decimal
    Private ReadOnly _previousIteration As Integer
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal previousLoss As Decimal,
                   ByVal previousIteration As Integer,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        '_userInputs = _entities
        _previousLoss = previousLoss
        _previousIteration = previousIteration
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
        Indicator.RSI.CalculateRSI(14, _hkPayload, _rsiPayload)
        _previousTradingDay = _signalPayload.Where(Function(x)
                                                       Return x.Key.Date = _tradingDate.Date
                                                   End Function).FirstOrDefault.Value.PreviousCandlePayload.PayloadDate
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item3
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = ConvertFloorCeling(GetAverageHighestATR(signalCandle) * 1.5, _parentStrategy.TickSize, RoundOfType.Celing)
                Dim lossPL As Decimal = -500 + _previousLoss
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = 500 - _previousLoss
                Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousIteration + 1,
                                    .Supporting3 = signal.Item5
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousIteration + 1,
                                    .Supporting3 = signal.Item5
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
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            If _direction = Trade.TradeExecutionDirection.None Then
                Dim rsiDirection As Trade.TradeExecutionDirection = GetRSIConfirmDirection(hkCandle)
                If rsiDirection = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.CandleColor = Color.Red Then
                    Dim buyLevel As Decimal = ConvertFloorCeling(Math.Round(hkCandle.High, 2), _parentStrategy.TickSize, RoundOfType.Celing)
                    If buyLevel = Math.Round(hkCandle.High, 2) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                        buyLevel = buyLevel + buffer
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, buyLevel, hkCandle, Trade.TradeExecutionDirection.Buy, "")
                ElseIf rsiDirection = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.CandleColor = Color.Green Then
                    Dim sellLevel As Decimal = ConvertFloorCeling(Math.Round(hkCandle.Low, 2), _parentStrategy.TickSize, RoundOfType.Floor)
                    If sellLevel = Math.Round(hkCandle.Low, 2) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                        sellLevel = sellLevel - buffer
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell, "")
                End If
            Else
                'If _direction = Trade.TradeExecutionDirection.Buy AndAlso Math.Round(hkCandle.High, 2) = Math.Round(hkCandle.Open, 2) Then
                '    Dim buyLevel As Decimal = ConvertFloorCeling(Math.Round(hkCandle.High, 2), _parentStrategy.TickSize, RoundOfType.Celing)
                '    If buyLevel = Math.Round(hkCandle.High, 2) Then
                '        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                '        buyLevel = buyLevel + buffer
                '    End If
                '    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle, Trade.TradeExecutionDirection.Buy)
                'ElseIf _direction = Trade.TradeExecutionDirection.Sell AndAlso Math.Round(hkCandle.Low, 2) = Math.Round(hkCandle.Open, 2) Then
                '    Dim sellLevel As Decimal = ConvertFloorCeling(Math.Round(hkCandle.Low, 2), _parentStrategy.TickSize, RoundOfType.Floor)
                '    If sellLevel = Math.Round(hkCandle.Low, 2) Then
                '        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                '        sellLevel = sellLevel - buffer
                '    End If
                '    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell)
                'End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 AndAlso signalCandle IsNot Nothing Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)
            Dim previousDayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                       If x.Key.Date = _previousTradingDay.Date Then
                                                                           Return x.Value
                                                                       Else
                                                                           Return Decimal.MinValue
                                                                       End If
                                                                   End Function)
            If todayHighestATR <> Decimal.MinValue Then
                ret = (todayHighestATR + previousDayHighestATR) / 2
            End If
        End If
        Return ret
    End Function

    Private Function GetRSIConfirmDirection(ByVal signalCandle As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        Dim rsiDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        Dim rsiTime As Date = Date.MinValue
        Dim rsiDirectionTime As Tuple(Of Trade.TradeExecutionDirection, Date) = GetRSIDirection(signalCandle.PayloadDate)
        If rsiDirectionTime IsNot Nothing AndAlso rsiDirectionTime.Item1 <> Trade.TradeExecutionDirection.None Then
            Dim prersiDirectionTime As Tuple(Of Trade.TradeExecutionDirection, Date) = rsiDirectionTime
            While True
                rsiDirectionTime = GetRSIDirection(prersiDirectionTime.Item2)
                If rsiDirectionTime Is Nothing OrElse rsiDirectionTime.Item1 <> prersiDirectionTime.Item1 Then
                    rsiDirection = prersiDirectionTime.Item1
                    rsiTime = prersiDirectionTime.Item2
                    Exit While
                Else
                    prersiDirectionTime = rsiDirectionTime
                End If
            End While
        End If
        If rsiDirection <> Trade.TradeExecutionDirection.None AndAlso rsiTime <> Date.MinValue Then
            For Each runningPayload In _hkPayload
                If runningPayload.Key >= rsiTime AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                    If rsiDirection = Trade.TradeExecutionDirection.Buy Then
                        If runningPayload.Value.CandleColor = Color.Green Then
                            ret = Trade.TradeExecutionDirection.Buy
                            Exit For
                        End If
                    ElseIf rsiDirection = Trade.TradeExecutionDirection.Sell Then
                        If runningPayload.Value.CandleColor = Color.Red Then
                            ret = Trade.TradeExecutionDirection.Sell
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetRSIDirection(ByVal signalTime As Date) As Tuple(Of Trade.TradeExecutionDirection, Date)
        Dim ret As Tuple(Of Trade.TradeExecutionDirection, Date) = Nothing
        For Each runningPayload In _hkPayload.OrderByDescending(Function(x)
                                                                    Return x.Key
                                                                End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key < signalTime Then
                Dim rsi As Decimal = _rsiPayload(runningPayload.Key)
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                    Dim preRsi As Decimal = _rsiPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                    If rsi >= 70 AndAlso preRsi < 70 Then
                        ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Sell, runningPayload.Key)
                        Exit For
                    ElseIf rsi <= 30 AndAlso preRsi > 30 Then
                        ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Buy, runningPayload.Key)
                        Exit For
                    End If
                Else
                    If rsi >= 70 Then
                        ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Sell, runningPayload.Key)
                        Exit For
                    ElseIf rsi <= 30 Then
                        ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Buy, runningPayload.Key)
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function
End Class