Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LossMakeupRainbowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerStock As Decimal
        Public MaxProfitPerStock As Decimal
    End Class
#End Region

    Private _sma1Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _sma2Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _sma3Payload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _checkSignal As Boolean = False
    Private _firstCandleOfTheDay As Payload = Nothing
    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue
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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.SMA.CalculateSMA(2, Payload.PayloadFields.Close, _signalPayload, _sma1Payload)
        Dim sma1Payload As Dictionary(Of Date, Payload) = Nothing
        Common.ConvertDecimalToPayload(Payload.PayloadFields.Close, _sma1Payload, sma1Payload)
        Indicator.SMA.CalculateSMA(2, Payload.PayloadFields.Close, sma1Payload, _sma2Payload)
        Dim sma2Payload As Dictionary(Of Date, Payload) = Nothing
        Common.ConvertDecimalToPayload(Payload.PayloadFields.Close, _sma2Payload, sma2Payload)
        Indicator.SMA.CalculateSMA(2, Payload.PayloadFields.Close, sma2Payload, _sma3Payload)

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    _firstCandleOfTheDay = runningPayload.Value
                    Exit For
                End If
            End If
        Next
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso Not IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item2
                    _slPoint = ConvertFloorCeling(GetHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                    _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open - _slPoint, _userInputs.MaxLossPerStock, Trade.TypeOfStock.Cash)
                    _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, currentTick.Open, _quantity, _userInputs.MaxProfitPerStock, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - currentTick.Open
                ElseIf lastExecutedOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload))
                    If signal.Item2.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate AndAlso
                        exitCandle.PayloadDate <> currentMinuteCandlePayload.PayloadDate Then
                        signalCandle = signal.Item2
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim slPoint As Decimal = _slPoint
                    Dim targetPoint As Decimal = _targetPoint
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = targetRemark,
                                    .Supporting4 = GetTradeNumber(currentMinuteCandlePayload) + 1,
                                    .Supporting5 = signal.Item4
                                }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), Trade.TypeOfStock.Cash)
                        targetRemark = "SL Makeup"

                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = targetRemark,
                                    .Supporting4 = GetTradeNumber(currentMinuteCandlePayload) + 1,
                                    .Supporting5 = signal.Item4
                                }
                    End If
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim slPoint As Decimal = _slPoint
                    Dim targetPoint As Decimal = _targetPoint
                    Dim targetRemark As String = "Normal"
                    Dim quantity As Integer = _quantity

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = targetRemark,
                                    .Supporting4 = GetTradeNumber(currentMinuteCandlePayload) + 1,
                                    .Supporting5 = signal.Item4
                                }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), Trade.TypeOfStock.Cash)
                        targetRemark = "SL Makeup"

                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _slPoint,
                                    .Supporting3 = targetRemark,
                                    .Supporting4 = GetTradeNumber(currentMinuteCandlePayload) + 1,
                                    .Supporting5 = signal.Item4
                                }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.Supporting5.ToUpper.Contains("LIMIT") AndAlso currentMinuteCandlePayload.PayloadDate > currentTrade.EntryTime Then
                Dim entryCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    triggerPrice = entryCandle.Low - _parentStrategy.CalculateBuffer(entryCandle.Low, RoundOfType.Floor)
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    triggerPrice = entryCandle.High + _parentStrategy.CalculateBuffer(entryCandle.High, RoundOfType.Floor)
                End If
            ElseIf currentTrade.Supporting5.ToUpper.Contains("MARKET") Then
                Dim signalCandle As Payload = currentTrade.SignalCandle
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    If currentTrade.Supporting5.ToUpper = "SMA3 Cut Market Entry" Then
                        triggerPrice = signalCandle.Low - _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Else
                        triggerPrice = Math.Min(signalCandle.Low, signalCandle.PreviousCandlePayload.Low) - _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    If currentTrade.Supporting5.ToUpper = "SMA3 Cut Market Entry" Then
                        triggerPrice = signalCandle.High + _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Else
                        triggerPrice = Math.Max(signalCandle.High, signalCandle.PreviousCandlePayload.High) + _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice AndAlso
                Math.Abs(triggerPrice - currentTrade.EntryPrice) < _slPoint Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, Math.Abs(triggerPrice - currentTrade.EntryPrice))
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            If Not _checkSignal Then
                If _firstCandleOfTheDay.Open > _sma1Payload(_firstCandleOfTheDay.PayloadDate) AndAlso candle.Low < _sma1Payload(candle.PayloadDate) Then
                    _checkSignal = True
                ElseIf _firstCandleOfTheDay.Open < _sma1Payload(_firstCandleOfTheDay.PayloadDate) AndAlso candle.High > _sma3Payload(candle.PayloadDate) Then
                    _checkSignal = True
                End If
            End If
            If _checkSignal Then
                Dim sma1 As Decimal = _sma1Payload(candle.PayloadDate)
                Dim sma2 As Decimal = _sma2Payload(candle.PayloadDate)
                Dim sma3 As Decimal = _sma3Payload(candle.PayloadDate)
                If sma1 > sma2 AndAlso sma2 > sma3 Then
                    If IsSignalValid(candle, Trade.TradeExecutionDirection.Buy) Then
                        If candle.Low <= sma3 Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Buy, "SMA3 Cut Market Entry")
                        ElseIf candle.CandleColor = Color.Red Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Buy, "Red Candle Market Entry")
                        ElseIf currentTick.Open <= Math.Max(candle.Low, sma3) Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Buy, "Candle Low/SMA3 Limit Entry")
                        End If
                    End If
                ElseIf sma1 < sma2 AndAlso sma2 < sma3 Then
                    If IsSignalValid(candle, Trade.TradeExecutionDirection.Sell) Then
                        If candle.High >= sma3 Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Sell, "SMA3 Cut Market Entry")
                        ElseIf candle.CandleColor = Color.Green Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Sell, "Green Candle Market Entry")
                        ElseIf currentTick.Open >= Math.Min(candle.High, sma3) Then
                            ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, String)(True, candle, Trade.TradeExecutionDirection.Sell, "Candle High/SMA3 Limit Entry")
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsSignalValid(ByVal signalCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim candle As Payload = signalCandle
        While candle IsNot Nothing
            If candle.PayloadDate.Date <> _tradingDate.Date Then Exit While
            Dim sma1 As Decimal = _sma1Payload(candle.PayloadDate)
            Dim sma2 As Decimal = _sma2Payload(candle.PayloadDate)
            Dim sma3 As Decimal = _sma3Payload(candle.PayloadDate)
            If direction = Trade.TradeExecutionDirection.Buy Then
                If sma1 > sma2 AndAlso sma2 > sma3 Then
                    If candle.Close > sma1 Then
                        ret = True
                        Exit While
                    End If
                Else
                    Exit While
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If sma1 < sma2 AndAlso sma2 < sma3 Then
                    If candle.Close < sma1 Then
                        ret = True
                        Exit While
                    End If
                Else
                    Exit While
                End If
            End If

            candle = candle.PreviousCandlePayload
        End While
        Return ret
    End Function

    Private Function IsAnyTradeOfTheStockTargetReached(ByVal currentMinutePayload As Payload) As Boolean
        Dim ret As Boolean = False
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            Dim targetTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                            Return x.ExitCondition = Trade.TradeExitCondition.Target AndAlso
                                                                            x.AdditionalTrade = False AndAlso x.Supporting3 = "Normal"
                                                                        End Function)
            If targetTrades IsNot Nothing AndAlso targetTrades.Count > 0 Then
                ret = True
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            If _firstCandleOfTheDay IsNot Nothing AndAlso _firstCandleOfTheDay.PreviousCandlePayload IsNot Nothing Then
                ret = _atrPayload.Max(Function(x)
                                          If x.Key.Date >= _firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date AndAlso
                                            x.Key <= signalCandle.PayloadDate Then
                                              Return x.Value
                                          Else
                                              Return Decimal.MinValue
                                          End If
                                      End Function)
            End If
        End If
        Return ret
    End Function

    Private Function GetTradeNumber(ByVal currentMinutePayload As Payload) As Integer
        Dim ret As Integer = 0
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            Dim mainTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                          Return x.Supporting3 = "Normal"
                                                                      End Function)
            ret = mainTrades.Count
        End If
        Return ret
    End Function
End Class