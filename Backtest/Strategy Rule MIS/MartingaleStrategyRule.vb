Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
        Public MaxProfitPerStock As Decimal
        Public ReverseSignalExit As Boolean
    End Class
#End Region

    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
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
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)
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
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item2
                    _slPoint = ConvertFloorCeling(GetHighestATR(signalCandle) * _userInputs.ATRMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    _targetPoint = _slPoint
                    If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                        _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open + _targetPoint, _userInputs.MaxProfitPerStock, Trade.TypeOfStock.Cash)
                        Me.MaxProfitOfThisStock = _userInputs.MaxProfitPerStock
                    Else
                        _quantity = Me.LotSize
                        Me.MaxProfitOfThisStock = _parentStrategy.CalculatePL(_tradingSymbol, currentTick.Open, currentTick.Open + _targetPoint, _quantity, Me.LotSize, _parentStrategy.StockType)
                    End If
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item2.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate Then
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
                                    .Supporting4 = 1
                                }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), _parentStrategy.StockType)
                        If _parentStrategy.StockType <> Trade.TypeOfStock.Cash Then
                            quantity = Math.Ceiling(quantity / Me.LotSize) * Me.LotSize
                        End If
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        targetPoint = targetPrice - entryPrice
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
                                    .Supporting3 = targetRemark
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
                                    .Supporting3 = targetRemark
                                }

                    Dim pl As Decimal = _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    If pl < 0 Then
                        quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, Math.Abs(pl), _parentStrategy.StockType)
                        If _parentStrategy.StockType <> Trade.TypeOfStock.Cash Then
                            quantity = Math.Ceiling(quantity / Me.LotSize) * Me.LotSize
                        End If
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                        targetPoint = entryPrice - targetPrice
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
                                    .Supporting3 = targetRemark
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
        If _userInputs.ReverseSignalExit AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item3 Then
                    ret = New Tuple(Of Boolean, String)(True, "Opposite Direction Signal")
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim vwap As Decimal = _vwapPayload(candle.PayloadDate)
            If candle.Open > vwap AndAlso candle.Close > vwap Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
            ElseIf candle.Open < vwap AndAlso candle.Close < vwap Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function

    Private Function IsAnyTradeOfTheStockTargetReached(ByVal currentMinutePayload As Payload) As Boolean
        Dim ret As Boolean = False
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            Dim targetTrades As List(Of Trade) = completeTrades.FindAll(Function(x)
                                                                            Return x.ExitCondition = Trade.TradeExitCondition.Target AndAlso
                                                                            x.AdditionalTrade = False AndAlso x.Supporting2 = "Normal"
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
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                ret = _atrPayload.Max(Function(x)
                                          If x.Key.Date >= firstCandle.PreviousCandlePayload.PayloadDate.Date AndAlso
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
End Class