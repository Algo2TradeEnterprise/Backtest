Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AlwaysInTradeMartingaleStrategyRule
    Inherits StrategyRule

    Private _oneSlabProfit As Decimal = Decimal.MinValue
    Private _25SlabProfit As Decimal = Decimal.MinValue
    Private _2SlabProfit As Decimal = Decimal.MinValue

    Private ReadOnly _25Slab As Decimal = 2.5
    Private ReadOnly _2Slab As Decimal = 2

    Private ReadOnly _slab As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _slab = slab
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        Dim parameter3 As PlaceOrderParameters = Nothing
        Dim parameter4 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso Not IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload) Then
            If _25SlabProfit = Decimal.MinValue Then
                Dim price As Decimal = currentTick.Open
                _25SlabProfit = _parentStrategy.CalculatePL(_tradingSymbol, price, price + _slab * _25Slab, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                _2SlabProfit = _parentStrategy.CalculatePL(_tradingSymbol, price, price + _slab * _2Slab, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                _oneSlabProfit = _parentStrategy.CalculatePL(_tradingSymbol, price, price + _slab, Me.LotSize, Me.LotSize, _parentStrategy.StockType)
                If _oneSlabProfit <= 0 Then
                    Throw New ApplicationException(String.Format("One Slab Profit PL:{0}", _oneSlabProfit))
                End If

                If Me.MaxProfitOfThisStock = Decimal.MaxValue Then Me.MaxProfitOfThisStock = _25SlabProfit
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 + buffer
                    Dim slPoint As Decimal = _slab + 2 * buffer
                    Dim quantity As Integer = Me.LotSize
                    Dim targetPoint As Decimal = ConvertFloorCeling(_slab * _25Slab, _parentStrategy.TickSize, RoundOfType.Celing)

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = entryPrice - slPoint,
                                .Target = entryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = _slab,
                                .Supporting2 = "Normal",
                                .Supporting3 = 1
                            }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload)
                    If pl < 0 Then
                        quantity = Math.Ceiling((_2SlabProfit + Math.Abs(pl) - _25SlabProfit) / _oneSlabProfit) * Me.LotSize
                        If quantity <> 0 Then
                            Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                            targetPoint = Math.Min(_slab, targetPrice - entryPrice)
                            parameter3 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice - slPoint,
                                    .Target = entryPrice + targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = _slab,
                                    .Supporting2 = "SL Makeup",
                                    .Supporting3 = quantity / Me.LotSize,
                                    .Supporting4 = pl,
                                    .Supporting5 = String.Format("({0}+{1}-{2})/{3}", _2SlabProfit, Math.Abs(pl), _25SlabProfit, _oneSlabProfit)
                                }
                        End If
                    End If
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signal.Item2 - buffer
                    Dim slPoint As Decimal = _slab + 2 * buffer
                    Dim quantity As Integer = Me.LotSize
                    Dim targetPoint As Decimal = ConvertFloorCeling(_slab * _25Slab, _parentStrategy.TickSize, RoundOfType.Celing)

                    parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = entryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = _slab,
                                    .Supporting2 = "Normal",
                                    .Supporting3 = 1
                                }

                    Dim pl As Decimal = GetStockPotentialPL(currentMinuteCandlePayload)
                    If pl < 0 Then
                        quantity = Math.Ceiling((_2SlabProfit + Math.Abs(pl) - _25SlabProfit) / _oneSlabProfit) * Me.LotSize
                        If quantity <> 0 Then
                            Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(pl), Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                            targetPoint = Math.Min(_slab, entryPrice - targetPrice)
                            parameter4 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Math.Abs(quantity),
                                    .Stoploss = entryPrice + slPoint,
                                    .Target = entryPrice - targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = currentMinuteCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = _slab,
                                    .Supporting2 = "SL Makeup",
                                    .Supporting3 = quantity / Me.LotSize,
                                    .Supporting4 = pl,
                                    .Supporting5 = String.Format("({0}+{1}-{2})/{3}", _2SlabProfit, Math.Abs(pl), _25SlabProfit, _oneSlabProfit)
                                }
                        End If
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter2)
        End If
        If parameter3 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter3)
        End If
        If parameter4 IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter4)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastSpecificTrades(currentTick, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If lastExecutedTrade.EntryDirection <> currentTrade.EntryDirection Then
                    If currentTrade.EntryPrice <> lastExecutedTrade.PotentialStopLoss OrElse currentTrade.EntryTime < lastExecutedTrade.EntryTime Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        End If
        If IsAnyTradeOfTheStockTargetReached(currentTick) Then
            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting2 = "Normal" Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim slabLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Sell)
                If currentTick.Open >= slabLevel + currentTrade.StoplossBuffer AndAlso
                    slabLevel - _slab - currentTrade.StoplossBuffer > currentTrade.PotentialStopLoss Then
                    triggerPrice = slabLevel - _slab - currentTrade.StoplossBuffer
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim slabLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Buy)
                If currentTick.Open <= slabLevel - currentTrade.StoplossBuffer AndAlso
                    slabLevel + _slab + currentTrade.StoplossBuffer < currentTrade.PotentialStopLoss Then
                    triggerPrice = slabLevel + _slab + currentTrade.StoplossBuffer
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Moved at {0}", triggerPrice))
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

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastSpecificTrades(candle, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If lastExecutedOrder Is Nothing Then
            Dim buyLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Buy)
            Dim sellLevel As Decimal = GetSlabBasedLevel(currentTick.Open, Trade.TradeExecutionDirection.Sell)
            If candle.High >= buyLevel AndAlso candle.Low <= sellLevel Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = New Tuple(Of Boolean, Decimal)(True, buyLevel)
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    ret = New Tuple(Of Boolean, Decimal)(True, sellLevel)
                End If
            ElseIf candle.High >= buyLevel Then
                If direction = Trade.TradeExecutionDirection.Sell Then
                    ret = New Tuple(Of Boolean, Decimal)(True, sellLevel)
                End If
            ElseIf candle.Low <= sellLevel Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    ret = New Tuple(Of Boolean, Decimal)(True, buyLevel)
                End If
            End If
        Else
            If lastExecutedOrder.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If lastExecutedOrder.EntryDirection <> direction Then
                    If lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        ret = New Tuple(Of Boolean, Decimal)(True, lastExecutedOrder.PotentialStopLoss + lastExecutedOrder.StoplossBuffer)
                    ElseIf lastExecutedOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        ret = New Tuple(Of Boolean, Decimal)(True, lastExecutedOrder.PotentialStopLoss - lastExecutedOrder.StoplossBuffer)
                    End If
                End If
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
    Private Function GetStockPotentialPL(ByVal currentMinutePayload As Payload) As Decimal
        Dim ret As Decimal = 0
        Dim completeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
        If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then
            ret += completeTrades.Sum(Function(x)
                                          Return x.PLAfterBrokerage
                                      End Function)
        End If
        Dim runningTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinutePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Inprogress)
        If runningTrades IsNot Nothing AndAlso runningTrades.Count > 0 Then
            For Each runningTrade In runningTrades
                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, runningTrade.EntryPrice, runningTrade.PotentialStopLoss, runningTrade.Quantity, runningTrade.LotSize, _parentStrategy.StockType)
                ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    ret += _parentStrategy.CalculatePL(runningTrade.TradingSymbol, runningTrade.PotentialStopLoss, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, _parentStrategy.StockType)
                End If
            Next
        End If
        Return ret
    End Function
End Class