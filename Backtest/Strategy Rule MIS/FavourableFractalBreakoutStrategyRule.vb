Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FavourableFractalBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public MaxProfitPerTrade As Decimal
    End Class

    Public Class StockData
        Public InstrumentName As String
        Public LotSize As Integer
        Public ChangePer As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim triggerPrice As Decimal = signal.Item2 + buffer
                    Dim atr As Decimal = _atrPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim stoplossPrice As Decimal = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - buffer
                    Dim slPoint As Decimal = Math.Max(atr, triggerPrice - stoplossPrice)
                    stoplossPrice = ConvertFloorCeling(triggerPrice - slPoint, _parentStrategy.TickSize, RoundOfType.Floor)
                    If ((triggerPrice - stoplossPrice) / stoplossPrice) * 100 > 1 Then
                        stoplossPrice = ConvertFloorCeling(triggerPrice - triggerPrice * 1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, triggerPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                    Dim targetPrice As Decimal = Decimal.MaxValue
                    If _userInputs.MaxProfitPerTrade = Decimal.MaxValue Then
                        targetPrice = triggerPrice + 1000000
                    Else
                        targetPrice = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, triggerPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                    End If

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss")
                                }
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim triggerPrice As Decimal = signal.Item2 - buffer
                    Dim atr As Decimal = _atrPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim stoplossPrice As Decimal = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - buffer
                    Dim slPoint As Decimal = Math.Max(atr, stoplossPrice - triggerPrice)
                    stoplossPrice = ConvertFloorCeling(triggerPrice + slPoint, _parentStrategy.TickSize, RoundOfType.Floor)
                    If ((stoplossPrice - triggerPrice) / triggerPrice) * 100 > 1 Then
                        stoplossPrice = ConvertFloorCeling(triggerPrice + triggerPrice * 1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoplossPrice, triggerPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                    Dim targetPrice As Decimal = Decimal.MinValue
                    If _userInputs.MaxProfitPerTrade = Decimal.MaxValue Then
                        targetPrice = triggerPrice - 1000000
                    Else
                        targetPrice = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, triggerPrice, quantity, _userInputs.MaxProfitPerTrade, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                    End If

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = triggerPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss")
                                }
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signal As Tuple(Of Boolean, Decimal) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
                If signal IsNot Nothing Then
                    If signalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                If _fractalHighPayload(candle.PayloadDate) < _fractalHighPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal)(True, _fractalHighPayload(candle.PayloadDate))
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If _fractalLowPayload(candle.PayloadDate) > _fractalLowPayload(candle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal)(True, _fractalLowPayload(candle.PayloadDate))
                End If
            End If
        End If
        Return ret
    End Function

    Public Shared Function GetStockData(ByVal allStocks As List(Of StockData), ByVal numberOfStock As Integer) As Dictionary(Of String, StockDetails)
        Dim ret As Dictionary(Of String, StockDetails) = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            For i As Integer = 1 To Math.Ceiling(numberOfStock / 2)
                Dim stock As StockData = Nothing
                stock = GetNthStockFromTop(allStocks, i)
                If stock IsNot Nothing Then
                    If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = stock.InstrumentName,
                                                 .LotSize = stock.LotSize,
                                                 .EligibleToTakeTrade = True}
                    ret.Add(detailsOfStock.StockName, detailsOfStock)
                End If
            Next
            For i As Integer = 1 To Math.Floor(numberOfStock / 2)
                Dim stock As StockData = Nothing
                stock = GetNthStockFromBottom(allStocks, i)
                If stock IsNot Nothing Then
                    If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = stock.InstrumentName,
                                                 .LotSize = stock.LotSize,
                                                 .EligibleToTakeTrade = True}
                    ret.Add(detailsOfStock.StockName, detailsOfStock)
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetNthStockFromTop(ByVal allStocks As List(Of StockData), ByVal stockNumber As Integer) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            For Each runningStock In allStocks.OrderByDescending(Function(x)
                                                                     Return x.ChangePer
                                                                 End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetNthStockFromBottom(ByVal allStocks As List(Of StockData), ByVal stockNumber As Integer) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            For Each runningStock In allStocks.OrderBy(Function(x)
                                                           Return x.ChangePer
                                                       End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
End Class