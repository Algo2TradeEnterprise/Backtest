Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowStoplossWickStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumInvestmentPerStock As Decimal
        Public MinStoploss As Decimal
        Public MaxStoploss As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private _userInputs As StrategyRuleEntities
    Private _quantity As Integer

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _quantity = Integer.MinValue
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
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
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If _quantity = Integer.MinValue Then
                _quantity = _parentStrategy.CalculateQuantityFromInvestment(_lotSize, _userInputs.MinimumInvestmentPerStock, currentMinuteCandlePayload.PreviousCandlePayload.Close, _parentStrategy.StockType, True)
            End If

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                Dim targetPoint As Decimal = signalCandleSatisfied.Item2
                Dim targetRemark As Decimal = "Original Target"
                If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, _quantity, Math.Abs(lastExecutedTrade.PLAfterBrokerage), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                    targetPoint = targetPrice - signalCandle.Open
                    targetRemark = "SL Makeup Trade"
                End If
                If signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim slPrice As Decimal = Decimal.MinValue
                    If signalCandle.CandleColor = Color.Red Then
                        slPrice = signalCandle.Open
                    Else
                        slPrice = signalCandle.Close
                    End If
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.High + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _quantity,
                        .Stoploss = slPrice,
                        .Target = .EntryPrice + targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                        .Supporting2 = targetRemark,
                        .Supporting3 = targetPoint,
                        .Supporting4 = (.EntryPrice - .Stoploss)
                    }
                ElseIf signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim slPrice As Decimal = Decimal.MinValue
                    If signalCandle.CandleColor = Color.Green Then
                        slPrice = signalCandle.Open
                    Else
                        slPrice = signalCandle.Close
                    End If
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.Low - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _quantity,
                        .Stoploss = slPrice,
                        .Target = .EntryPrice - targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                        .Supporting2 = targetRemark,
                        .Supporting3 = targetPoint,
                        .Supporting4 = (.Stoploss - .EntryPrice)
                    }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            'If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
            '    Me.MaxProfitOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice + parameter.EntryPrice * _parentStrategy.StockMaxProfitPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            'End If
            'If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
            '    Me.MaxLossOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice - parameter.EntryPrice * _parentStrategy.StockMaxLossPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            'End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            Dim firstDirectionToCheck As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            If candle.CandleColor = Color.Green Then
                firstDirectionToCheck = Trade.TradeExecutionDirection.Buy
            ElseIf candle.CandleColor = Color.Red Then
                firstDirectionToCheck = Trade.TradeExecutionDirection.Sell
            Else
                firstDirectionToCheck = Trade.TradeExecutionDirection.Buy
            End If
            Dim buySLPrice As Decimal = GetStoplossPrice(candle, Trade.TradeExecutionDirection.Buy)
            Dim sellSLPrice As Decimal = GetStoplossPrice(candle, Trade.TradeExecutionDirection.Sell)
            If firstDirectionToCheck = Trade.TradeExecutionDirection.Buy Then
                If Math.Abs(buySLPrice) >= _userInputs.MinStoploss AndAlso Math.Abs(buySLPrice) <= _userInputs.MaxStoploss Then
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, candle.High, _quantity, Math.Abs(buySLPrice) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, targetPrice - candle.High, Trade.TradeExecutionDirection.Buy)
                ElseIf Math.Abs(sellSLPrice) >= _userInputs.MinStoploss AndAlso Math.Abs(sellSLPrice) <= _userInputs.MaxStoploss Then
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, candle.Low, _quantity, Math.Abs(sellSLPrice) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, candle.Low - targetPrice, Trade.TradeExecutionDirection.Sell)
                End If
            ElseIf firstDirectionToCheck = Trade.TradeExecutionDirection.Sell Then
                If Math.Abs(sellSLPrice) >= _userInputs.MinStoploss AndAlso Math.Abs(sellSLPrice) <= _userInputs.MaxStoploss Then
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, candle.Low, _quantity, Math.Abs(sellSLPrice) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, candle.Low - targetPrice, Trade.TradeExecutionDirection.Sell)
                ElseIf Math.Abs(buySLPrice) >= _userInputs.MinStoploss AndAlso Math.Abs(buySLPrice) <= _userInputs.MaxStoploss Then
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, candle.High, _quantity, Math.Abs(buySLPrice) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, targetPrice - candle.High, Trade.TradeExecutionDirection.Buy)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetStoplossPrice(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
            Dim slPoint As Decimal = candle.CandleWicks.Top + buffer
            ret = _parentStrategy.CalculatePL(_tradingSymbol, candle.High, candle.High - slPoint, _quantity, _lotSize, _parentStrategy.StockType)
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
            Dim slPoint As Decimal = candle.CandleWicks.Bottom + buffer
            ret = _parentStrategy.CalculatePL(_tradingSymbol, candle.Low + slPoint, candle.Low, _quantity, _lotSize, _parentStrategy.StockType)
        End If
        Return ret
    End Function
End Class