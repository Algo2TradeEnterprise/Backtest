Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LimitBuyBelowFractalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerStock As Decimal
        Public AdjustTarget As Boolean
        Public AdjustTargetForTurnover As Decimal
        Public QuantityFromFractalLow As Boolean
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _firstCandleOfTheDay As Payload = Nothing

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

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                _firstCandleOfTheDay = runningPayload.Value
                Exit For
            End If
        Next

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim fractalLow As Decimal = _fractalLowPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
            Dim fractalHigh As Decimal = _fractalHighPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
            If _firstCandleOfTheDay IsNot Nothing Then
                If _fractalLowPayload(_firstCandleOfTheDay.PreviousCandlePayload.PayloadDate) = fractalLow Then
                    fractalLow = Decimal.MinValue
                Else
                    _firstCandleOfTheDay = Nothing
                End If
            End If

            If fractalLow <> Decimal.MinValue AndAlso currentMinuteCandle.PreviousCandlePayload.Close < fractalLow AndAlso fractalHigh > fractalLow Then
                If _signalCandle Is Nothing Then
                    _signalCandle = currentMinuteCandle.PreviousCandlePayload
                End If
            End If
            If _signalCandle IsNot Nothing Then
                If _fractalLowPayload(_signalCandle.PayloadDate) <> fractalLow Then
                    _signalCandle = Nothing
                End If
            End If
            Dim signalCandle As Payload = Nothing
            If _signalCandle IsNot Nothing AndAlso currentTick.Open <= _signalCandle.Low Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing Then
                    If lastExecutedOrder.Supporting2 <> fractalLow AndAlso lastExecutedOrder.Supporting3 <> fractalHigh Then
                        signalCandle = _signalCandle
                    End If
                Else
                    signalCandle = _signalCandle
                End If
                _signalCandle = Nothing
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = currentTick.Open
                If _userInputs.QuantityFromFractalLow Then entryPrice = fractalLow
                Dim targetPrice As Decimal = fractalHigh
                Dim quantity As Integer = CalculateQuantity(currentMinuteCandle, entryPrice, targetPrice)
                If _userInputs.AdjustTarget AndAlso quantity * entryPrice > _userInputs.AdjustTargetForTurnover Then
                    While quantity * entryPrice > _userInputs.AdjustTargetForTurnover
                        targetPrice += _parentStrategy.TickSize
                        quantity = CalculateQuantity(currentMinuteCandle, entryPrice, targetPrice)
                    End While
                End If
                entryPrice = currentTick.Open
                If quantity > 0 Then
                    parameter = New PlaceOrderParameters With {
                                    .entryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .quantity = quantity,
                                    .Stoploss = 0,
                                    .Target = targetPrice,
                                    .Buffer = 0,
                                    .signalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = fractalLow,
                                    .Supporting3 = fractalHigh,
                                    .Supporting4 = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
            Dim price As Decimal = CalculateTargetPrice(currentTick, lastExecutedTrade.EntryPrice)
            If currentTrade.PotentialTarget <> price Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, price, price - currentTrade.EntryPrice)
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function CalculateQuantity(ByVal candle As Payload, ByVal entryPrice As Decimal, ByVal targetPrice As Decimal) As Integer
        Dim ret As Integer = 0
        Dim unrealizedPL As Decimal = 0
        Dim inProgressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        If inProgressTrades IsNot Nothing AndAlso inProgressTrades.Count > 0 Then
            For Each runningTrade In inProgressTrades
                unrealizedPL += _parentStrategy.CalculatePL(_tradingSymbol, runningTrade.EntryPrice, targetPrice, runningTrade.Quantity, runningTrade.LotSize, Trade.TypeOfStock.Futures)
            Next
        End If
        Dim plToAchive As Decimal = _userInputs.MaxProfitPerStock - unrealizedPL
        If plToAchive > 0 Then
            Dim qty As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, targetPrice, plToAchive, Trade.TypeOfStock.Futures)
            ret = Math.Ceiling(qty / Me.LotSize) * Me.LotSize
        End If
        Return ret
    End Function

    Private Function CalculateTargetPrice(ByVal candle As Payload, ByVal entryPrice As Decimal) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim inProgressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        If inProgressTrades IsNot Nothing AndAlso inProgressTrades.Count > 0 Then
            For target As Decimal = entryPrice To Decimal.MaxValue Step _parentStrategy.TickSize
                Dim totalPl As Decimal = 0
                For Each runningTrade In inProgressTrades
                    totalPl += _parentStrategy.CalculatePL(_tradingSymbol, runningTrade.EntryPrice, target, runningTrade.Quantity, runningTrade.LotSize, Trade.TypeOfStock.Futures)
                Next

                If totalPl >= _userInputs.MaxProfitPerStock Then
                    ret = target
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
End Class