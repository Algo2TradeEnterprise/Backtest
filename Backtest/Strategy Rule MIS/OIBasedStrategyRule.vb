Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class OIBasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxStoplossPerStock As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region


    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _oiDetails As List(Of OIData)
    Private ReadOnly _slab As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal,
                   ByVal oiDetails As List(Of OIData))
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _slab = slab
        _oiDetails = oiDetails
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
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, _parentStrategy.TradeType) AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                ElseIf lastExecutedTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3
                    Dim targetPrice As Decimal = signalCandleSatisfied.Item4
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, slPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                ElseIf signalCandleSatisfied.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3
                    Dim targetPrice As Decimal = signalCandleSatisfied.Item4
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, slPrice, entryPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = targetPrice,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    If currentTrade.EntryDirection <> signalCandleSatisfied.Item5 Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
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

    Private _potentialHighEntryPrice As Decimal = Decimal.MinValue
    Private _potentialLowEntryPrice As Decimal = Decimal.MinValue

    Private Function GetSignalCandle(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing Then
            If _potentialHighEntryPrice = Decimal.MinValue AndAlso _potentialLowEntryPrice = Decimal.MinValue Then
                If _oiDetails IsNot Nothing AndAlso _oiDetails.Count > 0 Then
                    Dim availableOIData As OIData = _oiDetails.Find(Function(x)
                                                                        Return x.Time = currentCandle.PayloadDate
                                                                    End Function)
                    If currentCandle.PayloadDate.Date <> _tradingDate.Date Then
                        availableOIData = _oiDetails.Find(Function(x)
                                                              Return x.Time <= currentTick.PayloadDate
                                                          End Function)
                    End If
                    If availableOIData IsNot Nothing Then
                        If availableOIData.CPCOIChange >= 1000 OrElse availableOIData.CPCOIChange <= -1000 OrElse
                            availableOIData.PCCOIChange >= 1000 OrElse availableOIData.PCCOIChange <= -1000 Then
                            If availableOIData.CPCOI >= 100 OrElse availableOIData.CPCOI <= -100 OrElse
                                availableOIData.PCCOI >= 100 OrElse availableOIData.PCCOI <= -100 Then
                                _potentialHighEntryPrice = GetSlabBasedLevel(currentCandle.Close, Trade.TradeExecutionDirection.Buy)
                                _potentialLowEntryPrice = GetSlabBasedLevel(currentCandle.Close, Trade.TradeExecutionDirection.Sell)
                            End If
                        End If
                    End If
                End If
            End If
            If _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
                Dim middlePoint As Decimal = (_potentialHighEntryPrice + _potentialLowEntryPrice) / 2
                Dim range As Decimal = _potentialHighEntryPrice - middlePoint
                If currentTick.Open >= middlePoint + range * 60 / 100 Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialHighEntryPrice, ConvertFloorCeling(middlePoint, _parentStrategy.TickSize, RoundOfType.Floor), _potentialHighEntryPrice + _slab * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy)
                ElseIf currentTick.Open <= middlePoint - range * 60 / 100 Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialLowEntryPrice, ConvertFloorCeling(middlePoint, _parentStrategy.TickSize, RoundOfType.Celing), _potentialLowEntryPrice - _slab * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
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

    Public Class OIData
        Public Time As Date
        Public SumOfPutsOI As Long
        Public SumOfCallsOI As Long
        Public CPCOI As Decimal
        Public PCCOI As Decimal
        Public CTROI As Decimal
        Public PTROI As Decimal
        Public SumOfPutsOIChange As Long
        Public SumOfCallsOIChange As Long
        Public CPCOIChange As Decimal
        Public PCCOIChange As Decimal
        Public CTROIChange As Decimal
        Public PTROIChange As Decimal
    End Class
End Class