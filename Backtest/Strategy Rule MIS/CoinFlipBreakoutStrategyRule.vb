Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class CoinFlipBreakoutStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPercentage As Decimal
        Public TargetPercentage As Decimal
        Public MaxStoplossPerTrade
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim lastTradeEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)

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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso currentTick.PayloadDate <= lastTradeEntryTime AndAlso Me.EligibleToTakeTrade Then

            _direction = GetTradeDirectionForEntry()
            If _direction = Trade.TradeExecutionDirection.Buy Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentMinuteCandlePayload.PreviousCandlePayload.High, RoundOfType.Floor)
                Dim entryPrice As Decimal = currentMinuteCandlePayload.PreviousCandlePayload.High + buffer
                If currentTick.Open >= entryPrice Then
                    Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim targetPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.TargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxStoplossPerTrade, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + slPoint,
                                .Target = .EntryPrice - targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = _direction.ToString
                            }
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentMinuteCandlePayload.PreviousCandlePayload.Low, RoundOfType.Floor)
                Dim entryPrice As Decimal = currentMinuteCandlePayload.PreviousCandlePayload.Low - buffer
                If currentTick.Open <= entryPrice Then
                    Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim targetPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.TargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxStoplossPerTrade, _parentStrategy.StockType)

                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - slPoint,
                            .Target = .EntryPrice + targetPoint,
                            .Buffer = buffer,
                            .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss"),
                            .Supporting2 = _direction.ToString
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
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetTradeDirectionForEntry() As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        Dim r As Random = New Random
        Dim direction As Integer = r.Next(0, 2)
        If direction = 0 Then
            ret = Trade.TradeExecutionDirection.Sell
        ElseIf direction = 1 Then
            ret = Trade.TradeExecutionDirection.Buy
        End If
        Return ret
    End Function

End Class