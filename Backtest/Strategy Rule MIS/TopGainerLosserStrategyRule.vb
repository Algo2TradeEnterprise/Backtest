Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TopGainerLosserStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetPercentage As Decimal
        Public StoplossPercentage As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Decimal
    Private ReadOnly _time As Date

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Decimal,
                   ByVal time As Date)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _direction = direction
        _time = time.AddMinutes(1)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _time.Hour, _time.Minute, _time.Second)

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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = currentMinuteCandlePayload
            If signalCandle IsNot Nothing Then
                Dim slPoint As Decimal = ConvertFloorCeling(currentTick.Open * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                Dim targetPoint As Decimal = ConvertFloorCeling(currentTick.Open * _userInputs.TargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                If _direction < 0 Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = currentTick.Open
                    Dim stoploss As Decimal = entry - slPoint
                    Dim target As Decimal = entry + targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market
                                }
                ElseIf _direction > 0 Then
                    Dim buffer As Decimal = 0
                    Dim entry As Decimal = currentTick.Open
                    Dim stoploss As Decimal = entry + slPoint
                    Dim target As Decimal = entry - targetPoint
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss")
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
End Class