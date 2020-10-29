Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class OpeningPriceOptionsATRTraillingStrategyRule
    Inherits StrategyRule

    Private _atrTrailingLowPayload As Dictionary(Of Date, Decimal)
    Private _atrTrailingLowColorPayload As Dictionary(Of Date, Color)
    Private _atrTrailingHighPayload As Dictionary(Of Date, Decimal)
    Private _atrTrailingHighColorPayload As Dictionary(Of Date, Color)
    Private _rsiPayload As Dictionary(Of Date, Decimal)

    Public DummyCandle As Payload = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal controller As Integer,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, controller, canceller)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATRTrailingStop.CalculateATRTrailingStop(14, 5, _signalPayload, _atrTrailingLowPayload, _atrTrailingLowColorPayload)
        Indicator.ATRTrailingStop.CalculateATRTrailingStop(14, 10, _signalPayload, _atrTrailingHighPayload, _atrTrailingHighColorPayload)
        Indicator.RSI.CalculateRSI(4, _signalPayload, _rsiPayload)
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        DummyCandle = currentCandle
        If Not Controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso currentCandle.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
                Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
                _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
                _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
                _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
                _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 Then
                Dim signalCandle As Payload = Nothing
                Dim entryRemark As String = ""
                If _atrTrailingHighColorPayload(currentCandle.PreviousCandlePayload.PayloadDate) = Color.Green AndAlso
                    currentCandle.PreviousCandlePayload.Close > _atrTrailingHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    Dim activeInstruments As List(Of OpeningPriceOptionsATRTraillingStrategyRule) = CType(Me.ControllerInstrument, OpeningPriceOptionsATRTraillingStrategyRule).GetActiveInstruments()
                    If activeInstruments Is Nothing OrElse activeInstruments.Count = 0 Then
                        If _atrTrailingLowColorPayload(currentCandle.PreviousCandlePayload.PayloadDate) = Color.Green Then
                            If currentCandle.PreviousCandlePayload.Close < _atrTrailingLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                                signalCandle = currentCandle.PreviousCandlePayload
                                entryRemark = "ATR Trailing"
                            ElseIf _rsiPayload(currentCandle.PreviousCandlePayload.PayloadDate) <= 10 Then
                                signalCandle = currentCandle.PreviousCandlePayload
                                entryRemark = "RSI"
                            End If
                        End If
                    End If
                End If

                If signalCandle IsNot Nothing Then
                    Dim buffer As Decimal = 0
                    Dim entryPrice As Decimal = currentTick.Open
                    Dim stoploss As Decimal = ConvertFloorCeling(_atrTrailingHighPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim slPoint As Decimal = entryPrice - stoploss
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * 75 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim brkevnPoints As Decimal = _parentStrategy.GetBreakevenPoint(signalCandle.TradingSymbol, entryPrice, Me.LotSize, Trade.TradeExecutionDirection.Buy, Me.LotSize, Trade.TypeOfStock.Futures)
                    If targetPoint > brkevnPoints Then
                        Dim plToAchive As Decimal = 500 - _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate)
                        Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(signalCandle.TradingSymbol, entryPrice, entryPrice + targetPoint, plToAchive, Trade.TypeOfStock.Futures)
                        quantity = Math.Ceiling(quantity / Me.LotSize) * Me.LotSize

                        Dim parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - 1000000000,
                                        .Target = .EntryPrice + targetPoint,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = slPoint,
                                        .Supporting3 = entryRemark
                                    }

                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentCandle.PreviousCandlePayload.Close < _atrTrailingHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Stoploss")
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Public Function GetActiveInstruments() As List(Of OpeningPriceOptionsATRTraillingStrategyRule)
        Dim ret As List(Of OpeningPriceOptionsATRTraillingStrategyRule) = Nothing
        If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            For Each runningInstrument As OpeningPriceOptionsATRTraillingStrategyRule In Me.DependentInstrument
                If runningInstrument.DummyCandle IsNot Nothing AndAlso
                    _parentStrategy.IsTradeActive(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS) Then
                    If ret Is Nothing Then ret = New List(Of OpeningPriceOptionsATRTraillingStrategyRule)
                    ret.Add(runningInstrument)
                End If
            Next
        End If
        Return ret
    End Function
End Class