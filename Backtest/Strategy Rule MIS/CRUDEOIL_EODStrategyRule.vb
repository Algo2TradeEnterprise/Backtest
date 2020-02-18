Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class CRUDEOIL_EODStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public NumberOfLots As Integer
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _stoplossPoint As Decimal = 15
    Private _eodPayload As Dictionary(Of Date, Payload) = Nothing
    Private _minAvg As Decimal = Decimal.MinValue
    Private _maxAvg As Decimal = Decimal.MinValue

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

        _eodPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Commodity, _tradingSymbol, _tradingDate.AddDays(-15), _tradingDate)
        If _eodPayload IsNot Nothing AndAlso _eodPayload.Count > 0 Then
            Dim rulePayload As Dictionary(Of Date, Payload) = Nothing
            Dim ctr As Integer = 0
            For Each runningPayload In _eodPayload.OrderByDescending(Function(x)
                                                                         Return x.Key
                                                                     End Function)
                If runningPayload.Key.Date <> _tradingDate.Date Then
                    If rulePayload Is Nothing Then rulePayload = New Dictionary(Of Date, Payload)
                    rulePayload.Add(runningPayload.Key, runningPayload.Value)

                    ctr += 1
                    If ctr >= 5 Then Exit For
                End If
            Next
            If rulePayload IsNot Nothing AndAlso rulePayload.Count > 0 Then
                Dim highOpen As List(Of Decimal) = Nothing
                For Each runningRulePayload In rulePayload
                    If highOpen Is Nothing Then highOpen = New List(Of Decimal)
                    highOpen.Add(runningRulePayload.Value.High - runningRulePayload.Value.Open)
                Next

                Dim openLow As List(Of Decimal) = Nothing
                For Each runningRulePayload In rulePayload
                    If openLow Is Nothing Then openLow = New List(Of Decimal)
                    openLow.Add(runningRulePayload.Value.Open - runningRulePayload.Value.Low)
                Next

                Dim min As List(Of Decimal) = Nothing
                For i = 0 To openLow.Count - 1
                    If min Is Nothing Then min = New List(Of Decimal)
                    Dim result As Decimal = If(Math.Min(highOpen(i), openLow(i)) < 10, 10, Math.Min(highOpen(i), openLow(i)))
                    min.Add(result)
                Next

                Dim max As List(Of Decimal) = Nothing
                For i = 0 To openLow.Count - 1
                    If max Is Nothing Then max = New List(Of Decimal)
                    max.Add(Math.Max(highOpen(i), openLow(i)))
                Next

                _minAvg = ConvertFloorCeling(min.Average(), _parentStrategy.TickSize, RoundOfType.Floor)
                _maxAvg = ConvertFloorCeling(max.Average(), _parentStrategy.TickSize, RoundOfType.Floor)
            End If
        End If
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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignal(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim quantity As Decimal = Me.LotSize * _userInputs.NumberOfLots
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - _stoplossPoint,
                                .Target = .EntryPrice + _maxAvg,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = _maxAvg,
                                .Supporting2 = _minAvg,
                                .Supporting3 = _eodPayload.LastOrDefault.Value.Open
                            }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim quantity As Decimal = Me.LotSize * _userInputs.NumberOfLots

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + _stoplossPoint,
                                .Target = .EntryPrice - _maxAvg,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = _maxAvg,
                                .Supporting2 = _minAvg,
                                .Supporting3 = _eodPayload.LastOrDefault.Value.Open
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignal(currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item3 <> currentTrade.EntryDirection Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetSignal(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = Nothing
        If _minAvg <> Decimal.MinValue AndAlso _maxAvg <> Decimal.MinValue Then
            Dim currentDayOpen As Decimal = _eodPayload.LastOrDefault.Value.Open
            Dim buyPrice As Decimal = currentDayOpen + _minAvg
            Dim sellPrice As Decimal = currentDayOpen - _minAvg

            Dim range As Decimal = (buyPrice - sellPrice) / 2
            If currentTick.Open > buyPrice - range * 60 / 100 Then
                ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, buyPrice, Trade.TradeExecutionDirection.Buy)
            ElseIf currentTick.Open < sellPrice + range * 60 / 100 Then
                ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, sellPrice, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function
End Class