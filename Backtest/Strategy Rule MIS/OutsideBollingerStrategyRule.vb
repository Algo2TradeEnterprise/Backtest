Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class OutsideBollingerStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumInvestmentPerStock As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private _quantity As Integer
    Private _lastSignalCandle As Payload
    Private _upperBollingerPayloads As Dictionary(Of Date, Decimal)
    Private _lowerBollingerPayloads As Dictionary(Of Date, Decimal)
    Private _smaPayloads As Dictionary(Of Date, Decimal)

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

        Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _signalPayload, _upperBollingerPayloads, _lowerBollingerPayloads, _smaPayloads)
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

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = _lastSignalCandle
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If _quantity = Integer.MinValue Then
                    _quantity = _parentStrategy.CalculateQuantityFromInvestment(1, _userInputs.MinimumInvestmentPerStock, signalCandleSatisfied.Item2, _parentStrategy.StockType, True)
                End If

                If signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signalCandle.High + buffer,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = _quantity,
                                    .Stoploss = signalCandle.Low - buffer,
                                    .Target = .EntryPrice + (.EntryPrice - .Stoploss) * _userInputs.TargetMultiplier,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                ElseIf signalCandleSatisfied.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signalCandle.Low - buffer,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = _quantity,
                                    .Stoploss = signalCandle.High + buffer,
                                    .Target = .EntryPrice - (.Stoploss - .EntryPrice) * _userInputs.TargetMultiplier,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
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
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If (currentTrade.SignalCandle.PayloadDate <> _lastSignalCandle.PayloadDate) OrElse
                    (currentTrade.EntryDirection <> signalCandleSatisfied.Item3) Then
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
            'If candle.Volume > candle.PreviousCandlePayload.Volume Then
            '    Dim firstCandleOfTheDay As Payload = GetFirstCandleOfTheDay(candle.PayloadDate.Date)
            '    If firstCandleOfTheDay IsNot Nothing Then
            '        Dim firstCandleOfThePreviousDay As Payload = GetFirstCandleOfTheDay(firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date)
            '        If firstCandleOfThePreviousDay IsNot Nothing AndAlso firstCandleOfTheDay.Volume > firstCandleOfThePreviousDay.Volume Then
            '            Dim previous20Payoads As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(_signalPayload, candle.PayloadDate, 20, False)
            '            If previous20Payoads IsNot Nothing AndAlso previous20Payoads.Count > 0 Then
            '                Dim maxVolume As Long = previous20Payoads.Max(Function(x)
            '                                                                  Return x.Value.Volume
            '                                                              End Function)
            '                If candle.Volume > maxVolume Then
            If candle.Low > _upperBollingerPayloads(candle.PayloadDate) AndAlso
                                    candle.PreviousCandlePayload.Low > _upperBollingerPayloads(candle.PreviousCandlePayload.PayloadDate) Then
                _lastSignalCandle = candle
            ElseIf candle.High < _lowerBollingerPayloads(candle.PayloadDate) AndAlso
                candle.PreviousCandlePayload.High < _lowerBollingerPayloads(candle.PreviousCandlePayload.PayloadDate) Then
                _lastSignalCandle = candle
            End If
            '                End If
            '            End If
            '        End If
            '    End If
            'End If

            If _lastSignalCandle IsNot Nothing Then
                Dim buyPrice As Decimal = _lastSignalCandle.High
                Dim sellPrice As Decimal = _lastSignalCandle.Low
                Dim middle As Decimal = (buyPrice + sellPrice) / 2
                Dim range As Decimal = (buyPrice - middle) * 30 / 100
                If currentTick.Open > middle + range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, buyPrice, Trade.TradeExecutionDirection.Buy)
                ElseIf currentTick.Open < middle - range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection)(True, sellPrice, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetFirstCandleOfTheDay(ByVal day As Date) As Payload
        Dim ret As Payload = Nothing
        Dim requiredPayloads As IEnumerable(Of Payload) = _signalPayload.Values.Where(Function(x)
                                                                                          Return x.PayloadDate.Date = day.Date
                                                                                      End Function)
        If requiredPayloads IsNot Nothing AndAlso requiredPayloads.Count > 0 Then
            ret = requiredPayloads.OrderBy(Function(x)
                                               Return x.PayloadDate
                                           End Function).FirstOrDefault
        End If
        Return ret
    End Function
End Class