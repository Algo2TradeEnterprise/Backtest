Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKForwardBreakoutAdaptiveMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
        Public MaxCapitalPerTrade As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastOrder.ExitTime, _signalPayload))
                    If signal.Item4.PayloadDate >= exitCandle.PayloadDate Then
                        signalCandle = signal.Item4
                    End If
                Else
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item3

                Dim atr As Decimal = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim iteration As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - atr, lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, signal.Item5, Trade.TypeOfStock.Cash)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = signal.Item5,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = iteration,
                                                            .Supporting3 = atr
                                                        }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                currentCandle.PreviousCandlePayload.Close < currentTrade.SignalCandle.Low Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                currentCandle.PreviousCandlePayload.Close > currentTrade.SignalCandle.High Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            Else
                Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.EntryPrice <> signal.Item2 OrElse
                        currentTrade.PotentialStopLoss <> signal.Item3 OrElse
                        currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate OrElse
                        currentTrade.EntryDirection <> signal.Item5 Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            If Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.Low, 2) Then
                Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel = Math.Round(hkCandle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                    buyLevel = buyLevel + buffer
                End If
                Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If sellLevel = Math.Round(hkCandle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                    sellLevel = sellLevel - buffer
                End If
                Dim atr As Decimal = ConvertFloorCeling(_atrPayload(hkCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel - sellLevel <= atr Then
                    Dim initQty As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, buyLevel, buyLevel - atr, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim requiredCapital As Decimal = buyLevel * initQty / _parentStrategy.MarginMultiplier
                    If requiredCapital <= _userInputs.MaxCapitalPerTrade Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, sellLevel, hkCandle, Trade.TradeExecutionDirection.Buy)
                    Else
                        'Console.WriteLine(String.Format("{0}: Neglected for higher capital,{1},{2}", _tradingSymbol, hkCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), requiredCapital))
                    End If
                End If
            ElseIf Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.High, 2) Then
                Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If sellLevel = Math.Round(hkCandle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(sellLevel, RoundOfType.Floor)
                    sellLevel = sellLevel - buffer
                End If
                Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel = Math.Round(hkCandle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(buyLevel, RoundOfType.Floor)
                    buyLevel = buyLevel + buffer
                End If
                Dim atr As Decimal = ConvertFloorCeling(_atrPayload(hkCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel - sellLevel <= atr Then
                    Dim initQty As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, sellLevel + atr, sellLevel, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    Dim requiredCapital As Decimal = sellLevel * initQty / _parentStrategy.MarginMultiplier
                    If requiredCapital <= _userInputs.MaxCapitalPerTrade Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, buyLevel, hkCandle, Trade.TradeExecutionDirection.Sell)
                    Else
                        'Console.WriteLine(String.Format("{0}: Neglected for higher capital,{1},{2}", _tradingSymbol, hkCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), requiredCapital))
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class