Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MultiIndicatorStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public RSIOverBought As Decimal
        Public RSIOverSold As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _rsiPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _psarPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _psarTrendPayload As Dictionary(Of Date, Color) = Nothing
    Private _aroonHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _aroonLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _macdPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _macdSignalPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _macdHistogramPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Close, _signalPayload, _smaPayload)
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)
        Indicator.RSI.CalculateRSI(14, _signalPayload, _rsiPayload)
        Indicator.ParabolicSAR.CalculatePSAR(0.02, 0.2, _signalPayload, _psarPayload, _psarTrendPayload)
        Indicator.AROON.CalculateAROON(14, _signalPayload, _aroonHighPayload, _aroonLowPayload)
        Indicator.MACD.CalculateMACD(12, 26, 9, _signalPayload, _macdPayload, _macdSignalPayload, _macdHistogramPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item2
                ElseIf lastExecutedOrder IsNot Nothing Then
                    Dim lastOrderExitTime As Date = _parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload)
                    If signal.Item2.PayloadDate >= lastOrderExitTime Then
                        signalCandle = signal.Item2
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryBuffer As Decimal = CalculateEntryBuffer(signalCandle.High)
                    Dim entryPrice As Decimal = signalCandle.High + entryBuffer
                    Dim stoplossBuffer As Decimal = CalculateStoplossBuffer(signalCandle.Low)
                    Dim stoploss As Decimal = signalCandle.Low
                    Dim slRemark As String = "Candle Low"
                    Dim sma As Decimal = ConvertFloorCeling(_smaPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    If sma < stoploss Then
                        stoploss = sma
                        slRemark = "SMA"
                    End If
                    stoploss = stoploss - stoplossBuffer
                    Dim slPoint As Decimal = entryPrice - stoploss
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim target As Decimal = entryPrice + targetPoint
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_instrumentName, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    If _parentStrategy.StockType = Trade.TypeOfStock.Commodity Then quantity = Me.LotSize

                    parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .EntryBuffer = entryBuffer,
                                        .StoplossBuffer = stoplossBuffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = slRemark,
                                        .Supporting3 = sma,
                                        .Supporting4 = GetIndicatorString(signalCandle)
                                    }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryBuffer As Decimal = CalculateEntryBuffer(signalCandle.Low)
                    Dim entryPrice As Decimal = signalCandle.Low - entryBuffer
                    Dim stoplossBuffer As Decimal = CalculateStoplossBuffer(signalCandle.High)
                    Dim stoploss As Decimal = signalCandle.High
                    Dim slRemark As String = "Candle High"
                    Dim sma As Decimal = ConvertFloorCeling(_smaPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    If sma > stoploss Then
                        stoploss = sma
                        slRemark = "SMA"
                    End If
                    stoploss = stoploss + stoplossBuffer
                    Dim slPoint As Decimal = stoploss - entryPrice
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim target As Decimal = entryPrice - targetPoint
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_instrumentName, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    If _parentStrategy.StockType = Trade.TypeOfStock.Commodity Then quantity = Me.LotSize

                    parameter = New PlaceOrderParameters With {
                                        .EntryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .Quantity = quantity,
                                        .Stoploss = stoploss,
                                        .Target = target,
                                        .EntryBuffer = entryBuffer,
                                        .StoplossBuffer = stoplossBuffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                        .Supporting2 = slRemark,
                                        .Supporting3 = sma,
                                        .Supporting4 = GetIndicatorString(signalCandle)
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
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing Then
                Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing Then
                    If signal.Item1 Then
                        If currentTrade.EntryDirection <> signal.Item3 OrElse currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
                    Else
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                Else
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then

        End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            If _rsiPayload(candle.PayloadDate) > _userInputs.RSIOverBought Then
                If _macdPayload(candle.PayloadDate) > _macdSignalPayload(candle.PayloadDate) Then
                    If _aroonHighPayload(candle.PayloadDate) > _aroonLowPayload(candle.PayloadDate) Then
                        If _psarTrendPayload(candle.PayloadDate) = Color.Green Then
                            If candle.Close > _vwapPayload(candle.PayloadDate) Then
                                If candle.Close > _smaPayload(candle.PayloadDate) Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Buy)
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf _rsiPayload(candle.PayloadDate) < _userInputs.RSIOverSold Then
                If _macdPayload(candle.PayloadDate) < _macdSignalPayload(candle.PayloadDate) Then
                    If _aroonHighPayload(candle.PayloadDate) < _aroonLowPayload(candle.PayloadDate) Then
                        If _psarTrendPayload(candle.PayloadDate) = Color.Red Then
                            If candle.Close < _vwapPayload(candle.PayloadDate) Then
                                If candle.Close < _smaPayload(candle.PayloadDate) Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, candle, Trade.TradeExecutionDirection.Sell)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetIndicatorString(ByVal signalCandle As Payload) As String
        Dim ret As String = Nothing
        ret = String.Format("SMA:{0},VWAP:{1},RSI:{2},PSAR:{3},PSAR Trend:{4},AROON High:{5},AROON Low:{6},MACD:{7},MACD Signal:{8},MACD Histogram:{9}",
                            Math.Round(_smaPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_vwapPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_rsiPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_psarPayload(signalCandle.PayloadDate), 4),
                            _psarTrendPayload(signalCandle.PayloadDate).Name,
                            Math.Round(_aroonHighPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_aroonLowPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_macdPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_macdSignalPayload(signalCandle.PayloadDate), 4),
                            Math.Round(_macdHistogramPayload(signalCandle.PayloadDate), 4))
        Return ret
    End Function

    Private Function CalculateEntryBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = 0
        If _parentStrategy.StockType = Trade.TypeOfStock.Commodity Then
            If _tradingSymbol.Contains("CRUDE") Then
                ret = 2
            ElseIf _tradingSymbol.Contains("SILVER") Then
                ret = 5
            ElseIf _tradingSymbol.Contains("NATURALGAS") Then
                ret = 0.1
            ElseIf _tradingSymbol.Contains("COPPER") Then
                ret = 0.1
            ElseIf _tradingSymbol.Contains("NICKEL") Then
                ret = 0.2
            ElseIf _tradingSymbol.Contains("ZINC") Then
                ret = 0.1
            ElseIf _tradingSymbol.Contains("LEAD") Then
                ret = 0.1
            End If
        Else
            If price <= 400 Then
                ret = 0.1
            ElseIf price <= 800 Then
                ret = 0.2
            ElseIf price <= 1200 Then
                ret = 0.3
            ElseIf price <= 1600 Then
                ret = 0.4
            ElseIf price <= 2000 Then
                ret = 0.5
            ElseIf price <= 2400 Then
                ret = 0.6
            Else
                ret = 0.7
            End If
        End If
        Return ret
    End Function

    Private Function CalculateStoplossBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = 0
        If _parentStrategy.StockType = Trade.TypeOfStock.Commodity Then
            If _tradingSymbol.Contains("CRUDE") Then
                ret = 2
            ElseIf _tradingSymbol.Contains("SILVER") Then
                ret = 5
            ElseIf _tradingSymbol.Contains("NATURALGAS") Then
                ret = 0.2
            ElseIf _tradingSymbol.Contains("COPPER") Then
                ret = 0.2
            ElseIf _tradingSymbol.Contains("NICKEL") Then
                ret = 0.2
            ElseIf _tradingSymbol.Contains("ZINC") Then
                ret = 0.2
            ElseIf _tradingSymbol.Contains("LEAD") Then
                ret = 0.2
            End If
        Else
            If price <= 400 Then
                ret = 0.1
            ElseIf price <= 800 Then
                ret = 0.2
            ElseIf price <= 1200 Then
                ret = 0.3
            ElseIf price <= 1600 Then
                ret = 0.4
            ElseIf price <= 2000 Then
                ret = 0.5
            ElseIf price <= 2400 Then
                ret = 0.6
            Else
                ret = 0.7
            End If
        End If
        Return ret
    End Function
End Class