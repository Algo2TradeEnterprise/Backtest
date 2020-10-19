Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MADirectionBasedSMIEntryStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
        Public SignalTimeframe As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _smiPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _xMinutePayload As Dictionary(Of Date, Payload) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.SMI.CalculateSMI(10, 3, 3, 10, _hkPayload, _smiPayload, Nothing)

        Dim exchangeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)
        _xMinutePayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _userInputs.SignalTimeframe, exchangeStartTime)
        Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Close, _xMinutePayload, _smaPayload)

        Dim eodPayload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Cash, _tradingSymbol, _tradingDate.AddDays(-150), _tradingDate)
        If eodPayload IsNot Nothing AndAlso eodPayload.Count > 0 Then
            Dim eodSMAPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Close, eodPayload, eodSMAPayload)
            Dim currentDayCandle As Payload = eodPayload(_tradingDate.Date)
            If currentDayCandle.PreviousCandlePayload.Low > eodSMAPayload(currentDayCandle.PreviousCandlePayload.PayloadDate) Then
                _direction = Trade.TradeExecutionDirection.Buy
            ElseIf currentDayCandle.PreviousCandlePayload.High < eodSMAPayload(currentDayCandle.PreviousCandlePayload.PayloadDate) Then
                _direction = Trade.TradeExecutionDirection.Sell
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastOrder.ExitTime, _signalPayload))
                    If currentMinuteCandle.PayloadDate > exitCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                Else
                    signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim slPoint As Decimal = ConvertFloorCeling(_atrPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item2 - slPoint
                If signal.Item4 = Trade.TradeExecutionDirection.Sell Then stoploss = signal.Item2 + slPoint

                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, _userInputs.MaxProfitPerTrade, signal.Item4, Trade.TypeOfStock.Cash)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = signal.Item4,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
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
            Dim signalCandle As Payload = _xMinutePayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _xMinutePayload, _userInputs.SignalTimeframe)).PreviousCandlePayload
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso signalCandle.Low > _smaPayload(signalCandle.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso signalCandle.High < _smaPayload(signalCandle.PayloadDate) Then
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
            If ret Is Nothing Then
                Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            Dim signalCandle As Payload = _xMinutePayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _xMinutePayload, _userInputs.SignalTimeframe)).PreviousCandlePayload
            If signalCandle.PayloadDate.Date = _tradingDate.Date Then
                If _direction = Trade.TradeExecutionDirection.Buy AndAlso signalCandle.Low > _smaPayload(signalCandle.PayloadDate) Then
                    If _smiPayload(hkCandle.PayloadDate) <= -40 Then
                        If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                            Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                            If entryPrice = Math.Round(hkCandle.High, 2) Then
                                entryPrice = entryPrice + _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                            End If
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, _direction)
                        End If
                    End If
                ElseIf _direction = Trade.TradeExecutionDirection.Sell AndAlso signalCandle.High < _smaPayload(signalCandle.PayloadDate) Then
                    If _smiPayload(hkCandle.PayloadDate) >= 40 Then
                        If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                            Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                            If entryPrice = Math.Round(hkCandle.Low, 2) Then
                                entryPrice = entryPrice - _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                            End If
                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, _direction)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class