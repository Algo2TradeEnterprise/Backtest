Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class AjitJhaORBOptionStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxStoplossPoint As Decimal
        Public MinStoplossPoint As Decimal
        Public StoplossPercentage As Decimal
    End Class
#End Region

    Public Remarks As String = Nothing

    Private _buyLevel As Decimal = Decimal.MaxValue
    Private _sellLevel As Decimal = Decimal.MinValue
    Private _swingHighPayload As Dictionary(Of Date, Payload) = Nothing
    Private _swingLowPayload As Dictionary(Of Date, Payload) = Nothing

    Private _buyORBTriggered As Boolean = False
    Private _sellORBTriggered As Boolean = False

    Private _peOptionStrikeList As Dictionary(Of Decimal, AjitJhaORBOptionStrategyRule) = Nothing
    Private _ceOptionStrikeList As Dictionary(Of Decimal, AjitJhaORBOptionStrategyRule) = Nothing

    Public ReadOnly Property DummyCandle As Payload = Nothing

    Private ReadOnly _tradeStartTime As Date = Date.MinValue
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal controller As Integer,
                   ByVal strikeGap As Decimal,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, controller, strikeGap, canceller)
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        If Me.Controller Then
            CalculateSwingHighLow(_signalPayload, _swingHighPayload, _swingLowPayload)

            _buyLevel = _inputPayload.Max(Function(x)
                                              If x.Key.Date = _tradingDate.Date AndAlso x.Key < _tradeStartTime Then
                                                  Return x.Value.High
                                              Else
                                                  Return Decimal.MinValue
                                              End If
                                          End Function)
            _sellLevel = _inputPayload.Min(Function(x)
                                               If x.Key.Date = _tradingDate.Date AndAlso x.Key < _tradeStartTime Then
                                                   Return x.Value.Low
                                               Else
                                                   Return Decimal.MaxValue
                                               End If
                                           End Function)
        End If
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
        If Me.Controller Then
            If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
                For Each runningInstrument In Me.DependentInstrument
                    runningInstrument.EligibleToTakeTrade = False
                    Dim tradingSymbol As String = runningInstrument.TradingSymbol
                    If tradingSymbol.EndsWith("PE") Then
                        Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "PE", tradingSymbol)
                        Dim strike As String = strikeData.Substring(5)
                        If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike) Then
                            If _peOptionStrikeList Is Nothing Then _peOptionStrikeList = New Dictionary(Of Decimal, AjitJhaORBOptionStrategyRule)
                            _peOptionStrikeList.Add(Val(strike), runningInstrument)
                        End If
                    ElseIf tradingSymbol.EndsWith("CE") Then
                        Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "CE", tradingSymbol)
                        Dim strike As String = strikeData.Substring(5)
                        If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike) Then
                            If _ceOptionStrikeList Is Nothing Then _ceOptionStrikeList = New Dictionary(Of Decimal, AjitJhaORBOptionStrategyRule)
                            _ceOptionStrikeList.Add(Val(strike), runningInstrument)
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        _DummyCandle = currentTick
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
           currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not Controller AndAlso Me.ForceTakeTrade Then
                Dim stoplosss As Decimal = ConvertFloorCeling(currentTick.Open * _userInputs.StoplossPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                stoplosss = Math.Max(_userInputs.MinStoplossPoint, Math.Min(_userInputs.MaxStoplossPoint, stoplosss))

                Dim parameter1 As PlaceOrderParameters = Nothing
                parameter1 = New PlaceOrderParameters With {
                                                .EntryPrice = currentTick.Open,
                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                .Quantity = Me.LotSize,
                                                .Stoploss = .EntryPrice - stoplosss,
                                                .Target = .EntryPrice + stoplosss,
                                                .Buffer = 0,
                                                .SignalCandle = currentMinuteCandle,
                                                .OrderType = Trade.TypeOfOrder.Market,
                                                .Supporting1 = stoplosss,
                                                .Supporting2 = Me.Remarks
                                            }
                Dim parameter2 As PlaceOrderParameters = Nothing
                parameter2 = New PlaceOrderParameters With {
                                                .EntryPrice = currentTick.Open,
                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                .Quantity = Me.LotSize,
                                                .Stoploss = .EntryPrice - stoplosss,
                                                .Target = .EntryPrice + 10000000,
                                                .Buffer = 0,
                                                .SignalCandle = currentMinuteCandle,
                                                .OrderType = Trade.TypeOfOrder.Market,
                                                .Supporting1 = stoplosss,
                                                .Supporting2 = Me.Remarks
                                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter1, parameter2})
                Me.ForceTakeTrade = False
                Me.ForceCancelTrade = False
            ElseIf Controller Then
                If Not IsActiveTrade(Trade.TradeExecutionDirection.Buy) Then
                    Dim takeTrade As Boolean = False
                    Dim condition As String = Nothing
                    Dim lastEntryTrade As Trade = _parentStrategy.GetOverallLastEntryTrade(_tradingDate)
                    If _buyORBTriggered OrElse _sellORBTriggered Then
                        Dim swingHigh As Payload = _swingHighPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                        If swingHigh IsNot Nothing AndAlso currentTick.Open >= swingHigh.High AndAlso swingHigh.PayloadDate >= lastEntryTrade.SignalCandle.PayloadDate AndAlso
                            Not IsSignalTriggered(swingHigh.High, Trade.TradeExecutionDirection.Buy, swingHigh.PayloadDate.AddMinutes(_parentStrategy.SignalTimeFrame), currentMinuteCandle.PayloadDate) Then
                            takeTrade = True
                            condition = String.Format("Swing High({0}) triggered", swingHigh.PayloadDate.ToString("HH:mm"))
                        End If
                    End If
                    If Not _buyORBTriggered AndAlso Not takeTrade AndAlso currentTick.Open >= _buyLevel Then
                        takeTrade = True
                        condition = "Opening Range Buy Breakout"
                    End If
                    If takeTrade Then
                        'If lastEntryTrade Is Nothing OrElse currentMinuteCandle.PayloadDate >= lastEntryTrade.SignalCandle.PayloadDate.AddMinutes(30) Then
                        Dim strikePrice As Decimal = Math.Floor(currentTick.Open / Me.StrikeGap) * Me.StrikeGap
                        If _ceOptionStrikeList.ContainsKey(strikePrice) Then
                            Dim instrumentToTrade As AjitJhaORBOptionStrategyRule = _ceOptionStrikeList(strikePrice)
                            If instrumentToTrade IsNot Nothing Then
                                instrumentToTrade.ForceTakeTrade = True
                                instrumentToTrade.EligibleToTakeTrade = True
                                instrumentToTrade.Remarks = condition
                            End If
                        Else
                            _parentStrategy.SkipCurrentDay = True
                        End If
                        'End If
                    End If
                Else
                    _buyORBTriggered = True
                End If
                If Not IsActiveTrade(Trade.TradeExecutionDirection.Sell) Then
                    Dim takeTrade As Boolean = False
                    Dim condition As String = Nothing
                    Dim lastEntryTrade As Trade = _parentStrategy.GetOverallLastEntryTrade(_tradingDate)
                    If _buyORBTriggered OrElse _sellORBTriggered Then
                        Dim swingLow As Payload = _swingLowPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                        If swingLow IsNot Nothing AndAlso currentTick.Open <= swingLow.Low AndAlso swingLow.PayloadDate >= lastEntryTrade.SignalCandle.PayloadDate AndAlso
                            Not IsSignalTriggered(swingLow.Low, Trade.TradeExecutionDirection.Sell, swingLow.PayloadDate.AddMinutes(_parentStrategy.SignalTimeFrame), currentMinuteCandle.PayloadDate) Then
                            takeTrade = True
                            condition = String.Format("Swing Low({0}) triggered", swingLow.PayloadDate.ToString("HH:mm"))
                        End If
                    End If
                    If Not _sellORBTriggered AndAlso Not takeTrade AndAlso currentTick.Open <= _sellLevel Then
                        takeTrade = True
                        condition = "Opening Range Sell Breakout"
                    End If
                    If takeTrade Then
                        'If lastEntryTrade Is Nothing OrElse currentMinuteCandle.PayloadDate >= lastEntryTrade.SignalCandle.PayloadDate.AddMinutes(30) Then
                        Dim strikePrice As Decimal = Math.Ceiling(currentTick.Open / Me.StrikeGap) * Me.StrikeGap
                        If _peOptionStrikeList.ContainsKey(strikePrice) Then
                            Dim instrumentToTrade As AjitJhaORBOptionStrategyRule = _peOptionStrikeList(strikePrice)
                            If instrumentToTrade IsNot Nothing Then
                                instrumentToTrade.ForceTakeTrade = True
                                instrumentToTrade.EligibleToTakeTrade = True
                                instrumentToTrade.Remarks = condition
                            End If
                        Else
                            _parentStrategy.SkipCurrentDay = True
                        End If
                        'End If
                    End If
                Else
                    _sellORBTriggered = True
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        _DummyCandle = currentTick
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
            Dim slPoint As Decimal = currentTrade.Supporting1
            If plPoint >= slPoint Then
                Dim mul As Integer = Math.Floor(plPoint / slPoint)
                Dim triggerPrice As Decimal = currentTrade.EntryPrice + slPoint * (mul - 1)
                If currentTrade.PotentialStopLoss < triggerPrice Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Stoploss move to {0}", mul - 1))
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsActiveTrade(direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        If direction = Trade.TradeExecutionDirection.Buy Then
            For Each runningInstrument In _ceOptionStrikeList
                If runningInstrument.Value.DummyCandle IsNot Nothing Then
                    If _parentStrategy.IsTradeActive(runningInstrument.Value.DummyCandle, Trade.TypeOfTrade.MIS) Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            For Each runningInstrument In _peOptionStrikeList
                If runningInstrument.Value.DummyCandle IsNot Nothing Then
                    If _parentStrategy.IsTradeActive(runningInstrument.Value.DummyCandle, Trade.TypeOfTrade.MIS) Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

#Region "Indicator"
    Private Sub CalculateSwingHighLow(inputPayload As Dictionary(Of Date, Payload), ByRef outputHighPayload As Dictionary(Of Date, Payload), ByRef outputLowPayload As Dictionary(Of Date, Payload))
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            For Each runningPayload In inputPayload.Keys
                Dim swingHigh As Payload = Nothing
                Dim swingLow As Payload = Nothing

                Dim previousnInputPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, runningPayload, 7, True)
                If previousnInputPayload IsNot Nothing AndAlso previousnInputPayload.Count = 7 Then
                    Dim middleCandle As KeyValuePair(Of Date, Payload) = Common.GetPayloadAt(previousnInputPayload.ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value), runningPayload, -4)
                    Dim previousHigh As Decimal = previousnInputPayload.Max(Function(x)
                                                                                If x.Key < middleCandle.Key Then
                                                                                    Return x.Value.High
                                                                                Else
                                                                                    Return Decimal.MinValue
                                                                                End If
                                                                            End Function)
                    Dim previousLow As Decimal = previousnInputPayload.Min(Function(x)
                                                                               If x.Key < middleCandle.Key Then
                                                                                   Return x.Value.Low
                                                                               Else
                                                                                   Return Decimal.MaxValue
                                                                               End If
                                                                           End Function)
                    Dim nextHigh As Decimal = previousnInputPayload.Max(Function(x)
                                                                            If x.Key > middleCandle.Key Then
                                                                                Return x.Value.High
                                                                            Else
                                                                                Return Decimal.MinValue
                                                                            End If
                                                                        End Function)
                    Dim nextLow As Decimal = previousnInputPayload.Min(Function(x)
                                                                           If x.Key > middleCandle.Key Then
                                                                               Return x.Value.Low
                                                                           Else
                                                                               Return Decimal.MaxValue
                                                                           End If
                                                                       End Function)
                    If middleCandle.Value.High > previousHigh AndAlso middleCandle.Value.High > nextHigh Then
                        swingHigh = middleCandle.Value
                    Else
                        swingHigh = outputHighPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                    End If
                    If middleCandle.Value.Low < previousLow AndAlso middleCandle.Value.Low < nextLow Then
                        swingLow = middleCandle.Value
                    Else
                        swingLow = outputLowPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate)
                    End If
                End If

                If outputHighPayload Is Nothing Then outputHighPayload = New Dictionary(Of Date, Payload)
                outputHighPayload.Add(runningPayload, swingHigh)
                If outputLowPayload Is Nothing Then outputLowPayload = New Dictionary(Of Date, Payload)
                outputLowPayload.Add(runningPayload, swingLow)
            Next
        End If
    End Sub
#End Region
End Class