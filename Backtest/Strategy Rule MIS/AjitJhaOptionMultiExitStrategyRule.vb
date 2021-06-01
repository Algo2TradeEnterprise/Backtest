Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class AjitJhaOptionMultiExitStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public SpotMaxStoplossPercentage As Decimal
        Public OptionMaxStoplossPoint As Decimal
        Public OptionMinStoplossPointOnExpiry As Decimal
        Public TargetToMaxStoplossMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public Remarks As String = Nothing

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _peOptionStrikeList As Dictionary(Of Decimal, AjitJhaOptionMultiExitStrategyRule) = Nothing
    Private _ceOptionStrikeList As Dictionary(Of Decimal, AjitJhaOptionMultiExitStrategyRule) = Nothing

    Private ReadOnly _tradeStartTime As Date = Date.MinValue
    Public ReadOnly Property DummyCandle As Payload = Nothing

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

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
        If Me.Controller Then
            If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
                Indicator.EMA.CalculateEMA(20, Payload.PayloadFields.Close, _hkPayload, _emaPayload)

                For Each runningInstrument In Me.DependentInstrument
                    runningInstrument.EligibleToTakeTrade = False
                    Dim tradingSymbol As String = runningInstrument.TradingSymbol
                    If tradingSymbol.EndsWith("PE") Then
                        Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "PE", tradingSymbol)
                        Dim strike As String = strikeData.Substring(5)
                        If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike) Then
                            If _peOptionStrikeList Is Nothing Then _peOptionStrikeList = New Dictionary(Of Decimal, AjitJhaOptionMultiExitStrategyRule)
                            _peOptionStrikeList.Add(Val(strike), runningInstrument)
                        End If
                    ElseIf tradingSymbol.EndsWith("CE") Then
                        Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "CE", tradingSymbol)
                        Dim strike As String = strikeData.Substring(5)
                        If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike) Then
                            If _ceOptionStrikeList Is Nothing Then _ceOptionStrikeList = New Dictionary(Of Decimal, AjitJhaOptionMultiExitStrategyRule)
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
            If Me.ForceTakeTrade AndAlso Not Controller Then
                Dim signalCandle As Payload = _hkPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                If signalCandle IsNot Nothing AndAlso signalCandle.PreviousCandlePayload IsNot Nothing AndAlso
                    signalCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                    signalCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim stoploss As Decimal = ConvertFloorCeling(Math.Min(signalCandle.PreviousCandlePayload.Low, signalCandle.PreviousCandlePayload.PreviousCandlePayload.Low), _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim buffer As Decimal = CalculateBuffer(entryPrice)
                    entryPrice = entryPrice + buffer
                    stoploss = stoploss - buffer
                    If (entryPrice - stoploss) <= _userInputs.OptionMaxStoplossPoint Then
                        If (_tradingDate.DayOfWeek = DayOfWeek.Thursday AndAlso (entryPrice - stoploss) >= _userInputs.OptionMinStoplossPointOnExpiry) OrElse
                        (_tradingDate.DayOfWeek <> DayOfWeek.Thursday) Then
                            Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(Me.TradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Futures)
                            quantity = Math.Floor(quantity / Me.LotSize) * Me.LotSize
                            If quantity > 0 Then
                                Dim maxStoplossPL As Decimal = _parentStrategy.CalculatePL(Me.TradingSymbol, entryPrice, entryPrice - _userInputs.OptionMaxStoplossPoint, quantity, Me.LotSize, Trade.TypeOfStock.Futures)
                                Dim targetPL As Decimal = Math.Abs(maxStoplossPL) * _userInputs.TargetToMaxStoplossMultiplier
                                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me.TradingSymbol, entryPrice, quantity, targetPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Futures)
                                If target - entryPrice < entryPrice - stoploss Then target = entryPrice + (entryPrice - stoploss)

                                Dim parameter1 As PlaceOrderParameters = Nothing
                                parameter1 = New PlaceOrderParameters With {
                                                .EntryPrice = entryPrice,
                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                .Quantity = quantity,
                                                .Stoploss = stoploss,
                                                .Target = target,
                                                .Buffer = buffer,
                                                .SignalCandle = signalCandle,
                                                .OrderType = Trade.TypeOfOrder.SL,
                                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                .Supporting2 = Me.Remarks,
                                                .Supporting3 = (target - entryPrice),
                                                .Supporting4 = 1
                                            }

                                Dim parameter2 As PlaceOrderParameters = Nothing
                                parameter2 = New PlaceOrderParameters With {
                                                .EntryPrice = entryPrice,
                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                .Quantity = quantity,
                                                .Stoploss = stoploss,
                                                .Target = .EntryPrice + 100000000,
                                                .Buffer = buffer,
                                                .SignalCandle = signalCandle,
                                                .OrderType = Trade.TypeOfOrder.SL,
                                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                .Supporting2 = Me.Remarks,
                                                .Supporting3 = (target - entryPrice),
                                                .Supporting4 = 2
                                            }

                                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter1, parameter2})
                            Else
                                Console.WriteLine(String.Format("{0}: Unable to take trade for quntity. Entry={1}, Stoploss={2}, Quantity={3}, Signal Candle={4}, Direction={5}, Condition={6}",
                                                            Me.TradingSymbol, entryPrice, stoploss, quantity, signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), Me.Direction.ToString, Me.Remarks))
                            End If
                        Else
                            Console.WriteLine(String.Format("{0}: Unable to take trade for lower stoploss on expiry. Entry={1}, Stoploss={2}, Signal Candle={3}, Direction={4}, Condition={5}",
                                                            Me.TradingSymbol, entryPrice, stoploss, signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), Me.Direction.ToString, Me.Remarks))
                        End If
                    Else
                        Console.WriteLine(String.Format("{0}: Unable to take trade for higher stoploss. Entry={1}, Stoploss={2}, Signal Candle={3}, Direction={4}, Condition={5}",
                                                            Me.TradingSymbol, entryPrice, stoploss, signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), Me.Direction.ToString, Me.Remarks))
                    End If
                End If
                Me.ForceTakeTrade = False
            ElseIf Controller Then
                If Not IsActiveTrade() Then
                    Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                            Dim instrumentToTrade As AjitJhaOptionMultiExitStrategyRule = _ceOptionStrikeList.Where(Function(x)
                                                                                                                        Return x.Key <= signal.Item2
                                                                                                                    End Function).OrderByDescending(Function(y)
                                                                                                                                                        Return y.Key
                                                                                                                                                    End Function).FirstOrDefault.Value
                            If instrumentToTrade IsNot Nothing Then
                                instrumentToTrade.ForceTakeTrade = True
                                instrumentToTrade.EligibleToTakeTrade = True
                                instrumentToTrade.Direction = signal.Item3
                                instrumentToTrade.Remarks = signal.Item4
                            End If
                        ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                            Dim instrumentToTrade As AjitJhaOptionMultiExitStrategyRule = _peOptionStrikeList.Where(Function(x)
                                                                                                                        Return x.Key >= signal.Item2
                                                                                                                    End Function).OrderBy(Function(y)
                                                                                                                                              Return y.Key
                                                                                                                                          End Function).FirstOrDefault.Value
                            If instrumentToTrade IsNot Nothing Then
                                instrumentToTrade.ForceTakeTrade = True
                                instrumentToTrade.EligibleToTakeTrade = True
                                instrumentToTrade.Direction = signal.Item3
                                instrumentToTrade.Remarks = signal.Item4
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        _DummyCandle = currentTick
        If Me.ForceCancelTrade AndAlso Not Controller Then
            ret = New Tuple(Of Boolean, String)(True, "Reverse Signal")

            If currentTrade.Supporting4 = 2 Then Me.ForceCancelTrade = False
        ElseIf Controller Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            For Each runningInstrument As AjitJhaOptionMultiExitStrategyRule In Me.DependentInstrument
                If runningInstrument.DummyCandle IsNot Nothing Then
                    Dim lastTrade As Trade = _parentStrategy.GetLastTradeOfTheStock(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS)
                    If lastTrade IsNot Nothing Then
                        If lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = GetEntrySignal(currentMinuteCandle, currentTick)
                            If signal IsNot Nothing AndAlso signal.Item1 Then
                                If signal.Item3 = Trade.TradeExecutionDirection.Buy AndAlso lastTrade.TradingSymbol.EndsWith("PE") Then
                                    runningInstrument.ForceCancelTrade = True
                                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell AndAlso lastTrade.TradingSymbol.EndsWith("CE") Then
                                    runningInstrument.ForceCancelTrade = True
                                End If
                            End If
                        ElseIf lastTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                            Dim hkCandle As Payload = _hkPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                            If hkCandle.CandleColor = Color.Green AndAlso lastTrade.TradingSymbol.EndsWith("PE") Then
                                runningInstrument.ForceCancelTrade = True
                            ElseIf hkCandle.CandleColor = Color.Red AndAlso lastTrade.TradingSymbol.EndsWith("CE") Then
                                runningInstrument.ForceCancelTrade = True
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting3
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(Me.TradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, Me.LotSize, Trade.TypeOfStock.Futures)
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Breakeven at {1}", slPoint, currentTick.PayloadDate.ToString("HH:mm:ss")))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            If signalCandle IsNot Nothing AndAlso signalCandle.PreviousCandlePayload IsNot Nothing AndAlso
                signalCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                signalCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                If CInt(signalCandle.Open) = CInt(signalCandle.Low) Then
                    Dim condition As String = Nothing
                    If signalCandle.Close > _emaPayload(signalCandle.PayloadDate) AndAlso
                        signalCandle.PreviousCandlePayload.Close < _emaPayload(signalCandle.PreviousCandlePayload.PayloadDate) Then
                        If signalCandle.PreviousCandlePayload.CandleColor = Color.Red OrElse
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red Then
                            condition = "Condition 1"
                        ElseIf signalCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish AndAlso
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                            condition = "Condition 2"
                        End If
                    ElseIf signalCandle.Close > _emaPayload(signalCandle.PayloadDate) Then
                        If (signalCandle.PreviousCandlePayload.CandleColor = Color.Red AndAlso
                            signalCandle.PreviousCandlePayload.Low < _emaPayload(signalCandle.PreviousCandlePayload.PayloadDate)) OrElse
                            (signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red AndAlso
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.Low < _emaPayload(signalCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate)) Then
                            condition = "Condition 3"
                        End If
                    End If
                    If condition IsNot Nothing Then
                        'Dim buffer As Decimal = CalculateBuffer(signalCandle.High)
                        Dim buffer As Decimal = 0
                        Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                        Dim stoploss As Decimal = ConvertFloorCeling(Math.Min(signalCandle.PreviousCandlePayload.Low, signalCandle.PreviousCandlePayload.PreviousCandlePayload.Low), _parentStrategy.TickSize, RoundOfType.Floor) - buffer
                        If entryPrice - stoploss < entryPrice * _userInputs.SpotMaxStoplossPercentage / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)(True, entryPrice, Trade.TradeExecutionDirection.Buy, condition)
                        Else
                            Console.WriteLine(String.Format("Neglected because of bigger stoploss on spot. Signal Candle:{0}, Direction:BUY, Condition:{1}",
                                                            signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), condition))
                        End If
                    End If
                ElseIf CInt(signalCandle.Open) = CInt(signalCandle.High) Then
                    Dim condition As String = Nothing
                    If signalCandle.Close < _emaPayload(signalCandle.PayloadDate) AndAlso
                        signalCandle.PreviousCandlePayload.Close > _emaPayload(signalCandle.PreviousCandlePayload.PayloadDate) Then
                        If signalCandle.PreviousCandlePayload.CandleColor = Color.Green OrElse
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green Then
                            condition = "Condition 1"
                        ElseIf signalCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish AndAlso
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                            condition = "Condition 2"
                        End If
                    ElseIf signalCandle.Close < _emaPayload(signalCandle.PayloadDate) Then
                        If (signalCandle.PreviousCandlePayload.CandleColor = Color.Green AndAlso
                            signalCandle.PreviousCandlePayload.High > _emaPayload(signalCandle.PreviousCandlePayload.PayloadDate)) OrElse
                            (signalCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green AndAlso
                            signalCandle.PreviousCandlePayload.PreviousCandlePayload.High > _emaPayload(signalCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate)) Then
                            condition = "Condition 3"
                        End If
                    End If
                    If condition IsNot Nothing Then
                        'Dim buffer As Decimal = CalculateBuffer(signalCandle.Low)
                        Dim buffer As Decimal = 0
                        Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor) - buffer
                        Dim stoploss As Decimal = ConvertFloorCeling(Math.Max(signalCandle.PreviousCandlePayload.High, signalCandle.PreviousCandlePayload.PreviousCandlePayload.High), _parentStrategy.TickSize, RoundOfType.Celing) + buffer
                        If stoploss - entryPrice < entryPrice * _userInputs.SpotMaxStoplossPercentage / 100 Then
                            ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, String)(True, entryPrice, Trade.TradeExecutionDirection.Sell, condition)
                        Else
                            Console.WriteLine(String.Format("Neglected because of bigger stoploss on spot. Signal Candle:{0}, Direction:SELL, Condition:{1}",
                                                            signalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"), condition))
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsActiveTrade() As Boolean
        Dim ret As Boolean = False
        If Me.Controller AndAlso Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
            For Each runningInstrument As AjitJhaOptionMultiExitStrategyRule In Me.DependentInstrument
                If runningInstrument.DummyCandle IsNot Nothing Then
                    If _parentStrategy.IsTradeActive(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS) OrElse
                        _parentStrategy.IsTradeOpen(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS) Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function CalculateBuffer(ByVal price As Decimal) As Decimal
        Return ConvertFloorCeling(price * 0.5 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
    End Function
End Class