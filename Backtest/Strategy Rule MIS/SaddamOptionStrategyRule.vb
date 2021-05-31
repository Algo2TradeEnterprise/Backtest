Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class SaddamOptionStrategyRule
    Inherits StrategyRule

    Private _firstTimeEntryDone As Boolean = False
    Private _straddleCE As SaddamOptionStrategyRule
    Private _straddlePE As SaddamOptionStrategyRule
    Private _nakedCE As SaddamOptionStrategyRule
    Private _nakedPE As SaddamOptionStrategyRule

    Private _nakedSTCE As SaddamOptionStrategyRule
    Private _nakedSTPE As SaddamOptionStrategyRule

    Private _peOptionStrikeList As Dictionary(Of Decimal, StrategyRule) = Nothing
    Private _ceOptionStrikeList As Dictionary(Of Decimal, StrategyRule) = Nothing

    Private _signalCandle As Payload = Nothing
    Private _signalDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Private _entryDone As Boolean = False

    Private _bullLevel As Decimal = Decimal.MaxValue
    Private _bullLevelTime As Date = Date.MinValue
    Private _bearLevel As Decimal = Decimal.MinValue
    Private _bearLevelTime As Date = Date.MinValue

    Private _straddleStartTime As Date = Date.MinValue

    Private _rsiPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _supertrendPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _supertrendColorPayload As Dictionary(Of Date, Color) = Nothing

    Protected _EMAPayload As Dictionary(Of Date, Decimal) = Nothing

    Public AdjustmentDone As Boolean = False

    Private ReadOnly _buffer As Decimal
    Private ReadOnly _tradeStartTime As Date
    Private ReadOnly _stoplossPoint As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal controller As Integer,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, controller, canceller)

        _buffer = 5
        _stoplossPoint = 50
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If Me.Controller Then
            Indicator.RSI.CalculateRSI(14, _signalPayload, _rsiPayload)
            Indicator.Supertrend.CalculateSupertrend(7, 2, _signalPayload, _supertrendPayload, _supertrendColorPayload)
        Else
            Indicator.EMA.CalculateEMA(9, Payload.PayloadFields.Close, _signalPayload, _EMAPayload)
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
                            If _peOptionStrikeList Is Nothing Then _peOptionStrikeList = New Dictionary(Of Decimal, StrategyRule)
                            _peOptionStrikeList.Add(Val(strike), runningInstrument)
                        End If
                    ElseIf tradingSymbol.EndsWith("CE") Then
                        Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "CE", tradingSymbol)
                        Dim strike As String = strikeData.Substring(5)
                        If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike) Then
                            If _ceOptionStrikeList Is Nothing Then _ceOptionStrikeList = New Dictionary(Of Decimal, StrategyRule)
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
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not Me.Controller Then
            Dim quantity As Integer = Me.LotSize
            Dim slPoint As Decimal = _stoplossPoint
            Dim supporting As String = Me.Comment
            If supporting IsNot Nothing AndAlso supporting.StartsWith("+++") Then
                slPoint = 10000000
                supporting = supporting.Replace("+++", "")
            End If
            Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + slPoint,
                            .Target = .EntryPrice - 10000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = supporting
                        }

            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            Me.ForceTakeTrade = False
            Me.ForceCancelTrade = False
        ElseIf Me.Controller Then
            If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
                currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                If Not _firstTimeEntryDone Then
                    If _straddleCE Is Nothing AndAlso _straddlePE Is Nothing Then
                        Dim currentATM As Decimal = GetATMStrike(currentTick.Open, _ceOptionStrikeList.Keys.ToList)
                        If currentATM <> Decimal.MinValue AndAlso
                            _ceOptionStrikeList.ContainsKey(currentATM) AndAlso
                            _peOptionStrikeList.ContainsKey(currentATM) Then

                            _straddleCE = _ceOptionStrikeList(currentATM)
                            _straddlePE = _peOptionStrikeList(currentATM)
                        End If
                    End If

                    _straddleCE.ForceTakeTrade = True
                    _straddleCE.Comment = "+++Straddle CE Entry"
                    _straddleCE.EligibleToTakeTrade = True

                    _straddlePE.ForceTakeTrade = True
                    _straddlePE.Comment = "+++Straddle PE Entry"
                    _straddlePE.EligibleToTakeTrade = True

                    _firstTimeEntryDone = True
                    _straddleStartTime = currentTick.PayloadDate
                Else
                    If _signalCandle IsNot Nothing Then
                        If Not _entryDone Then
                            If _signalDirection = Trade.TradeExecutionDirection.Buy AndAlso
                                currentMinuteCandle.PreviousCandlePayload.Close < _bullLevel Then
                                _signalCandle = Nothing
                                _signalDirection = Trade.TradeExecutionDirection.None
                            ElseIf _signalDirection = Trade.TradeExecutionDirection.Sell AndAlso
                                currentMinuteCandle.PreviousCandlePayload.Close > _bearLevel Then
                                _signalCandle = Nothing
                                _signalDirection = Trade.TradeExecutionDirection.None
                            End If
                        Else
                            If _signalDirection = Trade.TradeExecutionDirection.Buy AndAlso
                                currentMinuteCandle.PreviousCandlePayload.Close < _signalCandle.Low AndAlso
                                currentTick.Open < currentMinuteCandle.PreviousCandlePayload.Low Then
                                If currentTick.Open > _bearLevel AndAlso currentTick.Open < _bullLevel Then
                                    _firstTimeEntryDone = False
                                End If
                                _nakedPE.ForceCancelTrade = True
                                _nakedPE.Comment = "Below breakout candle Naked PE Exit"
                                _signalCandle = Nothing
                                _signalDirection = Trade.TradeExecutionDirection.None
                            ElseIf _signalDirection = Trade.TradeExecutionDirection.Sell AndAlso
                                currentMinuteCandle.PreviousCandlePayload.Close > _signalCandle.High AndAlso
                                currentTick.Open > currentMinuteCandle.PreviousCandlePayload.High Then
                                If currentTick.Open > _bearLevel AndAlso currentTick.Open < _bullLevel Then
                                    _firstTimeEntryDone = False
                                End If
                                _nakedCE.ForceCancelTrade = True
                                _nakedCE.Comment = "Above breakout candle Naked PE Exit"
                                _signalCandle = Nothing
                                _signalDirection = Trade.TradeExecutionDirection.None
                            End If
                        End If
                        If _signalDirection = Trade.TradeExecutionDirection.Buy Then
                            If currentTick.Open >= _signalCandle.High + _buffer Then
                                If Not _entryDone Then
                                    _entryDone = True
                                    Me.AdjustmentDone = True

                                    _straddleCE.ForceCancelTrade = True
                                    _straddleCE.Comment = String.Format("Straddle CE Exit on Bull level({0} at {1})", _bullLevel, _bullLevelTime.ToString("HH:mm"))

                                    _straddlePE.ForceModifyTrade = True

                                    Dim nearestOTMPEStrike As Decimal = _peOptionStrikeList.Keys.Max(Function(x)
                                                                                                         If x <= _bullLevel - 100 Then
                                                                                                             Return x
                                                                                                         Else
                                                                                                             Return Decimal.MinValue
                                                                                                         End If
                                                                                                     End Function)
                                    _nakedPE = _peOptionStrikeList(nearestOTMPEStrike)
                                    _nakedPE.ForceTakeTrade = True
                                    _nakedPE.Comment = String.Format("Naked PE Entry on Bull level({0} at {1})", _bullLevel, _bullLevelTime.ToString("HH:mm"))
                                    _nakedPE.EligibleToTakeTrade = True
                                End If
                            End If
                            If currentTick.Open > _bullLevel AndAlso _nakedSTCE Is Nothing AndAlso
                                _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Red Then
                                Dim startingCandle As Payload = GetCurrentSupertrendStartingCandle(currentMinuteCandle)
                                If currentTick.Open <= startingCandle.Low Then
                                    Dim nakedPECandle As Payload = _nakedPE.GetCurrentCandle(currentTick.PayloadDate)
                                    Dim lowestDistance As Decimal = Decimal.MaxValue
                                    Dim lowestDistanceStrike As Decimal = Decimal.MinValue
                                    For Each runningStrike In _ceOptionStrikeList
                                        Dim currentCandle As Payload = CType(runningStrike.Value, SaddamOptionStrategyRule).GetCurrentCandle(currentTick.PayloadDate)
                                        If currentCandle IsNot Nothing Then
                                            Dim distance As Decimal = Math.Abs(currentCandle.Open - nakedPECandle.Open)
                                            If distance < lowestDistance Then
                                                lowestDistance = distance
                                                lowestDistanceStrike = runningStrike.Key
                                            End If
                                        End If
                                    Next

                                    If lowestDistanceStrike <> Decimal.MinValue Then
                                        _nakedSTCE = _ceOptionStrikeList(lowestDistanceStrike)
                                        _nakedSTCE.ForceTakeTrade = True
                                        _nakedSTCE.Comment = "+++Low Break Supertrend Red above bull level"
                                        _nakedSTCE.EligibleToTakeTrade = True
                                    End If
                                End If
                            End If
                            If currentTick.Open > _bullLevel Then
                                Dim dayHigh As Decimal = _inputPayload.Max(Function(x)
                                                                               If x.Key.Date = _tradingDate.Date AndAlso x.Key <= currentTick.PayloadDate Then
                                                                                   Return x.Value.High
                                                                               Else
                                                                                   Return Decimal.MinValue
                                                                               End If
                                                                           End Function)

                                If currentTick.Open >= dayHigh AndAlso Not _parentStrategy.IsAnyTradeActive() Then
                                    Dim otmPEStrike As Decimal = _peOptionStrikeList.Keys.Max(Function(x)
                                                                                                  If x <= currentTick.Open - 200 Then
                                                                                                      Return x
                                                                                                  Else
                                                                                                      Return Decimal.MinValue
                                                                                                  End If
                                                                                              End Function)

                                    Me.AdjustmentDone = True
                                    _peOptionStrikeList(otmPEStrike).ForceTakeTrade = True
                                    _peOptionStrikeList(otmPEStrike).Comment = "Day High Entry"
                                    _peOptionStrikeList(otmPEStrike).EligibleToTakeTrade = True
                                End If
                            End If
                        ElseIf _signalDirection = Trade.TradeExecutionDirection.Sell Then
                            If currentTick.Open <= _signalCandle.Low - _buffer Then
                                If Not _entryDone Then
                                    _entryDone = True
                                    Me.AdjustmentDone = True

                                    _straddlePE.ForceCancelTrade = True
                                    _straddlePE.Comment = String.Format("Straddle PE Exit on Bear level({0} at {1})", _bearLevel, _bearLevelTime.ToString("HH:mm"))

                                    _straddleCE.ForceModifyTrade = True

                                    Dim nearestOTMCEStrike As Decimal = _ceOptionStrikeList.Keys.Min(Function(x)
                                                                                                         If x >= _bearLevel + 100 Then
                                                                                                             Return x
                                                                                                         Else
                                                                                                             Return Decimal.MaxValue
                                                                                                         End If
                                                                                                     End Function)
                                    _nakedCE = _ceOptionStrikeList(nearestOTMCEStrike)
                                    _nakedCE.ForceTakeTrade = True
                                    _nakedCE.Comment = String.Format("Naked CE Entry on Bear level({0} at {1})", _bearLevel, _bearLevelTime.ToString("HH:mm"))
                                    _nakedCE.EligibleToTakeTrade = True
                                End If
                            End If
                            If currentTick.Open < _bearLevel AndAlso _nakedSTPE Is Nothing AndAlso
                                _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Green Then
                                Dim startingCandle As Payload = GetCurrentSupertrendStartingCandle(currentMinuteCandle)
                                If currentTick.Open >= startingCandle.High Then
                                    Dim nakedCECandle As Payload = _nakedCE.GetCurrentCandle(currentTick.PayloadDate)
                                    Dim lowestDistance As Decimal = Decimal.MaxValue
                                    Dim lowestDistanceStrike As Decimal = Decimal.MinValue
                                    For Each runningStrike In _peOptionStrikeList
                                        Dim currentCandle As Payload = CType(runningStrike.Value, SaddamOptionStrategyRule).GetCurrentCandle(currentTick.PayloadDate)
                                        If currentCandle IsNot Nothing Then
                                            Dim distance As Decimal = Math.Abs(currentCandle.Open - nakedCECandle.Open)
                                            If distance < lowestDistance Then
                                                lowestDistance = distance
                                                lowestDistanceStrike = runningStrike.Key
                                            End If
                                        End If
                                    Next

                                    If lowestDistanceStrike <> Decimal.MinValue Then
                                        _nakedSTPE = _peOptionStrikeList(lowestDistanceStrike)
                                        _nakedSTPE.ForceTakeTrade = True
                                        _nakedSTPE.Comment = "+++High Break on Supertrend Green below bear level"
                                        _nakedSTPE.EligibleToTakeTrade = True
                                    End If
                                End If
                            End If
                            If currentTick.Open < _bearLevel Then
                                Dim dayLow As Decimal = _inputPayload.Min(Function(x)
                                                                              If x.Key.Date = _tradingDate.Date AndAlso x.Key <= currentTick.PayloadDate Then
                                                                                  Return x.Value.Low
                                                                              Else
                                                                                  Return Decimal.MaxValue
                                                                              End If
                                                                          End Function)

                                If currentTick.Open <= dayLow AndAlso Not _parentStrategy.IsAnyTradeActive() Then
                                    Dim otmCEStrike As Decimal = _ceOptionStrikeList.Keys.Min(Function(x)
                                                                                                  If x >= currentTick.Open + 200 Then
                                                                                                      Return x
                                                                                                  Else
                                                                                                      Return Decimal.MaxValue
                                                                                                  End If
                                                                                              End Function)

                                    Me.AdjustmentDone = True
                                    _ceOptionStrikeList(otmCEStrike).ForceTakeTrade = True
                                    _ceOptionStrikeList(otmCEStrike).Comment = "Day Low Entry"
                                    _ceOptionStrikeList(otmCEStrike).EligibleToTakeTrade = True
                                End If
                            End If
                        End If
                    End If

                    If _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Green AndAlso _nakedSTCE IsNot Nothing Then
                        Dim startingCandle As Payload = GetCurrentSupertrendStartingCandle(currentMinuteCandle)
                        If currentTick.Open > startingCandle.High Then
                            _nakedSTCE.ForceCancelTrade = True
                            _nakedSTCE.Comment = "High Break on Supertrend Green"
                        End If
                    End If
                    If _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Red AndAlso _nakedSTPE IsNot Nothing Then
                        Dim startingCandle As Payload = GetCurrentSupertrendStartingCandle(currentMinuteCandle)
                        If currentTick.Open < startingCandle.Low Then
                            _nakedSTPE.ForceCancelTrade = True
                            _nakedSTPE.Comment = "Low Break on Supertrend Red"
                        End If
                    End If

                    If currentMinuteCandle.PreviousCandlePayload.Close > _bullLevel AndAlso
                        _signalDirection <> Trade.TradeExecutionDirection.Buy Then
                        If _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Green Then
                            _signalCandle = currentMinuteCandle.PreviousCandlePayload
                            _signalDirection = Trade.TradeExecutionDirection.Buy

                            _entryDone = False
                        End If
                    ElseIf currentMinuteCandle.PreviousCandlePayload.Close < _bearLevel AndAlso
                        _signalDirection <> Trade.TradeExecutionDirection.Sell Then
                        If _supertrendColorPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) = Color.Red Then
                            _signalCandle = currentMinuteCandle.PreviousCandlePayload
                            _signalDirection = Trade.TradeExecutionDirection.Sell

                            _entryDone = False
                        End If
                    End If

                    If _bullLevel = Decimal.MaxValue Then
                        Dim bullCandle As Payload = GetBullCandle(currentMinuteCandle)
                        If bullCandle IsNot Nothing Then
                            _bullLevel = bullCandle.High
                            _bullLevelTime = bullCandle.PayloadDate
                        End If
                    End If
                    If _bearLevel = Decimal.MinValue Then
                        Dim bearCandle As Payload = GetBearCandle(currentMinuteCandle)
                        If bearCandle IsNot Nothing Then
                            _bearLevel = bearCandle.Low
                            _bearLevelTime = bearCandle.PayloadDate
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
        If Not Me.Controller Then
            If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                If Me.ForceCancelTrade Then
                    If Me.Comment.ToUpper.Contains("NAKED") AndAlso currentTrade.Supporting1.ToUpper.Contains("NAKED") Then
                        ret = New Tuple(Of Boolean, String)(True, Me.Comment)
                        Me.ForceCancelTrade = False
                    ElseIf Me.Comment.ToUpper.Contains("SUPERTREND") AndAlso currentTrade.Supporting1.ToUpper.Contains("SUPERTREND") Then
                        ret = New Tuple(Of Boolean, String)(True, Me.Comment)
                        Me.ForceCancelTrade = False
                    ElseIf Me.Comment.ToUpper.Contains("STRADDLE") AndAlso currentTrade.Supporting1.ToUpper.Contains("STRADDLE") Then
                        ret = New Tuple(Of Boolean, String)(True, Me.Comment)
                        Me.ForceCancelTrade = False
                    End If
                Else
                    If currentTrade.PotentialStopLoss <= currentTrade.EntryPrice + _stoplossPoint AndAlso currentTrade.Supporting1.ToUpper.Contains("STRADDLE") Then
                        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                        Dim ema As Decimal = _EMAPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                        If currentMinuteCandle.PreviousCandlePayload.Close > ema AndAlso
                        currentTick.Open > currentMinuteCandle.PreviousCandlePayload.High Then
                            ret = New Tuple(Of Boolean, String)(True, "EMA Stoploss")
                        End If
                    End If
                End If
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Me.ForceModifyTrade Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, currentTrade.EntryPrice + _stoplossPoint, _stoplossPoint)
                Me.ForceModifyTrade = False
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function GetCurrentSupertrendStartingCandle(currentCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        If _supertrendColorPayload IsNot Nothing AndAlso _supertrendColorPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) Then
            Dim currentColor As Color = _supertrendColorPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            For Each runningPayload In _currentDayPayload.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                If runningPayload.Value.PreviousCandlePayload.PayloadDate <= currentCandle.PreviousCandlePayload.PayloadDate Then
                    If _supertrendColorPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate) <> currentColor Then
                        ret = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetBullCandle(currentCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        If currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
            If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High > currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High Then
                If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High > currentCandle.PreviousCandlePayload.PreviousCandlePayload.High Then
                    Dim potentialSignalCandle As Payload = currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload
                    If potentialSignalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                        Dim rsi As Decimal = _rsiPayload(potentialSignalCandle.PayloadDate)
                        If rsi > 60 Then
                            ret = potentialSignalCandle
                        Else
                            Dim dayLowRSI As Decimal = _rsiPayload.Min(Function(x)
                                                                           If x.Key.Date = _tradingDate.Date AndAlso
                                                                           x.Key <= potentialSignalCandle.PayloadDate Then
                                                                               Return x.Value
                                                                           Else
                                                                               Return Decimal.MaxValue
                                                                           End If
                                                                       End Function)
                            If rsi > dayLowRSI + 20 Then
                                ret = potentialSignalCandle
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetBearCandle(currentCandle As Payload) As Payload
        Dim ret As Payload = Nothing
        If currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
            If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low < currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
                If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low < currentCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                    Dim potentialSignalCandle As Payload = currentCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload
                    If potentialSignalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                        Dim rsi As Decimal = _rsiPayload(potentialSignalCandle.PayloadDate)
                        If rsi < 40 Then
                            ret = potentialSignalCandle
                        Else
                            Dim dayHighRSI As Decimal = _rsiPayload.Max(Function(x)
                                                                            If x.Key.Date = _tradingDate.Date AndAlso
                                                                                x.Key <= potentialSignalCandle.PayloadDate Then
                                                                                Return x.Value
                                                                            Else
                                                                                Return Decimal.MinValue
                                                                            End If
                                                                        End Function)
                            If rsi < dayHighRSI - 20 Then
                                ret = potentialSignalCandle
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetATMStrike(price As Decimal, allStrikes As List(Of Decimal)) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If allStrikes IsNot Nothing AndAlso allStrikes.Count > 0 Then
            Dim upperStrikes As List(Of Decimal) = allStrikes.FindAll(Function(x)
                                                                          Return x >= price
                                                                      End Function)
            Dim lowerStrikes As List(Of Decimal) = allStrikes.FindAll(Function(x)
                                                                          Return x <= price
                                                                      End Function)
            Dim upperStrikePrice As Decimal = Decimal.MaxValue
            Dim lowerStrikePrice As Decimal = Decimal.MinValue
            If upperStrikes IsNot Nothing AndAlso upperStrikes.Count > 0 Then
                upperStrikePrice = upperStrikes.OrderBy(Function(x)
                                                            Return x
                                                        End Function).FirstOrDefault
            End If
            If lowerStrikes IsNot Nothing AndAlso lowerStrikes.Count > 0 Then
                lowerStrikePrice = lowerStrikes.OrderBy(Function(x)
                                                            Return x
                                                        End Function).LastOrDefault
            End If

            If upperStrikePrice <> Decimal.MaxValue AndAlso lowerStrikePrice <> Decimal.MinValue Then
                If upperStrikePrice - price < price - lowerStrikePrice Then
                    ret = upperStrikePrice
                Else
                    ret = lowerStrikePrice
                End If
            ElseIf upperStrikePrice <> Decimal.MaxValue Then
                ret = upperStrikePrice
            ElseIf lowerStrikePrice <> Decimal.MinValue Then
                ret = lowerStrikePrice
            End If
        End If
        Return ret
    End Function

    Private Function GetCurrentTick(currentCandle As Payload, currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If currentCandle IsNot Nothing Then
            Dim selectedTicks As List(Of Payload) = currentCandle.Ticks.FindAll(Function(x)
                                                                                    Return x.PayloadDate <= currentTime
                                                                                End Function)
            If selectedTicks IsNot Nothing AndAlso selectedTicks.Count > 0 Then
                ret = selectedTicks.LastOrDefault
            Else
                ret = currentCandle.Ticks.FirstOrDefault
            End If
        End If
        Return ret
    End Function

    Protected Function GetCurrentCandle(currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _currentDayPayload IsNot Nothing AndAlso _currentDayPayload.Count > 0 Then
            ret = _currentDayPayload.Where(Function(x)
                                               Return x.Key <= currentTime
                                           End Function).LastOrDefault.Value
        End If
        Return ret
    End Function
End Class