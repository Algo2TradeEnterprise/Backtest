Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class OptionLadderStrategyRule
    Inherits StrategyRule

    Private _startingPrice As Decimal = Decimal.MinValue
    Private _nextCEStrike As Decimal = Decimal.MinValue
    Private _nextPEStrike As Decimal = Decimal.MinValue
    Private _takeTrade As Boolean = False
    Private _firstTimeEntry As Boolean = False
    Private _stockCount As Integer = Integer.MinValue

    Private _peOptionStrikeList As Dictionary(Of Decimal, StrategyRule) = Nothing
    Private _ceOptionStrikeList As Dictionary(Of Decimal, StrategyRule) = Nothing

    Private ReadOnly _strikeDifference As Decimal = 200
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not Controller Then
            Dim quantity As Integer = Me.LotSize

            Dim parameter As PlaceOrderParameters = Nothing
            parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + 10000000,
                            .Target = .EntryPrice - 10000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = "Force Entry"
                        }

            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            Me.ForceTakeTrade = False
        ElseIf Controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            Dim parameter As PlaceOrderParameters = Nothing
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                If _startingPrice = Decimal.MinValue Then
                    Dim nearestUpperStrike As Decimal = _ceOptionStrikeList.Keys.Min(Function(x)
                                                                                         If x > currentTick.Open Then
                                                                                             Return x
                                                                                         Else
                                                                                             Return Decimal.MaxValue
                                                                                         End If
                                                                                     End Function)

                    Dim nearestLoweStrike As Decimal = _ceOptionStrikeList.Keys.Max(Function(x)
                                                                                        If x <= currentTick.Open Then
                                                                                            Return x
                                                                                        Else
                                                                                            Return Decimal.MinValue
                                                                                        End If
                                                                                    End Function)

                    If nearestUpperStrike - currentTick.Open > currentTick.Open - nearestLoweStrike Then
                        _startingPrice = nearestLoweStrike
                    Else
                        _startingPrice = nearestUpperStrike
                    End If
                    _nextPEStrike = _startingPrice - _strikeDifference
                    _nextCEStrike = _startingPrice + _strikeDifference
                    _takeTrade = True
                    _firstTimeEntry = True
                End If
                If _takeTrade Then
                    If _firstTimeEntry Then
                        _firstTimeEntry = False
                        '_stockCount = Math.Floor((currentTick.Open * 10 / 100) / 100)
                        _stockCount = 5
                        For counter As Integer = 0 To _stockCount - 1
                            _ceOptionStrikeList(_nextCEStrike + counter * _strikeDifference).ForceTakeTrade = True
                            _ceOptionStrikeList(_nextCEStrike + counter * _strikeDifference).EligibleToTakeTrade = True

                            _peOptionStrikeList(_nextPEStrike - counter * _strikeDifference).ForceTakeTrade = True
                            _peOptionStrikeList(_nextPEStrike - counter * _strikeDifference).EligibleToTakeTrade = True
                        Next

                        'Dim counter As Integer = 0
                        'For Each runningInstrument In _ceOptionStrikeList.OrderBy(Function(x)
                        '                                                              Return x.Key
                        '                                                          End Function)
                        '    If runningInstrument.Key >= _nextCEStrike Then
                        '        runningInstrument.Value.ForceTakeTrade = True
                        '        runningInstrument.Value.EligibleToTakeTrade = True
                        '        counter += 1
                        '        If counter >= _stockCount Then Exit For
                        '    End If
                        'Next

                        'counter = 0
                        'For Each runningInstrument In _peOptionStrikeList.OrderByDescending(Function(x)
                        '                                                                        Return x.Key
                        '                                                                    End Function)
                        '    If runningInstrument.Key <= _nextPEStrike Then
                        '        runningInstrument.Value.ForceTakeTrade = True
                        '        runningInstrument.Value.EligibleToTakeTrade = True
                        '        counter += 1
                        '        If counter >= _stockCount Then Exit For
                        '    End If
                        'Next
                        _takeTrade = False
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If Me.ForceCancelTrade AndAlso Not Controller Then
            If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                ret = New Tuple(Of Boolean, String)(True, "Force Exit")
                Me.ForceCancelTrade = False
            End If
        ElseIf Controller Then
            If _nextCEStrike <> Decimal.MinValue AndAlso currentTick.Open >= _nextCEStrike Then
                _ceOptionStrikeList(_nextCEStrike).ForceCancelTrade = True
                _peOptionStrikeList(_nextPEStrike - ((_stockCount - 1) * _strikeDifference)).ForceCancelTrade = True
                _nextPEStrike = _nextPEStrike + _strikeDifference
                _peOptionStrikeList(_nextPEStrike).ForceTakeTrade = True
                _peOptionStrikeList(_nextPEStrike).EligibleToTakeTrade = True
                _nextCEStrike = _nextCEStrike + _strikeDifference
                _ceOptionStrikeList(_nextCEStrike + ((_stockCount - 1) * _strikeDifference)).ForceTakeTrade = True
                _ceOptionStrikeList(_nextCEStrike + ((_stockCount - 1) * _strikeDifference)).EligibleToTakeTrade = True
            ElseIf _nextPEStrike <> Decimal.MinValue AndAlso currentTick.Open <= _nextPEStrike Then
                _peOptionStrikeList(_nextPEStrike).ForceCancelTrade = True
                _ceOptionStrikeList(_nextCEStrike + ((_stockCount - 1) * _strikeDifference)).ForceCancelTrade = True
                _nextCEStrike = _nextCEStrike - _strikeDifference
                _ceOptionStrikeList(_nextCEStrike).ForceTakeTrade = True
                _ceOptionStrikeList(_nextCEStrike).EligibleToTakeTrade = True
                _nextPEStrike = _nextPEStrike - _strikeDifference
                _peOptionStrikeList(_nextPEStrike - ((_stockCount - 1) * _strikeDifference)).ForceTakeTrade = True
                _peOptionStrikeList(_nextPEStrike - ((_stockCount - 1) * _strikeDifference)).EligibleToTakeTrade = True
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
End Class