﻿Imports Backtest.StrategyHelper
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
        'Public StoplossPercentage As Decimal
    End Class
#End Region

    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _rsiPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _currentDayPaload As Dictionary(Of Date, Payload) = Nothing
    Private _eodPayload As Dictionary(Of Date, Payload) = Nothing
    Private _weeklyPayload As Dictionary(Of Date, Payload) = Nothing

    Private _weeklyHigh As Decimal = Decimal.MinValue
    Private _weeklyLow As Decimal = Decimal.MinValue

    Private ReadOnly _lastTradeEntryTime As Date
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
        _lastTradeEntryTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.SMA.CalculateSMA(20, Payload.PayloadFields.Volume, _signalPayload, _smaPayload)
        Indicator.EMA.CalculateEMA(20, Payload.PayloadFields.Close, _signalPayload, _emaPayload)
        Indicator.RSI.CalculateRSI(14, _signalPayload, _rsiPayload)
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)

        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If _currentDayPaload Is Nothing Then _currentDayPaload = New Dictionary(Of Date, Payload)
                _currentDayPaload.Add(runningPayload.Key, runningPayload.Value)
            End If
        Next

        _eodPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Cash, _tradingSymbol, _tradingDate.AddDays(-300), _tradingDate.AddDays(-1))
        If _eodPayload IsNot Nothing AndAlso _eodPayload.Count >= 150 Then
            _weeklyPayload = Common.ConvertDayPayloadsToWeek(_eodPayload)

            Dim weeklySubPayload As Dictionary(Of Date, Payload) = Nothing
            Dim counter As Integer = 0
            For Each runningPayload In _weeklyPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                If weeklySubPayload Is Nothing Then weeklySubPayload = New Dictionary(Of Date, Payload)
                weeklySubPayload.Add(runningPayload.Key, runningPayload.Value)

                counter += 1
                If counter >= 10 Then Exit For
            Next
            _weeklyHigh = weeklySubPayload.Min(Function(x)
                                                   Return x.Value.High
                                               End Function)
            _weeklyLow = weeklySubPayload.Max(Function(x)
                                                  Return x.Value.Low
                                              End Function)
        Else
            Me.EligibleToTakeTrade = False
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso currentMinuteCandle.PayloadDate <= _lastTradeEntryTime Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal) = GetSignalCandle(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing Then
                    If currentMinuteCandle.PayloadDate > lastExecutedOrder.ExitTime Then
                        signalCandle = signal.Item3
                    End If
                Else
                    signalCandle = signal.Item3
                End If
            End If
            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = CalculateBuffer(signal.Item2)
                Dim entryPrice As Decimal = signal.Item2
                Dim slPer As Decimal = signal.Item6
                Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * slPer / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, Math.Abs(_userInputs.MaxLossPerTrade / 2) * -1, _parentStrategy.StockType)
                Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)

                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = slPoint,
                                    .Supporting4 = targetPoint,
                                    .Supporting5 = "Normal"
                                }

                    parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = slPoint,
                                    .Supporting4 = targetPoint,
                                    .Supporting5 = "Trailling"
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = slPoint,
                                    .Supporting4 = targetPoint,
                                    .Supporting5 = "Normal"
                                }

                    parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = slPoint,
                                    .Supporting4 = targetPoint,
                                    .Supporting5 = "Trailling"
                                }
                End If
            End If
        End If
        If parameter1 IsNot Nothing AndAlso parameter2 IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter1, parameter2})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            If currentMinuteCandle.PayloadDate <= _lastTradeEntryTime Then
                Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal) = GetSignalCandle(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            Else
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting5.ToUpper = "TRAILLING" Then
            Dim slPoint As Decimal = currentTrade.Supporting3
            Dim targetPoint As Decimal = currentTrade.Supporting4
            Dim movementPoint As Decimal = ConvertFloorCeling(slPoint / 2, _parentStrategy.TickSize, RoundOfType.Celing)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso currentTick.Open >= currentTrade.EntryPrice + targetPoint Then
                Dim gain As Decimal = currentTick.Open - currentTrade.EntryPrice
                Dim extraGain As Decimal = gain - targetPoint
                Dim multiplier As Integer = Math.Floor(extraGain / movementPoint) + 1
                Dim triggerPrice As Decimal = (currentTrade.EntryPrice - slPoint) + (movementPoint * multiplier)
                If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss < triggerPrice Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move {0}", multiplier))
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso currentTick.Open <= currentTrade.EntryPrice - targetPoint Then
                Dim gain As Decimal = currentTrade.EntryPrice - currentTick.Open
                Dim extraGain As Decimal = gain - targetPoint
                Dim multiplier As Integer = Math.Floor(extraGain / movementPoint) + 1
                Dim triggerPrice As Decimal = (currentTrade.EntryPrice + slPoint) - (movementPoint * multiplier)
                If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss > triggerPrice Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move {0}", multiplier))
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function CalculateBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = Nothing
        If price <= 300 Then
            ret = ConvertFloorCeling(price * 0.1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
        Else
            ret = ConvertFloorCeling(price * 0.05 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            If signalCandle.High > _weeklyHigh Then
                Dim lastestCandle As Payload = GetCurrentDayCandle(signalCandle)
                If signalCandle.Close > lastestCandle.PreviousCandlePayload.High Then
                    Dim latestPayload As Dictionary(Of Date, Payload) = Utilities.Strings.DeepClone(Of Dictionary(Of Date, Payload))(_eodPayload)
                    latestPayload.Add(lastestCandle.PayloadDate, lastestCandle)

                    If signalCandle.Volume > _smaPayload(signalCandle.PayloadDate) Then
                        Dim emaCls20 As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.EMA_Close_20).Item1
                        If signalCandle.Close > emaCls20 Then
                            Dim emaCls50 As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.EMA_Close_50).Item1
                            If _emaPayload(signalCandle.PayloadDate) > emaCls50 Then
                                Dim macd As Tuple(Of Decimal, Decimal) = GetIndicatorLatestValue(latestPayload, IndicatorType.MACD_26_12_9)
                                If macd.Item1 > macd.Item2 Then
                                    Dim cci As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.CCI_20).Item1
                                    If cci > 100 Then
                                        If _rsiPayload(signalCandle.PayloadDate) > 60 Then
                                            If (signalCandle.High / lastestCandle.Low) - 1 <= 1 Then
                                                If signalCandle.Close >= 100 Then
                                                    If signalCandle.Close > _vwapPayload(signalCandle.PayloadDate) Then
                                                        If signalCandle.Close > signalCandle.Open Then
                                                            Dim remark As String = String.Format("5 Minute High({0})>Weekly High({1}).{2}5 Minute Close({3})>1 Day Ago High({4}).{5}5 Minute Volume({6})>5 Minute SMA Volume_20({7}).{8}5 Minute Close({9})>Latest EMA Close_20({10}).{11}5 Minute EMA Close_20({12})>Latest EMA Close_50({13}).{14}Latest MACD Line({15})>Latest MACD Signal({16}).{17}Latest CCI({18})>100.{19}5 Minute RSI({20})>60.{21}(5 Minute High({22})/Latest Low({23}))-1<=1.015.{24}Latest Close({25})>=100.{26}5 Minute Close({27})>5 Minute VWAP({28}).{29}5 Minute Close({30})>5 Minute Open({31}).",
                                                                                                signalCandle.High, _weeklyHigh,
                                                                                                vbNewLine, signalCandle.Close, lastestCandle.PreviousCandlePayload.High,
                                                                                                vbNewLine, signalCandle.Volume, _smaPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, signalCandle.Close, emaCls20,
                                                                                                vbNewLine, _emaPayload(signalCandle.PayloadDate), emaCls50,
                                                                                                vbNewLine, macd.Item1, macd.Item2,
                                                                                                vbNewLine, cci,
                                                                                                vbNewLine, _rsiPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, signalCandle.High, lastestCandle.Low,
                                                                                                vbNewLine, lastestCandle.Close,
                                                                                                vbNewLine, signalCandle.Close, _vwapPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, signalCandle.Close, signalCandle.Open)

                                                            Dim buffer As Decimal = CalculateBuffer(signalCandle.High)
                                                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal)(True, signalCandle.High + buffer, signalCandle, Trade.TradeExecutionDirection.Buy, remark, (signalCandle.High / lastestCandle.Low) - 1)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf signalCandle.Low < _weeklyLow Then
                Dim lastestCandle As Payload = GetCurrentDayCandle(signalCandle)
                If signalCandle.Close < lastestCandle.PreviousCandlePayload.Low Then
                    Dim latestPayload As Dictionary(Of Date, Payload) = Utilities.Strings.DeepClone(Of Dictionary(Of Date, Payload))(_eodPayload)
                    latestPayload.Add(lastestCandle.PayloadDate, lastestCandle)

                    If signalCandle.Volume > _smaPayload(signalCandle.PayloadDate) Then
                        Dim emaCls20 As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.EMA_Close_20).Item1
                        If signalCandle.Close < emaCls20 Then
                            Dim emaCls50 As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.EMA_Close_50).Item1
                            If _emaPayload(signalCandle.PayloadDate) < emaCls50 Then
                                Dim macd As Tuple(Of Decimal, Decimal) = GetIndicatorLatestValue(latestPayload, IndicatorType.MACD_26_12_9)
                                If macd.Item1 < macd.Item2 Then
                                    Dim cci As Decimal = GetIndicatorLatestValue(latestPayload, IndicatorType.CCI_20).Item1
                                    If cci < -100 Then
                                        If _rsiPayload(signalCandle.PayloadDate) < 40 Then
                                            If (lastestCandle.High / signalCandle.Low) - 1 <= 1 Then
                                                If lastestCandle.Close >= 100 Then
                                                    If signalCandle.Close < _vwapPayload(signalCandle.PayloadDate) Then
                                                        If signalCandle.Close < signalCandle.Open Then
                                                            Dim remark As String = String.Format("5 Minute Low({0})<Weekly Low({1}).{2}5 Minute Close({3})<1 Day Ago Low({4}).{5}5 Minute Volume({6})>5 Minute SMA Volume_20({7}).{8}5 Minute Close({9})<Latest EMA Close_20({10}).{11}5 Minute EMA Close_20({12})<Latest EMA Close_50({13}).{14}Latest MACD Line({15})<Latest MACD Signal({16}).{17}Latest CCI({18})<-100.{19}5 Minute RSI({20})<40.{21}(Latest High({22})/5 Minute Low({23}))-1<=1.{24}Latest Close({25})>=100.{26}5 Minute Close({27})<5 Minute VWAP({28}).{29}5 Minute Close({30})<5 Minute Open({31}).",
                                                                                                signalCandle.Low, _weeklyLow,
                                                                                                vbNewLine, signalCandle.Close, lastestCandle.PreviousCandlePayload.Low,
                                                                                                vbNewLine, signalCandle.Volume, _smaPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, signalCandle.Close, emaCls20,
                                                                                                vbNewLine, _emaPayload(signalCandle.PayloadDate), emaCls50,
                                                                                                vbNewLine, macd.Item1, macd.Item2,
                                                                                                vbNewLine, cci,
                                                                                                vbNewLine, _rsiPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, lastestCandle.High, signalCandle.Low,
                                                                                                vbNewLine, lastestCandle.Close,
                                                                                                vbNewLine, signalCandle.Close, _vwapPayload(signalCandle.PayloadDate),
                                                                                                vbNewLine, signalCandle.Close, signalCandle.Open)

                                                            Dim buffer As Decimal = CalculateBuffer(signalCandle.Low)
                                                            ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection, String, Decimal)(True, signalCandle.Low - buffer, signalCandle, Trade.TradeExecutionDirection.Sell, remark, (signalCandle.High / lastestCandle.Low) - 1)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetCurrentDayCandle(ByVal signalCandle As Payload) As Payload
        Dim ret As Payload = Nothing

        Dim open As Decimal = _currentDayPaload.FirstOrDefault.Value.Open
        Dim low As Decimal = _currentDayPaload.Min(Function(x)
                                                       If x.Key <= signalCandle.PayloadDate Then
                                                           Return x.Value.Low
                                                       Else
                                                           Return Decimal.MaxValue
                                                       End If
                                                   End Function)
        Dim high As Decimal = _currentDayPaload.Max(Function(x)
                                                        If x.Key <= signalCandle.PayloadDate Then
                                                            Return x.Value.High
                                                        Else
                                                            Return Decimal.MinValue
                                                        End If
                                                    End Function)
        Dim close As Decimal = signalCandle.Close
        Dim volume As Long = _currentDayPaload.Sum(Function(x)
                                                       If x.Key <= signalCandle.PayloadDate Then
                                                           Return x.Value.Volume
                                                       Else
                                                           Return 0
                                                       End If
                                                   End Function)


        ret = New Payload(Payload.CandleDataSource.Calculated) With {
            .Open = open,
            .Low = low,
            .High = high,
            .Close = close,
            .Volume = volume,
            .CumulativeVolume = volume,
            .TradingSymbol = _tradingSymbol,
            .PayloadDate = _tradingDate.Date,
            .PreviousCandlePayload = _eodPayload.LastOrDefault.Value
        }

        Return ret
    End Function

    Private Function GetIndicatorLatestValue(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal typeOfIndicator As IndicatorType) As Tuple(Of Decimal, Decimal)
        Dim ret As Tuple(Of Decimal, Decimal) = Nothing
        Select Case typeOfIndicator
            Case IndicatorType.EMA_Close_20
                Dim emaPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.EMA.CalculateEMA(20, Payload.PayloadFields.Close, inputPayload, emaPayload)
                ret = New Tuple(Of Decimal, Decimal)(emaPayload.LastOrDefault.Value, Decimal.MinValue)
            Case IndicatorType.EMA_Close_50
                Dim emaPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.EMA.CalculateEMA(50, Payload.PayloadFields.Close, inputPayload, emaPayload)
                ret = New Tuple(Of Decimal, Decimal)(emaPayload.LastOrDefault.Value, Decimal.MinValue)
            Case IndicatorType.MACD_26_12_9
                Dim macdPayload As Dictionary(Of Date, Decimal) = Nothing
                Dim macdSignalPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.MACD.CalculateMACD(12, 26, 9, inputPayload, macdPayload, macdSignalPayload, Nothing)
                ret = New Tuple(Of Decimal, Decimal)(macdPayload.LastOrDefault.Value, macdSignalPayload.LastOrDefault.Value)
            Case IndicatorType.CCI_20
                Dim cciPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.CCI.CalculateCCI(20, inputPayload, cciPayload)
                ret = New Tuple(Of Decimal, Decimal)(cciPayload.LastOrDefault.Value, Decimal.MinValue)
            Case Else
                Throw New NotImplementedException
        End Select
        Return ret
    End Function

    Enum IndicatorType
        EMA_Close_20
        EMA_Close_50
        MACD_26_12_9
        CCI_20
    End Enum
End Class