Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TIICNCStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Linear = 1
        AP
        GP
    End Enum

    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
        Public QuntityForLinear As Integer
        Public AdditionalProfitPercentage As Decimal
    End Class
#End Region

    Private _tiiPayload As Dictionary(Of Date, Decimal)
    Private _signalLinePayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)

    Private _doneForCurrentDay As Boolean = False

    Private ReadOnly _userInputs As StrategyRuleEntities

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

        Indicator.TrendIntensityIndex.CalculateTII(Payload.PayloadFields.Close, 20, 1, _signalPayload, _tiiPayload, _signalLinePayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload, Boolean) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload)
            If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close AndAlso
                lastExecutedTrade.ExitTime >= currentMinuteCandlePayload.PayloadDate Then
                signalReceivedForEntry = Nothing
            End If
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                signalCandle = signalReceivedForEntry.Item3

                If signalCandle IsNot Nothing Then
                    Dim lastEntryPrice As Decimal = Decimal.MinValue
                    Dim lastEntryATR As Decimal = Decimal.MinValue
                    Dim lowestATR As Decimal = _atrPayload(signalCandle.PayloadDate)
                    Dim firstEntryDate As String = currentTick.PayloadDate.ToString("dd-MM-yyyy")
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        lastEntryPrice = lastExecutedTrade.EntryPrice
                        lastEntryATR = lastExecutedTrade.Supporting2
                        lowestATR = lastExecutedTrade.Supporting3
                        firstEntryDate = lastExecutedTrade.Supporting4
                    End If
                    If lastEntryPrice = Decimal.MinValue OrElse signalReceivedForEntry.Item2 <= lastEntryPrice - lastEntryATR Then
                        Dim quantity As Integer = 1
                        If lastEntryPrice <> Decimal.MinValue Then
                            Select Case _userInputs.QuantityType
                                Case TypeOfQuantity.AP
                                    quantity = lastExecutedTrade.Quantity + 1
                                Case TypeOfQuantity.GP
                                    quantity = lastExecutedTrade.Quantity * 2
                                Case TypeOfQuantity.Linear
                                    quantity = _userInputs.QuntityForLinear
                            End Select
                        End If

                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalReceivedForEntry.Item2, RoundOfType.Floor)
                        If signalReceivedForEntry.Item4 OrElse currentMinuteCandlePayload.High >= signalReceivedForEntry.Item2 + buffer Then
                            If signalReceivedForEntry.Item4 Then buffer = 0
                            parameter = New PlaceOrderParameters With {
                                        .EntryPrice = signalReceivedForEntry.Item2 + buffer,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .Quantity = quantity,
                                        .Stoploss = .EntryPrice - 1000000,
                                        .Target = .EntryPrice + 1000000,
                                        .Buffer = buffer,
                                        .SignalCandle = signalCandle,
                                        .OrderType = Trade.TypeOfOrder.Market,
                                        .Supporting2 = _atrPayload(signalCandle.PayloadDate),
                                        .Supporting3 = Math.Min(lowestATR, _atrPayload(signalCandle.PayloadDate)),
                                        .Supporting4 = firstEntryDate
                                    }

                            Dim totalCapitalUsedWithoutMargin As Decimal = 0
                            Dim totalQuantity As Decimal = 0
                            Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                            If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                                For Each runningTrade In openActiveTrades
                                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                        totalCapitalUsedWithoutMargin += runningTrade.EntryPrice * runningTrade.Quantity
                                        totalQuantity += runningTrade.Quantity
                                    End If
                                Next
                            End If
                            totalCapitalUsedWithoutMargin += parameter.EntryPrice * parameter.Quantity
                            totalQuantity += parameter.Quantity
                            Dim averageTradePrice As Decimal = totalCapitalUsedWithoutMargin / totalQuantity
                            If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                                For Each runningTrade In openActiveTrades
                                    runningTrade.UpdateTrade(Supporting1:=ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor),
                                                             Supporting3:=parameter.Supporting3)
                                Next
                            End If
                            parameter.Supporting1 = ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        End If
                    End If
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        'If Not _doneForCurrentDay Then
        '    Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentTick, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
        '    If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
        '        Dim additionalProfit As Decimal = 0
        '        For Each runningTrade In openActiveTrades
        '            Dim dayDifference As Long = DateDiff(DateInterval.Day, runningTrade.EntryTime.Date, currentTick.PayloadDate.Date) - 1
        '            If dayDifference > 0 Then
        '                additionalProfit += ConvertFloorCeling(CDec(runningTrade.Supporting2) * dayDifference * _userInputs.AdditionalProfitPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
        '            End If
        '        Next
        '        For Each runningTrade In openActiveTrades
        '            runningTrade.UpdateTrade(Supporting4:=ConvertFloorCeling(additionalProfit, Me._parentStrategy.TickSize, RoundOfType.Floor))
        '        Next
        '    End If
        '    _doneForCurrentDay = True
        'End If
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim lowestATR As Decimal = ConvertFloorCeling(currentTrade.Supporting3, _parentStrategy.TickSize, RoundOfType.Floor)
            Dim averagePrice As Decimal = currentTrade.Supporting1
            Dim startingDay As Date = Convert.ToDateTime(currentTrade.Supporting4)
            Dim dayDifference As Long = DateDiff(DateInterval.Day, startingDay.Date, currentTick.PayloadDate.Date) + 1
            Dim additionalProfit As Decimal = ConvertFloorCeling(Math.Log(dayDifference), _parentStrategy.TickSize, RoundOfType.Floor)
            If currentTick.High >= averagePrice + lowestATR + additionalProfit Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, averagePrice + lowestATR + additionalProfit, "Target")
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

#Region "Entry Rule"
    Private Function GetSignalForEntry(ByVal candle As Payload) As Tuple(Of Boolean, Decimal, Payload, Boolean)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Boolean) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If candle.High > candle.PreviousCandlePayload.High Then
                If _tiiPayload(candle.PayloadDate) < 20 AndAlso _tiiPayload(candle.PreviousCandlePayload.PayloadDate) < 20 Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Boolean)(True, candle.High, candle, False)
                End If
            ElseIf _tiiPayload(candle.PayloadDate) > 20 AndAlso _tiiPayload(candle.PayloadDate) < 80 AndAlso
                _tiiPayload(candle.PreviousCandlePayload.PayloadDate) < 20 Then
                ret = New Tuple(Of Boolean, Decimal, Payload, Boolean)(True, candle.Close, candle, True)
            End If
        End If
        Return ret
    End Function
#End Region

End Class
