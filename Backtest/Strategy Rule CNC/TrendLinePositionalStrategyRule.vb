Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class TrendLinePositionalStrategyRule
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
        Public TargetMultiplier As Decimal
        Public PartialExit As Boolean
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _swingHighTrendLine As Dictionary(Of Date, TrendLineVeriables)
    Private _swingLowTrendLine As Dictionary(Of Date, TrendLineVeriables)

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.SwingHighLowTrendLine.CalculateSwingHighLowTrendLine(_signalPayload, True, _swingHighTrendLine, _swingLowTrendLine)
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

            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                signalCandle = signalReceivedForEntry.Item3

                If signalCandle IsNot Nothing Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalReceivedForEntry.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = signalReceivedForEntry.Item5,
                                .Stoploss = .EntryPrice - 1000000,
                                .Target = .EntryPrice + 1000000,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting2 = signalReceivedForEntry.Item4,
                                .Supporting3 = signalReceivedForEntry.Item6,
                                .Supporting4 = signalReceivedForEntry.Item7.M,
                                .Supporting5 = signalReceivedForEntry.Item7.C
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
                            runningTrade.UpdateTrade(Supporting1:=ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor))
                        Next
                    End If
                    parameter.Supporting1 = ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Not _userInputs.PartialExit Then
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                Dim atr As Decimal = currentTrade.Supporting2
                Dim averagePrice As Decimal = currentTrade.Supporting1
                If currentDayPayload.High >= averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                    Dim price As Decimal = averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, String)(True, price, "Target")
                End If
            Else
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                Dim atr As Decimal = currentTrade.Supporting2
                Dim entryPrice As Decimal = currentTrade.EntryPrice
                If currentDayPayload.High >= entryPrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                    Dim price As Decimal = entryPrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, String)(True, price, "Target")
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

#Region "Entry Rule"
    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            If currentDayPayload.PreviousCandlePayload IsNot Nothing Then
                Dim quantity As Integer = 1
                Dim currentDayTrendLine As TrendLineVeriables = _swingHighTrendLine(currentDayPayload.PayloadDate)
                If currentDayTrendLine IsNot Nothing AndAlso currentDayTrendLine.M <> Decimal.MinValue Then
                    Dim currentDayEntryPrice As Decimal = currentDayTrendLine.M * currentDayTrendLine.X + currentDayTrendLine.C
                    Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                    If lastExecutedTrade Is Nothing Then
                        Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentDayPayload.PreviousCandlePayload.PayloadDate), Me._parentStrategy.TickSize, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(currentDayEntryPrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        If currentDayPayload.High >= entryPrice AndAlso currentDayPayload.Low <= entryPrice Then
                            ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables)(True, entryPrice, currentDayPayload, atr, quantity, quantity, currentDayTrendLine)
                        End If
                    Else
                        If lastExecutedTrade.Supporting4 <> currentDayTrendLine.M OrElse lastExecutedTrade.Supporting5 <> currentDayTrendLine.C Then
                            Dim atr As Decimal = lastExecutedTrade.Supporting2
                            Dim entryPrice As Decimal = ConvertFloorCeling(currentDayEntryPrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
                            If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                atr = ConvertFloorCeling(_atrPayload(currentDayPayload.PreviousCandlePayload.PayloadDate), Me._parentStrategy.TickSize, RoundOfType.Floor)
                            End If
                            If currentDayPayload.High >= entryPrice AndAlso currentDayPayload.Low <= entryPrice Then
                                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                    Select Case _userInputs.QuantityType
                                        Case TypeOfQuantity.AP
                                            quantity = lastExecutedTrade.Quantity + 1
                                        Case TypeOfQuantity.GP
                                            quantity = lastExecutedTrade.Quantity * 2
                                        Case TypeOfQuantity.Linear
                                            quantity = lastExecutedTrade.Quantity
                                    End Select
                                End If
                                Dim startingQuantity As Integer = lastExecutedTrade.Supporting3
                                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                    startingQuantity = quantity
                                End If
                                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                    ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables)(True, entryPrice, currentDayPayload, atr, quantity, startingQuantity, currentDayTrendLine)
                                ElseIf lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                    If entryPrice < lastExecutedTrade.EntryPrice Then
                                        ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer, TrendLineVeriables)(True, entryPrice, currentDayPayload, atr, quantity, startingQuantity, currentDayTrendLine)
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
#End Region

End Class