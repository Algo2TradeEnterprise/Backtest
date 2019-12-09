Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ATRPositionalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Linear = 1
        AP
        GP
    End Enum
    Enum AveragingType
        Averaging = 1
        Pyramiding
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
        Public TypeOfAveraging As AveragingType
        Public TargetMultiplier As Decimal
        Public EntryATRMultiplier As Decimal
        Public PartialExit As Boolean
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _highestPrice As Decimal
    Private ReadOnly _investment As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal highestPrice As Decimal,
                   ByVal investment As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _highestPrice = highestPrice
        _investment = investment
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim monthlyPayload As Dictionary(Of Date, Payload) = Common.ConvertDayPayloadsToMonth(_signalPayload)
        Indicator.ATR.CalculateATR(14, monthlyPayload, _atrPayload)
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

            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer) = GetSignalForEntry(currentTick)
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
                                .Supporting1 = _highestPrice,
                                .Supporting3 = signalReceivedForEntry.Item4,
                                .Supporting4 = _investment,
                                .Supporting5 = signalReceivedForEntry.Item6
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
                            runningTrade.UpdateTrade(Supporting2:=ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor))
                        Next
                    End If
                    parameter.Supporting2 = ConvertFloorCeling(averageTradePrice, Me._parentStrategy.TickSize, RoundOfType.Floor)
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
                Dim atr As Decimal = currentTrade.Supporting3
                Dim averagePrice As Decimal = currentTrade.Supporting2
                If currentDayPayload.High >= averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor) Then
                    Dim price As Decimal = averagePrice + ConvertFloorCeling(atr * _userInputs.TargetMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, String)(True, price, "Target")
                End If
            Else
                Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
                Dim atr As Decimal = currentTrade.Supporting3
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
    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            If currentDayPayload.PreviousCandlePayload IsNot Nothing Then
                Dim quantity As Integer = Math.Floor(_investment / currentDayPayload.Open)
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                If lastExecutedTrade Is Nothing Then
                    Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                    Dim atr As Decimal = ConvertFloorCeling(_atrPayload(previousMonth), Me._parentStrategy.TickSize, RoundOfType.Floor)
                    Dim entryPrice As Decimal = ConvertFloorCeling(_highestPrice - atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    If currentDayPayload.Low <= entryPrice Then
                        ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer)(True, entryPrice, currentDayPayload, atr, quantity, quantity)
                    End If
                Else
                    If _userInputs.TypeOfAveraging = AveragingType.Pyramiding Then
                        Dim multiplier As Integer = 1
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentDayPayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                            If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                                Dim lastHighestPrice As Decimal = lastExecutedTrade.Supporting1
                                If _highestPrice = lastHighestPrice Then
                                    Dim tradeForSameLevel As List(Of Trade) = openActiveTrades.FindAll(Function(x)
                                                                                                           Return CDec(x.Supporting1) = _highestPrice
                                                                                                       End Function)
                                    If tradeForSameLevel IsNot Nothing AndAlso tradeForSameLevel.Count > 0 Then
                                        multiplier = tradeForSameLevel.Count + 1
                                    End If
                                End If
                            End If
                        End If
                        Dim atr As Decimal = lastExecutedTrade.Supporting3
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                            atr = ConvertFloorCeling(_atrPayload(previousMonth), Me._parentStrategy.TickSize, RoundOfType.Floor)
                        End If
                        Dim entryPrice As Decimal = ConvertFloorCeling(_highestPrice - atr * _userInputs.EntryATRMultiplier * multiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        If currentDayPayload.Low <= entryPrice Then
                            If multiplier > 1 Then
                                Select Case _userInputs.QuantityType
                                    Case TypeOfQuantity.AP
                                        quantity = lastExecutedTrade.Quantity + 1
                                    Case TypeOfQuantity.GP
                                        quantity = lastExecutedTrade.Quantity * 2
                                    Case TypeOfQuantity.Linear
                                        quantity = lastExecutedTrade.Quantity
                                End Select
                            ElseIf multiplier = 1 AndAlso lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                quantity = lastExecutedTrade.Quantity
                            End If
                            Dim startingQuantity As Integer = lastExecutedTrade.Supporting5
                            If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                startingQuantity = quantity
                            End If
                            ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer)(True, entryPrice, currentDayPayload, atr, quantity, startingQuantity)
                        End If
                    ElseIf _userInputs.TypeOfAveraging = AveragingType.Averaging Then
                        Dim atr As Decimal = lastExecutedTrade.Supporting3
                        Dim entryPrice As Decimal = ConvertFloorCeling(lastExecutedTrade.EntryPrice - atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            Dim previousMonth As Date = New Date(currentDayPayload.PayloadDate.Year, currentDayPayload.PayloadDate.Month, 1).AddMonths(-1)
                            atr = ConvertFloorCeling(_atrPayload(previousMonth), Me._parentStrategy.TickSize, RoundOfType.Floor)
                            entryPrice = ConvertFloorCeling(lastExecutedTrade.ExitPrice - atr * _userInputs.EntryATRMultiplier, Me._parentStrategy.TickSize, RoundOfType.Floor)
                        End If
                        If currentDayPayload.Low <= entryPrice Then
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
                            Dim startingQuantity As Integer = lastExecutedTrade.Supporting5
                            If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                startingQuantity = quantity
                            End If
                            ret = New Tuple(Of Boolean, Decimal, Payload, Decimal, Integer, Integer)(True, entryPrice, currentDayPayload, atr, quantity, startingQuantity)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function
#End Region

End Class
