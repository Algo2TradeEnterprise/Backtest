Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class NikhilCrudeOilStrategy
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public StoplossPoint As Decimal
        Public TargetPoint As Decimal
        Public Buffer As Decimal
        Public InitialNumberOfLots As Integer
        Public WithTimeConstraint As Boolean
    End Class
#End Region

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection) = GetEntrySignal(currentCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastSignalEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, currentCandle.PayloadDate.Hour, 45, 0)
                If Not _userInputs.WithTimeConstraint OrElse currentTick.PayloadDate < lastSignalEntryTime Then
                    If signal.Item4 = Trade.TradeExecutionDirection.Buy AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signal.Item2,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = signal.Item3,
                                                                .Stoploss = .EntryPrice - _userInputs.StoplossPoint,
                                                                .Target = .EntryPrice + _userInputs.TargetPoint,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentCandle.PreviousCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.SL,
                                                                .Supporting1 = currentCandle.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss")
                                                            }

                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell AndAlso
                    Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = signal.Item2,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                .Quantity = signal.Item3,
                                                                .Stoploss = .EntryPrice + _userInputs.StoplossPoint,
                                                                .Target = .EntryPrice - _userInputs.TargetPoint,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentCandle.PreviousCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.SL,
                                                                .Supporting1 = currentCandle.PreviousCandlePayload.PayloadDate.ToString("HH:mm:ss")
                                                            }

                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing Then
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim exitTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, currentTrade.EntryTime.Hour, 59, 45)
            If currentTick.PayloadDate >= exitTime Then
                ret = New Tuple(Of Boolean, String)(True, "Hour Exit")
            Else
                If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                    Dim oppositeTrade As Trade = GetOppositeTrade(currentTrade.SignalCandle, currentTrade.EntryDirection)
                    If oppositeTrade IsNot Nothing AndAlso oppositeTrade.ExitCondition = Trade.TradeExitCondition.Target AndAlso
                        currentTrade.Quantity >= oppositeTrade.Quantity * 2 Then
                        ret = New Tuple(Of Boolean, String)(True, "Quantity Correction")
                    End If
                End If
            End If
            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open AndAlso _userInputs.WithTimeConstraint Then
                Dim lastSignalEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, currentCandle.PayloadDate.Hour, 45, 0)
                If currentTick.PayloadDate >= lastSignalEntryTime Then
                    ret = New Tuple(Of Boolean, String)(True, "Time Constraint Exit")
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            currentCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            Dim signalTrades As List(Of Trade) = GetAllTrades(signalCandle)
            If signalTrades IsNot Nothing AndAlso signalTrades.Count > 0 AndAlso signalTrades.Count < 2 Then
                If signalCandle.CandleColor = Color.Green Then
                    Dim oppositeTrade As Trade = GetOppositeTrade(signalCandle, Trade.TradeExecutionDirection.Sell)
                    If oppositeTrade IsNot Nothing Then
                        If oppositeTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, oppositeTrade.PotentialStopLoss, Me.LotSize * _userInputs.InitialNumberOfLots * 2, Trade.TradeExecutionDirection.Sell)
                        ElseIf oppositeTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            If oppositeTrade.ExitCondition = Trade.TradeExitCondition.Target Then
                                ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, oppositeTrade.PotentialStopLoss, Me.LotSize * _userInputs.InitialNumberOfLots, Trade.TradeExecutionDirection.Sell)
                            End If
                        End If
                    End If
                ElseIf signalCandle.CandleColor = Color.Red Then
                    Dim oppositeTrade As Trade = GetOppositeTrade(signalCandle, Trade.TradeExecutionDirection.Buy)
                    If oppositeTrade IsNot Nothing Then
                        If oppositeTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, oppositeTrade.PotentialStopLoss, Me.LotSize * _userInputs.InitialNumberOfLots * 2, Trade.TradeExecutionDirection.Buy)
                        ElseIf oppositeTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            If oppositeTrade.ExitCondition = Trade.TradeExitCondition.Target Then
                                ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, oppositeTrade.PotentialStopLoss, Me.LotSize * _userInputs.InitialNumberOfLots, Trade.TradeExecutionDirection.Buy)
                            End If
                        End If
                    End If
                End If
            ElseIf signalTrades Is Nothing OrElse signalTrades.Count = 0 Then
                If signalCandle.CandleColor = Color.Green Then
                    ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, signalCandle.Close + _userInputs.Buffer, Me.LotSize * _userInputs.InitialNumberOfLots, Trade.TradeExecutionDirection.Buy)
                ElseIf signalCandle.CandleColor = Color.Red Then
                    ret = New Tuple(Of Boolean, Decimal, Integer, Trade.TradeExecutionDirection)(True, signalCandle.Close - _userInputs.Buffer, Me.LotSize * _userInputs.InitialNumberOfLots, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAllTrades(ByVal signalCandle As Payload) As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Close)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If openTrades IsNot Nothing Then allTrades.AddRange(openTrades)
        If inprogressTrades IsNot Nothing Then allTrades.AddRange(inprogressTrades)
        If closeTrades IsNot Nothing Then allTrades.AddRange(closeTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            ret = allTrades.FindAll(Function(x)
                                        Return x.SignalCandle.PayloadDate = signalCandle.PayloadDate
                                    End Function)
        End If
        Return ret
    End Function

    Private Function GetOppositeTrade(ByVal signalCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Trade
        Dim ret As Trade = Nothing
        Dim signalTrades As List(Of Trade) = GetAllTrades(signalCandle)
        If signalTrades IsNot Nothing AndAlso signalTrades.Count > 0 Then
            ret = signalTrades.Find(Function(x)
                                        Return x.EntryDirection <> direction
                                    End Function)
        End If
        Return ret
    End Function
End Class