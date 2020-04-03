Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FractalTrendLineStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TrendLineType
        FractalU = 1
        FractalNormal
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TypeOfTrendLine As TrendLineType
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _fractalHighTrendLine As Dictionary(Of Date, TrendLineVeriables)
    Private _fractalLowTrendLine As Dictionary(Of Date, TrendLineVeriables)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _tradeStartTime As Date
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        If _userInputs.TypeOfTrendLine = TrendLineType.FractalNormal Then
            Indicator.FractalBandsTrendLine.CalculateFractalBandsTrendLine(_signalPayload, _fractalHighTrendLine, _fractalLowTrendLine, _fractalHighPayload, _fractalLowPayload)
        ElseIf _userInputs.TypeOfTrendLine = TrendLineType.FractalU Then
            Indicator.FractalUTrendLive.CalculateFractalUTrendLine(_signalPayload, _fractalHighTrendLine, _fractalLowTrendLine, _fractalHighPayload, _fractalLowPayload)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim signal As Tuple(Of Boolean, Decimal, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)

                    parameter1 = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2 + buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = Me.LotSize,
                                .Stoploss = .EntryPrice - 10000000000000,
                                .Target = .EntryPrice + 10000000000000,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item3
                            }
                End If
            End If
            If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim signal As Tuple(Of Boolean, Decimal, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)

                    parameter2 = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2 - buffer,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = Me.LotSize,
                                .Stoploss = .EntryPrice + 10000000000000,
                                .Target = .EntryPrice - 10000000000000,
                                .Buffer = buffer,
                                .SignalCandle = currentMinuteCandlePayload.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item3
                            }
                End If
            End If
        End If
        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter2)
        End If
        If parameterList IsNot Nothing AndAlso parameterList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameterList)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, String) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Decimal, String)
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        If candle IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                Dim trendline As TrendLineVeriables = _fractalHighTrendLine(candle.PayloadDate)
                Dim fractalEntryPoint As Decimal = _fractalHighPayload(candle.PayloadDate)
                Dim anchorCandle As Payload = Common.GetPayloadAt(_signalPayload, candle.PayloadDate, (trendline.X + 1) * -1).Value.Value
                Dim trendlinePoint As Decimal = ConvertFloorCeling(trendline.M * trendline.X + trendline.C, _parentStrategy.TickSize, RoundOfType.Floor)
                If fractalEntryPoint > trendlinePoint Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, fractalEntryPoint, "Fractal")
                ElseIf candle.High > trendlinePoint Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, candle.High, "Candle close above trendline")
                Else
                    ret = New Tuple(Of Boolean, Decimal, String)(True, anchorCandle.High, "Anchor candle")
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                Dim trendline As TrendLineVeriables = _fractalLowTrendLine(candle.PayloadDate)
                Dim fractalEntryPoint As Decimal = _fractalLowPayload(candle.PayloadDate)
                Dim anchorCandle As Payload = Common.GetPayloadAt(_signalPayload, candle.PayloadDate, (trendline.X + 1) * -1).Value.Value
                Dim trendlinePoint As Decimal = ConvertFloorCeling(trendline.M * trendline.X + trendline.C, _parentStrategy.TickSize, RoundOfType.Floor)
                If fractalEntryPoint < trendlinePoint Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, fractalEntryPoint, "Fractal")
                ElseIf candle.Low < trendlinePoint Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, candle.Low, "Candle close below trendline")
                Else
                    ret = New Tuple(Of Boolean, Decimal, String)(True, anchorCandle.Low, "Anchor candle")
                End If
            End If
        End If
        Return ret
    End Function
End Class