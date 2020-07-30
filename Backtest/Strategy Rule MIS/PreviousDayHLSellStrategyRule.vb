Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class PreviousDayHLSellStrategyRule
    Inherits StrategyRule

    Public ReadOnly Property StoplossLevel As Decimal = Decimal.MinValue
    Public DummyCandle As Payload = Nothing
    Private _firstTime As Boolean = True
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _buyLevel As Decimal = Decimal.MinValue
    Private _sellLevel As Decimal = Decimal.MinValue

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
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Sub CompletePairProcessing()
        MyBase.CompletePairProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
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
            If currentTick.TradingSymbol.EndsWith("CE") Then
                Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "CE", TradingSymbol)
                Dim strike As Decimal = strikeData.Substring(5)
                _StoplossLevel = strike + currentTick.Open
            ElseIf currentTick.TradingSymbol.EndsWith("PE") Then
                Dim strikeData As String = Utilities.Strings.GetTextBetween("BANKNIFTY", "PE", TradingSymbol)
                Dim strike As Decimal = strikeData.Substring(5)
                _StoplossLevel = strike - currentTick.Open
            End If
        ElseIf Controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                If _firstTime Then
                    If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
                        For Each runningInstruments In Me.DependentInstrument
                            runningInstruments.EligibleToTakeTrade = True
                            runningInstruments.ForceTakeTrade = True
                        Next
                    End If
                    _firstTime = False
                End If
                If _buyLevel = Decimal.MinValue OrElse _sellLevel = Decimal.MinValue Then
                    If Me.DependentInstrument IsNot Nothing AndAlso Me.DependentInstrument.Count > 0 Then
                        For Each runningInstruments In Me.DependentInstrument
                            If CType(runningInstruments, PreviousDayHLSellStrategyRule).StoplossLevel <> Decimal.MinValue AndAlso
                            CType(runningInstruments, PreviousDayHLSellStrategyRule).DummyCandle IsNot Nothing Then
                                If CType(runningInstruments, PreviousDayHLSellStrategyRule).DummyCandle.TradingSymbol.EndsWith("CE") Then
                                    _buyLevel = CType(runningInstruments, PreviousDayHLSellStrategyRule).StoplossLevel
                                ElseIf CType(runningInstruments, PreviousDayHLSellStrategyRule).DummyCandle.TradingSymbol.EndsWith("PE") Then
                                    _sellLevel = CType(runningInstruments, PreviousDayHLSellStrategyRule).StoplossLevel
                                End If
                            End If
                        Next
                    End If
                Else
                    If Not _parentStrategy.IsTradeOpen(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) AndAlso
                        Not _parentStrategy.IsTradeActive(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS) Then
                        Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                        If currentTick.Open >= _buyLevel + atr Then
                            Dim parameter As PlaceOrderParameters = Nothing
                            parameter = New PlaceOrderParameters With {
                                            .EntryPrice = currentTick.Open,
                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                            .Quantity = Me.LotSize,
                                            .Stoploss = _buyLevel - atr,
                                            .Target = .EntryPrice + 10000000,
                                            .Buffer = 0,
                                            .SignalCandle = currentMinuteCandlePayload,
                                            .OrderType = Trade.TypeOfOrder.Market,
                                            .Supporting1 = ""
                                        }

                            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                        ElseIf currentTick.Open <= _sellLevel - atr Then
                            Dim parameter As PlaceOrderParameters = Nothing
                            parameter = New PlaceOrderParameters With {
                                            .EntryPrice = currentTick.Open,
                                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                            .Quantity = Me.LotSize,
                                            .Stoploss = _sellLevel + atr,
                                            .Target = .EntryPrice - 10000000,
                                            .Buffer = 0,
                                            .SignalCandle = currentMinuteCandlePayload,
                                            .OrderType = Trade.TypeOfOrder.Market,
                                            .Supporting1 = ""
                                        }

                            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Me.DummyCandle = currentTick
        Await Task.Delay(0).ConfigureAwait(False)
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