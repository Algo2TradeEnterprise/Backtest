Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class BelowFractalLowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) Then
            Dim signal As Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim entryPrice As Decimal = signal.Item2
                Dim quantity As Integer = signal.Item3
                Dim iteration As Integer = signal.Item4
                Dim remarks As String = signal.Item5
                Dim fractalLow As Decimal = signal.Item6
                Dim multipler As Integer = signal.Item7

                parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 1000000000,
                            .Target = .EntryPrice + 1000000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = iteration,
                            .Supporting2 = remarks,
                            .Supporting3 = fractalLow,
                            .Supporting4 = multipler
                        }
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
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

    Private Function GetSignalForEntry(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)
        Dim ret As Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer) = Nothing
        Dim lastTrade As Trade = GetLastOrder(currentCandle)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, currentTick.Open)

        Dim candleToCheck As Payload = currentCandle.PreviousCandlePayload
        Dim lastTradeEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 15, 29, 45)
        If currentTick.PayloadDate >= lastTradeEntryTime Then
            candleToCheck = currentCandle
        End If

        Dim entryPrice As Decimal = candleToCheck.Close
        Dim fractalLow As Decimal = _fractalLowPayload(candleToCheck.PayloadDate)
        Dim atr As Decimal = _atrPayload(candleToCheck.PayloadDate)
        If lastTrade Is Nothing Then
            If candleToCheck.Close < fractalLow Then
                ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "First Trade", fractalLow, 1)
            End If
        Else
            Dim lastUsedFractal As Decimal = lastTrade.Supporting3
            If IsDifferentFractal(fractalLow, candleToCheck.PayloadDate, lastUsedFractal, lastTrade.EntryTime) Then
                If candleToCheck.Close < fractalLow Then
                    If entryPrice < lastTrade.EntryPrice Then
                        If lastTrade.EntryPrice - entryPrice >= atr Then
                            ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "(NF)(Reset) Below last entry", fractalLow, 1)
                        End If
                    Else
                        ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "(NF)(Reset) Above last entry", fractalLow, 1)
                    End If
                End If
            Else
                Dim multiplier As Integer = lastTrade.Supporting4 * 2
                entryPrice = Utilities.Numbers.ConvertFloorCeling(lastTrade.EntryPrice - atr * multiplier, _parentStrategy.TickSize, Utilities.Numbers.NumberManipulation.RoundOfType.Floor)
                If currentTick.Low <= entryPrice Then
                    iteration = Val(lastTrade.Supporting1) + 1
                    quantity = GetQuantity(iteration, entryPrice)
                    ret = New Tuple(Of Boolean, Decimal, Integer, Integer, String, Decimal, Integer)(True, entryPrice, quantity, iteration, "Below last entry", fractalLow, multiplier)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsDifferentFractal(ByVal currentFractal As Decimal, ByVal currentFratalTime As Date,
                                        ByVal lastUsedFractal As Decimal, ByVal lastusedFratalTime As Date) As Boolean
        Dim ret As Boolean = False
        If currentFractal <> lastUsedFractal Then
            ret = True
        Else
            For Each runningFractal In _fractalLowPayload.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                If runningFractal.Key <= currentFratalTime AndAlso runningFractal.Key > lastusedFratalTime Then
                    If runningFractal.Value <> currentFractal Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital * Math.Pow(2, iterationNumber - 1)
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastExecutedTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function
End Class