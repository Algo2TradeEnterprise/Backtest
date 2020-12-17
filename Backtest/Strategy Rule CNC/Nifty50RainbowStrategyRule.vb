Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class Nifty50RainbowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
        Public MaxIteration As Integer
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _nifty50Payload As Dictionary(Of Date, Payload) = Nothing
    Private _rainbowPayload As Dictionary(Of Date, Indicator.RainbowMA) = Nothing

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

        _nifty50Payload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, "NIFTY 50", _tradingDate.AddYears(-2), _tradingDate)
        Indicator.RainbowMovingAverage.CalculateRainbowMovingAverage(2, _nifty50Payload, _rainbowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If currentDayCandlePayload IsNot Nothing AndAlso currentDayCandlePayload.PreviousCandlePayload IsNot Nothing Then
            Dim signals As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)) = GetSignalForEntry(currentDayCandlePayload)
            If signals IsNot Nothing AndAlso signals.Count > 0 Then
                For Each runningSignal In signals
                    Dim entryPrice As Decimal = runningSignal.Item2
                    Dim quantity As Integer = runningSignal.Item3
                    Dim iteration As Integer = runningSignal.Item4
                    Dim remarks As String = runningSignal.Item5

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = entryPrice,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = iteration,
                                                                .Supporting2 = remarks,
                                                                .Supporting3 = runningSignal.Item6.ToString("dd-MMM-yyyy HH:mm:ss")
                                                            }

                    If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
                    parameters.Add(parameter)
                Next
            End If
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
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

    Private Function GetSignalForEntry(ByVal candle As Payload) As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
        Dim ret As List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, candle.Open)
        Dim rainbow As Indicator.RainbowMA = _rainbowPayload(candle.PayloadDate)
        Dim nifty50Candle As Payload = _nifty50Payload(candle.PayloadDate)
        If nifty50Candle.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
            Dim previousOutsideRainbow As Tuple(Of Trade.TradeExecutionDirection, Date) = GetLastOutsideRainbow(candle)
            If previousOutsideRainbow IsNot Nothing AndAlso previousOutsideRainbow.Item1 = Trade.TradeExecutionDirection.Sell AndAlso previousOutsideRainbow.Item2 <> Date.MinValue Then
                Dim entryPrice As Decimal = candle.Close
                If lastTrade Is Nothing Then
                    If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                    ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "First Trade", previousOutsideRainbow.Item2))
                Else
                    Dim lastSignalTime As Date = Date.ParseExact(lastTrade.Supporting3, "dd-MMM-yyyy HH:mm:ss", Nothing)
                    If previousOutsideRainbow.Item2 <> lastSignalTime Then
                        If entryPrice < lastTrade.EntryPrice Then
                            If lastTrade.EntryPrice - entryPrice >= atr Then
                                If Val(lastTrade.Supporting1) < _userInputs.MaxIteration Then
                                    iteration = Val(lastTrade.Supporting1) + 1
                                    quantity = GetQuantity(iteration, entryPrice)
                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                                    ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "Below last entry", previousOutsideRainbow.Item2))
                                Else
                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                                    ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "(Reset) Max Iteration", previousOutsideRainbow.Item2))
                                End If
                                'Else
                                '    Console.WriteLine(String.Format("Trade Neglected for ATR on {0}", candle.PayloadDate.ToString("dd-MMM-yyyy")))
                            End If
                        Else
                            If ret Is Nothing Then ret = New List(Of Tuple(Of Boolean, Decimal, Integer, Integer, String, Date))
                            ret.Add(New Tuple(Of Boolean, Decimal, Integer, Integer, String, Date)(True, entryPrice, quantity, iteration, "(Reset) Above last entry", previousOutsideRainbow.Item2))
                        End If
                    End If
                End If
            End If
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

    Private Function GetLastOutsideRainbow(ByVal candle As Payload) As Tuple(Of Trade.TradeExecutionDirection, Date)
        Dim ret As Tuple(Of Trade.TradeExecutionDirection, Date) = Nothing
        For Each runningPayload In _nifty50Payload.OrderByDescending(Function(x)
                                                                         Return x.Key
                                                                     End Function)
            If runningPayload.Key <= candle.PreviousCandlePayload.PayloadDate Then
                Dim rainbow As Indicator.RainbowMA = _rainbowPayload(runningPayload.Key)
                If runningPayload.Value.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Buy, runningPayload.Key)
                    Exit For
                ElseIf runningPayload.Value.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    ret = New Tuple(Of Trade.TradeExecutionDirection, Date)(Trade.TradeExecutionDirection.Sell, runningPayload.Key)
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function
End Class