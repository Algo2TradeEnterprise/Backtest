Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SwingCNCStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Flat = 1
        AP
        GP
        Misc
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialCapital As Integer
        Public QuantityType As TypeOfQuantity
        Public MaxIteration As Integer
    End Class
#End Region

    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private ReadOnly _tradeStartTime As Date
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
        _userInputs = entities
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        _lastTradeEntryTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime Then
            Dim signal As Tuple(Of Boolean, Payload, Integer, Integer, Date, String) = GetSignalForEntry(currentMinuteCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim quantity As Integer = signal.Item3
                Dim iteration As Integer = signal.Item4

                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 1000000000,
                            .Target = .EntryPrice + 1000000000,
                            .Buffer = 0,
                            .SignalCandle = signal.Item2,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = iteration,
                            .Supporting2 = signal.Item6,
                            .Supporting3 = signal.Item5.ToString("dd-MMM-yyyy HH:mm:ss")
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Integer, Integer, Date, String)
        Dim ret As Tuple(Of Boolean, Payload, Integer, Integer, Date, String) = Nothing
        Dim lastTrade As Trade = GetLastOrder(candle)
        Dim atr As Decimal = _atrPayload(candle.PreviousCandlePayload.PayloadDate)
        Dim iteration As Integer = 1
        Dim quantity As Integer = GetQuantity(iteration, currentTick.Open)
        If candle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
            Dim payloadToCheck As Payload = candle.PreviousCandlePayload
            If currentTick.PayloadDate >= New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 15, 29, 30) Then
                payloadToCheck = candle
            End If
            Dim swingData As Indicator.Swing = _swingPayload(payloadToCheck.PayloadDate)
            atr = _atrPayload(payloadToCheck.PayloadDate)
            If payloadToCheck.Close <= swingData.SwingLow Then
                If lastTrade IsNot Nothing Then
                    Dim lastSignalTime As Date = Date.ParseExact(lastTrade.Supporting3, "dd-MMM-yyyy HH:mm:ss", Nothing)
                    If swingData.SwingLowTime <> lastSignalTime Then
                        If payloadToCheck.Close < lastTrade.EntryPrice Then
                            If lastTrade.EntryPrice - payloadToCheck.Close >= atr Then
                                iteration = Val(lastTrade.Supporting1) + 1
                                quantity = GetQuantity(iteration, currentTick.Open)
                                ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, candle, quantity, iteration, swingData.SwingLowTime, "Below last entry")
                            End If
                        Else
                            If payloadToCheck.Open > swingData.SwingLow Then
                                ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, candle, quantity, iteration, swingData.SwingLowTime, "(Reset) Above last entry")
                            End If
                        End If
                    End If
                Else
                    ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, payloadToCheck, quantity, iteration, swingData.SwingLowTime, "First Trade")
                End If
            End If
        Else
            If ((candle.PreviousCandlePayload.Close - candle.Open) / candle.PreviousCandlePayload.Close) * 100 >= 1 Then
                Dim openTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, candle.PayloadDate.Hour, candle.PayloadDate.Minute, 9)
                If lastTrade IsNot Nothing Then
                    Dim lastSignalTime As Date = Date.ParseExact(lastTrade.Supporting3, "dd-MMM-yyyy HH:mm:ss", Nothing)
                    If openTime <> lastSignalTime Then
                        If candle.Open < lastTrade.EntryPrice Then
                            If lastTrade.EntryPrice - candle.Open >= atr Then
                                iteration = Val(lastTrade.Supporting1) + 1
                                quantity = GetQuantity(iteration, currentTick.Open)
                                ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, candle, quantity, iteration, openTime, "Below last entry")
                            End If
                        Else
                            ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, candle, quantity, iteration, openTime, "(Reset) Above last entry")
                        End If
                    End If
                Else
                    ret = New Tuple(Of Boolean, Payload, Integer, Integer, Date, String)(True, candle, quantity, iteration, openTime, "First Trade")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Integer
        Dim capital As Decimal = _userInputs.InitialCapital
        Select Case _userInputs.QuantityType
            Case TypeOfQuantity.AP
                capital = _userInputs.InitialCapital * iterationNumber
            Case TypeOfQuantity.GP
                capital = _userInputs.InitialCapital * (Math.Pow(2, (iterationNumber - 1)))
            Case TypeOfQuantity.Misc
                Dim multiplier As Integer = Math.Floor(iterationNumber / _userInputs.MaxIteration)
                capital = _userInputs.InitialCapital * Math.Pow(_userInputs.MaxIteration, multiplier)
        End Select
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            ret = Me._parentStrategy.GetLastEntryTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
        End If
        Return ret
    End Function
End Class
