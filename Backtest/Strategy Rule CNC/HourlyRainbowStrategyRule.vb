Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class HourlyRainbowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfExit
        ATR = 1
        Percentage
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ExitType As TypeOfExit
        Public ExitValue As Decimal
    End Class
#End Region

    Public LastTick As Payload = Nothing
    Public EntrySpotATR As Decimal = Decimal.MinValue

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _rainbowPayload As Dictionary(Of Date, Indicator.RainbowMA)
    Private _lastTrade As Trade = Nothing

    Private ReadOnly _controller As Boolean
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)
        _userInputs = _entities
        If controller = 1 Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _controller Then
            Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
            Indicator.RainbowMovingAverage.CalculateRainbowMovingAverage(7, _signalPayload, _rainbowPayload)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso _userInputs.ExitType = TypeOfExit.ATR AndAlso _controller Then
            Dim entryTrades As List(Of Trade) = GetEntryTrades()
            If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
                Dim atr As Decimal = _atrPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                Dim exitHelper As Decimal = Math.Max(Val(entryTrades.LastOrDefault.Supporting3), atr)
                For Each runningTrade In entryTrades
                    runningTrade.UpdateTrade(Supporting3:=exitHelper)
                Next
            End If
        End If
        If (Me.ForceTakeTrade OrElse Me.ContractRolloverForceEntry) AndAlso Not _controller Then
            Dim totalQuantity As Integer = Me.LotSize
            Dim avgPrice As Decimal = currentTick.Open
            Dim entryATR As Decimal = Me.EntrySpotATR
            Dim exitHelper As Decimal = Decimal.MinValue
            If _userInputs.ExitType = TypeOfExit.ATR Then
                exitHelper = Me.EntrySpotATR
            Else
                exitHelper = currentTick.Open * Me.LotSize
            End If
            Dim tradeID As String = System.Guid.NewGuid.ToString()

            Dim entryTrades As List(Of Trade) = GetEntryTrades()
            If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
                Dim totalValue As Decimal = Val(entryTrades.LastOrDefault.Supporting1) * Val(entryTrades.LastOrDefault.Supporting2)
                Dim totalQty As Decimal = Val(entryTrades.LastOrDefault.Supporting2)

                totalValue += currentTick.Open * Me.LotSize
                totalQty += Me.LotSize
                avgPrice = totalValue / totalQty
                totalQuantity = totalQty

                entryATR = entryTrades.LastOrDefault.Supporting4
                tradeID = entryTrades.LastOrDefault.Supporting5

                If _userInputs.ExitType = TypeOfExit.ATR Then
                    exitHelper = Math.Max(Val(entryTrades.LastOrDefault.Supporting3), Me.EntrySpotATR)
                Else
                    exitHelper += Val(entryTrades.LastOrDefault.Supporting3)
                End If

                For Each runningTrade In entryTrades
                    runningTrade.UpdateTrade(Supporting1:=avgPrice, Supporting2:=totalQuantity, Supporting3:=exitHelper)
                Next
            End If

            If Me.ContractRolloverForceEntry AndAlso _lastTrade IsNot Nothing Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = _lastTrade.Quantity,
                                .Stoploss = .EntryPrice - 100000,
                                .Target = .EntryPrice + 100000,
                                .Buffer = 0,
                                .SignalCandle = currentCandle.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = _lastTrade.Supporting1 + (currentTick.Open - _lastTrade.ExitPrice),
                                .Supporting2 = _lastTrade.Supporting2,
                                .Supporting3 = _lastTrade.Supporting3,
                                .Supporting4 = _lastTrade.Supporting4,
                                .Supporting5 = _lastTrade.Supporting5
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ContractRolloverForceEntry = False
                _lastTrade = Nothing
            ElseIf Me.ForceTakeTrade Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = Me.LotSize,
                                .Stoploss = .EntryPrice - 100000,
                                .Target = .EntryPrice + 100000,
                                .Buffer = 0,
                                .SignalCandle = currentCandle.PreviousCandlePayload,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = avgPrice,
                                .Supporting2 = totalQuantity,
                                .Supporting3 = exitHelper,
                                .Supporting4 = entryATR,
                                .Supporting5 = tradeID
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ForceTakeTrade = False
            End If
        ElseIf _controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.CNC) AndAlso
                currentCandle.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                Dim lastEntryTrade As Trade = GetLastEntryTrades()
                If lastEntryTrade Is Nothing OrElse lastEntryTrade.SignalCandle.PayloadDate <> currentCandle.PreviousCandlePayload.PayloadDate Then
                    If currentCandle.PreviousCandlePayload.CandleColor = Color.Green Then
                        Dim rainbow As Indicator.RainbowMA = _rainbowPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                        Dim atr As Decimal = _atrPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                        If currentCandle.PreviousCandlePayload.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                            If IsValidRainbow(currentCandle) Then
                                If lastEntryTrade Is Nothing OrElse currentCandle.PreviousCandlePayload.Close <= Val(lastEntryTrade.Supporting1) - atr Then
                                    If currentTick.PayloadDate <= currentCandle.PayloadDate.AddMinutes(1) Then
                                        Me.AnotherPairInstrument.ForceTakeTrade = True
                                        CType(Me.AnotherPairInstrument, HourlyRainbowStrategyRule).EntrySpotATR = atr
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

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _userInputs.ExitType = TypeOfExit.Percentage Then
                Dim avgPrice As Decimal = currentTrade.Supporting1
                Dim totalQty As Decimal = currentTrade.Supporting2
                Dim capital As Decimal = currentTrade.Supporting3
                Dim pl As Decimal = _parentStrategy.CalculatePL(TradingSymbol, avgPrice, currentTick.Open, totalQty, Me.LotSize, Trade.TypeOfStock.Futures)
                If pl >= capital * _userInputs.ExitValue / 100 Then
                    ret = New Tuple(Of Boolean, String)(True, "Percentage Target Hit")
                End If
            ElseIf _userInputs.ExitType = TypeOfExit.ATR Then
                Dim avgPrice As Decimal = currentTrade.Supporting1
                Dim atr As Decimal = currentTrade.Supporting4
                If currentTick.Open >= avgPrice + atr * _userInputs.ExitValue Then
                    ret = New Tuple(Of Boolean, String)(True, "ATR Target Hit")
                End If
            End If
            If ret Is Nothing Then
                If (Me.ContractRollover OrElse Me.BlankDayExit) Then
                    'Dim currentMinutePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                    'If currentMinutePayload.PayloadDate = _signalPayload.LastOrDefault.Key Then
                    ret = New Tuple(Of Boolean, String)(True, If(Me.ContractRollover, "Contract Rollover Exit", "Blank Day Exit"))
                    Me.ForceCancellationDone = True
                    If _lastTrade IsNot Nothing Then
                        Dim ttlQty As Integer = _lastTrade.Quantity + currentTrade.Quantity
                        _lastTrade.UpdateTrade(Quantity:=ttlQty)
                    Else
                        _lastTrade = Utilities.Strings.DeepClone(Of Trade)(currentTrade)
                        _lastTrade.UpdateTrade(ExitPrice:=currentTick.Open)
                    End If
                    'End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsValidRainbow(ByVal signalCandle As Payload) As Boolean
        Dim ret As Boolean = False
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key < signalCandle.PreviousCandlePayload.PayloadDate Then
                Dim rainbow As Indicator.RainbowMA = _rainbowPayload(runningPayload.Key)
                If runningPayload.Value.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    If runningPayload.Value.CandleColor = Color.Green Then
                        ret = False
                        Exit For
                    End If
                ElseIf runningPayload.Value.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    If runningPayload.Value.CandleColor = Color.Red Then
                        ret = True
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function

    Private Function IsActiveTrade() As Boolean
        Return (_parentStrategy.IsTradeActive(Me.LastTick, Trade.TypeOfTrade.CNC) OrElse
            _parentStrategy.IsTradeActive(CType(Me.AnotherPairInstrument, HourlyRainbowStrategyRule).LastTick, Trade.TypeOfTrade.CNC))
    End Function

    Private Function GetEntryTrades() As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        Dim myTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(Me.LastTick, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        Dim myFutTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(CType(Me.AnotherPairInstrument, HourlyRainbowStrategyRule).LastTick, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If myTrades IsNot Nothing AndAlso myFutTrades IsNot Nothing Then
            ret = New List(Of Trade)
            ret.AddRange(myTrades)
            ret.AddRange(myFutTrades)
        Else
            If myTrades IsNot Nothing Then ret = myTrades
            If myFutTrades IsNot Nothing Then ret = myFutTrades
        End If
        Return ret
    End Function

    Private Function GetLastEntryTrades() As Trade
        Dim ret As Trade = Nothing
        Dim allTrades As List(Of Trade) = GetEntryTrades()
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            ret = allTrades.OrderBy(Function(x)
                                        Return x.EntryTime
                                    End Function).LastOrDefault
        End If
        Return ret
    End Function
End Class
