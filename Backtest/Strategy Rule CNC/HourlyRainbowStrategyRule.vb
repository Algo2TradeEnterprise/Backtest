Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper

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

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _rainbowPayload As Dictionary(Of Date, Indicator.RainbowMA)

    Private _dependentInstruments As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
    Private _strikeGap As Decimal = 50

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal nextTradingDay As Date,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, nextTradingDay, entities, canceller)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.RainbowMovingAverage.CalculateRainbowMovingAverage(7, _signalPayload, _rainbowPayload)

        Dim sampleCandle As Payload = _signalPayload.LastOrDefault.Value
        Dim runningTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(sampleCandle, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If runningTrades IsNot Nothing AndAlso runningTrades.Count > 0 Then
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            For Each runningTrade In runningTrades
                If Not _dependentInstruments.ContainsKey(runningTrade.SupportingTradingSymbol) Then
                    Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                    If runningTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                        inputPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, runningTrade.SupportingTradingSymbol, _tradingDate.AddDays(-5), _tradingDate)
                    Else
                        inputPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, runningTrade.SupportingTradingSymbol, _tradingDate.AddDays(-5), _tradingDate)
                    End If
                    If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                        _dependentInstruments.Add(runningTrade.SupportingTradingSymbol, inputPayload)
                    End If
                End If
            Next
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.CNC) AndAlso
            currentCandle.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim lastEntryTrade As Trade = _parentStrategy.GetLastEntryTradeOfTheStock(currentCandle, Trade.TypeOfTrade.CNC)
            If lastEntryTrade Is Nothing OrElse lastEntryTrade.SignalCandle.PayloadDate <> currentCandle.PreviousCandlePayload.PayloadDate Then
                If currentCandle.PreviousCandlePayload.CandleColor = Color.Green Then
                    Dim rainbow As Indicator.RainbowMA = _rainbowPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                    Dim atr As Decimal = _atrPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                    If currentCandle.PreviousCandlePayload.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                        If IsValidRainbow(currentCandle) Then
                            If lastEntryTrade Is Nothing OrElse lastEntryTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close OrElse
                                currentCandle.PreviousCandlePayload.Close <= Val(lastEntryTrade.EntryPrice) - atr Then
                                If Not EnterTrade(currentCandle.PreviousCandlePayload, currentTick) Then
                                    Throw New NotImplementedException()
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
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If _userInputs.ExitType = TypeOfExit.Percentage Then
                Dim allTrades As List(Of Trade) = _parentStrategy.GetAllTradesByTag(currentTrade.Tag, currentTick)
                If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                    Dim totalPL As Decimal = 0
                    For Each runningTrade In allTrades
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                            totalPL += _parentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTick.PayloadDate)
                            totalPL += _parentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, runningTrade.EntryPrice, currentFOTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                        End If
                    Next
                    If totalPL >= currentTrade.Supporting5 * 30000 * _userInputs.ExitValue / 100 Then
                        ret = New Tuple(Of Boolean, String)(True, "Percentage Target Hit")
                    End If
                End If
            ElseIf _userInputs.ExitType = TypeOfExit.ATR Then
                Dim avgPrice As Decimal = currentTrade.Supporting1
                Dim atr As Decimal = currentTrade.Supporting4
                If currentTick.Open >= avgPrice + atr * _userInputs.ExitValue Then
                    ret = New Tuple(Of Boolean, String)(True, "ATR Target Hit")
                End If
            End If
            If ret Is Nothing AndAlso currentTick.PayloadDate.Hour >= 15 AndAlso currentTick.PayloadDate.Minute >= 25 Then
                Dim currentTradingSymbol As String = Nothing
                If currentTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                    currentTradingSymbol = GetFutureInstrumentNameFromCore(Me.TradingSymbol, _nextTradingDay)
                Else
                    currentTradingSymbol = GetOptionInstrumentNameFromCore(Me.TradingSymbol, _nextTradingDay)
                End If
                If Not currentTrade.SupportingTradingSymbol.Contains(currentTradingSymbol) Then
                    Dim currentFOTick As Payload = GetCurrentTick(currentTrade.SupportingTradingSymbol, currentTick.PayloadDate)
                    _parentStrategy.ExitTradeByForce(currentTrade, currentFOTick, "Contract Rollover")

                    If Not currentTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                        Dim currentFutTick As Payload = GetCurrentTick(GetFutureInstrumentNameFromCore(Me.TradingSymbol, _nextTradingDay), currentTick.PayloadDate)

                        Dim atm As Decimal = Math.Ceiling(currentFutTick.Open / _strikeGap) * _strikeGap
                        currentTradingSymbol = String.Format("{0}{1}PE", GetOptionInstrumentNameFromCore(Me.TradingSymbol, _nextTradingDay), atm)
                    End If

                    currentFOTick = GetCurrentTick(currentTradingSymbol, currentTick.PayloadDate)
                    EnterDuplicateTrade(currentTrade, currentFOTick)
                End If
            End If
        End If
        If ret IsNot Nothing Then
            Dim allTrades As List(Of Trade) = _parentStrategy.GetAllTradesByTag(currentTrade.Tag, currentTick)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Close Then
                        Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTick.PayloadDate)
                        _parentStrategy.ExitTradeByForce(runningTrade, currentFOTick, ret.Item2)
                    End If
                Next
            End If
        End If
        Return Nothing
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

    Private Function GetCurrentTick(ByVal tradingSymbol As String, ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _dependentInstruments Is Nothing OrElse Not _dependentInstruments.ContainsKey(tradingSymbol) Then
            Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
            If tradingSymbol.EndsWith("FUT") Then
                inputPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, tradingSymbol, _tradingDate.AddDays(-5), _tradingDate)
            Else
                inputPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol, _tradingDate.AddDays(-5), _tradingDate)
            End If
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                If _dependentInstruments Is Nothing Then _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
                _dependentInstruments.Add(tradingSymbol, inputPayload)
            End If
        End If
        If _dependentInstruments IsNot Nothing AndAlso _dependentInstruments.ContainsKey(tradingSymbol) Then
            Dim inputPayload As Dictionary(Of Date, Payload) = _dependentInstruments(tradingSymbol)
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim currentMinute As Date = New Date(currentTime.Year, currentTime.Month, currentTime.Day, currentTime.Hour, currentTime.Minute, 0)
                Dim currentCandle As Payload = inputPayload.Where(Function(x)
                                                                      Return x.Key <= currentMinute
                                                                  End Function).LastOrDefault.Value

                If currentCandle Is Nothing Then
                    currentCandle = inputPayload.Where(Function(x)
                                                           Return x.Key.Date = _tradingDate
                                                       End Function).FirstOrDefault.Value
                End If

                ret = currentCandle.Ticks.FindAll(Function(x)
                                                      Return x.PayloadDate <= currentTime
                                                  End Function).LastOrDefault
                If ret Is Nothing Then ret = currentCandle.Ticks.FirstOrDefault
            End If
        End If
        Return ret
    End Function

    Private Function EnterTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload) As Boolean
        Dim ret As Boolean = False
        Dim tradeID As String = System.Guid.NewGuid.ToString()
        Dim tradeNumber As Integer = 1
        Dim currentFutureTradingSymbol As String = GetFutureInstrumentNameFromCore(Me.TradingSymbol, _tradingDate)
        If currentFutureTradingSymbol IsNot Nothing Then
            Dim currentFutTick As Payload = GetCurrentTick(currentFutureTradingSymbol, currentSpotTick.PayloadDate)
            If currentFutTick IsNot Nothing Then
                Dim totalQuantity As Integer = Me.LotSize
                Dim avgPrice As Decimal = currentFutTick.Open
                Dim entryATR As Decimal = _atrPayload(signalCandle.PayloadDate)
                Dim exitHelper As Decimal = Decimal.MinValue
                If _userInputs.ExitType = TypeOfExit.ATR Then
                    exitHelper = _atrPayload(signalCandle.PayloadDate)
                Else
                    exitHelper = currentFutTick.Open * Me.LotSize
                End If
                Dim entryTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(signalCandle, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
                    Dim totalValue As Decimal = Val(entryTrades.LastOrDefault.Supporting1) * Val(entryTrades.LastOrDefault.Supporting2)
                    Dim totalQty As Decimal = Val(entryTrades.LastOrDefault.Supporting2)

                    totalValue += currentFutTick.Open * Me.LotSize
                    totalQty += Me.LotSize
                    avgPrice = totalValue / totalQty
                    totalQuantity = totalQty

                    entryATR = entryTrades.LastOrDefault.Supporting4
                    tradeID = entryTrades.LastOrDefault.Tag
                    tradeNumber = Val(entryTrades.LastOrDefault.Supporting5) + 1

                    If _userInputs.ExitType = TypeOfExit.ATR Then
                        exitHelper = Math.Max(Val(entryTrades.LastOrDefault.Supporting3), _atrPayload(signalCandle.PayloadDate))
                    Else
                        exitHelper += Val(entryTrades.LastOrDefault.Supporting3)
                    End If

                    For Each runningTrade In entryTrades
                        runningTrade.UpdateTrade(Supporting1:=avgPrice, Supporting2:=totalQuantity, Supporting3:=exitHelper, Supporting5:=tradeNumber)
                    Next
                End If

                Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_parentStrategy,
                                                        tradingSymbol:=signalCandle.TradingSymbol,
                                                        stockType:=Trade.TypeOfStock.Futures,
                                                        orderType:=Trade.TypeOfOrder.Market,
                                                        tradingDate:=currentFutTick.PayloadDate,
                                                        entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                        entryPrice:=currentFutTick.Open,
                                                        entryBuffer:=0,
                                                        squareOffType:=Trade.TypeOfTrade.CNC,
                                                        entryCondition:=Trade.TradeEntryCondition.Original,
                                                        entryRemark:="Original Entry",
                                                        quantity:=Me.LotSize,
                                                        lotSize:=Me.LotSize,
                                                        potentialTarget:=currentFutTick.Open + 100000,
                                                        targetRemark:=100000,
                                                        potentialStopLoss:=currentFutTick.Open - 100000,
                                                        stoplossBuffer:=0,
                                                        slRemark:=100000,
                                                        signalCandle:=signalCandle)

                runningFutTrade.UpdateTrade(Tag:=tradeID,
                                             SquareOffValue:=100000,
                                             Supporting1:=avgPrice,
                                             Supporting2:=totalQuantity,
                                             Supporting3:=exitHelper,
                                             Supporting4:=entryATR,
                                             Supporting5:=tradeNumber,
                                             SupportingTradingSymbol:=currentFutureTradingSymbol)

                If _parentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                    Dim atm As Decimal = Math.Ceiling(currentFutTick.Open / _strikeGap) * _strikeGap
                    Dim currentOptionTradingSymbol As String = String.Format("{0}{1}PE", GetOptionInstrumentNameFromCore(Me.TradingSymbol, _tradingDate), atm)
                    Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
                    If currentOptTick IsNot Nothing Then
                        Dim runningOptTrade As Trade = New Trade(originatingStrategy:=_parentStrategy,
                                                                tradingSymbol:=signalCandle.TradingSymbol,
                                                                stockType:=Trade.TypeOfStock.Futures,
                                                                orderType:=Trade.TypeOfOrder.Market,
                                                                tradingDate:=currentOptTick.PayloadDate,
                                                                entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                                entryPrice:=currentOptTick.Open,
                                                                entryBuffer:=0,
                                                                squareOffType:=Trade.TypeOfTrade.CNC,
                                                                entryCondition:=Trade.TradeEntryCondition.Original,
                                                                entryRemark:="Original Entry",
                                                                quantity:=Me.LotSize,
                                                                lotSize:=Me.LotSize,
                                                                potentialTarget:=currentOptTick.Open + 100000,
                                                                targetRemark:=100000,
                                                                potentialStopLoss:=currentOptTick.Open - 100000,
                                                                stoplossBuffer:=0,
                                                                slRemark:=100000,
                                                                signalCandle:=signalCandle)

                        runningOptTrade.UpdateTrade(Tag:=tradeID,
                                                     SquareOffValue:=100000,
                                                     Supporting1:=avgPrice,
                                                     Supporting2:=totalQuantity,
                                                     Supporting3:=exitHelper,
                                                     Supporting4:=entryATR,
                                                     Supporting5:=tradeNumber,
                                                     SupportingTradingSymbol:=currentOptionTradingSymbol)

                        If _parentStrategy.PlaceOrModifyOrder(runningOptTrade, Nothing) Then
                            If _parentStrategy.EnterTradeIfPossible(runningFutTrade, currentFutTick) Then
                                ret = _parentStrategy.EnterTradeIfPossible(runningOptTrade, currentOptTick)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function EnterDuplicateTrade(ByVal existingTrade As Trade, ByVal currentTick As Payload) As Boolean
        Dim ret As Boolean = False
        Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_parentStrategy,
                                                  tradingSymbol:=existingTrade.SignalCandle.TradingSymbol,
                                                  stockType:=Trade.TypeOfStock.Futures,
                                                  orderType:=Trade.TypeOfOrder.Market,
                                                  tradingDate:=currentTick.PayloadDate,
                                                  entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                  entryPrice:=currentTick.Open,
                                                  entryBuffer:=0,
                                                  squareOffType:=Trade.TypeOfTrade.CNC,
                                                  entryCondition:=Trade.TradeEntryCondition.Original,
                                                  entryRemark:="Original Entry",
                                                  quantity:=Me.LotSize,
                                                  lotSize:=Me.LotSize,
                                                  potentialTarget:=currentTick.Open + 100000,
                                                  targetRemark:=100000,
                                                  potentialStopLoss:=currentTick.Open - 100000,
                                                  stoplossBuffer:=0,
                                                  slRemark:=100000,
                                                  signalCandle:=existingTrade.SignalCandle)

        runningFutTrade.UpdateTrade(Tag:=existingTrade.Tag,
                                    SquareOffValue:=100000,
                                    Supporting1:=existingTrade.Supporting1,
                                    Supporting2:=existingTrade.Supporting2,
                                    Supporting3:=existingTrade.Supporting3,
                                    Supporting4:=existingTrade.Supporting4,
                                    Supporting5:=existingTrade.Supporting5,
                                    SupportingTradingSymbol:=currentTick.TradingSymbol)

        If _parentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
            ret = _parentStrategy.EnterTradeIfPossible(runningFutTrade, currentTick)
        End If
        Return ret
    End Function

#Region "Contract Helper"
    Private Function GetFutureInstrumentNameFromCore(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
        Dim ret As String = Nothing
        If coreInstrumentName = "NIFTY 50" Then
            coreInstrumentName = "NIFTY"
        ElseIf coreInstrumentName = "NIFTY BANK" Then
            coreInstrumentName = "BANKNIFTY"
        End If
        Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
        If tradingDate.Date > lastThursday.Date.AddDays(-2) Then
            ret = String.Format("{0}{1}FUT", coreInstrumentName, tradingDate.AddDays(10).ToString("yyMMM")).ToUpper
        Else
            ret = String.Format("{0}{1}FUT", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
        End If
        Return ret
    End Function

    Private Function GetOptionInstrumentNameFromCore(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
        Dim ret As String = Nothing
        If coreInstrumentName = "NIFTY 50" Then
            coreInstrumentName = "NIFTY"
        ElseIf coreInstrumentName = "NIFTY BANK" Then
            coreInstrumentName = "BANKNIFTY"
        End If
        Dim nextThursday As Date = GetNextThusrday(tradingDate)
        If nextThursday.Date = New Date(2020, 4, 2).Date Then nextThursday = New Date(2020, 4, 1)
        Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
        If tradingDate.Date > nextThursday.Date.AddDays(-2) Then
            nextThursday = GetNextThusrday(tradingDate.AddDays(3))
            If nextThursday.Date = New Date(2020, 4, 2).Date Then nextThursday = New Date(2020, 4, 1)
            If nextThursday.Date = lastThursday.Date Then
                ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
            Else
                ret = String.Format("{0}{1}", coreInstrumentName, GetExpiryString(nextThursday)).ToUpper
            End If
        Else
            If nextThursday.Date = lastThursday.Date Then
                ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
            Else
                ret = String.Format("{0}{1}", coreInstrumentName, GetExpiryString(nextThursday)).ToUpper
            End If
        End If
        Return ret
    End Function

    Private Function GetExpiryString(ByVal expiryDate As Date) As String
        If expiryDate.Month >= 10 Then
            Return String.Format("{0}{1}{2}", expiryDate.ToString("yy"), expiryDate.ToString("MMM").Substring(0, 1), expiryDate.ToString("dd"))
        Else
            Return String.Format("{0}{1}{2}", expiryDate.ToString("yy"), expiryDate.Month, expiryDate.ToString("dd"))
        End If
    End Function

    Private Function GetLastThusrdayOfMonth(ByVal dateTime As Date) As Date
        Dim ret As Date = Date.MinValue
        Dim lastDayOfMonth As Date = New Date(dateTime.Year, dateTime.Month, Date.DaysInMonth(dateTime.Year, dateTime.Month))
        While True
            If lastDayOfMonth.DayOfWeek = DayOfWeek.Thursday Then
                ret = lastDayOfMonth
                Exit While
            End If
            lastDayOfMonth = lastDayOfMonth.AddDays(-1)
        End While
        Return ret
    End Function

    Private Function GetNextThusrday(ByVal dateTime As Date) As Date
        Dim ret As Date = Date.MinValue
        Dim daysUntilTradingDay As Integer = (CInt(DayOfWeek.Thursday) - CInt(dateTime.DayOfWeek) + 7) Mod 7
        ret = dateTime.Date.AddDays(daysUntilTradingDay).Date
        Return ret
    End Function
#End Region

End Class
