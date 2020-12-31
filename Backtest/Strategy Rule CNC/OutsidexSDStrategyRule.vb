Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper

Public Class OutsidexSDStrategyRule
    Inherits PairStrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public EntrySD As Integer
        Public TargetPercentage As Decimal
    End Class
#End Region

    Private _ratioPayload As Dictionary(Of Date, Decimal)
    Private _plusSDPayload As Dictionary(Of Date, Decimal)
    Private _minusSDPayload As Dictionary(Of Date, Decimal)
    Private _smaPayload As Dictionary(Of Date, Decimal)

    Private _dependentInstruments As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal pairName As String,
                   ByVal tradingDate As Date,
                   ByVal nextTradingDay As Date,
                   ByVal tradingSymbol1 As String,
                   ByVal lotSize1 As Integer,
                   ByVal tradingSymbol2 As String,
                   ByVal lotSize2 As Integer,
                   ByVal entities As RuleEntities,
                   ByVal parentStrategy As Strategy,
                   ByVal canceller As CancellationTokenSource,
                   ByVal inputPayload1 As Dictionary(Of Date, Payload),
                   ByVal inputPayload2 As Dictionary(Of Date, Payload))
        MyBase.New(pairName, tradingDate, nextTradingDay, tradingSymbol1, lotSize1, tradingSymbol2, lotSize2, entities, parentStrategy, canceller, inputPayload1, inputPayload2)
        _userInputs = _Entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim basePayload As Dictionary(Of Date, Payload) = Nothing
        If _SignalPayload2.Count > _SignalPayload1.Count Then
            basePayload = _SignalPayload2
        Else
            basePayload = _SignalPayload1
        End If
        For Each runningPayload In basePayload.Keys
            Dim payload1 As Payload = _SignalPayload1.Where(Function(x)
                                                                Return x.Key <= runningPayload
                                                            End Function).LastOrDefault.Value
            If payload1 Is Nothing Then
                payload1 = _SignalPayload1.Where(Function(x)
                                                     Return x.Key.Date = _TradingDate
                                                 End Function).FirstOrDefault.Value
            End If

            Dim payload2 As Payload = _SignalPayload2.Where(Function(x)
                                                                Return x.Key <= runningPayload
                                                            End Function).LastOrDefault.Value
            If payload2 Is Nothing Then
                payload2 = _SignalPayload2.Where(Function(x)
                                                     Return x.Key.Date = _TradingDate
                                                 End Function).FirstOrDefault.Value
            End If

            If _ratioPayload Is Nothing Then _ratioPayload = New Dictionary(Of Date, Decimal)
            _ratioPayload.Add(runningPayload, payload1.Close / payload2.Close)
        Next

        Dim convertedPayload As Dictionary(Of Date, Payload) = Common.ConvertDecimalToPayload(Payload.PayloadFields.Close, _ratioPayload)
        Indicator.BollingerBands.CalculateBollingerBands(200, Payload.PayloadFields.Close, _userInputs.EntrySD, convertedPayload, _plusSDPayload, _minusSDPayload, _smaPayload)

        Dim allRunningTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_PairName, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If allRunningTrades IsNot Nothing AndAlso allRunningTrades.Count > 0 Then
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            For Each runningTrade In allRunningTrades
                If Not _dependentInstruments.ContainsKey(runningTrade.SupportingTradingSymbol) Then
                    Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                    If runningTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                        inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-5), _TradingDate)
                    Else
                        inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-5), _TradingDate)
                    End If
                    If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                        _dependentInstruments.Add(runningTrade.SupportingTradingSymbol, inputPayload)
                    End If
                End If
            Next
            _dependentInstruments.Add(_TradingSymbol1, _InputPayload1)
            _dependentInstruments.Add(_TradingSymbol2, _InputPayload2)
        Else
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            _dependentInstruments.Add(_TradingSymbol1, _InputPayload1)
            _dependentInstruments.Add(_TradingSymbol2, _InputPayload2)
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime)
        Dim tradeStartTime As Date = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day, _ParentStrategy.TradeStartTime.Hours, _ParentStrategy.TradeStartTime.Minutes, _ParentStrategy.TradeStartTime.Seconds)
        If currentTickTime >= tradeStartTime Then
            Dim lastEntryTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_PairName, _TradingDate, Trade.TypeOfTrade.CNC)
            If lastEntryTrade Is Nothing OrElse (lastEntryTrade.ExitRemark IsNot Nothing AndAlso lastEntryTrade.ExitRemark.ToUpper = "TARGET HIT") Then
                'Fresh Entry
                If _SignalPayload1.ContainsKey(currentMinute) AndAlso _SignalPayload2.ContainsKey(currentMinute) Then
                    Dim currentMinuteCandle1 As Payload = _SignalPayload1(currentMinute)
                    Dim currentMinuteCandle2 As Payload = _SignalPayload2(currentMinute)
                    If currentMinuteCandle1 IsNot Nothing AndAlso currentMinuteCandle1.PreviousCandlePayload IsNot Nothing AndAlso
                        currentMinuteCandle2 IsNot Nothing AndAlso currentMinuteCandle2.PreviousCandlePayload IsNot Nothing Then
                        Dim ratio As Decimal = _ratioPayload(currentMinuteCandle1.PreviousCandlePayload.PayloadDate)
                        Dim plusSD As Decimal = _plusSDPayload(currentMinuteCandle1.PreviousCandlePayload.PayloadDate)
                        Dim minusSD As Decimal = _minusSDPayload(currentMinuteCandle1.PreviousCandlePayload.PayloadDate)
                        Dim mean As Decimal = _smaPayload(currentMinuteCandle1.PreviousCandlePayload.PayloadDate)
                        Dim sd As Decimal = (plusSD - mean) / _userInputs.EntrySD
                        Dim zScore As Decimal = (ratio - mean) / sd
                        If zScore >= _userInputs.EntrySD OrElse
                            zScore <= _userInputs.EntrySD * -1 Then
                            Dim turnover1 As Decimal = currentMinuteCandle1.PreviousCandlePayload.Close * _LotSize1
                            Dim turnover2 As Decimal = currentMinuteCandle2.PreviousCandlePayload.Close * _LotSize2
                            Dim turnoverPer As Decimal = (turnover1 / (turnover1 + turnover2)) * 100
                            If turnoverPer >= 45 AndAlso turnoverPer <= 55 Then
                                _ParentStrategy.MaxZSCore = zScore
                                If Not EnterTrade(currentMinuteCandle1.PreviousCandlePayload, currentTickTime, zScore, mean, sd) Then
                                    Throw New NotImplementedException()
                                End If
                            Else
                                _ParentStrategy.NeglectedTradeCount += 1
                                Console.WriteLine(String.Format("Trade Ignored. Date:{0}, ZScore:{1}, Turnover%:{2}",
                                                                currentMinuteCandle1.PreviousCandlePayload.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                                                                zScore, Math.Round(turnoverPer, 2)))
                            End If
                        End If
                    End If
                End If
            Else
                'Averaging Logic
                If _SignalPayload1.ContainsKey(currentMinute) AndAlso _SignalPayload2.ContainsKey(currentMinute) Then
                    Dim currentMinuteCandle As Payload = _SignalPayload1(currentMinute)
                    If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing Then
                        Dim ratio As Decimal = _ratioPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                        Dim mean As Decimal = Val(lastEntryTrade.Supporting3)
                        Dim sd As Decimal = Val(lastEntryTrade.Supporting4)
                        Dim zScore As Decimal = (ratio - mean) / sd

                        If _ParentStrategy.MaxZSCore > 0 Then
                            _ParentStrategy.MaxZSCore = Math.Max(_ParentStrategy.MaxZSCore, zScore)
                        Else
                            _ParentStrategy.MaxZSCore = Math.Min(_ParentStrategy.MaxZSCore, zScore)
                        End If

                        If _ParentStrategy.MaxZSCore > 0 Then
                            If zScore <= _ParentStrategy.MaxZSCore - 1 Then
                                If zScore >= Val(lastEntryTrade.Supporting2) + 1 Then
                                    If Not EnterTrade(currentMinuteCandle.PreviousCandlePayload, currentTickTime, zScore, mean, sd) Then
                                        Throw New NotImplementedException()
                                    End If
                                End If
                            End If
                        Else
                            If zScore >= _ParentStrategy.MaxZSCore + 1 Then
                                If zScore <= Val(lastEntryTrade.Supporting2) - 1 Then
                                    If Not EnterTrade(currentMinuteCandle.PreviousCandlePayload, currentTickTime, zScore, mean, sd) Then
                                        Throw New NotImplementedException()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTickTime As Date, ByVal availableTrades As List(Of Trade)) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        Dim exitAtTarget As Boolean = False
        If availableTrades IsNot Nothing AndAlso availableTrades.Count > 0 Then
            For Each runningAvailableTrade In availableTrades
                If runningAvailableTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByTag(runningAvailableTrade.Tag, _PairName)
                    If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                        Dim totalPL As Decimal = 0
                        For Each runningTrade In allTrades
                            If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
                                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                    totalPL += _ParentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                                Else
                                    totalPL += _ParentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, runningTrade.ExitPrice, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                                End If
                            Else
                                Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                                If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                    totalPL += _ParentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, runningTrade.EntryPrice, currentFOTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                                Else
                                    totalPL += _ParentStrategy.CalculatePL(runningTrade.SupportingTradingSymbol, currentFOTick.Open, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                                End If
                            End If
                        Next

                        Dim entryDate As Date = Date.ParseExact(allTrades.LastOrDefault.Supporting6, "dd-MMM-yyyy HH:mm:ss", Nothing).Date
                        Dim plToAchive As Decimal = Math.Max(_ParentStrategy.GetMaxCapital(entryDate, currentTickTime), 60000) * _userInputs.TargetPercentage / 100

                        If totalPL >= plToAchive Then
                            exitAtTarget = True
                            Exit For
                        End If
                    End If
                End If
            Next

            If Not exitAtTarget AndAlso currentTickTime.Hour >= 15 AndAlso currentTickTime.Minute >= 25 Then
                For Each runningAvailableTrade In availableTrades
                    If runningAvailableTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim coreInstrumentName As String = Nothing
                        If runningAvailableTrade.SupportingTradingSymbol.StartsWith(_TradingSymbol2) Then
                            coreInstrumentName = _TradingSymbol2
                        Else
                            coreInstrumentName = _TradingSymbol1
                        End If
                        Dim currentSpotTick As Payload = GetCurrentTick(coreInstrumentName, currentTickTime)
                        Dim currentTradingSymbol As String = Nothing
                        If runningAvailableTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                            currentTradingSymbol = GetFutureInstrumentNameFromCore(coreInstrumentName, _NextTradingDay)
                        Else
                            currentTradingSymbol = GetOptionInstrumentNameFromCore(coreInstrumentName, _NextTradingDay)
                        End If
                        If Not runningAvailableTrade.SupportingTradingSymbol.Contains(currentTradingSymbol) Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningAvailableTrade.SupportingTradingSymbol, currentTickTime)
                            _ParentStrategy.ExitTradeByForce(runningAvailableTrade, currentFOTick, "Contract Rollover")

                            If Not runningAvailableTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                                Dim optionType As String = runningAvailableTrade.SupportingTradingSymbol.Substring(runningAvailableTrade.SupportingTradingSymbol.Count - 2)
                                currentTradingSymbol = GetCurrentATMOption(currentTickTime, GetOptionInstrumentNameFromCore(coreInstrumentName, _NextTradingDay), currentSpotTick.Open, optionType)
                            End If

                            currentFOTick = GetCurrentTick(currentTradingSymbol, currentTickTime)
                            EnterDuplicateTrade(runningAvailableTrade, currentFOTick)
                        End If
                    End If
                Next
            End If
        End If
        If exitAtTarget Then
            For Each runningAvailableTrade In availableTrades
                If runningAvailableTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByTag(runningAvailableTrade.Tag, _PairName)
                    If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                        For Each runningTrade In allTrades
                            If runningTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Close Then
                                Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                                _ParentStrategy.ExitTradeByForce(runningTrade, currentFOTick, "Target Hit")
                            End If
                        Next

                        Dim entryDate As Date = Date.ParseExact(allTrades.LastOrDefault.Supporting6, "dd-MMM-yyyy HH:mm:ss", Nothing).Date
                        Dim exitDate As Date = currentTickTime.ToString("dd-MMM-yyyy HH:mm:ss")
                        Dim numberOfDays As Integer = _TradingDate.Subtract(entryDate).Days
                        For Each runningTrade In allTrades
                            runningTrade.UpdateTrade(Supporting7:=exitDate, Supporting8:=numberOfDays)
                        Next
                    End If
                End If
            Next
        End If
    End Function

    Private Function GetCurrentTick(ByVal tradingSymbol As String, ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _dependentInstruments Is Nothing OrElse Not _dependentInstruments.ContainsKey(tradingSymbol) Then
            Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
            If tradingSymbol.EndsWith("FUT") Then
                inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, tradingSymbol, _TradingDate.AddDays(-5), _TradingDate)
            Else
                inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol, _TradingDate.AddDays(-5), _TradingDate)
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
                                                           Return x.Key.Date = _TradingDate
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

    Private Function EnterTrade(ByVal signalCandle As Payload, ByVal currentTickTime As Date, ByVal entryZScore As Decimal, ByVal mean As Decimal, ByVal sd As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim tradeID As String = System.Guid.NewGuid.ToString()
        Dim tradeNumber As Integer = 1
        Dim entryDate As String = currentTickTime.ToString("dd-MMM-yyyy HH:mm:ss")

        Dim entryTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_PairName, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
            tradeID = entryTrades.LastOrDefault.Tag
            tradeNumber = Val(entryTrades.LastOrDefault.Supporting1) + 1
            mean = Val(entryTrades.LastOrDefault.Supporting3)
            sd = Val(entryTrades.LastOrDefault.Supporting4)
            entryDate = entryTrades.LastOrDefault.Supporting6
        End If

        If entryZScore < 0 Then
            Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol1, currentTickTime)
            If EnterBuyTrade(signalCandle, currentSpotTick, _TradingSymbol1, _LotSize1, tradeID, tradeNumber, entryZScore, entryDate, mean, sd) Then
                currentSpotTick = GetCurrentTick(_TradingSymbol2, currentTickTime)
                If EnterSellTrade(signalCandle, currentSpotTick, _TradingSymbol2, _LotSize2, tradeID, tradeNumber, entryZScore, entryDate, mean, sd) Then
                    ret = True
                End If
            End If
        Else
            Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol1, currentTickTime)
            If EnterSellTrade(signalCandle, currentSpotTick, _TradingSymbol1, _LotSize1, tradeID, tradeNumber, entryZScore, entryDate, mean, sd) Then
                currentSpotTick = GetCurrentTick(_TradingSymbol2, currentTickTime)
                If EnterBuyTrade(signalCandle, currentSpotTick, _TradingSymbol2, _LotSize2, tradeID, tradeNumber, entryZScore, entryDate, mean, sd) Then
                    ret = True
                End If
            End If
        End If

        Return ret
    End Function

    Private Function EnterBuyTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload,
                                   ByVal tradingSymbol As String, ByVal lotSize As Integer,
                                   ByVal tradeID As String, ByVal tradeNumber As Integer,
                                   ByVal entryZScore As Decimal, ByVal entryDate As String,
                                   ByVal mean As Decimal, ByVal sd As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim currentFutureTradingSymbol As String = GetFutureInstrumentNameFromCore(tradingSymbol, _TradingDate)
        If currentFutureTradingSymbol IsNot Nothing Then
            Dim currentFutTick As Payload = GetCurrentTick(currentFutureTradingSymbol, currentSpotTick.PayloadDate)
            If currentFutTick IsNot Nothing Then
                Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                        tradingSymbol:=_PairName,
                                                        stockType:=Trade.TypeOfStock.Futures,
                                                        orderType:=Trade.TypeOfOrder.Market,
                                                        tradingDate:=currentFutTick.PayloadDate,
                                                        entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                        entryPrice:=currentFutTick.Open,
                                                        entryBuffer:=0,
                                                        squareOffType:=Trade.TypeOfTrade.CNC,
                                                        entryCondition:=Trade.TradeEntryCondition.Original,
                                                        entryRemark:="Original Entry",
                                                        quantity:=lotSize,
                                                        lotSize:=lotSize,
                                                        potentialTarget:=currentFutTick.Open + 100000,
                                                        targetRemark:=100000,
                                                        potentialStopLoss:=currentFutTick.Open - 100000,
                                                        stoplossBuffer:=0,
                                                        slRemark:=100000,
                                                        signalCandle:=signalCandle)

                runningFutTrade.UpdateTrade(Tag:=tradeID,
                                             SquareOffValue:=100000,
                                             Supporting1:=tradeNumber,
                                             Supporting2:=entryZScore,
                                             Supporting3:=mean,
                                             Supporting4:=sd,
                                             Supporting6:=entryDate,
                                             SupportingTradingSymbol:=currentFutureTradingSymbol)

                If _ParentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                    Dim currentOptionTradingSymbol As String = GetCurrentATMOption(currentSpotTick.PayloadDate, GetOptionInstrumentNameFromCore(tradingSymbol, _TradingDate), currentSpotTick.Open, "PE")
                    Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
                    If currentOptTick IsNot Nothing Then
                        Dim runningOptTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                                tradingSymbol:=_PairName,
                                                                stockType:=Trade.TypeOfStock.Options,
                                                                orderType:=Trade.TypeOfOrder.Market,
                                                                tradingDate:=currentOptTick.PayloadDate,
                                                                entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                                entryPrice:=currentOptTick.Open,
                                                                entryBuffer:=0,
                                                                squareOffType:=Trade.TypeOfTrade.CNC,
                                                                entryCondition:=Trade.TradeEntryCondition.Original,
                                                                entryRemark:="Original Entry",
                                                                quantity:=lotSize,
                                                                lotSize:=lotSize,
                                                                potentialTarget:=currentOptTick.Open + 100000,
                                                                targetRemark:=100000,
                                                                potentialStopLoss:=currentOptTick.Open - 100000,
                                                                stoplossBuffer:=0,
                                                                slRemark:=100000,
                                                                signalCandle:=signalCandle)

                        runningOptTrade.UpdateTrade(Tag:=tradeID,
                                                     SquareOffValue:=100000,
                                                     Supporting1:=tradeNumber,
                                                     Supporting2:=entryZScore,
                                                     Supporting3:=mean,
                                                     Supporting4:=sd,
                                                     Supporting6:=entryDate,
                                                     SupportingTradingSymbol:=currentOptionTradingSymbol)

                        If _ParentStrategy.PlaceOrModifyOrder(runningOptTrade, Nothing) Then
                            If _ParentStrategy.EnterTradeIfPossible(_PairName, _TradingDate, runningFutTrade, currentFutTick) Then
                                If _ParentStrategy.EnterTradeIfPossible(_PairName, _TradingDate, runningOptTrade, currentOptTick) Then
                                    ret = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function EnterSellTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload,
                                    ByVal tradingSymbol As String, ByVal lotSize As Integer,
                                    ByVal tradeID As String, ByVal tradeNumber As Integer,
                                    ByVal entryZScore As Decimal, ByVal entryDate As String,
                                    ByVal mean As Decimal, ByVal sd As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim currentFutureTradingSymbol As String = GetFutureInstrumentNameFromCore(tradingSymbol, _TradingDate)
        If currentFutureTradingSymbol IsNot Nothing Then
            Dim currentFutTick As Payload = GetCurrentTick(currentFutureTradingSymbol, currentSpotTick.PayloadDate)
            If currentFutTick IsNot Nothing Then
                Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                        tradingSymbol:=_PairName,
                                                        stockType:=Trade.TypeOfStock.Futures,
                                                        orderType:=Trade.TypeOfOrder.Market,
                                                        tradingDate:=currentFutTick.PayloadDate,
                                                        entryDirection:=Trade.TradeExecutionDirection.Sell,
                                                        entryPrice:=currentFutTick.Open,
                                                        entryBuffer:=0,
                                                        squareOffType:=Trade.TypeOfTrade.CNC,
                                                        entryCondition:=Trade.TradeEntryCondition.Original,
                                                        entryRemark:="Original Entry",
                                                        quantity:=lotSize,
                                                        lotSize:=lotSize,
                                                        potentialTarget:=currentFutTick.Open + 100000,
                                                        targetRemark:=100000,
                                                        potentialStopLoss:=currentFutTick.Open - 100000,
                                                        stoplossBuffer:=0,
                                                        slRemark:=100000,
                                                        signalCandle:=signalCandle)

                runningFutTrade.UpdateTrade(Tag:=tradeID,
                                             SquareOffValue:=100000,
                                             Supporting1:=tradeNumber,
                                             Supporting2:=entryZScore,
                                             Supporting3:=mean,
                                             Supporting4:=sd,
                                             Supporting6:=entryDate,
                                             SupportingTradingSymbol:=currentFutureTradingSymbol)

                If _ParentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                    Dim currentOptionTradingSymbol As String = GetCurrentATMOption(currentSpotTick.PayloadDate, GetOptionInstrumentNameFromCore(tradingSymbol, _TradingDate), currentSpotTick.Open, "CE")
                    Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
                    If currentOptTick IsNot Nothing Then
                        Dim runningOptTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                                tradingSymbol:=_PairName,
                                                                stockType:=Trade.TypeOfStock.Options,
                                                                orderType:=Trade.TypeOfOrder.Market,
                                                                tradingDate:=currentOptTick.PayloadDate,
                                                                entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                                entryPrice:=currentOptTick.Open,
                                                                entryBuffer:=0,
                                                                squareOffType:=Trade.TypeOfTrade.CNC,
                                                                entryCondition:=Trade.TradeEntryCondition.Original,
                                                                entryRemark:="Original Entry",
                                                                quantity:=lotSize,
                                                                lotSize:=lotSize,
                                                                potentialTarget:=currentOptTick.Open + 100000,
                                                                targetRemark:=100000,
                                                                potentialStopLoss:=currentOptTick.Open - 100000,
                                                                stoplossBuffer:=0,
                                                                slRemark:=100000,
                                                                signalCandle:=signalCandle)

                        runningOptTrade.UpdateTrade(Tag:=tradeID,
                                                     SquareOffValue:=100000,
                                                     Supporting1:=tradeNumber,
                                                     Supporting2:=entryZScore,
                                                     Supporting3:=mean,
                                                     Supporting4:=sd,
                                                     Supporting6:=entryDate,
                                                     SupportingTradingSymbol:=currentOptionTradingSymbol)

                        If _ParentStrategy.PlaceOrModifyOrder(runningOptTrade, Nothing) Then
                            If _ParentStrategy.EnterTradeIfPossible(_PairName, _TradingDate, runningFutTrade, currentFutTick) Then
                                If _ParentStrategy.EnterTradeIfPossible(_PairName, _TradingDate, runningOptTrade, currentOptTick) Then
                                    ret = True
                                End If
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
        Try
            Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                      tradingSymbol:=existingTrade.TradingSymbol,
                                                      stockType:=existingTrade.StockType,
                                                      orderType:=Trade.TypeOfOrder.Market,
                                                      tradingDate:=currentTick.PayloadDate,
                                                      entryDirection:=existingTrade.EntryDirection,
                                                      entryPrice:=currentTick.Open,
                                                      entryBuffer:=0,
                                                      squareOffType:=Trade.TypeOfTrade.CNC,
                                                      entryCondition:=Trade.TradeEntryCondition.Original,
                                                      entryRemark:="Original Entry",
                                                      quantity:=existingTrade.LotSize,
                                                      lotSize:=existingTrade.LotSize,
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
                                        Supporting6:=existingTrade.Supporting6,
                                        SupportingTradingSymbol:=currentTick.TradingSymbol)

            If _ParentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                ret = _ParentStrategy.EnterTradeIfPossible(runningFutTrade.TradingSymbol, _TradingDate, runningFutTrade, currentTick)
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function

#Region "Contract Helper"
    Private Function GetCurrentATMOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal, ByVal optionType As String) As String
        If optionType = "PE" Then
            Return GetCurrentATMPEOption(currentTickTime, expiryString, price)
        Else
            Return GetCurrentATMCEOption(currentTickTime, expiryString, price)
        End If
    End Function

    Private Function GetCurrentATMPEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
        Dim ret As String = Nothing
        Dim query As String = Nothing
        'If currentTickTime.Hour >= 12 AndAlso currentTickTime.Minute >= 15 Then
        query = "SELECT `TradingSymbol`,SUM(`Volume`) Vol
                        FROM `intraday_prices_opt_futures`
                        WHERE `SnapshotDate`='{0}'
                        AND `SnapshotTime`<='{1}'
                        AND `TradingSymbol` LIKE '{2}%{3}'
                        GROUP BY `TradingSymbol`
                        ORDER BY Vol DESC"
        query = String.Format(query, currentTickTime.ToString("yyyy-MM-dd"), currentTickTime.AddMinutes(-1).ToString("HH:mm:ss"), expiryString, "PE")
        'Else
        '    query = "SELECT `TradingSymbol`,`Volume` Vol
        '                FROM `eod_prices_opt_futures`
        '                WHERE `SnapshotDate`='{0}'
        '                AND `TradingSymbol` LIKE '{1}%{2}'
        '                ORDER BY Vol DESC"

        '    Dim previousTradingDay As Date = _ParentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Futures, _TradingDate)
        '    query = String.Format(query, previousTradingDay.ToString("yyyy-MM-dd"), expiryString, "PE")
        'End If

        Dim dt As DataTable = _ParentStrategy.Cmn.RunSelect(query)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim contracts As Dictionary(Of Decimal, Long) = New Dictionary(Of Decimal, Long)
            For Each runningRow As DataRow In dt.Rows
                Dim tradingSymbol As String = runningRow.Item("TradingSymbol")
                Dim volume As Long = runningRow.Item("Vol")
                Dim strike As String = Utilities.Strings.GetTextBetween(expiryString, "PE", tradingSymbol)
                If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike.Trim) Then
                    contracts.Add(Val(strike.Trim), volume)
                End If
            Next
            If contracts IsNot Nothing AndAlso contracts.Count > 0 Then
                Dim sumVol As Double = contracts.Sum(Function(x)
                                                         If x.Key >= price Then
                                                             Return x.Value
                                                         Else
                                                             Return 0
                                                         End If
                                                     End Function)
                If sumVol > 0 Then
                    Dim countVol As Double = contracts.Sum(Function(x)
                                                               If x.Key >= price AndAlso x.Value > 0 Then
                                                                   Return 1
                                                               Else
                                                                   Return 0
                                                               End If
                                                           End Function)
                    Dim avgVol As Double = sumVol / countVol
                    For Each runningContract In contracts.OrderBy(Function(x)
                                                                      Return x.Key
                                                                  End Function)
                        If runningContract.Key >= price Then
                            If runningContract.Value >= avgVol Then
                                ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                                Exit For
                            End If
                        End If
                    Next
                Else
                    sumVol = contracts.Sum(Function(x)
                                               If x.Key <= price Then
                                                   Return x.Value
                                               Else
                                                   Return 0
                                               End If
                                           End Function)

                    Dim countVol As Double = contracts.Sum(Function(x)
                                                               If x.Key <= price AndAlso x.Value > 0 Then
                                                                   Return 1
                                                               Else
                                                                   Return 0
                                                               End If
                                                           End Function)
                    Dim avgVol As Double = sumVol / countVol
                    For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                        If runningContract.Key <= price Then
                            If runningContract.Value >= avgVol Then
                                ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetCurrentATMCEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
        Dim ret As String = Nothing
        Dim query As String = Nothing
        'If currentTickTime.Hour >= 12 AndAlso currentTickTime.Minute >= 15 Then
        query = "SELECT `TradingSymbol`,SUM(`Volume`) Vol
                        FROM `intraday_prices_opt_futures`
                        WHERE `SnapshotDate`='{0}'
                        AND `SnapshotTime`<='{1}'
                        AND `TradingSymbol` LIKE '{2}%{3}'
                        GROUP BY `TradingSymbol`
                        ORDER BY Vol DESC"
        query = String.Format(query, currentTickTime.ToString("yyyy-MM-dd"), currentTickTime.AddMinutes(-1).ToString("HH:mm:ss"), expiryString, "CE")
        'Else
        '    query = "SELECT `TradingSymbol`,`Volume` Vol
        '                FROM `eod_prices_opt_futures`
        '                WHERE `SnapshotDate`='{0}'
        '                AND `TradingSymbol` LIKE '{1}%{2}'
        '                ORDER BY Vol DESC"

        '    Dim previousTradingDay As Date = _ParentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Futures, _TradingDate)
        '    query = String.Format(query, previousTradingDay.ToString("yyyy-MM-dd"), expiryString, "CE")
        'End If

        Dim dt As DataTable = _ParentStrategy.Cmn.RunSelect(query)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim contracts As Dictionary(Of Decimal, Long) = New Dictionary(Of Decimal, Long)
            For Each runningRow As DataRow In dt.Rows
                Dim tradingSymbol As String = runningRow.Item("TradingSymbol")
                Dim volume As Long = runningRow.Item("Vol")
                Dim strike As String = Utilities.Strings.GetTextBetween(expiryString, "CE", tradingSymbol)
                If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike.Trim) Then
                    contracts.Add(Val(strike.Trim), volume)
                End If
            Next
            If contracts IsNot Nothing AndAlso contracts.Count > 0 Then
                Dim sumVol As Double = contracts.Sum(Function(x)
                                                         If x.Key <= price Then
                                                             Return x.Value
                                                         Else
                                                             Return 0
                                                         End If
                                                     End Function)
                If sumVol > 0 Then
                    Dim countVol As Double = contracts.Sum(Function(x)
                                                               If x.Key <= price AndAlso x.Value > 0 Then
                                                                   Return 1
                                                               Else
                                                                   Return 0
                                                               End If
                                                           End Function)
                    Dim avgVol As Double = sumVol / countVol
                    For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                        If runningContract.Key <= price Then
                            If runningContract.Value >= avgVol Then
                                ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                                Exit For
                            End If
                        End If
                    Next
                Else
                    sumVol = contracts.Sum(Function(x)
                                               If x.Key >= price Then
                                                   Return x.Value
                                               Else
                                                   Return 0
                                               End If
                                           End Function)

                    Dim countVol As Double = contracts.Sum(Function(x)
                                                               If x.Key >= price AndAlso x.Value > 0 Then
                                                                   Return 1
                                                               Else
                                                                   Return 0
                                                               End If
                                                           End Function)
                    Dim avgVol As Double = sumVol / countVol
                    For Each runningContract In contracts.OrderBy(Function(x)
                                                                      Return x.Key
                                                                  End Function)
                        If runningContract.Key >= price Then
                            If runningContract.Value >= avgVol Then
                                ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetFutureInstrumentNameFromCore(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
        Dim ret As String = Nothing
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
        Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
        If tradingDate.Date > lastThursday.Date.AddDays(-2) Then
            ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.AddDays(10).ToString("yyMMM")).ToUpper
        Else
            ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
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
