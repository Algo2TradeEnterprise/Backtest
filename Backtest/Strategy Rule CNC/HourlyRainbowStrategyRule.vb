Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class HourlyRainbowStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public RainbowPeriod As Integer
        Public ATRMultiplier As Decimal
        Public SpotToOptionDelta As Decimal
        Public ExitAtATRPL As Boolean
        Public HalfPremiumExit As Boolean
        Public OptionStrikeDistance As Integer
    End Class

    Private Enum ExitType
        Target = 1
        Reverse
        HalfPremium
    End Enum
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _rainbowPayload As Dictionary(Of Date, Indicator.RainbowMA)

    Private _dependentInstruments As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal tradingDate As Date,
                   ByVal nextTradingDay As Date,
                   ByVal tradingSymbol As String,
                   ByVal lotSize As Integer,
                   ByVal entities As RuleEntities,
                   ByVal parentStrategy As Strategy,
                   ByVal canceller As CancellationTokenSource,
                   ByVal inputPayload As Dictionary(Of Date, Payload))
        MyBase.New(tradingDate, nextTradingDay, tradingSymbol, lotSize, entities, parentStrategy, canceller, inputPayload)
        _userInputs = _Entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _SignalPayload, _atrPayload)
        Indicator.RainbowMovingAverage.CalculateRainbowMovingAverage(_userInputs.RainbowPeriod, _SignalPayload, _rainbowPayload)

        Dim allRunningTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If allRunningTrades IsNot Nothing AndAlso allRunningTrades.Count > 0 Then
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            For Each runningTrade In allRunningTrades
                If Not _dependentInstruments.ContainsKey(runningTrade.SupportingTradingSymbol) Then
                    Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                    If runningTrade.SupportingTradingSymbol.EndsWith("FUT") Then
                        inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-30), _TradingDate)
                    Else
                        inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-30), _TradingDate)
                    End If
                    If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                        _dependentInstruments.Add(runningTrade.SupportingTradingSymbol, inputPayload)
                    End If
                End If
            Next
            _dependentInstruments.Add(_TradingSymbol, _InputPayload)
        Else
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            _dependentInstruments.Add(_TradingSymbol, _InputPayload)
        End If
    End Sub

    Private Function IsEntrySignalReceived(ByVal currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime)
        If Not _ParentStrategy.IsTradeActive(GetCurrentTick(_TradingSymbol, currentTickTime), Trade.TypeOfTrade.CNC) AndAlso
            _SignalPayload.ContainsKey(currentMinute) Then
            Dim currentCandle As Payload = _SignalPayload(currentMinute)
            If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
                Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
                If signalCandle.CandleColor = Color.Green AndAlso IsValidRainbowForBuy(signalCandle) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, signalCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf signalCandle.CandleColor = Color.Red AndAlso IsValidRainbowForSell(signalCandle) Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, signalCandle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsExitSignalReceived(ByVal currentTickTime As Date, ByVal typeOfExit As ExitType, ByVal currentTrade As Trade) As Boolean
        Dim ret As Boolean = False
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim optionType As String = currentTrade.SupportingTradingSymbol.Substring(currentTrade.SupportingTradingSymbol.Count - 2)
            If typeOfExit = ExitType.Target Then
                If currentTrade.EntryType = Trade.TypeOfEntry.Fresh Then      'Fresh Trade
                    ret = GetFreshTradePL(currentTrade, currentTickTime) > Math.Abs(currentTrade.PreviousLoss)
                Else    'SL Makeup Trade
                    ret = GetLossMakeupTradePL(currentTrade, currentTickTime) > Math.Abs(currentTrade.PreviousLoss)
                End If
            ElseIf typeOfExit = ExitType.Reverse Then
                Dim currentCandle As Payload = _SignalPayload(_ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime))
                If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
                    Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
                    If signalCandle.PayloadDate > currentTrade.SignalCandle.PayloadDate Then
                        If signalCandle.CandleColor = Color.Green AndAlso IsValidRainbowForBuy(signalCandle) Then
                            ret = True
                        ElseIf signalCandle.CandleColor = Color.Red AndAlso IsValidRainbowForSell(signalCandle) Then
                            ret = True
                        End If
                    End If
                End If
            ElseIf typeOfExit = ExitType.HalfPremium Then
                Dim currentOptTick As Payload = GetCurrentTick(currentTrade.SupportingTradingSymbol, currentTickTime)
                If _userInputs.HalfPremiumExit Then
                    If currentOptTick.Open <= currentTrade.EntryPrice / 2 Then
                        ret = True
                    End If
                Else
                    If currentOptTick.Open <= 0.05 Then
                        ret = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLossMakeupTradePL(ByVal currentTrade As Trade, ByVal currentTickTime As Date) As Decimal
        Dim ret As Decimal = 0
        Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByChildTag(currentTrade.ChildTag, _TradingSymbol)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            For Each runningTrade In allTrades
                If runningTrade.EntryType = Trade.TypeOfEntry.LossMakeup Then
                    If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentFOTick.Open, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        End If
                    ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, currentFOTick.Open, runningTrade.EntryPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.ExitPrice, runningTrade.EntryPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetFreshTradePL(ByVal currentTrade As Trade, ByVal currentTickTime As Date) As Decimal
        Dim ret As Decimal = 0
        Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByChildTag(currentTrade.ChildTag, _TradingSymbol)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            For Each runningTrade In allTrades
                If runningTrade.EntryType = Trade.TypeOfEntry.Fresh Then
                    If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentFOTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += runningTrade.PLAfterBrokerage
                        End If
                    ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, currentFOTick.Open, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += runningTrade.PLAfterBrokerage
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function

#Region "Public Functions"
    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime)
        If currentTickTime >= _TradeStartTime Then
            Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = IsEntrySignalReceived(currentTickTime)
            If ret IsNot Nothing AndAlso ret.Item1 Then
                EnterTrade(ret.Item2, currentTickTime, ret.Item3)
            End If
        End If
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTickTime As Date, ByVal availableTrades As List(Of Trade)) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If availableTrades IsNot Nothing AndAlso availableTrades.Count > 0 Then
            'Set Drawup Drawdown
            For Each runningTrade In availableTrades
                If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                    If runningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If currentFOTick.Open > runningTrade.MaxDrawUp Then runningTrade.MaxDrawUp = currentFOTick.Open
                        If currentFOTick.Open < runningTrade.MaxDrawDown Then runningTrade.MaxDrawDown = currentFOTick.Open
                    ElseIf runningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If currentFOTick.Open < runningTrade.MaxDrawUp Then runningTrade.MaxDrawUp = currentFOTick.Open
                        If currentFOTick.Open > runningTrade.MaxDrawDown Then runningTrade.MaxDrawDown = currentFOTick.Open
                    End If
                End If
            Next

            'Target Exit
            For Each runningTrade In availableTrades
                If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    If IsExitSignalReceived(currentTickTime, ExitType.Target, runningTrade) Then
                        Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                        _ParentStrategy.ExitTradeByForce(runningTrade, currentFOTick, "Target Hit")
                    End If
                End If
            Next

            'Reverse Exit
            If currentTickTime >= _TradeStartTime Then
                For Each runningTrade In availableTrades
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        If IsExitSignalReceived(currentTickTime, ExitType.Reverse, runningTrade) Then
                            Dim currentFOTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            _ParentStrategy.ExitTradeByForce(runningTrade, currentFOTick, "Reverse Exit")
                        End If
                    End If
                Next
            End If

            'Contract Rollover
            If currentTickTime.Hour >= 15 AndAlso currentTickTime.Minute >= 29 Then
                For Each runningTrade In availableTrades
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                        Dim currentOptionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
                        If Not runningTrade.SupportingTradingSymbol.Contains(currentOptionExpiryString) Then
                            Dim currentOptTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            _ParentStrategy.ExitTradeByForce(runningTrade, currentOptTick, "Contract Rollover")

                            Dim optionType As String = runningTrade.SupportingTradingSymbol.Substring(runningTrade.SupportingTradingSymbol.Count - 2)
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, currentOptionExpiryString, currentSpotTick.Open, optionType, runningTrade.SpotATR)

                            currentOptTick = GetCurrentTick(optionTradingSymbol, currentTickTime)
                            EnterDuplicateTrade(runningTrade, currentOptTick, False)
                        End If
                    End If
                Next
            End If

            'Half Premium
            If currentTickTime.Hour >= 15 AndAlso currentTickTime.Minute >= 29 Then
                For Each runningTrade In availableTrades
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        If IsExitSignalReceived(currentTickTime, ExitType.HalfPremium, runningTrade) Then
                            Dim currentOptTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                            Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                            Dim currentOptionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _TradingDate)
                            Dim optionType As String = runningTrade.SupportingTradingSymbol.Substring(runningTrade.SupportingTradingSymbol.Count - 2)
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, currentOptionExpiryString, currentSpotTick.Open, optionType, runningTrade.SpotATR)
                            If optionTradingSymbol <> runningTrade.SupportingTradingSymbol Then
                                _ParentStrategy.ExitTradeByForce(runningTrade, currentOptTick, "Half Premium")

                                currentOptTick = GetCurrentTick(optionTradingSymbol, currentTickTime)
                                EnterDuplicateTrade(runningTrade, currentOptTick, True)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Function

    Private Function GetCurrentTick(ByVal tradingSymbol As String, ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _dependentInstruments Is Nothing OrElse Not _dependentInstruments.ContainsKey(tradingSymbol) Then
            Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
            If tradingSymbol.EndsWith("FUT") Then
                inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures, tradingSymbol, _TradingDate.AddDays(-30), _TradingDate)
            Else
                inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol, _TradingDate.AddDays(-30), _TradingDate)
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
#End Region

#Region "Place Trade"
    Private Function EnterTrade(ByVal signalCandle As Payload, ByVal currentTickTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim childTag As String = System.Guid.NewGuid.ToString()
        Dim parentTag As String = childTag
        Dim tradeNumber As Integer = 1
        Dim entryType As Trade.TypeOfEntry = Trade.TypeOfEntry.Fresh
        Dim spotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
        Dim spotATR As Decimal = _atrPayload(signalCandle.PayloadDate)
        Dim qunaity As Integer = _LotSize
        Dim lossToRecover As Decimal = _ParentStrategy.CalculatePL(_TradingSymbol, spotTick.Open, ConvertFloorCeling(spotTick.Open + spotATR, _ParentStrategy.TickSize, RoundOfType.Celing), qunaity, _LotSize, Trade.TypeOfStock.Cash)

        Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC)
        If lastTrade IsNot Nothing Then
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByParentTag(lastTrade.ParentTag, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                Dim pl As Decimal = allTrades.Sum(Function(x)
                                                      Return x.PLAfterBrokerage
                                                  End Function)
                If pl < 0 Then
                    parentTag = lastTrade.ParentTag
                    tradeNumber = lastTrade.TradeNumber + 1
                    entryType = Trade.TypeOfEntry.LossMakeup
                    lossToRecover = pl
                    qunaity = lastTrade.Quantity * 2
                End If
            End If
        End If

        ret = EnterBuySellTrade(signalCandle, spotTick, spotATR, childTag, parentTag, tradeNumber, entryType, direction, lossToRecover, qunaity)

        Return ret
    End Function

    Private Function EnterBuySellTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload, ByVal currentSpotATR As Decimal,
                                       ByVal childTag As String, ByVal parentTag As String,
                                       ByVal tradeNumber As Integer, ByVal entryType As Trade.TypeOfEntry,
                                       ByVal direction As Trade.TradeExecutionDirection,
                                       ByVal previousLoss As Decimal, ByVal quantity As Integer) As Boolean
        Dim ret As Boolean = False
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentSpotTick.PayloadDate)
        Dim optionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
        Dim currentOptionTradingSymbol As String = Nothing
        If direction = Trade.TradeExecutionDirection.Buy Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "PE", currentSpotATR)
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "CE", currentSpotATR)
        End If
        Dim currentFutureTradingSymbol As String = GetFutureInstrumentNameFromCore(_TradingSymbol, _NextTradingDay)

        If currentOptionTradingSymbol IsNot Nothing AndAlso currentFutureTradingSymbol IsNot Nothing Then
            Dim currentFutTick As Payload = GetCurrentTick(currentFutureTradingSymbol, currentSpotTick.PayloadDate)
            Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
            If currentOptTick IsNot Nothing AndAlso currentFutTick IsNot Nothing Then
                'Dim quantity As Integer = _LotSize
                'If entryType = Trade.TypeOfEntry.LossMakeup Then
                '    Dim entryPrice As Decimal = currentOptTick.Open
                '    Dim targetPrice As Decimal = ConvertFloorCeling(entryPrice + currentSpotATR / _userInputs.SpotToOptionDelta, _ParentStrategy.TickSize, RoundOfType.Celing)
                '    For ctr As Integer = 1 To Integer.MaxValue
                '        Dim pl As Decimal = _ParentStrategy.CalculatePL(_TradingSymbol, entryPrice, targetPrice, ctr * _LotSize, _LotSize, Trade.TypeOfStock.Options)
                '        If pl >= Math.Abs(previousLoss) Then
                '            If _userInputs.ExitAtATRPL Then previousLoss = (pl - 1) * -1
                '            quantity = ctr * _LotSize + _LotSize
                '            Exit For
                '        End If
                '    Next
                'End If
                Dim tgt As Decimal = currentFutTick.Open + 100000
                Dim sl As Decimal = currentFutTick.Open - 100000
                If direction = Trade.TradeExecutionDirection.Sell Then
                    tgt = currentFutTick.Open - 100000
                    sl = currentFutTick.Open + 100000
                End If

                Dim runningFutTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                         tradingSymbol:=_TradingSymbol,
                                                         stockType:=Trade.TypeOfStock.Futures,
                                                         orderType:=Trade.TypeOfOrder.Market,
                                                         tradingDate:=currentFutTick.PayloadDate,
                                                         entryDirection:=direction,
                                                         entryPrice:=currentFutTick.Open,
                                                         entryBuffer:=0,
                                                         squareOffType:=Trade.TypeOfTrade.CNC,
                                                         entryCondition:=Trade.TradeEntryCondition.Original,
                                                         entryRemark:="Original Entry",
                                                         quantity:=quantity,
                                                         lotSize:=_LotSize,
                                                         potentialTarget:=tgt,
                                                         targetRemark:=100000,
                                                         potentialStopLoss:=sl,
                                                         stoplossBuffer:=0,
                                                         slRemark:=100000,
                                                         signalCandle:=signalCandle)

                runningFutTrade.UpdateTrade(ChildTag:=childTag,
                                            ParentTag:=parentTag,
                                            TradeNumber:=tradeNumber,
                                            EntryType:=entryType,
                                            SpotPrice:=currentSpotTick.Open,
                                            SpotATR:=currentSpotATR,
                                            PreviousLoss:=previousLoss,
                                            SupportingTradingSymbol:=currentFutureTradingSymbol)

                Dim runningOptTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                         tradingSymbol:=_TradingSymbol,
                                                         stockType:=Trade.TypeOfStock.Options,
                                                         orderType:=Trade.TypeOfOrder.Market,
                                                         tradingDate:=currentOptTick.PayloadDate,
                                                         entryDirection:=Trade.TradeExecutionDirection.Buy,
                                                         entryPrice:=currentOptTick.Open,
                                                         entryBuffer:=0,
                                                         squareOffType:=Trade.TypeOfTrade.CNC,
                                                         entryCondition:=Trade.TradeEntryCondition.Original,
                                                         entryRemark:="Original Entry",
                                                         quantity:=quantity,
                                                         lotSize:=_LotSize,
                                                         potentialTarget:=currentOptTick.Open + 100000,
                                                         targetRemark:=100000,
                                                         potentialStopLoss:=currentOptTick.Open - 100000,
                                                         stoplossBuffer:=0,
                                                         slRemark:=100000,
                                                         signalCandle:=signalCandle)

                runningOptTrade.UpdateTrade(ChildTag:=childTag,
                                            ParentTag:=parentTag,
                                            TradeNumber:=tradeNumber,
                                            EntryType:=entryType,
                                            SpotPrice:=currentSpotTick.Open,
                                            SpotATR:=currentSpotATR,
                                            PreviousLoss:=previousLoss,
                                            SupportingTradingSymbol:=currentOptionTradingSymbol)

                If _ParentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                    If _ParentStrategy.PlaceOrModifyOrder(runningOptTrade, Nothing) Then
                        If _ParentStrategy.EnterTradeIfPossible(_TradingSymbol, _TradingDate, runningFutTrade, currentFutTick) Then
                            If _ParentStrategy.EnterTradeIfPossible(_TradingSymbol, _TradingDate, runningOptTrade, currentOptTick) Then
                                ret = True
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Console.WriteLine(String.Format("Option Symbol not available. Core:{0}, DateTime:{1}", _TradingSymbol, currentSpotTick.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")))
            ret = True
        End If
        Return ret
    End Function

    Private Function EnterDuplicateTrade(ByVal existingTrade As Trade, ByVal currentTick As Payload, ByVal increaseQuantityIfRequired As Boolean) As Boolean
        Dim ret As Boolean = False
        Try
            Dim quantity As Integer = existingTrade.Quantity
            Dim remark As String = Nothing
            Dim runningTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
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
                                                    quantity:=quantity,
                                                    lotSize:=existingTrade.LotSize,
                                                    potentialTarget:=currentTick.Open + 100000,
                                                    targetRemark:=100000,
                                                    potentialStopLoss:=currentTick.Open - 100000,
                                                    stoplossBuffer:=0,
                                                    slRemark:=100000,
                                                    signalCandle:=existingTrade.SignalCandle)

            runningTrade.UpdateTrade(ChildTag:=existingTrade.ChildTag,
                                    ParentTag:=existingTrade.ParentTag,
                                    TradeNumber:=existingTrade.TradeNumber,
                                    EntryType:=existingTrade.EntryType,
                                    SpotPrice:=existingTrade.SpotPrice,
                                    SpotATR:=existingTrade.SpotATR,
                                    PreviousLoss:=existingTrade.PreviousLoss,
                                    SupportingTradingSymbol:=currentTick.TradingSymbol)

            If remark IsNot Nothing AndAlso remark.Trim <> "" Then
                runningTrade.UpdateTrade(Remark2:=remark)
            End If

            If _ParentStrategy.PlaceOrModifyOrder(runningTrade, Nothing) Then
                ret = _ParentStrategy.EnterTradeIfPossible(runningTrade.TradingSymbol, _TradingDate, runningTrade, currentTick)
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return ret
    End Function
#End Region

#Region "Contract Helper"
    Private Function GetCurrentATMOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal, ByVal optionType As String, ByVal atr As Decimal) As String
        If optionType = "PE" Then
            Return GetCurrentATMPEOption(currentTickTime, expiryString, price, atr)
        Else
            Return GetCurrentATMCEOption(currentTickTime, expiryString, price, atr)
        End If
    End Function

    Private Function GetCurrentATMPEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal, ByVal atr As Decimal) As String
        Dim ret As String = Nothing
        Dim query As String = Nothing
        query = "SELECT `TradingSymbol`,SUM(`Volume`) Vol
                FROM `intraday_prices_opt_futures`
                WHERE `SnapshotDate`='{0}'
                AND `SnapshotTime`<='{1}'
                AND `TradingSymbol` LIKE '{2}%{3}'
                GROUP BY `TradingSymbol`
                ORDER BY Vol DESC"
        query = String.Format(query, currentTickTime.ToString("yyyy-MM-dd"), currentTickTime.AddMinutes(-1).ToString("HH:mm:ss"), expiryString, "PE")

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
                If _userInputs.OptionStrikeDistance > 0 Then
                    Dim ctr As Integer = 0
                    For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                        If runningContract.Key <= price Then
                            ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                            ctr += 1
                            If ctr = Math.Abs(_userInputs.OptionStrikeDistance) Then Exit For
                        End If
                    Next
                Else
                    Dim ctr As Integer = 0
                    For Each runningContract In contracts.OrderBy(Function(x)
                                                                      Return x.Key
                                                                  End Function)
                        If runningContract.Key >= price Then
                            ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                            ctr += 1
                            If ctr = Math.Abs(_userInputs.OptionStrikeDistance) Then Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetCurrentATMCEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal, ByVal atr As Decimal) As String
        Dim ret As String = Nothing
        Dim query As String = Nothing
        query = "SELECT `TradingSymbol`,SUM(`Volume`) Vol
                FROM `intraday_prices_opt_futures`
                WHERE `SnapshotDate`='{0}'
                AND `SnapshotTime`<='{1}'
                AND `TradingSymbol` LIKE '{2}%{3}'
                GROUP BY `TradingSymbol`
                ORDER BY Vol DESC"
        query = String.Format(query, currentTickTime.ToString("yyyy-MM-dd"), currentTickTime.AddMinutes(-1).ToString("HH:mm:ss"), expiryString, "CE")

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
                If _userInputs.OptionStrikeDistance > 0 Then
                    Dim ctr As Integer = 0
                    For Each runningContract In contracts.OrderBy(Function(x)
                                                                      Return x.Key
                                                                  End Function)
                        If runningContract.Key >= price Then
                            ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                            ctr += 1
                            If ctr = Math.Abs(_userInputs.OptionStrikeDistance) Then Exit For
                        End If
                    Next
                Else
                    Dim ctr As Integer = 0
                    For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                        If runningContract.Key <= price Then
                            ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                            ctr += 1
                            If ctr = Math.Abs(_userInputs.OptionStrikeDistance) Then Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetOptionInstrumentExpiryString(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
        Dim ret As String = Nothing
        Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
        If tradingDate.Date > lastThursday.Date.AddDays(-2) Then
            ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.AddDays(10).ToString("yyMMM")).ToUpper
        Else
            ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
        End If
        Return ret
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
#End Region

#Region "Required Functions"
    Private Function IsValidRainbowForBuy(ByVal signalCandle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim rainbow As Indicator.RainbowMA = _rainbowPayload(signalCandle.PayloadDate)
        If signalCandle.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
            For Each runningPayload In _SignalPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                If runningPayload.Key < signalCandle.PayloadDate Then
                    rainbow = _rainbowPayload(runningPayload.Key)
                    If runningPayload.Value.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                        If runningPayload.Value.CandleColor = Color.Green Then
                            If runningPayload.Key <> signalCandle.PayloadDate Then
                                ret = False
                                Exit For
                            End If
                        End If
                    ElseIf runningPayload.Value.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
                        If runningPayload.Value.CandleColor = Color.Red Then
                            ret = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function
    Private Function IsValidRainbowForSell(ByVal signalCandle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim rainbow As Indicator.RainbowMA = _rainbowPayload(signalCandle.PayloadDate)
        If signalCandle.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
            For Each runningPayload In _SignalPayload.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                If runningPayload.Key < signalCandle.PayloadDate Then
                    rainbow = _rainbowPayload(runningPayload.Key)
                    If runningPayload.Value.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
                        If runningPayload.Value.CandleColor = Color.Red Then
                            If runningPayload.Key <> signalCandle.PayloadDate Then
                                ret = False
                                Exit For
                            End If
                        End If
                    ElseIf runningPayload.Value.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                        If runningPayload.Value.CandleColor = Color.Green Then
                            ret = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return ret
    End Function
#End Region

End Class