Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class PivotTrendOutsideBuyStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
        Public SpotToOptionDelta As Decimal
        Public ExitAtATRPL As Boolean
        Public MartingaleOnLossMakeup As Boolean
    End Class

    Private Enum ExitType
        Target = 1
        Reverse
        HalfPremium
    End Enum
#End Region

    Private _eodPayload As Dictionary(Of Date, Payload)
    Private _pivotTrendPayload As Dictionary(Of Date, Color)
    Private _atrPayload As Dictionary(Of Date, Decimal)

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

        _eodPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, _TradingSymbol, _TradingDate.AddDays(-200), _TradingDate)
        Indicator.PivotHighLow.CalculatePivotHighLowTrend(4, 3, _eodPayload, Nothing, Nothing, _pivotTrendPayload)
        Indicator.ATR.CalculateATR(14, _eodPayload, _atrPayload)

        Dim allRunningTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If allRunningTrades IsNot Nothing AndAlso allRunningTrades.Count > 0 Then
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            For Each runningTrade In allRunningTrades
                If Not _dependentInstruments.ContainsKey(runningTrade.SupportingTradingSymbol) Then
                    Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                    inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-30), _TradingDate)
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
        If Not _ParentStrategy.IsTradeActive(GetCurrentTick(_TradingSymbol, currentTickTime), Trade.TypeOfTrade.CNC) Then
            If _pivotTrendPayload IsNot Nothing AndAlso _pivotTrendPayload.ContainsKey(_TradingDate) Then
                Dim trend As Color = _pivotTrendPayload(_TradingDate)
                Dim previousTrend As Color = _pivotTrendPayload(_eodPayload(_TradingDate).PreviousCandlePayload.PayloadDate)
                Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC)
                Dim optionType As String = ""
                If lastTrade IsNot Nothing AndAlso lastTrade.SupportingTradingSymbol IsNot Nothing Then
                    optionType = lastTrade.SupportingTradingSymbol.Substring(lastTrade.SupportingTradingSymbol.Count - 2)
                End If
                If trend = Color.Green AndAlso previousTrend = Color.Red AndAlso optionType <> "CE" Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, _eodPayload(_TradingDate), Trade.TradeExecutionDirection.Buy)
                ElseIf trend = Color.Red AndAlso previousTrend = Color.Green AndAlso optionType <> "PE" Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, _eodPayload(_TradingDate), Trade.TradeExecutionDirection.Sell)
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
                    Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                    If optionType = "CE" AndAlso currentSpotTick.Open >= currentTrade.SpotPrice + currentTrade.SpotATR * _userInputs.ATRMultiplier Then
                        Dim pl As Decimal = GetFreshTradePL(currentTrade, currentTickTime)
                        If pl > 0 Then
                            ret = True
                        Else
                            If currentTrade.Remark1 Is Nothing Then
                                currentTrade.UpdateTrade(Remark1:=String.Format("ATR Reached but PL {0}", pl))
                            End If
                        End If
                    ElseIf optionType = "PE" AndAlso currentSpotTick.Open <= currentTrade.SpotPrice - currentTrade.SpotATR * _userInputs.ATRMultiplier Then
                        Dim pl As Decimal = GetFreshTradePL(currentTrade, currentTickTime)
                        If pl > 0 Then
                            ret = True
                        Else
                            If currentTrade.Remark1 Is Nothing Then
                                currentTrade.UpdateTrade(Remark1:=String.Format("ATR Reached but PL {0}", pl))
                            End If
                        End If
                    End If
                Else    'SL Makeup Trade
                    ret = GetLossMakeupTradePL(currentTrade, currentTickTime) > Math.Abs(currentTrade.PreviousLoss)
                End If
            ElseIf typeOfExit = ExitType.Reverse Then
                If _pivotTrendPayload IsNot Nothing AndAlso _pivotTrendPayload.ContainsKey(_TradingDate) Then
                    Dim trend As Color = _pivotTrendPayload(_TradingDate)
                    Dim previousTrend As Color = _pivotTrendPayload(_eodPayload(_TradingDate).PreviousCandlePayload.PayloadDate)
                    If trend = Color.Green AndAlso previousTrend = Color.Red AndAlso optionType <> "CE" Then
                        ret = True
                    ElseIf trend = Color.Red AndAlso previousTrend = Color.Green AndAlso optionType <> "PE" Then
                        ret = True
                    End If
                End If
            ElseIf typeOfExit = ExitType.HalfPremium Then
                Dim currentOptTick As Payload = GetCurrentTick(currentTrade.SupportingTradingSymbol, currentTickTime)
                If currentOptTick.Open <= currentTrade.EntryPrice / 2 Then
                    ret = True
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
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim currentOptTick As Payload = GetCurrentTick(currentTrade.SupportingTradingSymbol, currentTickTime)
                        If _userInputs.MartingaleOnLossMakeup Then
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentOptTick.Open, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentOptTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                        End If
                    Else
                        If _userInputs.MartingaleOnLossMakeup Then
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
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
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim currentOptTick As Payload = GetCurrentTick(currentTrade.SupportingTradingSymbol, currentTickTime)
                        ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentOptTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                    Else
                        ret += runningTrade.PLAfterBrokerage
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
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, currentOptionExpiryString, currentSpotTick.Open, optionType)

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
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, currentOptionExpiryString, currentSpotTick.Open, optionType)
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
#End Region

#Region "Place Trade"
    Private Function EnterTrade(ByVal signalCandle As Payload, ByVal currentTickTime As Date, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim childTag As String = System.Guid.NewGuid.ToString()
        Dim parentTag As String = childTag
        Dim tradeNumber As Integer = 1
        Dim entryType As Trade.TypeOfEntry = Trade.TypeOfEntry.Fresh
        Dim lossToRecover As Decimal = 0

        Dim spotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
        Dim spotATR As Decimal = _atrPayload(signalCandle.PayloadDate)
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
                End If
            End If
        End If

        ret = EnterBuyTrade(signalCandle, spotTick, spotATR, childTag, parentTag, tradeNumber, entryType, direction, lossToRecover)

        Return ret
    End Function

    Private Function EnterBuyTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload, ByVal currentSpotATR As Decimal,
                                   ByVal childTag As String, ByVal parentTag As String,
                                   ByVal tradeNumber As Integer, ByVal entryType As Trade.TypeOfEntry,
                                   ByVal diretion As Trade.TradeExecutionDirection, ByVal previousLoss As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentSpotTick.PayloadDate)
        Dim optionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
        Dim currentOptionTradingSymbol As String = Nothing
        If diretion = Trade.TradeExecutionDirection.Buy Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "CE")
        ElseIf diretion = Trade.TradeExecutionDirection.Sell Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "PE")
        End If

        If currentOptionTradingSymbol IsNot Nothing Then
            Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
            If currentOptTick IsNot Nothing Then
                Dim quantity As Integer = _LotSize
                If entryType = Trade.TypeOfEntry.LossMakeup Then
                    Dim entryPrice As Decimal = currentOptTick.Open
                    Dim targetPrice As Decimal = ConvertFloorCeling(entryPrice + currentSpotATR / _userInputs.SpotToOptionDelta, _ParentStrategy.TickSize, RoundOfType.Celing)
                    If _userInputs.MartingaleOnLossMakeup Then
                        For ctr As Integer = 1 To Integer.MaxValue
                            Dim pl As Decimal = _ParentStrategy.CalculatePL(_TradingSymbol, entryPrice, targetPrice, ctr * _LotSize, _LotSize, Trade.TypeOfStock.Options)
                            If pl >= Math.Abs(previousLoss) Then
                                If _userInputs.ExitAtATRPL Then previousLoss = (pl - 1) * -1
                                quantity = ctr * _LotSize + _LotSize
                                Exit For
                            End If
                        Next
                    Else
                        Dim pl As Decimal = _ParentStrategy.CalculatePL(_TradingSymbol, entryPrice, targetPrice, _LotSize, _LotSize, Trade.TypeOfStock.Options)
                        previousLoss = previousLoss - pl
                    End If
                End If

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

                If _ParentStrategy.PlaceOrModifyOrder(runningOptTrade, Nothing) Then
                    If _ParentStrategy.EnterTradeIfPossible(_TradingSymbol, _TradingDate, runningOptTrade, currentOptTick) Then
                        ret = True
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
            'If increaseQuantityIfRequired AndAlso _userInputs.IncreaseQuantityWithHalfPremium Then
            '    Dim increasedCapital As Decimal = currentTick.Open * quantity * 2
            '    If increasedCapital <= existingTrade.CapitalRequiredWithMargin * 120 / 100 Then
            '        quantity = quantity * 2
            '    Else
            '        remark = "Unable to increase quantity"
            '    End If
            'End If
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
                Dim upperContract As Dictionary(Of Decimal, Long) = Nothing
                For Each runningContract In contracts.OrderBy(Function(x)
                                                                  Return x.Key
                                                              End Function)
                    If runningContract.Key >= price Then
                        If upperContract Is Nothing Then upperContract = New Dictionary(Of Decimal, Long)
                        upperContract.Add(runningContract.Key, runningContract.Value)

                        If upperContract.Count >= 2 Then Exit For
                    End If
                Next

                Dim lowerContract As Dictionary(Of Decimal, Long) = Nothing
                For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                    If runningContract.Key <= price Then
                        If lowerContract Is Nothing Then lowerContract = New Dictionary(Of Decimal, Long)
                        lowerContract.Add(runningContract.Key, runningContract.Value)

                        If lowerContract.Count >= 2 Then Exit For
                    End If
                Next

                Dim sumVol As Long = 0
                Dim count As Integer = 0
                Dim volList As List(Of Long) = Nothing
                If upperContract IsNot Nothing AndAlso upperContract.Count > 0 Then
                    sumVol += upperContract.Sum(Function(x)
                                                    Return x.Value
                                                End Function)
                    count += upperContract.Count
                    If volList Is Nothing Then volList = New List(Of Long)
                    volList.AddRange(upperContract.Values.ToList)
                End If
                If lowerContract IsNot Nothing AndAlso lowerContract.Count > 0 Then
                    sumVol += lowerContract.Sum(Function(x)
                                                    Return x.Value
                                                End Function)
                    count += lowerContract.Count
                    If volList Is Nothing Then volList = New List(Of Long)
                    volList.AddRange(lowerContract.Values.ToList)
                End If
                Dim avgVol As Decimal = sumVol
                If count > 0 Then
                    avgVol = sumVol / count
                    Dim std As Long = CalculateStandardDeviationPA(volList.ToArray())
                    avgVol = avgVol - std
                End If

                For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                    If runningContract.Key <= price AndAlso runningContract.Value >= avgVol Then
                        ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                        Exit For
                    End If
                Next
                If ret Is Nothing Then
                    For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                                Return x.Key
                                                                            End Function)
                        If runningContract.Key <= price Then
                            ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                            Exit For
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
                Dim upperContract As Dictionary(Of Decimal, Long) = Nothing
                For Each runningContract In contracts.OrderBy(Function(x)
                                                                  Return x.Key
                                                              End Function)
                    If runningContract.Key >= price Then
                        If upperContract Is Nothing Then upperContract = New Dictionary(Of Decimal, Long)
                        upperContract.Add(runningContract.Key, runningContract.Value)

                        If upperContract.Count >= 2 Then Exit For
                    End If
                Next

                Dim lowerContract As Dictionary(Of Decimal, Long) = Nothing
                For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                            Return x.Key
                                                                        End Function)
                    If runningContract.Key <= price Then
                        If lowerContract Is Nothing Then lowerContract = New Dictionary(Of Decimal, Long)
                        lowerContract.Add(runningContract.Key, runningContract.Value)

                        If lowerContract.Count >= 2 Then Exit For
                    End If
                Next

                Dim sumVol As Long = 0
                Dim count As Integer = 0
                Dim volList As List(Of Long) = Nothing
                If upperContract IsNot Nothing AndAlso upperContract.Count > 0 Then
                    sumVol += upperContract.Sum(Function(x)
                                                    Return x.Value
                                                End Function)
                    count += upperContract.Count
                    If volList Is Nothing Then volList = New List(Of Long)
                    volList.AddRange(upperContract.Values.ToList)
                End If
                If lowerContract IsNot Nothing AndAlso lowerContract.Count > 0 Then
                    sumVol += lowerContract.Sum(Function(x)
                                                    Return x.Value
                                                End Function)
                    count += lowerContract.Count
                    If volList Is Nothing Then volList = New List(Of Long)
                    volList.AddRange(lowerContract.Values.ToList)
                End If
                Dim avgVol As Decimal = sumVol
                If count > 0 Then
                    avgVol = sumVol / count
                    Dim std As Long = CalculateStandardDeviationPA(volList.ToArray())
                    avgVol = avgVol - std
                End If

                For Each runningContract In contracts.OrderBy(Function(x)
                                                                  Return x.Key
                                                              End Function)
                    If runningContract.Key >= price AndAlso runningContract.Value >= avgVol Then
                        ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                        Exit For
                    End If
                Next
                If ret Is Nothing Then
                    For Each runningContract In contracts.OrderBy(Function(x)
                                                                      Return x.Key
                                                                  End Function)
                        If runningContract.Key >= price Then
                            ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                            Exit For
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

    Private Function CalculateStandardDeviationPA(ParamArray numbers() As Long) As Long
        Dim ret As Long = Nothing
        If numbers.Count > 0 Then
            Dim sum As Long = 0
            For i = 0 To numbers.Count - 1
                sum = sum + numbers(i)
            Next
            Dim mean As Long = sum / numbers.Count
            Dim sumVariance As Long = 0
            For j = 0 To numbers.Count - 1
                sumVariance = sumVariance + Math.Pow((numbers(j) - mean), 2)
            Next
            Dim sampleVariance As Long = sumVariance / numbers.Count
            Dim standardDeviation As Long = Math.Sqrt(sampleVariance)
            ret = standardDeviation
        End If
        Return ret
    End Function
#End Region

End Class