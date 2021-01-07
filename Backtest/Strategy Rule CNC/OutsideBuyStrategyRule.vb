Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper

Public Class OutsideBuyStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
    End Class

    Private Enum ExitType
        Target = 1
        Reverse
        HalfPremium
    End Enum
#End Region

    Private _eodPayload As Dictionary(Of Date, Payload)
    Private _hkPayload As Dictionary(Of Date, Payload)
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
        Indicator.HeikenAshi.ConvertToHeikenAshi(_eodPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _eodPayload, _atrPayload)

        Dim allRunningTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
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
            _dependentInstruments.Add(_TradingSymbol, _InputPayload)
        Else
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            _dependentInstruments.Add(_TradingSymbol, _InputPayload)
        End If
    End Sub

    Private Function IsEntrySignalReceived(ByVal currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If Not _ParentStrategy.IsTradeActive(GetCurrentTick(_TradingSymbol, currentTickTime), Trade.TypeOfTrade.CNC) Then
            If _hkPayload IsNot Nothing AndAlso _hkPayload.ContainsKey(_TradingDate) Then
                Dim hkCandle As Payload = _hkPayload(_TradingDate)
                Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC)
                Dim optionType As String = ""
                If lastTrade IsNot Nothing AndAlso lastTrade.SupportingTradingSymbol IsNot Nothing Then
                    optionType = lastTrade.SupportingTradingSymbol.Substring(lastTrade.SupportingTradingSymbol.Count - 2)
                End If
                If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish AndAlso optionType <> "CE" Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, hkCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish AndAlso optionType <> "PE" Then
                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, hkCandle, Trade.TradeExecutionDirection.Sell)
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
                If currentTrade.Supporting2.Contains("Fresh") Then
                    Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                    If optionType = "CE" AndAlso currentSpotTick.Open >= Val(currentTrade.Supporting3) + Val(currentTrade.Supporting4) * _userInputs.ATRMultiplier Then
                        ret = True
                    ElseIf optionType = "PE" AndAlso currentSpotTick.Open <= Val(currentTrade.Supporting3) - Val(currentTrade.Supporting4) * _userInputs.ATRMultiplier Then
                        ret = True
                    End If
                Else
                    Dim lossToRecover As Decimal = currentTrade.Supporting5
                    Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByTag(currentTrade.Tag, _TradingSymbol)
                    If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                        Dim plAchieved As Decimal = 0
                        For Each runningTrade In allTrades
                            If Not runningTrade.Supporting2.Contains("Fresh") Then
                                plAchieved += runningTrade.PLAfterBrokerage
                            End If
                        Next
                        If plAchieved >= Math.Abs(lossToRecover) Then
                            ret = True
                        End If
                    End If
                End If
            ElseIf typeOfExit = ExitType.Reverse Then
                If _hkPayload IsNot Nothing AndAlso _hkPayload.ContainsKey(_TradingDate) Then
                    Dim hkCandle As Payload = _hkPayload(_TradingDate)
                    If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish AndAlso optionType <> "CE" Then
                        ret = True
                    ElseIf hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish AndAlso optionType <> "PE" Then
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
                            EnterDuplicateTrade(runningTrade, currentOptTick)
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
                                EnterDuplicateTrade(runningTrade, currentOptTick)
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
        Dim tradeID As String = System.Guid.NewGuid.ToString()
        Dim tradeNumber As Integer = 1
        Dim remark As String = "Fresh"

        Dim spotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
        Dim closeTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Close)
        If closeTrades IsNot Nothing AndAlso closeTrades.Count Then
            Dim pl As Decimal = 0
            Dim qty As Integer = 0
            For Each runningTrade In closeTrades.OrderByDescending(Function(x)
                                                                       Return x.EntryTime
                                                                   End Function)
                If runningTrade.ExitRemark.ToUpper = "TARGET HIT" Then
                    Exit For
                Else
                    pl += runningTrade.PLAfterBrokerage
                    qty += runningTrade.Quantity
                End If
            Next

            If pl < 0 Then
                'tradeID = closeTrades.LastOrDefault.Tag
                tradeNumber = Val(closeTrades.LastOrDefault.Supporting1) + 1
                remark = "SL Makeup + Fresh"

                EnterBuyTrade(signalCandle, spotTick, tradeID, tradeNumber, "SL Makeup", direction, qty, spotTick.Open, _atrPayload(signalCandle.PayloadDate), pl)
            End If
        End If

        ret = EnterBuyTrade(signalCandle, spotTick, tradeID, tradeNumber, remark, direction, _LotSize, spotTick.Open, _atrPayload(signalCandle.PayloadDate), 0)

        Return ret
    End Function

    Private Function EnterBuyTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload,
                                   ByVal tradeID As String, ByVal tradeNumber As Integer, ByVal remark As String,
                                   ByVal diretion As Trade.TradeExecutionDirection, ByVal quantity As Integer,
                                   ByVal spotPrice As Decimal, ByVal spotAtr As Decimal, ByVal previousLoss As Decimal) As Boolean
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

                runningOptTrade.UpdateTrade(Tag:=tradeID,
                                            Supporting1:=tradeNumber,
                                            Supporting2:=remark,
                                            Supporting3:=spotPrice,
                                            Supporting4:=spotAtr,
                                            Supporting5:=previousLoss,
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
                                                      quantity:=existingTrade.Quantity,
                                                      lotSize:=existingTrade.LotSize,
                                                      potentialTarget:=currentTick.Open + 100000,
                                                      targetRemark:=100000,
                                                      potentialStopLoss:=currentTick.Open - 100000,
                                                      stoplossBuffer:=0,
                                                      slRemark:=100000,
                                                      signalCandle:=existingTrade.SignalCandle)

            runningFutTrade.UpdateTrade(Tag:=existingTrade.Tag,
                                        Supporting1:=existingTrade.Supporting1,
                                        Supporting2:=existingTrade.Supporting2,
                                        Supporting3:=existingTrade.Supporting3,
                                        Supporting4:=existingTrade.Supporting4,
                                        Supporting5:=existingTrade.Supporting5,
                                        Supporting6:=existingTrade.Supporting6,
                                        Supporting7:=existingTrade.Supporting7,
                                        Supporting8:=existingTrade.Supporting8,
                                        Supporting9:=existingTrade.Supporting9,
                                        SupportingTradingSymbol:=currentTick.TradingSymbol)

            If _ParentStrategy.PlaceOrModifyOrder(runningFutTrade, Nothing) Then
                ret = _ParentStrategy.EnterTradeIfPossible(runningFutTrade.TradingSymbol, _TradingDate, runningFutTrade, currentTick)
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
                For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                            Return x.Value
                                                                        End Function)
                    If runningContract.Key <= price Then
                        ret = String.Format("{0}{1}PE", expiryString, runningContract.Key)
                        Exit For
                    End If
                Next
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
                For Each runningContract In contracts.OrderByDescending(Function(x)
                                                                            Return x.Value
                                                                        End Function)
                    If runningContract.Key >= price Then
                        ret = String.Format("{0}{1}CE", expiryString, runningContract.Key)
                        Exit For
                    End If
                Next
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
#End Region

End Class
