Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers
Imports Backtest.StrategyHelper

Public Class OptionBuyStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public ATRMultiplier As Decimal
        Public SpotToOptionDelta As Decimal
        Public ExitAtATRPL As Boolean
        Public OptionStrikeDistance As Integer
        Public NumberOfActiveStock As Integer
    End Class

    Private Enum ExitType
        Target = 1
    End Enum
#End Region

    Private _eodPayload As Dictionary(Of Date, Payload)
    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)

    Private _smiPayload As Dictionary(Of Date, Decimal)

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

        _eodPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, _TradingSymbol, _TradingDate.AddDays(-180), _TradingDate)
        Indicator.ATR.CalculateATR(14, _eodPayload, _atrPayload)
        Indicator.HeikenAshi.ConvertToHeikenAshi(_eodPayload, _hkPayload)

        Indicator.SMI.CalculateSMI(10, 3, 3, 10, _SignalPayload, _smiPayload, Nothing)

        Dim allRunningTrades As List(Of Trade) = _ParentStrategy.GetSpecificTrades(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If allRunningTrades IsNot Nothing AndAlso allRunningTrades.Count > 0 Then
            _dependentInstruments = New Dictionary(Of String, Dictionary(Of Date, Payload))
            For Each runningTrade In allRunningTrades
                If Not _dependentInstruments.ContainsKey(runningTrade.SupportingTradingSymbol) Then
                    Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                    inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, runningTrade.SupportingTradingSymbol, _TradingDate.AddDays(-50), _TradingDate)
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
        If Not _ParentStrategy.IsTradeActive(GetCurrentTick(_TradingSymbol, currentTickTime), Trade.TypeOfTrade.CNC) AndAlso
            currentTickTime >= _TradeStartTime Then
            Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime)
            If _SignalPayload.ContainsKey(currentMinute) Then
                Dim currentMinuteCandle As Payload = _SignalPayload(currentMinute)
                Dim currentDayHKCandle As Payload = _hkPayload(_TradingDate)
                Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC)
                If lastTrade Is Nothing OrElse lastTrade.EntryTime.Date <> _TradingDate.Date Then
                    If currentDayHKCandle.PreviousCandlePayload IsNot Nothing AndAlso currentDayHKCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                        Dim lossToRecover As Decimal = 0
                        If lastTrade IsNot Nothing AndAlso lastTrade.ExitRemark IsNot Nothing AndAlso lastTrade.ExitRemark.ToUpper <> "TARGET HIT" Then
                            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByParentTag(lastTrade.ParentTag, _TradingSymbol)
                            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                                Dim pl As Decimal = allTrades.Sum(Function(x)
                                                                      Return x.PLAfterBrokerage
                                                                  End Function)
                                If pl < 0 Then
                                    lossToRecover = pl
                                End If
                            End If
                        End If

                        If lossToRecover < 0 AndAlso
                            ((currentDayHKCandle.PreviousCandlePayload.CandleColor = Color.Red AndAlso
                            lastTrade.SupportingTradingSymbol.EndsWith("PE") AndAlso
                            lastTrade.EntryTime.Date = currentDayHKCandle.PreviousCandlePayload.PayloadDate) OrElse
                            (currentDayHKCandle.PreviousCandlePayload.CandleColor = Color.Green AndAlso
                            lastTrade.SupportingTradingSymbol.EndsWith("CE") AndAlso
                            lastTrade.EntryTime.Date = currentDayHKCandle.PreviousCandlePayload.PayloadDate)) Then
                            If currentDayHKCandle.PreviousCandlePayload.CandleColor = Color.Green Then
                                If _smiPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) <= -40 Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, currentMinuteCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                                End If
                            ElseIf currentDayHKCandle.PreviousCandlePayload.CandleColor = Color.Red Then
                                If _smiPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) >= 40 Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, currentMinuteCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                                End If
                            End If
                        Else
                            If currentDayHKCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish AndAlso
                                currentDayHKCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Red Then
                                If _smiPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) <= -40 Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, currentMinuteCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                                End If
                            ElseIf currentDayHKCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish AndAlso
                                currentDayHKCandle.PreviousCandlePayload.PreviousCandlePayload.CandleColor = Color.Green Then
                                If _smiPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) >= 40 Then
                                    ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, currentMinuteCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                                End If
                            End If
                        End If
                    End If
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
                        ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, currentOptTick.Open, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                    Else
                        ret += _ParentStrategy.CalculatePL(_TradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
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
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = IsEntrySignalReceived(currentTickTime)
        If ret IsNot Nothing AndAlso ret.Item1 Then
            If _ParentStrategy.GetNumberOfActiveStocks(_TradingDate, Trade.TypeOfTrade.CNC) < _userInputs.NumberOfActiveStock Then
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

            'EOD Exit
            If currentTickTime.Hour >= 15 AndAlso currentTickTime.Minute >= 25 Then
                For Each runningTrade In availableTrades
                    If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        Dim currentOptTick As Payload = GetCurrentTick(runningTrade.SupportingTradingSymbol, currentTickTime)
                        _ParentStrategy.ExitTradeByForce(runningTrade, currentOptTick, "EOD Exit")
                    End If
                Next
            End If
        End If
    End Function

    Private Function GetCurrentTick(ByVal tradingSymbol As String, ByVal currentTime As Date) As Payload
        Dim ret As Payload = Nothing
        If _dependentInstruments Is Nothing OrElse Not _dependentInstruments.ContainsKey(tradingSymbol) Then
            Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
            inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol, _TradingDate.AddDays(-50), _TradingDate)
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
        Dim spotATR As Decimal = _atrPayload(signalCandle.PayloadDate.Date)
        Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol, _TradingDate, Trade.TypeOfTrade.CNC)
        If lastTrade IsNot Nothing AndAlso lastTrade.ExitRemark IsNot Nothing AndAlso lastTrade.ExitRemark.ToUpper <> "TARGET HIT" Then
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

        ret = EnterBuyTrade(signalCandle, spotTick, spotATR, currentTickTime, childTag, parentTag, tradeNumber, entryType, direction, lossToRecover)

        Return ret
    End Function

    Private Function EnterBuyTrade(ByVal signalCandle As Payload, ByVal currentSpotTick As Payload, ByVal currentSpotATR As Decimal,
                                   ByVal currentTickTime As Date, ByVal childTag As String, ByVal parentTag As String,
                                   ByVal tradeNumber As Integer, ByVal entryType As Trade.TypeOfEntry,
                                   ByVal diretion As Trade.TradeExecutionDirection, ByVal previousLoss As Decimal) As Boolean
        Dim ret As Boolean = False
        Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime)
        Dim optionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
        Dim currentOptionTradingSymbol As String = Nothing
        If diretion = Trade.TradeExecutionDirection.Buy Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "CE", currentSpotATR)
        ElseIf diretion = Trade.TradeExecutionDirection.Sell Then
            currentOptionTradingSymbol = GetCurrentATMOption(currentSpotTick.PayloadDate, optionExpiryString, currentSpotTick.Open, "PE", currentSpotATR)
        End If

        If currentOptionTradingSymbol IsNot Nothing Then
            Dim currentOptTick As Payload = GetCurrentTick(currentOptionTradingSymbol, currentSpotTick.PayloadDate)
            If currentOptTick IsNot Nothing AndAlso currentOptTick.PayloadDate >= currentMinute AndAlso currentOptTick.Volume > 0 Then
                Dim quantity As Integer = _LotSize
                If entryType = Trade.TypeOfEntry.LossMakeup Then
                    Dim entryPrice As Decimal = currentOptTick.Open
                    Dim targetPrice As Decimal = ConvertFloorCeling(entryPrice + currentSpotATR / _userInputs.SpotToOptionDelta, _ParentStrategy.TickSize, RoundOfType.Celing)
                    For ctr As Integer = 1 To Integer.MaxValue
                        Dim pl As Decimal = _ParentStrategy.CalculatePL(_TradingSymbol, entryPrice, targetPrice, ctr * _LotSize, _LotSize, Trade.TypeOfStock.Options)
                        If pl >= Math.Abs(previousLoss) Then
                            If _userInputs.ExitAtATRPL Then previousLoss = (pl - 1) * -1
                            quantity = ctr * _LotSize + _LotSize
                            Exit For
                        End If
                    Next
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

    'Private Function EnterDuplicateTrade(ByVal existingTrade As Trade, ByVal currentTick As Payload, ByVal increaseQuantityIfRequired As Boolean) As Boolean
    '    Dim ret As Boolean = False
    '    Try
    '        Dim quantity As Integer = existingTrade.Quantity
    '        Dim remark As String = Nothing
    '        'If increaseQuantityIfRequired AndAlso _userInputs.IncreaseQuantityWithHalfPremium Then
    '        '    Dim increasedCapital As Decimal = currentTick.Open * quantity * 2
    '        '    If increasedCapital <= existingTrade.CapitalRequiredWithMargin * 120 / 100 Then
    '        '        quantity = quantity * 2
    '        '    Else
    '        '        remark = "Unable to increase quantity"
    '        '    End If
    '        'End If
    '        Dim runningTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
    '                                                tradingSymbol:=existingTrade.TradingSymbol,
    '                                                stockType:=existingTrade.StockType,
    '                                                orderType:=Trade.TypeOfOrder.Market,
    '                                                tradingDate:=currentTick.PayloadDate,
    '                                                entryDirection:=existingTrade.EntryDirection,
    '                                                entryPrice:=currentTick.Open,
    '                                                entryBuffer:=0,
    '                                                squareOffType:=Trade.TypeOfTrade.CNC,
    '                                                entryCondition:=Trade.TradeEntryCondition.Original,
    '                                                entryRemark:="Original Entry",
    '                                                quantity:=quantity,
    '                                                lotSize:=existingTrade.LotSize,
    '                                                potentialTarget:=currentTick.Open + 100000,
    '                                                targetRemark:=100000,
    '                                                potentialStopLoss:=currentTick.Open - 100000,
    '                                                stoplossBuffer:=0,
    '                                                slRemark:=100000,
    '                                                signalCandle:=existingTrade.SignalCandle)

    '        runningTrade.UpdateTrade(ChildTag:=existingTrade.ChildTag,
    '                                ParentTag:=existingTrade.ParentTag,
    '                                TradeNumber:=existingTrade.TradeNumber,
    '                                EntryType:=existingTrade.EntryType,
    '                                SpotPrice:=existingTrade.SpotPrice,
    '                                SpotATR:=existingTrade.SpotATR,
    '                                PreviousLoss:=existingTrade.PreviousLoss,
    '                                SupportingTradingSymbol:=currentTick.TradingSymbol)

    '        If remark IsNot Nothing AndAlso remark.Trim <> "" Then
    '            runningTrade.UpdateTrade(Remark2:=remark)
    '        End If

    '        If _ParentStrategy.PlaceOrModifyOrder(runningTrade, Nothing) Then
    '            ret = _ParentStrategy.EnterTradeIfPossible(runningTrade.TradingSymbol, _TradingDate, runningTrade, currentTick)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    '    Return ret
    'End Function
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
        query = "SELECT `TRADING_SYMBOL`
                FROM `active_instruments_futures`
                WHERE `TRADING_SYMBOL` LIKE '{0}%{1}'
                AND `AS_ON_DATE`=(SELECT MAX(`AS_ON_DATE`)
                FROM `active_instruments_futures`
                WHERE `AS_ON_DATE`<='{2}')"
        query = String.Format(query, expiryString, "PE", currentTickTime.ToString("yyyy-MM-dd"))

        Dim dt As DataTable = _ParentStrategy.Cmn.RunSelect(query)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim contracts As Dictionary(Of Decimal, Long) = New Dictionary(Of Decimal, Long)
            For Each runningRow As DataRow In dt.Rows
                Dim tradingSymbol As String = runningRow.Item("TRADING_SYMBOL")
                Dim strike As String = Utilities.Strings.GetTextBetween(expiryString, "PE", tradingSymbol)
                If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike.Trim) Then
                    contracts.Add(Val(strike.Trim), 1)
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
        query = "SELECT `TRADING_SYMBOL`
                FROM `active_instruments_futures`
                WHERE `TRADING_SYMBOL` LIKE '{0}%{1}'
                AND `AS_ON_DATE`=(SELECT MAX(`AS_ON_DATE`)
                FROM `active_instruments_futures`
                WHERE `AS_ON_DATE`<='{2}')"
        query = String.Format(query, expiryString, "CE", currentTickTime.ToString("yyyy-MM-dd"))

        Dim dt As DataTable = _ParentStrategy.Cmn.RunSelect(query)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim contracts As Dictionary(Of Decimal, Long) = New Dictionary(Of Decimal, Long)
            For Each runningRow As DataRow In dt.Rows
                Dim tradingSymbol As String = runningRow.Item("TRADING_SYMBOL")
                Dim strike As String = Utilities.Strings.GetTextBetween(expiryString, "CE", tradingSymbol)
                If strike IsNot Nothing AndAlso strike.Trim <> "" AndAlso IsNumeric(strike.Trim) Then
                    contracts.Add(Val(strike.Trim), 1)
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