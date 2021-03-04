Imports Algo2TradeBLL
Imports System.Threading
Imports Utilities.Numbers

Namespace StrategyHelper
    Public MustInherit Class StrategyRule

#Region "Events/Event handlers"
        Public Event DocumentDownloadComplete()
        Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        Public Event Heartbeat(ByVal msg As String)
        Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        'The below functions are needed to allow the derived classes to raise the above two events
        Protected Overridable Sub OnDocumentDownloadComplete()
            RaiseEvent DocumentDownloadComplete()
        End Sub
        Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
            RaiseEvent DocumentRetryStatus(currentTry, totalTries)
        End Sub
        Protected Overridable Sub OnHeartbeat(ByVal msg As String)
            RaiseEvent Heartbeat(msg)
        End Sub
        Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
            RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
        End Sub
#End Region

        Protected ReadOnly _Cts As CancellationTokenSource
        Protected ReadOnly _TradingDate As Date
        Protected ReadOnly _NextTradingDay As Date
        Protected ReadOnly _TradingSymbol As String
        Protected ReadOnly _LotSize As Integer
        Protected ReadOnly _ParentStrategy As Strategy
        Protected ReadOnly _InputMinPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _InputEODPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _TradeStartTime As Date

        Protected _SignalPayload As Dictionary(Of Date, Payload) = Nothing
        Protected _DependentInstrumentsPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
        Protected _DependentInstrumentsExpiryDate As Dictionary(Of String, Date) = Nothing

        Private _NSEHolidays As List(Of Date) = Nothing

        Public Sub New(ByVal canceller As CancellationTokenSource,
                       ByVal tradingDate As Date,
                       ByVal nextTradingDay As Date,
                       ByVal tradingSymbol As String,
                       ByVal lotSize As Integer,
                       ByVal parentStrategy As Strategy,
                       ByVal inputMinPayload As Dictionary(Of Date, Payload),
                       ByVal inputEODPayload As Dictionary(Of Date, Payload))
            _Cts = canceller
            _TradingDate = tradingDate.Date
            _NextTradingDay = nextTradingDay.Date
            _TradingSymbol = tradingSymbol
            _LotSize = lotSize
            _ParentStrategy = parentStrategy
            _InputMinPayload = inputMinPayload
            _InputEODPayload = inputEODPayload
            _TradeStartTime = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day, _ParentStrategy.TradeStartTime.Hours, _ParentStrategy.TradeStartTime.Minutes, _ParentStrategy.TradeStartTime.Seconds)

            _NSEHolidays = New List(Of Date) From {
                New Date(2021, 3, 11),
                New Date(2021, 5, 13),
                New Date(2021, 8, 19),
                New Date(2021, 11, 4)
            }
        End Sub

#Region "Public Functions"
        Public Overridable Sub CompletePreProcessing()
            If _ParentStrategy IsNot Nothing AndAlso _InputMinPayload IsNot Nothing AndAlso _InputMinPayload.Count > 0 AndAlso
                _InputEODPayload IsNot Nothing AndAlso _InputEODPayload.Count > 0 Then
                If _ParentStrategy.SignalTimeFrame >= 375 Then
                    _SignalPayload = _InputEODPayload
                ElseIf _ParentStrategy.SignalTimeFrame > 1 Then
                    Dim exchangeStartTime As Date = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day,
                                                             _ParentStrategy.ExchangeStartTime.Hours, _ParentStrategy.ExchangeStartTime.Minutes, _ParentStrategy.ExchangeStartTime.Seconds)
                    _SignalPayload = Common.ConvertPayloadsToXMinutes(_InputMinPayload, _ParentStrategy.SignalTimeFrame, exchangeStartTime)
                Else
                    _SignalPayload = _InputMinPayload
                End If

                _DependentInstrumentsPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                _DependentInstrumentsPayload.Add(_TradingSymbol.Trim.ToUpper, _InputMinPayload)
            End If
        End Sub

        Public Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task(Of List(Of Tuple(Of Trade, Payload)))
            Await Task.Delay(0).ConfigureAwait(False)
            Dim ret As List(Of Tuple(Of Trade, Payload)) = Nothing
            If Not _ParentStrategy.IsLogicalTradeActiveOfTheStock(_TradingSymbol) Then
                Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, 1)
                Dim signal As Tuple(Of Boolean, Payload, Trade.TradeDirection) = IsEntrySignalReceived(currentTickTime)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim lastCompleteTrade As Trade = _ParentStrategy.GetLastCompleteTradeOfTheStock(_TradingSymbol)

                    Dim childTag As String = System.Guid.NewGuid.ToString()
                    Dim parentTag As String = childTag
                    Dim iterationNumber As Integer = 1
                    Dim entryType As Trade.TypeOfEntry = Trade.TypeOfEntry.Fresh
                    Dim lossToRecover As Decimal = 0
                    Dim quantity As Integer = _LotSize
                    If lastCompleteTrade IsNot Nothing Then
                        If lastCompleteTrade.ExitType = Trade.TypeOfExit.Stoploss Then
                            entryType = Trade.TypeOfEntry.Stoploss
                            parentTag = lastCompleteTrade.ParentReference
                            iterationNumber = lastCompleteTrade.IterationNumber + 1
                            lossToRecover = GetParentTradesOverallPL(lastCompleteTrade.ParentReference, currentTickTime)
                            quantity = _LotSize * Math.Pow(2, iterationNumber - 1)
                        ElseIf lastCompleteTrade.ExitType = Trade.TypeOfExit.ContractRollover Then
                            entryType = Trade.TypeOfEntry.Rollover
                            parentTag = lastCompleteTrade.ParentReference
                            iterationNumber = lastCompleteTrade.IterationNumber + 1
                            lossToRecover = GetParentTradesOverallPL(lastCompleteTrade.ParentReference, currentTickTime)
                            quantity = _LotSize * Math.Pow(2, iterationNumber - 1)
                        End If
                    End If

                    Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)

                    Dim optionExpiryString As String = Nothing
                    If _TradingDate.DayOfWeek = DayOfWeek.Friday OrElse _TradingDate.DayOfWeek = DayOfWeek.Monday Then
                        optionExpiryString = GetCurrentOptionInstrumentExpiryString(_TradingSymbol, _TradingDate)
                    Else
                        optionExpiryString = GetNearOptionInstrumentExpiryString(_TradingSymbol, _TradingDate)
                    End If
                    Dim ceTradingSymbol As String = GetATMCEOption(currentTickTime, optionExpiryString, currentSpotTick.Open)
                    Dim peTradingSymbol As String = GetATMPEOption(currentTickTime, optionExpiryString, currentSpotTick.Open)
                    If ceTradingSymbol IsNot Nothing AndAlso peTradingSymbol IsNot Nothing Then
                        Dim ceOptTick As Payload = GetCurrentTick(ceTradingSymbol, currentSpotTick.PayloadDate)
                        Dim peOptTick As Payload = GetCurrentTick(peTradingSymbol, currentSpotTick.PayloadDate)
                        If ceOptTick IsNot Nothing AndAlso peOptTick IsNot Nothing AndAlso
                            ceOptTick.Volume > 0 AndAlso peOptTick.Volume > 0 AndAlso
                            ceOptTick.PayloadDate >= currentMinute AndAlso peOptTick.PayloadDate >= currentMinute Then
                            Dim ceEntryPrice As Decimal = ceOptTick.High
                            Dim peEntryPrice As Decimal = peOptTick.High
                            Dim requiredCapital As Decimal = ceEntryPrice * _LotSize + peEntryPrice * _LotSize
                            Dim potentialTarget As Decimal = requiredCapital * _ParentStrategy.RuleSettings.TargetCapitalPercentage / 100
                            Dim originalTarget As Decimal = potentialTarget
                            If iterationNumber > 1 Then
                                originalTarget = lastCompleteTrade.OriginalTarget
                                potentialTarget = originalTarget + originalTarget * (iterationNumber - 1) * 30 / 100
                                If lossToRecover < 0 Then potentialTarget = potentialTarget + Math.Abs(lossToRecover)
                            End If

                            Dim ceTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                            tradingSymbol:=ceTradingSymbol,
                                                            spotTradingSymbol:=_TradingSymbol,
                                                            stockType:=Trade.TypeOfStock.Options,
                                                            tradingDate:=currentTickTime.Date,
                                                            signalDirection:=Trade.TradeDirection.Buy,
                                                            entryDirection:=Trade.TradeDirection.Buy,
                                                            entryType:=entryType,
                                                            entryPrice:=ceEntryPrice,
                                                            quantity:=quantity,
                                                            lotSize:=_LotSize,
                                                            entrySignalCandle:=signal.Item2,
                                                            childReference:=childTag,
                                                            parentReference:=parentTag,
                                                            iterationNumber:=iterationNumber,
                                                            spotPrice:=currentSpotTick.Open,
                                                            requiredCapital:=ceEntryPrice * quantity + peEntryPrice * quantity,
                                                            previousLoss:=lossToRecover,
                                                            potentialTarget:=potentialTarget,
                                                            originalTarget:=originalTarget)

                            If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload))
                            ret.Add(New Tuple(Of Trade, Payload)(ceTrade, ceOptTick))

                            Dim peTrade As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                            tradingSymbol:=peTradingSymbol,
                                                            spotTradingSymbol:=_TradingSymbol,
                                                            stockType:=Trade.TypeOfStock.Options,
                                                            tradingDate:=currentTickTime.Date,
                                                            signalDirection:=Trade.TradeDirection.Buy,
                                                            entryDirection:=Trade.TradeDirection.Buy,
                                                            entryType:=entryType,
                                                            entryPrice:=peEntryPrice,
                                                            quantity:=quantity,
                                                            lotSize:=_LotSize,
                                                            entrySignalCandle:=signal.Item2,
                                                            childReference:=childTag,
                                                            parentReference:=parentTag,
                                                            iterationNumber:=iterationNumber,
                                                            spotPrice:=currentSpotTick.Open,
                                                            requiredCapital:=ceEntryPrice * quantity + peEntryPrice * quantity,
                                                            previousLoss:=lossToRecover,
                                                            potentialTarget:=potentialTarget,
                                                            originalTarget:=originalTarget)

                            If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload))
                            ret.Add(New Tuple(Of Trade, Payload)(peTrade, peOptTick))
                        End If
                    End If
                End If
            End If
            Return ret
        End Function

        Public Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTickTime As Date, ByVal availableTrades As List(Of Trade)) As Task(Of List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)))
            Await Task.Delay(0).ConfigureAwait(False)
            Dim ret As List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)) = Nothing
            If availableTrades IsNot Nothing AndAlso availableTrades.Count > 0 Then
                For Each runningTrade In availableTrades
                    Dim currentTick As Payload = GetCurrentTick(runningTrade.TradingSymbol, currentTickTime)
                    If currentTick IsNot Nothing Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Open Then
                            If runningTrade.EntryTime.Minute <> currentTickTime.Minute Then
                                If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.None, Nothing))
                            End If
                        ElseIf runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                            'Set Drawup drawdown
                            If runningTrade.EntryDirection = Trade.TradeDirection.Buy Then
                                If currentTick.Open > runningTrade.MaxDrawUp Then runningTrade.UpdateTrade(maxDrawUp:=currentTick.Open)
                                If currentTick.Open < runningTrade.MaxDrawDown Then runningTrade.UpdateTrade(maxDrawDown:=currentTick.Open)
                            ElseIf runningTrade.EntryDirection = Trade.TradeDirection.Sell Then
                                If currentTick.Open < runningTrade.MaxDrawUp Then runningTrade.UpdateTrade(maxDrawUp:=currentTick.Open)
                                If currentTick.Open > runningTrade.MaxDrawDown Then runningTrade.UpdateTrade(maxDrawDown:=currentTick.Open)
                            End If

                            'Exit Check
                            Dim currentPL As Decimal = GetChildTradesOverallPL(runningTrade.ChildReference, currentTickTime)
                            If currentPL >= runningTrade.PotentialTarget Then
                                If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.Target, Nothing))
                            ElseIf currentPL <= (runningTrade.RequiredCapital / 2) * -1 Then
                                If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.Stoploss, Nothing))
                            Else
                                Dim expiryDate As Date = GetExpiryDate(runningTrade.TradingSymbol)
                                expiryDate = New Date(expiryDate.Year, expiryDate.Month, expiryDate.Day, 15, 29, 0)
                                If currentTickTime >= expiryDate Then
                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                    ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.ContractRollover, Nothing))
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            Return ret
        End Function

        Public MustOverride Async Function UpdateCollectionsIfRequiredAsync(ByVal currentTickTime As Date) As Task
#End Region

#Region "Private Functions"
        Private Function GetCurrentTick(ByVal tradingSymbol As String, ByVal currentTime As Date) As Payload
            Dim ret As Payload = Nothing
            If Not _DependentInstrumentsPayload.ContainsKey(tradingSymbol.Trim.ToUpper) Then
                Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
                inputPayload = _ParentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol.Trim.ToUpper, _TradingDate.AddDays(-50), _TradingDate)
                If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                    _DependentInstrumentsPayload.Add(tradingSymbol.Trim.ToUpper, inputPayload)
                End If
            End If
            If _DependentInstrumentsPayload IsNot Nothing AndAlso _DependentInstrumentsPayload.ContainsKey(tradingSymbol.Trim.ToUpper) Then
                Dim inputPayload As Dictionary(Of Date, Payload) = _DependentInstrumentsPayload(tradingSymbol.Trim.ToUpper)
                If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                    Dim currentMinute As Date = New Date(currentTime.Year, currentTime.Month, currentTime.Day, currentTime.Hour, currentTime.Minute, 0)
                    Dim currentCandle As Payload = inputPayload.Where(Function(x)
                                                                          Return x.Key <= currentMinute
                                                                      End Function).LastOrDefault.Value

                    ret = currentCandle

                    'If currentCandle Is Nothing Then
                    '    currentCandle = inputPayload.Where(Function(x)
                    '                                           Return x.Key.Date = _TradingDate
                    '                                       End Function).FirstOrDefault.Value
                    'End If

                    'ret = currentCandle.Ticks.FindAll(Function(x)
                    '                                      Return x.PayloadDate <= currentTime
                    '                                  End Function).LastOrDefault

                    'If ret Is Nothing Then ret = currentCandle.Ticks.FirstOrDefault
                End If
            End If
            Return ret
        End Function

        Private Function GetExpiryDate(ByVal tradingSymbol As String) As Date
            Dim ret As Date = Date.MinValue
            If _DependentInstrumentsExpiryDate IsNot Nothing AndAlso _DependentInstrumentsExpiryDate.ContainsKey(tradingSymbol.ToUpper) Then
                ret = _DependentInstrumentsExpiryDate(tradingSymbol.ToUpper)
            Else
                Dim expiryDate As Date = _ParentStrategy.Cmn.GetExpiryDate(Common.DataBaseTable.Intraday_Futures_Options, tradingSymbol)
                If expiryDate <> Date.MinValue Then
                    If _DependentInstrumentsExpiryDate Is Nothing Then _DependentInstrumentsExpiryDate = New Dictionary(Of String, Date)
                    _DependentInstrumentsExpiryDate.Add(tradingSymbol.ToUpper, expiryDate)
                    ret = expiryDate
                End If
            End If
            Return ret
        End Function

        Private Function GetChildTradesOverallPL(ByVal tag As String, ByVal currentTickTime As Date) As Decimal
            Dim ret As Decimal = 0
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByChildTag(tag, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                            Dim currentTick As Payload = GetCurrentTick(runningTrade.TradingSymbol, currentTickTime)
                            If runningTrade.EntryDirection = Trade.TradeDirection.Buy Then
                                ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, runningTrade.EntryPrice, currentTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                            ElseIf runningTrade.EntryDirection = Trade.TradeDirection.Sell Then
                                ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, currentTick.Open, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                            End If
                        Else
                            ret += runningTrade.PLAfterBrokerage
                        End If
                    End If
                Next
            End If
            Return ret
        End Function

        Private Function GetParentTradesOverallPL(ByVal tag As String, ByVal currentTickTime As Date) As Decimal
            Dim ret As Decimal = 0
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByParentTag(tag, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                            Dim currentTick As Payload = GetCurrentTick(runningTrade.TradingSymbol, currentTickTime)
                            If runningTrade.EntryDirection = Trade.TradeDirection.Buy Then
                                ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, runningTrade.EntryPrice, currentTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                            ElseIf runningTrade.EntryDirection = Trade.TradeDirection.Sell Then
                                ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, currentTick.Open, runningTrade.EntryPrice, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
                            End If
                        Else
                            ret += runningTrade.PLAfterBrokerage
                        End If
                    End If
                Next
            End If
            Return ret
        End Function
#End Region

#Region "Protected Functions"
        Protected MustOverride Function IsEntrySignalReceived(ByVal currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)

        Protected Function GetCurrentOptionInstrumentExpiryString(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
            Dim ret As String = Nothing
            Dim expiryDate As Date = Date.MinValue
            If coreInstrumentName = "NIFTY BANK" OrElse coreInstrumentName = "NIFTY 50" Then
                expiryDate = GetNextThrusday(tradingDate)
            Else
                expiryDate = GetLastThusrdayOfMonth(tradingDate)
            End If
            If coreInstrumentName = "NIFTY BANK" Then coreInstrumentName = "BANKNIFTY"
            If coreInstrumentName = "NIFTY 50" Then coreInstrumentName = "NIFTY"

            While True
                If _ParentStrategy.Cmn.IsTradingDay(expiryDate.Date) Then
                    Exit While
                ElseIf Not _NSEHolidays.Contains(expiryDate.Date) Then
                    Exit While
                Else
                    expiryDate = expiryDate.AddDays(-1)
                End If
            End While

            Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
            While True
                If _ParentStrategy.Cmn.IsTradingDay(lastThursday.Date) Then
                    Exit While
                ElseIf Not _NSEHolidays.Contains(lastThursday.Date) Then
                    Exit While
                Else
                    lastThursday = lastThursday.AddDays(-1)
                End If
            End While

            If expiryDate.Date = lastThursday.Date Then
                ret = String.Format("{0}{1}", coreInstrumentName, expiryDate.ToString("yyMMM")).ToUpper
            Else
                If expiryDate.Month >= 10 Then
                    ret = String.Format("{0}{1}{2}{3}", coreInstrumentName, expiryDate.ToString("yy"), expiryDate.ToString("MMM").Substring(0, 1), expiryDate.ToString("dd")).ToUpper
                Else
                    ret = String.Format("{0}{1}", coreInstrumentName, expiryDate.ToString("yyMdd")).ToUpper
                End If
            End If
            Return ret
        End Function

        Protected Function GetNearOptionInstrumentExpiryString(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
            Dim ret As String = Nothing
            Dim expiryDate As Date = Date.MinValue
            If coreInstrumentName = "NIFTY BANK" OrElse coreInstrumentName = "NIFTY 50" Then
                expiryDate = GetNextThrusday(GetNextThrusday(tradingDate).AddDays(1))
            Else
                expiryDate = GetLastThusrdayOfMonth(GetLastThusrdayOfMonth(tradingDate).AddDays(10))
            End If
            If coreInstrumentName = "NIFTY BANK" Then coreInstrumentName = "BANKNIFTY"
            If coreInstrumentName = "NIFTY 50" Then coreInstrumentName = "NIFTY"

            While True
                If _ParentStrategy.Cmn.IsTradingDay(expiryDate.Date) Then
                    Exit While
                ElseIf Not _NSEHolidays.Contains(expiryDate.Date) Then
                    Exit While
                Else
                    expiryDate = expiryDate.AddDays(-1)
                End If
            End While

            Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
            While True
                If _ParentStrategy.Cmn.IsTradingDay(lastThursday.Date) Then
                    Exit While
                ElseIf Not _NSEHolidays.Contains(lastThursday.Date) Then
                    Exit While
                Else
                    lastThursday = lastThursday.AddDays(-1)
                End If
            End While

            Dim nextLastThursday As Date = GetLastThusrdayOfMonth(lastThursday.AddDays(10))
            While True
                If _ParentStrategy.Cmn.IsTradingDay(nextLastThursday.Date) Then
                    Exit While
                ElseIf Not _NSEHolidays.Contains(nextLastThursday.Date) Then
                    Exit While
                Else
                    nextLastThursday = lastThursday.AddDays(-1)
                End If
            End While

            If expiryDate.Date = lastThursday.Date OrElse expiryDate.Date = nextLastThursday.Date Then
                ret = String.Format("{0}{1}", coreInstrumentName, expiryDate.ToString("yyMMM")).ToUpper
            Else
                If expiryDate.Month >= 10 Then
                    ret = String.Format("{0}{1}{2}{3}", coreInstrumentName, expiryDate.ToString("yy"), expiryDate.ToString("MMM").Substring(0, 1), expiryDate.ToString("dd")).ToUpper
                Else
                    ret = String.Format("{0}{1}", coreInstrumentName, expiryDate.ToString("yyMdd")).ToUpper
                End If
            End If
            Return ret
        End Function
#End Region

#Region "Contract Helper"
        Private Function GetATMPEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
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
            Return ret
        End Function

        Private Function GetATMCEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
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

        Private Function GetNextThrusday(ByVal dateTime As Date) As Date
            Dim ret As Date = Date.MinValue
            Dim daysUntilTradingDay As Integer = (CInt(DayOfWeek.Thursday) - CInt(dateTime.DayOfWeek) + 7) Mod 7
            ret = dateTime.Date.AddDays(daysUntilTradingDay).Date
            Return ret
        End Function
#End Region
    End Class
End Namespace