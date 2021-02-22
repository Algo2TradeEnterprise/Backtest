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
        Protected _ATRPayload As Dictionary(Of Date, Decimal) = Nothing

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

                Indicator.ATR.CalculateATR(14, _InputEODPayload, _ATRPayload)
            End If
        End Sub

        Public Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task(Of List(Of Tuple(Of Trade, Payload)))
            Await Task.Delay(0).ConfigureAwait(False)
            Dim ret As List(Of Tuple(Of Trade, Payload)) = Nothing
            Dim lastTrade As Trade = _ParentStrategy.GetLastEntryTradeOfTheStock(_TradingSymbol)
            If Not _ParentStrategy.IsLogicalTradeActiveOfTheStock(_TradingSymbol) Then
                Dim signal As Tuple(Of Boolean, Payload, Trade.TradeDirection) = Nothing
                Dim currentMinute As Date = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, 1)
                If _ParentStrategy.SignalTimeFrame >= 375 Then
                    If currentTickTime >= _TradeStartTime Then
                        signal = IsFreshEntrySignalReceived(currentTickTime)
                        If signal IsNot Nothing AndAlso signal.Item1 Then
                            If _ParentStrategy.RuleSettings.TypeOfSignal = RuleEntities.SignalType.DifferentSignal Then
                                If lastTrade IsNot Nothing AndAlso lastTrade.ExitType = Trade.TypeOfExit.Target AndAlso
                                    lastTrade.EntrySignalCandle.PayloadDate = signal.Item2.PayloadDate Then
                                    signal = Nothing
                                Else
                                    If signal.Item2.PayloadDate.Date <> _TradingDate AndAlso currentTickTime < _TradeStartTime.AddMinutes(1) Then
                                        signal = Nothing
                                    End If
                                End If
                            Else
                                If signal.Item2.PayloadDate.Date <> _TradingDate AndAlso currentTickTime < _TradeStartTime.AddMinutes(1) Then
                                    signal = Nothing
                                End If
                            End If
                        End If
                    End If
                Else
                    currentMinute = _ParentStrategy.GetCurrentXMinuteCandleTime(currentTickTime, _ParentStrategy.SignalTimeFrame)
                    If (currentTickTime >= currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 2) OrElse currentTickTime >= _TradeStartTime) Then
                        signal = IsFreshEntrySignalReceived(currentTickTime)
                        If signal IsNot Nothing AndAlso signal.Item1 Then
                            If _ParentStrategy.RuleSettings.TypeOfSignal = RuleEntities.SignalType.DifferentSignal Then
                                If lastTrade IsNot Nothing AndAlso lastTrade.ExitType = Trade.TypeOfExit.Target AndAlso
                                    lastTrade.EntrySignalCandle.PayloadDate = signal.Item2.PayloadDate Then
                                    signal = Nothing
                                Else
                                    If signal.Item2.PayloadDate <> currentMinute AndAlso
                                        (currentTickTime < currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 1) OrElse currentTickTime < _TradeStartTime.AddMinutes(1)) Then
                                        signal = Nothing
                                    End If
                                End If
                            Else
                                If signal.Item2.PayloadDate <> currentMinute AndAlso
                                    (currentTickTime < currentMinute.AddMinutes(_ParentStrategy.SignalTimeFrame - 1) OrElse currentTickTime < _TradeStartTime.AddMinutes(1)) Then
                                    signal = Nothing
                                End If
                            End If
                        End If
                    End If
                End If

                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim targetReached As Boolean = True
                    Dim targetLeftPercentage As Decimal = 0
                    If signal.Item3 = Trade.TradeDirection.Buy Then
                        Dim highestHigh As Decimal = _SignalPayload.Max(Function(x)
                                                                            If x.Key > signal.Item2.PayloadDate AndAlso x.Key < _TradingDate Then
                                                                                Return x.Value.High
                                                                            Else
                                                                                Return Decimal.MinValue
                                                                            End If
                                                                        End Function)
                        If (_ParentStrategy.SignalTimeFrame >= 375 AndAlso signal.Item2.PayloadDate.Date <> _TradingDate.Date) OrElse
                            (_ParentStrategy.SignalTimeFrame < 375 AndAlso signal.Item2.PayloadDate <> currentMinute) Then
                            Dim todayHigh As Decimal = _InputMinPayload.Max(Function(x)
                                                                                If x.Key.Date = _TradingDate.Date AndAlso x.Key < currentTickTime Then
                                                                                    Return x.Value.High
                                                                                Else
                                                                                    Return Decimal.MinValue
                                                                                End If
                                                                            End Function)

                            highestHigh = Math.Max(highestHigh, todayHigh)
                        End If

                        Dim atr As Decimal = _ATRPayload(signal.Item2.PayloadDate.Date)
                        If highestHigh < signal.Item2.Close + atr Then
                            targetReached = False
                            If highestHigh <> Decimal.MinValue Then
                                targetLeftPercentage = ((atr - (highestHigh - signal.Item2.Close)) / atr) * 100
                            Else
                                targetLeftPercentage = 100
                            End If
                        End If
                    ElseIf signal.Item3 = Trade.TradeDirection.Sell Then
                        Dim lowestLow As Decimal = _SignalPayload.Min(Function(x)
                                                                          If x.Key > signal.Item2.PayloadDate AndAlso x.Key < _TradingDate Then
                                                                              Return x.Value.Low
                                                                          Else
                                                                              Return Decimal.MaxValue
                                                                          End If
                                                                      End Function)
                        If (_ParentStrategy.SignalTimeFrame >= 375 AndAlso signal.Item2.PayloadDate.Date <> _TradingDate.Date) OrElse
                            (_ParentStrategy.SignalTimeFrame < 375 AndAlso signal.Item2.PayloadDate <> currentMinute) Then
                            Dim todayLow As Decimal = _InputMinPayload.Min(Function(x)
                                                                               If x.Key.Date = _TradingDate.Date AndAlso x.Key < currentTickTime Then
                                                                                   Return x.Value.Low
                                                                               Else
                                                                                   Return Decimal.MaxValue
                                                                               End If
                                                                           End Function)

                            lowestLow = Math.Min(lowestLow, todayLow)
                        End If

                        Dim atr As Decimal = _ATRPayload(signal.Item2.PayloadDate.Date)
                        If lowestLow > signal.Item2.Close - atr Then
                            targetReached = False
                            If lowestLow <> Decimal.MaxValue Then
                                targetLeftPercentage = ((atr - (signal.Item2.Close - lowestLow)) / atr) * 100
                            Else
                                targetLeftPercentage = 100
                            End If
                        End If
                    End If
                    If Not targetReached AndAlso targetLeftPercentage >= 75 Then
                        Dim childTag As String = System.Guid.NewGuid.ToString()
                        Dim parentTag As String = childTag
                        Dim iterationNumber As Integer = 1
                        Dim entryType As Trade.TypeOfEntry = Trade.TypeOfEntry.Fresh
                        Dim lossToRecover As Decimal = 0
                        Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                        Dim spotATR As Decimal = _ATRPayload(signal.Item2.PayloadDate.Date)
                        Dim directionToCheck As Trade.TradeDirection = signal.Item3

                        Dim optionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
                        Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, optionExpiryString, currentSpotTick.Open, directionToCheck)
                        If optionTradingSymbol IsNot Nothing Then
                            Dim currentOptTick As Payload = GetCurrentTick(optionTradingSymbol, currentSpotTick.PayloadDate)
                            If currentOptTick IsNot Nothing Then
                                Dim quantity As Integer = _LotSize
                                Dim entryPrice As Decimal = currentOptTick.Open
                                Dim targetPrice As Decimal = ConvertFloorCeling(entryPrice + spotATR / 2, _ParentStrategy.TickSize, RoundOfType.Celing)
                                Dim potentialTarget As Decimal = _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, entryPrice, targetPrice, quantity, _LotSize, Trade.TypeOfStock.Options) - 1
                                If _ParentStrategy.RuleSettings.TypeOfTarget = RuleEntities.TargetType.CapitalPercentage Then
                                    Dim capital As Decimal = currentOptTick.Open * quantity / _ParentStrategy.MarginMultiplier
                                    If lossToRecover < 0 Then capital = capital + Math.Abs(lossToRecover)
                                    potentialTarget = capital * _ParentStrategy.RuleSettings.CapitalPercentage / 100

                                    If potentialTarget < 1000 AndAlso
                                        _ParentStrategy.RuleSettings.TypeOfQuantity = RuleEntities.QuantityType.Increase Then
                                        Dim multiplier As Integer = Math.Ceiling(1000 / potentialTarget)
                                        quantity = _LotSize * multiplier

                                        capital = currentOptTick.Open * quantity / _ParentStrategy.MarginMultiplier
                                        If lossToRecover < 0 Then capital = capital + Math.Abs(lossToRecover)
                                        potentialTarget = capital * _ParentStrategy.RuleSettings.CapitalPercentage / 100
                                    End If
                                End If

                                Dim tradeToPlace As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                                      tradingSymbol:=optionTradingSymbol,
                                                                      spotTradingSymbol:=_TradingSymbol,
                                                                      stockType:=Trade.TypeOfStock.Options,
                                                                      tradingDate:=currentTickTime.Date,
                                                                      signalDirection:=directionToCheck,
                                                                      entryDirection:=Trade.TradeDirection.Buy,
                                                                      entryType:=entryType,
                                                                      entryPrice:=currentOptTick.Open,
                                                                      quantity:=quantity,
                                                                      lotSize:=_LotSize,
                                                                      entrySignalCandle:=signal.Item2,
                                                                      childReference:=childTag,
                                                                      parentReference:=parentTag,
                                                                      iterationNumber:=iterationNumber,
                                                                      spotPrice:=currentSpotTick.Open,
                                                                      spotATR:=spotATR,
                                                                      previousLoss:=lossToRecover,
                                                                      potentialTarget:=Math.Max(potentialTarget, 1000))

                                If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload))
                                ret.Add(New Tuple(Of Trade, Payload)(tradeToPlace, currentOptTick))
                            End If
                        End If
                    End If
                End If
            Else
                If lastTrade IsNot Nothing AndAlso lastTrade.TradeCurrentStatus <> Trade.TradeStatus.Inprogress AndAlso lastTrade.TradeCurrentStatus <> Trade.TradeStatus.Open Then
                    Dim lastCompleteTrade As Trade = _ParentStrategy.GetLastCompleteTradeOfTheStock(_TradingSymbol)
                    If lastCompleteTrade IsNot Nothing AndAlso lastCompleteTrade.ParentReference = lastTrade.ParentReference Then
                        If lastCompleteTrade.ExitType = Trade.TypeOfExit.ContractRollover Then
                            Dim currentSpotTick As Payload = GetCurrentTick(lastCompleteTrade.SpotTradingSymbol, currentTickTime)
                            Dim currentOptionExpiryString As String = GetOptionInstrumentExpiryString(lastCompleteTrade.SpotTradingSymbol, _NextTradingDay)
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, currentOptionExpiryString, currentSpotTick.Open, lastCompleteTrade.SignalDirection)
                            If optionTradingSymbol IsNot Nothing Then
                                Dim currentOptTick As Payload = GetCurrentTick(optionTradingSymbol, currentTickTime)
                                If currentOptTick IsNot Nothing Then
                                    Dim lossToRecover As Decimal = GetOverallPL(lastTrade, Nothing)
                                    Dim potentialTarget As Decimal = lastCompleteTrade.PotentialTarget
                                    If _ParentStrategy.RuleSettings.TypeOfTarget = RuleEntities.TargetType.CapitalPercentage Then
                                        Dim capital As Decimal = currentOptTick.Open * lastCompleteTrade.Quantity / _ParentStrategy.MarginMultiplier
                                        If lossToRecover < 0 Then capital = capital + Math.Abs(lossToRecover)
                                        potentialTarget = capital * _ParentStrategy.RuleSettings.CapitalPercentage / 100
                                    End If

                                    Dim tradeToPlace As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                                          tradingSymbol:=optionTradingSymbol,
                                                                          spotTradingSymbol:=lastCompleteTrade.SpotTradingSymbol,
                                                                          stockType:=lastCompleteTrade.StockType,
                                                                          tradingDate:=currentTickTime.Date,
                                                                          signalDirection:=lastCompleteTrade.SignalDirection,
                                                                          entryDirection:=lastCompleteTrade.EntryDirection,
                                                                          entryType:=lastCompleteTrade.EntryType,
                                                                          entryPrice:=currentOptTick.Open,
                                                                          quantity:=lastCompleteTrade.Quantity,
                                                                          lotSize:=lastCompleteTrade.LotSize,
                                                                          entrySignalCandle:=lastCompleteTrade.EntrySignalCandle,
                                                                          childReference:=lastCompleteTrade.ChildReference,
                                                                          parentReference:=lastCompleteTrade.ParentReference,
                                                                          iterationNumber:=lastCompleteTrade.IterationNumber,
                                                                          spotPrice:=lastCompleteTrade.SpotPrice,
                                                                          spotATR:=lastCompleteTrade.SpotATR,
                                                                          previousLoss:=lastCompleteTrade.PreviousLoss,
                                                                          potentialTarget:=Math.Max(potentialTarget, 1000))
                                    tradeToPlace.UpdateTrade(contractRolloverEntry:=True)
                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload))
                                    ret.Add(New Tuple(Of Trade, Payload)(tradeToPlace, currentOptTick))
                                End If
                            End If
                        ElseIf lastCompleteTrade.ExitType = Trade.TypeOfExit.Reversal Then
                            Dim childTag As String = System.Guid.NewGuid.ToString()
                            Dim parentTag As String = lastTrade.ParentReference
                            Dim iterationNumber As Integer = lastCompleteTrade.IterationNumber + 1
                            Dim entryType As Trade.TypeOfEntry = Trade.TypeOfEntry.Reversal
                            Dim lossToRecover As Decimal = GetOverallPL(lastTrade, Nothing)
                            Dim currentSpotTick As Payload = GetCurrentTick(_TradingSymbol, currentTickTime)
                            Dim spotATR As Decimal = _ATRPayload(lastCompleteTrade.ExitSignalCandle.PayloadDate.Date)
                            Dim directionToCheck As Trade.TradeDirection = Trade.TradeDirection.None
                            If lastCompleteTrade.SignalDirection = Trade.TradeDirection.Buy Then
                                directionToCheck = Trade.TradeDirection.Sell
                            ElseIf lastCompleteTrade.SignalDirection = Trade.TradeDirection.Sell Then
                                directionToCheck = Trade.TradeDirection.Buy
                            End If

                            Dim optionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
                            Dim optionTradingSymbol As String = GetCurrentATMOption(currentTickTime, optionExpiryString, currentSpotTick.Open, directionToCheck)
                            If optionTradingSymbol IsNot Nothing Then
                                Dim currentOptTick As Payload = GetCurrentTick(optionTradingSymbol, currentSpotTick.PayloadDate)
                                If currentOptTick IsNot Nothing Then
                                    Dim quantity As Integer = _LotSize
                                    Dim entryPrice As Decimal = currentOptTick.Open
                                    Dim targetPrice As Decimal = ConvertFloorCeling(entryPrice + spotATR, _ParentStrategy.TickSize, RoundOfType.Celing)
                                    Dim potentialTarget As Decimal = Math.Abs(lossToRecover)
                                    For ctr As Integer = 1 To Integer.MaxValue
                                        Dim pl As Decimal = _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, entryPrice, targetPrice, ctr * _LotSize, _LotSize, Trade.TypeOfStock.Options)
                                        If pl >= lossToRecover * -1 Then
                                            quantity = ctr * _LotSize + _LotSize
                                            'potentialTarget = _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, entryPrice, targetPrice, quantity, _LotSize, Trade.TypeOfStock.Options)
                                            potentialTarget = pl - 1
                                            Exit For
                                        End If
                                    Next
                                    If _ParentStrategy.RuleSettings.TypeOfTarget = RuleEntities.TargetType.CapitalPercentage Then
                                        Dim capital As Decimal = currentOptTick.Open * quantity / _ParentStrategy.MarginMultiplier
                                        If lossToRecover < 0 Then capital = capital + Math.Abs(lossToRecover)
                                        potentialTarget = capital * _ParentStrategy.RuleSettings.CapitalPercentage / 100
                                    End If

                                    Dim tradeToPlace As Trade = New Trade(originatingStrategy:=_ParentStrategy,
                                                                          tradingSymbol:=optionTradingSymbol,
                                                                          spotTradingSymbol:=lastCompleteTrade.SpotTradingSymbol,
                                                                          stockType:=lastCompleteTrade.StockType,
                                                                          tradingDate:=currentTickTime.Date,
                                                                          signalDirection:=directionToCheck,
                                                                          entryDirection:=Trade.TradeDirection.Buy,
                                                                          entryType:=entryType,
                                                                          entryPrice:=currentOptTick.Open,
                                                                          quantity:=quantity,
                                                                          lotSize:=_LotSize,
                                                                          entrySignalCandle:=lastCompleteTrade.ExitSignalCandle,
                                                                          childReference:=childTag,
                                                                          parentReference:=parentTag,
                                                                          iterationNumber:=iterationNumber,
                                                                          spotPrice:=currentSpotTick.Open,
                                                                          spotATR:=spotATR,
                                                                          previousLoss:=lossToRecover,
                                                                          potentialTarget:=Math.Max(potentialTarget, 1000))

                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload))
                                    ret.Add(New Tuple(Of Trade, Payload)(tradeToPlace, currentOptTick))
                                End If
                            End If
                        Else
                            Throw New NotImplementedException
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
                            Dim targetReached As Boolean = False
                            If _ParentStrategy.RuleSettings.TypeOfTarget = RuleEntities.TargetType.CapitalPercentage Then
                                If GetOverallPL(runningTrade, currentTick) >= runningTrade.PotentialTarget Then
                                    targetReached = True
                                End If
                            ElseIf _ParentStrategy.RuleSettings.TypeOfTarget = RuleEntities.TargetType.ATR Then
                                If runningTrade.EntryType = Trade.TypeOfEntry.Fresh AndAlso
                                    GetFreshTradePL(runningTrade, currentTick) >= runningTrade.PotentialTarget Then
                                    targetReached = True
                                ElseIf runningTrade.EntryType = Trade.TypeOfEntry.Reversal AndAlso
                                    GetLossMakeupTradePL(runningTrade, currentTick) >= runningTrade.PotentialTarget Then
                                    targetReached = True
                                End If
                            Else
                                Throw New NotImplementedException
                            End If
                            If targetReached Then
                                If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.Target, Nothing))
                            Else
                                Dim reverseSignal As Tuple(Of Boolean, Payload) = IsReverseSignalReceived(runningTrade, currentTick)
                                If reverseSignal IsNot Nothing AndAlso reverseSignal.Item1 AndAlso reverseSignal.Item2 IsNot Nothing Then
                                    If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                    ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.Reversal, reverseSignal.Item2))
                                Else
                                    Dim expiryDate As Date = GetLastThusrdayOfMonth(_TradingDate).Date.AddDays(-3)
                                    expiryDate = New Date(expiryDate.Year, expiryDate.Month, expiryDate.Day, 15, 29, 0)
                                    If currentTickTime >= expiryDate Then
                                        Dim currentOptionExpiryString As String = GetOptionInstrumentExpiryString(_TradingSymbol, _NextTradingDay)
                                        If Not runningTrade.TradingSymbol.StartsWith(currentOptionExpiryString) Then
                                            If ret Is Nothing Then ret = New List(Of Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload))
                                            ret.Add(New Tuple(Of Trade, Payload, Trade.TypeOfExit, Payload)(runningTrade, currentTick, Trade.TypeOfExit.ContractRollover, Nothing))
                                        End If
                                    End If
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

        Private Function GetOverallPL(ByVal currentTrade As Trade, ByVal currentTick As Payload) As Decimal
            Dim ret As Decimal = 0
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByParentTag(currentTrade.ParentReference, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.TradeCurrentStatus <> Trade.TradeStatus.Cancel Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
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

        Private Function GetLossMakeupTradePL(ByVal currentTrade As Trade, ByVal currentTick As Payload) As Decimal
            Dim ret As Decimal = 0
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByChildTag(currentTrade.ChildReference, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.EntryType = Trade.TypeOfEntry.Reversal Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                            ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, runningTrade.EntryPrice, currentTick.Open, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        Else
                            ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, runningTrade.EntryPrice, runningTrade.ExitPrice, runningTrade.Quantity - runningTrade.LotSize, runningTrade.LotSize, runningTrade.StockType)
                        End If
                    End If
                Next
            End If
            Return ret
        End Function

        Private Function GetFreshTradePL(ByVal currentTrade As Trade, ByVal currentTick As Payload) As Decimal
            Dim ret As Decimal = 0
            Dim allTrades As List(Of Trade) = _ParentStrategy.GetAllTradesByChildTag(currentTrade.ChildReference, _TradingSymbol)
            If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                For Each runningTrade In allTrades
                    If runningTrade.EntryType = Trade.TypeOfEntry.Fresh Then
                        If runningTrade.TradeCurrentStatus = Trade.TradeStatus.Inprogress Then
                            ret += _ParentStrategy.CalculatePLAfterBrokerage(_TradingSymbol, runningTrade.EntryPrice, currentTick.Open, runningTrade.Quantity, runningTrade.LotSize, runningTrade.StockType)
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
        Protected MustOverride Function IsReverseSignalReceived(ByVal currentTrade As Trade, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload)

        Protected MustOverride Function IsFreshEntrySignalReceived(ByVal currentTickTime As Date) As Tuple(Of Boolean, Payload, Trade.TradeDirection)

        Protected Function GetOptionInstrumentExpiryString(ByVal coreInstrumentName As String, ByVal tradingDate As Date) As String
            Dim ret As String = Nothing
            Dim lastThursday As Date = GetLastThusrdayOfMonth(tradingDate)
            If tradingDate.Date > lastThursday.Date.AddDays(-3) Then
                ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.AddDays(10).ToString("yyMMM")).ToUpper
            Else
                ret = String.Format("{0}{1}", coreInstrumentName, tradingDate.ToString("yyMMM")).ToUpper
            End If
            Return ret
        End Function

        Protected Function GetCurrentATMOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal, ByVal signalDirection As Trade.TradeDirection) As String
            Dim ret As String = Nothing
            If signalDirection = Trade.TradeDirection.Buy Then
                ret = GetCurrentATMCEOption(currentTickTime, expiryString, price)
            ElseIf signalDirection = Trade.TradeDirection.Sell Then
                ret = GetCurrentATMPEOption(currentTickTime, expiryString, price)
            End If
            Return ret
        End Function
#End Region

#Region "Contract Helper"
        Private Function GetCurrentATMPEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
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

        Private Function GetCurrentATMCEOption(ByVal currentTickTime As Date, ByVal expiryString As String, ByVal price As Decimal) As String
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
#End Region
    End Class
End Namespace