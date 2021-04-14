Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class SectoralTopGainerTopLooserStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Public Direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public DummyCandle As Payload = Nothing
    Public GainLossPercentageData As Dictionary(Of Date, Decimal) = Nothing

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _lastDayClose As Decimal = Decimal.MinValue
    Private _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal rawInstrumentName As String,
                   ByVal entities As RuleEntities,
                   ByVal controller As Integer,
                   ByVal canceller As CancellationTokenSource)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, rawInstrumentName, entities, controller, canceller)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        If Me.Controller Then
            Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)

            _lastDayClose = _signalPayload.Where(Function(x)
                                                     Return x.Key.Date < _tradingDate.Date
                                                 End Function).LastOrDefault.Value.Close
        Else
            Dim currentDayPayload As Dictionary(Of Date, Payload) = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If currentDayPayload Is Nothing Then currentDayPayload = New Dictionary(Of Date, Payload)
                    currentDayPayload.Add(runningPayload.Key, runningPayload.Value)
                End If
            Next
            If currentDayPayload IsNot Nothing AndAlso currentDayPayload.Count > 0 Then
                Dim previousClose As Decimal = currentDayPayload.FirstOrDefault.Value.PreviousCandlePayload.Close
                For Each runningPayload In currentDayPayload
                    Dim gainLossPercentage As Decimal = ((runningPayload.Value.Close - previousClose) / previousClose) * 100
                    If GainLossPercentageData Is Nothing Then GainLossPercentageData = New Dictionary(Of Date, Decimal)
                    GainLossPercentageData.Add(runningPayload.Key, gainLossPercentage)
                Next
            End If
        End If
    End Sub

    Public Overrides Sub CompleteDependentProcessing()
        MyBase.CompleteDependentProcessing()
        If Me.Controller Then
            If Me.DependentInstruments IsNot Nothing AndAlso Me.DependentInstruments.Count > 0 Then
                For Each runningInstrument In Me.DependentInstruments
                    runningInstrument.EligibleToTakeTrade = False
                Next
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Me.DummyCandle = currentTick
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If Me.ForceTakeTrade AndAlso Not Controller Then
            Dim quantity As Integer = Me.LotSize
            Dim stoploss As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me.TradingSymbol, currentTick.Open, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Me.Direction, _parentStrategy.StockType)
            Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me.TradingSymbol, currentTick.Open, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Me.Direction, _parentStrategy.StockType)

            If Me.Direction = Trade.TradeExecutionDirection.Buy Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = Math.Abs(currentTick.Open - stoploss)
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.Direction = Trade.TradeExecutionDirection.None
                Me.ForceTakeTrade = False
            ElseIf Me.Direction = Trade.TradeExecutionDirection.Sell Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = Math.Abs(currentTick.Open - stoploss)
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.Direction = Trade.TradeExecutionDirection.None
                Me.ForceTakeTrade = False
            End If
        ElseIf Controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            Dim parameter As PlaceOrderParameters = Nothing
            If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
                currentMinuteCandle.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                Dim hkCandle As Payload = _hkPayload(currentMinuteCandle.PayloadDate)
                If hkCandle.PreviousCandlePayload IsNot Nothing Then
                    If hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish AndAlso
                        hkCandle.PreviousCandlePayload.High > _lastDayClose Then
                        Dim buyPrice As Decimal = ConvertFloorCeling(hkCandle.PreviousCandlePayload.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        If currentTick.Open >= buyPrice Then
                            Dim allSectors As List(Of String) = _parentStrategy.AllSectoralStockList.Keys.ToList
                            For Each runningSector In allSectors
                                If Not IsSectorActive(runningSector, Trade.TradeExecutionDirection.Buy) Then
                                    Dim ctr As Integer = 0
                                    While True
                                        ctr += 1
                                        Dim gainer As SectoralTopGainerTopLooserStrategyRule = GetSectoralTopGainerStock(runningSector, currentMinuteCandle, ctr)
                                        If gainer IsNot Nothing Then
                                            Dim stockPrice As Decimal = gainer._signalPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate).Close
                                            Dim stoploss As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me.TradingSymbol, stockPrice, gainer.LotSize, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                                            If stockPrice - stoploss > 0 Then
                                                gainer.EligibleToTakeTrade = True
                                                gainer.ForceTakeTrade = True
                                                gainer.Direction = Trade.TradeExecutionDirection.Buy
                                            End If
                                        Else
                                            Exit While
                                        End If
                                    End While
                                End If
                            Next
                        End If
                    ElseIf hkCandle.PreviousCandlePayload.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish AndAlso
                        hkCandle.PreviousCandlePayload.Low < _lastDayClose Then
                        Dim sellPrice As Decimal = ConvertFloorCeling(hkCandle.PreviousCandlePayload.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        If currentTick.Open <= sellPrice Then
                            Dim allSectors As List(Of String) = _parentStrategy.AllSectoralStockList.Keys.ToList
                            For Each runningSector In allSectors
                                If Not IsSectorActive(runningSector, Trade.TradeExecutionDirection.Sell) Then
                                    Dim ctr As Integer = 0
                                    While True
                                        ctr += 1
                                        Dim looser As SectoralTopGainerTopLooserStrategyRule = GetSectoralTopLooserStock(runningSector, currentMinuteCandle, ctr)
                                        If looser IsNot Nothing Then
                                            Dim stockPrice As Decimal = looser._signalPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate).Close
                                            Dim stoploss As Decimal = _parentStrategy.CalculatorTargetOrStoploss(Me.TradingSymbol, stockPrice, looser.LotSize, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TradeExecutionDirection.Sell, _parentStrategy.StockType)
                                            If stoploss - stockPrice > 0 Then
                                                looser.EligibleToTakeTrade = True
                                                looser.ForceTakeTrade = True
                                                looser.Direction = Trade.TradeExecutionDirection.Sell
                                            End If
                                        Else
                                            Exit While
                                        End If
                                    End While
                                End If
                            Next
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
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting1
            Dim potentialSL As Decimal = Decimal.MinValue
            Dim brkevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(Me.TradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                potentialSL = currentTrade.EntryPrice + brkevenPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                potentialSL = currentTrade.EntryPrice - brkevenPoint
            End If
            If potentialSL <> Decimal.MinValue AndAlso potentialSL <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, potentialSL, "Breakeven Movement")
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsSectorActive(ByVal sector As String, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = False
        Dim sectoralStocklist As List(Of String) = _parentStrategy.GetAllStockOfSector(sector)
        If sectoralStocklist IsNot Nothing AndAlso sectoralStocklist.Count > 0 Then
            If Me.DependentInstruments IsNot Nothing AndAlso Me.DependentInstruments.Count > 0 Then
                For Each runningInstrument As SectoralTopGainerTopLooserStrategyRule In Me.DependentInstruments
                    If sectoralStocklist.Contains(runningInstrument.RawInstrumentName) Then
                        If runningInstrument.ForceTakeTrade OrElse (runningInstrument.DummyCandle IsNot Nothing AndAlso
                            _parentStrategy.IsTradeActive(runningInstrument.DummyCandle, Trade.TypeOfTrade.MIS, direction)) Then
                            ret = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        Return ret
    End Function

    Private Function GetSectoralTopGainerStock(ByVal sector As String, ByVal currentMinuteCandle As Payload, ByVal position As Integer) As SectoralTopGainerTopLooserStrategyRule
        Dim ret As SectoralTopGainerTopLooserStrategyRule = Nothing
        Dim sectoralStocklist As List(Of String) = _parentStrategy.GetAllStockOfSector(sector)
        If sectoralStocklist IsNot Nothing AndAlso sectoralStocklist.Count > 0 Then
            If Me.DependentInstruments IsNot Nothing AndAlso Me.DependentInstruments.Count > 0 Then
                Dim gainLossOfTheMinute As Dictionary(Of String, Decimal) = Nothing
                For Each runningInstrument As SectoralTopGainerTopLooserStrategyRule In Me.DependentInstruments
                    If runningInstrument.GainLossPercentageData IsNot Nothing AndAlso sectoralStocklist.Contains(runningInstrument.RawInstrumentName) AndAlso
                        runningInstrument.GainLossPercentageData.ContainsKey(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                        If gainLossOfTheMinute Is Nothing Then gainLossOfTheMinute = New Dictionary(Of String, Decimal)
                        gainLossOfTheMinute.Add(runningInstrument.TradingSymbol, runningInstrument.GainLossPercentageData(currentMinuteCandle.PreviousCandlePayload.PayloadDate))
                    End If
                Next
                If gainLossOfTheMinute IsNot Nothing AndAlso gainLossOfTheMinute.Count > 0 Then
                    Dim counter As Integer = 0
                    For Each runningStock In gainLossOfTheMinute.OrderByDescending(Function(x)
                                                                                       Return x.Value
                                                                                   End Function)
                        counter += 1
                        If counter = position Then
                            ret = GetStrategyRuleByName(runningStock.Key)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSectoralTopLooserStock(ByVal sector As String, ByVal currentMinuteCandle As Payload, ByVal position As Integer) As SectoralTopGainerTopLooserStrategyRule
        Dim ret As SectoralTopGainerTopLooserStrategyRule = Nothing
        Dim sectoralStocklist As List(Of String) = _parentStrategy.GetAllStockOfSector(sector)
        If sectoralStocklist IsNot Nothing AndAlso sectoralStocklist.Count > 0 Then
            If Me.DependentInstruments IsNot Nothing AndAlso Me.DependentInstruments.Count > 0 Then
                Dim gainLossOfTheMinute As Dictionary(Of String, Decimal) = Nothing
                For Each runningInstrument As SectoralTopGainerTopLooserStrategyRule In Me.DependentInstruments
                    If runningInstrument.GainLossPercentageData IsNot Nothing AndAlso sectoralStocklist.Contains(runningInstrument.RawInstrumentName) AndAlso
                        runningInstrument.GainLossPercentageData.ContainsKey(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                        If gainLossOfTheMinute Is Nothing Then gainLossOfTheMinute = New Dictionary(Of String, Decimal)
                        gainLossOfTheMinute.Add(runningInstrument.TradingSymbol, runningInstrument.GainLossPercentageData(currentMinuteCandle.PreviousCandlePayload.PayloadDate))
                    End If
                Next
                If gainLossOfTheMinute IsNot Nothing AndAlso gainLossOfTheMinute.Count > 0 Then
                    Dim counter As Integer = 0
                    For Each runningStock In gainLossOfTheMinute.OrderBy(Function(x)
                                                                             Return x.Value
                                                                         End Function)
                        counter += 1
                        If counter = position Then
                            ret = GetStrategyRuleByName(runningStock.Key)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetStrategyRuleByName(ByVal tradingSymbol As String) As SectoralTopGainerTopLooserStrategyRule
        Dim ret As SectoralTopGainerTopLooserStrategyRule = Nothing
        If Me.DependentInstruments IsNot Nothing AndAlso Me.DependentInstruments.Count > 0 Then
            For Each runningInstrument In Me.DependentInstruments
                If runningInstrument.TradingSymbol.ToUpper = tradingSymbol.ToUpper Then
                    ret = runningInstrument
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
End Class