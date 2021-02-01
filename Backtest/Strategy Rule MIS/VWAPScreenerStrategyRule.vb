Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class VWAPScreenerStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Enum SLMovementType
        None = 1
        CostToCost
        Trailling
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
        Public SLMovement As SLMovementType
        Public RSILevel As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _mVwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _rsiPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _pivotPayload As Dictionary(Of Date, PivotPoints) = Nothing

    Private _lastDayMA As Decimal = Decimal.MinValue

    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)

        Indicator.VWAP.CalculateVWAP(_hkPayload, _vwapPayload)

        Dim convertedVWAPPayload As Dictionary(Of Date, Payload) = Nothing
        Common.ConvertDecimalToPayload(Payload.PayloadFields.Additional_Field, _vwapPayload, convertedVWAPPayload)
        Indicator.EMA.CalculateEMA(50, Payload.PayloadFields.Additional_Field, convertedVWAPPayload, _mVwapPayload)

        Indicator.RSI.CalculateRSI(14, _hkPayload, _rsiPayload)

        Indicator.Pivots.CalculatePivots(_hkPayload, _pivotPayload)

        If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
            Dim eodPayload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, _tradingSymbol, _tradingDate.AddYears(-1), _tradingDate.AddDays(-1))
            If eodPayload IsNot Nothing AndAlso eodPayload.Count > 0 Then
                Dim eodSMAPayload As Dictionary(Of Date, Decimal) = Nothing
                Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, eodPayload, eodSMAPayload)
                _lastDayMA = Math.Round(eodSMAPayload.LastOrDefault.Value, 2)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim runningCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If runningCandle IsNot Nothing AndAlso runningCandle.PreviousCandlePayload IsNot Nothing AndAlso runningCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String) = GetEntrySignal(runningCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastEntryTradeOfTheStock(runningCandle, _parentStrategy.TradeType)
                If lastOrder Is Nothing Then
                    signalCandle = runningCandle.PreviousCandlePayload
                ElseIf lastOrder IsNot Nothing Then
                    If lastOrder.SignalCandle.PayloadDate <> runningCandle.PreviousCandlePayload.PayloadDate Then
                        signalCandle = runningCandle.PreviousCandlePayload
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item3.EntryPrice
                    Dim stoploss As Decimal = signal.Item3.StoplossPrice
                    Dim target As Decimal = signal.Item3.TargetPrice
                    If _userInputs.SLMovement = SLMovementType.Trailling Then
                        target = entryPrice + 100000000
                    End If
                    Dim quantity As Integer = signal.Item4

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = entryPrice - stoploss,
                                    .Supporting4 = target - entryPrice
                                }
                ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item3.EntryPrice
                    Dim stoploss As Decimal = signal.Item3.StoplossPrice
                    Dim target As Decimal = signal.Item3.TargetPrice
                    If _userInputs.SLMovement = SLMovementType.Trailling Then
                        target = entryPrice - 100000000
                    End If
                    Dim quantity As Integer = signal.Item4

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = signal.Item5,
                                    .Supporting3 = stoploss - entryPrice,
                                    .Supporting4 = entryPrice - target
                                }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim runningCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim exitTrade As Boolean = False
            Dim reason As String = ""
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String) = GetEntrySignal(runningCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection = signal.Item2 Then
                    If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                        If currentTrade.EntryPrice <> signal.Item3.EntryPrice AndAlso currentTick.Open < signal.Item3.EntryPrice Then
                            exitTrade = True
                            reason = "New entry signal"
                        End If
                    ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                        If currentTrade.EntryPrice <> signal.Item3.EntryPrice AndAlso currentTick.Open > signal.Item3.EntryPrice Then
                            exitTrade = True
                            reason = "New entry signal"
                        End If
                    End If
                ElseIf currentTrade.EntryDirection <> signal.Item2 Then
                    exitTrade = True
                    reason = "Opposite direction signal"
                End If
            End If
            If Not exitTrade Then
                Dim hkSignalCandle As Payload = _hkPayload(currentTrade.SignalCandle.PayloadDate)
                If hkSignalCandle IsNot Nothing Then
                    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                        If currentTick.Open <= hkSignalCandle.Low Then
                            exitTrade = True
                            reason = "Candle Low Hit"
                        Else
                            Dim signalData As SignalDetails = CalculateEntryStoplossTarget(hkSignalCandle, currentTrade.EntryDirection)
                            If signalData IsNot Nothing Then
                                If currentTick.Open >= signalData.TargetPrice Then
                                    exitTrade = True
                                    reason = "Target Reached"
                                End If
                            End If
                        End If
                    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                        If currentTick.Open >= hkSignalCandle.High Then
                            exitTrade = True
                            reason = "Candle High Hit"
                        Else
                            Dim signalData As SignalDetails = CalculateEntryStoplossTarget(hkSignalCandle, currentTrade.EntryDirection)
                            If signalData IsNot Nothing Then
                                If currentTick.Open <= signalData.TargetPrice Then
                                    exitTrade = True
                                    reason = "Target Reached"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If exitTrade Then
                ret = New Tuple(Of Boolean, String)(True, reason)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.SLMovement <> SLMovementType.None AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting3
            Dim tgtPoint As Decimal = currentTrade.Supporting4
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remark As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim plPoint As Decimal = currentTick.Open - currentTrade.EntryPrice
                If _userInputs.SLMovement = SLMovementType.CostToCost Then
                    If plPoint >= tgtPoint / 2 Then
                        triggerPrice = currentTrade.EntryPrice
                        remark = "Cost To Cost"
                    End If
                ElseIf _userInputs.SLMovement = SLMovementType.Trailling Then
                    If plPoint >= slPoint Then
                        Dim mul As Integer = Math.Floor(plPoint / slPoint)
                        Dim potentialSL As Decimal = currentTrade.EntryPrice + (mul - 1) * slPoint
                        If potentialSL > currentTrade.PotentialStopLoss Then
                            triggerPrice = potentialSL
                            remark = String.Format("Move to {0}", mul - 1)
                        End If
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim plPoint As Decimal = currentTrade.EntryPrice - currentTick.Open
                If _userInputs.SLMovement = SLMovementType.CostToCost Then
                    If plPoint >= tgtPoint / 2 Then
                        triggerPrice = currentTrade.EntryPrice
                        remark = "Cost To Cost"
                    End If
                ElseIf _userInputs.SLMovement = SLMovementType.Trailling Then
                    If plPoint >= slPoint Then
                        Dim mul As Integer = Math.Floor(plPoint / slPoint)
                        Dim potentialSL As Decimal = currentTrade.EntryPrice - (mul - 1) * slPoint
                        If potentialSL < currentTrade.PotentialStopLoss Then
                            triggerPrice = potentialSL
                            remark = String.Format("Move to {0}", mul - 1)
                        End If
                    End If
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, remark)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

#Region "Signal Check"
    Private Function GetEntrySignal(ByVal runningCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String) = Nothing
        If runningCandle IsNot Nothing AndAlso runningCandle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(runningCandle.PreviousCandlePayload.PayloadDate)
            Dim vwap As Decimal = _vwapPayload(runningCandle.PreviousCandlePayload.PayloadDate)
            Dim mVWAP As Decimal = _mVwapPayload(runningCandle.PreviousCandlePayload.PayloadDate)
            Dim pivots As PivotPoints = _pivotPayload(runningCandle.PreviousCandlePayload.PayloadDate)
            Dim rsi As Decimal = _rsiPayload(runningCandle.PreviousCandlePayload.PayloadDate)

            Dim signalCandle As Payload = hkCandle
            If signalCandle IsNot Nothing AndAlso signalCandle.PreviousCandlePayload IsNot Nothing AndAlso
                signalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                Dim message As String = String.Format("{0} ->Signal Candle Time:{1}.",
                                                      _tradingSymbol,
                                                      signalCandle.PayloadDate.ToString("HH:mm:ss"))

                If vwap > mVWAP Then 'Buy
                    Dim takeTrade As Boolean = True
                    message = String.Format("{0} VWAP({1})>MVWAP({2})[BUY]. [INFO1]",
                                            message, Math.Round(vwap, 2), Math.Round(mVWAP, 2))

                    takeTrade = takeTrade And (signalCandle.CandleColor = Color.Green)
                    message = String.Format("{0} Signal Candle Color({1})=Green[{2}].",
                                            message, signalCandle.CandleColor.Name, signalCandle.CandleColor = Color.Green)

                    takeTrade = takeTrade And (signalCandle.PreviousCandlePayload.CandleColor = Color.Red)
                    message = String.Format("{0} Previous Candle Color({1})=Red[{2}].",
                                            message, signalCandle.PreviousCandlePayload.CandleColor.Name, signalCandle.PreviousCandlePayload.CandleColor = Color.Red)

                    takeTrade = takeTrade And (signalCandle.High > signalCandle.PreviousCandlePayload.High)
                    message = String.Format("{0} Signal Candle High({1})>Previous Candle High({2})[{3}].",
                                            message, Math.Round(signalCandle.High, 2),
                                            Math.Round(signalCandle.PreviousCandlePayload.High, 2),
                                            signalCandle.High > signalCandle.PreviousCandlePayload.High)

                    takeTrade = takeTrade And (signalCandle.Low > signalCandle.PreviousCandlePayload.Low)
                    message = String.Format("{0} Signal Candle Low:({1})>Previous Candle Low({2})[{3}]",
                                            message, Math.Round(signalCandle.Low, 2),
                                            Math.Round(signalCandle.PreviousCandlePayload.Low, 2),
                                            signalCandle.Low > signalCandle.PreviousCandlePayload.Low)

                    takeTrade = takeTrade And (signalCandle.High > vwap)
                    message = String.Format("{0} Signal Candle High({1})>VWAP({2})[{3}].",
                                            message, Math.Round(signalCandle.High, 2),
                                            Math.Round(vwap, 2),
                                            signalCandle.High > vwap)

                    takeTrade = takeTrade And (vwap > pivots.Pivot)
                    message = String.Format("{0} VWAP({1})>Central Pivot({2})[{3}].",
                                            message, Math.Round(vwap, 2),
                                            Math.Round(pivots.Pivot, 2),
                                            vwap > pivots.Pivot)

                    takeTrade = takeTrade And (mVWAP > pivots.Pivot)
                    message = String.Format("{0} MVWAP({1})>Central Pivot({2})[{3}].",
                                            message, Math.Round(mVWAP, 2),
                                            Math.Round(pivots.Pivot, 2),
                                            mVWAP > pivots.Pivot)

                    takeTrade = takeTrade And (rsi > _userInputs.RSILevel)
                    message = String.Format("{0} RSI({1})>RSI Level({2})[{3}].",
                                            message, Math.Round(rsi, 2),
                                            Math.Round(_userInputs.RSILevel, 2),
                                            rsi > _userInputs.RSILevel)

                    If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                        takeTrade = takeTrade And (currentTick.Open > _lastDayMA)
                        message = String.Format("{0} LTP({1})>Last Day MA({2})[{3}].",
                                                message, currentTick.Open, Math.Round(_lastDayMA, 2), currentTick.Open > _lastDayMA)
                    End If

                    If takeTrade Then
                        Dim entrySLTgt As SignalDetails = CalculateEntryStoplossTarget(signalCandle, Trade.TradeExecutionDirection.Buy)
                        If entrySLTgt IsNot Nothing Then
                            Dim entryPrice As Decimal = entrySLTgt.EntryPrice
                            Dim stoploss As Decimal = entrySLTgt.StoplossPrice
                            Dim slRemark As String = entrySLTgt.StoplossRemark

                            message = message.Replace("[INFO1]", String.Format("Entry:{0}, Stoploss:{1}({2}).", entryPrice, stoploss, slRemark))

                            Dim quantity As Integer = Me.LotSize
                            If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                                quantity = Math.Ceiling(Math.Abs(_userInputs.MaxLossPerTrade) / (entryPrice - stoploss))
                            End If

                            takeTrade = takeTrade And (quantity > 0)
                            message = String.Format("{0} Quantity({1})>0[{2}]",
                                            message,
                                            quantity,
                                            quantity > 0)

                            If takeTrade Then
                                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String)(True, Trade.TradeExecutionDirection.Buy, entrySLTgt, quantity, message)
                            End If
                        Else
                            message = message.Replace("[INFO1]", "Can not calculate Entry Stoploss Target")
                        End If
                    Else
                        message = message.Replace("[INFO1]", "")
                    End If
                ElseIf vwap < mVWAP Then 'Sell
                    Dim takeTrade As Boolean = True
                    message = String.Format("{0} VWAP({1})<MVWAP({2})[SELL]. [INFO1]",
                                            message, Math.Round(vwap, 2), Math.Round(mVWAP, 2))

                    takeTrade = takeTrade And (signalCandle.CandleColor = Color.Red)
                    message = String.Format("{0} Signal Candle Color({1})=Red[{2}].",
                                            message, signalCandle.CandleColor.Name, signalCandle.CandleColor = Color.Red)

                    takeTrade = takeTrade And (signalCandle.PreviousCandlePayload.CandleColor = Color.Green)
                    message = String.Format("{0} Previous Candle Color({1})=Green[{2}].",
                                            message, signalCandle.PreviousCandlePayload.CandleColor.Name, signalCandle.PreviousCandlePayload.CandleColor = Color.Green)

                    takeTrade = takeTrade And (signalCandle.High < signalCandle.PreviousCandlePayload.High)
                    message = String.Format("{0} Signal Candle High({1})<Previous Candle High({2})[{3}].",
                                            message, Math.Round(signalCandle.High, 2),
                                            Math.Round(signalCandle.PreviousCandlePayload.High, 2),
                                            signalCandle.High < signalCandle.PreviousCandlePayload.High)

                    takeTrade = takeTrade And (signalCandle.Low < signalCandle.PreviousCandlePayload.Low)
                    message = String.Format("{0} Signal Candle Low:({1})<Previous Candle Low({2})[{3}]",
                                            message, Math.Round(signalCandle.Low, 2),
                                            Math.Round(signalCandle.PreviousCandlePayload.Low, 2),
                                            signalCandle.Low < signalCandle.PreviousCandlePayload.Low)

                    takeTrade = takeTrade And (signalCandle.Low < vwap)
                    message = String.Format("{0} Signal Candle Low({1})<VWAP({2})[{3}].",
                                            message, Math.Round(signalCandle.Low, 2),
                                            Math.Round(vwap, 2),
                                            signalCandle.Low < vwap)

                    takeTrade = takeTrade And (vwap < pivots.Pivot)
                    message = String.Format("{0} VWAP({1})<Central Pivot({2})[{3}].",
                                            message, Math.Round(vwap, 2),
                                            Math.Round(pivots.Pivot, 2),
                                            vwap < pivots.Pivot)

                    takeTrade = takeTrade And (mVWAP < pivots.Pivot)
                    message = String.Format("{0} MVWAP({1})<Central Pivot({2})[{3}].",
                                            message, Math.Round(mVWAP, 2),
                                            Math.Round(pivots.Pivot, 2),
                                            mVWAP < pivots.Pivot)

                    takeTrade = takeTrade And (rsi < _userInputs.RSILevel)
                    message = String.Format("{0} RSI({1})<RSI Level({2})[{3}].",
                                            message, Math.Round(rsi, 2),
                                            Math.Round(_userInputs.RSILevel, 2),
                                            rsi < _userInputs.RSILevel)

                    If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                        takeTrade = takeTrade And (currentTick.Open < _lastDayMA)
                        message = String.Format("{0} LTP({1})<Last Day MA({2})[{3}].",
                                                message, currentTick.Open, Math.Round(_lastDayMA, 2), currentTick.Open < _lastDayMA)
                    End If

                    If takeTrade Then
                        Dim entrySLTgt As SignalDetails = CalculateEntryStoplossTarget(signalCandle, Trade.TradeExecutionDirection.Sell)
                        If entrySLTgt IsNot Nothing Then
                            Dim entryPrice As Decimal = entrySLTgt.EntryPrice
                            Dim stoploss As Decimal = entrySLTgt.StoplossPrice
                            Dim slRemark As String = entrySLTgt.StoplossRemark


                            message = message.Replace("[INFO1]", String.Format("Entry:{0}, Stoploss:{1}({2}).", entryPrice, stoploss, slRemark))

                            Dim quantity As Integer = Me.LotSize
                            If _parentStrategy.StockType = Trade.TypeOfStock.Cash Then
                                quantity = Math.Ceiling(Math.Abs(_userInputs.MaxLossPerTrade) / (stoploss - entryPrice))
                            End If

                            takeTrade = takeTrade And (quantity > 0)
                            message = String.Format("{0} Quantity({1})>0[{2}]",
                                                message,
                                                quantity,
                                                quantity > 0)

                            If takeTrade Then
                                ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, SignalDetails, Integer, String)(True, Trade.TradeExecutionDirection.Sell, entrySLTgt, quantity, message)
                            End If
                        Else
                            message = message.Replace("[INFO1]", "Can not calculate Entry Stoploss Target")
                        End If
                    Else
                        message = message.Replace("[INFO1]", "")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function CalculateEntryStoplossTarget(ByVal signalCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As SignalDetails
        Dim ret As SignalDetails = Nothing
        If signalCandle IsNot Nothing AndAlso signalCandle.PreviousCandlePayload IsNot Nothing Then
            Dim vwap As Decimal = _vwapPayload(signalCandle.PayloadDate)
            Dim mVWAP As Decimal = _mVwapPayload(signalCandle.PayloadDate)
            Dim pivots As PivotPoints = _pivotPayload(signalCandle.PayloadDate)
            Dim entryPrice As Decimal = Decimal.MinValue
            Dim stoplossPrice As Decimal = Decimal.MinValue
            Dim stoplossRemark As String = Nothing
            Dim targetPrice As Decimal = Decimal.MinValue
            Dim targetRemark As String = Nothing

            If direction = Trade.TradeExecutionDirection.Buy Then
                entryPrice = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                stoplossPrice = Decimal.MinValue
                If mVWAP < entryPrice Then
                    stoplossPrice = ConvertFloorCeling(mVWAP, _parentStrategy.TickSize, RoundOfType.Floor)
                    stoplossRemark = "MVWAP"

                    If pivots.Support3 > mVWAP AndAlso pivots.Support3 < vwap AndAlso pivots.Support3 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support3, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Support3"
                    End If
                    If pivots.Support2 > mVWAP AndAlso pivots.Support2 < vwap AndAlso pivots.Support2 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support2, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Support2"
                    End If
                    If pivots.Support1 > mVWAP AndAlso pivots.Support1 < vwap AndAlso pivots.Support1 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support1, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Support1"
                    End If
                    If pivots.Pivot > mVWAP AndAlso pivots.Pivot < vwap AndAlso pivots.Pivot < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Pivot, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Central Pivot"
                    End If
                    If pivots.Resistance3 > mVWAP AndAlso pivots.Resistance3 < vwap AndAlso pivots.Resistance3 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance3, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Resistance3"
                    End If
                    If pivots.Resistance2 > mVWAP AndAlso pivots.Resistance2 < vwap AndAlso pivots.Resistance2 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance2, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Resistance2"
                    End If
                    If pivots.Resistance1 > mVWAP AndAlso pivots.Resistance1 < vwap AndAlso pivots.Resistance1 < entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance1, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Resistance1"
                    End If

                    If signalCandle.PreviousCandlePayload.Low <= stoplossPrice Then
                        stoplossPrice = ConvertFloorCeling(signalCandle.PreviousCandlePayload.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        stoplossRemark = "Previous Candle Low"
                    End If

                    Dim entryBuffer As Decimal = CalculateTriggerBuffer(entryPrice)
                    entryPrice = entryPrice + entryBuffer
                    Dim slBuffer As Decimal = CalculateStoplossBuffer(stoplossPrice)
                    stoplossPrice = stoplossPrice - slBuffer

                    Dim targetPoint As Decimal = ConvertFloorCeling((entryPrice - stoplossPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    targetPrice = entryPrice + targetPoint
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                entryPrice = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                stoplossPrice = Decimal.MaxValue
                If mVWAP > entryPrice Then
                    stoplossPrice = ConvertFloorCeling(mVWAP, _parentStrategy.TickSize, RoundOfType.Celing)
                    stoplossRemark = "MVWAP"

                    If pivots.Resistance1 < mVWAP AndAlso pivots.Resistance1 > vwap AndAlso pivots.Resistance1 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance1, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Resistance1"
                    End If
                    If pivots.Resistance2 < mVWAP AndAlso pivots.Resistance2 > vwap AndAlso pivots.Resistance2 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance2, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Resistance2"
                    End If
                    If pivots.Resistance3 < mVWAP AndAlso pivots.Resistance3 > vwap AndAlso pivots.Resistance3 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Resistance3, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Resistance3"
                    End If
                    If pivots.Pivot < mVWAP AndAlso pivots.Pivot > vwap AndAlso pivots.Pivot > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Pivot, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Central Pivot"
                    End If
                    If pivots.Support1 < mVWAP AndAlso pivots.Support1 > vwap AndAlso pivots.Support1 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support1, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Support1"
                    End If
                    If pivots.Support2 < mVWAP AndAlso pivots.Support2 > vwap AndAlso pivots.Support2 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support2, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Support2"
                    End If
                    If pivots.Support3 < mVWAP AndAlso pivots.Support3 > vwap AndAlso pivots.Support3 > entryPrice Then
                        stoplossPrice = ConvertFloorCeling(pivots.Support3, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Support3"
                    End If

                    If signalCandle.PreviousCandlePayload.High >= stoplossPrice Then
                        stoplossPrice = ConvertFloorCeling(signalCandle.PreviousCandlePayload.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        stoplossRemark = "Previous Candle High"
                    End If

                    Dim entryBuffer As Decimal = CalculateTriggerBuffer(entryPrice)
                    entryPrice = entryPrice - entryBuffer
                    Dim slBuffer As Decimal = CalculateStoplossBuffer(stoplossPrice)
                    stoplossPrice = stoplossPrice + slBuffer

                    Dim targetPoint As Decimal = ConvertFloorCeling((stoplossPrice - entryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    targetPrice = entryPrice - targetPoint
                End If
            End If

            If entryPrice <> Decimal.MinValue AndAlso stoplossPrice <> Decimal.MinValue AndAlso
                stoplossPrice <> Decimal.MaxValue AndAlso targetPrice <> Decimal.MinValue Then
                ret = New SignalDetails With
                    {
                     .EntryPrice = entryPrice,
                     .StoplossPrice = stoplossPrice,
                     .StoplossRemark = stoplossRemark,
                     .TargetPrice = targetPrice,
                     .TargetRemark = targetRemark
                    }
            End If
        End If
        Return ret
    End Function
#End Region

#Region "Private Class"
    Private Class SignalDetails
        Public Property EntryPrice As Decimal
        Public Property StoplossPrice As Decimal
        Public Property StoplossRemark As String
        Public Property TargetPrice As Decimal
        Public Property TargetRemark As String
    End Class
#End Region

#Region "Required Functions"
    Private Function CalculateTriggerBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = _parentStrategy.TickSize
        If price <= 200 Then
            ret = 0.05
        ElseIf price > 200 AndAlso price <= 500 Then
            ret = 0.1
        ElseIf price > 500 AndAlso price <= 1000 Then
            ret = 0.1
        ElseIf price > 1000 Then
            ret = 0.5
        End If
        Return ret
    End Function

    Private Function CalculateLimitBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = _parentStrategy.TickSize
        If price <= 200 Then
            ret = 0.05
        ElseIf price > 200 AndAlso price <= 500 Then
            ret = 0.1
        ElseIf price > 500 AndAlso price <= 1000 Then
            ret = 0.2
        ElseIf price > 1000 Then
            ret = 0.5
        End If
        Return ret
    End Function

    Private Function CalculateStoplossBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = _parentStrategy.TickSize
        If price <= 300 Then
            ret = 0.5
        ElseIf price > 300 AndAlso price <= 1000 Then
            ret = 1
        ElseIf price > 1000 AndAlso price <= 2000 Then
            ret = 2
        ElseIf price > 2000 AndAlso price <= 5000 Then
            ret = 3
        ElseIf price > 5000 Then
            ret = 4
        End If
        Return ret
    End Function
#End Region
End Class