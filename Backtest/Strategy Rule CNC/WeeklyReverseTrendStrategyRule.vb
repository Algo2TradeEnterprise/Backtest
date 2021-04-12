Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class WeeklyReverseTrendStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetPercentage As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _weeklyPayload As Dictionary(Of Date, Payload)
    Private _lastTradingDayOfTheWeek As Date = Date.MinValue

    Private _macdLinePayload As Dictionary(Of Date, Decimal)
    Private _macdSignalLinePayload As Dictionary(Of Date, Decimal)


    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        _weeklyPayload = Common.ConvertDayPayloadsToWeek(_signalPayload)
        Dim startDateOfTheWeek As Date = Common.GetStartDateOfTheWeek(_tradingDate.Date, DayOfWeek.Monday)
        Dim endDateOfTheWeek As Date = Common.GetEndDateOfTheWeek(_tradingDate.Date, DayOfWeek.Monday)
        _lastTradingDayOfTheWeek = _signalPayload.Where(Function(x)
                                                            Return x.Key.Date >= startDateOfTheWeek.Date AndAlso x.Key.Date <= endDateOfTheWeek.Date
                                                        End Function).OrderBy(Function(y)
                                                                                  Return y.Key
                                                                              End Function).LastOrDefault.Key

        Indicator.MACD.CalculateMACD(12, 26, 9, _weeklyPayload, _macdLinePayload, _macdSignalLinePayload, Nothing)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentWeekCandle As Payload = _weeklyPayload(Common.GetStartDateOfTheWeek(currentTick.PayloadDate.Date, DayOfWeek.Monday))
        Dim currentDayCandle As Payload = _signalPayload(currentTick.PayloadDate.Date)
        If currentWeekCandle IsNot Nothing AndAlso currentDayCandle IsNot Nothing AndAlso _lastTradingDayOfTheWeek <> Date.MinValue AndAlso
            Not _parentStrategy.IsTradeActive(currentWeekCandle, Trade.TypeOfTrade.CNC) Then
            Dim signalCandle As Payload = Nothing
            If currentWeekCandle.PreviousCandlePayload IsNot Nothing AndAlso currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
                currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                If currentWeekCandle.PreviousCandlePayload.High > currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                    currentWeekCandle.PreviousCandlePayload.Low > currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                    If currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.High < currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                        currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.Low < currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
                        If currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High < currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                            currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low < currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.PreviousCandlePayload.Low Then
                            If _macdLinePayload(currentWeekCandle.PreviousCandlePayload.PayloadDate) > _macdSignalLinePayload(currentWeekCandle.PreviousCandlePayload.PayloadDate) Then
                                Dim lastTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentWeekCandle, Trade.TypeOfTrade.CNC)
                                If lastTrade Is Nothing OrElse lastTrade.ExitTime < currentWeekCandle.PayloadDate Then
                                    signalCandle = currentWeekCandle.PreviousCandlePayload
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If signalCandle IsNot Nothing AndAlso currentDayCandle.High >= signalCandle.High + _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor) Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Close, RoundOfType.Floor)

                Dim entryPrice As Decimal = signalCandle.High + buffer
                Dim stoplossPrice As Decimal = signalCandle.Low - buffer
                Dim targetPrice As Decimal = entryPrice + ConvertFloorCeling(entryPrice * _userInputs.TargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Celing)

                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromSL(_tradingSymbol, entryPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)
                If currentDayCandle.Open > entryPrice Then entryPrice = currentDayCandle.Open

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = entryPrice,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = quantity,
                                                                .Stoploss = stoplossPrice,
                                                                .Target = targetPrice,
                                                                .Buffer = buffer,
                                                                .SignalCandle = signalCandle,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = signalCandle.PayloadDate.ToString("dd-MMM-yyyy"),
                                                                .Supporting2 = _macdLinePayload(signalCandle.PayloadDate),
                                                                .Supporting3 = _macdSignalLinePayload(signalCandle.PayloadDate)
                                                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentDayCandle As Payload = _signalPayload(currentTick.PayloadDate.Date)
            If currentDayCandle.High >= currentTrade.PotentialTarget AndAlso currentDayCandle.Low <= currentTrade.PotentialStopLoss Then
                If currentDayCandle.CandleColor = Color.Red Then
                    Dim priceToExit As Decimal = currentTrade.PotentialTarget
                    If currentDayCandle.Open > priceToExit Then priceToExit = currentDayCandle.Open

                    ret = New Tuple(Of Boolean, Decimal, String)(True, priceToExit, "Target")
                Else
                    Dim priceToExit As Decimal = currentTrade.PotentialStopLoss
                    If currentDayCandle.Open < priceToExit Then priceToExit = currentDayCandle.Open

                    ret = New Tuple(Of Boolean, Decimal, String)(True, priceToExit, "Stoploos")
                End If
            ElseIf currentDayCandle.High >= currentTrade.PotentialTarget Then
                Dim priceToExit As Decimal = currentTrade.PotentialTarget
                If currentDayCandle.Open > priceToExit Then priceToExit = currentDayCandle.Open

                ret = New Tuple(Of Boolean, Decimal, String)(True, priceToExit, "Target")
            ElseIf currentDayCandle.Low <= currentTrade.PotentialStopLoss Then
                Dim priceToExit As Decimal = currentTrade.PotentialStopLoss
                If currentDayCandle.Open < priceToExit Then priceToExit = currentDayCandle.Open

                ret = New Tuple(Of Boolean, Decimal, String)(True, priceToExit, "Stoploos")
            Else
                If currentTick.PayloadDate.Date = _lastTradingDayOfTheWeek.Date Then
                    Dim startDateOfTheWeek As Date = Common.GetStartDateOfTheWeek(currentTrade.EntryTime, DayOfWeek.Monday)
                    If currentTick.PayloadDate.Date > startDateOfTheWeek.AddDays(8) Then
                        Dim priceToExit As Decimal = currentDayCandle.Close

                        ret = New Tuple(Of Boolean, Decimal, String)(True, priceToExit, "EOC Exit")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentWeekCandle As Payload = _weeklyPayload(Common.GetStartDateOfTheWeek(currentTick.PayloadDate.Date, DayOfWeek.Monday))
            If currentWeekCandle.PreviousCandlePayload IsNot Nothing AndAlso currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                If currentWeekCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate = currentTrade.SignalCandle.PayloadDate Then
                    Dim potentialStoploss As Decimal = currentWeekCandle.PreviousCandlePayload.Low - currentTrade.StoplossBuffer
                    If currentTrade.PotentialStopLoss <> potentialStoploss Then
                        ret = New Tuple(Of Boolean, Decimal, String)(True, potentialStoploss, "Moved")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function
End Class