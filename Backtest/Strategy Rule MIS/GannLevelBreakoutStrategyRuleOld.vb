Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class GannLevelBreakoutStrategyRuleOld
    Inherits StrategyRule

#Region "Entity"
    Enum StrategyType
        OnlyTwoOppositeDirectionTradesWithFib = 1
        InfiniteTradesWithFib
        InfiniteTradesWithoutFib
        OnlyTwoOppositeDirectionTradesWithoutFib
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxDiffPer As Decimal
        Public TypeOfStrategy As StrategyType
    End Class
#End Region

    Private _buyLevel As Decimal = Decimal.MinValue
    Private _buySLLevel As Decimal = Decimal.MinValue
    Private _buyRemarks As String = Nothing
    Private _sellLevel As Decimal = Decimal.MinValue
    Private _sellSLLevel As Decimal = Decimal.MinValue
    Private _sellRemarks As String = Nothing

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing

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

        Dim firstCandleOfTheDay As Payload = _signalPayload.Where(Function(x)
                                                                      Return x.Key.Date = _tradingDate.Date
                                                                  End Function).FirstOrDefault.Value
        Dim currentDayOpen As Decimal = firstCandleOfTheDay.Open
        Dim previousDayHigh As Decimal = _signalPayload.Max(Function(x)
                                                                If x.Key.Date = firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date Then
                                                                    Return x.Value.High
                                                                Else
                                                                    Return Decimal.MinValue
                                                                End If
                                                            End Function)
        Dim previousDayLow As Decimal = _signalPayload.Min(Function(x)
                                                               If x.Key.Date = firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date Then
                                                                   Return x.Value.Low
                                                               Else
                                                                   Return Decimal.MaxValue
                                                               End If
                                                           End Function)

        Dim gann As GannLevels = Common.CalculateGann(currentDayOpen)
        If _userInputs.TypeOfStrategy = StrategyType.InfiniteTradesWithoutFib OrElse
            _userInputs.TypeOfStrategy = StrategyType.OnlyTwoOppositeDirectionTradesWithoutFib Then
            _buyLevel = ConvertFloorCeling(gann.BuyAt, _parentStrategy.TickSize, RoundOfType.Celing)
            _buyRemarks = String.Format("Buy At:{0}", Math.Round(gann.BuyAt, 2))
            _sellLevel = ConvertFloorCeling(gann.SellAt, _parentStrategy.TickSize, RoundOfType.Floor)
            _sellRemarks = String.Format("Sell At:{0}", Math.Round(gann.SellAt, 2))
            _buySLLevel = _sellLevel
            _sellSLLevel = _buyLevel
        Else
            Dim fib As FibonacciLevels = Common.GetFibonacciLevels(previousDayHigh, previousDayLow)
            Dim buyData As Tuple(Of Decimal, Decimal) = GetMinDiff(gann.BuyAt, fib)
            Dim sellData As Tuple(Of Decimal, Decimal) = GetMinDiff(gann.SellAt, fib)
            If buyData IsNot Nothing AndAlso buyData.Item2 <= _userInputs.MaxDiffPer Then
                _buyLevel = ConvertFloorCeling(gann.BuyAt, _parentStrategy.TickSize, RoundOfType.Celing)
                _buyRemarks = String.Format("Buy At:{0}, Fib Level:{1}%, Diff:{2}%",
                                            Math.Round(gann.BuyAt, 2), buyData.Item1, Math.Round(buyData.Item2, 2))
                _buySLLevel = ConvertFloorCeling(gann.SellAt, _parentStrategy.TickSize, RoundOfType.Floor)
            Else
                _buyRemarks = String.Format("Buy At:{0}, Fib Level:{1}, Diff:{2}",
                                            Math.Round(gann.BuyAt, 2), "None", "Not Applicable")
            End If
            If sellData IsNot Nothing AndAlso sellData.Item2 <= _userInputs.MaxDiffPer Then
                _sellLevel = ConvertFloorCeling(gann.SellAt, _parentStrategy.TickSize, RoundOfType.Floor)
                _sellRemarks = String.Format("Sell At:{0}, Fib Level:{1}%, Diff:{2}%",
                                            Math.Round(gann.SellAt, 2), sellData.Item1, Math.Round(sellData.Item2, 2))
                _sellSLLevel = ConvertFloorCeling(gann.BuyAt, _parentStrategy.TickSize, RoundOfType.Celing)
            Else
                _sellRemarks = String.Format("Sell At:{0}, Fib Level:{1}, Diff:{2}",
                                            Math.Round(gann.SellAt, 2), "None", "Not Applicable")
            End If
        End If

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload, True)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso currentCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            If _buyLevel <> Decimal.MinValue AndAlso
                Not _parentStrategy.IsTradeOpen(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                Not _parentStrategy.IsTradeActive(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                Dim takeTrade As Boolean = False
                If _userInputs.TypeOfStrategy = StrategyType.OnlyTwoOppositeDirectionTradesWithFib OrElse
                    _userInputs.TypeOfStrategy = StrategyType.OnlyTwoOppositeDirectionTradesWithoutFib Then
                    If _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Is Nothing Then
                        takeTrade = True
                    End If
                Else
                    Dim lastOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                    If lastOrder IsNot Nothing Then
                        If lastOrder.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                            If currentCandle.PreviousCandlePayload.Close <= _buyLevel Then
                                takeTrade = True
                            Else
                                Console.WriteLine(String.Format("{0} -> BUY neglect at {1}", _tradingSymbol, currentCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")))
                            End If
                        Else
                            takeTrade = True
                        End If
                    Else
                        takeTrade = True
                    End If
                End If
                If takeTrade Then
                    Dim entryPrice As Decimal = _buyLevel
                    Dim quantity As Integer = Me.LotSize
                    If currentTick.Open < entryPrice Then
                        Dim parameter = New PlaceOrderParameters With {
                                        .entryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                        .quantity = quantity,
                                        .Stoploss = _buySLLevel,
                                        .Target = .EntryPrice + 1000000000,
                                        .Buffer = 0,
                                        .SignalCandle = currentCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = _buyRemarks,
                                        .Supporting2 = _sellRemarks
                                    }

                        If ret Is Nothing Then
                            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                        Else
                            ret.Item2.Add(parameter)
                        End If
                    End If
                End If
            End If
            If _sellLevel <> Decimal.MinValue AndAlso
                Not _parentStrategy.IsTradeOpen(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                Not _parentStrategy.IsTradeActive(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                Dim takeTrade As Boolean = False
                If _userInputs.TypeOfStrategy = StrategyType.OnlyTwoOppositeDirectionTradesWithFib OrElse
                    _userInputs.TypeOfStrategy = StrategyType.OnlyTwoOppositeDirectionTradesWithoutFib Then
                    If _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Is Nothing Then
                        takeTrade = True
                    End If
                Else
                    Dim lastOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentCandle, Trade.TypeOfTrade.MIS)
                    If lastOrder IsNot Nothing Then
                        If lastOrder.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                            If currentCandle.PreviousCandlePayload.Close >= _sellLevel Then
                                takeTrade = True
                            Else
                                Console.WriteLine(String.Format("{0} -> SELL neglect at {1}", _tradingSymbol, currentCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")))
                            End If
                        Else
                            takeTrade = True
                        End If
                    Else
                        takeTrade = True
                    End If
                End If
                If takeTrade Then
                    Dim entryPrice As Decimal = _sellLevel
                    Dim quantity As Integer = Me.LotSize
                    If currentTick.Open > entryPrice Then
                        Dim parameter = New PlaceOrderParameters With {
                                        .entryPrice = entryPrice,
                                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                        .quantity = quantity,
                                        .Stoploss = _sellSLLevel,
                                        .Target = .EntryPrice - 1000000000,
                                        .Buffer = 0,
                                        .SignalCandle = currentCandle,
                                        .OrderType = Trade.TypeOfOrder.SL,
                                        .Supporting1 = _buyRemarks,
                                        .Supporting2 = _sellRemarks
                                    }

                        If ret Is Nothing Then
                            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                        Else
                            ret.Item2.Add(parameter)
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

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentCandle.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
            Dim swingHL As Indicator.Swing = _swingPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim remark As String = Nothing
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If swingHL.SwingLowTime >= currentTrade.EntryTime Then
                    Dim potentialSL As Decimal = swingHL.SwingLow - atr
                    If potentialSL > currentTrade.PotentialStopLoss Then
                        triggerPrice = potentialSL
                        remark = String.Format("Trailling to {0}[Swing({1})-ATR({2})], Swing at {3}", triggerPrice, swingHL.SwingLow, atr, swingHL.SwingLowTime.ToString("HH:mm:ss"))
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If swingHL.SwingHighTime >= currentTrade.EntryTime Then
                    Dim potentialSL As Decimal = swingHL.SwingHigh + atr
                    If potentialSL < currentTrade.PotentialStopLoss Then
                        triggerPrice = potentialSL
                        remark = String.Format("Trailling to {0}[Swing({1})+ATR({2})], Swing at {3}", triggerPrice, swingHL.SwingHigh, atr, swingHL.SwingHighTime.ToString("HH:mm:ss"))
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

    Private Function GetMinDiff(ByVal price As Decimal, ByVal fib As FibonacciLevels) As Tuple(Of Decimal, Decimal)
        Dim ret As Tuple(Of Decimal, Decimal) = Nothing
        Dim diff0 As Decimal = (Math.Abs(fib.Level0 - price) / price) * 100
        Dim diff38 As Decimal = (Math.Abs(fib.Level38 - price) / price) * 100
        Dim diff50 As Decimal = (Math.Abs(fib.Level50 - price) / price) * 100
        Dim diff61 As Decimal = (Math.Abs(fib.Level61 - price) / price) * 100
        Dim diff100 As Decimal = (Math.Abs(fib.Level100 - price) / price) * 100
        Dim diff_38 As Decimal = (Math.Abs(fib.Level_38 - price) / price) * 100
        Dim diff_61 As Decimal = (Math.Abs(fib.Level_61 - price) / price) * 100
        Dim diff138 As Decimal = (Math.Abs(fib.Level138 - price) / price) * 100
        Dim diff161 As Decimal = (Math.Abs(fib.Level161 - price) / price) * 100
        If diff0 < Math.Min(diff38, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff_61, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(0, diff0)
        ElseIf diff38 < Math.Min(diff0, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff_61, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(38.2, diff38)
        ElseIf diff61 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff_61, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(61.8, diff61)
        ElseIf diff100 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff61, Math.Min(diff_38, Math.Min(diff_61, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(100, diff100)
        ElseIf diff_38 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_61, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(-38.2, diff_38)
        ElseIf diff_61 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff138, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(-61.8, diff_61)
        ElseIf diff138 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff_61, diff161)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(138.2, diff138)
        ElseIf diff161 < Math.Min(diff0, Math.Min(diff38, Math.Min(diff61, Math.Min(diff100, Math.Min(diff_38, Math.Min(diff_61, diff138)))))) Then
            ret = New Tuple(Of Decimal, Decimal)(161.8, diff161)
        End If
        Return ret
    End Function
End Class