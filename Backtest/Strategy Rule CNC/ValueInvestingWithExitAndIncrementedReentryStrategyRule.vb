Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class ValueInvestingWithExitAndIncrementedReEntryStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InitialInvestment As Decimal
        Public AmountOfIncreaseDesireEachPeriod As Decimal
    End Class
#End Region

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
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
        If currentDayCandlePayload IsNot Nothing Then
            Dim activeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
            If activeTrades IsNot Nothing AndAlso activeTrades.Count > 0 Then   'Continue Investing
                Dim firstTrade As Trade = activeTrades.OrderBy(Function(x)
                                                                   Return x.EntryTime
                                                               End Function).FirstOrDefault

                Dim lastTrade As Trade = activeTrades.OrderBy(Function(x)
                                                                  Return x.EntryTime
                                                              End Function).LastOrDefault

                Dim noOfSharedOwnedBeforeRebalancing As Integer = activeTrades.Sum(Function(x) x.Quantity)
                Dim totalValueBeforeRebalancing As Decimal = noOfSharedOwnedBeforeRebalancing * currentDayCandlePayload.Close
                Dim numberOfDays As Integer = _signalPayload.Where(Function(x)
                                                                       Return x.Key.Date >= firstTrade.SignalCandle.PayloadDate.Date AndAlso
                                                                       x.Key.Date <= currentDayCandlePayload.PayloadDate.Date
                                                                   End Function).Count
                Dim desiredValue As Decimal = _userInputs.InitialInvestment + lastTrade.Supporting5 + (numberOfDays - 1) * _userInputs.AmountOfIncreaseDesireEachPeriod
                Dim amountToInvest As Decimal = desiredValue - totalValueBeforeRebalancing
                Dim numberOfSharesToBuy As Decimal = Math.Round(amountToInvest / currentDayCandlePayload.Close)

                If numberOfSharesToBuy > 0 Then
                    Dim totalInvestedAmount As Decimal = activeTrades.Sum(Function(x)
                                                                              Return x.EntryPrice * x.Quantity
                                                                          End Function)
                    totalInvestedAmount += currentDayCandlePayload.Close * numberOfSharesToBuy

                    Dim noOfSharedOwnedAfterRebalancing As Integer = noOfSharedOwnedBeforeRebalancing + numberOfSharesToBuy

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = currentDayCandlePayload.Close,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = numberOfSharesToBuy,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = Val(lastTrade.Supporting1) + 1,  'Iteration
                                                                .Supporting2 = lastTrade.Supporting2,           'Tag
                                                                .Supporting3 = totalInvestedAmount,             'Total Amount Invested in this chain
                                                                .Supporting4 = noOfSharedOwnedAfterRebalancing, 'Total Quantity Traded in this chain
                                                                .Supporting5 = lastTrade.Supporting5            'Additional Investment
                                                            }

                    For Each runningTrade In activeTrades
                        runningTrade.UpdateTrade(Supporting3:=totalInvestedAmount, Supporting4:=noOfSharedOwnedAfterRebalancing)
                    Next

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
            Else    'New Investing start
                Dim additionalInvestment As Decimal = GetAdditionalInvestment(currentDayCandlePayload)
                Dim numberOfSharesToBuy As Decimal = Math.Round((_userInputs.InitialInvestment + additionalInvestment) / currentDayCandlePayload.Close)

                If numberOfSharesToBuy > 0 Then
                    Dim totalInvestedAmount As Decimal = currentDayCandlePayload.Close * numberOfSharesToBuy

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                .EntryPrice = currentDayCandlePayload.Close,
                                                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                .Quantity = numberOfSharesToBuy,
                                                                .Stoploss = .EntryPrice - 1000000000,
                                                                .Target = .EntryPrice + 1000000000,
                                                                .Buffer = 0,
                                                                .SignalCandle = currentDayCandlePayload,
                                                                .OrderType = Trade.TypeOfOrder.Market,
                                                                .Supporting1 = 1,                                  'Iteration
                                                                .Supporting2 = System.Guid.NewGuid.ToString(),     'Tag
                                                                .Supporting3 = totalInvestedAmount,                'Total Amount Invested in this chain
                                                                .Supporting4 = numberOfSharesToBuy,                'Total Quantity Traded in this chain
                                                                .Supporting5 = additionalInvestment                'Additional Investment
                                                            }

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
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
            Dim currentDayCandlePayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            Dim totalQuantityTradedInThisChain As Integer = currentTrade.Supporting4
            Dim totalAmountInvestedInThisChain As Decimal = currentTrade.Supporting3
            Dim totalAmountReturnInThisChain As Decimal = totalQuantityTradedInThisChain * currentDayCandlePayload.Close
            Dim returnPercentage As Decimal = Math.Round((totalAmountReturnInThisChain / totalAmountInvestedInThisChain - 1) * 100, 2)
            If returnPercentage >= 5 Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, currentDayCandlePayload.Close, "Target")
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetAdditionalInvestment(ByVal currentDayCandlePayload As Payload) As Decimal
        Dim ret As Decimal = 0
        Dim allTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentDayCandlePayload, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Close)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim tradeIDs As List(Of String) = New List(Of String)
            For Each runningTrade In allTrades
                If Not tradeIDs.Contains(runningTrade.Supporting2) Then
                    tradeIDs.Add(runningTrade.Supporting2)
                End If
            Next
            If tradeIDs IsNot Nothing AndAlso tradeIDs.Count >= 2 Then
                Dim tradeGroupData As List(Of TradeGroupDetails) = New List(Of TradeGroupDetails)
                For Each runningID In tradeIDs
                    Dim groupTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                              Return x.Supporting2 = runningID
                                                                          End Function)
                    If groupTrades IsNot Nothing AndAlso groupTrades.Count > 1 Then
                        Dim tradeData As TradeGroupDetails = New TradeGroupDetails
                        tradeData.TradeID = runningID
                        tradeData.NumberOfOngoingIterations = groupTrades.Count - 1
                        tradeData.TotalInvestmentOfOngoingIterations = groupTrades.Sum(Function(x)
                                                                                           If x.Supporting1 > 1 Then
                                                                                               Return x.CapitalRequiredWithMargin
                                                                                           Else
                                                                                               Return 0
                                                                                           End If
                                                                                       End Function)
                        tradeGroupData.Add(tradeData)
                    End If
                Next
                If tradeGroupData IsNot Nothing AndAlso tradeGroupData.Count > 0 Then
                    Dim maxTotalInvestment As Decimal = tradeGroupData.Max(Function(x) x.TotalInvestmentOfOngoingIterations)
                    Dim avgIteration As Decimal = tradeGroupData.Average(Function(x) x.NumberOfOngoingIterations)
                    Dim avgAvgIteration As Decimal = tradeGroupData.Average(Function(x) x.AverageInvestmentOfOngoingIterations)

                    ret = maxTotalInvestment - avgIteration * avgAvgIteration
                End If
            End If
        End If
        Return ret
    End Function

    Private Class TradeGroupDetails
        Public Property TradeID As String
        Public Property NumberOfOngoingIterations As Integer
        Public Property TotalInvestmentOfOngoingIterations As Decimal
        Public ReadOnly Property AverageInvestmentOfOngoingIterations As Decimal
            Get
                Return TotalInvestmentOfOngoingIterations / NumberOfOngoingIterations
            End Get
        End Property
    End Class
End Class