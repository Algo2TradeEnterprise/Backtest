Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL

Public Class NiftyBankMarketPairTradingStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public NumberOfLots As Integer
        Public StockSelectionDetails As List(Of StockSelectionDetails)
    End Class

    Public Class StockData
        Public InstrumentName As String
        Public LotSize As Integer
        Public ChangePer As Decimal
    End Class

    Public Class StockSelectionDetails
        Public StockPosition As PositionOfStock
        Public PositionNumber As Integer
        Public Direction As Trade.TradeExecutionDirection
    End Class

    Enum PositionOfStock
        Top = 1
        Bottom
        Middle1
        Middle2
    End Enum
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Trade.TradeExecutionDirection

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
               ByVal lotSize As Integer,
               ByVal parentStrategy As Strategy,
               ByVal tradingDate As Date,
               ByVal tradingSymbol As String,
               ByVal canceller As CancellationTokenSource,
               ByVal entities As RuleEntities,
               ByVal direction As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
        Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
        _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
        _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
        _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
        _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
        _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
        _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
        _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
        currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = currentMinuteCandlePayload

            If signalCandle IsNot Nothing Then
                If _direction = Trade.TradeExecutionDirection.Buy Then
                    'Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, _userInputs.MaxInvestmentPerStock, signalCandle.Open, _parentStrategy.StockType, True)
                    Dim quantity As Decimal = Me.LotSize * _userInputs.NumberOfLots

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 1000000,
                                .Target = .EntryPrice + 1000000,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market
                            }
                ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                    'Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, _userInputs.MaxInvestmentPerStock, signalCandle.Open, _parentStrategy.StockType, True)
                    Dim quantity As Decimal = Me.LotSize * _userInputs.NumberOfLots

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signalCandle.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice + 1000000,
                                .Target = .EntryPrice - 1000000,
                                .Buffer = 0,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.Market
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
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
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

    Public Shared Function GetStockData(ByVal allStocks As List(Of StockData), ByVal stockSelectionDetails As List(Of StockSelectionDetails)) As Dictionary(Of String, StockDetails)
        Dim ret As Dictionary(Of String, StockDetails) = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 AndAlso stockSelectionDetails IsNot Nothing AndAlso stockSelectionDetails.Count > 0 Then
            For Each runningStockSelection In stockSelectionDetails
                Dim stock As StockData = Nothing
                Select Case runningStockSelection.StockPosition
                    Case PositionOfStock.Top
                        stock = GetNthStockFromTop(allStocks, runningStockSelection.PositionNumber)
                    Case PositionOfStock.Bottom
                        stock = GetNthStockFromBottom(allStocks, runningStockSelection.PositionNumber)
                    Case PositionOfStock.Middle1
                        stock = GetMiddleStock1(allStocks)
                    Case PositionOfStock.Middle2
                        stock = GetMiddleStock2(allStocks)
                End Select
                If stock IsNot Nothing Then
                    If ret Is Nothing Then ret = New Dictionary(Of String, StockDetails)
                    Dim detailsOfStock As StockDetails = New StockDetails With
                                                {.StockName = stock.InstrumentName,
                                                 .LotSize = stock.LotSize,
                                                 .EligibleToTakeTrade = True,
                                                 .Supporting1 = If(runningStockSelection.Direction = Trade.TradeExecutionDirection.Buy, 1, -1)}
                    ret.Add(detailsOfStock.StockName, detailsOfStock)
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetNthStockFromTop(ByVal allStocks As List(Of StockData), ByVal stockNumber As Integer) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            For Each runningStock In allStocks.OrderByDescending(Function(x)
                                                                     Return x.ChangePer
                                                                 End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetNthStockFromBottom(ByVal allStocks As List(Of StockData), ByVal stockNumber As Integer) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            For Each runningStock In allStocks.OrderBy(Function(x)
                                                           Return x.ChangePer
                                                       End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetMiddleStock1(ByVal allStocks As List(Of StockData)) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            Dim stockNumber As Integer = Math.Floor(allStocks.Count / 2)
            For Each runningStock In allStocks.OrderByDescending(Function(x)
                                                                     Return x.ChangePer
                                                                 End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Shared Function GetMiddleStock2(ByVal allStocks As List(Of StockData)) As StockData
        Dim ret As StockData = Nothing
        If allStocks IsNot Nothing AndAlso allStocks.Count > 0 Then
            Dim ctr As Integer = 0
            Dim stockNumber As Integer = Math.Floor(allStocks.Count / 2) + 1
            For Each runningStock In allStocks.OrderByDescending(Function(x)
                                                                     Return x.ChangePer
                                                                 End Function)
                ctr += 1
                If ctr = stockNumber Then
                    ret = runningStock
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
End Class