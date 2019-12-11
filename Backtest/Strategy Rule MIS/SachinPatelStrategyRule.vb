Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SachinPatelStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public InvestmentPerStock As Decimal
        Public MaxAllowedLossPercentagePerStock As Decimal
        Public MaxStoplossPercentagePerTrade As Decimal
        Public MaxTargetPercentagePerTrade As Decimal
    End Class
#End Region

    Private _userInputs As StrategyRuleEntities
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _adxPayload As Dictionary(Of Date, Decimal)
    Private _diPlusPayload As Dictionary(Of Date, Decimal)
    Private _diMinusPayload As Dictionary(Of Date, Decimal)

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
        Indicator.ADX.CalculateADX(14, 14, _signalPayload, _adxPayload, _diPlusPayload, _diMinusPayload, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
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

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = CalculateBuffer(signalCandle.High)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim quantity As Decimal = CalculateQuantity(entryPrice)
                    Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.MaxStoplossPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    Dim targetPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.MaxTargetPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = entryPrice,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = quantity,
                        .Stoploss = .EntryPrice - slPoint,
                        .Target = .EntryPrice + targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL
                    }
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = CalculateBuffer(signalCandle.Low)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim quantity As Decimal = CalculateQuantity(entryPrice)
                    Dim slPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.MaxStoplossPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    Dim targetPoint As Decimal = ConvertFloorCeling(entryPrice * _userInputs.MaxTargetPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = entryPrice,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = quantity,
                        .Stoploss = .EntryPrice + slPoint,
                        .Target = .EntryPrice - targetPoint,
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL
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
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open < signalCandle.Low Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open > signalCandle.High Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = CDec(currentTrade.Supporting4)
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                triggerPrice = currentTrade.EntryPrice - slPoint
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                triggerPrice = currentTrade.EntryPrice + slPoint
            End If
            If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, slPoint)
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

    Public Async Function UpdateRequiedSupporting(currentTrade As Trade) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Close Then
            Dim exitCandleTime As Date = Me._parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.ExitTime, _signalPayload)
            currentTrade.UpdateTrade(Supporting1:=_adxPayload(currentTrade.SignalCandle.PayloadDate),
                                    Supporting2:=_diPlusPayload(currentTrade.SignalCandle.PayloadDate),
                                    Supporting3:=_diMinusPayload(currentTrade.SignalCandle.PayloadDate),
                                    Supporting4:=_adxPayload(exitCandleTime),
                                    Supporting5:=_diPlusPayload(exitCandleTime),
                                    Supporting6:=_diMinusPayload(exitCandleTime))
        End If
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim atr As Decimal = _atrPayload(candle.PayloadDate)
            Dim fractalHigh As Decimal = _fractalHighPayload(candle.PayloadDate)
            Dim fractalLow As Decimal = _fractalLowPayload(candle.PayloadDate)
            Dim adx As Decimal = _fractalLowPayload(candle.PayloadDate)
            Dim diPlus As Decimal = _diPlusPayload(candle.PayloadDate)
            Dim diMinus As Decimal = _diMinusPayload(candle.PayloadDate)

        End If
        Return ret
    End Function

    Private Function CalculateQuantity(ByVal stockPrice As Decimal) As Integer
        Dim ret As Integer = Integer.MinValue
        Dim maxAllowedLossOfTheStock As Decimal = _userInputs.InvestmentPerStock * _userInputs.MaxAllowedLossPercentagePerStock / 100
        Dim maxLossPerTrade As Decimal = ConvertFloorCeling(stockPrice * _userInputs.MaxStoplossPercentagePerTrade / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
        ret = Math.Floor(maxAllowedLossOfTheStock / maxLossPerTrade)
        Return ret
    End Function

    Private Function CalculateBuffer(ByVal price As Decimal) As Decimal
        Dim ret As Decimal = 0.05
        If price < 200 Then
            ret = 0.05
        ElseIf price >= 200 AndAlso price < 400 Then
            ret = 0.1
        ElseIf price >= 400 AndAlso price < 700 Then
            ret = 0.2
        ElseIf price >= 700 AndAlso price < 1000 Then
            ret = 0.25
        ElseIf price >= 1000 AndAlso price < 1500 Then
            ret = 0.3
        ElseIf price >= 1500 AndAlso price < 3000 Then
            ret = 0.4
        ElseIf price >= 3000 Then
            ret = 0.5
        End If
        Return ret
    End Function
End Class