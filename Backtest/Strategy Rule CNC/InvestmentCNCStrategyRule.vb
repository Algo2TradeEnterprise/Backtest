Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class InvestmentCNCStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Linear = 1
        GP
        AP
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public QuantityType As TypeOfQuantity
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _smaPayload As Dictionary(Of Date, Decimal)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _stockSMAPercentage As Decimal
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockSMAPer As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _stockSMAPercentage = stockSMAPer
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
        Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _signalPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso GetAboveSMAPercentage() > 90 Then
            Dim signalCandle As Payload = Nothing

            Dim signalReceivedForEntry As Tuple(Of Boolean, Decimal, Integer) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                Dim firstQuantity As Integer = 0
                Dim firstEntryDate As Date = currentTick.PayloadDate
                Dim quantity As Integer = signalReceivedForEntry.Item3
                firstQuantity = quantity
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                If lastExecutedTrade IsNot Nothing Then
                    If _userInputs.QuantityType = TypeOfQuantity.GP Then
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            quantity = lastExecutedTrade.Quantity * 2
                        End If
                    ElseIf _userInputs.QuantityType = TypeOfQuantity.AP Then
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            quantity = lastExecutedTrade.Quantity + CInt(lastExecutedTrade.Supporting1)
                            firstQuantity = lastExecutedTrade.Supporting1
                        End If
                    ElseIf _userInputs.QuantityType = TypeOfQuantity.Linear Then
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            quantity = lastExecutedTrade.Quantity
                        End If
                    End If
                    If lastExecutedTrade.EntryTime.Date <> _tradingDate.Date Then
                        signalCandle = currentMinuteCandlePayload
                    End If
                    If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                        firstEntryDate = Date.Parse(lastExecutedTrade.Supporting2)
                    End If
                Else
                    signalCandle = currentMinuteCandlePayload
                End If

                If signalCandle IsNot Nothing Then
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalReceivedForEntry.Item2,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = quantity,
                        .Stoploss = .EntryPrice - 1000000,
                        .Target = .EntryPrice + 1000000,
                        .Buffer = 0,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.Market,
                        .Supporting1 = firstQuantity,
                        .Supporting2 = firstEntryDate.ToString("yyyy-MM-dd")
                    }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
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
            Dim signalReceivedForExit As Tuple(Of Boolean, Decimal, String) = GetSignalForExit(currentTick, currentTrade)
            If signalReceivedForExit IsNot Nothing AndAlso signalReceivedForExit.Item1 Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, signalReceivedForExit.Item2, signalReceivedForExit.Item3)
            End If
        End If
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

    Private Function GetAboveSMAPercentage() As Decimal
        Dim ret As Decimal = Decimal.MinValue

        Dim avgChkFrom As Date = _tradingDate.AddYears(-2)
        Dim totalDayCount As Integer = 0
        Dim aboveSMACount As Integer = 0
        For Each runningPayload In _signalPayload.Values
            _cts.Token.ThrowIfCancellationRequested()
            If runningPayload.PayloadDate.Date >= avgChkFrom.Date AndAlso
                runningPayload.PayloadDate.Date <= _tradingDate.Date Then
                totalDayCount += 1
                If runningPayload.Close > _smaPayload(runningPayload.PayloadDate) Then
                    aboveSMACount += 1
                End If
            End If
        Next
        If totalDayCount <> 0 Then
            Dim aboveSMA200Avg As Decimal = Math.Round((aboveSMACount / totalDayCount) * 100, 2)
            ret = aboveSMA200Avg
        End If
        Return ret
    End Function

#Region "Entry Rule"

#Region "Fractal Low Entry"
    'Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer)
    '    Dim ret As Tuple(Of Boolean, Decimal, Integer) = Nothing
    '    If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
    '        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
    '        Dim fractalLow As Decimal = _fractalLowPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim fractalHigh As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim sma As Decimal = _smaPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim entryPoint As Decimal = fractalLow

    '        If entryPoint > 100 AndAlso (fractalHigh >= sma OrElse fractalLow >= sma) Then
    '            If currentDayPayload.Open >= entryPoint Then
    '                If currentDayPayload.Low <= entryPoint Then
    '                    Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
    '                    ret = New Tuple(Of Boolean, Decimal, Integer)(True, entryPoint, quantity)
    '                End If
    '            Else
    '                Dim tradeEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 15, 25, 0)
    '                If currentDayPayload.CandleColor = Color.Red Then
    '                    Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
    '                    Dim r As New Random()
    '                    Dim pricePercentage As Integer = r.Next(1, 5)
    '                    Dim direction As Integer = r.Next(0, 1)
    '                    Dim price As Decimal = Decimal.MinValue
    '                    If currentDayPayload.Close = currentDayPayload.Low Then
    '                        price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                    Else
    '                        If direction = 0 Then
    '                            price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                        ElseIf direction = 1 Then
    '                            price = currentDayPayload.Close - ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                        End If
    '                    End If
    '                    ret = New Tuple(Of Boolean, Decimal, Integer)(True, price, quantity)
    '                End If
    '            End If
    '        End If
    '    End If
    '    Return ret
    'End Function
#End Region

#Region "Fractal Low Outside Entry"
    'Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer)
    '    Dim ret As Tuple(Of Boolean, Decimal, Integer) = Nothing
    '    If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
    '        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
    '        Dim fractalLow As Decimal = _fractalLowPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim fractalHigh As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim sma As Decimal = _smaPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        Dim entryPoint As Decimal = fractalLow

    '        If entryPoint > 100 AndAlso (fractalHigh >= sma OrElse fractalLow >= sma) Then
    '            If currentDayPayload.High < fractalLow Then
    '                Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
    '                Dim r As New Random()
    '                Dim pricePercentage As Integer = r.Next(1, 5)
    '                Dim direction As Integer = r.Next(0, 1)
    '                Dim price As Decimal = Decimal.MinValue
    '                If currentDayPayload.Close = currentDayPayload.Low Then
    '                    price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                Else
    '                    If direction = 0 Then
    '                        price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                    ElseIf direction = 1 Then
    '                        price = currentDayPayload.Close - ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
    '                    End If
    '                End If
    '                ret = New Tuple(Of Boolean, Decimal, Integer)(True, price, quantity)
    '            End If
    '        End If
    '    End If
    '    Return ret
    'End Function
#End Region

#Region "Fractal Low Outside Entry"
    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Integer)
        Dim ret As Tuple(Of Boolean, Decimal, Integer) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            Dim fractalLow As Decimal = _fractalLowPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim fractalHigh As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim sma As Decimal = _smaPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim entryPoint As Decimal = fractalLow

            If entryPoint > 100 AndAlso (fractalHigh >= sma OrElse fractalLow >= sma) Then
                If currentDayPayload.High < fractalLow Then
                    Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
                    Dim r As New Random()
                    Dim pricePercentage As Integer = r.Next(1, 5)
                    Dim direction As Integer = r.Next(0, 1)
                    Dim price As Decimal = Decimal.MinValue
                    If currentDayPayload.Close = currentDayPayload.Low Then
                        price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    ElseIf currentDayPayload.Close = currentDayPayload.High Then
                        price = currentDayPayload.Close - ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
                    Else
                        If direction = 0 Then
                            price = currentDayPayload.Close + ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
                        ElseIf direction = 1 Then
                            price = currentDayPayload.Close - ConvertFloorCeling(currentDayPayload.Close * pricePercentage / 1000, Me._parentStrategy.TickSize, RoundOfType.Celing)
                        End If
                    End If
                    If price <= currentDayPayload.PreviousCandlePayload.Low Then
                        ret = New Tuple(Of Boolean, Decimal, Integer)(True, price, quantity)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
#End Region

#End Region

#Region "Exit Rule"

#Region "Fractal High Exit"
    'Private Function GetSignalForExit(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Tuple(Of Boolean, Decimal, String)
    '    Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
    '    If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
    '        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
    '        Dim exitPoint As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
    '        If currentDayPayload.High >= exitPoint Then
    '            ret = New Tuple(Of Boolean, Decimal, String)(True, exitPoint, "Target Hit")
    '        End If
    '    End If
    '    Return ret
    'End Function
#End Region

#Region "Upper High Fractal Exit"
    'Private Function GetSignalForExit(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Tuple(Of Boolean, Decimal, String)
    '    Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
    '    If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
    '        Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
    '        Dim exitPoint As Decimal = GetUpperHighFractal(Date.Parse(currentTrade.Supporting2), currentTick.PayloadDate.Date)
    '        If exitPoint <> Decimal.MinValue AndAlso currentDayPayload.High >= exitPoint Then
    '            ret = New Tuple(Of Boolean, Decimal, String)(True, exitPoint, "Target Hit")
    '        End If
    '    End If
    '    Return ret
    'End Function

    'Private Function GetUpperHighFractal(ByVal startDate As Date, ByVal endDate As Date) As Decimal
    '    Dim ret As Decimal = Decimal.MinValue
    '    If _fractalHighPayload IsNot Nothing AndAlso _fractalHighPayload.Count > 0 Then
    '        Dim previousFractal As Decimal = Decimal.MinValue
    '        For Each runningDate In _fractalHighPayload.Keys
    '            If runningDate.Date > startDate.Date AndAlso runningDate.Date < endDate.Date Then
    '                If previousFractal <> Decimal.MinValue AndAlso _fractalHighPayload(runningDate) > previousFractal Then
    '                    ret = _fractalHighPayload(runningDate)
    '                End If
    '                previousFractal = _fractalHighPayload(runningDate)
    '            End If
    '        Next
    '    End If
    '    Return ret
    'End Function
#End Region

#Region "15% Exit"
    Private Function GetSignalForExit(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Tuple(Of Boolean, Decimal, String)
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 AndAlso _signalPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _signalPayload(currentTick.PayloadDate.Date)
            Dim exitPoint As Decimal = Decimal.MinValue
            Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentDayPayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
            If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                Dim allRelatedTrades As List(Of Trade) = New List(Of Trade)
                allRelatedTrades.AddRange(openActiveTrades)
                Dim closedTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentDayPayload, Me._parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
                If closedTrades IsNot Nothing AndAlso closedTrades.Count > 0 Then
                    Dim relatedTrades As List(Of Trade) = closedTrades.FindAll(Function(x)
                                                                                   Return x.Supporting2 = currentTrade.Supporting2
                                                                               End Function)
                    If relatedTrades IsNot Nothing AndAlso relatedTrades.Count > 0 Then
                        allRelatedTrades.AddRange(relatedTrades)
                    End If
                End If
                If allRelatedTrades IsNot Nothing AndAlso allRelatedTrades.Count > 0 Then
                    Dim totalCapitalUsedWithoutMargin As Decimal = allRelatedTrades.Sum(Function(x)
                                                                                            Return x.EntryPrice * x.Quantity
                                                                                        End Function)
                    Dim totalQuantity As Decimal = allRelatedTrades.Sum(Function(x)
                                                                            Return x.Quantity
                                                                        End Function)
                    Dim averageTradePrice As Decimal = totalCapitalUsedWithoutMargin / totalQuantity
                    exitPoint = ConvertFloorCeling(averageTradePrice + averageTradePrice * 15 / 100, Me._parentStrategy.TickSize, RoundOfType.Floor)
                End If
            End If
            If exitPoint <> Decimal.MinValue AndAlso currentDayPayload.High >= exitPoint Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, exitPoint, "Target Hit")
            End If
        End If
        Return ret
    End Function
#End Region

#End Region
End Class
