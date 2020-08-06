Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayHKTrendGapStrategyRule
    Inherits MathematicalStrategyRule

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _previousDayHighestATR As Decimal = Decimal.MinValue

    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)

        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)

        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                _previousDayHighestATR = _atrPayload.Max(Function(x)
                                                             If x.Key.Date >= firstCandle.PreviousCandlePayload.PayloadDate.Date AndAlso
                                                             x.Key.Date < firstCandle.PayloadDate.Date Then
                                                                 Return x.Value
                                                             Else
                                                                 Return Decimal.MinValue
                                                             End If
                                                         End Function)
            End If
        End If
    End Sub

    Protected Overrides Function GetEntrySignal(currentCandle As Payload, currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)
        Dim ret As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If _direction = Trade.TradeExecutionDirection.Buy Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                Dim highestPoint As Decimal = GetHighestHigh(currentCandle.PayloadDate)
                Dim entryPrice As Decimal = highestPoint + buffer
                Dim targetPoint As Decimal = ConvertFloorCeling(GetAverageHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, 500, Trade.TypeOfStock.Cash)
                If currentCandle.Open > entryPrice Then
                    entryPrice = currentTick.Open
                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice + targetPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.Market, "Highest Point Gap")
                Else
                    If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                        If currentCandle.PreviousCandlePayload.High < currentCandle.PreviousCandlePayload.PreviousCandlePayload.High Then
                            entryPrice = currentCandle.PreviousCandlePayload.High + buffer
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, 500, Trade.TypeOfStock.Cash)

                            Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice + targetPoint}
                            ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.SL, "Favourable High")
                        End If
                    Else
                        Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice + targetPoint}
                        ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.SL, "Highest Point")
                    End If
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                Dim lowestPoint As Decimal = GetLowestLow(currentCandle.PayloadDate)
                Dim entryPrice As Decimal = lowestPoint - buffer
                Dim targetPoint As Decimal = ConvertFloorCeling(GetAverageHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, 500, Trade.TypeOfStock.Cash)
                If currentCandle.Open < entryPrice Then
                    entryPrice = currentTick.Open
                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice - targetPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.Market, "Lowest Point Gap")
                Else
                    If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                        If currentCandle.PreviousCandlePayload.Low > currentCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                            entryPrice = currentCandle.PreviousCandlePayload.Low - buffer
                            quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice + targetPoint, 500, Trade.TypeOfStock.Cash)

                            Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice - targetPoint}
                            ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.SL, "Favourable Low")
                        End If
                    Else
                        Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .TargetPrice = entryPrice - targetPoint}
                        ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.SL, "Lowest Point")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _previousDayHighestATR <> Decimal.MinValue AndAlso _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)
            If todayHighestATR <> Decimal.MinValue Then
                ret = (_previousDayHighestATR + todayHighestATR) / 2
            Else
                ret = _previousDayHighestATR
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestHigh(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Max(Function(x)
                                     If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentTime Then
                                         Return x.Value.High
                                     Else
                                         Return Decimal.MinValue
                                     End If
                                 End Function)
        Return ret
    End Function

    Private Function GetLowestLow(ByVal currentTime As Date) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Min(Function(x)
                                     If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentTime Then
                                         Return x.Value.Low
                                     Else
                                         Return Decimal.MaxValue
                                     End If
                                 End Function)
        Return ret
    End Function
End Class
