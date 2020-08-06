Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayHKTrendGapStrategyRule2
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
                Dim highestPoint As Decimal = GetHighestHigh(currentCandle, currentTick)
                Dim highestATR As Decimal = ConvertFloorCeling(GetAverageHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim entryPrice As Decimal = highestPoint - highestATR
                Dim slPoint As Decimal = highestATR * 2
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)
                If currentTick.Open <= entryPrice Then
                    entryPrice = currentTick.Open
                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice - slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.Market, "")
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                Dim lowestPoint As Decimal = GetLowestLow(currentCandle, currentTick)
                Dim highestATR As Decimal = ConvertFloorCeling(GetAverageHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                Dim entryPrice As Decimal = lowestPoint + highestATR
                Dim slPoint As Decimal = highestATR * 2
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)
                If currentTick.Open >= entryPrice Then
                    entryPrice = currentTick.Open
                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice + slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.Market, "")
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

    Private Function GetHighestHigh(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim candleHigh As Decimal = _signalPayload.Max(Function(x)
                                                           If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentCandle.PayloadDate Then
                                                               Return x.Value.High
                                                           Else
                                                               Return Decimal.MinValue
                                                           End If
                                                       End Function)
        Dim tickHigh As Decimal = currentCandle.Ticks.Max(Function(x)
                                                              If x.PayloadDate <= currentTick.PayloadDate Then
                                                                  Return x.High
                                                              Else
                                                                  Return Decimal.MinValue
                                                              End If
                                                          End Function)
        ret = Math.Max(candleHigh, tickHigh)
        Return ret
    End Function

    Private Function GetLowestLow(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim candleLow As Decimal = _signalPayload.Min(Function(x)
                                                          If x.Key.Date = _tradingDate.Date AndAlso x.Key < currentCandle.PayloadDate Then
                                                              Return x.Value.Low
                                                          Else
                                                              Return Decimal.MaxValue
                                                          End If
                                                      End Function)
        Dim tickLow As Decimal = currentCandle.Ticks.Min(Function(x)
                                                             If x.PayloadDate <= currentTick.PayloadDate Then
                                                                 Return x.Low
                                                             Else
                                                                 Return Decimal.MaxValue
                                                             End If
                                                         End Function)
        ret = Math.Min(candleLow, tickLow)
        Return ret
    End Function
End Class