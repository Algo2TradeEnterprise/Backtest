Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation
Imports System.IO

Public Class BollingerTouchFavourableBreakoutStrategyRule
    Inherits MathematicalStrategyRule

    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _niftyTrendPayload As Dictionary(Of Date, Decimal) = Nothing
    Private ReadOnly _minNiftyChangePer As Decimal = 0.1

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim trendFilename As String = Path.Combine(My.Application.Info.DirectoryPath, "NIFTY Trend.csv")
        If File.Exists(trendFilename) Then
            Dim dt As DataTable = Nothing
            Using csvHelper As New Utilities.DAL.CSVHelper(trendFilename, ",", _cts)
                dt = csvHelper.GetDataTableFromCSV(1)
            End Using
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim counter As Integer = 0
                For i = 0 To dt.Rows.Count - 1
                    Dim rowDateTime As Date = dt.Rows(i).Item("Date")
                    If rowDateTime.Date = _tradingDate.Date Then
                        'Dim trend As Decimal = dt.Rows(i).Item("Previous Day Close Trend %")
                        Dim trend As Decimal = dt.Rows(i).Item("Current Day Open Trend %")

                        If _niftyTrendPayload Is Nothing Then _niftyTrendPayload = New Dictionary(Of Date, Decimal)
                        _niftyTrendPayload.Add(rowDateTime, trend)
                    End If
                Next
            End If
        Else
            Throw New ApplicationException("Trend file not available")
        End If
        If _niftyTrendPayload Is Nothing OrElse _niftyTrendPayload.Count = 0 Then
            Throw New ApplicationException("NIFTY trend data not available")
        End If

        Indicator.BollingerBands.CalculateBollingerBands(50, Payload.PayloadFields.Close, 3, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Protected Overrides Function GetEntrySignal(currentCandle As Payload, currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)
        Dim ret As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentCandle, Trade.TypeOfTrade.MIS) Then
            If currentCandle.PreviousCandlePayload.Low < _bollingerLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                If _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                    _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) >= _minNiftyChangePer Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentCandle.PreviousCandlePayload.High + buffer
                    Dim slPoint As Decimal = entryPrice - (GetStoplossPrice(currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy) - buffer)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice - slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.SL, "")
                End If
            ElseIf currentCandle.PreviousCandlePayload.High > _bollingerHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                If _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                    _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) <= _minNiftyChangePer * -1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentCandle.PreviousCandlePayload.Low - buffer
                    Dim slPoint As Decimal = (GetStoplossPrice(currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell) + buffer) - entryPrice
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice + slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.SL, "")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetStoplossPrice(ByVal signalCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key <= signalCandle.PayloadDate Then
                If direction = Trade.TradeExecutionDirection.Buy Then
                    If ret = Decimal.MinValue Then ret = runningPayload.Value.Low
                    If runningPayload.Value.Low < _bollingerLowPayload(runningPayload.Key) Then
                        ret = Math.Min(ret, runningPayload.Value.Low)
                    Else
                        Exit For
                    End If
                ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                    If ret = Decimal.MinValue Then ret = runningPayload.Value.High
                    If runningPayload.Value.High > _bollingerHighPayload(runningPayload.Key) Then
                        ret = Math.Max(ret, runningPayload.Value.High)
                    Else
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function
End Class