Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation
Imports System.IO

Public Class Stock_NiftyTrendHKBreakoutStrategyRule
    Inherits MathematicalStrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _currentDayOpen As Decimal = Decimal.MinValue
    Private _niftyTrendPayload As Dictionary(Of Date, Decimal) = Nothing
    Private ReadOnly _minNiftyChangePer As Decimal = 0

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

        _currentDayOpen = _signalPayload.Where(Function(x)
                                                   Return x.Key.Date = _tradingDate.Date
                                               End Function).OrderBy(Function(y)
                                                                         Return y.Key
                                                                     End Function).FirstOrDefault.Value.Open

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
    End Sub

    Protected Overrides Function GetEntrySignal(currentCandle As Payload, currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)
        Dim ret As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentCandle, Trade.TypeOfTrade.MIS) Then
            Dim hkCandle As Payload = _hkPayload(currentCandle.PreviousCandlePayload.PayloadDate)
            If _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) > _minNiftyChangePer AndAlso
                currentCandle.PreviousCandlePayload.Close > _currentDayOpen Then
                If Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Buy)
                    Dim stoploss As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Sell)

                    Dim slPoint As Decimal = entryPrice - stoploss
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice - slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, hkCandle, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.SL, "")
                End If
            ElseIf _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _minNiftyChangePer AndAlso
                currentCandle.PreviousCandlePayload.Close < _currentDayOpen Then
                If Math.Round(hkCandle.Open, 2) = Math.Round(hkCandle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Sell)
                    Dim stoploss As Decimal = GetEntryPrice(hkCandle, Trade.TradeExecutionDirection.Buy)

                    Dim slPoint As Decimal = stoploss - entryPrice
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice - slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, hkCandle, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.SL, "")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetEntryPrice(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If candle IsNot Nothing AndAlso direction <> Trade.TradeExecutionDirection.None Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                ret = ConvertFloorCeling(candle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If ret = Math.Round(candle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                    ret = ret + buffer
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                ret = ConvertFloorCeling(candle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If ret = Math.Round(candle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                    ret = ret - buffer
                End If
            End If
        End If
        Return ret
    End Function
End Class