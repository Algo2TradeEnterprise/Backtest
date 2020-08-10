Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation
Imports System.IO

Public Class Stock_NiftyTrendVWAPStrategyRule
    Inherits MathematicalStrategyRule

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)
    End Sub

    Protected Overrides Function GetEntrySignal(currentCandle As Payload, currentTick As Payload) As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)
        Dim ret As Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentCandle, Trade.TypeOfTrade.MIS) Then
            If _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) > _minNiftyChangePer AndAlso
                currentCandle.PreviousCandlePayload.Close < _currentDayOpen Then
                If currentCandle.PreviousCandlePayload.Close > _vwapPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open

                    Dim slPoint As Decimal = ConvertFloorCeling(GetHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice - slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy, Trade.TypeOfOrder.Market, "")
                End If
            ElseIf _niftyTrendPayload.ContainsKey(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                _niftyTrendPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _minNiftyChangePer AndAlso
                currentCandle.PreviousCandlePayload.Close > _currentDayOpen Then
                If currentCandle.PreviousCandlePayload.Close < _vwapPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentTick.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = currentTick.Open

                    Dim slPoint As Decimal = ConvertFloorCeling(GetHighestATR(currentCandle.PreviousCandlePayload), _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, -500, Trade.TypeOfStock.Cash)

                    Dim entryData As EntryDetails = New EntryDetails With {.EntryPrice = entryPrice, .Quantity = quantity, .StoplossPrice = entryPrice + slPoint}
                    ret = New Tuple(Of Boolean, EntryDetails, Payload, Trade.TradeExecutionDirection, Trade.TypeOfOrder, String)(True, entryData, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell, Trade.TypeOfOrder.Market, "")
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            ret = _atrPayload.Max(Function(x)
                                      If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                          Return x.Value
                                      Else
                                          Return Decimal.MinValue
                                      End If
                                  End Function)
        End If
        Return ret
    End Function
End Class