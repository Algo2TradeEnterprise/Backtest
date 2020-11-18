Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class SwingCNCWithExitStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        Flat = 1
        AP
        GP
    End Enum
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public HigherTimeframe As Integer
        Public InitialCapital As Integer
        Public QuantityType As TypeOfQuantity
        Public MinimumExitPercentage As Decimal
    End Class
#End Region

    Private _htPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _maPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _tradeStartTime As Date
    Private ReadOnly _lastTradeEntryTime As Date
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
        _tradeStartTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        _lastTradeEntryTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _signalPayload, _maPayload)
        If _userInputs.HigherTimeframe < 375 Then
            _htPayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _userInputs.HigherTimeframe, New Date(Now.Year, Now.Month, Now.Day, 9, 15, 0))
        Else
            _htPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Cash, _tradingSymbol, _tradingDate.AddYears(-1), _tradingDate.AddDays(-1))
        End If
        Indicator.ATR.CalculateATR(14, _htPayload, _atrPayload, True)
        Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _htPayload, _bollingerHighPayload, _bollingerLowPayload, Nothing)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso currentMinuteCandle.PayloadDate >= _tradeStartTime Then
            Dim signal As Tuple(Of Boolean, Payload, Integer, Integer, String) = GetSignalForEntry(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim quantity As Integer = signal.Item3
                Dim iteration As Integer = signal.Item4
                Dim totalInvestedAmount As Decimal = currentTick.Open * quantity
                Dim totalQuantity As Long = quantity
                Dim potentialTarget As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, totalInvestedAmount / totalQuantity, totalQuantity, totalInvestedAmount * _userInputs.MinimumExitPercentage / 100, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)
                Dim startDate As String = _tradingDate.ToString("dd-MMM-yyyy")

                Dim activeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandle, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
                If activeTrades IsNot Nothing AndAlso activeTrades.Count > 0 Then
                    totalInvestedAmount += activeTrades.Sum(Function(x)
                                                                Return x.EntryPrice * x.Quantity
                                                            End Function)
                    totalQuantity += activeTrades.Sum(Function(x)
                                                          Return x.Quantity
                                                      End Function)
                    potentialTarget = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, totalInvestedAmount / totalQuantity, totalQuantity, totalInvestedAmount * _userInputs.MinimumExitPercentage / 100, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                    For Each runningTrade In activeTrades
                        runningTrade.UpdateTrade(Supporting2:=potentialTarget, Supporting3:=totalInvestedAmount, Supporting4:=totalQuantity)
                    Next

                    startDate = activeTrades.LastOrDefault.Supporting5
                End If


                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 1000000000,
                            .Target = .EntryPrice + 1000000000,
                            .Buffer = 0,
                            .SignalCandle = signal.Item2,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = iteration,
                            .Supporting2 = potentialTarget,
                            .Supporting3 = totalInvestedAmount,
                            .Supporting4 = totalQuantity,
                            .Supporting5 = startDate
                        }
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim totalInvest As Decimal = currentTrade.Supporting3
            Dim totalQty As Decimal = currentTrade.Supporting4
            Dim averagePrice As Decimal = totalInvest / totalQty

            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim htPayloadToCheck As Payload = Nothing
            If _userInputs.HigherTimeframe > 375 Then
                htPayloadToCheck = _htPayload.LastOrDefault.Value
            Else
                htPayloadToCheck = _htPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _htPayload, _userInputs.HigherTimeframe)).PreviousCandlePayload
            End If
            If currentMinuteCandle.PreviousCandlePayload.Close >= _bollingerHighPayload(htPayloadToCheck.PayloadDate) Then
                Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, averagePrice, currentMinuteCandle.PreviousCandlePayload.Close, totalQty, Me.LotSize, Trade.TypeOfStock.Cash)
                If pl >= totalInvest * _userInputs.MinimumExitPercentage / 100 Then
                    ret = New Tuple(Of Boolean, String)(True, "Bollinger Target")
                End If
            End If
            If ret Is Nothing Then
                Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, averagePrice, currentTick.Open, totalQty, Me.LotSize, Trade.TypeOfStock.Cash)
                Dim startDate As Date = Date.ParseExact(currentTrade.Supporting5, "dd-MMM-yyyy", Nothing)
                If pl >= totalInvest * _userInputs.MinimumExitPercentage * (currentTick.PayloadDate.Date.Subtract(startDate.Date).Days + 1) / 100 Then
                    ret = New Tuple(Of Boolean, String)(True, String.Format("{0}% Target",
                                                                            _userInputs.MinimumExitPercentage * (currentTick.PayloadDate.Date.Subtract(startDate.Date).Days + 1)))
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
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

    Private Function GetSignalForEntry(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Integer, Integer, String)
        Dim ret As Tuple(Of Boolean, Payload, Integer, Integer, String) = Nothing
        Dim lastTrade As Trade = GetLastOrder(currentCandle)
        'Dim atr As Decimal = _atrPayload(currentCandle.PreviousCandlePayload.PayloadDate)
        'Dim iteration As Integer = 1
        'Dim quantity As Integer = GetQuantity(iteration, currentTick.Open)
        If currentCandle.PreviousCandlePayload IsNot Nothing AndAlso currentCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle.PreviousCandlePayload
            Dim htPayloadToCheck As Payload = Nothing
            If _userInputs.HigherTimeframe > 375 Then
                htPayloadToCheck = _htPayload.LastOrDefault.Value
            Else
                htPayloadToCheck = _htPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _htPayload, _userInputs.HigherTimeframe)).PreviousCandlePayload
            End If

            Dim ma As Decimal = _maPayload(signalCandle.PayloadDate)
            Dim atr As Decimal = _atrPayload(htPayloadToCheck.PayloadDate)
            If signalCandle.Close > ma AndAlso signalCandle.CandleColor = Color.Green Then
                Dim crossedCandle As Payload = Nothing
                Dim chkCandle As Payload = signalCandle.PreviousCandlePayload
                While chkCandle IsNot Nothing
                    If chkCandle.Low <= _maPayload(chkCandle.PayloadDate) Then
                        crossedCandle = chkCandle
                        Exit While
                    Else
                        If chkCandle.Close > _maPayload(chkCandle.PayloadDate) AndAlso
                            chkCandle.CandleColor = Color.Green Then
                            Exit While
                        End If
                    End If
                    chkCandle = chkCandle.PreviousCandlePayload
                End While

                If crossedCandle IsNot Nothing Then
                    If lastTrade IsNot Nothing Then
                        If crossedCandle.PayloadDate <> lastTrade.SignalCandle.PayloadDate Then
                            If lastTrade.EntryPrice - signalCandle.Close >= atr Then
                                Dim iteration As Integer = Val(lastTrade.Supporting1) + 1
                                Dim quantity As Long = GetQuantity(iteration, currentTick.Open)
                                ret = New Tuple(Of Boolean, Payload, Integer, Integer, String)(True, currentCandle, quantity, iteration, "Below last entry")
                            End If
                        End If
                    Else
                        Dim iteration As Integer = 1
                        Dim quantity As Long = GetQuantity(iteration, currentTick.Open)
                        ret = New Tuple(Of Boolean, Payload, Integer, Integer, String)(True, crossedCandle, quantity, iteration, "First Trade")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetQuantity(ByVal iterationNumber As Integer, ByVal price As Decimal) As Long
        Dim capital As Decimal = _userInputs.InitialCapital
        Select Case _userInputs.QuantityType
            Case TypeOfQuantity.Flat
                capital = _userInputs.InitialCapital
            Case TypeOfQuantity.AP
                capital = _userInputs.InitialCapital * iterationNumber
            Case TypeOfQuantity.GP
                capital = _userInputs.InitialCapital * (Math.Pow(2, (iterationNumber - 1)))
            Case Else
                Throw New NotImplementedException
        End Select
        Return _parentStrategy.CalculateQuantityFromInvestment(1, capital, price, Trade.TypeOfStock.Cash, False)
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            Dim lastEntry As Trade = Me._parentStrategy.GetLastEntryTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
            If lastEntry IsNot Nothing AndAlso lastEntry.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                ret = lastEntry
            End If
        End If
        Return ret
    End Function
End Class