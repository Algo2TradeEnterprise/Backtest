Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class InvestmentCNCStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum TypeOfQuantity
        NormalQuantity = 1
        MartingaleBasedQuantity
        TargetBasedQuantity
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
    Private ReadOnly _EODPayload As Dictionary(Of Date, Payload)
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockSMAPer As Decimal,
                   ByVal eodPayload As Dictionary(Of Date, Payload))
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _stockSMAPercentage = stockSMAPer
        _EODPayload = eodPayload
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_EODPayload, _fractalHighPayload, _fractalLowPayload)
        Indicator.SMA.CalculateSMA(200, Payload.PayloadFields.Close, _EODPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing

            Dim signalReceivedForEntry As Tuple(Of Boolean, Integer) = GetSignalForEntry(currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                Dim firstQuantity As Integer = 0
                Dim quantity As Integer = 0
                Select Case _userInputs.QuantityType
                    Case TypeOfQuantity.NormalQuantity, TypeOfQuantity.MartingaleBasedQuantity
                        quantity = 1
                    Case TypeOfQuantity.TargetBasedQuantity
                        quantity = signalReceivedForEntry.Item2
                End Select
                firstQuantity = quantity
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, _parentStrategy.TradeType)
                If lastExecutedTrade IsNot Nothing Then
                    If _userInputs.QuantityType = TypeOfQuantity.MartingaleBasedQuantity Then
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            quantity = lastExecutedTrade.Quantity * 2
                        End If
                    ElseIf _userInputs.QuantityType = TypeOfQuantity.TargetBasedQuantity Then
                        If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            quantity = lastExecutedTrade.Quantity + CInt(lastExecutedTrade.Supporting1)
                        End If
                    End If
                    If lastExecutedTrade.EntryTime.Date <> _tradingDate.Date Then
                        signalCandle = currentMinuteCandlePayload
                    End If
                Else
                    signalCandle = currentMinuteCandlePayload
                End If

                If signalCandle IsNot Nothing Then
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = currentTick.Open,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = quantity,
                        .Stoploss = .EntryPrice - 1000000,
                        .Target = .EntryPrice + 1000000,
                        .Buffer = 0,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.Market,
                        .Supporting1 = firstQuantity
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim signalReceivedForExit As Tuple(Of Boolean, String) = GetSignalForExit(currentTick)
            If signalReceivedForExit IsNot Nothing AndAlso signalReceivedForExit.Item1 Then
                ret = New Tuple(Of Boolean, String)(True, signalReceivedForExit.Item2)
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

    Private Function GetSignalForEntry(ByVal currentTick As Payload) As Tuple(Of Boolean, Integer)
        Dim ret As Tuple(Of Boolean, Integer) = Nothing
        If _EODPayload IsNot Nothing AndAlso _EODPayload.Count > 0 AndAlso _EODPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _EODPayload(currentTick.PayloadDate.Date)
            Dim fractalLow As Decimal = _fractalLowPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim fractalHigh As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim sma As Decimal = _smaPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            Dim entryPoint As Decimal = fractalLow

            If entryPoint > 100 AndAlso (fractalHigh >= sma OrElse fractalLow >= sma) Then
                If currentDayPayload.Open >= entryPoint Then
                    If currentTick.Open <= entryPoint Then
                        Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
                        ret = New Tuple(Of Boolean, Integer)(True, quantity)
                    End If
                Else
                    Dim tradeEntryTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 15, 25, 0)
                    If currentTick.PayloadDate >= tradeEntryTime Then
                        If currentDayPayload.CandleColor = Color.Red Then
                            Dim quantity As Integer = Math.Ceiling(100 / (_fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate) - entryPoint))
                            ret = New Tuple(Of Boolean, Integer)(True, quantity)
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSignalForExit(ByVal currentTick As Payload) As Tuple(Of Boolean, String)
        Dim ret As Tuple(Of Boolean, String) = Nothing
        If _EODPayload IsNot Nothing AndAlso _EODPayload.Count > 0 AndAlso _EODPayload.ContainsKey(currentTick.PayloadDate.Date) Then
            Dim currentDayPayload As Payload = _EODPayload(currentTick.PayloadDate.Date)
            Dim exitPoint As Decimal = _fractalHighPayload(currentDayPayload.PreviousCandlePayload.PayloadDate)
            If currentTick.Open >= exitPoint Then
                ret = New Tuple(Of Boolean, String)(True, "Target Hit")
            End If
        End If
        Return ret
    End Function
End Class
