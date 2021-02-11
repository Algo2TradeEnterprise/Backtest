Imports Algo2TradeBLL
Imports System.Threading
Imports NLog

Namespace StrategyHelper
    Public MustInherit Class StrategyRule

#Region "Logging and Status Progress"
        Public Shared logger As Logger = LogManager.GetCurrentClassLogger
#End Region

#Region "Events/Event handlers"
        Public Event DocumentDownloadComplete()
        Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        Public Event Heartbeat(ByVal msg As String)
        Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        'The below functions are needed to allow the derived classes to raise the above two events
        Protected Overridable Sub OnDocumentDownloadComplete()
            RaiseEvent DocumentDownloadComplete()
        End Sub
        Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
            RaiseEvent DocumentRetryStatus(currentTry, totalTries)
        End Sub
        Protected Overridable Sub OnHeartbeat(ByVal msg As String)
            RaiseEvent Heartbeat(msg)
        End Sub
        Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
            RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
        End Sub
#End Region

        Protected ReadOnly _TradingDate As Date
        Protected ReadOnly _NextTradingDay As Date
        Protected ReadOnly _TradingSymbol As String
        Protected ReadOnly _LotSize As Integer
        Protected ReadOnly _ParentStrategy As Strategy
        Protected ReadOnly _Cts As CancellationTokenSource
        Protected ReadOnly _InputMinPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _InputEODPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _TradeStartTime As Date

        Protected _SignalPayload As Dictionary(Of Date, Payload)

        Public Sub New(ByVal tradingDate As Date,
                       ByVal nextTradingDay As Date,
                       ByVal tradingSymbol As String,
                       ByVal lotSize As Integer,
                       ByVal parentStrategy As Strategy,
                       ByVal canceller As CancellationTokenSource,
                       ByVal inputMinPayload As Dictionary(Of Date, Payload),
                       ByVal inputEODPayload As Dictionary(Of Date, Payload))
            _TradingDate = tradingDate
            _NextTradingDay = nextTradingDay
            _TradingSymbol = tradingSymbol
            _LotSize = lotSize
            _ParentStrategy = parentStrategy
            _Cts = canceller
            _InputMinPayload = inputMinPayload
            _InputEODPayload = inputEODPayload
            _TradeStartTime = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day, _ParentStrategy.TradeStartTime.Hours, _ParentStrategy.TradeStartTime.Minutes, _ParentStrategy.TradeStartTime.Seconds)
        End Sub

        Public Overridable Sub CompletePreProcessing()
            If _ParentStrategy IsNot Nothing AndAlso _InputMinPayload IsNot Nothing AndAlso _InputMinPayload.Count > 0 Then
                If _ParentStrategy.SignalTimeFrame > 1 Then
                    Dim exchangeStartTime As Date = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day,
                                                             _ParentStrategy.ExchangeStartTime.Hours, _ParentStrategy.ExchangeStartTime.Minutes, _ParentStrategy.ExchangeStartTime.Seconds)
                    _SignalPayload = Common.ConvertPayloadsToXMinutes(_InputMinPayload, _ParentStrategy.SignalTimeFrame, exchangeStartTime)
                Else
                    _SignalPayload = _InputMinPayload
                End If
            End If
        End Sub

        Public MustOverride Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task
        Public MustOverride Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTickTime As Date, ByVal availableTrades As List(Of Trade)) As Task

    End Class
End Namespace