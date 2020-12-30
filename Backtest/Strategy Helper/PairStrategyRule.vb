Imports Algo2TradeBLL
Imports System.Threading
Imports NLog

Namespace StrategyHelper
    Public MustInherit Class PairStrategyRule

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

        Protected ReadOnly _PairName As String
        Protected ReadOnly _TradingDate As Date
        Protected ReadOnly _NextTradingDay As Date
        Protected ReadOnly _TradingSymbol1 As String
        Protected ReadOnly _LotSize1 As Integer
        Protected ReadOnly _TradingSymbol2 As String
        Protected ReadOnly _LotSize2 As Integer
        Protected ReadOnly _Entities As RuleEntities
        Protected ReadOnly _ParentStrategy As Strategy
        Protected ReadOnly _Cts As CancellationTokenSource
        Protected ReadOnly _InputPayload1 As Dictionary(Of Date, Payload)
        Protected ReadOnly _InputPayload2 As Dictionary(Of Date, Payload)

        Protected _SignalPayload1 As Dictionary(Of Date, Payload)
        Protected _SignalPayload2 As Dictionary(Of Date, Payload)

        Public Sub New(ByVal pairName As String,
                       ByVal tradingDate As Date,
                       ByVal nextTradingDay As Date,
                       ByVal tradingSymbol1 As String,
                       ByVal lotSize1 As Integer,
                       ByVal tradingSymbol2 As String,
                       ByVal lotSize2 As Integer,
                       ByVal entities As RuleEntities,
                       ByVal parentStrategy As Strategy,
                       ByVal canceller As CancellationTokenSource,
                       ByVal inputPayload1 As Dictionary(Of Date, Payload),
                       ByVal inputPayload2 As Dictionary(Of Date, Payload))
            _PairName = pairName
            _TradingDate = tradingDate
            _NextTradingDay = nextTradingDay
            _TradingSymbol1 = tradingSymbol1
            _LotSize1 = lotSize1
            _TradingSymbol2 = tradingSymbol2
            _LotSize2 = lotSize2
            _Entities = entities
            _ParentStrategy = parentStrategy
            _Cts = canceller
            _InputPayload1 = inputPayload1
            _InputPayload2 = inputPayload2
        End Sub

        Public Overridable Sub CompletePreProcessing()
            If _ParentStrategy IsNot Nothing AndAlso
                _InputPayload1 IsNot Nothing AndAlso _InputPayload1.Count > 0 AndAlso
                _InputPayload2 IsNot Nothing AndAlso _InputPayload2.Count > 0 Then
                If _ParentStrategy.SignalTimeFrame > 1 Then
                    Dim exchangeStartTime As Date = New Date(_TradingDate.Year, _TradingDate.Month, _TradingDate.Day,
                                                             _ParentStrategy.ExchangeStartTime.Hours, _ParentStrategy.ExchangeStartTime.Minutes, _ParentStrategy.ExchangeStartTime.Seconds)
                    _SignalPayload1 = Common.ConvertPayloadsToXMinutes(_InputPayload1, _ParentStrategy.SignalTimeFrame, exchangeStartTime)
                    _SignalPayload2 = Common.ConvertPayloadsToXMinutes(_InputPayload2, _ParentStrategy.SignalTimeFrame, exchangeStartTime)
                Else
                    _SignalPayload1 = _InputPayload1
                    _SignalPayload2 = _InputPayload2
                End If

            End If
        End Sub

        Public MustOverride Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTickTime As Date) As Task
        Public MustOverride Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTickTime As Date, ByVal availableTrades As List(Of Trade)) As Task

    End Class
End Namespace