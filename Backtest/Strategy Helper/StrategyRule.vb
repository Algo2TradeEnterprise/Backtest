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

        Public MaxProfitOfThisStock As Decimal = Decimal.MaxValue
        Public MaxLossOfThisStock As Decimal = Decimal.MinValue
        Public LotSize As Integer

        Protected ContractRollover As Boolean = False
        Protected BlankDayExit As Boolean = False
        Protected ForceCancellationDone As Boolean = False
        Protected ContractRolloverForceEntry As Boolean = False

        Protected ForceTakeTrade As Boolean = False
        Protected ForceCancelTrade As Boolean = False

        Public EligibleToTakeTrade As Boolean = True
        Public ReadOnly Property TradingSymbol As String
        Public ReadOnly Property InstrumentName As String
            Get
                If Me.TradingSymbol.Contains("FUT") Then
                    Return Me.TradingSymbol.Remove(Me.TradingSymbol.Count - 8)
                Else
                    Return Me.TradingSymbol
                End If
            End Get
        End Property

        Protected _signalPayload As Dictionary(Of Date, Payload)

        Protected ReadOnly _parentStrategy As Strategy
        Protected ReadOnly _tradingDate As Date
        Protected _inputPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _cts As CancellationTokenSource
        Protected ReadOnly _entities As RuleEntities
        Protected ReadOnly _nextTradingDay As Date

        'Public AnotherPairInstrument As StrategyRule

        Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                       ByVal lotSize As Integer,
                       ByVal parentStrategy As Strategy,
                       ByVal tradingDate As Date,
                       ByVal tradingSymbol As String,
                       ByVal nextTradingDay As Date,
                       ByVal entities As RuleEntities,
                       ByVal canceller As CancellationTokenSource)
            _inputPayload = inputPayload
            Me.LotSize = lotSize
            _parentStrategy = parentStrategy
            _tradingDate = tradingDate
            Me.TradingSymbol = tradingSymbol
            _cts = canceller
            _entities = entities
            _nextTradingDay = nextTradingDay

            EligibleToTakeTrade = True
        End Sub

        Public Overridable Sub CompletePreProcessing()
            If _parentStrategy IsNot Nothing AndAlso _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
                Dim xMinutePayload As Dictionary(Of Date, Payload) = Nothing
                If _parentStrategy.SignalTimeFrame > 1 Then
                    Dim exchangeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day,
                                                             _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)
                    xMinutePayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _parentStrategy.SignalTimeFrame, exchangeStartTime)
                Else
                    xMinutePayload = _inputPayload
                End If
                If _parentStrategy.UseHeikenAshi Then
                    Indicator.HeikenAshi.ConvertToHeikenAshi(xMinutePayload, _signalPayload)
                Else
                    _signalPayload = xMinutePayload
                End If
            End If
        End Sub

        Public Overridable Sub CompletePairProcessing()

        End Sub

        Public Overridable Async Function UpdateRequiredCollectionsAsync(ByVal currentTick As Payload) As Task
            Await Task.Delay(0).ConfigureAwait(False)
        End Function

        Public Sub UpdateInputPayloadAndTradingSymbol(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tradingSymbol As String)
            _inputPayload = inputPayload
            _TradingSymbol = tradingSymbol
        End Sub


        Public MustOverride Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Public MustOverride Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Public MustOverride Async Function IsTriggerReceivedForExitCNCEODOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Public MustOverride Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Public MustOverride Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))

        Public Overridable Function IsSignalTriggered(ByVal entryPrice As Decimal, ByVal entryDirection As Trade.TradeExecutionDirection, ByVal startTime As Date, ByVal endTime As Date) As Boolean
            Dim ret As Boolean = False
            startTime = _parentStrategy.GetCurrentXMinuteCandleTime(startTime, _inputPayload)
            endTime = _parentStrategy.GetCurrentXMinuteCandleTime(endTime, _inputPayload)
            If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
                For Each runningPayload In _inputPayload.Keys
                    If runningPayload >= startTime AndAlso runningPayload < endTime Then
                        If entryDirection = Trade.TradeExecutionDirection.Buy Then
                            If _inputPayload(runningPayload).High >= entryPrice Then
                                ret = True
                                Exit For
                            End If
                        ElseIf entryDirection = Trade.TradeExecutionDirection.Sell Then
                            If _inputPayload(runningPayload).Low <= entryPrice Then
                                ret = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
            Return ret
        End Function
    End Class
End Namespace