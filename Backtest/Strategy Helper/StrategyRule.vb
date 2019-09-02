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

        Public EligibleToTakeTrade As Boolean = True
        Protected _signalPayload As Dictionary(Of Date, Payload)

        Protected ReadOnly _lotSize As Integer
        Protected ReadOnly _parentStrategy As Strategy
        Protected ReadOnly _tradingDate As Date
        Protected ReadOnly _tradingSymbol As String
        Protected ReadOnly _inputPayload As Dictionary(Of Date, Payload)
        Protected ReadOnly _cts As CancellationTokenSource
        Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                       ByVal lotSize As Integer,
                       ByVal parentStrategy As Strategy,
                       ByVal tradingDate As Date,
                       ByVal tradingSymbol As String,
                       ByVal canceller As CancellationTokenSource)
            _inputPayload = inputPayload
            _lotSize = lotSize
            _parentStrategy = parentStrategy
            _tradingDate = tradingDate
            _tradingSymbol = tradingSymbol
            _cts = canceller

            EligibleToTakeTrade = True
        End Sub

        Public Overridable Sub CompletePreProcessing()
            If _parentStrategy IsNot Nothing AndAlso _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
                Dim xMinutePayload As Dictionary(Of Date, Payload) = Nothing
                If _parentStrategy.SignalTimeFrame > 1 Then
                    xMinutePayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _parentStrategy.SignalTimeFrame)
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

        Public Overridable Async Function UpdateRequiredCollectionsAsync(ByVal currentTick As Payload) As Task
            Await Task.Delay(0).ConfigureAwait(False)
        End Function

        Public MustOverride Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Public MustOverride Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Public MustOverride Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Public MustOverride Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))

        Public Overridable Function IsSignalTriggered(ByVal entryPrice As Decimal, ByVal entryDirection As Trade.TradeExecutionDirection, ByVal startTime As Date, ByVal endTime As Date) As Boolean
            Dim ret As Boolean = False
            If _inputPayload IsNot Nothing AndAlso _inputPayload.Count > 0 Then
                For Each runningPayload In _inputPayload.Keys
                    If runningPayload >= startTime AndAlso runningPayload <= endTime Then
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