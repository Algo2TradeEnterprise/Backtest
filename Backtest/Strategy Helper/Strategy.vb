Imports System.Threading
Imports Utilities.DAL
Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.IO
Imports NLog

Namespace StrategyHelper
    Public MustInherit Class Strategy

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

#Region "Constructor"
        Public Sub New(ByVal canceller As CancellationTokenSource,
                       ByVal exchangeStartTime As TimeSpan,
                       ByVal exchangeEndTime As TimeSpan,
                       ByVal tradeStartTime As TimeSpan,
                       ByVal lastTradeEntryTime As TimeSpan,
                       ByVal eodExitTime As TimeSpan,
                       ByVal tickSize As Decimal,
                       ByVal marginMultiplier As Decimal,
                       ByVal timeframe As Integer,
                       ByVal heikenAshiCandle As Boolean,
                       ByVal stockType As Trade.TypeOfStock,
                       ByVal tradeType As Trade.TypeOfTrade,
                       ByVal databaseTable As Common.DataBaseTable,
                       ByVal dataSource As SourceOfData,
                       ByVal initialCapital As Decimal,
                       ByVal usableCapital As Decimal,
                       ByVal minimumEarnedCapitalToWithdraw As Decimal,
                       ByVal amountToBeWithdrawn As Decimal)
            _canceller = canceller
            Me.ExchangeStartTime = exchangeStartTime
            Me.ExchangeEndTime = exchangeEndTime
            Me.TradeStartTime = tradeStartTime
            Me.LastTradeEntryTime = lastTradeEntryTime
            Me.EODExitTime = eodExitTime
            Me.TickSize = tickSize
            Me.MarginMultiplier = marginMultiplier
            Me.SignalTimeFrame = timeframe
            Me.UseHeikenAshi = heikenAshiCandle
            Me.StockType = stockType
            Me.TradeType = tradeType
            Me.DatabaseTable = databaseTable
            Me.DataSource = dataSource
            Me.InitialCapital = initialCapital
            Me.UsableCapital = usableCapital
            Me.MinimumEarnedCapitalToWithdraw = minimumEarnedCapitalToWithdraw
            Me.AmountToBeWithdrawn = amountToBeWithdrawn

            Cmn = New Common(_canceller)
            AddHandler Cmn.Heartbeat, AddressOf OnHeartbeat
            AddHandler Cmn.WaitingFor, AddressOf OnWaitingFor
            AddHandler Cmn.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            AddHandler Cmn.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
        End Sub
#End Region

#Region "Constants"
        Const MARGIN_MULTIPLIER As Integer = 30
        Const DEFAULT_TICK_SIZE As Decimal = 0.05
        Const EOD_EXIT_TIME As String = "15:00:00"
        Const LAST_TRADE_ENTRY_TIME As String = "15:00:00"
        Const TRADE_START_TIME As String = "09:15:00"
        Const EXCHANGE_START_TIME As String = "09:15:00"
        Const EXCHANGE_END_TIME As String = "15:30:00"
#End Region

#Region "Enum"
        Enum SourceOfData
            Live = 1
            Database
            None
        End Enum

        Enum MTMTrailingType
            RealtimeTrailing = 1
            FixedSlabTrailing
            LogSlabTrailing
            None
        End Enum
#End Region

#Region "Variables"
        Protected ReadOnly _canceller As CancellationTokenSource
        Protected TradesTaken As Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
        Protected CapitalMovement As Dictionary(Of Date, List(Of Capital))
        Protected AvailableCapital As Decimal = Decimal.MinValue

        Private StockNumberOfTradeBuffer As Dictionary(Of Date, Dictionary(Of String, Integer)) = Nothing

        Public ReadOnly Cmn As Common = Nothing
        Public ReadOnly ExchangeStartTime As TimeSpan = TimeSpan.Parse(EXCHANGE_START_TIME)
        Public ReadOnly ExchangeEndTime As TimeSpan = TimeSpan.Parse(EXCHANGE_END_TIME)
        Public ReadOnly TradeStartTime As TimeSpan = TimeSpan.Parse(TRADE_START_TIME)
        Public ReadOnly LastTradeEntryTime As TimeSpan = TimeSpan.Parse(LAST_TRADE_ENTRY_TIME)
        Public ReadOnly EODExitTime As TimeSpan = TimeSpan.Parse(EOD_EXIT_TIME)
        Public ReadOnly TickSize As Decimal = DEFAULT_TICK_SIZE
        Public ReadOnly MarginMultiplier As Decimal = MARGIN_MULTIPLIER
        Public ReadOnly SignalTimeFrame As Integer = 1
        Public ReadOnly UseHeikenAshi As Boolean = False
        Public ReadOnly StockType As Trade.TypeOfStock
        Public ReadOnly TradeType As Trade.TypeOfTrade
        Public ReadOnly DatabaseTable As Common.DataBaseTable
        Public ReadOnly DataSource As SourceOfData = SourceOfData.Database
        Public ReadOnly InitialCapital As Decimal = Decimal.MaxValue
        Public ReadOnly UsableCapital As Decimal = Decimal.MaxValue
        Public ReadOnly MinimumEarnedCapitalToWithdraw As Decimal = Decimal.MaxValue
        Public ReadOnly AmountToBeWithdrawn As Decimal = Decimal.MaxValue


        Public NumberOfTradeableStockPerDay As Integer = Integer.MaxValue
        Public ExitOnOverAllFixedTargetStoploss As Boolean = False
        Public OverAllProfitPerDay As Decimal = Decimal.MaxValue
        Public OverAllLossPerDay As Decimal = Decimal.MinValue
        Public StockMaxProfitPercentagePerDay As Double = Decimal.MaxValue
        Public StockMaxLossPercentagePerDay As Double = Decimal.MinValue
        Public ExitOnStockFixedTargetStoploss As Boolean = False
        Public StockMaxProfitPerDay As Double = Decimal.MaxValue
        Public StockMaxLossPerDay As Double = Decimal.MinValue
        Public NumberOfTradesPerStockPerDay As Integer = Integer.MaxValue
        Public NumberOfTradesPerDay As Integer = Integer.MaxValue
        Public AllowBothDirectionEntryAtSameTime As Boolean = False
        Public TickBasedStrategy As Boolean = False
        Public TrailingStoploss As Boolean = False
        Public TypeOfMTMTrailing As MTMTrailingType = MTMTrailingType.None
        Public MTMSlab As Decimal = Decimal.MinValue
        Public MovementSlab As Decimal = Decimal.MinValue
        Public RealtimeTrailingPercentage As Decimal = Decimal.MinValue

        Public MaxZSCore As Decimal = Decimal.MinValue
        Public NeglectedTradeCount As Integer = 0
#End Region

#Region "Public Calculated Property"
        Public ReadOnly Property TotalPLAfterBrokerage(ByVal currentDate As Date) As Decimal
            Get
                Dim ret As Decimal = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            ret += StockPLAfterBrokerage(currentDate, stock)
                        Next
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property TotalPLBeforeBrokerage(ByVal currentDate As Date) As Decimal
            Get
                Dim ret As Decimal = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            ret += StockPLBeforeBrokerage(currentDate, stock)
                        Next
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockPLPoint(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Decimal
            Get
                Dim ret As Decimal = Nothing
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    ret = TradesTaken(currentDate.Date)(stockTradingSymbol).Sum(Function(x)
                                                                                    If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                        Return x.PLPoint
                                                                                    Else
                                                                                        Return 0
                                                                                    End If
                                                                                End Function)
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockPLAfterBrokerage(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Decimal
            Get
                Dim ret As Decimal = Nothing
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    ret = TradesTaken(currentDate.Date)(stockTradingSymbol).Sum(Function(x)
                                                                                    If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                        Return x.PLAfterBrokerage
                                                                                    Else
                                                                                        Return 0
                                                                                    End If
                                                                                End Function)
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockPLBeforeBrokerage(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Decimal
            Get
                Dim ret As Decimal = Nothing
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    ret = TradesTaken(currentDate.Date)(stockTradingSymbol).Sum(Function(x)
                                                                                    If x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                                                                                        Return x.PLBeforeBrokerage
                                                                                    Else
                                                                                        Return 0
                                                                                    End If
                                                                                End Function)
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockNumberOfTrades(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    Dim tradeList As List(Of Trade) = TradesTaken(currentDate.Date)(stockTradingSymbol).FindAll(Function(x)
                                                                                                                    Return x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso
                                                                                                              x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open
                                                                                                                End Function)
                    If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                        Dim artnrGroups = From a In tradeList
                                          Group a By Key = a.Tag Into Group
                                          Select artnr = Key, numbersCount = Group.Count()

                        If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                            ret = artnrGroups.Count
                        End If
                    End If
                End If
                If StockNumberOfTradeBuffer IsNot Nothing AndAlso StockNumberOfTradeBuffer.ContainsKey(currentDate.Date) AndAlso
                    StockNumberOfTradeBuffer(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    ret += StockNumberOfTradeBuffer(currentDate.Date)(stockTradingSymbol)
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockNumberOfStoplossTrades(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    Dim tradeList As List(Of Trade) = TradesTaken(currentDate.Date)(stockTradingSymbol).FindAll(Function(x)
                                                                                                                    Return x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso
                                                                                                              x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open AndAlso
                                                                                                              x.ExitCondition = Trade.TradeExitCondition.StopLoss
                                                                                                                End Function)
                    If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                        Dim artnrGroups = From a In tradeList
                                          Group a By Key = a.Tag Into Group
                                          Select artnr = Key, numbersCount = Group.Count()

                        If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                            ret = artnrGroups.Count
                        End If
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockNumberOfTargetTrades(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    Dim tradeList As List(Of Trade) = TradesTaken(currentDate.Date)(stockTradingSymbol).FindAll(Function(x)
                                                                                                                    Return x.ExitCondition <> Trade.TradeExitCondition.Cancelled AndAlso
                                                                                                              x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open AndAlso
                                                                                                              x.ExitCondition = Trade.TradeExitCondition.Target
                                                                                                                End Function)
                    If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                        Dim artnrGroups = From a In tradeList
                                          Group a By Key = a.Tag Into Group
                                          Select artnr = Key, numbersCount = Group.Count()

                        If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                            ret = artnrGroups.Count
                        End If
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property TotalNumberOfTrades(ByVal currentDate As Date) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            ret += StockNumberOfTrades(currentDate, stock)
                        Next
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property StockNumberOfTradesWithoutBreakevenExit(ByVal currentDate As Date, ByVal stockTradingSymbol As String) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) AndAlso TradesTaken(currentDate.Date).ContainsKey(stockTradingSymbol) Then
                    ret = Me.StockNumberOfTrades(currentDate, stockTradingSymbol)
                    Dim beakevenTradeList As List(Of Trade) = TradesTaken(currentDate.Date)(stockTradingSymbol).FindAll(Function(x)
                                                                                                                            Return x.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso x.PLPoint > 0
                                                                                                                        End Function)
                    If beakevenTradeList IsNot Nothing AndAlso beakevenTradeList.Count > 0 Then
                        Dim artnrGroups = From a In beakevenTradeList
                                          Group a By Key = a.Tag Into Group
                                          Select artnr = Key, numbersCount = Group.Count()

                        If artnrGroups IsNot Nothing AndAlso artnrGroups.Count > 0 Then
                            ret = ret - artnrGroups.Count
                        End If
                    End If
                End If
                Return ret
            End Get
        End Property
        Public ReadOnly Property TotalNumberOfTradesWithoutBreakevenExit(ByVal currentDate As Date) As Integer
            Get
                Dim ret As Integer = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            ret += StockNumberOfTradesWithoutBreakevenExit(currentDate, stock)
                        Next
                    End If
                End If
                Return ret
            End Get
        End Property

        Private _TotalMaxDrawDownTime As Date = Date.MinValue
        Private _TotalMaxDrawDownPLAfterBrokerage As Decimal = Decimal.MaxValue
        Public ReadOnly Property TotalMaxDrawDownPLAfterBrokerage(ByVal currentDate As Date, ByVal currentTime As Date) As Decimal
            Get
                Dim pl As Decimal = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            pl += StockPLAfterBrokerage(currentDate, stock)
                        Next
                    End If
                    '_TotalMaxDrawDownPLAfterBrokerage = Math.Min(_TotalMaxDrawDownPLAfterBrokerage, pl)
                    If pl < _TotalMaxDrawDownPLAfterBrokerage Then
                        Me._TotalMaxDrawDownPLAfterBrokerage = pl
                        Me._TotalMaxDrawDownTime = currentTime
                    End If
                End If
                Return _TotalMaxDrawDownPLAfterBrokerage
            End Get
        End Property

        Private _TotalMaxDrawUpTime As Date = Date.MinValue
        Private _TotalMaxDrawUpPLAfterBrokerage As Decimal = Decimal.MinValue
        Public ReadOnly Property TotalMaxDrawUpPLAfterBrokerage(ByVal currentDate As Date, ByVal currentTime As Date) As Decimal
            Get
                Dim pl As Decimal = 0
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(currentDate.Date) Then
                    Dim stockTrades As Dictionary(Of String, List(Of Trade)) = TradesTaken(currentDate.Date)
                    If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                        For Each stock In stockTrades.Keys
                            pl += StockPLAfterBrokerage(currentDate, stock)
                        Next
                    End If
                    '_TotalMaxDrawUpPLAfterBrokerage = Math.Max(_TotalMaxDrawUpPLAfterBrokerage, pl)
                    If pl > _TotalMaxDrawUpPLAfterBrokerage Then
                        Me._TotalMaxDrawUpPLAfterBrokerage = pl
                        Me._TotalMaxDrawUpTime = currentTime
                    End If
                End If
                Return _TotalMaxDrawUpPLAfterBrokerage
            End Get
        End Property
#End Region

#Region "Public Functions"
        Public Function IsTradeActive(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TypeOfTrade, Optional ByVal tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None) As Boolean
            Dim ret As Boolean = False
            Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date
            If tradeType = Trade.TypeOfTrade.MIS Then
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                    Dim tradeList As List(Of Trade) = Nothing
                    If tradeDirection = Trade.TradeExecutionDirection.None Then
                        tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                           Return x.SquareOffType = tradeType AndAlso
                                                                                                   x.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress
                                                                                                       End Function)
                    Else
                        tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                           Return x.SquareOffType = tradeType AndAlso
                                                                                                   x.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                                                                                                   x.EntryDirection = tradeDirection
                                                                                                       End Function)
                    End If
                    ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
                End If
            ElseIf tradeType = Trade.TypeOfTrade.CNC Then
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    Dim tradeList As List(Of Trade) = Nothing
                    For Each runningDate In TradesTaken.Keys
                        If TradesTaken(runningDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                            For Each runningTrade In TradesTaken(runningDate)(currentMinutePayload.TradingSymbol)
                                If tradeDirection = Trade.TradeExecutionDirection.None Then
                                    If runningTrade.SquareOffType = tradeType AndAlso runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                        If tradeList Is Nothing Then tradeList = New List(Of Trade)
                                        tradeList.Add(runningTrade)
                                    End If
                                Else
                                    If runningTrade.SquareOffType = tradeType AndAlso
                                        runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso
                                        runningTrade.EntryDirection = tradeDirection Then
                                        If tradeList Is Nothing Then tradeList = New List(Of Trade)
                                        tradeList.Add(runningTrade)
                                    End If
                                End If
                            Next
                        End If
                    Next
                    ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
                End If
            End If
            Return ret
        End Function

        Public Function IsTradeOpen(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TypeOfTrade, Optional ByVal tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None) As Boolean
            Dim ret As Boolean = False
            Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date
            If tradeType = Trade.TypeOfTrade.MIS Then
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                    Dim tradeList As List(Of Trade) = Nothing
                    If tradeDirection = Trade.TradeExecutionDirection.None Then
                        tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                           Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = Trade.TradeExecutionStatus.Open
                                                                                                       End Function)
                    Else
                        tradeList = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                           Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = Trade.TradeExecutionStatus.Open AndAlso x.EntryDirection = tradeDirection
                                                                                                       End Function)
                    End If
                    ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
                End If
            ElseIf tradeType = Trade.TypeOfTrade.CNC Then
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    Dim tradeList As List(Of Trade) = Nothing
                    For Each runningDate In TradesTaken.Keys
                        If TradesTaken(runningDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                            For Each runningTrade In TradesTaken(runningDate)(currentMinutePayload.TradingSymbol)
                                If tradeDirection = Trade.TradeExecutionDirection.None Then
                                    If runningTrade.SquareOffType = tradeType AndAlso runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                                        If tradeList Is Nothing Then tradeList = New List(Of Trade)
                                        tradeList.Add(runningTrade)
                                    End If
                                Else
                                    If runningTrade.SquareOffType = tradeType AndAlso
                                        runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open AndAlso
                                        runningTrade.EntryDirection = tradeDirection Then
                                        If tradeList Is Nothing Then tradeList = New List(Of Trade)
                                        tradeList.Add(runningTrade)
                                    End If
                                End If
                            Next
                        End If
                    Next
                    ret = tradeList IsNot Nothing AndAlso tradeList.Count > 0
                End If
            End If

            Return ret
        End Function

        Public Function CalculateBuffer(ByVal price As Decimal, ByVal floorOrCeiling As RoundOfType) As Decimal
            Dim bufferPrice As Decimal = Nothing
            'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
            bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, TickSize, floorOrCeiling)
            Return bufferPrice
        End Function

        Public Function PlaceOrModifyOrder(ByVal currentTrade As Trade, ByVal modifyTrade As Trade) As Boolean
            Dim ret As Boolean = False
            Dim usableTrade As Trade = Nothing
            Dim capitalToBeAdded As Decimal = Nothing
            Dim capitalToBeReleased As Decimal = Nothing
            If modifyTrade Is Nothing Then  'Addition
                usableTrade = currentTrade
                capitalToBeAdded = currentTrade.CapitalRequiredWithMargin
                capitalToBeReleased = 0
            Else    'Modification
                If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then Throw New ApplicationException("Supplied trade is not open, cannot modify")
                usableTrade = modifyTrade
                capitalToBeAdded = modifyTrade.CapitalRequiredWithMargin
                capitalToBeReleased = currentTrade.CapitalRequiredWithMargin
            End If
            Dim tradeDate As Date = currentTrade.TradingDate.Date
            Dim tradingSymbol As String = currentTrade.TradingSymbol

            Dim lastTradingTime As Date = New Date(usableTrade.TradingDate.Year, usableTrade.TradingDate.Month, usableTrade.TradingDate.Day, LastTradeEntryTime.Hours, LastTradeEntryTime.Minutes, LastTradeEntryTime.Seconds)

            If usableTrade.EntryTime < lastTradingTime AndAlso Me.AvailableCapital >= capitalToBeAdded Then
                currentTrade.UpdateTrade(usableTrade)

                With usableTrade
                    currentTrade.UpdateTrade(EntryPrice:= .EntryPrice, PotentialTarget:= .PotentialTarget, PotentialStopLoss:= .PotentialStopLoss)
                    currentTrade.UpdateTrade(EntryTime:= .EntryTime, Quantity:= .Quantity, TradeCurrentStatus:=Trade.TradeExecutionStatus.Open)
                End With

                If modifyTrade Is Nothing Then
                    If currentTrade.EntryTime < lastTradingTime Then
                        If TradesTaken.ContainsKey(tradeDate) Then
                            If TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
                                TradesTaken(tradeDate)(tradingSymbol).Add(currentTrade)
                            Else
                                TradesTaken(tradeDate).Add(tradingSymbol, New List(Of Trade) From {currentTrade})
                            End If
                        Else
                            TradesTaken.Add(tradeDate, New Dictionary(Of String, List(Of Trade)) From {{tradingSymbol, New List(Of Trade) From {currentTrade}}})
                        End If
                    End If
                End If

                InsertCapitalRequired(currentTrade.EntryTime, capitalToBeAdded, capitalToBeReleased, "Placed Or Modify Order")
                ret = True
            Else
                If Me.AvailableCapital < capitalToBeAdded Then
                    Console.WriteLine(String.Format("Trade Neglected:{0},{1},{2}", tradeDate.ToShortDateString, tradingSymbol, usableTrade.EntryTime.ToString))
                End If
            End If
            Return ret
        End Function

        Public Function GetOpenActiveTrades(ByVal currentMinutePayload As Payload, ByVal tradeType As Trade.TypeOfTrade, ByVal direction As Trade.TradeExecutionDirection) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If currentMinutePayload IsNot Nothing Then
                Dim tradeDate As Date = currentMinutePayload.PayloadDate.Date
                If tradeType = Trade.TypeOfTrade.MIS Then
                    If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                        ret = TradesTaken(tradeDate)(currentMinutePayload.TradingSymbol).FindAll(Function(x)
                                                                                                     Return x.SquareOffType = tradeType AndAlso
                                                                                                     (x.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress OrElse
                                                                                                     x.TradeCurrentStatus = Trade.TradeExecutionStatus.Open) AndAlso
                                                                                                     x.EntryDirection = direction
                                                                                                 End Function)
                    End If
                ElseIf tradeType = Trade.TypeOfTrade.CNC Then
                    If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                        For Each runningDate In TradesTaken.Keys
                            If TradesTaken(runningDate).ContainsKey(currentMinutePayload.TradingSymbol) Then
                                For Each runningTrade In TradesTaken(runningDate)(currentMinutePayload.TradingSymbol)
                                    If runningTrade.SquareOffType = tradeType AndAlso runningTrade.EntryDirection = direction AndAlso
                                        (runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress OrElse
                                        runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open) Then
                                        If ret Is Nothing Then ret = New List(Of Trade)
                                        ret.Add(runningTrade)
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            Return ret
        End Function

        Public Function GetSpecificTrades(ByVal tradingSymbol As String, ByVal tradingDate As Date, ByVal tradeType As Trade.TypeOfTrade, ByVal tradeStatus As Trade.TradeExecutionStatus) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If tradingSymbol IsNot Nothing Then
                Dim tradeDate As Date = tradingDate.Date
                If tradeType = Trade.TypeOfTrade.MIS Then
                    If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 AndAlso TradesTaken.ContainsKey(tradeDate) AndAlso TradesTaken(tradeDate).ContainsKey(tradingSymbol) Then
                        ret = TradesTaken(tradeDate)(tradingSymbol).FindAll(Function(x)
                                                                                Return x.SquareOffType = tradeType AndAlso x.TradeCurrentStatus = tradeStatus
                                                                            End Function)
                    End If
                ElseIf tradeType = Trade.TypeOfTrade.CNC Then
                    If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                        For Each runningDate In TradesTaken.Keys
                            If TradesTaken(runningDate).ContainsKey(tradingSymbol) Then
                                For Each runningTrade In TradesTaken(runningDate)(tradingSymbol)
                                    If runningTrade.SquareOffType = tradeType AndAlso runningTrade.TradeCurrentStatus = tradeStatus Then
                                        If ret Is Nothing Then ret = New List(Of Trade)
                                        ret.Add(runningTrade)
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            Return ret
        End Function

        Public Function GetLastEntryTradeOfTheStock(ByVal tradingSymbol As String, ByVal tradingDate As Date, ByVal tradeType As Trade.TypeOfTrade) As Trade
            Dim ret As Trade = Nothing
            If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                Dim completeTrades As List(Of Trade) = GetSpecificTrades(tradingSymbol, tradingDate, tradeType, Trade.TradeExecutionStatus.Close)
                Dim inprogressTrades As List(Of Trade) = GetSpecificTrades(tradingSymbol, tradingDate, tradeType, Trade.TradeExecutionStatus.Inprogress)
                Dim allTrades As List(Of Trade) = New List(Of Trade)
                If completeTrades IsNot Nothing AndAlso completeTrades.Count > 0 Then allTrades.AddRange(completeTrades)
                If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then allTrades.AddRange(inprogressTrades)
                If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
                    ret = allTrades.OrderBy(Function(x)
                                                Return x.EntryTime
                                            End Function).LastOrDefault
                End If
            End If
            Return ret
        End Function

        Public Function GetAllTradesByTag(ByVal tag As String, ByVal tradingSymbol As String) As List(Of Trade)
            Dim ret As List(Of Trade) = Nothing
            If tradingSymbol IsNot Nothing Then
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    For Each runningDate In TradesTaken.Keys
                        If TradesTaken(runningDate).ContainsKey(tradingSymbol) Then
                            For Each runningTrade In TradesTaken(runningDate)(tradingSymbol)
                                If runningTrade.Tag = tag Then
                                    If ret Is Nothing Then ret = New List(Of Trade)
                                    ret.Add(runningTrade)
                                End If
                            Next
                        End If
                    Next
                End If
            End If
            Return ret
        End Function

        Public Sub ExitTradeByForce(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal exitRemark As String)
            If currentTrade Is Nothing Then Throw New ApplicationException("Supplied trade is nothing, cannot exit")

            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                CancelTrade(currentTrade, currentPayload, exitRemark)
            Else
                If currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Inprogress Then Throw New ApplicationException("Supplied trade is not active, cannot exit")
                currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                        ExitPrice:=currentPayload.Open,
                                        ExitCondition:=Trade.TradeExitCondition.ForceExit,
                                        ExitRemark:=exitRemark,
                                        TradeCurrentStatus:=Trade.TradeExecutionStatus.Close)
                InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin + currentTrade.PLAfterBrokerage, "Exit Trade By Force")
            End If
        End Sub

        Public Function EnterTradeIfPossible(ByVal tradingSymbol As String, ByVal tradingDate As Date, ByVal currentTrade As Trade, ByVal currentPayload As Payload) As Boolean
            Dim ret As Boolean = False
            Dim reverseSignalExit As Boolean = False
            If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then Throw New ApplicationException("Supplied trade is not open, cannot enter")

            'Dim previousRunningTrades As List(Of Trade) = GetSpecificTrades(tradingSymbol, tradingDate, currentTrade.SquareOffType, Trade.TradeExecutionStatus.Inprogress)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentPayload.High >= currentTrade.EntryPrice Then
                    'If Not AllowBothDirectionEntryAtSameTime AndAlso previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                    '    For Each previousRunningTrade In previousRunningTrades
                    '        If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    '            ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                    '            reverseSignalExit = True
                    '        End If
                    '    Next
                    'End If
                    Dim targetPoint As Decimal = currentTrade.PotentialTarget - currentTrade.EntryPrice
                    currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    currentTrade.UpdateTrade(PotentialTarget:=currentTrade.EntryPrice + targetPoint)
                    ret = True
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentPayload.Low <= currentTrade.EntryPrice Then
                    'If Not AllowBothDirectionEntryAtSameTime AndAlso previousRunningTrades IsNot Nothing AndAlso previousRunningTrades.Count > 0 Then
                    '    For Each previousRunningTrade In previousRunningTrades
                    '        If previousRunningTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    '            ExitTradeByForce(previousRunningTrade, currentPayload, "Opposite direction trade trigerred")
                    '            reverseSignalExit = True
                    '        End If
                    '    Next
                    'End If
                    Dim targetPoint As Decimal = currentTrade.EntryPrice - currentTrade.PotentialTarget
                    currentTrade.UpdateTrade(EntryPrice:=currentPayload.Open, EntryTime:=currentPayload.PayloadDate, TradeCurrentStatus:=Trade.TradeExecutionStatus.Inprogress)
                    currentTrade.UpdateTrade(PotentialTarget:=currentTrade.EntryPrice - targetPoint)
                    ret = True
                End If
            End If

            Return ret
        End Function

        Public Sub CancelTrade(ByVal currentTrade As Trade, ByVal currentPayload As Payload, ByVal exitRemark As String)
            If currentTrade Is Nothing OrElse currentTrade.TradeCurrentStatus <> Trade.TradeExecutionStatus.Open Then
                Throw New ApplicationException("Supplied trade is not open, cannot cancel")
            End If
            currentTrade.UpdateTrade(ExitTime:=currentPayload.PayloadDate,
                                 ExitCondition:=Trade.TradeExitCondition.Cancelled,
                                 ExitRemark:=exitRemark,
                                 TradeCurrentStatus:=Trade.TradeExecutionStatus.Cancel)

            InsertCapitalRequired(currentTrade.ExitTime, 0, currentTrade.CapitalRequiredWithMargin, "Cancel Trade")

            'TradesTaken(currentTrade.TradingDate.Date)(currentTrade.TradingSymbol).Remove(currentTrade)
        End Sub

        Public Function CalculatePL(ByVal stockName As String, ByVal buyPrice As Decimal, ByVal sellPrice As Decimal, ByVal quantity As Integer, ByVal lotSize As Integer, ByVal typeOfStock As Trade.TypeOfStock) As Decimal
            Dim potentialBrokerage As New Calculator.BrokerageAttributes
            Dim calculator As New Calculator.BrokerageCalculator(_canceller)

            Select Case typeOfStock
                Case Trade.TypeOfStock.Cash
                    If Me.TradeType = Trade.TypeOfTrade.MIS Then
                        calculator.Intraday_Equity(buyPrice, sellPrice, quantity, potentialBrokerage)
                    ElseIf Me.TradeType = Trade.TypeOfTrade.CNC Then
                        calculator.Delivery_Equity(buyPrice, sellPrice, quantity, potentialBrokerage)
                    End If
                Case Trade.TypeOfStock.Currency
                    calculator.Currency_Futures(buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
                Case Trade.TypeOfStock.Commodity
                    calculator.Commodity_MCX(stockName, buyPrice, sellPrice, quantity / lotSize, potentialBrokerage)
                Case Trade.TypeOfStock.Futures
                    calculator.FO_Futures(buyPrice, sellPrice, quantity, potentialBrokerage)
                Case Trade.TypeOfStock.Options
                    calculator.FO_Options(buyPrice, sellPrice, quantity, potentialBrokerage)
            End Select

            Return potentialBrokerage.NetProfitLoss
        End Function

        Public Function GetCurrentXMinuteCandleTime(ByVal lowerTFTime As Date) As Date
            Dim ret As Date = Nothing
            If Me.ExchangeStartTime.Minutes Mod Me.SignalTimeFrame = 0 Then
                ret = New Date(lowerTFTime.Year,
                                lowerTFTime.Month,
                                lowerTFTime.Day,
                                lowerTFTime.Hour,
                                Math.Floor(lowerTFTime.Minute / Me.SignalTimeFrame) * Me.SignalTimeFrame, 0)
            Else
                Dim exchangeTime As Date = New Date(lowerTFTime.Year, lowerTFTime.Month, lowerTFTime.Day, Me.ExchangeStartTime.Hours, Me.ExchangeStartTime.Minutes, 0)
                Dim currentTime As Date = New Date(lowerTFTime.Year, lowerTFTime.Month, lowerTFTime.Day, lowerTFTime.Hour, lowerTFTime.Minute, 0)
                Dim timeDifference As Double = currentTime.Subtract(exchangeTime).TotalMinutes
                Dim adjustedTimeDifference As Integer = Math.Floor(timeDifference / Me.SignalTimeFrame) * Me.SignalTimeFrame
                ret = exchangeTime.AddMinutes(adjustedTimeDifference)
                'Dim currentMinute As Date = exchangeTime.AddMinutes(adjustedTimeDifference)
                'ret = New Date(lowerTFTime.Year,
                '                lowerTFTime.Month,
                '                lowerTFTime.Day,
                '                currentMinute.Hour,
                '                currentMinute.Minute, 0)
            End If
            Return ret
        End Function

        Public Function GetMaxCapital(ByVal startTime As Date, ByVal endTime As Date) As Decimal
            Dim ret As Decimal = Decimal.MinValue
            If Me.CapitalMovement IsNot Nothing AndAlso CapitalMovement.Count > 0 Then
                For Each runningDate In CapitalMovement.Keys
                    If runningDate.Date >= startTime.Date AndAlso runningDate.Date <= endTime.Date Then
                        For Each runningCapital In CapitalMovement(runningDate)
                            If runningCapital.CapitalExhaustedDateTime >= startTime AndAlso
                                runningCapital.CapitalExhaustedDateTime <= endTime Then
                                ret = Math.Max(ret, runningCapital.RunningCapital)
                            End If
                        Next
                    End If
                Next
            End If
            Return ret
        End Function
#End Region

#Region "Public MustOverride Function"
        Public MustOverride Async Function TestStrategyAsync(startDate As Date, endDate As Date, ByVal filename As String) As Task
#End Region

#Region "Private Functions"
        Private Sub InsertCapitalRequired(ByVal currentDate As Date, ByVal capitalToBeInserted As Decimal, ByVal capitalToBeReleased As Decimal, ByVal remarks As String)
            Dim capitalRequired As Capital = New Capital
            Dim capitalToBeAdded As Decimal = 0

            If CapitalMovement IsNot Nothing AndAlso CapitalMovement.Count > 0 Then
                If Me.TradeType = Trade.TypeOfTrade.MIS Then
                    If CapitalMovement.ContainsKey(currentDate.Date) Then
                        Dim capitalList As List(Of Capital) = CapitalMovement(currentDate.Date)
                        capitalToBeAdded = capitalList.LastOrDefault.RunningCapital
                    End If
                ElseIf Me.TradeType = Trade.TypeOfTrade.CNC Then
                    Dim capitalList As List(Of Capital) = CapitalMovement.LastOrDefault.Value
                    capitalToBeAdded = capitalList.LastOrDefault.RunningCapital
                End If
            Else
                If CapitalMovement Is Nothing Then CapitalMovement = New Dictionary(Of Date, List(Of Capital))
            End If
            If Not CapitalMovement.ContainsKey(currentDate.Date) Then CapitalMovement.Add(currentDate.Date, New List(Of Capital))

            Me.AvailableCapital = Me.AvailableCapital - capitalToBeInserted + capitalToBeReleased
            With capitalRequired
                .TradingDate = currentDate.Date
                .CapitalExhaustedDateTime = currentDate
                .RunningCapital = capitalToBeInserted + capitalToBeAdded - capitalToBeReleased
                .CapitalReleased = capitalToBeReleased
                .AvailableCapital = Me.AvailableCapital
                .Remarks = remarks
            End With
            CapitalMovement(currentDate.Date).Add(capitalRequired)
        End Sub
#End Region

#Region "Deserialize"
        Public Function Deserialize(ByVal inputFilePath As String, ByVal retryCounter As Integer) As Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
            Dim ret As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
            If inputFilePath IsNot Nothing AndAlso File.Exists(inputFilePath) Then
                Using stream As New FileStream(inputFilePath, FileMode.Open)
                    Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                    Dim counter As Integer = 0
                    Dim totalSize As Long = 0
                    While stream.Position <> stream.Length
                        If totalSize <> 0 Then OnHeartbeat(String.Format("Deserializing Trades collection {0}/{1} Retry Counter:{2}", counter, totalSize, retryCounter))
                        Dim temp As KeyValuePair(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
                        Dim tempData As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = binaryFormatter.Deserialize(stream)
                        For Each runningDate In tempData.Keys
                            Dim stockData As Dictionary(Of String, List(Of Trade)) = Nothing
                            For Each stock In tempData(runningDate).Keys
                                Dim tradeList As List(Of Trade) = tempData(runningDate)(stock).FindAll(Function(x)
                                                                                                           Return x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                       End Function)
                                If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                                    For Each runningTrade In tradeList
                                        runningTrade.UpdateOriginatingStrategy(Me)
                                    Next
                                    If stockData Is Nothing Then stockData = New Dictionary(Of String, List(Of Trade))
                                    stockData.Add(stock, tradeList)
                                End If
                            Next
                            If stockData IsNot Nothing AndAlso stockData.Count > 0 Then
                                temp = New KeyValuePair(Of Date, Dictionary(Of String, List(Of Trade)))(runningDate, stockData)
                                If ret Is Nothing Then
                                    ret = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                                    totalSize = Math.Ceiling(stream.Length / stream.Position)
                                End If
                                ret.Add(temp.Key, temp.Value)
                            End If
                        Next
                        counter += 1
                    End While
                End Using
            End If
            Return ret
        End Function
#End Region

#Region "Print to excel direct"
        Public Overridable Sub PrintArrayToExcel(ByVal fileName As String, Optional ByVal tradesFilename As String = Nothing, Optional ByVal capitalFileName As String = Nothing)
            For retryCounter As Integer = 1 To 20 Step 1
                Try
                    Dim allTradesData As Dictionary(Of Date, Dictionary(Of String, List(Of Trade))) = Nothing
                    Dim allCapitalData As Dictionary(Of Date, List(Of Capital)) = Nothing
                    If tradesFilename IsNot Nothing AndAlso capitalFileName IsNot Nothing AndAlso File.Exists(tradesFilename) AndAlso File.Exists(capitalFileName) Then
                        OnHeartbeat("Deserializing Trades collections")
                        allTradesData = Deserialize(tradesFilename, retryCounter)
                        OnHeartbeat("Deserializing Capital collections")
                        allCapitalData = Utilities.Strings.DeserializeToCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName)
                    Else
                        allTradesData = TradesTaken
                        allCapitalData = CapitalMovement
                    End If
                    If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                        OnHeartbeat("Calculating supporting values")

                        Dim cts As New CancellationTokenSource

                        Dim totalTrades As Integer = allTradesData.Values.Sum(Function(x)
                                                                                  Return x.Values.Sum(Function(y)
                                                                                                          Return y.FindAll(Function(z)
                                                                                                                               Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                                           End Function).Count
                                                                                                      End Function)
                                                                              End Function)

                        Dim totalPositiveTrades As Integer = allTradesData.Values.Sum(Function(x)
                                                                                          Return x.Values.Sum(Function(y)
                                                                                                                  Return y.FindAll(Function(z)
                                                                                                                                       Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                            z.PLAfterBrokerage > 0
                                                                                                                                   End Function).Count
                                                                                                              End Function)
                                                                                      End Function)

                        Dim sumOfPositiveTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                          Return x.Values.Sum(Function(y)
                                                                                                                  Return y.Sum(Function(z)
                                                                                                                                   If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                       z.PLAfterBrokerage > 0 Then
                                                                                                                                       Return z.PLAfterBrokerage
                                                                                                                                   Else
                                                                                                                                       Return 0
                                                                                                                                   End If
                                                                                                                               End Function)
                                                                                                              End Function)
                                                                                      End Function)

                        Dim sumOfNegativeTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                          Return x.Values.Sum(Function(y)
                                                                                                                  Return y.Sum(Function(z)
                                                                                                                                   If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                       z.PLAfterBrokerage <= 0 Then
                                                                                                                                       Return z.PLAfterBrokerage
                                                                                                                                   Else
                                                                                                                                       Return 0
                                                                                                                                   End If
                                                                                                                               End Function)
                                                                                                              End Function)
                                                                                      End Function)

                        Dim totalDurationInTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                            Return x.Values.Sum(Function(y)
                                                                                                                    Return y.Sum(Function(z)
                                                                                                                                     If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                         Return z.DurationOfTrade.TotalMinutes
                                                                                                                                     Else
                                                                                                                                         Return 0
                                                                                                                                     End If
                                                                                                                                 End Function)
                                                                                                                End Function)
                                                                                        End Function)

                        Dim totalDurationInPositiveTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                                    Return x.Values.Sum(Function(y)
                                                                                                                            Return y.Sum(Function(z)
                                                                                                                                             If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                      z.PLAfterBrokerage > 0 Then
                                                                                                                                                 Return z.DurationOfTrade.TotalMinutes
                                                                                                                                             Else
                                                                                                                                                 Return 0
                                                                                                                                             End If
                                                                                                                                         End Function)
                                                                                                                        End Function)
                                                                                                End Function)

                        Dim totalDurationInNegativeTrades As Decimal = allTradesData.Values.Sum(Function(x)
                                                                                                    Return x.Values.Sum(Function(y)
                                                                                                                            Return y.Sum(Function(z)
                                                                                                                                             If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel AndAlso
                                                                                                                                      z.PLAfterBrokerage <= 0 Then
                                                                                                                                                 Return z.DurationOfTrade.TotalMinutes
                                                                                                                                             Else
                                                                                                                                                 Return 0
                                                                                                                                             End If
                                                                                                                                         End Function)
                                                                                                                        End Function)
                                                                                                End Function)

                        Dim largestWinningTrade As Decimal = allTradesData.Values.Max(Function(x)
                                                                                          Return x.Values.Max(Function(y)
                                                                                                                  Return y.Max(Function(z)
                                                                                                                                   If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                       Return z.PLAfterBrokerage
                                                                                                                                   Else
                                                                                                                                       Return Decimal.MinValue
                                                                                                                                   End If
                                                                                                                               End Function)
                                                                                                              End Function)
                                                                                      End Function)

                        Dim largestLosingTrade As Decimal = allTradesData.Values.Min(Function(x)
                                                                                         Return x.Values.Min(Function(y)
                                                                                                                 Return y.Min(Function(z)
                                                                                                                                  If z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel Then
                                                                                                                                      Return z.PLAfterBrokerage
                                                                                                                                  Else
                                                                                                                                      Return Decimal.MaxValue
                                                                                                                                  End If
                                                                                                                              End Function)
                                                                                                             End Function)
                                                                                     End Function)

                        Dim winRatio As Decimal = Math.Round((totalPositiveTrades / totalTrades) * 100, 2)
                        Dim riskReward As Decimal = 0
                        If totalPositiveTrades <> 0 AndAlso (totalTrades - totalPositiveTrades) <> 0 Then
                            riskReward = Math.Round(Math.Abs((sumOfPositiveTrades / totalPositiveTrades) / (sumOfNegativeTrades / (totalTrades - totalPositiveTrades))), 2)
                        End If

                        Dim strategyOutputData As StrategyOutput = New StrategyOutput
                        With strategyOutputData
                            .WinRatio = winRatio
                            .NetProfit = sumOfPositiveTrades + sumOfNegativeTrades
                            .GrossProfit = sumOfPositiveTrades
                            .GrossLoss = sumOfNegativeTrades
                            .TotalTrades = totalTrades
                            .TotalWinningTrades = totalPositiveTrades
                            .TotalLosingTrades = totalTrades - totalPositiveTrades
                            .AverageTrades = (sumOfPositiveTrades + sumOfNegativeTrades) / totalTrades
                            .AverageWinningTrades = If(totalPositiveTrades <> 0, sumOfPositiveTrades / totalPositiveTrades, 0)
                            .AverageLosingTrades = If((totalTrades - totalPositiveTrades) <> 0, sumOfNegativeTrades / (totalTrades - totalPositiveTrades), 0)
                            .RiskReward = riskReward
                            .LargestWinningTrade = largestWinningTrade
                            .LargestLosingTrade = largestLosingTrade
                            .AverageDurationInTrades = totalDurationInTrades / totalTrades
                            .AverageDurationInWinningTrades = If(totalPositiveTrades <> 0, totalDurationInPositiveTrades / totalPositiveTrades, 0)
                            .AverageDurationInLosingTrades = If((totalTrades - totalPositiveTrades) <> 0, totalDurationInNegativeTrades / (totalTrades - totalPositiveTrades), 0)
                        End With

                        Dim maxNumberOfDays As Decimal = Decimal.MinValue
                        For Each runningDate In allTradesData
                            For Each runningStock In runningDate.Value
                                If runningStock.Value IsNot Nothing AndAlso runningStock.Value.Count > 0 Then
                                    For Each runningTrade In runningStock.Value
                                        Try
                                            If runningTrade.Supporting8 IsNot Nothing AndAlso
                                            runningTrade.Supporting8.Trim <> "" AndAlso
                                            IsNumeric(runningTrade.Supporting8) Then
                                                maxNumberOfDays = Math.Max(Val(runningTrade.Supporting8), maxNumberOfDays)
                                            End If
                                        Catch ex As Exception
                                            Throw ex
                                        End Try
                                    Next
                                End If
                            Next
                        Next

                        Dim distinctTagList As List(Of String) = New List(Of String)
                        Dim runningTradeTag As String = Nothing
                        For Each runningDate In allTradesData
                            For Each runningStock In runningDate.Value
                                If runningStock.Value IsNot Nothing AndAlso runningStock.Value.Count > 0 Then
                                    For Each runningTrade In runningStock.Value
                                        If Not distinctTagList.Contains(runningTrade.Tag) Then
                                            distinctTagList.Add(runningTrade.Tag)
                                        End If
                                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                                            runningTradeTag = runningTrade.Tag
                                        End If
                                    Next
                                End If
                            Next
                        Next

                        Dim tradeCount As Integer = distinctTagList.Count
                        Dim pl As Decimal = strategyOutputData.NetProfit
                        Dim maxCapital As Decimal = allCapitalData.Values.Max(Function(x)
                                                                                  Return x.Max(Function(y)
                                                                                                   Return y.RunningCapital
                                                                                               End Function)
                                                                              End Function)

                        If runningTradeTag IsNot Nothing AndAlso runningTradeTag.Trim <> "" Then
                            Dim plToDeduct As Decimal = 0
                            Dim startTime As Date = Date.MinValue
                            For Each runningDate In allTradesData
                                For Each runningStock In runningDate.Value
                                    If runningStock.Value IsNot Nothing AndAlso runningStock.Value.Count > 0 Then
                                        For Each runningTrade In runningStock.Value
                                            If runningTrade.Tag = runningTradeTag Then
                                                plToDeduct += runningTrade.PLAfterBrokerage
                                                startTime = Date.ParseExact(runningTrade.Supporting6, "dd-MMM-yyyy HH:mm:ss", Nothing)
                                            End If
                                        Next
                                    End If
                                Next
                            Next

                            tradeCount = tradeCount - 1
                            pl = pl - plToDeduct
                            If startTime <> Date.MinValue Then
                                maxCapital = allCapitalData.Values.Max(Function(x)
                                                                           Return x.Max(Function(y)
                                                                                            If y.CapitalExhaustedDateTime < startTime Then
                                                                                                Return y.RunningCapital
                                                                                            Else
                                                                                                Return Decimal.MinValue
                                                                                            End If
                                                                                        End Function)
                                                                       End Function)
                            End If
                        End If

                        If tradeCount = 0 Then
                            fileName = String.Format("PL {0},Cap {1},ROI {2},TrdNmbr {3},MaxDays {4},{5}.xlsx",
                                             Math.Round(pl, 0),
                                             "∞",
                                             "∞",
                                             tradeCount,
                                             "∞",
                                             fileName)
                        Else
                            Dim roi As Decimal = (pl / maxCapital) * 100
                            fileName = String.Format("PL {0},Cap {1},ROI {2},LgclTrd {3},NglctTrd {4},MaxDays {5},{6}.xlsx",
                                             Math.Round(pl, 0),
                                             Math.Round(maxCapital, 0),
                                             Math.Round(roi, 0),
                                             tradeCount,
                                             Me.NeglectedTradeCount,
                                             Math.Round(maxNumberOfDays, 0),
                                             fileName)
                        End If

                        Dim filepath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output", fileName)
                        If File.Exists(filepath) Then File.Delete(filepath)

                        OnHeartbeat("Opening Excel.....")
                        Using excelWriter As New ExcelHelper(filepath, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, cts)
                            excelWriter.CreateNewSheet("Capital")
                            excelWriter.CreateNewSheet("Data")
                            excelWriter.SetActiveSheet("Data")

                            Dim rowCtr As Integer = 0
                            Dim colCtr As Integer = 0

                            Dim rowCount As Integer = 0
                            If allTradesData IsNot Nothing AndAlso allTradesData.Count > 0 Then
                                rowCount = allTradesData.Sum(Function(x)
                                                                 Dim stockTrades = x.Value
                                                                 Return stockTrades.Sum(Function(y)
                                                                                            Dim trades = y.Value.FindAll((Function(z)
                                                                                                                              Return z.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                                          End Function))
                                                                                            Return trades.Count
                                                                                        End Function)
                                                             End Function)
                            End If

                            Dim mainRawData(rowCount, 0) As Object

                            If rowCtr = 0 Then
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Trading Date"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Trading Symbol"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Entry Direction"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Buy Price"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Sell Price"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Quantity"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Entry Time"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Exit Time"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Exit Remark"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "PL Point"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "PL Before Brokerage"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "PL After Brokerage"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Signal Candle Time"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Month"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Tag"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Trade Number"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Entry Z-Score"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Mean"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "SD"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Correlation"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Start Date"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "End Date"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Duration"
                                colCtr += 1
                                If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                mainRawData(rowCtr, colCtr) = "Continues Z-Score"

                                rowCtr += 1
                            End If
                            Dim dateCtr As Integer = 0
                            For Each tempkeys In allTradesData.Keys
                                dateCtr += 1
                                OnHeartbeat(String.Format("Excel printing for Date: {0} [{1} of {2}]", tempkeys.Date.ToShortDateString, dateCtr, allTradesData.Count))
                                Dim stockTrades As Dictionary(Of String, List(Of Trade)) = allTradesData(tempkeys)
                                If stockTrades IsNot Nothing AndAlso stockTrades.Count > 0 Then
                                    Dim stockCtr As Integer = 0
                                    For Each stock In stockTrades.Keys
                                        stockCtr += 1
                                        If stockTrades.ContainsKey(stock) Then
                                            Dim tradeList As List(Of Trade) = stockTrades(stock).FindAll(Function(x)
                                                                                                             Return x.TradeCurrentStatus <> Trade.TradeExecutionStatus.Cancel
                                                                                                         End Function)
                                            Dim tradeCtr As Integer = 0
                                            If tradeList IsNot Nothing AndAlso tradeList.Count > 0 Then
                                                For Each tradeTaken In tradeList.OrderBy(Function(x)
                                                                                             Return x.EntryTime
                                                                                         End Function)
                                                    tradeCtr += 1
                                                    OnHeartbeat(String.Format("Excel printing: {0} of {1}", tradeCtr, tradeList.Count))
                                                    colCtr = 0

                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.TradingDate.ToString("dd-MMM-yyyy")
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.SupportingTradingSymbol
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.EntryDirection.ToString
                                                    colCtr += 1
                                                    If tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                        mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                        colCtr += 1
                                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                        mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                        colCtr += 1
                                                    ElseIf tradeTaken.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                        mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.ExitPrice, 4)
                                                        colCtr += 1
                                                        If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                        mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.EntryPrice, 4)
                                                        colCtr += 1
                                                    End If
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Quantity
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.EntryTime.ToString("yyyy-MM-dd HH:mm:ss")
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.ExitTime.ToString("yyyy-MM-dd HH:mm:ss")
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.ExitRemark
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLPoint, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = Math.Round(tradeTaken.PLBeforeBrokerage, 4)
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.PLAfterBrokerage
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.SignalCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss")
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = String.Format("{0}-{1}", tradeTaken.TradingDate.ToString("yyyy"), tradeTaken.TradingDate.ToString("MM"))
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Tag
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting1
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting2
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting3
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting4
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting5
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting6
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting7
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting8
                                                    colCtr += 1
                                                    If colCtr > UBound(mainRawData, 2) Then ReDim Preserve mainRawData(UBound(mainRawData, 1), 0 To UBound(mainRawData, 2) + 1)
                                                    mainRawData(rowCtr, colCtr) = tradeTaken.Supporting9

                                                    rowCtr += 1
                                                Next
                                            End If
                                        End If
                                    Next
                                End If
                            Next

                            Dim range As String = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                            RaiseEvent Heartbeat("Writing from memory to excel...")
                            excelWriter.WriteArrayToExcel(mainRawData, range)
                            Erase mainRawData
                            mainRawData = Nothing

                            excelWriter.SetActiveSheet("Capital")
                            rowCtr = 0
                            colCtr = 0

                            rowCount = 0
                            If allCapitalData IsNot Nothing AndAlso allCapitalData.Count > 0 Then
                                rowCount = allCapitalData.Values.Sum(Function(x)
                                                                         Return x.Count
                                                                     End Function)
                            End If

                            Dim capitalRawData(rowCount, 0) As Object

                            If rowCtr = 0 Then
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Trading Date"
                                colCtr += 1
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Capital Exhausted Date Time"
                                colCtr += 1
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Running Capital"
                                colCtr += 1
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Capital Released"
                                colCtr += 1
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Available Capital"
                                colCtr += 1
                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                capitalRawData(rowCtr, colCtr) = "Remarks"

                                rowCtr += 1
                            End If

                            If allCapitalData IsNot Nothing AndAlso allCapitalData.Count > 0 Then
                                Dim stockCtr As Integer = 0
                                For Each item In allCapitalData.Keys
                                    stockCtr += 1
                                    If allCapitalData.ContainsKey(item) Then
                                        Dim capitalList As List(Of Capital) = allCapitalData(item)
                                        Dim itemCtr As Integer = 0
                                        If capitalList IsNot Nothing AndAlso capitalList.Count > 0 Then
                                            For Each runningData In capitalList
                                                itemCtr += 1
                                                OnHeartbeat(String.Format("Excel printing: {0} of {1}", itemCtr, capitalList.Count))
                                                colCtr = 0

                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.TradingDate.ToString("dd-MMM-yyyy")
                                                colCtr += 1
                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.CapitalExhaustedDateTime.ToString("HH:mm:ss")
                                                colCtr += 1
                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(mainRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.RunningCapital
                                                colCtr += 1
                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.CapitalReleased
                                                colCtr += 1
                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.AvailableCapital
                                                colCtr += 1
                                                If colCtr > UBound(capitalRawData, 2) Then ReDim Preserve capitalRawData(UBound(capitalRawData, 1), 0 To UBound(capitalRawData, 2) + 1)
                                                capitalRawData(rowCtr, colCtr) = runningData.Remarks

                                                rowCtr += 1
                                            Next
                                        End If
                                    End If
                                Next
                            End If

                            range = excelWriter.GetNamedRange(1, rowCount, 1, colCtr)
                            OnHeartbeat("Writing from memory to excel...")
                            excelWriter.WriteArrayToExcel(capitalRawData, range)
                            Erase capitalRawData
                            capitalRawData = Nothing

                            excelWriter.SetActiveSheet("Data")
                            OnHeartbeat("Saving excel...")
                            excelWriter.SaveExcel()
                        End Using

                        If tradesFilename IsNot Nothing AndAlso File.Exists(tradesFilename) Then File.Delete(tradesFilename)
                        If capitalFileName IsNot Nothing AndAlso File.Exists(capitalFileName) Then File.Delete(capitalFileName)
                        Exit For
                    Else
                        Exit For
                    End If
                Catch ex As System.OutOfMemoryException
                    retryCounter += 1
                    GC.Collect()
                End Try
                Thread.Sleep(30000)
            Next
        End Sub
#End Region

    End Class
End Namespace