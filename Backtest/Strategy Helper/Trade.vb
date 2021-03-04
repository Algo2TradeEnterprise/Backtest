Imports Algo2TradeBLL

Namespace StrategyHelper
    <Serializable>
    Public Class Trade

#Region "Enum"
        Public Enum TypeOfStock
            Cash = 1
            Currency
            Commodity
            Futures
            Options
            None
        End Enum
        Public Enum TradeDirection
            Buy = 1
            Sell
            None
        End Enum
        Public Enum TradeStatus
            Open = 1
            Inprogress
            Complete
            Cancel
            None
        End Enum
        Public Enum TypeOfExit
            Target = 1
            Stoploss
            Reversal
            ContractRollover
            None
        End Enum
        Enum TypeOfEntry
            Fresh = 1
            Reversal
            Stoploss
            Rollover
            None
        End Enum
#End Region

#Region "Constructor"
        Public Sub New(ByVal originatingStrategy As Strategy,
                       ByVal tradingSymbol As String,
                       ByVal spotTradingSymbol As String,
                       ByVal stockType As TypeOfStock,
                       ByVal tradingDate As Date,
                       ByVal signalDirection As TradeDirection,
                       ByVal entryDirection As TradeDirection,
                       ByVal entryType As TypeOfEntry,
                       ByVal entryPrice As Decimal,
                       ByVal quantity As Long,
                       ByVal lotSize As Integer,
                       ByVal entrySignalCandle As Payload,
                       ByVal childReference As String,
                       ByVal parentReference As String,
                       ByVal iterationNumber As Integer,
                       ByVal spotPrice As Decimal,
                       ByVal requiredCapital As Decimal,
                       ByVal previousLoss As Decimal,
                       ByVal potentialTarget As Decimal)
            _OriginatingStrategy = originatingStrategy
            _TradingSymbol = tradingSymbol
            _SpotTradingSymbol = spotTradingSymbol
            _StockType = stockType
            _TradingDate = tradingDate
            _SignalDirection = signalDirection
            _EntryDirection = entryDirection
            _EntryType = entryType
            _EntryPrice = entryPrice
            _Quantity = quantity
            _LotSize = lotSize
            _EntrySignalCandle = entrySignalCandle
            _ChildReference = childReference
            _ParentReference = parentReference
            _IterationNumber = iterationNumber
            _SpotPrice = spotPrice
            _RequiredCapital = requiredCapital
            _PreviousLoss = previousLoss
            _PotentialTarget = potentialTarget
        End Sub
#End Region

#Region "Variables"
        <NonSerialized>
        Private _OriginatingStrategy As Strategy

        Public ReadOnly Property TradingSymbol As String
        Public ReadOnly Property SpotTradingSymbol As String
        Public ReadOnly Property StockType As TypeOfStock
        Public ReadOnly Property TradingDate As Date
        Public ReadOnly Property SignalDirection As TradeDirection
        Public ReadOnly Property EntryDirection As TradeDirection
        Public ReadOnly Property EntryType As TypeOfEntry
        Public ReadOnly Property EntryTime As Date
        Public ReadOnly Property EntryPrice As Decimal
        Public ReadOnly Property Quantity As Long
        Public ReadOnly Property LotSize As Integer
        Public ReadOnly Property ExitTime As Date
        Public ReadOnly Property ExitPrice As Decimal
        Public ReadOnly Property ExitType As TypeOfExit
        Public ReadOnly Property EntrySignalCandle As Payload
        Public ReadOnly Property ExitSignalCandle As Payload
        Public ReadOnly Property TradeCurrentStatus As TradeStatus
        Public ReadOnly Property ChildReference As String
        Public ReadOnly Property ParentReference As String
        Public ReadOnly Property IterationNumber As Integer
        Public ReadOnly Property SpotPrice As Decimal
        Public ReadOnly Property RequiredCapital As Decimal
        Public ReadOnly Property PreviousLoss As Decimal
        Public ReadOnly Property PotentialTarget As Decimal
        Public ReadOnly Property ContractRolloverEntry As Boolean

        Public ReadOnly Property CapitalRequiredWithMargin As Decimal
            Get
                Return Me.EntryPrice * Me.Quantity / _OriginatingStrategy.MarginMultiplier
            End Get
        End Property

        Public ReadOnly Property PLPoint As Decimal
            Get
                If Me.TradeCurrentStatus = TradeStatus.Complete Then
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return (Me.ExitPrice - Me.EntryPrice)
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return (Me.EntryPrice - Me.ExitPrice)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return 0
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property PLBeforeBrokerage As Decimal
            Get
                If Me.TradeCurrentStatus = TradeStatus.Complete Then
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return ((Me.ExitPrice - Me.EntryPrice) * Me.Quantity)
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return ((Me.EntryPrice - Me.ExitPrice) * Me.Quantity)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return 0
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property PLAfterBrokerage As Decimal
            Get
                If Me.TradeCurrentStatus = TradeStatus.Complete Then
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, EntryPrice, ExitPrice, Quantity, LotSize, StockType)
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, ExitPrice, EntryPrice, Quantity, LotSize, StockType)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If Me.EntryDirection = TradeDirection.Buy Then
                        Return 0
                    ElseIf Me.EntryDirection = TradeDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property MaxDrawUp As Decimal
        Public ReadOnly Property MaxDrawDown As Decimal

        Public ReadOnly Property MaxDrawUpPL As Decimal
            Get
                If EntryDirection = TradeDirection.Buy Then
                    Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, EntryPrice, MaxDrawUp, Quantity, LotSize, StockType)
                ElseIf EntryDirection = TradeDirection.Sell Then
                    Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, MaxDrawUp, EntryPrice, Quantity, LotSize, StockType)
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property MaxDrawDownPL As Decimal
            Get
                If EntryDirection = TradeDirection.Buy Then
                    Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, EntryPrice, MaxDrawDown, Quantity, LotSize, StockType)
                ElseIf EntryDirection = TradeDirection.Sell Then
                    Return _OriginatingStrategy.CalculatePLAfterBrokerage(SpotTradingSymbol, MaxDrawDown, EntryPrice, Quantity, LotSize, StockType)
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property
#End Region

#Region "Public Fuction"
        Public Sub UpdateTrade(Optional ByVal tradingSymbol As String = Nothing,
                               Optional ByVal spotTradingSymbol As String = Nothing,
                               Optional ByVal stockType As TypeOfStock = TypeOfStock.None,
                               Optional ByVal tradingDate As Date = Nothing,
                               Optional ByVal signalDirection As TradeDirection = TradeDirection.None,
                               Optional ByVal entryDirection As TradeDirection = TradeDirection.None,
                               Optional ByVal entryType As TypeOfEntry = TypeOfEntry.None,
                               Optional ByVal entryTime As Date = Nothing,
                               Optional ByVal entryPrice As Decimal = Decimal.MinValue,
                               Optional ByVal quantity As Integer = Integer.MinValue,
                               Optional ByVal lotSize As Integer = Integer.MinValue,
                               Optional ByVal exitTime As Date = Nothing,
                               Optional ByVal exitPrice As Decimal = Decimal.MinValue,
                               Optional ByVal exitType As TypeOfExit = TypeOfExit.None,
                               Optional ByVal entrySignalCandle As Payload = Nothing,
                               Optional ByVal exitSignalCandle As Payload = Nothing,
                               Optional ByVal tradeCurrentStatus As TradeStatus = TradeStatus.None,
                               Optional ByVal childReference As String = Nothing,
                               Optional ByVal parentReference As String = Nothing,
                               Optional ByVal iterationNumber As Integer = Integer.MinValue,
                               Optional ByVal spotPrice As Decimal = Decimal.MinValue,
                               Optional ByVal requiredCapital As Decimal = Decimal.MinValue,
                               Optional ByVal previousLoss As Decimal = Decimal.MinValue,
                               Optional ByVal potentialTarget As Decimal = Decimal.MinValue,
                               Optional ByVal maxDrawUp As Decimal = Decimal.MinValue,
                               Optional ByVal maxDrawDown As Decimal = Decimal.MinValue,
                               Optional ByVal contractRolloverEntry As Boolean = False)
            If tradingSymbol IsNot Nothing Then _TradingSymbol = tradingSymbol
            If spotTradingSymbol IsNot Nothing Then _SpotTradingSymbol = spotTradingSymbol
            If stockType <> TypeOfStock.None Then _StockType = stockType
            If tradingDate <> Nothing AndAlso tradingDate <> Date.MinValue Then _TradingDate = tradingDate
            If signalDirection <> TradeDirection.None Then _SignalDirection = signalDirection
            If entryDirection <> TradeDirection.None Then _EntryDirection = entryDirection
            If entryType <> TypeOfEntry.None Then _EntryType = entryType
            If entryTime <> Nothing AndAlso entryTime <> Date.MinValue Then _EntryTime = entryTime
            If entryPrice <> Decimal.MinValue Then _EntryPrice = entryPrice
            If quantity <> Integer.MinValue Then _Quantity = quantity
            If lotSize <> Integer.MinValue Then _LotSize = lotSize
            If exitTime <> Nothing AndAlso exitTime <> Date.MinValue Then _ExitTime = exitTime
            If exitPrice <> Decimal.MinValue Then _ExitPrice = exitPrice
            If exitType <> TypeOfExit.None Then _ExitType = exitType
            If entrySignalCandle IsNot Nothing Then _EntrySignalCandle = entrySignalCandle
            If exitSignalCandle IsNot Nothing Then _ExitSignalCandle = exitSignalCandle
            If tradeCurrentStatus <> TradeStatus.None Then _TradeCurrentStatus = tradeCurrentStatus
            If childReference IsNot Nothing Then _ChildReference = childReference
            If parentReference IsNot Nothing Then _ParentReference = parentReference
            If iterationNumber <> Integer.MinValue Then _IterationNumber = iterationNumber
            If spotPrice <> Decimal.MinValue Then _SpotPrice = spotPrice
            If requiredCapital <> Decimal.MinValue Then _RequiredCapital = requiredCapital
            If previousLoss <> Decimal.MinValue Then _PreviousLoss = previousLoss
            If potentialTarget <> Decimal.MinValue Then _PotentialTarget = potentialTarget
            If maxDrawUp <> Decimal.MinValue Then _MaxDrawUp = maxDrawUp
            If maxDrawDown <> Decimal.MinValue Then _MaxDrawDown = maxDrawDown
            If contractRolloverEntry Then _ContractRolloverEntry = contractRolloverEntry
        End Sub

        Public Sub UpdateOriginatingStrategy(ByVal originatingStrategy As Strategy)
            Me._OriginatingStrategy = originatingStrategy
        End Sub
#End Region
    End Class
End Namespace