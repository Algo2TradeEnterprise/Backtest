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
        Public Enum TradeExecutionDirection
            Buy = 1
            Sell
            None
        End Enum
        Public Enum TradeExecutionStatus
            Open = 1
            Inprogress
            Close
            Cancel
            None
        End Enum
        Public Enum TradeEntryCondition
            Original = 1
            Reversal
            Onward
            None
        End Enum
        Public Enum TradeExitCondition
            Target = 1
            StopLoss
            EndOfDay
            Cancelled
            ForceExit
            None
        End Enum
        Public Enum TypeOfTrade
            MIS = 1
            CNC
            None
        End Enum
        Enum TypeOfOrder
            Market
            SL
        End Enum
#End Region

#Region "Constructor"
        Public Sub New(ByVal originatingStrategy As Strategy,
                       ByVal tradingSymbol As String,
                       ByVal stockType As TypeOfStock,
                       ByVal orderType As TypeOfOrder,
                       ByVal tradingDate As Date,
                       ByVal entryDirection As TradeExecutionDirection,
                       ByVal entryPrice As Decimal,
                       ByVal entryBuffer As Decimal,
                       ByVal squareOffType As TypeOfTrade,
                       ByVal entryCondition As TradeEntryCondition,
                       ByVal entryRemark As String,
                       ByVal quantity As Integer,
                       ByVal lotSize As Integer,
                       ByVal potentialTarget As Decimal,
                       ByVal targetRemark As String,
                       ByVal potentialStopLoss As Decimal,
                       ByVal stoplossBuffer As Decimal,
                       ByVal slRemark As String,
                       ByVal signalCandle As Payload)
            Me._OriginatingStrategy = originatingStrategy
            Me._TradingSymbol = tradingSymbol
            Me._StockType = stockType
            Me._OrderType = orderType
            Me._EntryTime = tradingDate
            Me._TradingDate = tradingDate.Date
            Me._EntryDirection = entryDirection
            Me._EntryPrice = Math.Round(entryPrice, 4)
            Me._EntryBuffer = entryBuffer
            Me._SquareOffType = squareOffType
            Me._EntryCondition = entryCondition
            Me._EntryRemark = entryRemark
            Me._Quantity = quantity
            Me._LotSize = lotSize
            Me._PotentialTarget = Math.Round(potentialTarget, 4)
            Me._TargetRemark = targetRemark
            Me._PotentialStopLoss = Math.Round(potentialStopLoss, 4)
            Me._StoplossBuffer = stoplossBuffer
            Me._SLRemark = slRemark
            Me._StoplossSetTime = Me._EntryTime
            Me._SignalCandle = signalCandle
            Me._TradeUpdateTimeStamp = Me._EntryTime
            Me._TradeCurrentStatus = TradeExecutionStatus.None
            Me._ExitTime = Date.MinValue
            Me._ExitPrice = Decimal.MinValue
            Me._ExitCondition = TradeExitCondition.None
            Me._ExitRemark = Nothing
        End Sub
#End Region

#Region "Variables"
        <NonSerialized>
        Private _OriginatingStrategy As Strategy
        Public ReadOnly Property TradingSymbol As String
        Public ReadOnly Property CoreTradingSymbol As String
            Get
                If TradingSymbol.Contains("FUT") Then
                    Return TradingSymbol.Remove(TradingSymbol.Count - 8)
                Else
                    Return TradingSymbol
                End If
            End Get
        End Property
        Public ReadOnly Property StockType As TypeOfStock
        Public ReadOnly Property OrderType As TypeOfOrder
        Public ReadOnly Property EntryTime As DateTime
        Public ReadOnly Property TradingDate As Date '''''
        Public ReadOnly Property EntryDirection As TradeExecutionDirection
        Public ReadOnly Property FirstEntryDirection As TradeExecutionDirection
        Public ReadOnly Property AdditionalTrade As Boolean
        Public ReadOnly Property EntryPrice As Decimal
        Public ReadOnly Property EntryBuffer As Decimal
        Public ReadOnly Property SquareOffType As TypeOfTrade
        Public ReadOnly Property EntryCondition As TradeEntryCondition
        Public ReadOnly Property EntryRemark As String
        Public ReadOnly Property Quantity As Integer
        Public ReadOnly Property LotSize As Integer
        Public ReadOnly Property PotentialTarget As Decimal
        Public ReadOnly Property TargetRemark As String
        Public ReadOnly Property PotentialStopLoss As Decimal
        Public ReadOnly Property StoplossBuffer As Decimal
        Public ReadOnly Property SLRemark As String
        Public ReadOnly Property StoplossSetTime As Date
        Public ReadOnly Property SignalCandle As Payload
        Public ReadOnly Property TradeUpdateTimeStamp As Date
        Public ReadOnly Property TradeCurrentStatus As TradeExecutionStatus
        Public ReadOnly Property ExitTime As DateTime
        Public ReadOnly Property ExitPrice As Decimal
        Public ReadOnly Property ExitCondition As TradeExitCondition
        Public ReadOnly Property ExitRemark As String
        Public ReadOnly Property Tag As String
        Public ReadOnly Property SquareOffValue As Decimal
        Public ReadOnly Property Supporting1 As String
        Public ReadOnly Property Supporting2 As String
        Public ReadOnly Property Supporting3 As String
        Public ReadOnly Property Supporting4 As String
        Public ReadOnly Property Supporting5 As String
        Public ReadOnly Property Supporting6 As String
        Public ReadOnly Property Supporting7 As String
        Public ReadOnly Property Supporting8 As String
        Public ReadOnly Property Supporting9 As String
        Public ReadOnly Property SupportingTradingSymbol As String

        Public ReadOnly Property CapitalRequiredWithMargin As Decimal
            Get
                Return Me.EntryPrice * Me.Quantity / _OriginatingStrategy.MarginMultiplier
            End Get
        End Property

        Public ReadOnly Property PLPoint As Decimal
            Get
                If Me.TradeCurrentStatus = TradeExecutionStatus.Close Then
                    If Me.EntryDirection = TradeExecutionDirection.Buy Then
                        Return (Me.ExitPrice - Me.EntryPrice)
                    ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                        Return (Me.EntryPrice - Me.ExitPrice)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If Me.EntryDirection = TradeExecutionDirection.Buy Then
                        Return 0
                    ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property PLBeforeBrokerage As Decimal
            Get
                If Me.TradeCurrentStatus = TradeExecutionStatus.Close Then
                    If Me.EntryDirection = TradeExecutionDirection.Buy Then
                        Return ((Me.ExitPrice - Me.EntryPrice) * Me.Quantity)
                    ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                        Return ((Me.EntryPrice - Me.ExitPrice) * Me.Quantity)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If Me.EntryDirection = TradeExecutionDirection.Buy Then
                        Return 0
                    ElseIf Me.EntryDirection = TradeExecutionDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property PLAfterBrokerage As Decimal
            Get
                If Me.TradeCurrentStatus = TradeExecutionStatus.Close Then
                    If EntryDirection = TradeExecutionDirection.Buy Then
                        Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, EntryPrice, ExitPrice, Quantity, LotSize, StockType)
                    ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                        Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, ExitPrice, EntryPrice, Quantity, LotSize, StockType)
                    Else
                        Return Decimal.MinValue
                    End If
                Else
                    If EntryDirection = TradeExecutionDirection.Buy Then
                        Return 0
                    ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                        Return 0
                    Else
                        Return Decimal.MinValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property DurationOfTrade As TimeSpan
            Get
                Return Me.ExitTime - Me.EntryTime
            End Get
        End Property

        Public Property MaxDrawUp As Decimal
        Public Property MaxDrawDown As Decimal

        Public ReadOnly Property MaxDrawUpPL As Decimal
            Get
                If EntryDirection = TradeExecutionDirection.Buy Then
                    Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, EntryPrice, MaxDrawUp, Quantity, LotSize, StockType)
                ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                    Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, MaxDrawUp, EntryPrice, Quantity, LotSize, StockType)
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property

        Public ReadOnly Property MaxDrawDownPL As Decimal
            Get
                If EntryDirection = TradeExecutionDirection.Buy Then
                    Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, EntryPrice, MaxDrawDown, Quantity, LotSize, StockType)
                ElseIf EntryDirection = TradeExecutionDirection.Sell Then
                    Return _OriginatingStrategy.CalculatePL(CoreTradingSymbol, MaxDrawDown, EntryPrice, Quantity, LotSize, StockType)
                Else
                    Return Decimal.MinValue
                End If
            End Get
        End Property
#End Region

#Region "Public Fuction"
        Public Sub UpdateTrade(Optional ByVal TradingSymbol As String = Nothing,
                                Optional ByVal StockType As TypeOfStock = TypeOfStock.None,
                                Optional ByVal EntryTime As Date = Nothing,
                                Optional ByVal TradingDate As Date = Nothing,
                                Optional ByVal EntryDirection As TradeExecutionDirection = TradeExecutionDirection.None,
                                Optional ByVal EntryPrice As Decimal = Decimal.MinValue,
                                Optional ByVal EntryBuffer As Decimal = Decimal.MinValue,
                                Optional ByVal SquareOffType As TypeOfTrade = TypeOfTrade.None,
                                Optional ByVal EntryCondition As TradeEntryCondition = TradeEntryCondition.None,
                                Optional ByVal EntryRemark As String = Nothing,
                                Optional ByVal Quantity As Integer = Integer.MinValue,
                                Optional ByVal PotentialTarget As Decimal = Decimal.MinValue,
                                Optional ByVal TargetRemark As String = Nothing,
                                Optional ByVal PotentialStopLoss As Decimal = Decimal.MinValue,
                                Optional ByVal StoplossBuffer As Decimal = Decimal.MinValue,
                                Optional ByVal SLRemark As String = Nothing,
                                Optional ByVal StoplossSetTime As Date = Nothing,
                                Optional ByVal SignalCandle As Payload = Nothing,
                                Optional ByVal TradeCurrentStatus As TradeExecutionStatus = TradeExecutionStatus.None,
                                Optional ByVal ExitTime As Date = Nothing,
                                Optional ByVal ExitPrice As Decimal = Decimal.MinValue,
                                Optional ByVal ExitCondition As TradeExitCondition = TradeExitCondition.None,
                                Optional ByVal ExitRemark As String = Nothing,
                                Optional ByVal Tag As String = Nothing,
                                Optional ByVal SquareOffValue As Decimal = Decimal.MinValue,
                                Optional ByVal AdditionalTrade As Boolean = False,
                                Optional ByVal Supporting1 As String = Nothing,
                                Optional ByVal Supporting2 As String = Nothing,
                                Optional ByVal Supporting3 As String = Nothing,
                                Optional ByVal Supporting4 As String = Nothing,
                                Optional ByVal Supporting5 As String = Nothing,
                                Optional ByVal Supporting6 As String = Nothing,
                                Optional ByVal Supporting7 As String = Nothing,
                                Optional ByVal Supporting8 As String = Nothing,
                                Optional ByVal Supporting9 As String = Nothing,
                                Optional ByVal SupportingTradingSymbol As String = Nothing)


            If TradingSymbol IsNot Nothing Then _TradingSymbol = TradingSymbol
            If StockType <> TypeOfStock.None Then _StockType = StockType
            If EntryTime <> Nothing OrElse EntryTime <> Date.MinValue Then _EntryTime = EntryTime
            If TradingDate <> Nothing OrElse TradingDate <> Date.MinValue Then _TradingDate = TradingDate
            If EntryDirection <> TradeExecutionDirection.None Then _EntryDirection = EntryDirection
            If EntryPrice <> Decimal.MinValue Then _EntryPrice = Math.Round(EntryPrice, 4)
            If EntryBuffer <> Decimal.MinValue Then _EntryBuffer = EntryBuffer
            If SquareOffType <> TypeOfTrade.None Then _SquareOffType = SquareOffType
            If EntryCondition <> TradeEntryCondition.None Then _EntryCondition = EntryCondition
            If EntryRemark IsNot Nothing Then _EntryRemark = EntryRemark
            If Quantity <> Integer.MinValue Then _Quantity = Quantity
            If PotentialTarget <> Decimal.MinValue Then _PotentialTarget = Math.Round(PotentialTarget, 4)
            If TargetRemark IsNot Nothing Then _TargetRemark = TargetRemark
            If PotentialStopLoss <> Decimal.MinValue Then _PotentialStopLoss = Math.Round(PotentialStopLoss, 4)
            If StoplossBuffer <> Decimal.MinValue Then _StoplossBuffer = StoplossBuffer
            If SLRemark IsNot Nothing Then _SLRemark = SLRemark
            If StoplossSetTime <> Nothing OrElse StoplossSetTime <> Date.MinValue Then _StoplossSetTime = StoplossSetTime
            If SignalCandle IsNot Nothing Then _SignalCandle = SignalCandle
            If TradeCurrentStatus <> TradeExecutionStatus.None Then _TradeCurrentStatus = TradeCurrentStatus
            If ExitTime <> Nothing OrElse ExitTime <> Date.MinValue Then _ExitTime = ExitTime
            If ExitPrice <> Decimal.MinValue Then _ExitPrice = ExitPrice
            If ExitCondition <> TradeExitCondition.None Then _ExitCondition = ExitCondition
            If ExitRemark IsNot Nothing Then _ExitRemark = ExitRemark
            If Tag IsNot Nothing Then _Tag = Tag
            If SquareOffValue <> Decimal.MinValue Then _SquareOffValue = SquareOffValue
            If AdditionalTrade Then _AdditionalTrade = AdditionalTrade
            If Supporting1 IsNot Nothing Then _Supporting1 = Supporting1
            If Supporting2 IsNot Nothing Then _Supporting2 = Supporting2
            If Supporting3 IsNot Nothing Then _Supporting3 = Supporting3
            If Supporting4 IsNot Nothing Then _Supporting4 = Supporting4
            If Supporting5 IsNot Nothing Then _Supporting5 = Supporting5
            If Supporting6 IsNot Nothing Then _Supporting6 = Supporting6
            If Supporting7 IsNot Nothing Then _Supporting7 = Supporting7
            If Supporting8 IsNot Nothing Then _Supporting8 = Supporting8
            If Supporting9 IsNot Nothing Then _Supporting9 = Supporting9
            If SupportingTradingSymbol IsNot Nothing Then _SupportingTradingSymbol = SupportingTradingSymbol

            If Me._ExitTime <> Nothing OrElse Me._ExitTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._ExitTime
            ElseIf Me._StoplossSetTime <> Nothing OrElse Me._StoplossSetTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._StoplossSetTime
            ElseIf Me._EntryTime <> Nothing OrElse Me._EntryTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._EntryTime
            End If
        End Sub
        Public Sub UpdateTrade(ByVal tradeToBeUsed As Trade)
            If tradeToBeUsed Is Nothing Then Exit Sub
            With tradeToBeUsed
                If .TradingSymbol IsNot Nothing Then _TradingSymbol = .TradingSymbol
                If .StockType <> TypeOfStock.None Then _StockType = .StockType
                If .EntryTime <> Nothing OrElse .EntryTime <> Date.MinValue Then _EntryTime = .EntryTime
                If .TradingDate <> Nothing OrElse .TradingDate <> Date.MinValue Then _TradingDate = .TradingDate
                If .EntryDirection <> TradeExecutionDirection.None Then _EntryDirection = .EntryDirection
                If .EntryPrice <> Decimal.MinValue Then _EntryPrice = .EntryPrice
                If .EntryBuffer <> Decimal.MinValue Then _EntryBuffer = .EntryBuffer
                If .SquareOffType <> TypeOfTrade.None Then _SquareOffType = .SquareOffType
                If .EntryCondition <> TradeEntryCondition.None Then _EntryCondition = .EntryCondition
                If .EntryRemark IsNot Nothing Then _EntryRemark = .EntryRemark
                If .Quantity <> Integer.MinValue Then _Quantity = .Quantity
                If .PotentialTarget <> Decimal.MinValue Then _PotentialTarget = .PotentialTarget
                If .TargetRemark IsNot Nothing Then _TargetRemark = .TargetRemark
                If .PotentialStopLoss <> Decimal.MinValue Then _PotentialStopLoss = .PotentialStopLoss
                If .StoplossBuffer <> Decimal.MinValue Then _StoplossBuffer = .StoplossBuffer
                If .SLRemark IsNot Nothing Then _SLRemark = .SLRemark
                If .StoplossSetTime <> Nothing OrElse .StoplossSetTime <> Date.MinValue Then _StoplossSetTime = .StoplossSetTime
                If .SignalCandle IsNot Nothing Then _SignalCandle = .SignalCandle
                If .TradeCurrentStatus <> TradeExecutionStatus.None Then _TradeCurrentStatus = .TradeCurrentStatus
                If .ExitTime <> Nothing OrElse .ExitTime <> Date.MinValue Then _ExitTime = .ExitTime
                If .ExitPrice <> Decimal.MinValue Then _ExitPrice = .ExitPrice
                If .ExitCondition <> TradeExitCondition.None Then _ExitCondition = .ExitCondition
                If .ExitRemark IsNot Nothing Then _ExitRemark = .ExitRemark
                If .Tag IsNot Nothing Then _Tag = .Tag
                If .SquareOffValue <> Decimal.MinValue Then _SquareOffValue = .SquareOffValue
                If .Supporting1 IsNot Nothing Then _Supporting1 = .Supporting1
                If .Supporting2 IsNot Nothing Then _Supporting2 = .Supporting2
                If .Supporting3 IsNot Nothing Then _Supporting3 = .Supporting3
                If .Supporting4 IsNot Nothing Then _Supporting4 = .Supporting4
                If .Supporting5 IsNot Nothing Then _Supporting5 = .Supporting5
                If .Supporting6 IsNot Nothing Then _Supporting6 = .Supporting6
                If .Supporting7 IsNot Nothing Then _Supporting7 = .Supporting7
                If .Supporting8 IsNot Nothing Then _Supporting8 = .Supporting8
                If .Supporting9 IsNot Nothing Then _Supporting9 = .Supporting9
                If .SupportingTradingSymbol IsNot Nothing Then _SupportingTradingSymbol = .SupportingTradingSymbol
            End With

            If Me._ExitTime <> Nothing OrElse Me._ExitTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._ExitTime
            ElseIf Me._StoplossSetTime <> Nothing OrElse Me._StoplossSetTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._StoplossSetTime
            ElseIf Me._EntryTime <> Nothing OrElse Me._EntryTime <> Date.MinValue Then
                Me._TradeUpdateTimeStamp = Me._EntryTime
            End If
        End Sub

        Public Sub UpdateOriginatingStrategy(ByVal originatingStrategy As Strategy)
            Me._OriginatingStrategy = originatingStrategy
        End Sub
#End Region
    End Class
End Namespace