Imports Algo2TradeBLL
Namespace StrategyHelper
    Public Class PlaceOrderParameters
        Public EntryDirection As Trade.TradeExecutionDirection
        Public EntryPrice As Decimal
        Public Target As Decimal
        Public Stoploss As Decimal
        Public Quantity As Integer
        Public Buffer As Decimal
        Public SignalCandle As Payload
        Public Supporting1 As String
        Public Supporting2 As String
        Public Supporting3 As String
        Public Supporting4 As String
        Public Supporting5 As String
    End Class
End Namespace