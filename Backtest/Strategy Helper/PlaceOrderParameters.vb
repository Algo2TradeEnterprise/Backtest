Imports Algo2TradeBLL
Namespace StrategyHelper
    Public Class PlaceOrderParameters
        Public EntryDirection As Trade.TradeExecutionDirection
        Public EntryPrice As Decimal
        Public Target As Decimal
        Public Stoploss As Decimal
        Public Quantity As Integer
        Public EntryBuffer As Decimal
        Public StoplossBuffer As Decimal
        Public SignalCandle As Payload
        Public OrderType As Trade.TypeOfOrder
        Public Supporting1 As String
        Public Supporting2 As String
        Public Supporting3 As String
        Public Supporting4 As String
        Public Supporting5 As String
        Public Used As Boolean
    End Class
End Namespace