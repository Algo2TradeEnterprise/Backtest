Namespace StrategyHelper
    Public Class StockDetails
        Public TradingSymbol As String
        Public StockName As String
        Public LotSize As Integer
        Public EligibleToTakeTrade As Boolean
        Public PlaceOrderTrigger As Tuple(Of Boolean, List(Of PlaceOrderParameters))
        'Public ExitOrderTrigger As Tuple(Of Boolean, String)
        'Public ModifyStoplossTrigger As Tuple(Of Boolean, Decimal, String)
        'Public ModifyTargetTrigger As Tuple(Of Boolean, Decimal, String)
        Public PlaceOrderDoneForTheMinute As Boolean
        Public ExitOrderDoneForTheMinute As Boolean
        Public CancelOrderDoneForTheMinute As Boolean
        Public ModifyStoplossOrderDoneForTheMinute As Boolean
        Public ModifyTargetOrderDoneForTheMinute As Boolean
        Public Supporting1 As Decimal
        Public Supporting2 As Decimal
        Public Supporting3 As Decimal
        Public Supporting4 As Decimal
        Public Supporting5 As Decimal
        Public Supporting6 As Decimal
        Public SupportingDate As Date
    End Class
End Namespace