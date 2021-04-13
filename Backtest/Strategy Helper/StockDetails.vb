Namespace StrategyHelper
    Public Class StockDetails
        Public StockName As String
        Public RawInstrumentName As String
        Public OptionStock As Boolean
        Public Controller As Boolean
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
        Public StockType As Trade.TypeOfStock = Trade.TypeOfStock.None
    End Class
End Namespace