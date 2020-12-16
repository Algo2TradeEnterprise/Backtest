Namespace Indicator
    Public Module HMA
        Public Sub CalculateHMA(ByVal hmaPeriod As Integer, ByVal hmaField As Payload.PayloadFields, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Decimal))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim nWMAPayload As Dictionary(Of Date, Decimal) = Nothing
                WMA.CalculateWMA(hmaPeriod, hmaField, inputPayload, nWMAPayload)
                Dim n2WMAPayload As Dictionary(Of Date, Decimal) = Nothing
                WMA.CalculateWMA(hmaPeriod / 2, hmaField, inputPayload, n2WMAPayload)

                Dim midInputPayload As Dictionary(Of Date, Decimal) = Nothing
                For Each runningPayload In inputPayload.Keys
                    Dim inputVal As Decimal = 2 * n2WMAPayload(runningPayload) - nWMAPayload(runningPayload)
                    If midInputPayload Is Nothing Then midInputPayload = New Dictionary(Of Date, Decimal)
                    midInputPayload.Add(runningPayload, inputVal)
                Next

                Dim midInputConPayload As Dictionary(Of Date, Payload) = Common.ConvertDecimalToPayload(Payload.PayloadFields.Additional_Field, midInputPayload)

                WMA.CalculateWMA(CInt(Math.Sqrt(hmaPeriod)), Payload.PayloadFields.Additional_Field, midInputConPayload, outputPayload)
            End If
        End Sub
    End Module
End Namespace