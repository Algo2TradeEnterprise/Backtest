Namespace Indicator
    Public Module AnchoredVWAP
        Public Sub CalculateAnchoredVWAP(ByVal anchoredDateTime As Date, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Decimal))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim finalPriceToBeAdded As Decimal = 0
                Dim avgPrice As Decimal = Decimal.MinValue
                Dim avgPriceStarVolume As Decimal = Decimal.MinValue
                Dim cumAvgPriceStarVolume As Decimal = 0
                Dim cumVolume As Long = 0
                For Each runningInputPayload In inputPayload
                    If runningInputPayload.Key >= anchoredDateTime Then
                        avgPrice = (runningInputPayload.Value.High + runningInputPayload.Value.Low + runningInputPayload.Value.Close) / 3
                        avgPriceStarVolume = avgPrice * runningInputPayload.Value.Volume
                        cumAvgPriceStarVolume += avgPriceStarVolume
                        cumVolume += runningInputPayload.Value.Volume
                        If cumVolume <> 0 Then
                            finalPriceToBeAdded = cumAvgPriceStarVolume / cumVolume
                        End If
                        If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Decimal)
                        outputPayload.Add(runningInputPayload.Key, finalPriceToBeAdded)
                    End If
                Next
            End If
        End Sub
    End Module
End Namespace