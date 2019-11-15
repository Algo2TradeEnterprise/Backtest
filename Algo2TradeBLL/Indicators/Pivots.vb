Namespace Indicator
    Public Module Pivots
        Public Sub CalculatePivots(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, PivotPoints))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                If inputPayload.Count < 2 Then
                    Throw New ApplicationException("Can't Calculate Pivots")
                End If
                For Each runningInputPayload In inputPayload
                    Dim pivotPointsData As PivotPoints = New PivotPoints
                    If runningInputPayload.Value.PreviousCandlePayload IsNot Nothing Then
                        Dim prevHigh As Decimal = runningInputPayload.Value.High
                        Dim prevLow As Decimal = runningInputPayload.Value.Low
                        Dim prevClose As Decimal = runningInputPayload.Value.Close
                        pivotPointsData.Pivot = (prevHigh + prevLow + prevClose) / 3
                        pivotPointsData.Support1 = (2 * pivotPointsData.Pivot) - prevHigh
                        pivotPointsData.Resistance1 = (2 * pivotPointsData.Pivot) - prevLow
                        pivotPointsData.Support2 = pivotPointsData.Pivot - (prevHigh - prevLow)
                        pivotPointsData.Resistance2 = pivotPointsData.Pivot + (prevHigh - prevLow)
                        pivotPointsData.Support3 = pivotPointsData.Support2 - (prevHigh - prevLow)
                        pivotPointsData.Resistance3 = pivotPointsData.Resistance2 + (prevHigh - prevLow)
                    End If
                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, PivotPoints)
                    outputPayload.Add(runningInputPayload.Key, pivotPointsData)
                Next
            End If
        End Sub
    End Module
End Namespace
