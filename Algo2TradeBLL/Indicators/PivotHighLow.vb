Namespace Indicator
    Public Module PivotHighLow
        Public Class Pivot
            Public Property PivotHigh As Decimal
            Public Property PivotHighTime As Date
            Public Property PivotLow As Decimal
            Public Property PivotLowTime As Date
        End Class
        Public Sub CalculatePivotHighLow(ByVal period As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Pivot))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                If inputPayload.Count <= period * 2 + 1 Then
                    Throw New ApplicationException("Can not calculate pivot high low")
                End If
                For Each runningPayload In inputPayload.Keys
                    Dim pivotData As Pivot = Nothing

                    Dim previousNInputPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, runningPayload, period, True)
                    If previousNInputPayload IsNot Nothing AndAlso previousNInputPayload.Count = period Then
                        Dim highestHigh As Decimal = previousNInputPayload.Max(Function(x)
                                                                                   Return x.Value.High
                                                                               End Function)
                        Dim lowestLow As Decimal = previousNInputPayload.Min(Function(x)
                                                                                 Return x.Value.Low
                                                                             End Function)

                        Dim lastCandleTime As Date = previousNInputPayload.Min(Function(x)
                                                                                   Return x.Key
                                                                               End Function)

                        Dim pivotCandle As Payload = inputPayload(lastCandleTime).PreviousCandlePayload
                        If pivotCandle IsNot Nothing Then
                            Dim prePreviousNInputPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, pivotCandle.PayloadDate, period, False)
                            If prePreviousNInputPayload IsNot Nothing AndAlso prePreviousNInputPayload.Count = period Then
                                Dim preHighestHigh As Decimal = prePreviousNInputPayload.Max(Function(x)
                                                                                                 Return x.Value.High
                                                                                             End Function)
                                Dim preLowestLow As Decimal = prePreviousNInputPayload.Min(Function(x)
                                                                                               Return x.Value.Low
                                                                                           End Function)

                                If pivotCandle.High > highestHigh AndAlso pivotCandle.High > preHighestHigh Then
                                    If pivotData Is Nothing Then pivotData = New Pivot
                                    pivotData.PivotHigh = pivotCandle.High
                                    pivotData.PivotHighTime = pivotCandle.PayloadDate
                                Else
                                    If outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) IsNot Nothing Then
                                        If pivotData Is Nothing Then pivotData = New Pivot
                                        pivotData.PivotHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHigh
                                        pivotData.PivotHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHighTime
                                    End If
                                End If
                                If pivotCandle.Low < lowestLow AndAlso pivotCandle.Low < preLowestLow Then
                                    If pivotData Is Nothing Then pivotData = New Pivot
                                    pivotData.PivotLow = pivotCandle.Low
                                    pivotData.PivotLowTime = pivotCandle.PayloadDate
                                Else
                                    If outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) IsNot Nothing Then
                                        If pivotData Is Nothing Then pivotData = New Pivot
                                        pivotData.PivotLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLow
                                        pivotData.PivotLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLowTime
                                    End If
                                End If
                            Else
                                If outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) IsNot Nothing Then
                                    pivotData = New Pivot With {
                                        .PivotHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHigh,
                                        .PivotHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHighTime,
                                        .PivotLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLow,
                                        .PivotLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLowTime
                                    }
                                End If
                            End If
                        Else
                            If outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate) IsNot Nothing Then
                                pivotData = New Pivot With {
                                    .PivotHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHigh,
                                    .PivotHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotHighTime,
                                    .PivotLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLow,
                                    .PivotLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).PivotLowTime
                                }
                            End If
                        End If
                    End If

                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Pivot)
                    outputPayload.Add(runningPayload, pivotData)
                Next
            End If
        End Sub
    End Module
End Namespace