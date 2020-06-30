Namespace Indicator
    Public Module SwingHighLow
        Public Class Swing
            Public Property SwingHigh As Decimal
            Public Property SwingHighTime As Date
            Public Property SwingLow As Decimal
            Public Property SwingLowTime As Date
        End Class
        Public Sub CalculateSwingHighLow(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal strict As Boolean, ByRef outputPayload As Dictionary(Of Date, Swing))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                For Each runningPayload In inputPayload.Keys
                    Dim swingData As Swing = New Swing
                    If strict Then
                        If inputPayload(runningPayload).PreviousCandlePayload Is Nothing Then
                            swingData.SwingHigh = inputPayload(runningPayload).High
                            swingData.SwingHighTime = inputPayload(runningPayload).PayloadDate
                            swingData.SwingLow = inputPayload(runningPayload).Low
                            swingData.SwingLowTime = inputPayload(runningPayload).PayloadDate
                        ElseIf inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload Is Nothing Then
                            If inputPayload(runningPayload).PreviousCandlePayload.High > inputPayload(runningPayload).High Then
                                swingData.SwingHigh = inputPayload(runningPayload).PreviousCandlePayload.High
                                swingData.SwingHighTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHigh
                                swingData.SwingHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHighTime
                            End If
                            If inputPayload(runningPayload).PreviousCandlePayload.Low < inputPayload(runningPayload).Low Then
                                swingData.SwingLow = inputPayload(runningPayload).PreviousCandlePayload.Low
                                swingData.SwingLowTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLow
                                swingData.SwingLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLowTime
                            End If
                        Else
                            If inputPayload(runningPayload).PreviousCandlePayload.High > inputPayload(runningPayload).High AndAlso
                                inputPayload(runningPayload).PreviousCandlePayload.High > inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High Then
                                swingData.SwingHigh = inputPayload(runningPayload).PreviousCandlePayload.High
                                swingData.SwingHighTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHigh
                                swingData.SwingHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHighTime
                            End If
                            If inputPayload(runningPayload).PreviousCandlePayload.Low < inputPayload(runningPayload).Low AndAlso
                                inputPayload(runningPayload).PreviousCandlePayload.Low < inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low Then
                                swingData.SwingLow = inputPayload(runningPayload).PreviousCandlePayload.Low
                                swingData.SwingLowTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLow
                                swingData.SwingLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLowTime
                            End If
                        End If
                    Else
                        If inputPayload(runningPayload).PreviousCandlePayload Is Nothing Then
                            swingData.SwingHigh = inputPayload(runningPayload).High
                            swingData.SwingHighTime = inputPayload(runningPayload).PayloadDate
                            swingData.SwingLow = inputPayload(runningPayload).Low
                            swingData.SwingLowTime = inputPayload(runningPayload).PayloadDate
                        ElseIf inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload Is Nothing Then
                            If inputPayload(runningPayload).PreviousCandlePayload.High >= inputPayload(runningPayload).High Then
                                swingData.SwingHigh = inputPayload(runningPayload).PreviousCandlePayload.High
                                swingData.SwingHighTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHigh
                                swingData.SwingHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHighTime
                            End If
                            If inputPayload(runningPayload).PreviousCandlePayload.Low <= inputPayload(runningPayload).Low Then
                                swingData.SwingLow = inputPayload(runningPayload).PreviousCandlePayload.Low
                                swingData.SwingLowTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLow
                                swingData.SwingLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLowTime
                            End If
                        Else
                            If inputPayload(runningPayload).PreviousCandlePayload.High >= inputPayload(runningPayload).High AndAlso
                                inputPayload(runningPayload).PreviousCandlePayload.High >= inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High Then
                                swingData.SwingHigh = inputPayload(runningPayload).PreviousCandlePayload.High
                                swingData.SwingHighTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingHigh = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHigh
                                swingData.SwingHighTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingHighTime
                            End If
                            If inputPayload(runningPayload).PreviousCandlePayload.Low <= inputPayload(runningPayload).Low AndAlso
                                inputPayload(runningPayload).PreviousCandlePayload.Low <= inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low Then
                                swingData.SwingLow = inputPayload(runningPayload).PreviousCandlePayload.Low
                                swingData.SwingLowTime = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                            Else
                                swingData.SwingLow = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLow
                                swingData.SwingLowTime = outputPayload(inputPayload(runningPayload).PreviousCandlePayload.PayloadDate).SwingLowTime
                            End If
                        End If
                    End If

                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Swing)
                    outputPayload.Add(runningPayload, swingData)
                Next
            End If
        End Sub
    End Module
End Namespace
