Namespace Indicator
    Public Module AROON
        Public Sub CalculateAROON(ByVal period As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputHighPayload As Dictionary(Of Date, Decimal), ByRef outputLowPayload As Dictionary(Of Date, Decimal))
            'Using WILDER Formula
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                If inputPayload.Count < period + 1 Then
                    Throw New ApplicationException("Can't Calculate AROON")
                End If
                Dim counter As Integer = 0
                For Each runningPayload In inputPayload
                    counter += 1
                    Dim highAroon As Decimal = 0
                    Dim lowAroon As Decimal = 0
                    If counter >= period + 1 Then
                        Dim subPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, runningPayload.Key, period + 1, True)
                        If subPayload IsNot Nothing AndAlso subPayload.Count > 0 Then
                            Dim highestHigh As Decimal = subPayload.Max(Function(x)
                                                                            Return x.Value.High
                                                                        End Function)
                            Dim lowestLow As Decimal = subPayload.Min(Function(x)
                                                                          Return x.Value.Low
                                                                      End Function)
                            Dim highestCandleNumber As Integer = Integer.MinValue
                            Dim lowestCandleNumber As Integer = Integer.MinValue
                            Dim subCounter As Integer = 0
                            For Each runningSubPayload In subPayload
                                subCounter += 1
                                If highestCandleNumber = Integer.MinValue AndAlso runningSubPayload.Value.High = highestHigh Then
                                    highestCandleNumber = subCounter
                                End If
                                If lowestCandleNumber = Integer.MinValue AndAlso runningSubPayload.Value.Low = lowestLow Then
                                    lowestCandleNumber = subCounter
                                End If
                                If highestCandleNumber <> Integer.MinValue AndAlso lowestCandleNumber <> Integer.MinValue Then
                                    Exit For
                                End If
                            Next
                            highAroon = ((period - (period + 1 - highestCandleNumber)) / period) * 100
                            lowAroon = ((period - (period + 1 - lowestCandleNumber)) / period) * 100
                        End If
                    End If
                    If outputHighPayload Is Nothing Then outputHighPayload = New Dictionary(Of Date, Decimal)
                    outputHighPayload.Add(runningPayload.Key, highAroon)
                    If outputLowPayload Is Nothing Then outputLowPayload = New Dictionary(Of Date, Decimal)
                    outputLowPayload.Add(runningPayload.Key, lowAroon)
                Next
            End If
        End Sub
    End Module
End Namespace
