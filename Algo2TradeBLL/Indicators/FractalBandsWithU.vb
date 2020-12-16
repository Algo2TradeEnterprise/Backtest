Namespace Indicator
    Public Module FractalBandsWithU
        Public Sub CalculateFractalWithU(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputHighPayload As Dictionary(Of Date, Decimal), ByRef outputLowPayload As Dictionary(Of Date, Decimal),
                                         ByRef highFractalUFormingCandle As Dictionary(Of Date, Date), ByRef highFractalUMiddleCandle As Dictionary(Of Date, Date),
                                         ByRef lowFractalUFormingCandle As Dictionary(Of Date, Date), ByRef lowFractalUMiddleCandle As Dictionary(Of Date, Date))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Dim previousHighFractal As Decimal = 0
                Dim previousLowFractal As Decimal = 0
                Dim highFractal As Decimal = 0
                Dim lowFractal As Decimal = 0

                Dim highU1 As Boolean = True
                Dim highU2 As Boolean = True
                Dim potentialHighUFormingCandle As Date = Date.MinValue
                Dim highUFormingCandle As Date = Date.MinValue
                Dim highUMiddleCandle As Date = Date.MinValue
                Dim lowU1 As Boolean = True
                Dim lowU2 As Boolean = True
                Dim lowUFormingCandle As Date = Date.MinValue
                Dim lowUMiddleCandle As Date = Date.MinValue

                For Each runningPayload In inputPayload.Keys
                    If inputPayload(runningPayload).PreviousCandlePayload IsNot Nothing AndAlso
                        inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                        If inputPayload(runningPayload).PreviousCandlePayload.High < inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High AndAlso
                            inputPayload(runningPayload).High < inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High Then
                            If IsFractalHighSatisfied(inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload, False) Then
                                highFractal = inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.High
                            End If
                        End If
                        If inputPayload(runningPayload).PreviousCandlePayload.Low > inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low AndAlso
                            inputPayload(runningPayload).Low > inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low Then
                            If IsFractalLowSatisfied(inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload, False) Then
                                lowFractal = inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.Low
                            End If
                        End If
                    End If

                    If previousHighFractal <> 0 Then
                        If Not highU1 Then
                            If highFractal = previousHighFractal Then
                                highU1 = True
                            End If
                        ElseIf highU1 AndAlso Not highU2 Then
                            If highFractal > previousHighFractal Then
                                highU2 = True
                                potentialHighUFormingCandle = inputPayload(runningPayload).PreviousCandlePayload.PreviousCandlePayload.PayloadDate
                            ElseIf highFractal < previousHighFractal Then
                                highU1 = False
                            End If
                        ElseIf highU1 AndAlso highU2 Then
                            If highFractal < previousHighFractal Then
                                highUMiddleCandle = inputPayload(runningPayload).PreviousCandlePayload.PayloadDate
                                highUFormingCandle = potentialHighUFormingCandle
                                highU1 = False
                                highU2 = False
                            ElseIf highFractal > previousHighFractal Then
                                highU1 = False
                                highU2 = False
                                potentialHighUFormingCandle = Nothing
                            End If
                        End If
                    End If

                    If highFractalUFormingCandle Is Nothing Then highFractalUFormingCandle = New Dictionary(Of Date, Date)
                    highFractalUFormingCandle.Add(runningPayload, highUFormingCandle)
                    If highFractalUMiddleCandle Is Nothing Then highFractalUMiddleCandle = New Dictionary(Of Date, Date)
                    highFractalUMiddleCandle.Add(runningPayload, highUMiddleCandle)

                    If outputHighPayload Is Nothing Then outputHighPayload = New Dictionary(Of Date, Decimal)
                    outputHighPayload.Add(runningPayload, highFractal)
                    If outputLowPayload Is Nothing Then outputLowPayload = New Dictionary(Of Date, Decimal)
                    outputLowPayload.Add(runningPayload, lowFractal)
                    previousHighFractal = highFractal
                    previousLowFractal = lowFractal
                Next
            End If
        End Sub

        Private Function IsFractalHighSatisfied(ByVal candidateCandle As Payload, ByVal checkOnlyPrevious As Boolean) As Boolean
            Dim ret As Boolean = False
            If candidateCandle IsNot Nothing AndAlso
            candidateCandle.PreviousCandlePayload IsNot Nothing AndAlso
            candidateCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                If checkOnlyPrevious AndAlso candidateCandle.PreviousCandlePayload.High < candidateCandle.High Then
                    ret = True
                ElseIf candidateCandle.PreviousCandlePayload.High < candidateCandle.High AndAlso
                    candidateCandle.PreviousCandlePayload.PreviousCandlePayload.High < candidateCandle.High Then
                    ret = True
                ElseIf candidateCandle.PreviousCandlePayload.High = candidateCandle.High Then
                    ret = IsFractalHighSatisfied(candidateCandle.PreviousCandlePayload, checkOnlyPrevious)
                ElseIf candidateCandle.PreviousCandlePayload.High > candidateCandle.High Then
                    ret = False
                ElseIf candidateCandle.PreviousCandlePayload.PreviousCandlePayload.High = candidateCandle.High Then
                    ret = IsFractalHighSatisfied(candidateCandle.PreviousCandlePayload.PreviousCandlePayload, True)
                ElseIf candidateCandle.PreviousCandlePayload.PreviousCandlePayload.High > candidateCandle.High Then
                    ret = False
                End If
            End If
            Return ret
        End Function
        Private Function IsFractalLowSatisfied(ByVal candidateCandle As Payload, ByVal checkOnlyPrevious As Boolean) As Boolean
            Dim ret As Boolean = False
            If candidateCandle IsNot Nothing AndAlso
            candidateCandle.PreviousCandlePayload IsNot Nothing AndAlso
            candidateCandle.PreviousCandlePayload.PreviousCandlePayload IsNot Nothing Then
                If checkOnlyPrevious AndAlso candidateCandle.PreviousCandlePayload.Low > candidateCandle.Low Then
                    ret = True
                ElseIf candidateCandle.PreviousCandlePayload.Low > candidateCandle.Low AndAlso
                    candidateCandle.PreviousCandlePayload.PreviousCandlePayload.Low > candidateCandle.Low Then
                    ret = True
                ElseIf candidateCandle.PreviousCandlePayload.Low = candidateCandle.Low Then
                    ret = IsFractalLowSatisfied(candidateCandle.PreviousCandlePayload, checkOnlyPrevious)
                ElseIf candidateCandle.PreviousCandlePayload.Low < candidateCandle.Low Then
                    ret = False
                ElseIf candidateCandle.PreviousCandlePayload.PreviousCandlePayload.Low = candidateCandle.Low Then
                    ret = IsFractalLowSatisfied(candidateCandle.PreviousCandlePayload.PreviousCandlePayload, True)
                ElseIf candidateCandle.PreviousCandlePayload.PreviousCandlePayload.Low < candidateCandle.Low Then
                    ret = False
                End If
            End If
            Return ret
        End Function
    End Module
End Namespace