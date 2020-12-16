Namespace Indicator
    Public Module ChaikinMoneyFlow
        Public Sub CalculateCMF(ByVal Period As Integer, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Decimal), Optional neglectValidation As Boolean = False)
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                outputPayload = New Dictionary(Of Date, Decimal)
                Dim moneyFlowVolumeData As Dictionary(Of Date, Decimal) = Nothing
                For Each runningInputPayload In inputPayload
                    Dim moneyFlowMultiplier As Decimal = Decimal.MinValue
                    If (runningInputPayload.Value.High - runningInputPayload.Value.Low) <> 0 Then
                        moneyFlowMultiplier = ((runningInputPayload.Value.Close - runningInputPayload.Value.Low) - (runningInputPayload.Value.High - runningInputPayload.Value.Close)) / (runningInputPayload.Value.High - runningInputPayload.Value.Low)
                    Else
                        moneyFlowMultiplier = ((runningInputPayload.Value.Close - runningInputPayload.Value.Low) - (runningInputPayload.Value.High - runningInputPayload.Value.Close))
                    End If
                    Dim moneyFlowVolume As Decimal = moneyFlowMultiplier * runningInputPayload.Value.Volume
                    If moneyFlowVolumeData Is Nothing Then moneyFlowVolumeData = New Dictionary(Of Date, Decimal)
                    moneyFlowVolumeData.Add(runningInputPayload.Key, moneyFlowVolume)

                    Dim previousNInputFieldPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, runningInputPayload.Key, Period, True)
                    If previousNInputFieldPayload IsNot Nothing AndAlso previousNInputFieldPayload.Count > 0 Then
                        Dim previousNCMFFieldPayload As List(Of KeyValuePair(Of Date, Decimal)) = Common.GetSubPayload(moneyFlowVolumeData, runningInputPayload.Key, Period, True)
                        Dim totalVolume As Long = previousNInputFieldPayload.Sum(Function(x) x.Value.Volume)
                        If totalVolume <> 0 Then
                            Dim totalMoneyFlowVolume As Decimal = previousNCMFFieldPayload.Sum(Function(x) x.Value)
                            outputPayload.Add(runningInputPayload.Key, totalMoneyFlowVolume / totalVolume)
                        Else
                            outputPayload.Add(runningInputPayload.Key, moneyFlowVolume)
                        End If
                    Else
                        outputPayload.Add(runningInputPayload.Key, moneyFlowVolume)
                    End If
                Next
            End If
        End Sub
    End Module
End Namespace