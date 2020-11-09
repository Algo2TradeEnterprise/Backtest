Namespace Indicator
    Public Module WMA
        Public Sub CalculateWMA(ByVal wmaPeriod As Integer, ByVal wmaField As Payload.PayloadFields, ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputPayload As Dictionary(Of Date, Decimal))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                For Each runningPayload In inputPayload
                    Dim wmaValue As Decimal = 0

                    Dim previousNInputFieldPayload As List(Of KeyValuePair(Of Date, Payload)) = Common.GetSubPayload(inputPayload, runningPayload.Key, wmaPeriod, True)
                    If previousNInputFieldPayload IsNot Nothing AndAlso previousNInputFieldPayload.Count > 0 Then
                        Dim ctr As Integer = 0
                        For Each runningSubPayload In previousNInputFieldPayload.OrderBy(Function(x)
                                                                                             Return x.Key
                                                                                         End Function)
                            ctr += 1
                            Select Case wmaField
                                Case Payload.PayloadFields.Close
                                    wmaValue += runningSubPayload.Value.Close * ctr
                                Case Payload.PayloadFields.High
                                    wmaValue += runningSubPayload.Value.High * ctr
                                Case Payload.PayloadFields.Low
                                    wmaValue += runningSubPayload.Value.Low * ctr
                                Case Payload.PayloadFields.Open
                                    wmaValue += runningSubPayload.Value.Open * ctr
                                Case Payload.PayloadFields.Volume
                                    wmaValue += runningSubPayload.Value.Volume * ctr
                                Case Payload.PayloadFields.Additional_Field
                                    wmaValue += runningSubPayload.Value.Additional_Field * ctr
                                Case Else
                                    Throw New NotImplementedException
                            End Select
                        Next

                        wmaValue = wmaValue / (ctr * (ctr + 1) / 2)
                    End If

                    If outputPayload Is Nothing Then outputPayload = New Dictionary(Of Date, Decimal)
                    outputPayload.Add(runningPayload.Key, wmaValue)
                Next
            End If
        End Sub

    End Module
End Namespace