Namespace Indicator
    Public Module FractalBandsTrendLine
        Public Sub CalculateFractalBandsTrendLine(ByVal inputPayload As Dictionary(Of Date, Payload), ByRef outputHighPayload As Dictionary(Of Date, TrendLineVeriables), ByRef outputLowPayload As Dictionary(Of Date, TrendLineVeriables), ByRef fractalHighPayload As Dictionary(Of Date, Decimal), ByRef fractalLowPayload As Dictionary(Of Date, Decimal))
            If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
                Indicator.FractalBands.CalculateFractal(inputPayload, fractalHighPayload, fractalLowPayload)
                For Each runningPayload In inputPayload
                    Dim highLine As TrendLineVeriables = New TrendLineVeriables
                    Dim lowLine As TrendLineVeriables = New TrendLineVeriables

                    Dim lastHighSpikeCandle As Payload = GetSpikeCandleFractalHigh(inputPayload, fractalHighPayload, runningPayload.Key)
                    If lastHighSpikeCandle IsNot Nothing Then
                        Dim lastHighCandle As Payload = GetHighestCandleOutsideFractal(inputPayload, fractalHighPayload, lastHighSpikeCandle.PayloadDate)
                        If lastHighCandle IsNot Nothing Then
                            Dim x1 As Decimal = 0
                            Dim y1 As Decimal = lastHighCandle.High
                            Dim x2 As Decimal = inputPayload.Where(Function(x)
                                                                       Return x.Key > lastHighCandle.PayloadDate AndAlso x.Key <= lastHighSpikeCandle.PayloadDate
                                                                   End Function).Count
                            Dim y2 As Decimal = fractalHighPayload(lastHighSpikeCandle.PayloadDate)

                            Dim trendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(x1, y1, x2, y2)
                            If trendLine IsNot Nothing Then
                                highLine.M = trendLine.M
                                highLine.C = trendLine.C
                                highLine.X = inputPayload.Where(Function(x)
                                                                    Return x.Key > lastHighCandle.PayloadDate AndAlso x.Key <= runningPayload.Value.PayloadDate
                                                                End Function).Count
                            End If
                        End If
                    End If

                    Dim lastLowSpikeCandle As Payload = GetSpikeCandleFractalLow(inputPayload, fractalLowPayload, runningPayload.Key)
                    If lastLowSpikeCandle IsNot Nothing Then
                        Dim lastLowCandle As Payload = GetLowestCandleOutsideFractal(inputPayload, fractalLowPayload, lastLowSpikeCandle.PayloadDate)
                        If lastLowCandle IsNot Nothing Then
                            Dim x1 As Decimal = 0
                            Dim y1 As Decimal = lastLowCandle.Low
                            Dim x2 As Decimal = inputPayload.Where(Function(x)
                                                                       Return x.Key > lastLowCandle.PayloadDate AndAlso x.Key <= lastLowSpikeCandle.PayloadDate
                                                                   End Function).Count
                            Dim y2 As Decimal = fractalLowPayload(lastLowSpikeCandle.PayloadDate)

                            Dim trendLine As TrendLineVeriables = Common.GetEquationOfTrendLine(x1, y1, x2, y2)
                            If trendLine IsNot Nothing Then
                                lowLine.M = trendLine.M
                                lowLine.C = trendLine.C
                                lowLine.X = inputPayload.Where(Function(x)
                                                                   Return x.Key > lastLowCandle.PayloadDate AndAlso x.Key <= runningPayload.Value.PayloadDate
                                                               End Function).Count
                            End If
                        End If
                    End If

                    If outputHighPayload Is Nothing Then outputHighPayload = New Dictionary(Of Date, TrendLineVeriables)
                    outputHighPayload.Add(runningPayload.Key, highLine)
                    If outputLowPayload Is Nothing Then outputLowPayload = New Dictionary(Of Date, TrendLineVeriables)
                    outputLowPayload.Add(runningPayload.Key, lowLine)
                Next
            End If
        End Sub
        Private Function GetHighestCandleOutsideFractal(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal beforeThisTime As Date) As Payload
            Dim ret As Payload = Nothing
            If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = fractalPayload.Where(Function(x)
                                                                                                                 Return x.Key <= beforeThisTime
                                                                                                             End Function)
                If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                    Dim firstCandleTime As Date = Date.MinValue
                    Dim lastCandleTime As Date = Date.MinValue
                    Dim startChecking As Boolean = False
                    For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                     Return x.Key
                                                                                 End Function)
                        If inputPayload(runningPayload.Key).High <= fractalPayload(runningPayload.Key) Then
                            startChecking = True
                        End If
                        If startChecking Then
                            If lastCandleTime = Date.MinValue AndAlso inputPayload(runningPayload.Key).High > fractalPayload(runningPayload.Key) Then
                                lastCandleTime = runningPayload.Key
                            End If
                            Dim previousCandle As Payload = inputPayload(runningPayload.Key).PreviousCandlePayload
                            If lastCandleTime <> Date.MinValue AndAlso previousCandle IsNot Nothing Then
                                If previousCandle.High <= fractalPayload(previousCandle.PayloadDate) Then
                                    firstCandleTime = runningPayload.Key
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If firstCandleTime <> Date.MinValue AndAlso lastCandleTime <> Date.MinValue Then
                        For Each runningPayload In inputPayload
                            If runningPayload.Key >= firstCandleTime AndAlso runningPayload.Key <= lastCandleTime Then
                                If ret Is Nothing Then
                                    ret = runningPayload.Value
                                ElseIf runningPayload.Value.High >= ret.High Then
                                    ret = runningPayload.Value
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            Return ret
        End Function

        Private Function GetSpikeCandleFractalHigh(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal beforeThisTime As Date) As Payload
            Dim ret As Payload = Nothing
            If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = fractalPayload.Where(Function(x)
                                                                                                                 Return x.Key <= beforeThisTime
                                                                                                             End Function)
                If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                    For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                     Return x.Key
                                                                                 End Function)
                        Dim previousPayload As Payload = inputPayload(runningPayload.Key).PreviousCandlePayload
                        If previousPayload IsNot Nothing Then
                            If fractalPayload(runningPayload.Key) < fractalPayload(previousPayload.PayloadDate) Then
                                ret = previousPayload
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            Return ret
        End Function

        Private Function GetLowestCandleOutsideFractal(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal beforeThisTime As Date) As Payload
            Dim ret As Payload = Nothing
            If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = fractalPayload.Where(Function(x)
                                                                                                                 Return x.Key <= beforeThisTime
                                                                                                             End Function)
                If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                    Dim firstCandleTime As Date = Date.MinValue
                    Dim lastCandleTime As Date = Date.MinValue
                    Dim startChecking As Boolean = False
                    For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                     Return x.Key
                                                                                 End Function)
                        If inputPayload(runningPayload.Key).Low >= fractalPayload(runningPayload.Key) Then
                            startChecking = True
                        End If
                        If startChecking Then
                            If lastCandleTime = Date.MinValue AndAlso inputPayload(runningPayload.Key).Low < fractalPayload(runningPayload.Key) Then
                                lastCandleTime = runningPayload.Key
                            End If
                            Dim previousCandle As Payload = inputPayload(runningPayload.Key).PreviousCandlePayload
                            If lastCandleTime <> Date.MinValue AndAlso previousCandle IsNot Nothing Then
                                If previousCandle.Low >= fractalPayload(previousCandle.PayloadDate) Then
                                    firstCandleTime = runningPayload.Key
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If firstCandleTime <> Date.MinValue AndAlso lastCandleTime <> Date.MinValue Then
                        For Each runningPayload In inputPayload
                            If runningPayload.Key >= firstCandleTime AndAlso runningPayload.Key <= lastCandleTime Then
                                If ret Is Nothing Then
                                    ret = runningPayload.Value
                                ElseIf runningPayload.Value.Low <= ret.Low Then
                                    ret = runningPayload.Value
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            Return ret
        End Function

        Private Function GetSpikeCandleFractalLow(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal beforeThisTime As Date) As Payload
            Dim ret As Payload = Nothing
            If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
                Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = fractalPayload.Where(Function(x)
                                                                                                                 Return x.Key <= beforeThisTime
                                                                                                             End Function)
                If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                    For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                     Return x.Key
                                                                                 End Function)
                        Dim previousPayload As Payload = inputPayload(runningPayload.Key).PreviousCandlePayload
                        If previousPayload IsNot Nothing Then
                            If fractalPayload(runningPayload.Key) > fractalPayload(previousPayload.PayloadDate) Then
                                ret = previousPayload
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            Return ret
        End Function
    End Module
End Namespace