Imports RDotNet

Public Class R_Arima_Forecast
    Public Shared Function InvokeR_ArimaForecast( _
                                                RE As REngine, _
                                                HistoryFromYear As Integer, HistoryFromMonth As Integer, _
                                                ByRef HistoryFactRecords As List(Of HistoryFactRecord), _
                                                ByRef HistoryRegressionInputRecords As List(Of RegressionInputRecord), _
                                                ByRef FutureRegressionInputRecords As List(Of RegressionInputRecord), ForecastHead As Integer) As List(Of ForecastRecord)
        If HistoryFactRecords.Count = 0 Or HistoryRegressionInputRecords.Count = 0 Then Return New List(Of ForecastRecord)
        Dim ForecastRecords As New List(Of ForecastRecord)
        Dim YearMonthVector = RE.CreateCharacterVector(HistoryFactRecords.Count), HistoryFactVector = RE.CreateNumericVector(HistoryFactRecords.Count)
        For idx As Integer = 0 To HistoryFactRecords.Count - 1
            'Dim YearWeekDate As Date = DateAdd(DateInterval.Day, (HistoryFactRecords(idx).Week - 1) * 7, New Date(HistoryFactRecords(idx).Year, 1, 1))
            YearMonthVector(idx) = HistoryFactRecords(idx).HistoryDate.ToString("yyyy-MM-dd")
            HistoryFactVector(idx) = HistoryFactRecords(idx).HistoryValue
        Next

        RE.SetSymbol("YM", YearMonthVector) : RE.SetSymbol("FactValue", HistoryFactVector)
        Dim RegNameSet As New ArrayList
        For colIdx As Integer = 0 To HistoryRegressionInputRecords(0).InputValues.Count - 1
            Dim RegVector = RE.CreateNumericVector(HistoryRegressionInputRecords.Count), RegName As String = "Reg" + (colIdx + 1).ToString()
            For rowIdx As Integer = 0 To HistoryRegressionInputRecords.Count - 1
                RegVector(rowIdx) = HistoryRegressionInputRecords(rowIdx).InputValues(colIdx).Value
            Next
            RE.SetSymbol(RegName, RegVector)
            RegNameSet.Add(RegName)
        Next
        Dim RegNameArray() As String = RegNameSet.ToArray(GetType(String))

        RE.Evaluate("modelfitsample <- data.frame(YM,FactValue," + String.Join(",", RegNameArray) + ")")
        RE.Evaluate("modelfitsample$YM<-as.Date(modelfitsample$YM, '%Y-%m-%d')")
        For regIdx As Integer = 0 To RegNameArray.Length - 1
            RegNameArray(regIdx) = "modelfitsample$" + RegNameArray(regIdx)
        Next
        RE.Evaluate("xreg<-cbind(" + String.Join(",", RegNameArray) + ")")
        RE.Evaluate("myTs <- ts(modelfitsample$FactValue, frequency=12, start=c(" + HistoryFromYear.ToString() + "," + HistoryFromMonth.ToString() + "))")
        Try
            RE.Evaluate("modArima <- auto.arima(myTs, xreg=xreg)")
        Catch ex As EvaluationException
            Console.WriteLine("EvaluationException:" + ex.Message)
            Return ForecastRecords
        End Try


        RegNameSet = New ArrayList()
        For colIdx As Integer = 0 To FutureRegressionInputRecords(0).InputValues.Count - 1
            Dim RegVector = RE.CreateNumericVector(FutureRegressionInputRecords.Count), RegName As String = "FReg" + (colIdx + 1).ToString()
            For rowIdx As Integer = 0 To FutureRegressionInputRecords.Count - 1
                RegVector(rowIdx) = FutureRegressionInputRecords(rowIdx).InputValues(colIdx).Value
            Next
            RE.SetSymbol(RegName, RegVector)
            RegNameSet.Add(RegName)
        Next



        RE.Evaluate("forexreg<-cbind(" + String.Join(",", RegNameSet.ToArray()) + ")")
        RE.Evaluate(String.Format("fore<-forecast(modArima,h={0},xreg=head(forexreg,{0}))", ForecastHead))
        RE.Evaluate("fore<-as.data.frame(fore)")
        Dim foreFrame = RE.Evaluate("fore").AsDataFrame()

        Dim Forecasts = foreFrame(0).AsVector(), Low80 = foreFrame(1).AsVector(), High80 = foreFrame(2).AsVector(), Low95 = foreFrame(3).AsVector(), High95 = foreFrame(4).AsVector()

        Dim ForeFromDate As New Date(HistoryFactRecords(0).Year, 1)

        Dim HistoryLatestDate = From q In HistoryFactRecords Order By q.HistoryDate Descending Take 1

        ForeFromDate = DateAdd(DateInterval.Month, 1, HistoryLatestDate.First.HistoryDate)


        For idx As Integer = 0 To Forecasts.Count - 1
            Dim ForecastRecord1 As New ForecastRecord
            With ForecastRecord1
                .ForecastDate = DateAdd(DateInterval.Month, idx, ForeFromDate)
                '.Year = DateAdd(DateInterval.Day, (idx + 1) * 7, ForeFromDate).Year
                '.Week = DateAdd(DateInterval.Day, (idx + 1) * 7, ForeFromDate).DayOfYear / 7 + 1
                .Forecast = Forecasts(idx) : .Low80 = Low80(idx) : .High80 = High80(idx) : .Low95 = Low95(idx) : .High95 = High95(idx)
            End With
            ForecastRecords.Add(ForecastRecord1)
        Next

        Return ForecastRecords
    End Function

    Public Class HistoryFactRecord
        Public Property Year As Integer : Public Property Week As Integer : Public Property HistoryDate As Date : Public Property HistoryValue As Object
        Public Sub New(y As Integer, w As Integer, v As Object)
            Me.Year = y : Me.Week = w : Me.HistoryValue = v
        End Sub

        Public Sub New(dt As Date, v As Object)
            Me.HistoryDate = dt : Me.HistoryValue = v
        End Sub

        Public Sub New()
            Me.Year = 0 : Me.Week = 0 : Me.HistoryValue = Nothing : Me.HistoryDate = Date.MaxValue
        End Sub
    End Class

    Public Class RegressionInputRecord
        Public Property Year As Integer : Public Property Week As Integer : Public Property HistoryDate As Date

        Public Property InputValues As List(Of ColNameValue)
        Public Sub New()
            InputValues = New List(Of ColNameValue)
        End Sub
    End Class

    Public Class ColNameValue
        Public Property ColName As String : Public Property Value As Object
        Public Sub New(name As String, v As Object)
            Me.ColName = name : Me.Value = v
        End Sub

        Public Sub New()
            Me.ColName = String.Empty : Me.Value = Nothing
        End Sub
    End Class

End Class
