Imports ExcelDna.Integration
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' SEE https://min-api.cryptocompare.com/documentation and https://github.com/Excel-DNA

Public Module CryptoGetFunctions

    ' =CGPrice("BTC", "USD")
    <ExcelFunction(Description:="Get price")>
    Public Function CGPrice(crypto As String, fiat As String, limit As Integer) As Double
        Try
            Dim json As String = CgAPI("/data/price?fsym=" & crypto & "&tsyms=" & fiat)
            Dim result = JsonConvert.DeserializeObject(json)
            Return result(fiat)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGDailyHistory("BTC", "USD", 10)
    <ExcelFunction(Description:="Get daily historical data (H, L, O, VF, VT, C)")>
    Public Function CGDailyHistory(crypto As String, fiat As String, limit As Integer) As Object
        Try
            Dim json As String = CgAPI("/data/v2/histoday?fsym=" & crypto & "&tsym=" & fiat & "&limit=" & limit)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("Data")
            Dim r(limit - 1, 6) As Double
            For i As Integer = 0 To limit - 1
                r(i, 0) = result(i)("time")
                r(i, 1) = result(i)("high")
                r(i, 2) = result(i)("low")
                r(i, 3) = result(i)("open")
                r(i, 4) = result(i)("volumefrom")
                r(i, 5) = result(i)("volumeto")
                r(i, 6) = result(i)("close")
            Next
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGHourlyHistory("BTC", "USD", 10)
    <ExcelFunction(Description:="Get hourly historical data (H, L, O, VF, VT, C)")>
    Public Function CGHourlyHistory(crypto As String, fiat As String, limit As Integer) As Object
        Try
            Dim json As String = CgAPI("/data/v2/histohour?fsym=" & crypto & "&tsym=" & fiat & "&limit=" & limit)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("Data")
            Dim r(limit - 1, 6) As Double
            For i As Integer = 0 To limit - 1
                r(i, 0) = result(i)("time")
                r(i, 1) = result(i)("high")
                r(i, 2) = result(i)("low")
                r(i, 3) = result(i)("open")
                r(i, 4) = result(i)("volumefrom")
                r(i, 5) = result(i)("volumeto")
                r(i, 6) = result(i)("close")
            Next
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGMinuteHistory("BTC", "USD", 10)
    <ExcelFunction(Description:="Get historical data by minute (H, L, O, VF, VT, C)")>
    Public Function CGMinuteHistory(crypto As String, fiat As String, limit As Integer) As Object
        Try
            Dim json As String = CgAPI("/data/v2/histominute?fsym=" & crypto & "&tsym=" & fiat & "&limit=" & limit)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("Data")
            Dim r(limit - 1, 6) As Double
            For i As Integer = 0 To limit - 1
                r(i, 0) = result(i)("time")
                r(i, 1) = result(i)("high")
                r(i, 2) = result(i)("low")
                r(i, 3) = result(i)("open")
                r(i, 4) = result(i)("volumefrom")
                r(i, 5) = result(i)("volumeto")
                r(i, 6) = result(i)("close")
            Next
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGAddresses("BTC")
    <ExcelFunction(Description:="1) Zero balance all time, 2) Unique all time, 3) New 4) Active")>
    Public Function CGAddresses(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/blockchain/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")
            Dim r(5, 0) As Double
            r(0, 0) = result("id")
            r(1, 0) = result("time")
            r(2, 0) = result("zero_balance_addresses_all_time")
            r(3, 0) = result("unique_addresses_all_time")
            r(4, 0) = result("new_addresses")
            r(5, 0) = result("active_addresses")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGAvgTransactions("BTC")
    <ExcelFunction(Description:="1) Avg value, 2) Count, 3) Count all time, 3) Count large")>
    Public Function CGAvgTransactions(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/blockchain/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")
            Dim r(3, 0) As Double
            r(0, 0) = result("average_transaction_value")
            r(1, 0) = result("transaction_count")
            r(2, 0) = result("transaction_count_all_time")
            r(3, 0) = result("large_transaction_count")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGBlock("BTC")
    <ExcelFunction(Description:="1) Height, 2) Hashrate, 3) Difficulty, 3) Time 4) Size, 5) Supply")>
    Public Function CGBlock(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/blockchain/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")
            Dim r(5, 0) As Double
            r(0, 0) = result("block_height")
            r(1, 0) = result("hashrate")
            r(2, 0) = result("difficulty")
            r(3, 0) = result("block_time")
            r(4, 0) = result("block_size")
            r(5, 0) = result("current_supply")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGBlockDaily("BTC", "10")
    <ExcelFunction(Description:="1) Height, 2) Hashrate, 3) Difficulty, 3) Time 4) Size, 5) Supply")>
    Public Function CGBlockDaily(crypto As String, limit As String) As Object
        Try
            Dim json As String = CgAPI("/data/blockchain/histo/day?fsym=" & crypto & "&limit=" & limit)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("Data")
            Dim r(limit - 1, 6) As Double
            For i As Integer = 0 To limit - 1
                r(i, 0) = result(i)("time")
                r(i, 1) = result(i)("block_height")
                r(i, 2) = result(i)("hashrate")
                r(i, 3) = result(i)("difficulty")
                r(i, 4) = result(i)("block_time")
                r(i, 5) = result(i)("block_size")
                r(i, 6) = result(i)("current_supply")
            Next
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGITBinOutVar("BTC")
    <ExcelFunction(Description:="1) Value, 2) Score, 3) Bear Threshold, 4) Bull Threshold")>
    Public Function CGITBinOutVar(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/tradingsignals/intotheblock/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("inOutVar")
            Dim r(3, 0) As Double
            r(0, 0) = result("value")
            r(1, 0) = result("score")
            r(2, 0) = result("score_threshold_bearish")
            r(3, 0) = result("score_threshold_bullish")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGITBlargetxsVar("BTC")
    <ExcelFunction(Description:="1) Value, 2) Score, 3) Bear Threshold, 4) Bull Threshold")>
    Public Function CGITBlargetxsVar(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/tradingsignals/intotheblock/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("largetxsVar")
            Dim r(3, 0) As Double
            r(0, 0) = result("value")
            r(1, 0) = result("score")
            r(2, 0) = result("score_threshold_bearish")
            r(3, 0) = result("score_threshold_bullish")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGITBaddressesNetGrowth("BTC")
    <ExcelFunction(Description:="1) Value, 2) Score, 3) Bear Threshold, 4) Bull Threshold")>
    Public Function CGITBaddressesNetGrowth(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/tradingsignals/intotheblock/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("addressesNetGrowth")
            Dim r(3, 0) As Double
            r(0, 0) = result("value")
            r(1, 0) = result("score")
            r(2, 0) = result("score_threshold_bearish")
            r(3, 0) = result("score_threshold_bullish")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGITBconcentrationVar("BTC")
    <ExcelFunction(Description:="1) Value, 2) Score, 3) Bear Threshold, 4) Bull Threshold")>
    Public Function CGITBconcentrationVar(crypto As String) As Object
        Try
            Dim json As String = CgAPI("/data/tradingsignals/intotheblock/latest?fsym=" & crypto)
            Dim result = JsonConvert.DeserializeObject(json)("Data")("concentrationVar")
            Dim r(3, 0) As Double
            r(0, 0) = result("value")
            r(1, 0) = result("score")
            r(2, 0) = result("score_threshold_bearish")
            r(3, 0) = result("score_threshold_bullish")
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

    ' =CGCustomData("BTC", "USD", "PRICE|SUPPLY|MKTCAP")
    <ExcelFunction(Description:="Pipe-delimited measures: TYPE | FLAGS | PRICE | LASTUPDATE | MEDIAN | LASTVOLUME | LASTVOLUMETO | LASTTRADEID | VOLUMEDAY | VOLUMEDAYTO | VOLUME24HOUR | VOLUME24HOURTO | OPENDAY | HIGHDAY | LOWDAY | OPEN24HOUR | HIGH24HOUR | LOW24HOUR | VOLUMEHOUR | VOLUMEHOURTO | OPENHOUR | HIGHHOUR | LOWHOUR | TOPTIERVOLUME24HOUR | TOPTIERVOLUME24HOURTO | CHANGE24HOUR | CHANGEPCT24HOUR | CHANGEDAY | CHANGEPCTDAY | CHANGEHOUR | CHANGEPCTHOUR | SUPPLY | MKTCAP | TOTALVOLUME24H | TOTALVOLUME24HTO | TOTALTOPTIERVOLUME24H | TOTALTOPTIERVOLUME24HTO")>
    Public Function CGCustomData(crypto As String, fiat As String, measures As String) As Object
        Try
            Dim json As String = CgAPI("/data/pricemultifull?fsyms=" & crypto & "&tsyms=" & fiat)
            Dim ms() = measures.Split("|")
            Dim result(ms.GetUpperBound(0))
            Dim r(ms.GetUpperBound(0), 0) As Double
            For i As Integer = ms.GetLowerBound(0) To ms.GetUpperBound(0)
                ms(i) = Trim(ms(i))
                result(i) = JsonConvert.DeserializeObject(json)("RAW")(crypto)(fiat)(ms(i))
                r(i, 0) = result(i)
            Next
            Return ArrayResizer.ResizeDoubles(r)
        Catch ex As Exception
            Return ExcelError.ExcelErrorValue
        End Try
    End Function

End Module
