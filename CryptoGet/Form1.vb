Public Class Form1
    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim HTML As String = "
<html>
<head>
<style type='text/css'>
    body {padding: 16px; background: #f5f5f5; color: #1e1e1e; margin: 0; font-family: Calibri, 'Times Romain'}
    code {color: #000; font-size: 16px;}
</style>
</head>
<body>
<h3>CryptoGet Add-In</h3>
<p>This add-in relies on cryptocompare.com API. Visit their API pages to obtain API key (free with some limitations) and for detailed explanation with regards to provided data.
The functions below are mostly self-adjustable array functions, so you don't need Ctrl-Shift-Return while entering them. For modifications/support, contact author at sundayheap.com.
</p>
<p>All date/time values are provided at source as Unix timestamps. To convert to Excel date, use:<br>
<strong>=( UnixTimestamp - DATE(1970,1,1)) * 86400</strong><br><br><br>

<code>=CGPrice(""BTC"", ""USD"")</code><br>
<ul>
<li>Price</li>
</ul>

<code>=CGDailyHistory(""BTC"", ""USD"", 10)</code><br>
<ul>
<li>Time</li>
<li>High</li>
<li>Low</li>
<li>Open</li>
<li>VolumeFrom</li>
<li>VolumeTo</li>
<li>Close</li>
</ul>

<code>=CGHourlyHistory(""BTC"", ""USD"", 10)</code><br>
<ul>
<li>Time</li>
<li>High</li>
<li>Low</li>
<li>Open</li>
<li>VolumeFrom</li>
<li>VolumeTo</li>
<li>Close</li>
</ul>

<code>=CGMinuteHistory(""BTC"", ""USD"", 10)</code><br>
<ul>
<li>Time</li>
<li>High</li>
<li>Low</li>
<li>Open</li>
<li>VolumeFrom</li>
<li>VolumeTo</li>
<li>Close</li>
</ul>

<code>=CGAddresses(""BTC"")</code><br>
<p></p>
<ul>
<li>ID</li>
<li>Time</li>
<li>Zero balance addresses all time</li>
<li>Unique addresses all time</li>
<li>New addresses</li>
<li>Active addresses</li>
</ul>

<code>=CGAvgTransactions(""BTC"")</code><br>
<p></p>
<ul>
<li>Average transaction value</li>
<li>Transaction count</li>
<li>Transaction count all time</li>
<li>Large transaction count</li>
</ul>

<code>=CGBlock(""BTC"")</code><br>
<p></p>
<ul>
<li>Block height</li>
<li>Hash rate</li>
<li>Difficulty</li>
<li>Block time</li>
<li>Block size</li>
<li>Current supply</li>
</ul>

<code>=CGBlockDaily(""BTC"", ""10"")</code><br>
<p></p>
<ul>
<li>Time</li>
<li>Block height</li>
<li>Hash rate</li>
<li>Difficulty</li>
<li>Block time</li>
<li>Block size</li>
<li>Current supply</li>
</ul>

<code>=CGITBinOutVar(""BTC"")</code><br>
<ul>
<li>Value</li>
<li>Score</li>
<li>Score threashold bearish</li>
<li>Score threashold bullish</li>
</ul>

<code>=CGITBlargetxsVar(""BTC"")</code><br>
<ul>
<li>Value</li>
<li>Score</li>
<li>Score threashold bearish</li>
<li>Score threashold bullish</li>

</ul>

<code>=CGITBaddressesNetGrowth(""BTC"")</code><br>
<ul>
<li>Value</li>
<li>Score</li>
<li>Score threashold bearish</li>
<li>Score threashold bullish</li>
</ul>

<code>=CGITBconcentrationVar(""BTC"")</code><br>
<ul>
<li>Value</li>
<li>Score</li>
<li>Score threashold bearish</li>
<li>Score threashold bullish</li>
</ul>

<code>=CGCustomData(""BTC"", ""USD"", ""PRICE|SUPPLY|MKTCAP"")</code><br>
<p></p>
<p>Pipe-delimited measures: TYPE | FLAGS | PRICE | LASTUPDATE | MEDIAN | LASTVOLUME | LASTVOLUMETO | LASTTRADEID | VOLUMEDAY | VOLUMEDAYTO | VOLUME24HOUR | VOLUME24HOURTO | OPENDAY | HIGHDAY | LOWDAY | OPEN24HOUR | HIGH24HOUR | LOW24HOUR | VOLUMEHOUR | VOLUMEHOURTO | OPENHOUR | HIGHHOUR | LOWHOUR | TOPTIERVOLUME24HOUR | TOPTIERVOLUME24HOURTO | CHANGE24HOUR | CHANGEPCT24HOUR | CHANGEDAY | CHANGEPCTDAY | CHANGEHOUR | CHANGEPCTHOUR | SUPPLY | MKTCAP | TOTALVOLUME24H | TOTALVOLUME24HTO | TOTALTOPTIERVOLUME24H | TOTALTOPTIERVOLUME24HTO</p>
</body>
</html>




"
        WebBrowser1.DocumentText = "0"
        WebBrowser1.Document.OpenNew(True)
        WebBrowser1.Document.Write(HTML)
        WebBrowser1.Refresh()
    End Sub

    Private Sub Form1_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        WebBrowser1.Select()
    End Sub

    Private Sub btnSetApiKey_Click(sender As Object, e As EventArgs) Handles btnSetApiKey.Click
        Dim ExistingApiKey As String = GetSetting("CryptoGet", "creds", "api", "")
        Dim ApiKey As String = InputBox("Enter API key as obtained from cryptocompare.com:", "API Key", ExistingApiKey)
        If ApiKey.Length() > 6 Then
            SaveSetting("CryptoGet", "creds", "api", ApiKey)
        End If
    End Sub
End Class