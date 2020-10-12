Imports System.IO
Imports System.Net
Imports System.Text

Module CryptoGetAPI
    Friend Function cgAPI(url As String) As String
        Dim objResponse As HttpWebResponse
        Dim objReader As StreamReader
        Dim ExistingApiKey As String = GetSetting("CryptoGet", "creds", "api", "583fdb583d69827c15db099308608c97abe5b09dca32be4de94dfe6e39fbbe9f")
        Dim myRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("https://min-api.cryptocompare.com" & url), HttpWebRequest)
        myRequest.Headers.Add("Authorization", "Basic " + ExistingApiKey)
        Try
            objResponse = DirectCast(myRequest.GetResponse(), HttpWebResponse)
            objReader = New StreamReader(objResponse.GetResponseStream())
            Return objReader.ReadToEnd()
        Catch ex As WebException
            Return ex.Status
        End Try
    End Function

End Module
