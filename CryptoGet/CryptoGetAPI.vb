Imports System.IO
Imports System.Net
Imports System.Text

Module CryptoGetAPI
    Friend Function CgAPI(url As String) As String
        Dim objResponse As HttpWebResponse
        Dim objReader As StreamReader
        Dim ExistingApiKey As String = GetSetting("CryptoGet", "creds", "api", "")
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
