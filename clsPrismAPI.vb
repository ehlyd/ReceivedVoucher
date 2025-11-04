Imports System.Net
Imports System.Text
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json

Public Class clsPrismAPI

    Dim authNonce As String, authNonceResponse As String

    Dim intConnectTries As Integer
    Dim invPostedDate As String


    Public Function IsAPI_LoginSuccessfull() As Boolean
        Try

ReTry:      authNonce = GetAuthNonce()

            If authNonce <> "" Then

                authNonceResponse = Get_AuthNonceResponse(authNonce)

                authSession = GetAuthSession()

                If authSession.Contains("(401) Unauthorized") Then

                    If intConnectTries < 3 Then

                        GoTo ReTry
                    Else
                        'Return "Unable to get Auth-Session"
                        Throw New Exception("Unable to get Auth-Session after 3 tries. Please check the API credentials.")
                    End If

                Else

                    Dim str1stSession As String
                    str1stSession = Get1stSession()

                    If str1stSession.Contains("200") Then
                        Dim strWebClient As String
                        strWebClient = GetWebClient()

                        If strWebClient.Contains("200") Then
                            Dim str3rdSession As String
                            str3rdSession = Get3rdSession()

                            If str3rdSession.Contains("200") Then

                                'Return "Authorized"
                                Return True

                            Else
                                'Return "Unable to get third session"
                                Throw New Exception("Unable to get third session.")
                            End If

                        Else
                            'Return "Unable to get Workstation"
                            Throw New Exception("Unable to get Workstation.")
                        End If

                    Else
                        'Return "Unable to get first session"
                        Throw New Exception("Unable to get first session.")
                    End If

                End If

            Else
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Function GetAuthSession() As String
        Try
            intConnectTries += 1
            Dim url As String = APIUrl & "/v1/rest/auth?usr=" & APIUserName & "&pwd=" & APIPassword

            Dim request As HttpWebRequest = WebRequest.Create(url)

            request.Method = "GET"
            request.Headers.Add("Auth-Nonce", authNonce)
            request.Headers.Add("Auth-Nonce-Response", authNonceResponse)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim response As HttpWebResponse = request.GetResponse()

            If response.StatusCode = HttpStatusCode.OK Then
                Dim header As String = response.Headers.Get("Auth-Session")
                Return header
            Else
                Return ""
            End If

            response.Close()

        Catch ex As Exception
            If ex.Message.Contains("(401) Unauthorized") Then
                Return "(401) Unauthorized"
            Else
                Throw ex
            End If

        End Try

    End Function

    Function Get1stSession() As String
        Dim url As String = APIUrl & "/v1/rest/session"

        Try
            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Headers.Add("Auth-Session", authSession)
            request.Method = "GET"

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim response As HttpWebResponse = request.GetResponse()

            If response.StatusCode = HttpStatusCode.OK Then
                Return response.StatusCode
            Else
                Return "Error " & response.StatusCode
            End If

            response.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function GetWebClient() As String
        Dim url As String = APIUrl & "/v1/rest/sit?ws=" & currWokstation

        Try
            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Headers.Add("Auth-Session", authSession)
            request.Method = "GET"

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim response As HttpWebResponse = request.GetResponse()

            If response.StatusCode = HttpStatusCode.OK Then
                Return response.StatusCode
            Else
                Return "Error " & response.StatusCode
            End If

            response.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function Get3rdSession() As String
        Dim url As String = APIUrl & "/v1/rest/session"

        Try
            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Headers.Add("Auth-Session", authSession)
            request.Method = "GET"

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim response As HttpWebResponse = request.GetResponse()

            If response.StatusCode = HttpStatusCode.OK Then

                Dim responseStream As Stream = response.GetResponseStream()
                Dim reader As New StreamReader(responseStream)

                ' Read response content
                Dim responseContent As String = reader.ReadToEnd()

                Dim responseObject As Object = JsonConvert.DeserializeObject(responseContent)

                For Each dataObject In responseObject
                    currStoreSid = Replace(dataObject("storesid").ToString, """", "")
                    currStoreCode = Replace(dataObject("storecode").ToString, """", "")
                    currSBSNo = Replace(dataObject("subsidiarynumber").ToString, """", "")
                Next

                Return response.StatusCode
            Else
                Return "Error " & response.StatusCode
            End If

            response.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Function GetAuthNonce() As String
    '    Try

    '        Dim url As String = APIUrl & "/v1/rest/auth"
    '        Dim request As WebRequest = HttpWebRequest.Create(url)

    '        ' Set request method
    '        request.Method = "GET"

    '        Dim response As WebResponse = request.GetResponse()
    '        Dim headers As WebHeaderCollection = response.Headers

    '        ' Check if Auth-Nonce header exists
    '        If headers.AllKeys.Contains("Auth-Nonce") Then
    '            Return headers("Auth-Nonce")
    '        Else
    '            Return ""
    '        End If

    '        response.Close()

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Function GetAuthNonce() As String
        Try

            Dim url As String = APIUrl & "/v1/rest/auth"
            Dim request As HttpWebRequest = WebRequest.Create(url)

            request.Method = "GET"
            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim response As HttpWebResponse = request.GetResponse()
            If response.StatusCode = HttpStatusCode.OK Then
                Dim header As String = response.Headers.Get("Auth-Nonce")
                Return header
            Else
                Return ""
            End If

            response.Close()

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function Get_AuthNonceResponse(ByVal lngAuthNonce As Long) As Long
        Try
            Dim dbldiv13 As Double, mod59 As Double

            dbldiv13 = Math.Round(lngAuthNonce / 13, 0)
            mod59 = dbldiv13 Mod 99999

            Return mod59 * 17

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ReadJsonFromFile(ByVal filePath As String) As String
        Try
            Using reader As New StreamReader(filePath)
                Return reader.ReadToEnd()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ReadResponseStream(ByVal response As HttpWebResponse) As String
        Try
            Dim reader As New StreamReader(response.GetResponseStream())
            Dim responseString As String = reader.ReadToEnd()
            reader.Close()
            Return responseString
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function ChangeCurrent_StoreSID(ByVal strSBS_SID As String, ByVal strStoreSID As String) As Boolean
        Try

            Dim json As String = "{""subsidiarysid"":""" & strSBS_SID & """,""storesid"":""" & strStoreSID & """}"

            Dim url As String = APIUrl & "/api/security/altersession?action=changevalues"

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            WriteToFile("Changing store SID to " & strStoreSID & " with subsidiary SID " & strSBS_SID & "...")

            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim responseString As String = ""

            If response.StatusCode = HttpStatusCode.OK Then

                responseString = ReadResponseStream(response)
                Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                For Each dataObject In responseObject
                    currStoreSid = Replace(dataObject("storesid").ToString, """", "")
                    currStoreCode = dataObject("storecode").ToString
                    currSBSNo = dataObject("subsidiarynumber").ToString
                Next

                WriteToFile("Current store SID was changed successfully.")

                Return True
            Else
                Return False
                WriteToFile("Error changing current store SID: " & response.StatusCode & " " & response.StatusDescription)
            End If

            response.Close()

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function IsGenerateVoucher_Successfull(ByVal EmpSID As String, ByVal VoucherSID As String) As Boolean

        Try

            Dim json As String = "{""data"":[{""clerksid"":""" & EmpSID & """,""asnsidlist"":""" & VoucherSID & """,""doupdatevoucher"":false,""originapplication"":""RProPrismWeb""}]}"

            Console.WriteLine(json)

            Dim url As String = APIUrl & "/api/backoffice/receiving?action=convertasntovoucher"

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("Voucher generated successfully.")
                    Return True

                Else

                    WriteToFile("Error generating voucher: " & response.StatusCode & " " & response.StatusDescription)
                    Return False

                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error generating voucher: " & webex.Message)
                Return False
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub ReceiveVoucherItem(VoucherSID As String, VoucherItemSID As String, rowVersion As Int16, Barcode As String, Qty As Int16)
        Try

            Dim json As String = "{
                                    ""data"":
                                    [
                                        {""rowversion"":" & rowVersion & ",
                                        ""qty"":" & Qty & ",
                                        ""upc"":""" & Barcode & """
                                        }
                                    ]
                                }"

            Dim url As String = APIUrl & "/api/backoffice/receiving/" & VoucherSID & "/recvitem/" & VoucherItemSID & "?filter=rowversion,eq," & rowVersion

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "PUT"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("Qty received.")

                Else
                    WriteToFile("Error receiving voucher item: " & response.StatusCode & " " & response.StatusDescription)
                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error receiving voucher item: " & webex.Message)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub AddExtraItem(VoucherSID As String, VoucherItemSID As String, Barcode As String, Qty As Int16, StyleNo As String, Price As Double)
        Try

            Dim json As String = "{
                                      ""data"": [
                                        {
                                          ""recvItem"": {
                                            ""resource"": ""recvitem"",
                                            ""endpoint"": ""backoffice/receiving/:vousid/recvitem/:sid/"",
                                            ""dirty"": {},
                                            ""originapplication"": ""RProPrismWeb"",
                                            ""itemsid"": """ & VoucherItemSID & """,
                                            ""qty"": " & Qty & ",
                                            ""upc"": " & Barcode & ",
                                            ""description1"": """ & StyleNo & """,
                                            ""serialno"": null,
                                            ""lotnumber"": null,
                                            ""serialtype"": 0,
                                            ""lottype"": 0,
                                            ""price"": " & Price & ",
                                            ""vousid"": """ & VoucherSID & """
                                          }
                                        }
                                      ]
                                    }"


            Dim url As String = APIUrl & "/api/backoffice/receiving/" & VoucherSID & "?action=AddConsolidateVouItem"

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("Extra item added.")

                Else
                    WriteToFile("Error adding extra item: " & response.StatusCode & " " & response.StatusDescription)
                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error adding extra item: " & webex.Message)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub ApproveVoucher(VoucherSID As String, EmpSID As String, rowVersion As Integer, ApproveDate As DateTime)
        Try

            Dim customFormat As String = "yyyy-MM-dd'T'HH:mm:ss.fff'Z'"
            'Dim myDate As DateTime = DateTime.UtcNow
            'Dim strApprovedDate As String = myDate.ToString(customFormat)

            Dim strApprovedDate As String = ApproveDate.ToString(customFormat)

            Dim json As String = "{
                                    ""data"":
                                    [
                                        {""rowversion"":" & rowVersion & ",
                                        ""status"":4,
                                        ""held"":0,
                                        ""approvbysid"":""" & EmpSID & """,
                                        ""approvdate"":""" & strApprovedDate & """,
                                        ""approvstatus"":2,
                                        ""publishstatus"":2
                                        }
                                    ]
                                }"

            Dim url As String = APIUrl & "/api/backoffice/receiving/" & VoucherSID

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "PUT"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("Voucher approved.")

                Else
                    WriteToFile("Error approving voucher: " & response.StatusCode & " " & response.StatusDescription)
                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error approving voucher: " & webex.Message)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub UpdatePONote(POSID As String, rowVersion As Integer)
        Try



            Dim json As String = "{
                                    ""data"":
                                    [
                                        {
                                            ""rowversion"":" & rowVersion & ",
                                            ""note"":""C""
                                        }
                                    ]
                                }"

            Dim url As String = APIUrl & "/api/backoffice/purchaseorder/" & POSID

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "PUT"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("PO note updated to ""C"".")

                Else
                    WriteToFile("Error updating PO note: " & response.StatusCode & " " & response.StatusDescription)
                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error updating PO Note: " & webex.Message)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub AddVoucherComments(VoucherSID As String, Comment As String)
        Try

            Dim json As String = "{
                                  ""data"": [
                                    {
                                      ""originapplication"": ""RProPrismWeb"",
                                      ""comments"": """ & Comment & """,
                                      ""vousid"": """ & VoucherSID & """
                                    }
                                  ]
                                }"

            Dim url As String = APIUrl & "/api/backoffice/receiving/" & VoucherSID & "/recvcomment?comments=" & Replace(Comment, " ", "+")

            Dim request As HttpWebRequest = WebRequest.Create(url)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Accept = "application/json, text/plain, version=2"
            request.Headers.Add("Auth-Session", authSession)

            request.ServicePoint.ConnectionLimit = 10
            request.ServicePoint.MaxIdleTime = 5 * 1000
            request.Timeout = 60000

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(json)
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            Try

                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseString As String = ""

                If response.StatusCode = HttpStatusCode.OK Then

                    responseString = ReadResponseStream(response)
                    Dim responseObject As Object = JsonConvert.DeserializeObject(responseString)

                    WriteToFile("Comments added successfully.")

                Else
                    WriteToFile("Error adding comments to voucher: " & response.StatusCode & " " & response.StatusDescription)
                End If

                response.Close()

            Catch webex As WebException
                WriteToFile("Error adding comments to voucher: " & webex.Message)
            End Try

        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class
