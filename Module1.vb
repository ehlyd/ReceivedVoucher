Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Newtonsoft.Json
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Module Module1

    Public APIUrl, APIUserName, APIPassword As String

    Public strServer, strDatabase, strUserID, strPswrd As String
    Public strRPDataSource, strRPUserID, strRPPswrd As String
    Public strEBSDataSource, strEBSUserID, strEBSPswrd As String

    Public strEmailSender, strEmailRecipient, strEmailPswrd As String
    Public strEmailReceivedVoucher, strEmailReceivedVoucherCC As String
    Public objStream As StreamWriter

    Public currWokstation, currStoreSid, currStoreCode, currSBSNo, strDBStoreCode, strUseDBStoreCode As String
    Public authSession As String

    Sub Main()

        Dim configFile As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim appSettings As AppSettingsSection = configFile.Sections("appSettings")

        With appSettings
            strServer = .Settings("SQLServer").Value
            strDatabase = .Settings("SQLDatabase").Value
            strUserID = .Settings("SQLUserID").Value
            strPswrd = .Settings("SQLPswrd").Value
            currWokstation = .Settings("Workstation").Value
        End With

        Dim mclsSQL As New clsSQLDB, dt As DataTable
        mclsSQL.OpenDB()
        dt = mclsSQL.GetDataSet("select * from BrandIntegration_Settings where upper(IntegrationName)='RECEIVED VOUCHER BY PRISM API' and upper(BRAND)='JACADI'").Tables(0)
        If dt.Rows.Count <> 0 Then
            Dim setValue As String

            For Each dRow As DataRow In dt.Rows
                setValue = IIf(IsDBNull(dRow.Item("SettingValue")), "", dRow.Item("SettingValue"))

                Select Case UCase(dRow.Item("SettingName"))

                    Case "ORACLEDATASOURCE"
                        strEBSDataSource = setValue

                    Case "ORACLEUSERID"
                        strEBSUserID = setValue

                    Case "ORACLEPSWRD"
                        strEBSPswrd = setValue

                    Case "RP_DATASOURCE"
                        strRPDataSource = setValue

                    Case "RP_USERID"
                        strRPUserID = setValue

                    Case "RP_PSWRD"
                        strRPPswrd = setValue

                    Case "EMAIL_SENDER"
                        strEmailSender = setValue

                    Case "EMAIL_PASSWORD"
                        strEmailPswrd = setValue

                    Case "EMAIL_RECIPIENT"
                        strEmailRecipient = setValue

                    Case "API_URL"
                        APIUrl = setValue

                    Case "API_USERNAME"
                        APIUserName = setValue

                    Case "API_PASSWORD"
                        APIPassword = setValue

                    Case "EMAIL_RECEIVEDVOUCHER"
                        strEmailReceivedVoucher = setValue

                    Case "EMAIL_RECEIVEDVOUCHER_CC"
                        strEmailReceivedVoucherCC = setValue

                    Case "STORECODE"
                        strDBStoreCode = setValue

                    Case "CHANGE_STORECODE"
                        strUseDBStoreCode = dRow.Item("SettingValue")

                End Select
            Next

        End If
        mclsSQL.CloseDB()


        Dim f1 As New Form1
        f1.Show()
    End Sub

    Public Function GetCustomerSID() As String
        Try
            Dim dt As DataTable
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()
            dt = mclsOra.GetDataSet("select sid from rps.customer where active=1 and cust_type=0 and rownum=1 order by created_datetime").Tables(0)
            If dt.Rows.Count <> 0 Then
                Return dt.Rows(0).Item(0)
            Else
                Return ""
            End If
            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetItemSID(ByVal strBarcode As String) As String
        Try
            Dim dt As DataTable
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()
            dt = mclsOra.GetDataSet("SELECT sid FROM RPS.INVN_SBS_ITEM	WHERE UPC='" & strBarcode & "'").Tables(0)
            If dt.Rows.Count <> 0 Then
                Return dt.Rows(0).Item(0)
            Else
                Return ""
            End If
            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub Create_LOGFile()
        If Dir(Application.StartupPath & "\LOG", vbDirectory) = vbNullString Then
            MkDir(Application.StartupPath & "\LOG")
        End If
        objStream = New StreamWriter(Application.StartupPath & "\LOG\AppLog_" & Format(Now, "yyyyMMdd") & ".txt", True)

    End Sub

    Public Sub WriteToFile(ByVal tmpStr As String)
        If objStream Is Nothing Then
            objStream = New StreamWriter(Application.StartupPath & "\LOG\AppLog_" & Format(Now, "yyyyMMdd") & ".txt", True)
        End If

        objStream.WriteLine(Now.ToLongTimeString & " " & tmpStr)
        objStream.Close()
        objStream = Nothing
    End Sub

    Public Sub SendEmail(ByVal toAddress As String, ByVal subject As String, Optional ByVal strFileAttachment As String = "", Optional ByVal tblHTML As StringBuilder = Nothing, Optional ByVal strEmailCC As String = "", Optional ByVal EmailBCC As String = "")

        Try

            Dim strHtmlLineFeed As New StringBuilder()
            strHtmlLineFeed.AppendLine("<br>")

            Dim message As New MailMessage()
            message.From = New MailAddress(strEmailSender)
            message.To.Add(toAddress)

            If strEmailCC <> "" Then message.CC.Add(strEmailCC)
            If EmailBCC <> "" Then message.Bcc.Add(EmailBCC)

            'message.To.Add("idabu@aseelah.com")
            'message.To.Add(" idabu@aseelah.com, akhan@alaseel.com")

            message.Subject = subject
            message.IsBodyHtml = True

            If strFileAttachment <> "" Then
                Dim objAttachment As Attachment
                objAttachment = New Attachment(strFileAttachment)
                message.Attachments.Add(objAttachment)

                strFileAttachment = ""
            End If

            Dim msgBody As String = ""

            If subject.Contains("Successful") Then

                msgBody = "Below is the list of successful voucher receiving from Salasa Replenishment" & tblHTML.ToString()
                'msgBody = msgBody & strHtmlLineFeed.ToString & "Details are also available in the attached file." & tblHTML.ToString()

            Else

                msgBody = "Attached is a list of replenishments with SKUs not in the system."

            End If

            msgBody = msgBody & strHtmlLineFeed.ToString & "Automated email - please do not reply." & strHtmlLineFeed.ToString

            message.Body = msgBody

            Dim smtpClient As New SmtpClient("smtp-mail.outlook.com", 587)
            smtpClient.Credentials = New NetworkCredential(strEmailSender, strEmailPswrd)
            smtpClient.EnableSsl = True

            smtpClient.Send(message)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub SendEmail_AppError(ByVal body As String)

        Dim subject As String

        Dim smtpClient As New SmtpClient("smtp-mail.outlook.com", 587)
        smtpClient.Credentials = New NetworkCredential(strEmailSender, strEmailPswrd)
        smtpClient.EnableSsl = True

        subject = "Error in Prism API Voucher Receiving"

        'strEmailRecipient = "idabu@aseelah.com"

        Dim mailMessage As New MailMessage(strEmailSender, strEmailRecipient, subject, body)
        Try
            smtpClient.Send(mailMessage)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub ExportToExcel_EPPlus(dtExcelData As DataTable, strFilePath As String)

        Try
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial

            Using package As New ExcelPackage()
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Sheet1")

                'worksheet.Cells.Style.Font.Name = "Arial" ' Or "Calibri", "Tahoma", "Segoe UI"
                'worksheet.Cells.Style.Font.Name = "Segoe UI"

                For col As Integer = 0 To dtExcelData.Columns.Count - 1
                    worksheet.Cells(1, col + 1).Value = dtExcelData.Columns(col).ColumnName
                    worksheet.Cells(1, col + 1).Style.Font.Bold = True
                    worksheet.Cells(1, col + 1).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(1, col + 1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray)
                Next

                For row As Integer = 0 To dtExcelData.Rows.Count - 1
                    For col As Integer = 0 To dtExcelData.Columns.Count - 1

                        worksheet.Cells(row + 2, col + 1).Value = dtExcelData.Rows(row)(col)

                    Next

                Next

                worksheet.Columns.AutoFit()

                package.SaveAs(New FileInfo(strFilePath))

                package.Dispose()

            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function CreateHTMLTable(ByVal dataTable As DataTable) As StringBuilder
        Try
            Dim tableHtml As StringBuilder = New StringBuilder()
            tableHtml.AppendLine("<br>")
            tableHtml.AppendLine("<br>")
            tableHtml.AppendLine("<table border=""1""")
            tableHtml.AppendLine("<tbody>")

            ' Build table header row
            tableHtml.AppendLine("<tr>")
            For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                tableHtml.AppendFormat("<th>{0}</th>", dataTable.Columns(colIndex).ColumnName)
            Next
            tableHtml.AppendLine("</tr>")

            ' Build table data rows
            For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                tableHtml.AppendLine("<tr>")
                For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                    tableHtml.AppendFormat("<td style='text-align:center; vertical-align:middle'>{0}</td>", dataTable.Rows(rowIndex).Item(colIndex))
                Next
                tableHtml.AppendLine("</tr>")
            Next

            tableHtml.AppendLine("</tbody>")
            tableHtml.AppendLine("</table>")
            tableHtml.AppendLine("<br>")
            tableHtml.AppendLine("<br>")

            Return tableHtml
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Module
