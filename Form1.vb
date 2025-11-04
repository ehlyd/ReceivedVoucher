Imports System.Net
Imports System.Net.Http.Headers
Imports System.Text

Public Class Form1

    'Private Function GenerateASNVoucher(PONo As String, ASN_No As String, PkgNo As String, StoreCode As String) As String
    Private Function GenerateASNVoucher(VoucherSID) As Boolean
        Try

            'Dim VoucherSID As String = ""
            Dim strEmpSID As String = ""

            Dim strQuery As String

            Dim dt As DataTable

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            'strQuery = "SELECT V.SID FROM RPS.VOUCHER V INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID " _
            '            & "WHERE STATUS=3 and HELD=0 And VOU_CLASS=2 And PO_NO='" & PONo & "' AND ASN_NO='" & ASN_No _
            '            & "' AND PKG_NO='" & PkgNo & "' AND ST.STORE_CODE='" & StoreCode & "'"

            strQuery = "SELECT V.SID,V.PO_NO,SL.SLIP_NO FROM RPS.VOUCHER V INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID LEFT OUTER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID " _
                        & "WHERE V.STATUS=3 and V.HELD=0 And V.VOU_CLASS=2 and V.SID='" & VoucherSID & "'"

            dt = mclsOra.GetDataSet(strQuery).Tables(0)


            ''if no PO found then check for transfer slip 
            'If dt.Rows.Count = 0 Then
            '    dt = Nothing

            '    strQuery = "SELECT V.SID FROM RPS.VOUCHER V INNER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID " _
            '            & "WHERE V.STATUS=3 and V.HELD=0 And V.VOU_CLASS=2 And SL.SLIP_NO='" & PONo & "' AND V.ASN_NO='" & ASN_No _
            '            & "' AND V.PKG_NO='" & PkgNo & "' AND ST.STORE_CODE='" & StoreCode & "'"
            '    dt = mclsOra.GetDataSet(strQuery).Tables(0)

            '    WriteToFile("Generating voucher for Slip No.: " & PONo & ", ASN No.: " & ASN_No & ", Box No.: " & PkgNo & ", Store Code: " & StoreCode)
            'Else
            '    WriteToFile("Generating voucher for PO No.: " & PONo & ", ASN No.: " & ASN_No & ", Box No.: " & PkgNo & ", Store Code: " & StoreCode)
            'End If


            If dt.Rows.Count <> 0 Then

                If IsDBNull(dt.Rows(0).Item("PO_NO")) Then
                    WriteToFile("Generating voucher for Slip No.: " & dt.Rows(0).Item("SLIP_NO") & ", Voucher SID: " & VoucherSID)
                Else
                    WriteToFile("Generating voucher for PO No.: " & dt.Rows(0).Item("PO_NO") & ", Voucher SID: " & VoucherSID)
                End If

                'VoucherSID = dt.Rows(0).Item(0)

                dt = mclsOra.GetDataSet("Select SID FROM RPS.EMPLOYEE WHERE upper(USER_NAME)='PRISM_CUSTOM'").Tables(0)
                If dt.Rows.Count <> 0 Then
                    strEmpSID = dt.Rows(0).Item(0)

                    mclsOra.CloseDB()

                    Dim mclsAPI As New clsPrismAPI
                    If authSession = "" Then
                        If Not mclsAPI.IsAPI_LoginSuccessfull Then
                            WriteToFile("API Login failed.")
                            Exit Function
                        End If
                    End If


                    'If mclsAPI.IsAPI_LoginSuccessfull Then

                    If mclsAPI.IsGenerateVoucher_Successfull(strEmpSID, VoucherSID) Then

                            'Return VoucherSID
                            Return True

                        End If

                    'End If

                Else
                    WriteToFile("PRISM_CUSTOM user not found.")
                End If

            Else

                'VoucherSID = GetPendingVoucher(PONo, ASN_No, PkgNo, StoreCode)
                'Return VoucherSID

                If GetPendingVoucher(VoucherSID) Then
                    Return True
                Else
                    WriteToFile("No record found for Voucher SID: " & VoucherSID)
                End If

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Private Function GetPendingVoucher(PONo As String, ASN_No As String, PkgNo As String, StoreCode As String) As String
    Private Function GetPendingVoucher(VoucherSID) As Boolean
        Try
            'Dim VoucherSID As String = ""
            Dim strQuery As String
            Dim dt As DataTable

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            'strQuery = "SELECT V.SID FROM RPS.VOUCHER V INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID " _
            '            & "WHERE (STATUS<>4 or HELD=1) AND VOU_CLASS<>2 And PO_NO='" & PONo & "' AND ASN_NO='" & ASN_No _
            '            & "' AND PKG_NO='" & PkgNo & "' AND ST.STORE_CODE='" & StoreCode & "'"

            strQuery = "SELECT V.SID FROM RPS.VOUCHER V INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID " _
                        & "WHERE (V.STATUS<>4 or V.HELD=1) AND V.VOU_CLASS<>2 AND V.SID='" & VoucherSID & "'"

            dt = mclsOra.GetDataSet(strQuery).Tables(0)

            'if no PO found then check for transfer slip
            'If dt.Rows.Count = 0 Then
            '    dt = Nothing

            '    strQuery = "SELECT V.SID FROM RPS.VOUCHER V INNER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID " _
            '            & "WHERE (V.STATUS<>4 or V.HELD=1) AND V.VOU_CLASS<>2 And SL.SLIP_NO='" & PONo & "' AND V.ASN_NO='" & ASN_No _
            '            & "' AND V.PKG_NO='" & PkgNo & "' AND ST.STORE_CODE='" & StoreCode & "'"
            '    dt = mclsOra.GetDataSet(strQuery).Tables(0)

            '    WriteToFile("Getting pending voucher for Slip No.: " & PONo & ", ASN No.: " & ASN_No & ", Box No.: " & PkgNo & ", Store Code: " & StoreCode)
            'Else
            '    WriteToFile("Getting pending voucher for PO No.: " & PONo & ", ASN No.: " & ASN_No & ", Box No.: " & PkgNo & ", Store Code: " & StoreCode)
            'End If

            WriteToFile(dt.Rows.Count & " pending voucher(s) found.")
            If dt.Rows.Count <> 0 Then
                'VoucherSID = dt.Rows(0).Item(0)
                Return True
            End If

            'Return VoucherSID

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetVoucherDetails(VoucherSID As String) As DataTable
        Try

            Dim strQuery As String
            Dim dt As DataTable

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            'strQuery = "Select V.*,NVL(S.QTY_RECEIVED,0)QTY_RECEIVED,S.UPDATED_AT FROM 
            '            (SELECT V.ROW_VERSION VOUCHER_ROW_VERSION, V.MODIFIED_DATETIME,V.POST_DATE,V.STATUS,V.HELD, V.VOU_TYPE,V.VOU_CLASS,
            '            V.PO_NO,V.PKG_NO,V.ASN_NO, V.VOU_NO,VI.SID VOU_ITEM_SID,VI.ROW_VERSION VOU_ITEM_ROW_VERSION, UPC,VI.ORIG_QTY,VI.QTY,
            '            SB.SBS_NO,ST.STORE_CODE 
            '            FROM RPS.VOUCHER V INNER JOIN RPS.VOU_ITEM VI ON V.SID=VI.VOU_SID
            '            INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
            '            INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
            '            INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=ST.SBS_SID 
            '            WHERE V.SID='" & VoucherSID & "')V
            '            LEFT OUTER JOIN 
            '            (SELECT H.*,D.* FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
            '            ON H.REPLE_ID=D.REPLE_HEADERID)S ON V.PO_NO=Substr(po_num,1,instr(po_num,'-')-1)
            '            AND V.ASN_NO=S.BL_NUM AND V.SBS_NO=S.SBS_NO AND V.STORE_CODE=S.STORE_CODE
            '            AND V.UPC=S.SKU WHERE NVL(S.QTY_RECEIVED,0)<>0"

            strQuery = "Select V.*,NVL(S.QTY_RECEIVED,0)QTY_RECEIVED,S.UPDATED_AT FROM 
                        (SELECT V.SID,V.ROW_VERSION VOUCHER_ROW_VERSION, V.MODIFIED_DATETIME,V.POST_DATE,V.STATUS,V.HELD, V.VOU_TYPE,V.VOU_CLASS,
                        V.PO_NO,V.PKG_NO,V.ASN_NO, V.VOU_NO,VI.SID VOU_ITEM_SID,VI.ROW_VERSION VOU_ITEM_ROW_VERSION,NVL(I.ALU,I.UPC) SKU,I.UPC UPC,VI.ORIG_QTY,VI.QTY,
                        SB.SBS_NO,ST.STORE_CODE 
                        FROM RPS.VOUCHER V INNER JOIN RPS.VOU_ITEM VI ON V.SID=VI.VOU_SID
                        INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
                        INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
                        INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=ST.SBS_SID 
                        WHERE V.SID='" & VoucherSID & "')V
                        LEFT OUTER JOIN 
                        (SELECT H.*,D.* FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
                        ON H.REPLE_ID=D.REPLE_HEADERID)S ON V.SID=S.VOU_SID                        
                        AND V.SKU=S.SKU WHERE NVL(S.QTY_RECEIVED,0)<>0"

            dt = mclsOra.GetDataSet(strQuery).Tables(0)

            ''if no PO found then check for transfer slip
            'If dt.Rows.Count = 0 Then
            '    dt = Nothing

            '    strQuery = "SELECT V.*,NVL(S.QTY_RECEIVED,0)QTY_RECEIVED,S.UPDATED_AT FROM 
            '            (SELECT V.SID,V.ROW_VERSION VOUCHER_ROW_VERSION, V.MODIFIED_DATETIME,V.POST_DATE,V.STATUS,V.HELD, V.VOU_TYPE,V.VOU_CLASS,
            '            V.PO_NO,V.PKG_NO,V.ASN_NO,SL.SLIP_NO,V.VOU_NO,VI.SID VOU_ITEM_SID,VI.ROW_VERSION VOU_ITEM_ROW_VERSION, UPC,VI.ORIG_QTY,VI.QTY,
            '            SB.SBS_NO,ST.STORE_CODE 
            '            FROM RPS.VOUCHER V INNER JOIN RPS.VOU_ITEM VI ON V.SID=VI.VOU_SID
            '            LEFT OUTER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID
            '            INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
            '            INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
            '            INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=ST.SBS_SID 
            '            WHERE V.SID='" & VoucherSID & "')V
            '            LEFT OUTER JOIN 
            '            (SELECT H.*,D.* FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
            '            ON H.REPLE_ID=D.REPLE_HEADERID)S ON V.SLIP_NO=Substr(po_num,1,instr(po_num,'-')-1)
            '            AND V.ASN_NO=S.BL_NUM AND V.SBS_NO=S.SBS_NO AND V.STORE_CODE=S.STORE_CODE
            '            AND V.UPC=S.SKU WHERE NVL(S.QTY_RECEIVED,0)<>0"

            '    dt = mclsOra.GetDataSet(strQuery).Tables(0)

            'End If

            If dt.Rows.Count <> 0 Then
                Return dt
            Else
                Return Nothing
            End If
            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
            'Return Nothing
        End Try
    End Function

    Private Sub ReceiveVoucherItem(strVoucherSID As String)
        Try
            Dim dtVoucherItem As DataTable
            Dim strEmpSID As String = ""
            Dim dt As DataTable

            Dim intVoucherRowVersion As Integer = 0
            Dim dtVoucherRowVersion As DataTable
            'Dim UpdatedAt As DateTime

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()
            Dim mclsAPI As New clsPrismAPI

            If authSession = "" Then
                If Not mclsAPI.IsAPI_LoginSuccessfull Then
                    WriteToFile("API Login failed.")
                    Exit Sub
                End If
            End If

            dtVoucherItem = GetVoucherDetails(strVoucherSID)

            If Not IsNothing(dtVoucherItem) Then
                If dtVoucherItem.Rows.Count <> 0 Then

                    'UpdatedAt = dtVoucherItem.Rows(0).Item("UPDATED_AT")


                    For Each dRow As DataRow In dtVoucherItem.Rows
                        WriteToFile("Receiving Sku: " & dRow.Item("SKU") & ", Qty: " & dRow.Item("QTY_RECEIVED"))
                        mclsAPI.ReceiveVoucherItem(strVoucherSID, dRow.Item("VOU_ITEM_SID"), dRow.Item("VOU_ITEM_ROW_VERSION"), dRow.Item("UPC"), dRow.Item("QTY_RECEIVED"))
                    Next

                Else

                    WriteToFile("No matching voucher item found to receive for Voucher SID: " & strVoucherSID)

                End If

                'Else
                'WriteToFile(strEmpSID & "Voucher item found but with 0 receive qty for Voucher SID: " & strVoucherSID)
            End If

            Dim dtExtraItem As DataTable
            dtExtraItem = GetExtraItems(strVoucherSID)
            If dtExtraItem.Rows.Count <> 0 Then

                For Each dRow As DataRow In dtExtraItem.Rows
                    WriteToFile("Adding extra item: " & dRow.Item("UPC") & ", Qty: " & dRow.Item("QTY_RECEIVED"))
                    mclsAPI.AddExtraItem(strVoucherSID, dRow.Item("ITEM_SID"), dRow.Item("UPC"), dRow.Item("QTY_RECEIVED"), dRow.Item("STYLENO"), dRow.Item("PRICE"))
                Next

            End If

            dtVoucherRowVersion = mclsOra.GetDataSet("SELECT ROW_VERSION FROM RPS.VOUCHER WHERE SID='" & strVoucherSID & "' AND STATUS=3").Tables(0)
            If dtVoucherRowVersion.Rows.Count > 0 Then

                dt = mclsOra.GetDataSet("SELECT SID FROM RPS.EMPLOYEE WHERE upper(USER_NAME)='PRISM_CUSTOM'").Tables(0)
                If dt.Rows.Count <> 0 Then strEmpSID = dt.Rows(0).Item(0)

                intVoucherRowVersion = dtVoucherRowVersion.Rows(0).Item(0)

                'mclsAPI.ApproveVoucher(strVoucherSID, strEmpSID, intVoucherRowVersion, UpdatedAt)
                mclsAPI.ApproveVoucher(strVoucherSID, strEmpSID, intVoucherRowVersion, Now)

                mclsOra.ExecuteNonQuery("UPDATE XXASH_SALASA_REPLE_HEADER H SET H.RETAILPRO_RECEIVED='Y',H.MODIFIED_DATE=SYSDATE WHERE VOU_SID='" & strVoucherSID & "' AND H.RETAILPRO_RECEIVED<>'Y'")

            End If


            WriteToFile("Voucher receiving process completed.")


            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ReceiveExtraItemManual(strVoucherSID As String)
        Dim dtExtraItem As DataTable
        dtExtraItem = GetExtraItems(strVoucherSID)
        If dtExtraItem.Rows.Count <> 0 Then

            Dim mclsAPI As New clsPrismAPI
            If mclsAPI.IsAPI_LoginSuccessfull = False Then
                WriteToFile("API Login failed.")
                Exit Sub
            End If


            For Each dRow As DataRow In dtExtraItem.Rows
                WriteToFile("Adding extra item: " & dRow.Item("SKU") & ", Qty: " & dRow.Item("QTY_RECEIVED"))
                mclsAPI.AddExtraItem(strVoucherSID, dRow.Item("ITEM_SID"), dRow.Item("SKU"), dRow.Item("QTY_RECEIVED"), dRow.Item("STYLENO"), dRow.Item("PRICE"))
            Next

            WriteToFile("Extra item added succesfully.")

        End If
    End Sub

    Private Function GetExtraItems(VoucherSID) As DataTable
        Try
            Dim dtExtraItem As DataTable
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            Dim strQuery As String
            mclsOra.OpenDB()
            strQuery = "SELECT H.*,D.*,IM.UPC,IM.SID ITEM_SID,IM.DESCRIPTION1 STYLENO,P.PRICE  FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
                        ON H.REPLE_ID=D.REPLE_HEADERID
                        LEFT OUTER JOIN RPS.INVN_SBS_ITEM IM ON NVL(IM.ALU,IM.UPC)=D.SKU
                        LEFT OUTER JOIN
                        (SELECT NVL(ALU,UPC)SKU, PL.PRICE FROM RPS.INVN_SBS_ITEM I 
                        INNER JOIN RPS.INVN_SBS_PRICE PL ON PL.INVN_SBS_ITEM_SID=I.SID
                        INNER JOIN RPS.PRICE_LEVEL p ON P.SID=PL.PRICE_LVL_SID
                        INNER JOIN RPS.SUBSIDIARY s ON S.SID=I.SBS_SID
                        AND P.PRICE_LVL=1)P ON P.SKU=D.SKU
                        WHERE EXISTS(
                        SELECT V.* FROM RPS.VOUCHER V INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=V.SBS_SID 
                        INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
                        INNER JOIN RPS.VOU_ITEM VI ON VI.VOU_SID=V.SID
                        INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
                        WHERE V.SID='" & VoucherSID & "'
                        AND H.VOU_SID=V.SID
                        AND H.SBS_NO=SB.SBS_NO AND H.STORE_CODE=ST.STORE_CODE)
                        AND NOT EXISTS
                        (SELECT VI.* FROM RPS.VOUCHER V INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=V.SBS_SID 
                        INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
                        INNER JOIN RPS.VOU_ITEM VI ON VI.VOU_SID=V.SID
                        INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
                        WHERE V.SID='" & VoucherSID & "'
                        AND H.VOU_SID=V.SID
                        AND H.SBS_NO=SB.SBS_NO AND H.STORE_CODE=ST.STORE_CODE
                        AND D.SKU=NVL(I.ALU,I.UPC))"
            dtExtraItem = mclsOra.GetDataSet(strQuery).Tables(0)

            ''if no PO found then check for transfer slip
            'If dtExtraItem.Rows.Count = 0 Then
            '    dtExtraItem = Nothing
            '    strQuery = "SELECT H.*,D.*,IM.SID ITEM_SID,IM.DESCRIPTION1 STYLENO,P.PRICE  FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
            '                ON H.REPLE_ID=D.REPLE_HEADERID
            '                LEFT OUTER JOIN RPS.INVN_SBS_ITEM IM ON IM.UPC=D.SKU
            '                LEFT OUTER JOIN
            '                (SELECT UPC, PL.PRICE FROM RPS.INVN_SBS_ITEM I 
            '                INNER JOIN RPS.INVN_SBS_PRICE PL ON PL.INVN_SBS_ITEM_SID=I.SID
            '                INNER JOIN RPS.PRICE_LEVEL p ON P.SID=PL.PRICE_LVL_SID
            '                INNER JOIN RPS.SUBSIDIARY s ON S.SID=I.SBS_SID
            '                AND P.PRICE_LVL=1)P ON P.UPC=D.SKU
            '                WHERE EXISTS(
            '                SELECT V.* FROM RPS.VOUCHER V INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=V.SBS_SID 
            '                INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
            '                INNER JOIN RPS.VOU_ITEM VI ON VI.VOU_SID=V.SID
            '                LEFT OUTER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID
            '                INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
            '                WHERE V.SID='" & VoucherSID & "'
            '                AND H.VOU_SID=V.SID
            '                AND H.SBS_NO=SB.SBS_NO AND H.STORE_CODE=ST.STORE_CODE)
            '                AND NOT EXISTS
            '                (SELECT VI.* FROM RPS.VOUCHER V INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=V.SBS_SID 
            '                LEFT OUTER JOIN RPS.SLIP SL ON SL.VOU_SID=V.SID
            '                INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
            '                INNER JOIN RPS.VOU_ITEM VI ON VI.VOU_SID=V.SID
            '                INNER JOIN RPS.INVN_SBS_ITEM I ON I.SID=VI.ITEM_SID 
            '                WHERE V.SID='" & VoucherSID & "'
            '                AND H.PO_NUM=SL.SLIP_NO||'-'||V.ASN_NO AND H.SBS_NO=SB.SBS_NO AND H.STORE_CODE=ST.STORE_CODE
            '                AND D.SKU=I.UPC)"
            '    dtExtraItem = mclsOra.GetDataSet(strQuery).Tables(0)

            'End If

            mclsOra.CloseDB()

            Return dtExtraItem

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Create_LOGFile()

            DailyVoucherReceiving()

        Catch ex As Exception
            WriteToFile(ex.Message)
            SendEmail_AppError(ex.Message)
        Finally
            End
        End Try
    End Sub

    'Private Sub ReceiveNewASNVoucher(PONo As String, ASN_No As String, PkgNo As String, StoreCode As String)
    'Private Sub ReceiveNewASNVoucher(VoucherSID)
    '    Try
    '        'Dim strVoucherSID As String = ""
    '        'strVoucherSID = GenerateASNVoucher(PONo, ASN_No, PkgNo, StoreCode)
    '        'GenerateASNVoucher(VoucherSID)

    '        'If strVoucherSID <> "" Then
    '        '    ReceiveVoucherItem(strVoucherSID)
    '        'Else
    '        '    WriteToFile("Voucher not found.")
    '        'End If

    '        If GenerateASNVoucher(VoucherSID) Then
    '            ReceiveVoucherItem(VoucherSID)
    '        End If

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Function ChangeCurrentStoreSID() As Boolean
        Try
            Dim strSBSID As String = "", strStoreSID As String = ""
            Dim dt As DataTable

            WriteToFile("Changing store code.")
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            WriteToFile("Getting SID and SBS_SID of store code " & strDBStoreCode & "...")
            mclsOra.OpenDB()
            dt = mclsOra.GetDataSet("select sid,sbs_sid from rps.store where store_code='" & strDBStoreCode & "'").Tables(0)
            If dt.Rows.Count <> 0 Then
                strSBSID = dt.Rows(0).Item("sbs_sid")
                strStoreSID = dt.Rows(0).Item("sid")

                Dim mclsAPI As New clsPrismAPI
                If authSession = "" Then
                    If Not mclsAPI.IsAPI_LoginSuccessfull Then
                        WriteToFile("API Login failed.")
                        Exit Function
                    End If
                End If

                Return mclsAPI.ChangeCurrent_StoreSID(strSBSID, strStoreSID)
            Else
                WriteToFile("Unable to change store.  Store code not found.")
            End If
            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub DailyVoucherReceiving()
        Try

            Dim strQuery As String, dt As DataTable, strVouPONo As String
            Dim VoucherSID As String = ""

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            Dim mclsAPI As New clsPrismAPI
            If authSession = "" Then
                If Not mclsAPI.IsAPI_LoginSuccessfull Then

                    WriteToFile("API Login failed.")
                    Exit Sub

                Else

                    If currStoreCode <> strDBStoreCode Then
                        If UCase(strUseDBStoreCode) = "Y" Then
                            If ChangeCurrentStoreSID() = False Then Exit Sub
                        End If
                    End If

                End If
            End If

            WriteToFile("Getting replenishment with ""ready for your review"" status and not yet received in Retail PRO...")
            strQuery = "SELECT * FROM XXASH_SALASA_REPLE_HEADER WHERE NVL(RETAILPRO_RECEIVED,'N')='N' AND TRUNC(MODIFIED_DATE)>=TRUNC(SYSDATE) And STATUS ='ready for your review'
                        AND SBS_NO='" & currSBSNo & "' AND STORE_CODE='" & currStoreCode & "' ORDER BY REPLE_ID"
            dt = mclsOra.GetDataSet(strQuery).Tables(0)
            WriteToFile(dt.Rows.Count & " record(s) found.")

            If dt.Rows.Count <> 0 Then

                'mclsOra.ExecuteNonQuery("BEGIN XXASH_DROP_TMPTABLE('XXASHTMPREPLEID'); END;")
                'mclsOra.ExecuteNonQuery("commit")
                'mclsOra.ExecuteNonQuery("CREATE TABLE XXASHTMPREPLEID(REPLE_ID INTEGER)")
                'mclsOra.ExecuteNonQuery("commit")

                For Each dRow As DataRow In dt.Rows
                    'strVouPONo = Mid(dRow.Item("PO_NUM"), 1, dRow.Item("PO_NUM").ToString.LastIndexOf("-"))

                    strVouPONo = Mid(dRow.Item("BL_NUM"), 1, dRow.Item("BL_NUM").ToString.LastIndexOf("-"))
                    VoucherSID = dRow.Item("VOU_SID").ToString

                    WriteToFile("Processing received qty for PO-ASN No.: " & dRow.Item("BL_NUM") & ", Box No.: " & dRow.Item("CONTAINER_NUM"))

                    If IsReplenishmentHaveMissingSKUs(VoucherSID) = False Then

                        'ReceiveNewASNVoucher(strVouPONo, dRow.Item("BL_NUM"), dRow.Item("CONTAINER_NUM"), dRow.Item("STORE_CODE"))
                        'ReceiveNewASNVoucher(VoucherSID)

                        If GenerateASNVoucher(VoucherSID) Then
                            ReceiveVoucherItem(VoucherSID)
                        End If

                        'mclsOra.ExecuteNonQuery("INSERT INTO XXASHTMPREPLEID (REPLE_ID) VALUES(" & dRow.Item("REPLE_ID") & ")")
                        'WriteToFile("Received qty was successfully processed for PO No.: " & dRow.Item("PO_NUM") & ", Box No.: " & dRow.Item("CONTAINER_NUM"))

                    Else

                        WriteToFile("Received qty was not processed due to missing SKU(s) for PO No.: " & dRow.Item("PO_NUM") & ", Box No.: " & dRow.Item("CONTAINER_NUM"))

                    End If

                Next

                'UPDATE po note with "C"
                For Each dRow As DataRow In dt.Rows
                    If UCase(dRow.Item("COMMENTS")) = "SHIPMENT" Then
                        strVouPONo = Mid(dRow.Item("BL_NUM"), 1, dRow.Item("BL_NUM").ToString.LastIndexOf("-"))
                        UpdatePONote(strVouPONo)
                    End If
                Next

                SendSuccessfulVoucherReceivedtoEmail()
                SendMissingSKUtoEmail()

            Else
                WriteToFile("No record found.")
            End If

            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub UpdatePONote(strPONo As String)
        Try
            Dim strQuery As String
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            Dim dt As DataTable
            Dim strPOSID As String, intPORowVersion As Integer

            mclsOra.OpenDB()
            strQuery = "SELECT COUNT(V.PKG_NO) VOU_COUNT FROM rps.VOUCHER V INNER JOIN RPS.STORE ST ON ST.SID=V.STORE_SID 
                        INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=ST.SBS_SID 
                        WHERE SBS_NO='" & currSBSNo & "' AND ST.STORE_CODE='" & currStoreCode & "' AND V.STATUS<>4 AND V.VOU_CLASS=2
                        AND PO_NO='" & strPONo & "'"
            dt = mclsOra.GetDataSet(strQuery).Tables(0)
            If dt.Rows(0).Item("VOU_COUNT") = 0 Then
                dt = Nothing

                strQuery = "SELECT P.SID,P.ROW_VERSION FROM RPS.PO P INNER JOIN RPS.STORE ST ON ST.SID=P.STORE_SID 
                            INNER JOIN RPS.SUBSIDIARY SB ON SB.SID=ST.SBS_SID 
                            WHERE PO_NO='" & strPONo & "' AND SBS_NO='" & currSBSNo & "' AND ST.STORE_CODE='" & currStoreCode & "'"
                dt = mclsOra.GetDataSet(strQuery).Tables(0)
                If dt.Rows.Count <> 0 Then
                    strPOSID = dt.Rows(0).Item("SID")
                    intPORowVersion = dt.Rows(0).Item("ROW_VERSION")

                    Dim mclsAPI As New clsPrismAPI

                    If authSession = "" Then
                        If Not mclsAPI.IsAPI_LoginSuccessfull Then
                            WriteToFile("API Login failed.")
                            Exit Sub
                        End If
                    End If

                    WriteToFile("Updating PO Note for PO No.: " & strPONo)
                    mclsAPI.UpdatePONote(strPOSID, intPORowVersion)

                End If

            End If
            mclsOra.CloseDB()
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SendSuccessfulVoucherReceivedtoEmail()
        Try
            Dim strQuery As String
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            Dim dtSummary As DataTable ', dtDetail As DataTable
            strQuery = "Select SBS_NAME BRAND,H.STORE_CODE,ST.STORE_NAME,H.REPLE_ID,
                                CASE WHEN INSTR(H.BL_NUM,'-')=0 THEN H.BL_NUM ELSE SUBSTR(H.BL_NUM,1,INSTR(H.BL_NUM,'-')-1) END PO_NO,
                                SUBSTR(H.BL_NUM,INSTR(H.BL_NUM,'-')+1,5)ASN_NO,H.TRACKING_NUM INVOICE_NO,H.CONTAINER_NUM BOX_NO,VOU_NO,V.POST_DATE,
                                SUM(D.QTY_REQUEST)SHIP_QTY,SUM(D.QTY_RECEIVED)RCV_QTY
                                FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D ON H.REPLE_ID=D.REPLE_HEADERID                                                            
                                INNER JOIN RPS.SUBSIDIARY SB ON SB.SBS_NO=H.SBS_NO 
                                INNER JOIN RPS.STORE ST ON ST.STORE_CODE=H.STORE_CODE 
                                LEFT OUTER JOIN RPS.VOUCHER V ON V.SID=H.VOU_SID
                                where NVL(RETAILPRO_RECEIVED,'N')='Y' AND NVL(VOU_RCV_EMAIL_SENT,'N')='N'
                                GROUP BY H.REPLE_ID,CASE WHEN INSTR(H.BL_NUM,'-')=0 THEN H.BL_NUM ELSE SUBSTR(H.BL_NUM,1,INSTR(H.BL_NUM,'-')-1) END,
                                SUBSTR(H.BL_NUM,INSTR(H.BL_NUM,'-')+1,5),H.TRACKING_NUM,H.CONTAINER_NUM,UPDATED_AT,
                                SBS_NAME,H.STORE_CODE,ST.STORE_NAME,VOU_NO,V.POST_DATE
                                ORDER BY SBS_NAME,ST.STORE_NAME,H.REPLE_ID"
            dtSummary = mclsOra.GetDataSet(strQuery).Tables(0)

            If dtSummary.Rows.Count <> 0 Then

                'strQuery = "SELECT SBS_NAME BRAND,H.STORE_CODE,ST.STORE_NAME,H.REPLE_ID,CASE WHEN INSTR(H.BL_NUM,'-')=0 THEN H.BL_NUM ELSE SUBSTR(H.BL_NUM,1,INSTR(H.BL_NUM,'-')-1) END PO_NO,
                '                SUBSTR(H.BL_NUM,INSTR(H.BL_NUM,'-')+1,5)ASN_NO,H.TRACKING_NUM INVOICE_NO,H.CONTAINER_NUM BOX_NO,                                    
                '                    VOU_NO,TO_CHAR(V.POST_DATE,'YYYY-MM-DD HH:MI:SS AM')POST_DATE,D.SKU,
                '                    SUM(D.QTY_REQUEST)SHIP_QTY,SUM(D.QTY_RECEIVED)RCV_QTY
                '                    FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D ON H.REPLE_ID=D.REPLE_HEADERID                                    
                '                    INNER JOIN RPS.SUBSIDIARY SB ON SB.SBS_NO=H.SBS_NO 
                '                    INNER JOIN RPS.STORE ST ON ST.STORE_CODE=H.STORE_CODE
                '                    LEFT OUTER JOIN RPS.VOUCHER V ON V.SID=H.VOU_SID
                '                    where NVL(RETAILPRO_RECEIVED,'N')='Y' AND NVL(VOU_RCV_EMAIL_SENT,'N')='N'
                '                    GROUP BY H.REPLE_ID,CASE WHEN INSTR(H.BL_NUM,'-')=0 THEN H.BL_NUM ELSE SUBSTR(H.BL_NUM,1,INSTR(H.BL_NUM,'-')-1) END,
                '                    SUBSTR(H.BL_NUM,INSTR(H.BL_NUM,'-')+1,5),H.TRACKING_NUM,H.CONTAINER_NUM,VOU_NO,TO_CHAR(V.POST_DATE,'YYYY-MM-DD HH:MI:SS AM'),
                '                    SBS_NAME,H.STORE_CODE,ST.STORE_NAME,D.SKU
                '                    ORDER BY SBS_NAME,ST.STORE_NAME,H.REPLE_ID"
                'dtDetail = mclsOra.GetDataSet(strQuery).Tables(0)

                Dim strfileAttachment As String = ""
                'If dtDetail.Rows.Count <> 0 Then

                '    If Dir(System.Windows.Forms.Application.StartupPath & "\RECEIVED_REPLENISHMENT", vbDirectory) = vbNullString Then
                '        MkDir(System.Windows.Forms.Application.StartupPath & "\RECEIVED_REPLENISHMENT")
                '    End If

                '    strfileAttachment = System.Windows.Forms.Application.StartupPath & "\RECEIVED_REPLENISHMENT\ReceivedReplenishment_" & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
                '    ExportToExcel_EPPlus(dtDetail, strfileAttachment)

                'End If

                Dim tblRepleRecv As StringBuilder

                tblRepleRecv = CreateHTMLTable(dtSummary)


                WriteToFile("Sending successful voucher receiving email to " & strEmailRecipient & " ...")
                SendEmail(strEmailReceivedVoucher, "Successful voucher receiving from Salasa Replenishment", strfileAttachment, tblRepleRecv, strEmailReceivedVoucherCC)
                WriteToFile("Email sent.")

                mclsOra.ExecuteNonQuery("UPDATE XXASH_SALASA_REPLE_HEADER H SET H.VOU_RCV_EMAIL_SENT='Y' WHERE NVL(H.RETAILPRO_RECEIVED,'N')='Y' AND NVL(H.VOU_RCV_EMAIL_SENT,'N')='N'")

            End If

            mclsOra.CloseDB()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function IsReplenishmentHaveMissingSKUs(VoucherSID As String) As Boolean
        Try
            Dim strQuery As String


            Dim mclsSQL As New clsSQLDB
            mclsSQL.OpenDB()
            mclsSQL.ExecuteNonQuery("IF OBJECT_ID('_tmpMissingSKUs', 'U') IS NOT NULL DROP TABLE _tmpMissingSKUs")
            strQuery = "CREATE TABLE [_tmpMissingSKUs]([REPLENISHMENT_ID] [int] NULL,[PO_NUM] [varchar](30) NULL,
	                        [CONTAINER_NUM] [varchar](100) NULL,
	                        [SKU] [varchar](30) NULL,[RCV_QTY] [INT] NULL
                        ) ON [PRIMARY]"
            mclsSQL.ExecuteNonQuery(strQuery)

            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()
            Dim dt As DataTable

            WriteToFile("Checking missing SKU(s)...")
            strQuery = "SELECT DISTINCT H.REPLE_ID, H.PO_NUM,H.CONTAINER_NUM, D.SKU,D.QTY_RECEIVED FROM XXASH_SALASA_REPLE_HEADER H INNER JOIN XXASH_SALASA_REPLE_DETAIL D
                        ON H.REPLE_ID=D.REPLE_HEADERID
                        LEFT OUTER JOIN RPS.INVN_SBS_ITEM I ON D.SKU=NVL(I.ALU,I.UPC)
                        WHERE H.VOU_SID='" & VoucherSID & "' AND I.SID IS NULL"
            dt = mclsOra.GetDataSet(strQuery).Tables(0)
            WriteToFile(dt.Rows.Count & " missing SKU(s) found.")

            mclsOra.CloseDB()

            If dt.Rows.Count <> 0 Then

                For Each dRow As DataRow In dt.Rows

                    strQuery = "insert into _tmpMissingSKUs (REPLENISHMENT_ID,PO_NUM,CONTAINER_NUM,SKU) values(" & dRow.Item("REPLE_ID") & ",'" & dRow.Item("PO_NUM") & "','" _
                    & dRow.Item("CONTAINER_NUM") & "','" & dRow.Item("SKU") & "'," & dRow.Item("QTY_RECEIVED") & ")"
                    mclsSQL.ExecuteNonQuery(strQuery)

                    WriteToFile("SKU " & dRow.Item("SKU") & " from Replenishment ID: " & dRow.Item("REPLE_ID") & " PO.No-ASN.No: " & dRow.Item("PO_NUM") & ", Box No: " & dRow.Item("CONTAINER_NUM") & " does not exists in Retail Pro.")
                Next

                mclsSQL.CloseDB()

                Return True

            Else
                mclsSQL.CloseDB()
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub SendMissingSKUtoEmail()
        Try
            Dim dtMissingSKU As DataTable
            Dim mclsSQL As New clsSQLDB
            mclsSQL.OpenDB()

            dtMissingSKU = mclsSQL.GetDataSet("SELECT * FROM _tmpMissingSKUs").Tables(0)

            If dtMissingSKU.Rows.Count <> 0 Then

                If Dir(System.Windows.Forms.Application.StartupPath & "\MISSING_SKU", vbDirectory) = vbNullString Then
                    MkDir(System.Windows.Forms.Application.StartupPath & "\MISSING_SKU")
                End If

                Dim strfileAttachment As String = System.Windows.Forms.Application.StartupPath & "\MISSING_SKU\MissingSKU_" & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
                ExportToExcel_EPPlus(dtMissingSKU, strfileAttachment)

                WriteToFile("Sending missing SKUs email to: " & strEmailReceivedVoucher & ", CC: " & strEmailRecipient & " ...")
                SendEmail(strEmailReceivedVoucher, "Undefined SKU from Salasa Replenishment", strfileAttachment, Nothing, strEmailRecipient)
                WriteToFile("Email sent.")
            End If

            mclsSQL.ExecuteNonQuery("If OBJECT_ID('_tmpMissingSKUs', 'U') IS NOT NULL DROP TABLE _tmpMissingSKUs")

            mclsSQL.CloseDB()

        Catch ex As Exception
            If ex.Message.Contains("Invalid object name") Then
            Else
                Throw ex
            End If
        End Try
    End Sub

    'Private Sub ReceivePendingVoucher(PONo As String, ASN_No As String, PkgNo As String, StoreCode As String)
    Private Sub ReceivePendingVoucher(strVoucherSID As String)
        Try
            'Dim strVoucherSID As String = ""
            'strVoucherSID = GetPendingVoucher(PONo, ASN_No, PkgNo, StoreCode)

            'If strVoucherSID <> "" Then ReceiveVoucherItem(strVoucherSID)

            ReceiveVoucherItem(strVoucherSID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Private Sub ApprovePendingVoucher(PONo As String, ASN_No As String, PkgNo As String, StoreCode As String)
    Private Sub ApprovePendingVoucher(strVoucherSID As String)
        Try
            Dim dt As DataTable
            Dim dtVoucherRowVersion As DataTable
            Dim strEmpSID As String = ""
            'Dim strVoucherSID As String = ""

            Dim intVoucherRowVersion As Integer = 0
            Dim mclsAPI As New clsPrismAPI
            Dim mclsOra As New clsOracleDB(strRPDataSource, strRPUserID, strRPPswrd)
            mclsOra.OpenDB()

            'strVoucherSID = GetPendingVoucher(PONo, ASN_No, PkgNo, StoreCode)
            If strVoucherSID <> "" Then

                dt = mclsOra.GetDataSet("SELECT SID FROM RPS.EMPLOYEE WHERE upper(USER_NAME)='PRISM_CUSTOM'").Tables(0)
                If dt.Rows.Count <> 0 Then strEmpSID = dt.Rows(0).Item(0)

                If strEmpSID <> "" Then
                    If authSession = "" Then
                        If Not mclsAPI.IsAPI_LoginSuccessfull Then
                            WriteToFile("API Login failed.")
                            Exit Sub
                        End If
                    End If

                    dtVoucherRowVersion = mclsOra.GetDataSet("SELECT ROW_VERSION FROM RPS.VOUCHER WHERE SID='" & strVoucherSID & "' AND STATUS=4").Tables(0)
                    If dtVoucherRowVersion.Rows.Count > 0 Then
                        intVoucherRowVersion = dtVoucherRowVersion.Rows(0).Item(0)

                        mclsAPI.ApproveVoucher(strVoucherSID, strEmpSID, intVoucherRowVersion, Now)

                    End If

                End If
            End If

            mclsOra.CloseDB()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


End Class
