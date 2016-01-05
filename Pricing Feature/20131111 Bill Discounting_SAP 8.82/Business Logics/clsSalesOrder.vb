Public Class clsSalesOrder
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oBP As SAPbobsCOM.BusinessPartners
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Validate Customer Ref no "
    Private Function ValiateCustomerRerNo(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strdocnum, strrefno, strCardCode, strSQL As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strrefno = oApplication.Utilities.GetEditText(aForm, "14")
        strdocnum = oApplication.Utilities.GetEditText(aForm, "8")
        strCardCode = oApplication.Utilities.GetEditText(aForm, "4")
        If strrefno <> "" And strCardCode <> "" And aForm.TypeEx = frm_SalesOrder Then
            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                strSQL = "Select  * from ORDR where cardcode='" & strCardCode & "' and  NumAtCard='" & strrefno.Trim() & "'"
            ElseIf (aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                strSQL = "Select  * from ORDR where cardcode='" & strCardCode & "' and  NumAtCard='" & strrefno.Trim() & "' and docnum <>" & strdocnum
            End If
            oTemp.DoQuery(strSQL)
            If oTemp.RecordCount > 0 Then
                ' oApplication.Utilities.Message("Customer Reference number already exists: " & strrefno, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                If oApplication.SBO_Application.MessageBox("LPO Number already exists, Do you want to continue? ", , "Yes", "No") = 2 Then
                    Return False
                Else
                    Return True

                End If

                'Return False
            End If
        End If
        Return True
    End Function
#End Region

#Region "Populate Free Item Discount"
    Private Function populatefreeitemdiscount(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            aform.Freeze(True)
            Dim strSeries, stritem As String
            oCombobox = aform.Items.Item("88").Specific
            strSeries = oCombobox.Selected.Description
            oMatrix = aform.Items.Item("38").Specific
            If aform.TypeEx <> frm_SalesOrder Then
                aform.Freeze(False)
                Return True
            End If

            Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
            strLocalCurrency = oApplication.Utilities.GetLocalCurrency()
            strsystemcurrency = oApplication.Utilities.GetSystemCurrency()
            strCardCode = oApplication.Utilities.getBPCurrency(oApplication.Utilities.GetEditText(aform, "4"))
            strdoctotal = oApplication.Utilities.GetEditText(aform, "29")
            strdoctotal = strdoctotal.Replace(strLocalCurrency, "")
            strdoctotal = strdoctotal.Replace(strCardCode, "")
            strdoctotal = strdoctotal.Replace(strsystemcurrency, "")
            Dim dbldoctotal As Double
            Try
                dbldoctotal = CDbl(strdoctotal)
            Catch ex As Exception
                dbldoctotal = 0
            End Try

            If dbldoctotal <> 0 And (strSeries = "GFC" Or strSeries = "SR") Then
                oApplication.Utilities.Message("For GFC / SR Series Sales order the document total should be Zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Freeze(False)
                Return False
            End If

            If ValiateCustomerRerNo(aform) = False Then
                aform.Freeze(False)
                Return False
            End If
            Dim dblRowGPrice, dblGPRice As Double
            dblGPRice = 0
            dblRowGPrice = 0
            If (strSeries = "GFC" Or strSeries = "SR") Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    stritem = oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & stritem & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        stritem = oTest.Fields.Item("ItemCode").Value
                    Else
                        stritem = ""
                    End If
                    If stritem <> "" Then
                        oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value <> "F" Then
                            oApplication.Utilities.Message("For GFC Series the Item Type should be Free: Line no : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            aform.Freeze(False)
                            Return False
                        End If
                    End If
                    dblRowGPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_GPrice", intRow))
                    dblGPRice = dblGPRice + dblRowGPrice
                Next
            Else
                For intRow As Integer = 1 To oMatrix.RowCount
                    stritem = oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                    If stritem <> "" Then
                        oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value = "F" Then
                            dblRowGPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_GPrice", intRow))
                            dblGPRice = dblGPRice + dblRowGPrice
                        End If
                    End If

                Next
            End If
            aform.Freeze(False)
            Dim intFormType As Integer
            intFormType = aform.Type
            intFormType = intFormType * -1
            Dim oUDFFOrm As SAPbouiCOM.Form
            Try
                oUDFFOrm = oApplication.SBO_Application.Forms.GetForm(intFormType.ToString, aform.TypeCount)
                oApplication.Utilities.SetEditText(oUDFFOrm, "U_Z_GPrice", dblGPRice)
            Catch ex As Exception
            End Try
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False

        End Try
    End Function


    Private Sub calculateGrossPr(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strSeries, stritem As String
            oCombobox = aform.Items.Item("88").Specific
            strSeries = oCombobox.Selected.Description
            oMatrix = aform.Items.Item("38").Specific

            Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
            strLocalCurrency = oApplication.Utilities.GetLocalCurrency()
            strsystemcurrency = oApplication.Utilities.GetSystemCurrency()
            strCardCode = oApplication.Utilities.getBPCurrency(oApplication.Utilities.GetEditText(aform, "4"))
            strdoctotal = oApplication.Utilities.GetEditText(aform, "29")

            strdoctotal = strdoctotal.Replace(strLocalCurrency, "")
            strdoctotal = strdoctotal.Replace(strCardCode, "")
            strdoctotal = strdoctotal.Replace(strsystemcurrency, "")

            Dim dblRowGPrice, dblGPRice As Double
            dblGPRice = 0
            dblRowGPrice = 0
            If (strSeries = "GFC" Or strSeries = "SR") Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    stritem = oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & stritem & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        stritem = oTest.Fields.Item("ItemCode").Value
                    Else
                        stritem = ""
                    End If
                    If stritem <> "" Then
                        oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value <> "F" Then
                            dblRowGPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_GPrice", intRow))
                            dblGPRice = dblGPRice + dblRowGPrice
                        End If
                    End If
                Next
            Else
                For intRow As Integer = 1 To oMatrix.RowCount
                    stritem = oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                    If stritem <> "" Then
                        oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value = "F" Then
                            dblRowGPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_GPrice", intRow))
                            dblGPRice = dblGPRice + dblRowGPrice
                        End If
                    End If
                Next
            End If
            aform.Freeze(False)
            Dim intFormType As Integer
            intFormType = aform.Type
            intFormType = intFormType * -1
            Dim oUDFFOrm As SAPbouiCOM.Form
            Try
                oUDFFOrm = oApplication.SBO_Application.Forms.GetForm(intFormType.ToString, aform.TypeCount)
                oApplication.Utilities.SetEditText(oUDFFOrm, "U_Z_GPrice", dblGPRice)
            Catch ex As Exception
            End Try
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Calcualte Discount"
    Private Function CalculateDiscount(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim dblCorton, dblPieces, dblItemPieces, dblprice, dblSellingPrice, dblNoofPiece, dblDiscount As Double
        Dim strItemCode, strCardCode, strPostingDate As String
        Dim dtPostingDate As Date
        Dim OTemprec, oRecSet, oDiscRec As SAPbobsCOM.Recordset
        Dim dblGProce As Double = 0
        Try
            aform.Freeze(True)
            If aform.TypeEx = frm_SalesOrder Or aform.TypeEx = frm_Invoice Or aform.TypeEx = frm_ARCreditMemo Or aform.TypeEx = frm_PurchaseOrder Or aform.TypeEx = frm_GRPO Or aform.TypeEx = frm_APInvoice Or aform.TypeEx = frm_APCreditnote Then
                oMatrix = aform.Items.Item("38").Specific
                OTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strPostingDate = oApplication.Utilities.GetEditText(aform, "10")
                strCardCode = oApplication.Utilities.GetEditText(aform, "4")
                If strPostingDate <> "" Then
                    dtPostingDate = oApplication.Utilities.GetDateTimeValue(strPostingDate)
                Else
                    oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.RowCount
                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)

                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        strItemCode = oTest.Fields.Item("ItemCode").Value
                    Else
                        strItemCode = ""
                    End If
                    If strItemCode <> "" Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "14", intRow, "")
                        Try
                            dblCorton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Carton", intRow))
                        Catch ex As Exception
                            dblCorton = 0
                        End Try
                        Try
                            dblPieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pieces", intRow))
                        Catch ex As Exception
                            dblPieces = 0
                        End Try

                        OTemprec.DoQuery("Select isnull(SalPackUn,1) from OITM where ItemCode='" & strItemCode & "'")
                        dblItemPieces = OTemprec.Fields.Item(0).Value
                        strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                        oRecSet.DoQuery(strSQL)
                        If oRecSet.RecordCount > 0 Then
                            Dim strSql As String
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                            oDiscRec.DoQuery(strSql)
                            If oDiscRec.RecordCount > 0 Then
                                dblItemPieces = oDiscRec.Fields.Item("U_Z_No_Pices").Value
                                dblprice = oDiscRec.Fields.Item("U_Z_Price").Value
                                If dblPieces >= dblItemPieces Then
                                    oApplication.Utilities.Message("No of Pieces should be less than than the special prices for alternative  UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    aform.Freeze(False)
                                    oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                    Return False
                                End If
                                dblDiscount = oDiscRec.Fields.Item("U_Z_Discount").Value
                                dblNoofPiece = dblItemPieces
                                dblSellingPrice = oDiscRec.Fields.Item("U_Z_SellPrice").Value
                                ' dblSellingPrice = dblItemPieces * dblprice
                                If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, dblDiscount)
                                Else
                                    Try
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, dblDiscount)
                                    Catch ex As Exception

                                    End Try
                                End If
                            Else
                                dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                                dblNoofPiece = dblItemPieces
                                Dim strBP As String
                                strBP = oApplication.Utilities.GetEditText(aform, "4")
                                dblprice = oApplication.Utilities.GetB1Price(strItemCode, strBP)
                                dblSellingPrice = dblItemPieces * dblprice
                                dblSellingPrice = dblSellingPrice
                            End If
                        Else
                            dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                            dblNoofPiece = dblItemPieces
                        End If
                    End If
                    If dblPieces >= dblItemPieces Then
                        oApplication.Utilities.Message("No of Pieces should be less than Sales UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aform.Freeze(False)
                        oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                        Return False
                    End If

                    Dim dblLinetotal As Double
                    oApplication.Utilities.SetMatrixValues(oMatrix, "14", intRow, "")
                    If dblItemPieces = 0 Then
                        dblLinetotal = (dblCorton * dblprice) + 0 ' (dblPieces * dblprice / dblItemPieces)
                    Else
                        dblLinetotal = (dblCorton * dblprice) + (dblPieces * dblprice / dblItemPieces)
                    End If

                    oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                    aform.Freeze(False)
                    dblprice = dblLinetotal / dblItemPieces
                    Dim strPrice As String

                    aform.Freeze(True)
                    strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow)
                    Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
                    Dim oCurRS As SAPbobsCOM.Recordset
                    oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oCurRS.DoQuery("select currcode from OCRN")
                    For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
                        strPrice = strPrice.Replace(oCurRS.Fields.Item(0).Value, "")
                        oCurRS.MoveNext()
                    Next
                    If strPrice = "" Then
                        dblprice = 0
                    Else
                        dblprice = oApplication.Utilities.getDocumentQuantity(strPrice)
                    End If
                    dblPieces = (dblCorton * dblItemPieces) + dblPieces
                    ' dblprice = dblLinetotal / dblPieces
                    dblSellingPrice = dblItemPieces * dblprice
                    dblSellingPrice = Math.Round(dblSellingPrice, 3)
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice.ToString)

                    If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If (dblPieces > 0) Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                        ElseIf dblPieces < 0 Then
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                            Catch ex As Exception

                            End Try
                        Else
                            dblPieces = 1
                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)

                        End If
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                    Else
                        Try
                            If (dblPieces > 0) Then
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                            ElseIf dblPieces < 0 Then
                                Try
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                                Catch ex As Exception

                                End Try
                            Else
                                dblPieces = 1
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)

                            End If
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                        Catch ex As Exception

                        End Try
                    End If
                    Dim dblPricePerCarton, dblNoofPices, dblLineQty, dblGrossprice As Double
                    dblLineQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow))
                    dblNoofPices = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow))
                    dblPricePerCarton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                    Try
                        dblGrossprice = (dblPricePerCarton / dblNoofPices) * dblLineQty
                    Catch ex As Exception
                        dblGrossprice = 0
                    End Try

                    If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                        oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value = "F" Then
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                dblGProce = dblGProce + dblGrossprice
                            Catch ex As Exception
                            End Try
                        End If
                    Else
                        Try
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                            oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                            If oCombobox.Selected.Value = "F" Then
                                Try
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                    dblGProce = dblGProce + dblGrossprice
                                Catch ex As Exception
                                End Try
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If

            aform.Freeze(False)
            Dim intFormType As Integer
            intFormType = aform.Type
            intFormType = intFormType * -1
            Dim oUDFFOrm As SAPbouiCOM.Form
            Try
                oUDFFOrm = oApplication.SBO_Application.Forms.GetForm(intFormType.ToString, aform.TypeCount)
                oApplication.Utilities.SetEditText(oUDFFOrm, "U_Z_GPrice", dblGProce)
            Catch ex As Exception

            End Try
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End Try
    End Function
#End Region

#Region "Validate Purchase Currency"
    Private Function ValidatePurchaseCurrency(ByVal aform As SAPbouiCOM.Form) As Boolean
        If aform.TypeEx = frm_PurchaseOrder Or aform.TypeEx = frm_GRPO Or aform.TypeEx = frm_APInvoice Or aform.TypeEx = frm_APCreditnote Then
            Dim strDoccurrency, strDocCur, strItemCode, strPostingDate, strSQL As String
            Dim dtPostingDate As Date
            Dim OTemprec, oRecSet, oDiscRec As SAPbobsCOM.Recordset
            oCombobox = aform.Items.Item("70").Specific
            strDocCur = oCombobox.Selected.Value
            If strDocCur = "L" Then
                strDoccurrency = oApplication.Utilities.GetLocalCurrency()
            ElseIf strDocCur = "S" Then
                strDoccurrency = oApplication.Utilities.GetSystemCurrency
            Else
                oCombobox = aform.Items.Item("63").Specific
                strDoccurrency = oCombobox.Selected.Value
            End If

            strPostingDate = oApplication.Utilities.GetEditText(aform, "10")
            strCardCode = oApplication.Utilities.GetEditText(aform, "4")
            If strPostingDate <> "" Then
                dtPostingDate = oApplication.Utilities.GetDateTimeValue(strPostingDate)
            Else
                oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aform.Freeze(False)
                Return False
            End If
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aform.Items.Item("38").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                If oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value <> "" Then
                    strItemCode = oMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                    If strItemCode <> "" Then
                        strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                        oRecSet.DoQuery(strSQL)
                        If oRecSet.RecordCount > 0 Then
                            strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_Currency='" & strDoccurrency & "' and  T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                            oDiscRec.DoQuery(strSQL)
                            If oDiscRec.RecordCount <= 0 Then
                                oApplication.Utilities.Message("Document currency and Special price for alt.UoM Currency should be same : Item Code : " & strItemCode & " in Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oMatrix.Columns.Item("1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                Return False
                            End If
                        Else
                            ' oApplication.Utilities.Message("Item code not defined in  Special price for alt.UoM Currency : Item Code : " & strItemCode & " in Line No : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'oMatrix.Columns.Item("1").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                            ' Return False
                        End If
                    End If
                End If
            Next
            Return True
        End If
        Return True
    End Function
#End Region

#Region "Validation"
    Private Function Validate(ByVal aForm As SAPbouiCOM.Form, ByVal aItemUID As String) As Boolean
        Dim dblCorton, dblPieces, dblItemPieces, dblprice, dblSellingPrice, dblNoofPiece, dblDiscount, dblCartonPrice As Double
        Dim strItemCode, strCardCode, strPostingDate As String
        Dim dtPostingDate As Date
        Dim OTemprec, oRecSet, oDiscRec As SAPbobsCOM.Recordset
        Try
            If aForm.Title.ToUpper.Contains("APPROVED") Then
                Return True
            End If

            aForm.Freeze(True)
            If populatefreeitemdiscount(aForm) = False Then
                aForm.Freeze(False)
                Return False
            End If

            If Validate_NoofPieces(aForm, aItemUID) = False Then
                aForm.Freeze(False)
                Return False
            End If
            If aForm.TypeEx = frm_PurchaseOrder Or aForm.TypeEx = frm_GRPO Or aForm.TypeEx = frm_APInvoice Or aForm.TypeEx = frm_APCreditnote Then
                If ValidatePurchaseCurrency(oForm) = False Then
                    aForm.Freeze(False)
                    Return False
                End If

            End If

            If aItemUID <> "btnDis" Then
                If oApplication.SBO_Application.MessageBox("Are you sure the Discount calculation process are completed?", , "Yes", "No") = 2 Then
                    aForm.Freeze(False)
                    Return False
                Else
                    aForm.Freeze(False)
                    Return True
                End If
                aForm.Freeze(False)
                Return True
            End If

            'If CalculateDiscount(aForm) = False Then
            '    Return False
            'End If
          
           


            Dim dblGProce As Double = 0
            If aForm.TypeEx = frm_SalesOrder Or aForm.TypeEx = frm_Invoice Or aForm.TypeEx = frm_ARCreditMemo Or aForm.TypeEx = frm_PurchaseOrder Or aForm.TypeEx = frm_GRPO Or aForm.TypeEx = frm_APInvoice Or aForm.TypeEx = frm_APCreditnote Then
                oMatrix = aForm.Items.Item("38").Specific
                OTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strPostingDate = oApplication.Utilities.GetEditText(aForm, "10")
                strCardCode = oApplication.Utilities.GetEditText(aForm, "4")
                If strPostingDate <> "" Then
                    dtPostingDate = oApplication.Utilities.GetDateTimeValue(strPostingDate)
                Else
                    oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.RowCount
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        strItemCode = oTest.Fields.Item("ItemCode").Value
                    Else
                        strItemCode = ""
                    End If
                    If strItemCode <> "" Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "14", intRow, "")
                        Try
                            dblCorton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Carton", intRow))
                        Catch ex As Exception
                            dblCorton = 0
                        End Try
                        Try
                            dblPieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pieces", intRow))
                        Catch ex As Exception
                            dblPieces = 0
                        End Try
                        OTemprec.DoQuery("Select isnull(SalPackUn,1) from OITM where ItemCode='" & strItemCode & "'")
                        dblItemPieces = OTemprec.Fields.Item(0).Value
                        Dim otemp As SAPbobsCOM.Recordset
                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        'strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                        'otemp.DoQuery(strSQL)
                        'If otemp.RecordCount > 0 Then
                        '    strSQL = strSQL
                        'Else
                        '    strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  isnull(T1.U_Z_Default,'N')='Y' order by T1.DocEntry Desc"
                        'End If

                        strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"

                        oRecSet.DoQuery(strSQL)
                        If oRecSet.RecordCount > 0 Then
                            strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                            strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                            oRecSet.DoQuery(strSQL)

                        Else
                            strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where  isnull(U_Z_Default,'N')='Y' and T0.U_Z_ItemCode='" & strItemCode & "'" ' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"
                            oRecSet.DoQuery(strSQL)
                        End If


                        '  strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        ' strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"

                        oRecSet.DoQuery(strSQL)
                        dblDiscount = 0
                        Dim strDIscCode, strDiscName As String
                        strDIscCode = ""
                        strDiscName = ""
                        If oRecSet.RecordCount > 0 Then
                            strDIscCode = oRecSet.Fields.Item("U_Z_Dis_Code").Value
                            strDiscName = oRecSet.Fields.Item("U_Z_Dis_Name").Value

                            Dim strSql As String
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                            oDiscRec.DoQuery(strSql)
                            If oDiscRec.RecordCount > 0 Then
                                dblItemPieces = oDiscRec.Fields.Item("U_Z_No_Pices").Value
                                dblprice = oDiscRec.Fields.Item("U_Z_Price").Value
                                dblCartonPrice = dblprice
                                If dblPieces >= dblItemPieces Then
                                    oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                    oApplication.Utilities.Message("No of Pieces should be less than than the special prices for alternative  UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    aForm.Freeze(False)
                                    oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)

                                    Return False
                                End If
                                dblDiscount = oDiscRec.Fields.Item("U_Z_Discount").Value
                                dblNoofPiece = dblItemPieces
                                'dblSellingPrice = oDiscRec.Fields.Item("U_Z_SellPrice").Value
                                dblSellingPrice = dblCartonPrice
                                'dblSellingPrice = dblItemPieces * dblprice
                                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                                Else
                                    Try
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                                    Catch ex As Exception

                                    End Try
                                End If
                            Else
                                dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                                dblNoofPiece = dblItemPieces
                                Dim strBP As String
                                strBP = oApplication.Utilities.GetEditText(aForm, "4")
                                dblprice = oApplication.Utilities.GetB1Price(strItemCode, strBP)
                                dblSellingPrice = dblItemPieces * dblprice
                                dblCartonPrice = dblSellingPrice
                                'dblSellingPrice = dblSellingPrice
                                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                                Else
                                    Try
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If
                        Else
                            dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                            dblNoofPiece = dblItemPieces
                            Dim strBP As String
                            strBP = oApplication.Utilities.GetEditText(aForm, "4")
                            dblprice = oApplication.Utilities.GetB1Price(strItemCode, strBP)
                            dblSellingPrice = dblItemPieces * dblprice
                            dblCartonPrice = dblSellingPrice
                            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                            Else
                                Try
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisCode", intRow, strDIscCode)
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DiscName", intRow, strDiscName)
                        If dblPieces > dblItemPieces Then
                            oApplication.Utilities.Message("No of Pieces should be less than Sales UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aForm.Freeze(False)
                            oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                            Return False
                        End If
                        Dim dblLinetotal As Double
                        ' oApplication.Utilities.SetMatrixValues(oMatrix, "14", intRow, "")
                        dblprice = dblCartonPrice
                        dblLinetotal = (dblCorton * dblprice) + (dblPieces * dblprice / dblItemPieces)
                        Try
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                            Catch ex As Exception
                                oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)
                            End Try
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                        End Try
                        dblprice = dblLinetotal / dblItemPieces
                        Dim strPrice As String
                        strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow)
                        Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
                        Dim oCurRS As SAPbobsCOM.Recordset
                        oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oCurRS.DoQuery("select currcode from OCRN")
                        For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
                            strPrice = strPrice.Replace(oCurRS.Fields.Item(0).Value, "")
                            oCurRS.MoveNext()
                        Next
                        If strPrice = "" Then
                            dblprice = 0
                        Else
                            dblprice = oApplication.Utilities.getDocumentQuantity(strPrice)
                        End If

                        ' dblprice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow))

                        dblPieces = (dblCorton * dblItemPieces) + dblPieces
                        dblprice = dblLinetotal / dblPieces
                        'dblSellingPrice = dblItemPieces * dblprice
                        dblSellingPrice = dblCartonPrice
                        dblSellingPrice = Math.Round(dblSellingPrice, 3)
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice.ToString)
                        If dblPieces = 0 Then
                            dblPieces = 1
                        End If
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                        Else
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                            Catch ex As Exception
                            End Try
                        End If
                        Try
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                        Catch ex As Exception
                            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aForm.Freeze(False)
                            Return False
                        End Try

                        Dim dblPricePerCarton, dblNoofPices, dblLineQty, dblGrossprice As Double
                        dblLineQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow))
                        dblNoofPices = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow))
                        dblPricePerCarton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                        Try
                            dblGrossprice = (dblPricePerCarton / dblNoofPices) * dblLineQty
                        Catch ex As Exception
                            dblGrossprice = 0
                        End Try


                        Dim dblDiscountPer, dblDiscAmt As Double
                        dblDiscountPer = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                        dblDiscAmt = dblGrossprice * dblDiscountPer / 100
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", intRow, dblDiscAmt)
                        dblLinetotal = Math.Round(dblLinetotal, 6) - Math.Round(dblDiscAmt, 6)
                        'oMatrix.Columns.Item("21").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                        'oApplication.SBO_Application.SendKeys("{TAB}")
                        Try
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                            Catch ex As Exception
                                oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)
                            End Try
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                        End Try

                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                            oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                            If oCombobox.Selected.Value = "F" Then
                                Try
                                    'oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                    dblGProce = dblGProce + dblGrossprice
                                Catch ex As Exception
                                End Try
                            End If
                        Else
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                                oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                                If oCombobox.Selected.Value = "F" Then
                                    Try
                                        ' oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                        dblGProce = dblGProce + dblGrossprice
                                    Catch ex As Exception
                                    End Try
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                Next
            End If
            aForm.Freeze(False)
            Dim intFormType As Integer
            intFormType = aForm.Type
            intFormType = intFormType * -1
            Dim oUDFFOrm As SAPbouiCOM.Form
            Try
                oUDFFOrm = oApplication.SBO_Application.Forms.GetForm(intFormType.ToString, aForm.TypeCount)
                oApplication.Utilities.SetEditText(oUDFFOrm, "U_Z_GPrice", dblGProce)
            Catch ex As Exception

            End Try
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
        Return True
    End Function

    Private Function Validate_NoofPieces(ByVal aForm As SAPbouiCOM.Form, ByVal aItemUID As String) As Boolean
        Dim dblCorton, dblPieces, dblItemPieces, dblprice, dblSellingPrice, dblNoofPiece, dblDiscount, dblCartonPrice As Double
        Dim strItemCode, strCardCode, strPostingDate As String
        Dim dtPostingDate As Date
        Dim OTemprec, oRecSet, oDiscRec As SAPbobsCOM.Recordset

        Try
            aForm.Freeze(True)
            Dim dblGProce As Double = 0
            If aForm.TypeEx = frm_SalesOrder Or aForm.TypeEx = frm_Invoice Or aForm.TypeEx = frm_ARCreditMemo Or aForm.TypeEx = frm_PurchaseOrder Or aForm.TypeEx = frm_GRPO Or aForm.TypeEx = frm_APInvoice Or aForm.TypeEx = frm_APCreditnote Then
                oMatrix = aForm.Items.Item("38").Specific
                OTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strPostingDate = oApplication.Utilities.GetEditText(aForm, "10")
                strCardCode = oApplication.Utilities.GetEditText(aForm, "4")
                If strPostingDate <> "" Then
                    dtPostingDate = oApplication.Utilities.GetDateTimeValue(strPostingDate)
                Else
                    oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
                For intRow As Integer = 1 To oMatrix.RowCount
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        strItemCode = oTest.Fields.Item("ItemCode").Value
                    Else
                        strItemCode = ""
                    End If
                    If strItemCode <> "" Then
                        Try
                            dblCorton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Carton", intRow))
                        Catch ex As Exception
                            dblCorton = 0
                        End Try
                        Try
                            dblPieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pieces", intRow))
                        Catch ex As Exception
                            dblPieces = 0
                        End Try

                        OTemprec.DoQuery("Select isnull(SalPackUn,1) from OITM where ItemCode='" & strItemCode & "'")
                        dblItemPieces = OTemprec.Fields.Item(0).Value
                        dblCartonPrice = GetPrice(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                        dblItemPieces = GetPrice(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow))
                        dblDiscount = GetPrice(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                        Dim otemp As SAPbobsCOM.Recordset
                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        'strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                        'otemp.DoQuery(strSQL)
                        'If otemp.RecordCount > 0 Then
                        '    strSQL = strSQL
                        'Else
                        '    strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  isnull(T1.U_Z_Default,'N')='Y' order by T1.DocEntry Desc"
                        'End If

                        strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"

                        oRecSet.DoQuery(strSQL)
                        If oRecSet.RecordCount > 0 Then
                            Dim strSql As String
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                            oDiscRec.DoQuery(strSql)
                            If oDiscRec.RecordCount > 0 Then
                                dblItemPieces = oDiscRec.Fields.Item("U_Z_No_Pices").Value
                                ' dblCartonPrice = oDiscRec.Fields.Item("U_Z_Price").Value
                                If dblPieces >= dblItemPieces Then
                                    oApplication.Utilities.Message("No of Pieces should be less than than the special prices for alternative  UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    aForm.Freeze(False)
                                    oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                    Return False
                                End If
                            Else

                            End If
                        Else
                            dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                            dblNoofPiece = dblItemPieces
                            Dim strBP As String
                            strBP = oApplication.Utilities.GetEditText(aForm, "4")
                            dblprice = oApplication.Utilities.GetB1Price(strItemCode, strBP)
                            dblSellingPrice = dblItemPieces * dblprice
                            ' dblCartonPrice = dblSellingPrice
                            dblSellingPrice = dblSellingPrice
                        End If
                        If dblPieces > dblItemPieces Then
                            oApplication.Utilities.Message("No of Pieces should be less than Sales UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aForm.Freeze(False)
                            oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                            Return False
                        End If
                        Dim dblLinetotal As Double
                        dblprice = dblCartonPrice
                        dblLinetotal = (dblCorton * dblprice) + (dblPieces * dblprice / dblItemPieces)
                        Try

                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                            Catch ex As Exception
                                oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)

                            End Try
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                        End Try
                        Dim strPrice As String
                        strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow)
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                        Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
                        Dim oCurRS As SAPbobsCOM.Recordset
                        oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oCurRS.DoQuery("select currcode from OCRN")
                        For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
                            strPrice = strPrice.Replace(oCurRS.Fields.Item(0).Value, "")
                            oCurRS.MoveNext()
                        Next
                        If strPrice = "" Then
                            dblprice = 0
                        Else
                            dblprice = oApplication.Utilities.getDocumentQuantity(strPrice)
                        End If
                        dblSellingPrice = dblCartonPrice
                        dblSellingPrice = Math.Round(dblSellingPrice, 3)
                        dblPieces = (dblCorton * dblItemPieces) + dblPieces
                        If dblPieces = 0 Then
                            dblPieces = 1

                        End If
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                        Else
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                            Catch ex As Exception
                            End Try
                        End If

                        Dim dblPricePerCarton, dblNoofPices, dblLineQty, dblGrossprice As Double
                        dblLineQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow))
                        dblNoofPices = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow))
                        strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow)
                        oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oCurRS.DoQuery("select currcode from OCRN")
                        For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
                            strPrice = strPrice.Replace(oCurRS.Fields.Item(0).Value, "")
                            oCurRS.MoveNext()
                        Next
                        If strPrice = "" Then
                            dblprice = 0
                        Else
                            dblprice = oApplication.Utilities.getDocumentQuantity(strPrice)
                        End If

                        '  dblSellingPrice = dblNoofPices * dblprice
                        dblSellingPrice = dblCartonPrice
                        '  oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice)
                        dblPricePerCarton = dblSellingPrice ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                        Try
                            dblGrossprice = (dblPricePerCarton / dblNoofPices) * dblLineQty
                        Catch ex As Exception
                            dblGrossprice = 0
                        End Try


                        Dim dblDiscountPer, dblDiscAmt As Double
                        dblDiscountPer = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                        dblDiscAmt = dblGrossprice * dblDiscountPer / 100
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", intRow, dblDiscAmt)
                        dblLinetotal = dblLinetotal - dblDiscAmt
                        'oMatrix.Columns.Item("21").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                        'oApplication.SBO_Application.SendKeys("{TAB}")
                        Try
                            dblLinetotal = Math.Round(dblLinetotal, 6)
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                            Catch ex As Exception
                                oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)

                            End Try
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                        End Try

                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                            oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                            If oCombobox.Selected.Value = "F" Then
                                Try
                                    '  oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                    dblGProce = dblGProce + dblGrossprice
                                Catch ex As Exception
                                End Try
                            End If
                        Else
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                                oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(intRow).Specific
                                If oCombobox.Selected.Value = "F" Then
                                    Try
                                        ' oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "100")
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                        dblGProce = dblGProce + dblGrossprice
                                    Catch ex As Exception
                                    End Try
                                End If
                            Catch ex As Exception
                            End Try
                        End If

                    End If
                Next
                Dim intFormType As Integer
                intFormType = aForm.Type
                intFormType = intFormType * -1
                Dim oUDFFOrm As SAPbouiCOM.Form
                Try
                    oUDFFOrm = oApplication.SBO_Application.Forms.GetForm(intFormType.ToString, aForm.TypeCount)
                    oApplication.Utilities.SetEditText(oUDFFOrm, "U_Z_GPrice", dblGProce)
                Catch ex As Exception
                End Try
            End If
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
        Return True
    End Function
    Private Sub PopulateQuantity(ByVal aForm As SAPbouiCOM.Form, ByVal arow As Integer, Optional ByVal aBool As Boolean = False)
        Dim dblCorton, dblPieces, dblItemPieces, dblprice, dblSellingPrice, dblNoofPiece, dblDiscount, dblCartonPrice As Double
        Dim strItemCode, strCardCode, strPostingDate As String
        Dim dtPostingDate As Date
        Dim OTemprec, oRecSet, oDiscRec As SAPbobsCOM.Recordset
        Try
            aForm.Freeze(True)
            If aForm.TypeEx = frm_SalesOrder Or aForm.TypeEx = frm_Invoice Or aForm.TypeEx = frm_ARCreditMemo Or aForm.TypeEx = frm_PurchaseOrder Or aForm.TypeEx = frm_GRPO Or aForm.TypeEx = frm_APInvoice Or aForm.TypeEx = frm_APCreditnote Then
                oMatrix = aForm.Items.Item("38").Specific
                OTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strPostingDate = oApplication.Utilities.GetEditText(aForm, "10")
                strCardCode = oApplication.Utilities.GetEditText(aForm, "4")
                If strPostingDate <> "" Then
                    dtPostingDate = oApplication.Utilities.GetDateTimeValue(strPostingDate)
                Else
                    dtPostingDate = Now.Date
                End If
                For intRow As Integer = arow To arow
                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "' and isnull(TreeType,'')<>'S'")
                    If oTest.RecordCount > 0 Then
                        strItemCode = oTest.Fields.Item("ItemCode").Value
                    Else
                        strItemCode = ""
                    End If
                    If strItemCode <> "" Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "14", intRow, "")
                        Try
                            dblCorton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Carton", intRow))
                        Catch ex As Exception
                            dblCorton = 0
                        End Try
                        dblCartonPrice = 0
                        Try
                            dblPieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pieces", intRow))
                        Catch ex As Exception
                            dblPieces = 0
                        End Try
                        OTemprec.DoQuery("Select isnull(SalPackUn,1) from OITM where ItemCode='" & strItemCode & "'")
                        dblItemPieces = OTemprec.Fields.Item(0).Value

                        Dim strSql As String
                        dblDiscount = 0
                        Dim otemp As SAPbobsCOM.Recordset
                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strSql = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        'strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"
                        'otemp.DoQuery(strSql)
                        'If otemp.RecordCount > 0 Then
                        '    strSql = strSql
                        'Else
                        '    strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  isnull(T1.U_Z_Default,'N')='Y' order by T1.DocEntry Desc"
                        'End If

                        strSql = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                        strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"

                        oRecSet.DoQuery(strSql)
                        If oRecSet.RecordCount > 0 Then
                            strSql = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"
                            oRecSet.DoQuery(strSql)

                        Else
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where  isnull(U_Z_Default,'N')='Y' and T0.U_Z_ItemCode='" & strItemCode & "'" ' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"
                            oRecSet.DoQuery(strSql)

                        End If

                        Dim strDIscCode, strDiscName As String
                        strDIscCode = ""
                        strDiscName = ""

                        If oRecSet.RecordCount > 0 Then
                            strDIscCode = oRecSet.Fields.Item("U_Z_Dis_Code").Value
                            strDiscName = oRecSet.Fields.Item("U_Z_Dis_Name").Value
                            strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & strItemCode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                            oDiscRec.DoQuery(strSql)
                            If oDiscRec.RecordCount > 0 Then
                                dblItemPieces = oDiscRec.Fields.Item("U_Z_No_Pices").Value
                                dblprice = oDiscRec.Fields.Item("U_Z_Price").Value
                                dblCartonPrice = oDiscRec.Fields.Item("U_Z_Price").Value
                                dblDiscount = oDiscRec.Fields.Item("U_Z_Discount").Value
                                dblNoofPiece = dblItemPieces
                                dblSellingPrice = oDiscRec.Fields.Item("U_Z_SellPrice").Value
                                'dblSellingPrice = dblItemPieces * dblprice
                                If aBool = False Then
                                    dblDiscount = dblDiscount
                                Else
                                    dblDiscount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                                End If
                                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    'oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, dblDiscount)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                                Else
                                    Try
                                        '  oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, dblDiscount)
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                                    Catch ex As Exception
                                    End Try
                                End If
                            Else
                                dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                                If aBool = False Then
                                    dblDiscount = dblDiscount
                                Else
                                    dblDiscount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                                End If
                                dblNoofPiece = dblItemPieces
                            End If
                        Else
                            dblDiscount = 0 'oDiscRec.Fields.Item("U_Z_Discount").Value
                            If aBool = False Then
                                dblDiscount = dblDiscount
                            Else
                                dblDiscount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                            End If
                            dblNoofPiece = dblItemPieces
                        End If

                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisCode", intRow, strDIscCode)
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DiscName", intRow, strDiscName)
                    End If
                    Dim dblLinetotal As Double
                    dblprice = dblCartonPrice
                    If (dblPieces <= 0) Then
                        dblPieces = 0
                    End If
                    If dblItemPieces = 0 Then
                        dblLinetotal = (dblCorton * dblprice) + (dblPieces * dblprice / dblItemPieces)
                    Else
                        dblLinetotal = (dblCorton * dblprice) '+ (dblPieces * dblprice / dblItemPieces)
                    End If
                    dblLinetotal = (dblCorton * dblprice) + (dblPieces * dblprice / dblItemPieces)
                    Try
                        dblLinetotal = Math.Round(dblLinetotal, 6)
                        Try
                            oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)

                        End Try
                    Catch ex As Exception
                        oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                    End Try

                    Dim strPrice As String
                    oApplication.SBO_Application.SendKeys("{TAB}")
                    strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "14", intRow)
                    Dim strLocalCurrency, strBPCurrency, strsystemcurrency, strdoctotal, strCurrency As String
                    Dim oCurRS As SAPbobsCOM.Recordset
                    oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oCurRS.DoQuery("select currcode from OCRN")
                    For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
                        strPrice = strPrice.Replace(oCurRS.Fields.Item(0).Value, "")
                        oCurRS.MoveNext()
                    Next
                    If strPrice = "" Then
                        dblprice = 0
                    Else
                        dblprice = oApplication.Utilities.getDocumentQuantity(strPrice)
                    End If

                    dblPieces = (dblCorton * dblItemPieces) + dblPieces
                    dblSellingPrice = dblCartonPrice ' dblItemPieces * dblprice
                    dblSellingPrice = Math.Round(dblSellingPrice, 3)
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_SPrice", intRow, dblSellingPrice.ToString)
                    Try
                        If (dblPieces > 0) Then
                            oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                        ElseIf dblPieces < 0 Then
                            Try
                                oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces.ToString)
                            Catch ex As Exception

                            End Try
                        Else
                            dblPieces = 1
                            Try
                                'oApplication.Utilities.SetMatrixValues(oMatrix, "11", intRow, dblPieces)
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception

                    End Try
                    '  oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, dblDiscount)
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, dblDiscount)
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pack", intRow, dblItemPieces)
                    Dim dblPricePerCarton, dblNoofPices, dblLineQty, dblGrossprice As Double
                    dblLineQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow))
                    dblNoofPices = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow))
                    dblPricePerCarton = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                    Try
                        dblGrossprice = (dblPricePerCarton / dblNoofPices) * dblLineQty
                    Catch ex As Exception
                        dblGrossprice = 0
                    End Try
                    Dim dblDiscountPer, dblDiscAmt As Double
                    dblDiscountPer = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", intRow))
                    dblDiscAmt = dblGrossprice * dblDiscountPer / 100
                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", intRow, dblDiscAmt)
                    dblLinetotal = Math.Round(dblGrossprice, 6) - Math.Round(dblDiscAmt, 6)
                    ' oMatrix.Columns.Item("21").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                    Try
                        dblLinetotal = Math.Round(dblLinetotal, 6)
                        Try
                            oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, dblLinetotal)
                        Catch ex As Exception
                            oApplication.Utilities.SetMatrixValues(oMatrix, "23", intRow, dblLinetotal)

                        End Try
                    Catch ex As Exception
                        oApplication.Utilities.SetMatrixValues(oMatrix, "22", intRow, dblLinetotal)
                    End Try
                    '  oApplication.SBO_Application.SendKeys("{TAB}")
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                        Dim strSeries As String
                        oCombobox = aForm.Items.Item("88").Specific
                        strSeries = oCombobox.Selected.Description
                        If aForm.TypeEx = frm_SalesOrder And (strSeries = "GFC" Or strSeries = "SR") Then
                            oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(arow).Specific
                            If oCombobox.Selected.Value = "F" Then
                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, 0)
                                ' oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, 0)
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", intRow, dblGrossprice)
                            End If
                        End If
                    Else
                        Try
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, dblGrossprice)
                            Dim strSeries As String
                            oCombobox = aForm.Items.Item("88").Specific
                            strSeries = oCombobox.Selected.Description
                            If aForm.TypeEx = frm_SalesOrder And (strSeries = "GFC" Or strSeries = "SR") Then
                                oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(arow).Specific
                                'oCombobox.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                If oCombobox.Selected.Value = "F" Then
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "21", intRow, 0)
                                    ' oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", intRow, 0)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", intRow, "100")
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", intRow, dblGrossprice)
                                End If
                            End If

                        Catch ex As Exception

                        End Try
                    End If
                Next
            End If

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region


    Private Function GetPrice(ByVal aPrice As String) As Double
        Dim oCurRS As SAPbobsCOM.Recordset
        Dim dbPrice As Double
        oCurRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCurRS.DoQuery("select currcode from OCRN")
        For intRow11 As Integer = 0 To oCurRS.RecordCount - 1
            aPrice = aPrice.Replace(oCurRS.Fields.Item(0).Value, "")
            oCurRS.MoveNext()
        Next
        If aPrice = "" Then
            dbPrice = 0
        Else
            dbPrice = oApplication.Utilities.getDocumentQuantity(aPrice)
        End If
        Return dbPrice
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BPMaster Then
                If pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    oApplication.Utilities.AddControls(oForm, "btnDis", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Special Prices", , , 130)
                End If
                If pVal.Before_Action = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    If pVal.ItemUID = "btnDis" And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                        Dim oBj As New clsDiscMapping
                        oBj.databind(oApplication.Utilities.GetEditText(oForm, "5"))
                    End If
                End If
            End If

            If pVal.FormTypeEx = frm_PurchaseOrder Or pVal.FormTypeEx = frm_GRPO Or pVal.FormTypeEx = frm_APCreditnote Or pVal.FormTypeEx = frm_APInvoice Or pVal.FormTypeEx = frm_SalesOrder Or pVal.FormTypeEx = frm_Invoice Or pVal.FormTypeEx = frm_ARCreditMemo Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If Validate(oForm, pVal.ItemUID) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pVal.FormTypeEx = frm_PurchaseOrder Or pVal.FormTypeEx = frm_SalesOrder) Then
                                    If Validate(oForm, pVal.ItemUID) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And (pVal.ColUID = "11" Or pVal.ColUID = "14" Or pVal.ColUID = "U_Z_Pack" Or pVal.ColUID = "U_Z_GPrice") And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'oMatrix = oForm.Items.Item("38").Specific
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "btnDis", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Discount Calculation", , , 130)
                                Recalculate(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "14" And pVal.CharPressed = 9 Then
                                    If ValiateCustomerRerNo(oForm) = False Then
                                        oForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_SPrice" Or pVal.ColUID = "U_Z_Discount") And pVal.CharPressed = 9 Then
                                    Dim dblPricePerCarton, dblCartonPrice, dblLinetotal, dblsellingprice, dblPrice, dblitempieces, introw, dblNoofPices, dblCartons, dblLineQty, dblGrossprice As Double
                                    Dim strPrice As String
                                    oForm.Freeze(True)
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    introw = pVal.Row
                                    dblCartons = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Carton", introw))
                                    dblNoofPices = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pieces", introw))
                                    dblitempieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", introw))
                                    strPrice = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", introw)
                                    dblPrice = GetPrice(strPrice)
                                    dblsellingprice = dblPrice
                                    dblCartonPrice = dblPrice
                                    dblPrice = dblCartonPrice
                                    If dblitempieces = 0 Then
                                        dblLinetotal = (dblCartons * dblPrice) + 0 ' (dblNoofPices * dblPrice / dblitempieces)
                                    Else
                                        dblLinetotal = (dblCartons * dblPrice) + (dblNoofPices * dblPrice / dblitempieces)

                                    End If
                                    Dim dblDiscount As Double
                                    dblDiscount = GetPrice(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", introw))
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "14", introw, "")
                                    Try
                                        dblLinetotal = Math.Round(dblLinetotal, 6)
                                        Try
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "21", introw, dblLinetotal)
                                        Catch ex As Exception
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "23", introw, dblLinetotal)

                                        End Try
                                    Catch ex As Exception
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "22", introw, dblLinetotal)
                                    End Try
                                    dblPricePerCarton = dblsellingprice ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", intRow))
                                    Try
                                        dblLineQty = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", introw))
                                        dblNoofPices = dblitempieces 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", introw))
                                        dblPricePerCarton = dblCartonPrice  'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_SPrice", introw))
                                        Try
                                            dblGrossprice = (dblPricePerCarton / dblNoofPices) * dblLineQty
                                        Catch ex As Exception
                                            dblGrossprice = 0
                                        End Try

                                        Dim dblDiscountPer, dblDiscAmt As Double
                                        dblDiscountPer = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Discount", introw))
                                        dblDiscAmt = dblGrossprice * dblDiscountPer / 100
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisAmt", introw, dblDiscAmt)
                                        '  dblLinetotal = dblLinetotal - dblDiscAmt

                                        dblLinetotal = Math.Round(dblLinetotal, 6) - Math.Round(dblDiscAmt, 6)
                                        Try
                                            dblLinetotal = Math.Round(dblLinetotal, 6)
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "21", introw, dblLinetotal)
                                            Catch ex As Exception
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "23", introw, dblLinetotal)
                                            End Try
                                        Catch ex As Exception
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "22", introw, dblLinetotal)
                                        End Try
                                        oForm.Freeze(False)
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_GPrice", introw, dblGrossprice)
                                        Try
                                            If pVal.ColUID = "U_Z_Discount" Then
                                                oMatrix.Columns.Item("1").Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                            Else
                                                oMatrix.Columns.Item("U_Z_Discount").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                            End If
                                        Catch ex As Exception
                                        End Try
                                    Catch ex As Exception
                                        dblGrossprice = 0
                                    End Try
                                End If

                                '   If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_Carton" Or pVal.ColUID = "U_Z_Discount") And pVal.CharPressed = 9 Then
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_Carton") And pVal.CharPressed = 9 Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim strSeries, stritem As String
                                    stritem = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("Select * from OITM where ItemCode='" & stritem & "'")
                                    If oTest.RecordCount > 0 Then
                                        If stritem <> "" And stritem <> "*" Then
                                            oForm.Freeze(True)
                                            If pVal.ColUID = "U_Z_Discount" Then
                                                PopulateQuantity(oForm, pVal.Row, True)
                                            Else
                                                PopulateQuantity(oForm, pVal.Row)
                                            End If

                                            oCombobox = oForm.Items.Item("88").Specific
                                            Try
                                                strSeries = oCombobox.Selected.Description
                                            Catch ex As Exception
                                                strSeries = ""
                                            End Try
                                            ' PopulateQuantity(oForm, pVal.Row)
                                            If oForm.TypeEx = frm_SalesOrder And (strSeries = "GFC" Or strSeries = "SR") Then
                                                oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(pVal.Row).Specific
                                                oCombobox.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            End If
                                            oForm.Freeze(False)
                                            Try
                                                If pVal.ColUID = "U_Z_Discount" Then
                                                    oMatrix.Columns.Item("1").Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                                Else
                                                    oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                                End If

                                            Catch ex As Exception
                                            End Try

                                        End If
                                    End If
                                    'oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                End If
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.CharPressed = 9 Then
                                    Dim strSeries, stritem As String
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Try
                                        stritem = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Catch ex As Exception
                                        stritem = ""
                                    End Try
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("Select * from OITM where ItemCode='" & stritem & "' and isnull(TreeType,'')<>'S'")
                                    If oTest.RecordCount > 0 Then
                                        If stritem <> "" And stritem <> "*" Then
                                            oCombobox = oForm.Items.Item("88").Specific
                                            Try
                                                strSeries = oCombobox.Selected.Description
                                            Catch ex As Exception
                                                strSeries = ""
                                            End Try
                                            PopulateQuantity(oForm, pVal.Row)
                                            Try
                                                If oForm.TypeEx = frm_SalesOrder And (strSeries = "GFC" Or strSeries = "SR") Then
                                                    oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(pVal.Row).Specific
                                                    oCombobox.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                                End If
                                            Catch ex As Exception
                                            End Try
                                            ' oMatrix.Columns.Item("U_Z_Carton").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                        End If
                                        ' oForm.Freeze(True)
                                        oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oMatrix.Columns.Item("U_Z_Carton").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                        'oForm.Freeze(False)
                                    End If
                                End If
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_Pieces") And pVal.CharPressed = 9 Then
                                    Dim oTempRec As SAPbobsCOM.Recordset
                                    Dim dblItemPieces, dblPieces As Double
                                    Dim stritemcode As String
                                    oMatrix = oForm.Items.Item("38").Specific
                                    stritemcode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("Select * from OITM where ItemCode='" & stritemcode & "'")
                                    If oTest.RecordCount > 0 Then
                                        If stritemcode <> "" And stritemcode <> "*" Then
                                            oForm.Freeze(True)
                                            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            ' oTempRec.DoQuery("Select isnull(NumInSale,1) from OITM where ItemCode='" & stritemcode & "'")
                                            oTempRec.DoQuery("Select isnull(SalPackUn,1) from OITM where ItemCode='" & stritemcode & "' and isnull(TreeType,'')<>'S'")
                                            dblItemPieces = oTempRec.Fields.Item(0).Value
                                            Try
                                                dblPieces = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row))
                                            Catch ex As Exception
                                                dblPieces = 0
                                            End Try
                                            Dim oRecSet, oDiscRec As SAPbobsCOM.Recordset
                                            Dim strPostingdate As String
                                            Dim dtPostingdate As Date
                                            strCardCode = oApplication.Utilities.GetEditText(oForm, "4")
                                            strPostingdate = oApplication.Utilities.GetEditText(oForm, "10")
                                            strCardCode = oApplication.Utilities.GetEditText(oForm, "4")
                                            If strPostingdate <> "" Then
                                                dtPostingdate = oApplication.Utilities.GetDateTimeValue(strPostingdate)
                                            Else
                                                dtPostingdate = Now.Date
                                            End If
                                            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oDiscRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            ' oRecSet.DoQuery("Select * from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate order by U_Z_FromDate Desc")

                                            Dim otemp As SAPbobsCOM.Recordset
                                            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            'strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                                            'strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                                            'otemp.DoQuery(strSQL)
                                            'If otemp.RecordCount > 0 Then
                                            '    strSQL = strSQL
                                            'Else
                                            '    strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  isnull(T1.U_Z_Default,'N')='Y' order by T1.DocEntry Desc"
                                            'End If

                                            strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                                            strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"

                                            oRecSet.DoQuery(strSQL)
                                            If oRecSet.RecordCount > 0 Then
                                                strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                                                strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"
                                                oRecSet.DoQuery(strSQL)

                                            Else
                                                strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where  isnull(U_Z_Default,'N')='Y' and T0.U_Z_ItemCode='" & stritemcode & "'" ' and  T1.U_Z_Dis_Code in (" & strSql & ") order by T1.DocEntry Desc"
                                                oRecSet.DoQuery(strSQL)

                                            End If



                                            'strSQL = "Select U_Z_Dis_Code from [@Z_Dis_Mapping] where U_Z_CardCode='" & strCardCode & "' and '" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_FromDate and U_Z_ToDate"
                                            'strSQL = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  T1.U_Z_Dis_Code in (" & strSQL & ") order by T1.DocEntry Desc"

                                            oRecSet.DoQuery(strSQL)
                                            Dim strDIscCode, strDiscName As String
                                            strDIscCode = ""
                                            strDiscName = ""
                                            If oRecSet.RecordCount > 0 Then
                                                strDIscCode = oRecSet.Fields.Item("U_Z_Dis_Code").Value
                                                strDiscName = oRecSet.Fields.Item("U_Z_Dis_Name").Value
                                                Dim strSql As String
                                                strSql = "Select * from [@Z_DIS1] T0 inner join [@Z_ODIS] T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_ItemCode='" & stritemcode & "' and  T1.U_Z_Dis_Code='" & oRecSet.Fields.Item("U_Z_Dis_Code").Value & "'"
                                                oDiscRec.DoQuery(strSql)
                                                If oDiscRec.RecordCount > 0 Then
                                                    dblItemPieces = oDiscRec.Fields.Item("U_Z_No_Pices").Value
                                                    If dblPieces >= dblItemPieces Then
                                                        oApplication.Utilities.Message("No of Pieces should be less than the special prices for alternative UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oForm.Freeze(False)
                                                        oMatrix.Columns.Item("U_Z_Pieces").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                                        Exit Sub
                                                    End If
                                                End If
                                            End If


                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DisCode", pVal.Row, strDIscCode)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_DiscName", pVal.Row, strDiscName)
                                            If dblPieces >= dblItemPieces Then
                                                oApplication.Utilities.Message("No of Pieces should be less than Sales UoM", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oForm.Freeze(False)
                                                oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                                Exit Sub
                                            End If
                                            PopulateQuantity(oForm, pVal.Row)
                                            ' PopulateQuantity(oForm, pVal.Row)
                                            'oMatrix.Columns.Item("U_Z_Discount").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                            oForm.Freeze(False)
                                            oMatrix.Columns.Item("U_Z_SPrice").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_ItemType" Then
                                    Try
                                        If oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row) <> "" And oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row) <> "*" Then
                                            oForm.Freeze(True)
                                            oCombobox = oMatrix.Columns.Item("U_Z_ItemType").Cells.Item(pVal.Row).Specific
                                            If oCombobox.Selected.Value = "F" Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", pVal.Row, "100")
                                            Else
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Discount", pVal.Row, "0")
                                            End If
                                            oForm.Freeze(False)
                                            PopulateQuantity(oForm, pVal.Row, True)
                                            calculateGrossPr(oForm)
                                            oMatrix.Columns.Item("U_Z_SPrice").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                        End If
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "btnDis" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Validate(oForm, pVal.ItemUID)
                                End If
                                If pVal.ItemUID = "btnDis" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE) And (pVal.FormTypeEx = frm_PurchaseOrder Or pVal.FormTypeEx = frm_SalesOrder) Then
                                    Validate(oForm, pVal.ItemUID)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            If ex.Message.StartsWith("Input string") Then

            Else
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Reset Discount %"
    Private Sub resetDiscount(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("38").Specific
            For intRow As Integer = 1 To oMatrix.RowCount
                Try
                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", intRow, "0")
                Catch ex As Exception

                End Try
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try

    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                ' Case mnu_InvSO
                Case "1287"
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If oForm.TypeEx = frm_SalesOrder Then
                            resetDiscount(oForm)
                        End If
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Recalculate Cartons and pieces"
    Private Sub Recalculate(ByVal aForm As SAPbouiCOM.Form)
        Dim oForm As SAPbouiCOM.Form
        Dim strItem, strQty, strBastType, strpack As String
        Dim dblQty, dblPack As Double
        Dim intCarton, intPieces As Integer
        oForm = aForm
        oForm = aForm
        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If oForm.TypeEx = frm_GRPO Or oForm.TypeEx = frm_APInvoice Or oForm.TypeEx = frm_APCreditnote Then
                oForm.Freeze(True)
                oMatrix = oForm.Items.Item("38").Specific
                For intRow As Integer = 1 To oMatrix.RowCount
                    strItem = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                    If strItem <> "" Then
                        strBastType = oApplication.Utilities.getMatrixValues(oMatrix, "43", intRow)
                        If strBastType <> "-1" Then
                            strQty = oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow)
                            strpack = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow)
                            dblQty = oApplication.Utilities.getDocumentQuantity(strQty)
                            dblPack = oApplication.Utilities.getDocumentQuantity(strpack)
                            intCarton = Math.Floor(dblQty / dblPack)
                            intPieces = dblQty - intCarton * dblPack
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Carton", intRow, intCarton)
                            oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pieces", intRow, intPieces)
                        End If
                    End If
                Next
                oForm.Freeze(False)
            End If
        End If
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                Dim oForm As SAPbouiCOM.Form
                Dim strItem, strQty, strBastType, strpack As String
                Dim dblQty, dblPack As Double
                Dim intCarton, intPieces As Integer
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If oForm.TypeEx = frm_GRPO Then
                        oForm.Freeze(True)
                        oMatrix = oForm.Items.Item("38").Specific
                        For intRow As Integer = 1 To oMatrix.RowCount
                            strItem = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                            If strItem <> "" Then
                                strBastType = oApplication.Utilities.getMatrixValues(oMatrix, "43", intRow)
                                If strBastType <> "-1" Then
                                    strQty = oApplication.Utilities.getMatrixValues(oMatrix, "11", intRow)
                                    strpack = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_Pack", intRow)
                                    dblQty = oApplication.Utilities.getDocumentQuantity(strQty)
                                    dblPack = oApplication.Utilities.getDocumentQuantity(strpack)
                                    intCarton = dblQty / dblPack
                                    intPieces = intCarton * dblPack - dblQty
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Carton", intRow, intCarton)
                                    oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_Pieces", intRow, intPieces)
                                End If

                            End If

                        Next
                        oForm.Freeze(False)
                    End If
                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
End Class
