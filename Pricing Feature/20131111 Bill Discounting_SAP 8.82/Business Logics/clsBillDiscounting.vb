Imports System.IO
Public Class clsBillDiscounting
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private strSelectedFilePath, strSelectedFolderPath As String
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, spath As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_BillDiscount, frm_BillDiscount)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            oForm.DataSources.UserDataSources.Add("Bank", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("from", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("to", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("dtPosting", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("Batch", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("FileName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("BankFile", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("DisAmt", SAPbouiCOM.BoDataType.dt_SUM)
            oForm.DataSources.UserDataSources.Add("SelBatch", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("NetAmt", SAPbouiCOM.BoDataType.dt_SUM)
            oEditText = oForm.Items.Item("33").Specific
            oEditText.DataBind.SetBound(True, "", "SelBatch")

            oCombobox = oForm.Items.Item("9").Specific
            oCombobox.DataBind.SetBound(True, "", "Bank")
            oApplication.Utilities.FillComboBox(oCombobox, "Select BankCode,BankName from ODSC order by BankCode")
            oForm.Items.Item("9").DisplayDesc = True
            oEditText = oForm.Items.Item("12").Specific
            oEditText.DataBind.SetBound(True, "", "from")
            oEditText = oForm.Items.Item("15").Specific
            oEditText.DataBind.SetBound(True, "", "to")


            oEditText = oForm.Items.Item("22").Specific
            oEditText.DataBind.SetBound(True, "", "DisAmt")
            oEditText = oForm.Items.Item("32").Specific
            oEditText.DataBind.SetBound(True, "", "Batch")
            oEditText = oForm.Items.Item("26").Specific
            oEditText.DataBind.SetBound(True, "", "FileName")
            oEditText = oForm.Items.Item("29").Specific
            oEditText.DataBind.SetBound(True, "", "BankFile")
            oEditText = oForm.Items.Item("36").Specific
            oEditText.DataBind.SetBound(True, "", "NetAmt")

            oEditText = oForm.Items.Item("38").Specific
            oEditText.DataBind.SetBound(True, "", "dtPosting")
            oEditText.String = "t"
            oApplication.SBO_Application.SendKeys("{TAB}")
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try

    End Sub
#Region "Get Details"
    Private Sub GetDetails(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim strQuery, strSalesQuery, strReturnsQuery, strFromdate, strTodate, strbank, strCondition, strBPCondition, strBatch As String
            Dim dtFromdate, dttodate As Date
            Dim oTempRec As SAPbobsCOM.Recordset
            oCombobox = aForm.Items.Item("9").Specific
            strbank = oCombobox.Selected.Value
            If strbank = "" Then
                oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            Else
                strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
            End If
            strBatch = oApplication.Utilities.getEdittextvalue(aForm, "33")
            If strBatch = "" Then
                oApplication.Utilities.Message("Select the Batch Number", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            Else
                oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTempRec.DoQuery("Select isnull(U_Z_Imported,'N'),U_Z_DateFrom,U_Z_DateTo from [@Z_Bill_Export] where U_Z_BatchNumber='" & strBatch & "' and U_Z_BankCode='" & strbank & "'")
                If oTempRec.Fields.Item(0).Value = "Y" Then
                    oApplication.Utilities.Message("Selected batch already imported to SAP ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("6").Enabled = False
                Else
                    dtFromdate = oTempRec.Fields.Item(1).Value
                    dttodate = oTempRec.Fields.Item(2).Value
                    aForm.Items.Item("6").Enabled = True
                End If
            End If
            Dim strDate As String
            strDate = Now.Date
            Dim dtdate As Date
            ' dtdate = oApplication.Utilities.GetDateTimeValue(strDate)
            'oApplication.Utilities.SetEditText(aForm, "38", "t")
            'oApplication.SBO_Application.SendKeys("{TAB}")
            'aForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Right)


            strCondition = "1=1"
            ' strQuery = "Select isnull(T1.U_Z_KFHNO,''),x.CardCode,x.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " group by CardCode,Cardname "
            ' strQuery = strQuery & " union select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y'  and " & strCondition & " group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname group by isnull(T1.U_Z_KFHNO,''),X.CardCode,X.CardName"


            strQuery = "Select isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from "
            strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where T0.DocStatus<>'C' and isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
            strQuery = strQuery & " union all  select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and  T0.DocStatus<>'C' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode   group by isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName"



            oGrid = aForm.Items.Item("20").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)

            oGrid.Columns.Item(0).TitleObject.Caption = "Bank ref No"
            oGrid.Columns.Item(1).TitleObject.Caption = "Customer Code"
            oEditTextColumn = oGrid.Columns.Item(1)
            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            oGrid.Columns.Item(2).TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item(3).TitleObject.Caption = "Sales"
            oEditTextColumn = oGrid.Columns.Item(3)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(4).TitleObject.Caption = "Returns"
            oEditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(5).TitleObject.Caption = "Net Sales"
            oEditTextColumn = oGrid.Columns.Item(5)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oForm.Items.Item("20").Enabled = False
            oGrid.AutoResizeColumns()

            strSalesQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from OINV T0 where T0.DocStatus<>'C' and  isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and T0.CardCode in (" & strBPCondition & ") order by CardCode  "
            oGrid = aForm.Items.Item("23").Specific
            oGrid.DataTable.ExecuteQuery(strSalesQuery)

            oGrid.Columns.Item(0).TitleObject.Caption = "DocEntry"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
            oGrid.Columns.Item(1).TitleObject.Caption = "DocNum"
            oGrid.Columns.Item(2).TitleObject.Caption = "Customer Code"
            oEditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            oGrid.Columns.Item(3).TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item(4).TitleObject.Caption = "Document Total"
            oEditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oForm.Items.Item("23").Enabled = False
            oGrid.AutoResizeColumns()

            strReturnsQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from ORIN T0 where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and  T0.DocStatus<>'C' and isnull(U_Z_Exported,'N')='Y' and T0.CardCode in (" & strBPCondition & ")  order by CardCode"
            oGrid = aForm.Items.Item("24").Specific
            oGrid.DataTable.ExecuteQuery(strReturnsQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "DocEntry"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
            oGrid.Columns.Item(1).TitleObject.Caption = "DocNum"
            oGrid.Columns.Item(2).TitleObject.Caption = "Customer Code"
            oEditTextColumn = oGrid.Columns.Item(2)
            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            oGrid.Columns.Item(3).TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item(4).TitleObject.Caption = "Document Total"
            oEditTextColumn = oGrid.Columns.Item(4)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oForm.Items.Item("24").Enabled = False
            oGrid.AutoResizeColumns()
            Dim dblDisAmt, dblNetSales As Double
            oTempRec.DoQuery("Select * from [@Z_Bill_Export] where U_Z_BatchNumber='" & strBatch & "' and U_Z_BankCode='" & strbank & "'")
            If oTempRec.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aForm, "22", oTempRec.Fields.Item("U_Z_DiscountAmount").Value)
                dblDisAmt = oTempRec.Fields.Item("U_Z_DiscountAmount").Value


                strQuery = "Select sum(x.INV)-sum(x.RETU) from "
                strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where T0.DocStatus<>'C' and isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
                strQuery = strQuery & " union all  select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and  T0.DocStatus<>'C' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode "
                oTempRec.DoQuery(strQuery)
                dblNetSales = oTempRec.Fields.Item(0).Value
                dblDisAmt = dblNetSales - dblDisAmt
                oApplication.Utilities.SetEditText(aForm, "36", dblDisAmt)
            End If
            oForm.PaneLevel = 3
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


#Region "Post Incoming Documents"
    Private Function postincomingDocuments(ByVal aForm As SAPbouiCOM.Form) As Boolean

        Dim strbatchNumber, strPayCardCode, strCardName, strCashAccount, strQuery, strbankCode, strBankCreditAcct, strBankDebitAcct As String
        Dim dblNetSales, dblDiscountAmount As Double
        Dim oTempRec As SAPbobsCOM.Recordset
        Dim oDoc As SAPbobsCOM.Payments
        Dim dtPostingdate As Date
        Dim stPostingdate As String
        Dim blnLineExist As Boolean = False
        Dim blnDocumentCreated As Boolean = False
        Dim strType As String
        strbatchNumber = oApplication.Utilities.getEdittextvalue(aForm, "32")

        oCombobox = aForm.Items.Item("9").Specific
        strbankCode = oCombobox.Selected.Value
        If strbankCode = "" Then
            oApplication.Utilities.Message("Select the Bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRec.DoQuery("Select * from ODSC where BankCode='" & strbankCode & "'")
            If oTempRec.Fields.Item("U_CreditAc").Value = "" Or oTempRec.Fields.Item("U_DebitAc").Value = "" Then
                oApplication.Utilities.Message("Account mappings are missing for the selected bankcode", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                strBankCreditAcct = oTempRec.Fields.Item("U_CreditAc").Value
                strBankDebitAcct = oTempRec.Fields.Item("U_DebitAc").Value
                strCashAccount = oTempRec.Fields.Item("U_CashAc").Value
            End If
            oTempRec.DoQuery("Select AcctCode from OACT where Formatcode ='" & strBankCreditAcct & "'")
            If oTempRec.RecordCount > 0 Then
                strBankCreditAcct = oTempRec.Fields.Item(0).Value
            End If
            oTempRec.DoQuery("Select AcctCode from OACT where Formatcode ='" & strBankDebitAcct & "'")
            If oTempRec.RecordCount > 0 Then
                strBankDebitAcct = oTempRec.Fields.Item(0).Value
            End If

            oTempRec.DoQuery("Select AcctCode from OACT where Formatcode ='" & strCashAccount & "'")
            If oTempRec.RecordCount > 0 Then
                strBankCreditAcct = oTempRec.Fields.Item(0).Value
            End If
        End If

        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec.DoQuery("Select * from [@Z_Bill_Export] where U_Z_BankCode='" & strbankCode & "' and  U_Z_BatchNumber='" & strbatchNumber & "'")
        If oTempRec.RecordCount > 0 Then
            If oTempRec.Fields.Item("U_Z_Imported").Value = "Y" Then
                oApplication.Utilities.Message("Payment documents already created for the selected batch number", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        stPostingdate = oApplication.Utilities.GetEditText(aForm, "38")
        If stPostingdate = "" Then
            oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Items.Item("38").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Return False
        Else
            dtPostingdate = oApplication.Utilities.GetDateTimeValue(stPostingdate)
        End If
        dblDiscountAmount = oApplication.Utilities.getEdittextvalue(aForm, "22")
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oGrid = aForm.Items.Item("20").Specific
        Try
            spath = System.Windows.Forms.Application.StartupPath & "\Log\ImportLog.txt"
            If File.Exists(spath) Then
                File.Delete(spath)
            End If
            WriteErrorlog("Processing Payment Document creation....", spath)
            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Company.StartTransaction()
            Dim oBP As SAPbobsCOM.BusinessPartners
            oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCardName = oGrid.DataTable.GetValue("CardCode", intRow)
                strPayCardCode = strCardName
                If strCardName <> "" Then
                    strCardName = getCardCode(strCardName)
                End If
                dblNetSales = oGrid.DataTable.GetValue(5, intRow)
                If dblNetSales >= 0 Then
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    oDoc.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                    oDoc.CashSum = dblNetSales

                    strtype = "Incoming Payment"
                Else
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                    'oDoc.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
                    oDoc.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                    oDoc.CashSum = dblNetSales * -1
                    'oDoc.CashAccount = "161000"
                    strtype = "Vendor Payment"
                End If
                oDoc.DocDate = dtPostingdate
                oDoc.CashAccount = strCashAccount
                oDoc.CardCode = strPayCardCode ' strCardName
                oDoc.UserFields.Fields.Item("U_BatchNumber").Value = strbatchNumber
                strQuery = "select 'INV',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocTotal  from OINV T0 where T0.CardCode in (" & strCardName & ") and  isnull(U_Z_BatchNumber,'')='" & strbatchNumber & "' and isnull(U_Z_Exported,'N')='Y' and  DocStatus<>'C'   "
                strQuery = strQuery & " Union all select 'CRN',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocTotal  from ORIN T0 where CardCode  in (" & strCardName & ") and  isnull(U_Z_BatchNumber,'')='" & strbatchNumber & "' and isnull(U_Z_Exported,'N')='Y' and  DocStatus<>'C' "
                oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTempRec.DoQuery(strQuery)
                blnLineExist = False
                For intloop As Integer = 0 To oTempRec.RecordCount - 1
                    If intloop > 0 Then
                        oDoc.Invoices.Add()
                        oDoc.Invoices.SetCurrentLine(intloop)
                    End If
                    Dim dblsum As Double
                    If dblNetSales >= 0 Then
                        If oTempRec.Fields.Item(0).Value = "INV" Then
                            oDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                            dblsum = oTempRec.Fields.Item(5).Value
                            oDoc.Invoices.SumApplied = dblsum
                        Else
                            oDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                            dblsum = oTempRec.Fields.Item(5).Value
                            oDoc.Invoices.SumApplied = dblsum * -1
                        End If
                    Else
                        If oTempRec.Fields.Item(0).Value = "INV" Then
                            oDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                            dblsum = oTempRec.Fields.Item(5).Value
                            oDoc.Invoices.SumApplied = dblsum * -1
                        Else
                            oDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                            dblsum = oTempRec.Fields.Item(5).Value
                            oDoc.Invoices.SumApplied = dblsum
                        End If
                    End If
                    oDoc.Invoices.DocEntry = oTempRec.Fields.Item(1).Value
                    oDoc.Invoices.DiscountPercent = 0
                    blnLineExist = True
                    oTempRec.MoveNext()
                Next

                If blnLineExist = True Then
                    If oDoc.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        WriteErrorlog("Payment document creation failed : Customer name " & strCardName & ": Error : " & oApplication.Company.GetLastErrorDescription, spath)
                        If oApplication.Company.InTransaction() Then
                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        openLogFile()
                        Return False
                    Else
                        blnDocumentCreated = True
                        Dim strDocNum As String
                        oApplication.Company.GetNewObjectCode(strDocNum)
                        If strType = "Incoming Payment" Then
                            oTempRec.DoQuery("Select DocNum,TransID from ORCT where DocEntry=" & CInt(strDocNum))
                        Else
                            oTempRec.DoQuery("Select DocNum,TransID from OVPM where DocEntry=" & CInt(strDocNum))
                        End If
                        Dim intTransid As Integer
                        intTransid = oTempRec.Fields.Item(1).Value
                        WriteErrorlog("Payment document created successfully: DocNum Type : " & strType & " DocNum: " & oTempRec.Fields.Item("DocNum").Value, spath)
                        strQuery = "Update JDT1 set Ref2='" & strbatchNumber & "' where TransID=" & oTempRec.Fields.Item(1).Value
                        oTempRec.DoQuery(strQuery)
                        strQuery = "Update OJDT set U_BatchNumber='" & strbatchNumber & "' where TransID=" & intTransid
                        oTempRec.DoQuery(strQuery)
                    End If
                End If
            Next
            If dblDiscountAmount > 0 And blnDocumentCreated = True Then
                Dim oJE As SAPbobsCOM.JournalEntries
                oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJE.DueDate = dtPostingdate
                oJE.TaxDate = dtPostingdate
                oJE.Lines.AccountCode = strBankCreditAcct
                oJE.Lines.Credit = dblDiscountAmount
                oJE.Lines.Reference2 = strbatchNumber
                oJE.Lines.Add()
                oJE.Lines.SetCurrentLine(1)
                oJE.Lines.AccountCode = strBankDebitAcct
                oJE.Lines.Debit = dblDiscountAmount
                oJE.Lines.Reference2 = strbatchNumber
                oJE.UserFields.Fields.Item("U_BatchNumber").Value = strbatchNumber
                If oJE.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction() Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    openLogFile()
                    Return False
                Else
                    Dim stcode As String
                    oApplication.Company.GetNewObjectCode(stcode)
                    WriteErrorlog("Journal Entry created successfully: " & stcode, spath)
                End If
            End If
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            If blnDocumentCreated = True Then
                strQuery = "Update [@Z_Bill_Export] set U_Z_Imported='Y' where U_Z_BatchNumber='" & strbatchNumber & "'"
                oTempRec.DoQuery(strQuery)
                oTempRec.DoQuery("Update ODSC set U_BatchNumber=" & strbatchNumber & " where BankCode='" & strbankCode & "'")
            End If
            oApplication.Utilities.Message("Operation Completed successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            openLogFile()
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Return False
        End Try
    End Function
#End Region

#Region "GetCardCode"
    Private Function getCardCode(ByVal aCode As String) As String
        Dim st As String
        Dim otemp1 As SAPbobsCOM.Recordset
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1.DoQuery("Select CardCode from OCRD where FatherCard='" & aCode & "' and Fathertype='P'")
        st = ""
        st = "'" & aCode & "'"
        For intRow As Integer = 0 To otemp1.RecordCount - 1
            If st <> "" Then
                st = st & ",'" & otemp1.Fields.Item(0).Value & "'"
            Else
                st = "'" & otemp1.Fields.Item(0).Value & "'"
            End If
            otemp1.MoveNext()
        Next
        If st = "" Then
            Return "'" & aCode & "'"
        Else
            Return st
        End If

    End Function
#End Region

#Region "OpenLogFile"
    Private Sub openLogFile()
        Dim x As System.Diagnostics.ProcessStartInfo
        Dim spath As String
        x = New System.Diagnostics.ProcessStartInfo
        x.UseShellExecute = True
        spath = System.Windows.Forms.Application.StartupPath & "\Log\ImportLog.txt"
        If File.Exists(spath) Then
            x.FileName = spath
            System.Diagnostics.Process.Start(x)
            x = Nothing
            Exit Sub
        End If
    End Sub
#End Region

#Region "Validation"
    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        aform.Freeze(True)
        Dim strQuery, strSalesQuery, strReturnsQuery, strFromdate, strTodate, strbank, strCondition, strBPCondition As String
        Dim dtFromdate, dttodate As Date
        oCombobox = aform.Items.Item("9").Specific
        strbank = oCombobox.Selected.Value
        If strbank = "" Then
            oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Exit Function
        Else
            strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
        End If

        strFromdate = oApplication.Utilities.getEdittextvalue(aform, "12")
        strTodate = oApplication.Utilities.getEdittextvalue(aform, "15")
        If strFromdate <> "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
        Else
            oApplication.Utilities.Message("From date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End If
        If strTodate <> "" Then
            dttodate = oApplication.Utilities.GetDateTimeValue(strTodate)
        Else
            oApplication.Utilities.Message("To date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End If
        aform.Freeze(False)
        Return True
    End Function
#End Region

#Region "Change Step"
    Private Sub ChangeStep(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        Dim oStatic As SAPbouiCOM.StaticText
        oStatic = aForm.Items.Item("1").Specific
        Select Case aForm.PaneLevel
            Case "1"
                oStatic.Caption = "Step 1 of 4"
            Case "2"
                oStatic.Caption = "Step 2 of 4"
            Case "3"
                oStatic.Caption = "Step 3 of 4"
            Case "4"
                oStatic.Caption = "Step 3 of 4"
            Case "5"
                oStatic.Caption = "Step 3 of 4"
            Case "6"
                oStatic.Caption = "Step 3 of 4"
            Case "7"
                oStatic.Caption = "Step 4 of 4"
        End Select
        aForm.Freeze(False)
    End Sub
#End Region


#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New FolderBrowserDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.SelectedPath
                        strSelectedFilepath = oDialogBox.SelectedPath
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        If strSelectedFolderPath.EndsWith("\") Then
                            strSelectedFolderPath = strSelectedFilepath.Substring(0, strSelectedFolderPath.Length - 1)
                        End If
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region

#Region "Write into ErrorLog File"
    Private Sub WriteErrorHeader(ByVal apath As String, ByVal strMessage As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = Now.ToLocalTime & "---" & strMessage
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(apath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        aPath = aPath
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToLocalTime & "---" & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BillDiscount Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "33" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strbank, strGirdValue As String
                                    ' oGrid = oForm.Items.Item("1").Specific
                                    oCombobox = oForm.Items.Item("9").Specific
                                    strbank = oCombobox.Selected.Value
                                    If strbank = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getEdittextvalue(oForm, "33")
                                    Dim otemp As SAPbobsCOM.Recordset
                                    otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otemp.DoQuery("Select * from [@Z_Bill_Export] where  U_Z_BankCode='" & strbank & "' and U_Z_BatchNumber='" & strGirdValue & "'")
                                    If otemp.RecordCount > 0 Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "33", otemp.Fields.Item("U_Z_BatchNumber").Value)
                                        oApplication.Utilities.setEdittextvalue(oForm, "32", otemp.Fields.Item("U_Z_BatchNumber").Value)
                                        oApplication.Utilities.setEdittextvalue(oForm, "12", otemp.Fields.Item("U_Z_DateFrom").Value)
                                        oApplication.Utilities.setEdittextvalue(oForm, "15", otemp.Fields.Item("U_Z_DateTo").Value)
                                        strbank = ""
                                    Else
                                        strbank = strbank
                                    End If
                                    If strbank <> "" Then
                                        clsChooseFromList.ItemUID = "33"
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = 0 'pVal.Row
                                        clsChooseFromList.CFLChoice = "[@Z_Bill_Export]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "BatchNumber"
                                        clsChooseFromList.ItemCode = strbank
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

                                    Case "27"
                                        fillopen()
                                        oEditText = oForm.Items.Item("26").Specific
                                        oEditText.String = strSelectedFilepath
                                    Case "30"
                                        fillopen()
                                        oEditText = oForm.Items.Item("29").Specific
                                        oEditText.String = strSelectedFilepath

                                    Case "4"
                                        If oForm.PaneLevel = 2 Then
                                            If Validation(oForm) = False Then
                                                oForm.PaneLevel = 2
                                                oForm.Freeze(False)
                                                Exit Sub
                                            End If
                                            ChangeStep(oForm)
                                        End If
                                        If oForm.PaneLevel = 3 Or oForm.PaneLevel = 4 Or oForm.PaneLevel = 5 Or oForm.PaneLevel = 6 Then
                                            oForm.PaneLevel = 7
                                        Else
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                        End If
                                        If oForm.PaneLevel = 3 Then
                                            GetDetails(oForm)
                                        End If
                                        ChangeStep(oForm)
                                    Case "3"
                                        If oForm.PaneLevel = 3 Or oForm.PaneLevel = 4 Or oForm.PaneLevel = 5 Or oForm.PaneLevel = 6 Then
                                            oForm.PaneLevel = 2
                                        Else
                                            oForm.PaneLevel = oForm.PaneLevel - 1
                                        End If
                                        ChangeStep(oForm)
                                    Case "17"
                                        oForm.PaneLevel = 4
                                        ChangeStep(oForm)
                                    Case "18"
                                        oForm.PaneLevel = 5
                                        ChangeStep(oForm)
                                    Case "19"
                                        oForm.PaneLevel = 6
                                        ChangeStep(oForm)
                                    Case "6"
                                        If postincomingDocuments(oForm) Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Close()
                                        End If
                                    Case "34"
                                        Dim BatchNumber, strFile, strBankFile, strBank As String
                                        strFile = oApplication.Utilities.GetEditText(oForm, "26")
                                        strBankFile = oApplication.Utilities.GetEditText(oForm, "29")
                                        BatchNumber = oApplication.Utilities.getEdittextvalue(oForm, "32")
                                        oCombobox = oForm.Items.Item("9").Specific
                                        strBank = oCombobox.Selected.Value

                                        If strBankFile = "" Then
                                            oApplication.Utilities.Message("Exported Bank file Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                        If strFile = "" Then
                                            oApplication.Utilities.Message("Exported  file Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If

                                        If oApplication.Utilities.GenerateBankDBFFile(BatchNumber, strBankFile, strBank) = True Then
                                            If oApplication.Utilities.generateBillDiscountreport(BatchNumber, strFile, strBank) = True Then
                                            End If
                                        End If

                                        'If strBankFile <> "" Then
                                        '    oApplication.Utilities.GenerateBankDBFFile(BatchNumber, strBankFile, strBank)
                                        'Else
                                        '    oApplication.Utilities.Message("Exported Bank file Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '    Exit Sub
                                        'End If
                                        'If strFile <> "" Then
                                        '    If oApplication.Utilities.generateBillDiscountreport(BatchNumber, strFile, strBank) = True Then
                                        '    End If
                                        'Else
                                        '    oApplication.Utilities.Message("Exported  file Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        '    Exit Sub
                                        'End If

                                        
                                End Select
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_BillDiscount
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
