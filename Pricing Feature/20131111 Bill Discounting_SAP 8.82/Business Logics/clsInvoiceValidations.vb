
Imports System.IO
Public Class clsInvoiceValidations
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
#Region "Methods"
    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_InvoiceValidation, frm_Validations)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oGrid = oForm.Items.Item("1").Specific
        Try
            oForm.Freeze(True)
            oGrid.DataTable.ExecuteQuery("Select CardCode,CardCode,DocNum,CardName,DocDate,DocTotal,DocTotal from ORDR where 1=2")
            AddChooseFromList(oForm)
            FormatGrid(oGrid)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        aGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        aGrid.Columns.Item(0).TitleObject.Caption = "Series"
        oComboColumn = aGrid.Columns.Item(0)
        oComboColumn.ValidValues.Add("", "")
        Dim otemprec As SAPbobsCOM.Recordset
        otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' otemprec.DoQuery("select Series,SeriesName from NNM1 where ObjectCode=17 and SeriesName in ('Invoice','GFC')")
        Dim str As String = ""
        otemprec.DoQuery("select Series,SeriesName from NNM1 where ObjectCode=17")
        For intRow As Integer = 0 To otemprec.RecordCount - 1
            oComboColumn.ValidValues.Add(otemprec.Fields.Item(0).Value, otemprec.Fields.Item(1).Value)
            If otemprec.Fields.Item(1).Value.ToString.ToUpper() = "INVOICE" Then
                str = otemprec.Fields.Item(0).Value
                strInvoiceSeriesNumber = str
            End If
            otemprec.MoveNext()
        Next
        'oComboColumn.ValidValues.Add("Invoice", "Invoice")
        'oComboColumn.ValidValues.Add("GFC", "GFC")


        If str <> "" Then
            oComboColumn.SetSelectedValue(0, oComboColumn.ValidValues.Item(str))
        End If

        oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        aGrid.Columns.Item(1).TitleObject.Caption = "Sales Order Entry"
        oEditTextColumn = aGrid.Columns.Item(1)

        'oEditTextColumn.ChooseFromListUID = "CFL_1"
        'oEditTextColumn.ChooseFromListAlias = "DocEntry"
        'oEditTextColumn.LinkedObjectType = "17"

        oEditTextColumn.Visible = False
        aGrid.Columns.Item(2).TitleObject.Caption = "Sales Order Number"
        aGrid.Columns.Item(2).Editable = True
        aGrid.Columns.Item(3).TitleObject.Caption = "Customer Name"
        aGrid.Columns.Item(3).Editable = False
        aGrid.Columns.Item(4).TitleObject.Caption = "Document Date"
        aGrid.Columns.Item(4).Editable = False
        aGrid.Columns.Item(5).TitleObject.Caption = "Total Amount"
        aGrid.Columns.Item(5).Editable = True
        aGrid.Columns.Item(6).TitleObject.Caption = "DocTotal"
        aGrid.Columns.Item(6).Editable = False
        aGrid.Columns.Item(6).Visible = False
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub


    Private Sub AddChooseFromList(ByVal aform As SAPbouiCOM.Form)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFLs = oForm.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFL = oCFLs.Item("CFL_1")
        Dim oCond As SAPbouiCOM.Condition
        oCons = oCFL.GetConditions()
        oCond = oCons.Add()

        oCond.Alias = "DocStatus"
        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCond.CondVal = "O"
        oCFL.SetConditions(oCons)

        'oCFL = oCFLs.Add(oCFLCreationParams)



    End Sub

#Region "Add Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("1").Specific
        If oGrid.DataTable.GetValue(1, oGrid.DataTable.Rows.Count - 1) <> "" Then
            oGrid.DataTable.Rows.Add()
            oComboColumn = oGrid.Columns.Item(0)
            If strInvoiceSeriesNumber <> "" Then
                oComboColumn.SetSelectedValue(oGrid.DataTable.Rows.Count - 1, oComboColumn.ValidValues.Item(strInvoiceSeriesNumber))
            End If
            oGrid.Columns.Item(2).Click(oGrid.DataTable.Rows.Count - 1)
        End If
    End Sub
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("1").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                oGrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
    End Sub
#End Region

#Region "Validation"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim intDocEntry As Integer
            Dim dblDocTotal, dblLineTotal As Double
            Dim oTemprec As SAPbobsCOM.Recordset
            Dim strLocalCurrency, strDocCurrency As String
            oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aForm.Items.Item("1").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue(1, intRow) <> "" Then
                    intDocEntry = oGrid.DataTable.GetValue(1, intRow)
                    dblLineTotal = oGrid.DataTable.GetValue(5, intRow)
                    ' oTemprec.DoQuery("Select DocTotal,DocTotalFC,DocCur from ORDR where DocEntry=" & intDocEntry)
                    oTemprec.DoQuery("Select Round(DocTotal,3),Round(DocTotalFC,3),DocCur from ORDR where DocEntry=" & intDocEntry)
                    strLocalCurrency = oApplication.Utilities.GetLocalCurrency()
                    strDocCurrency = oTemprec.Fields.Item(2).Value
                    If strLocalCurrency <> strDocCurrency Then
                        dblDocTotal = oTemprec.Fields.Item(1).Value
                    Else
                        dblDocTotal = oTemprec.Fields.Item(0).Value
                    End If
                    Dim strLineTotal As String
                    Dim strtotal As String()
                    Dim intDecimal As Integer

                    'strLineTotal = dblLineTotal
                    'strtotal = strLineTotal.Split(CompanyDecimalSeprator)
                    'strLineTotal = strtotal(1)
                    'If strLineTotal.EndsWith("5") And strLineTotal.Length > 2 Then
                    '    dblLineTotal = Math.Round(dblLineTotal, 3)
                    '    dblLineTotal = dblLineTotal + 0.001
                    '    dblLineTotal = Math.Round(dblLineTotal, 3)
                    'Else
                    '    dblLineTotal = Math.Round(dblLineTotal, 3)
                    'End If

                    'dblLineTotal = Math.Round(dblLineTotal, 3)
                    'strLineTotal = dblDocTotal
                    'strtotal = strLineTotal.Split(CompanyDecimalSeprator)
                    'strLineTotal = strtotal(1)
                    'If strLineTotal.Length > 3 Then
                    '    strLineTotal = strLineTotal.Substring(0, 4)
                    'End If

                    'If strLineTotal.EndsWith("5") And strLineTotal.Length > 2 Then
                    '    dblDocTotal = Math.Round(dblDocTotal, 3)
                    '    dblDocTotal = dblDocTotal + 0.001
                    '    dblDocTotal = Math.Round(dblDocTotal, 3)
                    'Else
                    '    dblDocTotal = Math.Round(dblDocTotal, 3)
                    'End If

                    dblDocTotal = Math.Round(dblDocTotal, 3)
                    dblLineTotal = Math.Round(dblLineTotal, 3)
                    'dblDocTotal = Math.Round(dblDocTotal, 3) - Math.Round(dblLineTotal, 3)

                    dblDocTotal = dblDocTotal - dblLineTotal
                    dblDocTotal = Math.Round(dblDocTotal, 3)

                    Dim intDocNum, strSeries, intSelectedSeries As String
                    Dim otemp As SAPbobsCOM.Recordset
                    intDocNum = oGrid.DataTable.GetValue(2, intRow)
                    otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otemp.DoQuery("Select Series from ORDR where DocStatus<>'C' and  DocNum=" & intDocNum)
                    If otemp.RecordCount > 0 Then
                    Else
                        oApplication.Utilities.Message("Entered Sales order number does not exists . Line no : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item(2).Click(intRow, , 1)
                        Return False
                    End If


                    If (dblDocTotal <> 0) Then
                        oApplication.Utilities.Message("Entered amount does not match with Sales order total : Line no " & intRow + 1 & " : Sales order Number: " & oGrid.DataTable.GetValue(2, intRow), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item(5).Click(intRow, False, 1)
                        Return False
                    End If

                    oComboColumn = oGrid.Columns.Item(0)
                    If oComboColumn.GetSelectedValue(intRow).Value <> "" Then
                        intSelectedSeries = oComboColumn.GetSelectedValue(intRow).Value
                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        intDocNum = oGrid.DataTable.GetValue(2, intRow)
                        otemp.DoQuery("Select Series from ORDR where DocNum=" & intDocNum)
                        If otemp.RecordCount > 0 Then
                            strSeries = otemp.Fields.Item(0).Value
                            If intSelectedSeries <> CInt(strSeries) Then
                                oApplication.Utilities.Message("Selected series type is not matched with the entered Sales order document number: Line no : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oGrid.Columns.Item(2).Click(intRow)
                                Return False
                            End If
                        Else
                            oApplication.Utilities.Message("Entered Sales order number does not exists . Line no : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oGrid.Columns.Item(2).Click(intRow, , 1)
                            Return False
                        End If
                    Else
                        oApplication.Utilities.Message("Series is missng... Line no : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item(0).Click(intRow)
                        Return False
                    End If

                End If
            Next
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Create Invoice Documents"
    Private Function createInvoiceDocuments(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim blnLineExists As Boolean
        Dim intSODocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName, stMessagetext, spath, strChoice As String
        Dim strFileName1 As String
        oGrid = aform.Items.Item("1").Specific
        'If oApplication.Company.InTransaction() Then
        '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        'End If
        spath = System.Windows.Forms.Application.StartupPath & "\Log\ImportLog.txt"
        If File.Exists(spath) Then
            File.Delete(spath)
        End If
        WriteErrorlog("Processing Invoice creation....", spath)
        ' oApplication.Company.StartTransaction()
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue(1, intRow) <> "" Then
                oComboColumn = oGrid.Columns.Item(0)
                strChoice = oComboColumn.GetSelectedValue(intRow).Description
                intSODocEntry = oGrid.DataTable.GetValue(1, intRow)
                objremoteDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                If objremoteDoc.GetByKey(intSODocEntry) Then
                    WriteErrorlog("Processing Sales Order : " & objremoteDoc.DocNum, spath)
                    oApplication.Utilities.Message("Processing Sales Order : " & objremoteDoc.DocNum, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If objremoteDoc.DocumentStatus = SAPbobsCOM.BoStatus.bost_Close Then
                        WriteErrorlog("Sales Order already closed : " & objremoteDoc.DocNum, spath)
                    Else
                        If strChoice = "SR" Or strChoice = "GFC" Then
                            objMainDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                        Else
                            objMainDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        End If
                        objMainDoc.DocNum = objremoteDoc.DocNum
                        objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                        objMainDoc.DocDate = objremoteDoc.DocDate
                        objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                        objMainDoc.TaxDate = objremoteDoc.TaxDate
                        objMainDoc.CardCode = objremoteDoc.CardCode
                        objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                        '  objMainDoc.Comments = objremoteDoc.Comments
                        objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                        objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                        objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                        objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                        objMainDoc.TaxDate = objremoteDoc.TaxDate
                        objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                        objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                        objMainDoc.Address = objremoteDoc.Address
                        objMainDoc.Address2 = objremoteDoc.Address2
                        objMainDoc.AgentCode = objremoteDoc.AgentCode
                        objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                        objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                        objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                        Try
                            objMainDoc.UserFields.Fields.Item("U_Z_GPrice").Value = objremoteDoc.UserFields.Fields.Item("U_Z_GPrice").Value
                        Catch ex As Exception

                        End Try

                        If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                            objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                            objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                        Else
                            objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                        End If

                        objMainDoc.DocType = objremoteDoc.DocType
                        For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                            If objremoteDoc.Expenses.LineTotal > 0 Then
                                If IntExp > 0 Then
                                    objMainDoc.Expenses.Add()
                                    objMainDoc.Expenses.SetCurrentLine(IntExp)
                                End If
                                objremoteDoc.Expenses.SetCurrentLine(IntExp)
                                objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                                objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                                objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                                objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                                objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                                objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                                objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                                objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                                objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                                objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                                objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                            End If
                        Next
                        Dim intCount As Integer = 0
                        For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                            If intCount > 0 Then
                                objMainDoc.Lines.Add()
                            End If
                            objMainDoc.Lines.SetCurrentLine(intCount)
                            objremoteDoc.Lines.SetCurrentLine(intLoop)
                            If objremoteDoc.Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                If objremoteDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service Then
                                    objMainDoc.Lines.BaseType = "17"
                                    objMainDoc.Lines.BaseEntry = objremoteDoc.DocEntry
                                    objMainDoc.Lines.BaseLine = objremoteDoc.Lines.LineNum
                                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                    'objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    ' objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                    ' objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    objMainDoc.Lines.SalesPersonCode = objremoteDoc.Lines.SalesPersonCode
                                    intCount = intCount + 1
                                    blnLineExists = True
                                Else
                                    objMainDoc.Lines.BaseType = "17"
                                    objMainDoc.Lines.BaseEntry = objremoteDoc.DocEntry
                                    objMainDoc.Lines.BaseLine = objremoteDoc.Lines.LineNum
                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    objMainDoc.Lines.SalesPersonCode = objremoteDoc.Lines.SalesPersonCode
                                    intCount = intCount + 1
                                    blnLineExists = True
                                    If 1 = 1 Then
                                        For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                            If intSer > 0 Then
                                                objMainDoc.Lines.SerialNumbers.Add()
                                                objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                            End If
                                            objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                            objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                            objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                            objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                            objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                            objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                            objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                            objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                        Next
                                    End If
                                    If 1 = 1 Then
                                        For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                            If intSer > 0 Then
                                                objMainDoc.Lines.BatchNumbers.Add()
                                                objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                            End If
                                            objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                            objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                            objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                            objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                            objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                            objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                            objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                            objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                            objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                            objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                        Next
                                    End If
                                End If
                            End If
                        Next
                        If blnLineExists = True Then
                            If objMainDoc.Add <> 0 Then
                                WriteErrorlog("Failed to Convert Sales Order docuemnt :" & objremoteDoc.DocNum & ": Error :" & oApplication.Company.GetLastErrorDescription, spath)
                                'If oApplication.Company.InTransaction() Then
                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                'End If
                                openLogFile()
                                Return False

                            Else
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                Dim oInvoice As SAPbobsCOM.Documents
                                Dim intInvoiceDocentry As String
                                If strChoice = "GFC" Or strChoice = "SR" Then
                                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                Else
                                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                End If

                                If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                                    strDocNum = oInvoice.DocNum
                                    intInvoiceDocentry = oInvoice.DocEntry
                                    If strChoice = "GFC" Then
                                        Dim oTemp As SAPbobsCOM.Recordset
                                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oTemp.DoQuery("Update DLN1 set LineStatus='C' where DocEntry=" & intInvoiceDocentry)
                                        oTemp.DoQuery("Update ODLN set DocStatus='C' where DocEntry=" & intInvoiceDocentry)
                                        WriteErrorlog("Delivery Document based on Sales Order : " & objremoteDoc.DocNum & "  Created Successfully. DocNum : " & strDocNum, spath)
                                    Else
                                        WriteErrorlog("Invoice Document based on Sales Order : " & objremoteDoc.DocNum & "  Created Successfully. DocNum : " & strDocNum, spath)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        WriteErrorlog("Invoice Creation Completed...", spath)
        'If oApplication.Company.InTransaction() Then
        '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        'End If
        openLogFile()
        Return True
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
            If pVal.FormTypeEx = frm_Validations Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And pVal.ColUID = "DocNum" And pVal.CharPressed <> 9 Then
                                    Dim strdocnum As String
                                    Dim oTemp As SAPbobsCOM.Recordset
                                    oGrid = oForm.Items.Item("1").Specific
                                    oComboColumn = oGrid.Columns.Item(0)
                                    If oComboColumn.GetSelectedValue(pVal.Row).Value = "" Then
                                        oApplication.Utilities.Message("Series is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And pVal.ColUID = "DocNum" And pVal.CharPressed = 9 Then
                                    Dim strdocnum As String
                                    Dim oTemp As SAPbobsCOM.Recordset
                                    oGrid = oForm.Items.Item("1").Specific
                                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strdocnum = oGrid.DataTable.GetValue(2, pVal.Row)
                                    If strdocnum <> "" Then
                                        oTemp.DoQuery("Select DocEntry,CardName,DocDate,DocTotal from ORDR where DocStatus<>'C' and  DocNum=" & CInt(strdocnum))
                                        If oTemp.RecordCount > 0 Then
                                            oForm.Freeze(True)
                                            oGrid.DataTable.SetValue(1, pVal.Row, oTemp.Fields.Item(0).Value)
                                            oGrid.DataTable.SetValue(3, pVal.Row, oTemp.Fields.Item(1).Value)
                                            oGrid.DataTable.SetValue(4, pVal.Row, oTemp.Fields.Item(2).Value)
                                            oGrid.DataTable.SetValue(6, pVal.Row, oTemp.Fields.Item(3).Value)
                                            oForm.Freeze(False)
                                        Else
                                            oForm.Freeze(True)
                                            oApplication.Utilities.Message("Entered Sales order number does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oGrid.Columns.Item(2).Click(pVal.Row, , 1)
                                            oForm.Freeze(False)
                                        End If
                                    End If

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "4"
                                        AddRow(oForm)
                                    Case "5"
                                        DeleteRow(oForm)
                                    Case "3"
                                        If Validation(oForm) = True Then
                                            If createInvoiceDocuments(oForm) = True Then
                                                oApplication.Utilities.Message("Operation Completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                oForm.Close()
                                            End If
                                        End If
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        val = oDataTable.GetValue("DocEntry", 0)
                                        val1 = oDataTable.GetValue("DocNum", 0)
                                        If pVal.ItemUID = "1" And pVal.ColUID = "OCRD.CardCode" Then
                                            oGrid = oForm.Items.Item("1").Specific
                                            'MsgBox(oDataTable.GetValue("Series", 0))
                                            oGrid.DataTable.SetValue(1, pVal.Row, val)
                                            oGrid.DataTable.SetValue(2, pVal.Row, val1)
                                            oGrid.DataTable.SetValue(3, pVal.Row, oDataTable.GetValue("CardName", 0))
                                            oGrid.DataTable.SetValue(4, pVal.Row, oDataTable.GetValue("DocDate", 0))
                                            oGrid.DataTable.SetValue(6, pVal.Row, oDataTable.GetValue("DocTotal", 0))
                                            oGrid.DataTable.SetValue(0, pVal.Row, oDataTable.GetValue("Series", 0))
                                            If oGrid.DataTable.GetValue(1, oGrid.DataTable.Rows.Count - 1) <> "" Then
                                                oGrid.DataTable.Rows.Add()
                                            End If

                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    'MsgBox(ex.Message)
                                End Try


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
                Case mnu_validations
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
