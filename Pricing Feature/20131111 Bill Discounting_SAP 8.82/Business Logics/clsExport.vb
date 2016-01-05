Imports System.IO
Public Class clsExport
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oOptionButton As SAPbouiCOM.OptionBtn
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
    Dim strFileName As String
    Dim strSelectedFilepath, sPath, strSelectedFolderPath As String
    Dim dtDatatable As SAPbouiCOM.DataTable
    Dim blnErrorflag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
#Region "Methods"
    Private Sub LoadForm()
        Try
        
            oApplication.Utilities.LoadForm(xml_Export, frm_Export)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataSources.UserDataSources.Add("Bank", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("from", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("to", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("Batch", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("FileName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("BankFile", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("DisAmt", SAPbouiCOM.BoDataType.dt_SUM)
            oForm.DataSources.UserDataSources.Add("NetAmt", SAPbouiCOM.BoDataType.dt_SUM)
            oForm.DataSources.UserDataSources.Add("New", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("Old", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("Bat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

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

            oEditText = oForm.Items.Item("33").Specific
            oEditText.DataBind.SetBound(True, "", "NetAmt")

            oOptionButton = oForm.Items.Item("34").Specific
            oOptionButton.DataBind.SetBound(True, "", "New")
            oOptionButton = oForm.Items.Item("35").Specific
            oOptionButton.DataBind.SetBound(True, "", "Old")
            oOptionButton.GroupWith("34")
            oEditText = oForm.Items.Item("37").Specific
            oEditText.DataBind.SetBound(True, "", "Bat")
            'oForm.Items.Item("36").Visible = False
            'oForm.Items.Item("37").Visible = False
            oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
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

#Region "AddtoUDT"
    Private Function AddToUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strItem, strCode, strWhs, strBin, strwhsdesc, strbindesc, strTo, strHeaderRef, strConditionType, strfromdate, strCardCode As String
        Dim strbankCode, strbankName, stFromdate, strTodate, strBatchNumber, strReturnsQuery As String
        Dim dtFromdate, dtTodate As Date
        Dim intbatchNumber As Integer
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage, dblDiscountAmount As Double
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Grid
        oBPGrid = aform.Items.Item("20").Specific
        If oBPGrid.DataTable.GetValue(1, oBPGrid.DataTable.Rows.Count - 1).ToString = "" Then
            oApplication.Utilities.Message("No record found", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End If
        '  oCombobox = aform.Items.Item("6").Specific
        oCombobox = aform.Items.Item("9").Specific

        Dim strQuery, strSalesQuery, strbank, strCondition, strBPCondition As String
        'oCombobox = aform.Items.Item("6").Specific
        oCombobox = aform.Items.Item("9").Specific
        strbank = oCombobox.Selected.Value
        If strbank = "" Then
            oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        Else
            strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
        End If

        If oApplication.Utilities.getEdittextvalue(aform, "22") = "" Then
            If oApplication.SBO_Application.MessageBox("Discount Amount is empty. Do you want to Continue?", , "Yes", "No") = 2 Then
                aform.Freeze(False)
                Return False
            Else
                dblDiscountAmount = 0
            End If
        Else
            dblDiscountAmount = CDbl(oApplication.Utilities.getEdittextvalue(aform, "22"))
            If dblDiscountAmount <= 0 Then
                If oApplication.SBO_Application.MessageBox("Discount Amount should be greater than zero. Do you want to Continue?", , "Yes", "No") = 2 Then
                    aform.Freeze(False)
                    Return False
                Else
                    dblDiscountAmount = dblDiscountAmount
                End If
            End If
        End If

        'If oApplication.Utilities.getEdittextvalue(aform, "20") <> "" Then
        '    oApplication.Utilities.Message("Statement already generated", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    aform.Freeze(False)
        '    Return False
        'End If

        'If oApplication.Utilities.getEdittextvalue(aform, "32") <> "" Then
        '    oApplication.Utilities.Message("Statement already generated", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    aform.Freeze(False)
        '    Return True
        'End If

        strfromdate = oApplication.Utilities.getEdittextvalue(aform, "12")
        strTodate = oApplication.Utilities.getEdittextvalue(aform, "15")
        If strFromdate <> "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
        End If
        If strTodate <> "" Then
            dttodate = oApplication.Utilities.GetDateTimeValue(strTodate)
        End If
        If strFromdate <> "" Then
            strCondition = " T0.DocDate >='" & dtFromdate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = " 1=1"
        End If
        If strTodate <> "" Then
            strCondition = strCondition & " and T0.DocDate<='" & dttodate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = strCondition & " and 1=1"
        End If
        strbankCode = oCombobox.Selected.Value
        strbankName = oCombobox.Selected.Description

        '  strfromdate = oApplication.Utilities.getEdittextvalue(aform, "9")
        ' strTodate = oApplication.Utilities.getEdittextvalue(aform, "11")

        If strfromdate <> "" Then
            dtFromdate = oApplication.Utilities.GetDateTimeValue(strfromdate)
        End If
        If strTodate <> "" Then
            dtTodate = oApplication.Utilities.GetDateTimeValue(strTodate)
        End If
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim intBatchCount As Integer

        '  oTempRec.DoQuery("Select isnull(max(U_Z_BatchNumber),0) from [@Z_Bill_Export] where U_Z_Imported='Y' and U_Z_BankCode='" & strbankCode & "'")
        oTempRec.DoQuery("Select isnull(max(U_Z_BatchNumber),0) from [@Z_Bill_Export] where U_Z_BankCode='" & strbankCode & "'")
        intbatchNumber = oTempRec.Fields.Item(0).Value
        If intbatchNumber = 0 Then
            oTempRec.DoQuery("Select isnull(U_BatchNumber,1) from ODSC where BankCode='" & strbankCode & "'")
            intbatchNumber = CInt(oTempRec.Fields.Item(0).Value)
        Else
            intbatchNumber = intbatchNumber + 1
        End If

        If oApplication.Utilities.GetEditText(aform, "37") <> "" Then
            intbatchNumber = CInt(oApplication.Utilities.GetEditText(aform, "37"))
        End If

        oTempRec.DoQuery("Select * from [@Z_Bill_Export] where U_Z_BankCode='" & strbankCode & "' and U_Z_BatchNumber=" & intbatchNumber)
        If oTempRec.RecordCount <= 0 Then
            ousertable = oApplication.Company.UserTables.Item("Z_Bill_Export")
            strCode = oApplication.Utilities.getMaxCode("@Z_Bill_Export", "Code")
            ousertable.Code = strCode
            ousertable.Name = strCode
            ousertable.UserFields.Fields.Item("U_Z_BankCode").Value = strbankCode
            ousertable.UserFields.Fields.Item("U_Z_BankName").Value = strbankName
            ousertable.UserFields.Fields.Item("U_Z_DateFrom").Value = dtFromdate
            ousertable.UserFields.Fields.Item("U_Z_DateTo").Value = dtTodate
            ousertable.UserFields.Fields.Item("U_Z_BatchNumber").Value = intbatchNumber
            ousertable.UserFields.Fields.Item("U_Z_DiscountAmount").Value = dblDiscountAmount
            ousertable.UserFields.Fields.Item("U_Z_ExportDate").Value = Now.Date
            ousertable.UserFields.Fields.Item("U_Z_Exported").Value = "Y"
            ousertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
            If ousertable.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Else
            ousertable = oApplication.Company.UserTables.Item("Z_Bill_Export")
            strCode = oTempRec.Fields.Item("Code").Value

            If ousertable.GetByKey(strCode) Then
                ousertable.Code = strCode
                ousertable.Name = strCode
                ousertable.UserFields.Fields.Item("U_Z_BankCode").Value = strbankCode
                ousertable.UserFields.Fields.Item("U_Z_BankName").Value = strbankName
                ousertable.UserFields.Fields.Item("U_Z_DateFrom").Value = dtFromdate
                ousertable.UserFields.Fields.Item("U_Z_DateTo").Value = dtTodate
                ousertable.UserFields.Fields.Item("U_Z_BatchNumber").Value = intbatchNumber
                ousertable.UserFields.Fields.Item("U_Z_DiscountAmount").Value = dblDiscountAmount
                ousertable.UserFields.Fields.Item("U_Z_ExportDate").Value = Now.Date
                ousertable.UserFields.Fields.Item("U_Z_Exported").Value = "Y"
                ousertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                If ousertable.Update <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
        End If
        strReturnsQuery = "Update OINV set  U_Z_BatchNumber='" & intbatchNumber & "',U_Z_Exported='Y' where DocStatus<>'C' and  isnull(U_Z_Exported,'N')='N' and  DocEntry in (Select DocEntry from OINV T0 where  " & strCondition & " and T0.CardCode in (" & strBPCondition & "))"
        oTempRec.DoQuery(strReturnsQuery)
        strReturnsQuery = "Update ORIN set U_Z_BatchNumber='" & intbatchNumber & "',U_Z_Exported='Y' where  DocStatus<>'C' and  isnull(U_Z_Exported,'N')='N' and  DocEntry in (Select DocEntry from ORIN T0 where   " & strCondition & " and T0.CardCode in (" & strBPCondition & "))"
        oTempRec.DoQuery(strReturnsQuery)
        '  oApplication.Utilities.setEdittextvalue(aform, "20", intbatchNumber)
        oApplication.Utilities.setEdittextvalue(aform, "32", intbatchNumber)
        'oApplication.Utilities.generateBillDiscountreport(intbatchNumber.ToString)
        'oTempRec.DoQuery("Update ODSC set U_BatchNumber=" & intbatchNumber & " where BankCode='" & strbankCode & "'")
        oApplication.Utilities.Message("Operation completed successfully: Exported Batch Number : " & intbatchNumber, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function

#End Region

#Region "Validation"
    Private Function ValidateOpenbatches(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strbank As String
        oCombobox = aform.Items.Item("9").Specific
        strbank = oCombobox.Selected.Value
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strbank = "" Then
            Return True
        End If
        otemp.DoQuery("Select * from [@Z_Bill_Export] where U_Z_BankCode='" & strbank & "' and isnull(U_Z_Imported,'N')='N'")
        If otemp.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        aform.Freeze(True)
        Dim strQuery, strSalesQuery, strReturnsQuery, strFromdate, strTodate, strbank, strCondition, strBPCondition, strbatch As String
        Dim dtFromdate, dttodate As Date
        oCombobox = aform.Items.Item("9").Specific
        strbank = oCombobox.Selected.Value
        If strbank = "" Then
            oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        Else
            strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
        End If

        If aform.PaneLevel = 1 Then
            strbatch = oApplication.Utilities.GetEditText(aform, "37")
            oOptionButton = aform.Items.Item("35").Specific
            If oOptionButton.Selected = True Then
                If strbatch = "" Then
                    oApplication.Utilities.Message("Select the BatchNumber", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    aform.Items.Item("37").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
            Else
                aform.Freeze(False)
                Return True
            End If
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

        If dtFromdate > dttodate Then
            oApplication.Utilities.Message("From date should be less than To date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End If
        Dim OTemp As SAPbobsCOM.Recordset
        Dim dtLastEnddate As Date
        OTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        OTemp.DoQuery("Select * from [@Z_Bill_Export] where  U_Z_BankCode='" & strbank & "' order by Code desc")
        If OTemp.RecordCount > 0 Then
            oOptionButton = aform.Items.Item("34").Specific
            If oOptionButton.Selected = True Then

                dtLastEnddate = OTemp.Fields.Item("U_Z_DateTo").Value
                If dtFromdate <= dtLastEnddate Then
                    oApplication.Utilities.Message("From date should be greater than last bill generated date : " & dtLastEnddate, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aform.Freeze(False)
                    Return False
                End If
            End If
        Else


        End If
        If strFromdate <> "" Then
            strCondition = " T0.DocDate >='" & dtFromdate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = " 1=1"
        End If
        If strTodate <> "" Then
            strCondition = strCondition & " and T0.DocDate<='" & dttodate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = strCondition & " and 1=1"
        End If


        strQuery = "select T0.CardCode,T0.CardName   from ORDR  T0 where    T0.DocStatus='O'  and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")"
        Dim oValidateRS As SAPbobsCOM.Recordset
        oValidateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oValidateRS.DoQuery(strQuery)
        If oValidateRS.RecordCount > 0 Then
            oApplication.Utilities.Message("Sales orders are pending for the requested period", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End If

        aform.Freeze(False)
        Return True
    End Function
#End Region

#Region "Get Details"
    Private Sub GetDetails(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim strQuery, strSalesQuery, strReturnsQuery, strFromdate, strTodate, strbank, strCondition, strBPCondition As String
            Dim dtFromdate, dttodate As Date
            Dim otemprec As SAPbobsCOM.Recordset
            Dim intBatchNumber As Integer

            oCombobox = aForm.Items.Item("9").Specific
            strbank = oCombobox.Selected.Value
            If strbank = "" Then
                oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            Else
                strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
            End If
            Dim strBatch As String
            strBatch = oApplication.Utilities.getEdittextvalue(aForm, "37")
            If strBatch = "" Then
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery("Select isnull(max(U_Z_BatchNumber),0) from [@Z_Bill_Export] where U_Z_Imported='Y' and U_Z_BankCode='" & strbank & "'")
                intBatchNumber = otemprec.Fields.Item(0).Value
                If intBatchNumber = 0 Then
                    otemprec.DoQuery("Select isnull(U_BatchNumber,1) from ODSC where BankCode='" & strbank & "'")
                    intBatchNumber = CInt(otemprec.Fields.Item(0).Value)
                Else
                    intBatchNumber = intBatchNumber + 1
                End If

            Else
                intBatchNumber = CInt(strBatch)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery("Select isnull(U_Z_Imported,'N'),U_Z_DateFrom,U_Z_DateTo from [@Z_Bill_Export] where U_Z_BatchNumber='" & strBatch & "' and U_Z_BankCode='" & strbank & "'")
                If otemprec.Fields.Item(0).Value = "Y" Then
                    oApplication.Utilities.Message("Selected batch already imported to SAP ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("6").Enabled = False
                Else
                    dtFromdate = otemprec.Fields.Item(1).Value
                    dttodate = otemprec.Fields.Item(2).Value
                    'aForm.Items.Item("6").Enabled = True
                End If
            End If

            

            strFromdate = oApplication.Utilities.getEdittextvalue(aForm, "12")
            strTodate = oApplication.Utilities.getEdittextvalue(aForm, "15")
            If strFromdate <> "" Then
                dtFromdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
            Else
                oApplication.Utilities.Message("From date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            End If
            If strTodate <> "" Then
                dttodate = oApplication.Utilities.GetDateTimeValue(strTodate)
            Else
                oApplication.Utilities.Message("To date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            End If
            If strFromdate <> "" Then
                strCondition = " T0.DocDate >='" & dtFromdate.ToString("yyyy-MM-dd") & "'"
            Else
                strCondition = " 1=1"
            End If
            If strTodate <> "" Then
                strCondition = strCondition & " and T0.DocDate<='" & dttodate.ToString("yyyy-MM-dd") & "'"
            Else
                strCondition = strCondition & " and 1=1"
            End If

          
            'strQuery = "Select isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from "
            'strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
            'strQuery = strQuery & " union  select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode  group by isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName"

            strQuery = "Select isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from "
            strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
            strQuery = strQuery & " union  all select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and " & strCondition & " and  T0.DocStatus='O' and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode  group by isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName"


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
            'isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "'
            'strSalesQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from OINV T0 where  isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N' and  DocStatus='O'  and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  order by CardCode"
            strSalesQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from OINV T0 where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and  DocStatus='O'  and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  order by CardCode"
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

            'isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "'

            'strReturnsQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from ORIN T0 where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N'  and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") order by CardCode"
            strReturnsQuery = "select DocEntry,DocNum,CardCode,CardName,DocTotal  from ORIN T0 where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and " & strCondition & " and  T0.DocStatus='O' and T0.CardCode in (" & strBPCondition & ") order by CardCode"

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

            ' Dim oTempRec As SAPbobsCOM.Recordset
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "'

            'strQuery = "Select Sum(x.INV-x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname "
            'strQuery = strQuery & " union select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N'   and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname "

            strQuery = "Select sum(x.INV)-sum(x.RETU) from "
            strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
            strQuery = strQuery & " union  all select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode "

            'strQuery = "Select sum(x.INV)-sum(x.RETU) from "
            'strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where T0.DocStatus<>'C' and isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
            'strQuery = strQuery & " union all  select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where isnull(U_Z_BatchNumber,'')='" & strBatch & "' and isnull(U_Z_Exported,'N')='Y' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode "


            ' strQuery = "Select Sum(x.INV-x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname "
            'strQuery = strQuery & " union all select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber," & intBatchNumber & ")='" & intBatchNumber & "'  and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname "
            oTempRec.DoQuery(strQuery)
            Dim dblPercentage, dblDisAmt, dblnetSales As Double
            dblnetSales = oTempRec.Fields.Item(0).Value

            oTempRec.DoQuery("Select * from [ODSC] where BankCode='" & strbank & "'")
            If oTempRec.RecordCount > 0 Then
                dblPercentage = oTempRec.Fields.Item("U_DisRate").Value
                If dblPercentage > 0 Then
                    '   dblDisAmt = (dblnetSales / dblPercentage) * 100
                    dblDisAmt = (dblnetSales * dblPercentage) / 100
                Else
                    dblDisAmt = 0
                End If
                oApplication.Utilities.setEdittextvalue(aForm, "22", dblDisAmt)
                dblDisAmt = dblnetSales - dblDisAmt
                oApplication.Utilities.SetEditText(aForm, "33", dblDisAmt)
            End If
            oForm.PaneLevel = 3
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Get Net Total"
    Private Sub getnetAmount(ByVal aform As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        Dim strQuery, strSalesQuery, strReturnsQuery, strFromdate, strTodate, strbank, strCondition, strBPCondition As String
        Dim dtFromdate, dttodate As Date

        oCombobox = aform.Items.Item("9").Specific
        strbank = oCombobox.Selected.Value
        If strbank = "" Then
            oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Exit Sub
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
            Exit Sub
        End If
        If strTodate <> "" Then
            dttodate = oApplication.Utilities.GetDateTimeValue(strTodate)
        Else
            oApplication.Utilities.Message("To date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Exit Sub
        End If
        If strFromdate <> "" Then
            strCondition = " T0.DocDate >='" & dtFromdate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = " 1=1"
        End If
        If strTodate <> "" Then
            strCondition = strCondition & " and T0.DocDate<='" & dttodate.ToString("yyyy-MM-dd") & "'"
        Else
            strCondition = strCondition & " and 1=1"
        End If

        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strQuery = "Select sum(x.INV)-sum(x.RETU) from "
        strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  "
        strQuery = strQuery & " union  all select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where  " & strCondition & " and  T0.DocStatus='O' and T0.CardCode in (" & strBPCondition & ")  group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard  )   x  inner join OCRD T1 on T1.CardCode=x.CardCode "


        'strQuery = "Select Sum(x.INV-x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N' and  T0.DocStatus='O' and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname "
        'strQuery = strQuery & " union select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber,'')='' and isnull(U_Z_Exported,'N')='N'   and " & strCondition & " and T0.CardCode in (" & strBPCondition & ") group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname "
        oTempRec.DoQuery(strQuery)
        Dim dblPercentage, dblDisAmt, dblnetSales As Double
        dblnetSales = oTempRec.Fields.Item(0).Value

       
        dblDisAmt = oApplication.Utilities.GetEditText(aform, "22")
        dblDisAmt = dblnetSales - dblDisAmt
        oApplication.Utilities.SetEditText(aform, "33", dblDisAmt)
    End Sub
#End Region

#Region "Change Step"
    Private Sub ChangeStep(ByVal aForm As SAPbouiCOM.Form)
        aform.Freeze(True)
        Dim oStatic As SAPbouiCOM.StaticText
        oStatic = aform.Items.Item("1").Specific
        Select Case aform.PaneLevel
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
        aform.Freeze(False)
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Export Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "37" And pVal.CharPressed <> 9 Then
                                    oOptionButton = oForm.Items.Item("34").Specific
                                    If oOptionButton.Selected = True Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                If pVal.ItemUID = "9" Then
                                    If ValidateOpenbatches(oForm) = True Then
                                        oForm.Items.Item("35").Enabled = True
                                        oApplication.Utilities.SetEditText(oForm, "37", "")
                                    Else
                                        oApplication.Utilities.SetEditText(oForm, "37", "")
                                        oApplication.Utilities.SetEditText(oForm, "12", "")
                                        oApplication.Utilities.SetEditText(oForm, "15", "")
                                        oForm.Items.Item("34").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        Try
                                            oForm.Items.Item("35").Enabled = False
                                        Catch ex As Exception
                                        End Try

                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "22" And pVal.CharPressed = 9 Then
                                    getnetAmount(oForm)
                                End If

                                If pVal.ItemUID = "37" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strbank, strGirdValue As String
                                    ' oGrid = oForm.Items.Item("1").Specific
                                    oCombobox = oForm.Items.Item("9").Specific
                                    strbank = oCombobox.Selected.Value
                                    If strbank = "" Then
                                        Exit Sub
                                    End If
                                    oOptionButton = oForm.Items.Item("35").Specific
                                    If oOptionButton.Selected = False Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getEdittextvalue(oForm, "37")
                                    Dim otemp As SAPbobsCOM.Recordset
                                    If strGirdValue <> "" Then
                                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        otemp.DoQuery("Select * from [@Z_Bill_Export] where  U_Z_BankCode='" & strbank & "' and isnull(U_Z_Imported,'N')='N' and  U_Z_BatchNumber='" & strGirdValue & "'")
                                        If otemp.RecordCount > 0 Then
                                            oApplication.Utilities.setEdittextvalue(oForm, "33", otemp.Fields.Item("U_Z_BatchNumber").Value)
                                            oApplication.Utilities.setEdittextvalue(oForm, "32", otemp.Fields.Item("U_Z_BatchNumber").Value)
                                            oApplication.Utilities.setEdittextvalue(oForm, "12", otemp.Fields.Item("U_Z_DateFrom").Value)
                                            oApplication.Utilities.setEdittextvalue(oForm, "15", otemp.Fields.Item("U_Z_DateTo").Value)
                                            oForm.Items.Item("12").Enabled = False
                                            oForm.Items.Item("15").Enabled = False
                                            strbank = ""
                                        ElseIf otemp.RecordCount = 0 Then
                                            oApplication.Utilities.Message("No open batches available for this selected bank", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        Else
                                            strbank = strbank
                                        End If
                                    Else
                                        strbank = strbank
                                    End If

                                    If strbank <> "" Then
                                        clsChooseFromList.ItemUID = "37"
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = 0 'pVal.Row
                                        clsChooseFromList.CFLChoice = "[@Z_Bill_Export]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "BatchNumber"
                                        clsChooseFromList.ItemCode = strbank
                                        clsChooseFromList.Documentchoice = "Generation" ' oApplication.Utilities.GetDocType(oForm)
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
                                    Case "34"
                                        oForm.Items.Item("36").Visible = True
                                        Try
                                            oApplication.Utilities.SetEditText(oForm, "37", "")
                                            oForm.Items.Item("37").Enabled = False
                                        Catch ex As Exception
                                        End Try

                                        oForm.Items.Item("12").Enabled = True
                                        oForm.Items.Item("15").Enabled = True
                                    Case "35"
                                        oForm.Items.Item("36").Visible = True
                                        Try
                                            oForm.Items.Item("37").Enabled = True
                                            oForm.Items.Item("12").Enabled = False
                                            oForm.Items.Item("15").Enabled = False

                                        Catch ex As Exception

                                        End Try
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

                                        If oForm.PaneLevel = 1 Then
                                            If Validation(oForm) = False Then
                                                oForm.PaneLevel = 1
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
                                        Dim BatchNumber, strFile, strBankFile, strBank As String
                                        strFile = oApplication.Utilities.GetEditText(oForm, "26")
                                        strBankFile = oApplication.Utilities.GetEditText(oForm, "29")
                                        ' BatchNumber = oApplication.Utilities.getEdittextvalue(oForm, "32")
                                        oCombobox = oForm.Items.Item("9").Specific
                                        strBank = oCombobox.Selected.Value
                                        If strFile = "" Then
                                            oApplication.Utilities.Message("Exported File Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                        If strBankFile = "" Then
                                            oApplication.Utilities.Message("Exported Bank file Folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                        If AddToUDT(oForm) = False Then
                                        Else
                                            BatchNumber = oApplication.Utilities.getEdittextvalue(oForm, "32")
                                            If BatchNumber = "" Then
                                                oApplication.Utilities.Message("Statement process not completed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                Exit Sub
                                            Else
                                                If oApplication.Utilities.GenerateBankDBFFile(BatchNumber, strBankFile, strBank) = True Then
                                                    If oApplication.Utilities.generateBillDiscountreport(BatchNumber, strFile, strBank) = True Then
                                                        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                        oForm.Close()
                                                    End If
                                                End If
                                            End If
                                        End If

                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "16" Then
                                            If (oCFL.UniqueID = "6" Or oCFL.UniqueID = "7" Or oCFL.UniqueID = "8" Or oCFL.UniqueID = "9" Or oCFL.UniqueID = "10" Or oCFL.UniqueID = "11") Then
                                                val = oDataTable.GetValue("DocNum", 0)
                                            Else
                                                val = oDataTable.GetValue(0, 0)
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "16", val)
                                        ElseIf pVal.ItemUID = "18" Then
                                            If (oCFL.UniqueID = "6" Or oCFL.UniqueID = "7" Or oCFL.UniqueID = "8" Or oCFL.UniqueID = "9" Or oCFL.UniqueID = "10" Or oCFL.UniqueID = "11") Then
                                                val = oDataTable.GetValue("DocNum", 0)
                                            Else
                                                val = oDataTable.GetValue(0, 0)
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
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
                Case mnu_Export
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
