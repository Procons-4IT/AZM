
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class clsSalesAgingRpt
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
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private strReportviewOption As String

    Private ds As New BillDiscounting       '(dataset)

    Private blnFound As Boolean
    Private strContents(), str As String
    Private intI As Integer = 0
    Private strArr(1) As String
    Private oDRow As DataRow
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Private Sub AddChooseFromList(ByVal aform As SAPbouiCOM.Form)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFLs = oForm.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFL = oCFLs.Item("CFL_2")
        Dim oCond As SAPbouiCOM.Condition
        oCons = oCFL.GetConditions()
        oCond = oCons.Add()

        oCond.Alias = "CardType"
        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCond.CondVal = "C"
        oCFL.SetConditions(oCons)

        oCFL = oCFLs.Item("CFL_3")
        oCons = oCFL.GetConditions()
        oCond = oCons.Add()

        oCond.Alias = "CardType"
        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCond.CondVal = "C"
        oCFL.SetConditions(oCons)

        'oCFL = oCFLs.Add(oCFLCreationParams)


    End Sub



    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_SalesAgingRpt, frm_SalesOdrAging)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            oForm.DataSources.UserDataSources.Add("dtFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("dtTo", SAPbouiCOM.BoDataType.dt_DATE)

            oForm.DataSources.UserDataSources.Add("BPFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("BPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("frmBP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("frmTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("frmPro", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = oForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "dtFrom")
            oEditText = oForm.Items.Item("7").Specific
            oEditText.DataBind.SetBound(True, "", "dtTo")

            AddChooseFromList(oForm)

            oEditText = oForm.Items.Item("31").Specific
            oEditText.DataBind.SetBound(True, "", "BPFrom")
            oEditText.ChooseFromListUID = "CFL_2"
            oEditText.ChooseFromListAlias = "CardCode"

            oEditText = oForm.Items.Item("32").Specific
            oEditText.DataBind.SetBound(True, "", "BPTo")
            oEditText.ChooseFromListUID = "CFL_3"
            oEditText.ChooseFromListAlias = "CardCode"




            oCombobox = oForm.Items.Item("10").Specific
            oCombobox.DataBind.SetBound(True, "", "frmBP")

            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select GroupCode,GroupName from OCRG where GroupType='C'")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("10").DisplayDesc = True
            oCombobox = oForm.Items.Item("18").Specific
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select SlpCode,SlpName from OSLP order by SlpCode")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("18").DisplayDesc = True

            oCombobox = oForm.Items.Item("19").Specific
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select SlpCode,SlpName from OSLP order by SlpCode")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("19").DisplayDesc = True


            oCombobox = oForm.Items.Item("23").Specific
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("SELECT T0.[territryID], T0.[descript] FROM OTER T0")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("23").DisplayDesc = True

            oCombobox = oForm.Items.Item("24").Specific
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("SELECT T0.[territryID], T0.[descript] FROM OTER T0")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("24").DisplayDesc = True


            oCombobox = oForm.Items.Item("12").Specific
            oCombobox.DataBind.SetBound(True, "", "frmTo")
            otemp.DoQuery("select GroupCode,GroupName from OCRG where GroupType='C'")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("12").DisplayDesc = True


            oCombobox = oForm.Items.Item("14").Specific
            oCombobox.DataBind.SetBound(True, "", "frmPro")
            otemp.DoQuery("select GroupCode,GroupName from OCQG order by GroupCode")
            oCombobox.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.Items.Item("14").DisplayDesc = True


            oCombobox = oForm.Items.Item("28").Specific
            oCombobox.ValidValues.Add("1", "Salesman-Area")
            oCombobox.ValidValues.Add("2", "Area-Salesman")
            oCombobox.ValidValues.Add("3", "Channel - Salesman")
            oCombobox.ValidValues.Add("4", "CustomerType - Salesman")
            oCombobox.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue)


            oForm.Items.Item("28").DisplayDesc = True


            oCombobox = oForm.Items.Item("38").Specific
            oCombobox.ValidValues.Add("", "")
            oCombobox.ValidValues.Add("W", "Window")
            oCombobox.ValidValues.Add("P", "PDF")
            oCombobox.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.Items.Item("38").DisplayDesc = True

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try

    End Sub
#Region "Generate Report"
    Private Sub GeneratReport(ByVal aForm As SAPbouiCOM.Form)
        Dim strFromdate, strToDate, strFromBP, strToBP, strPro, strDateCon, strBPCond, strProcon As String
        Dim dtfromdate, dttodate As Date
        Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
        Dim strSlpFrom, strSlpTo, strCust1, strCust2, strcustCondtion, strAreaFrom, strAreaTo, strfrom, strto, strChannal, strCustType, strBranch, strChannalFroms, strChannalTo, strLocalCurrency, strSMNo, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oRecBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBalanceRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strFromdate = oApplication.Utilities.GetEditText(aForm, "16")
        If strFromdate = "" Then
            oApplication.Utilities.Message("Aging Date is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            dtAgingdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
        End If


        oCombobox = aForm.Items.Item("38").Specific
        If oCombobox.Selected.Value = "" Then
            oApplication.Utilities.Message("Report View option is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            strReportviewOption = oCombobox.Selected.Value
        End If

        'dtAgingdate = Now.Date
        strCust1 = oApplication.Utilities.GetEditText(aForm, "31")
        strCust2 = oApplication.Utilities.GetEditText(aForm, "32")

        oCombobox = aForm.Items.Item("10").Specific
        strFromBP = oCombobox.Selected.Value
        strChannalFroms = oCombobox.Selected.Description
        oCombobox = aForm.Items.Item("12").Specific
        strToBP = oCombobox.Selected.Value
        strChannalTo = oCombobox.Selected.Description

        oCombobox = aForm.Items.Item("14").Specific
        strPro = oCombobox.Selected.Value
        strCustType = oCombobox.Selected.Description

        oCombobox = aForm.Items.Item("18").Specific
        strSlpFrom = oCombobox.Selected.Value
        oCombobox = aForm.Items.Item("19").Specific
        strSlpTo = oCombobox.Selected.Value

        oCombobox = aForm.Items.Item("23").Specific
        strAreaFrom = oCombobox.Selected.Value

        oCombobox = aForm.Items.Item("24").Specific
        strAreaTo = oCombobox.Selected.Value
        strFromdate = oApplication.Utilities.GetEditText(aForm, "5")
        strToDate = oApplication.Utilities.GetEditText(aForm, "7")

        If strFromdate <> "" Then
            dtfromdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
            strDateCon = " T0.DocDate>='" & dtfromdate.ToString("yyyy-MM-dd") & "'"
        Else
            strDateCon = "1=1"
        End If

        If strToDate <> "" Then
            dttodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            strDateCon = strDateCon & " and T0.DocDate <='" & dttodate.ToString("yyyy-MM-dd") & "'"
        Else
            strDateCon = strDateCon & " and 1=1"
        End If

        If strCust1 <> "" Then
            'dtfromdate = oApplication.Utilities.GetDateTimeValue(strFromdate)
            strcustCondtion = " and T1.CardCode>='" & strCust1 & "'"
        Else
            strcustCondtion = " and 1=1"
        End If

        If strCust2 <> "" Then
            ' dttodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            strcustCondtion = strcustCondtion & " and T1.CardCode <='" & strCust2 & "'"
        Else
            strcustCondtion = strcustCondtion & " and 1=1"
        End If

        If strFromBP <> "" Then
            strBPCond = " T1.GroupCode >=" & CInt(strFromBP)
            strChannal = strChannalFroms
        Else
            strBPCond = "1=1"
        End If
        If strToBP <> "" Then
            strBPCond = strBPCond & " And T1.GroupCode<=" & CInt(strToBP)
            If strChannal <> "" Then
                strChannal = strChannal & "-" & strChannalTo
            Else
                strChannal = strChannalTo
            End If
        Else
            strBPCond = strBPCond & " and 1=1"
        End If
        If strChannal = "" Then
            strChannal = "All"
        End If
        If strPro <> "" Then
            strProcon = " and T1.QryGroup" & CInt(strPro) & "='Y'"
        Else
            strProcon = " and 1=1"
            strCustType = "All"
            strPro = 0
        End If
        oCombobox = aForm.Items.Item("28").Specific
        Try
            intReportChoice = CInt(oCombobox.Selected.Value)

        Catch ex As Exception
            intReportChoice = 1

        End Try

        Dim strsql, strcondition, strMainQuery, strLines, strSlpCon, strAreaCondtion As String
        strcondition = strDateCon & strBPCond & strProcon
        strLocalCurrency = oApplication.Utilities.GetLocalCurrency()

        If strSlpFrom <> "" Then
            strSlpCon = " and T1.SlpCode >=" & strSlpFrom
        Else
            strSlpCon = " and 1=1"
        End If

        If strSlpTo <> "" Then
            strSlpCon = strSlpCon & " and T1.SlpCode<=" & strSlpTo
        Else
            strSlpCon = strSlpCon & " and 1=1"
        End If

        If strAreaFrom <> "" Then
            strAreaCondtion = " and isnull(T1.[Territory],0) >=" & strAreaFrom
        Else
            strAreaCondtion = " and 1=1"
        End If

        If strAreaTo <> "" Then
            strAreaCondtion = strAreaCondtion & " and isnull(T1.[Territory],0)<=" & strAreaTo
        Else
            strAreaCondtion = strAreaCondtion & " and 1=1"
        End If


        ' strMainQuery = "select T0.CardCode,T0.SlpCode,Count(*) from ORDR T0 inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
        strMainQuery = "Select T1.CardCode,T1.SlpCode, Count(*) from OCRD T1 where CardType='C' and  "
        strMainQuery = strMainQuery & strBPCond & strcustCondtion & strProcon & strSlpCon & strAreaCondtion & " Group by T1.SlpCode,T1.CardCode order by T1.Cardcode"
        Dim oMainRs, oLineRs As SAPbobsCOM.Recordset
        oMainRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oLineRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRs.DoQuery(strMainQuery)
        Dim strBP, strBPCurrency, strAddressSQL, strReportCurrency As String
        Dim intSlpCode As Integer
        strReportCurrency = "L"
        ds.Clear()
        For intRow As Integer = 0 To oMainRs.RecordCount - 1
            strBP = oMainRs.Fields.Item(0).Value
            intSlpCode = oMainRs.Fields.Item(1).Value
            strBPCurrency = oApplication.Utilities.getBPCurrency(strBP)
            oApplication.Utilities.Message("Processing Cardcode : " & strBP, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strAddressSQL = "Select * from OSLP where slpCode=" & intSlpCode
            strBranch = ""
            oRecTemp.DoQuery(strAddressSQL)
            If strReportCurrency = "L" Then
                strBPCurrency = oApplication.Utilities.GetLocalCurrency()
                If strBPCurrency = "##" Then
                    strBPCurrency = oApplication.Company.GetLocalCurrency()
                Else
                    strBPCurrency = strBPCurrency
                End If
            ElseIf strReportCurrency = "B" Then
                If strBPCurrency = "##" Then
                    strBPCurrency = oApplication.Company.GetSystemCurrency()
                Else
                    strBPCurrency = strBPCurrency
                End If
            ElseIf strReportCurrency = "S" Then
                If strBPCurrency = "##" Then
                    strBPCurrency = oApplication.Company.GetSystemCurrency()
                Else
                    strBPCurrency = strBPCurrency
                End If
                strBPCurrency = oApplication.Company.GetSystemCurrency()
            Else
                strBPCurrency = strBPCurrency
            End If
            '  ds.Clear()
            '  ds.Clear()
            Dim strBPChoice As String = "C"
            Dim strSlpname As String
            Dim oTempBPRecSet As SAPbobsCOM.Recordset
            oTempBPRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oRecTemp.RecordCount > 0 Then
                oTempBPRecSet.DoQuery("Select * from OCRD where CardCode='" & strBP & "'")
                oDRow = ds.Tables("SOHeader").NewRow()
                oDRow.Item("SlpNo") = intSlpCode
                strSlpname = oRecTemp.Fields.Item("SlpName").Value
                oDRow.Item("SlpName") = oRecTemp.Fields.Item("SlpName").Value
                oDRow.Item("CardCode") = strBP
                oDRow.Item("CardName") = oTempBPRecSet.Fields.Item("CardName").Value
                oDRow.Item("CreLimit") = oTempBPRecSet.Fields.Item("CreditLine").Value
                oTempBPRecSet.DoQuery(" SELECT T0.[ExtraDays] FROM OCTG T0  INNER JOIN OCRD T1 ON T0.GroupNum = T1.GroupNum WHERE T1.[CardCode] ='" & strBP & "'")
                oDRow.Item("CreDays") = oTempBPRecSet.Fields.Item("ExtraDays").Value
                oDRow.Item("Channal") = strChannal
                oDRow.Item("Type") = strCustType
                oTempBPRecSet.DoQuery("SELECT T1.[GroupName], isnull(T2.[territryID],0), isnull(T2.[descript],''),T1.GroupCode FROM OCRD T0  INNER JOIN OCRG T1 ON T0.GroupCode = T1.GroupCode Left Outer JOIN OTER T2 ON T0.Territory = T2.territryID WHERE T0.[CardCode] ='" & strBP & "'")
                oDRow.Item("CustomerGroup") = oTempBPRecSet.Fields.Item(0).Value
                oDRow.Item("Area") = oTempBPRecSet.Fields.Item(2).Value
                Dim intID, intID1 As Integer
                Dim strValue, strValue1 As String
                Select Case intReportChoice
                    Case 1
                        intID = intSlpCode
                        strValue = strSlpname
                        intID1 = oTempBPRecSet.Fields.Item(1).Value
                        strValue1 = oTempBPRecSet.Fields.Item(2).Value
                    Case 2
                        intID1 = intSlpCode
                        strValue1 = strSlpname
                        intID = oTempBPRecSet.Fields.Item(1).Value
                        strValue = oTempBPRecSet.Fields.Item(2).Value
                    Case 3
                        intID = oTempBPRecSet.Fields.Item(3).Value
                        strValue = oTempBPRecSet.Fields.Item(0).Value
                        intID1 = intSlpCode
                        strValue1 = strSlpname
                    Case 4
                        intID = CInt(strPro)
                        strValue = strCustType
                        intID1 = intSlpCode
                        strValue1 = strSlpname
                End Select
                oDRow.Item("SortID1") = intID
                oDRow.Item("SortIDValue1") = strValue
                oDRow.Item("SortID2") = intID1
                oDRow.Item("SortIDValue2") = strValue1
                oDRow.Item("Choice") = intReportChoice
              


                If strFromdate <> "" Then
                    oDRow.Item("dtFrom") = dtfromdate
                End If
                oDRow.Item("Currency") = strBPCurrency
                If strToDate <> "" Then
                    oDRow.Item("dtto") = dttodate
                Else
                    oDRow.Item("dtto") = Now.Date
                End If
                oDRow.Item("Ageingdate") = dtAgingdate
                Dim st As String
                'st = "select isnull(sum(CheckSum),0) from RCT1 where DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "'" & strPDCDate & ")"
                Dim totRS As SAPbobsCOM.Recordset
                totRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                dtAging = Now.Date
                st = "select isnull(sum(T1.CheckSum),0) from RCT1 T1 where T1.duedate>getdate() and T1.DocNum in (Select T0.DocNum from ORCT T0 where " & strDateCon & " and  T0.CardCode='" & strBP.Replace("'", "''") & "')" ' & strPDCDate & ")"
                Dim totRS1 As SAPbobsCOM.Recordset
                totRS1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                totRS1.DoQuery(st)
                oDRow.Item("PDC") = totRS1.Fields.Item(0).Value
                If strLocalCurrency <> strBPCurrency Then
                    oDRow.Item("W00_30") = GetAgingDetails(strBP, "3=3", "30") '(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(intSlpCode))
                    oDRow.Item("W31_60") = GetAgingDetails(strBP, "3=3", "60") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(intSlpCode))
                    oDRow.Item("W61_90") = GetAgingDetails(strBP, "3=3", "90") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(intSlpCode))
                    oDRow.Item("W91_120") = GetAgingDetails(strBP, "3=3", "120") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(intSlpCode))
                    oDRow.Item("W121_150") = GetAgingDetails(strBP, "3=3", "150") 'GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(intSlpCode))
                    oDRow.Item("W151_180") = GetAgingDetails(strBP, "3=3", "180") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(intSlpCode))
                    oDRow.Item("W180_Above") = 0 ' GetAgingDetails(strBP, "3-3", "30") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(intSlpCode))
                Else
                    oDRow.Item("W00_30") = GetAgingDetails(strBP, "3=3", "30") '(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(intSlpCode))
                    oDRow.Item("W31_60") = GetAgingDetails(strBP, "3=3", "60") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(intSlpCode))
                    oDRow.Item("W61_90") = GetAgingDetails(strBP, "3=3", "90") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(intSlpCode))
                    oDRow.Item("W91_120") = GetAgingDetails(strBP, "3=3", "120") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(intSlpCode))
                    oDRow.Item("W121_150") = GetAgingDetails(strBP, "3=3", "150") 'GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(intSlpCode))
                    oDRow.Item("W151_180") = GetAgingDetails(strBP, "3=3", "180") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(intSlpCode))
                    oDRow.Item("W180_Above") = 0 ' GetAgingDetails(strBP, "3-3", "30") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(intSlpCode))
                End If

                If strLocalCurrency <> strBPCurrency Then
                    oDRow.Item("P00_301") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "30", intSlpCode) ' (strBP, "3=3", "30") '(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(intSlpCode))
                    oDRow.Item("P31_601") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "60", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(intSlpCode))
                    oDRow.Item("P61_901") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "90", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(intSlpCode))
                    oDRow.Item("P91_1201") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "120", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(intSlpCode))
                    oDRow.Item("P121_1501") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "150", intSlpCode) 'GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(intSlpCode))
                    oDRow.Item("P151_1801") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "180", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(intSlpCode))
                    oDRow.Item("P180_Above1") = 0 ' GetAgingDetails(strBP, "3-3", "30") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(intSlpCode))
                Else
                    oDRow.Item("P00_301") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "30", intSlpCode) ' (strBP, "3=3", "30") '(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(intSlpCode))
                    oDRow.Item("P31_601") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "60", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(intSlpCode))
                    oDRow.Item("P61_901") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "90", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(intSlpCode))
                    oDRow.Item("P91_1201") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "120", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(intSlpCode))
                    oDRow.Item("P121_1501") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "150", intSlpCode) 'GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(intSlpCode))
                    oDRow.Item("P151_1801") = GetAgingDetails_IncomingPayment(strBP, "1=1", "2=2", dtAgingdate, "180", intSlpCode) ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(intSlpCode))
                    oDRow.Item("P180_Above1") = 0 ' GetAgingDetails(strBP, "3-3", "30") ' GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(intSlpCode))
                End If
                ds.Tables("SOHeader").Rows.Add(oDRow)
            Else
                oDRow = ds.Tables("SOHeader").NewRow()
                oDRow.Item("CardCode") = strBP
                oDRow.Item("Currency") = strBPCurrency
                oDRow.Item("Title") = "Sales Order Aging"
                oDRow.Item("Ageingdate") = dtAging
                If strfrom <> "" Then
                    oDRow.Item("dtFrom") = dtFrom
                End If
                oDRow.Item("Currency") = strBPCurrency
                If strto <> "" Then
                    oDRow.Item("dtto") = dtTo
                Else
                    oDRow.Item("dtto") = Now.Date
                End If
                dtAging = Now.Date
                If strLocalCurrency <> strBPCurrency Then
                    oDRow.Item("W00_30") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(strBP))
                    oDRow.Item("W31_60") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(strBP))
                    oDRow.Item("W61_90") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(strBP))
                    oDRow.Item("W91_120") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(strBP))
                    oDRow.Item("W121_150") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(strBP))
                    oDRow.Item("W151_180") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(strBP))
                    oDRow.Item("W180_Above") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(strBP))
                Else
                    oDRow.Item("W00_30") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "30", CInt(strBP))
                    oDRow.Item("W31_60") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "60", CInt(strBP))
                    oDRow.Item("W61_90") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "90", CInt(strBP))
                    oDRow.Item("W91_120") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "120", CInt(strBP))
                    oDRow.Item("W121_150") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "150", CInt(strBP))
                    oDRow.Item("W151_180") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "180", CInt(strBP))
                    oDRow.Item("W180_Above") = GetAgingDetails(strBP, strBPCond, strProcon, strDateCon, dtAging, "181", CInt(strBP))
                End If
                Dim st As String
                ds.Tables("SOHeader").Rows.Add(oDRow)
            End If


            oMainRs.MoveNext()
        Next
        addCrystal(ds)
    End Sub


#Region "Add Crystal Report"
    Private Sub addCrystal(ByVal ds1 As DataSet)
        Dim cryRpt As New ReportDocument
        Dim strFilename As String
        Dim strReportFileName As String = "SOAccounts.rpt"
        'strReportviewOption = "P"
        '    strReportFileName = "AcctStatement_old.rpt"
        strFilename = System.Windows.Forms.Application.StartupPath & "\CustomerAgeingreport"
        'strFilename = aBankName & "_BatchNumber_" & aBatchNumber
        strFilename = strFilename & ".pdf"
        oApplication.Utilities.Message("Report Generation processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'strFilename = strFilename & ".pdf"
        If ds1.Tables.Item("SOHeader").Rows.Count > 0 Then
            'If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\" & strReportFileName)
            cryRpt.SetDataSource(ds1)
            oApplication.Utilities.Message("Report Generation processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If strReportviewOption = "W" Then
                Dim objPL As New frmReportViewer
                objPL.iniViewer = AddressOf objPL.GenerateReport
                objPL.rptViewer.ReportSource = cryRpt
                objPL.rptViewer.Refresh()
                objPL.rptViewer.Refresh()
                objPL.WindowState = FormWindowState.Maximized
                objPL.ShowDialog()
                ds1.Clear()
            Else
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New _
                DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                CrDiskFileDestinationOptions.DiskFileName = strFilename
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                ' blnReportGenerationFlag = True
                cryRpt.Close()
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                '  sPath = System.Windows.Forms.Application.StartupPath & "\ImportErrorLog.txt"
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                oApplication.Utilities.Message("Report exported into PDF File", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        Else
            oApplication.Utilities.Message("No data found", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

    End Sub
#End Region

    Private Function GetAgingDetails(ByVal strCardcode As String, ByVal aBPCond As String, ByVal aProCon As String, ByVal aDateCo As String, ByVal dtAgingDate As Date, ByVal strchoice As String, ByVal intSlpCode As Integer) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""
        Select Case strchoice
            Case "30"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and T0.CardCode='" & strCardcode & "' and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 151 and 180"
            Case "181"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        strsql = ""

        Select Case strchoice
            Case "30"
                strsql = "  select  sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0 inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 151 and 180"
            Case "181"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & aProCon & " where T0.SlpCode=" & intSlpCode & " and  T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = dblAmount + oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function

    Private Function GetAgingDetails_SalesOrder(ByVal strCardcode As String, ByVal aBPCond As String, ByVal aDateCo As String, ByVal dtAgingDate As Date, ByVal strchoice As String, ByVal intSlpCode As Integer) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""

        Select Case strchoice
            Case "30"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and    T0.CardCode='" & strCardcode & "' and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and    T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   >150"
            Case "181"
                strsql = "  select sum(T2.OpenQty*T2.Price) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and    T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        strsql = ""

        Select Case strchoice
            Case "30"
                strsql = "  select  sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and    T0.CardCode='" & strCardcode & "' and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0 inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and    T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   >150"
            Case "181"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where  T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and (T3.Seriesname<>'GFC' or T3.Seriesname<>'SR')) and   T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = dblAmount + oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function


    Private Function GetAgingDetails_IncomingPayment(ByVal strCardcode As String, ByVal aBPCond As String, ByVal aDateCo As String, ByVal dtAgingDate As Date, ByVal strchoice As String, ByVal intSlpCode As Integer) As Double
        Dim dblAmount As Double
        Dim strsql, strsql1 As String
        strsql = ""
        strsql1 = ""

        Select Case strchoice
            Case "30"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"


                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"

            Case "60"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   between 31 and 60"

                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 31 and 60 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"

            Case "90"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"
                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 61 and 90 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"
            Case "120"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   between 91 and 120"
                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 91 and 120 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"
            Case "150"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   between 121 and 150"
                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 121 and 150 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"
            Case "180"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')    >150"
                strsql1 = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') >150 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"
            Case "181"
                strsql = "  select sum(T0.[DocTotal]) from ORCT T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.CardCode='" & strCardcode & "' and  DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        strsql = ""
        '  strsql = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate()  and  DocNum in (Select Docentry from ORCT where  DATEDIFF(D,COALESCE(DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30 and CardCode='" & strCardcode.Replace("'", "''") & "' and " & aDateCo & ")"
        If strsql1 <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql1)
            dblAmount = dblAmount - oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function

    Private Function GetAgingDetails(ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""

        If strBPChoice = "C" Or strBPChoice = "P" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred  <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"

                Case "120"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"

                Case "180"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') >150"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
            End Select
        ElseIf strBPChoice = "S" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"


                Case "120"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                Case "180"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  >150"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
            End Select
        End If
        Dim oTemp As SAPbobsCOM.Recordset
        If strsql <> "" Then

            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If

        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select SlpCode from OCRD where CardCode='" & strCardcode & "'")
        dblAmount = dblAmount + GetAgingDetails_SalesOrder(strCardcode, "1=1", "2=2", dtAgingdate, strchoice, oTemp.Fields.Item(0).Value)
        Return dblAmount
    End Function

    Private Function GetFCAgingDetails(ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql, strReportCurrency As String
        strsql = ""
        strReportCurrency = "L"
        If strReportCurrency = "L" Or strReportCurrency = "B" Then
            If strBPChoice = "C" Or strBPChoice = "P" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        'strsql = " Select  isnull(Sum(isnull(FCDebit,0)),0) - isnull(Sum(isnull(FCCredit,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred<>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    Case "120"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                    Case "150"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                    Case "180"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 150"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <> 0 or JDT1.BalFCDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            ElseIf strBPChoice = "S" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <> 0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    Case "120"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                    Case "150"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                    Case "180"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 150 "
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <> 0 or JDT1.BalFCDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            End If
        Else
            If strBPChoice = "C" Or strBPChoice = "P" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"

                    Case "120"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 120"
                    Case "150"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                    Case "180"
                        strsql = " Select  isnull(Sum(isnull([BalFCDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 150"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.[BalScCred] <> 0 or JDT1.[BalScDeb] <> 0) and " & strCondition & " and (DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')) > 180"
                End Select
            ElseIf strBPChoice = "S" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <> 0 or JDT1.BalScDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <> 0 or JDT1.[BalScDeb] <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"

                    Case "120"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <> 0 or JDT1.[BalScDeb] <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                    Case "150"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and (JDT1.BalFCCred <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"

                    Case "180"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull([BalFCDeb],0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') >150"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <> 0 or JDT1.BalFCDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            End If
        End If
        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesOdrAging Then
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

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    GeneratReport(oForm)
                                End If

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
                                        val = oDataTable.GetValue("CardCode", 0)
                                        'val1 = oDataTable.GetValue("DocNum", 0)
                                        If pVal.ItemUID = "31" Or pVal.ItemUID = "32" Then ' And pVal.ColUID = "OCRD.CardCode" Then
                                            oApplication.Utilities.SetEditText(oForm, pVal.ItemUID, val)
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
                ' Case mnu_InvSO
                Case mnu_SalAgeRpt
                    LoadForm()
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
