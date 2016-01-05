Imports System
Imports System.Collections
Imports System.ComponentModel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Collections.Generic



Public Class clsBP

#Region "Declaration"
    Private blnCFL As Boolean
    Dim objSBOAPI As ClsSBO
    Dim objUtility As clsUtilities
    Dim objForm As SAPbouiCOM.Form
    Dim oRec As SAPbobsCOM.Recordset
    Dim strsql, strBPCode2 As String
    Public dblsum, dblcreditlimit, dblPaidtodate As Double
    Dim oFolder As SAPbouiCOM.Folder
    Dim oEdit, oEditBP As SAPbouiCOM.EditText
    Dim oStatic As SAPbouiCOM.StaticText
    Dim oItem, oItem1 As SAPbouiCOM.Item
    Dim oOption As SAPbouiCOM.OptionBtn
    Dim oButton As SAPbouiCOM.Button
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim oGrid, oGr As SAPbouiCOM.Grid
    Dim objDataTable As SAPbouiCOM.DataTable
    Dim oColumn As SAPbouiCOM.Column
    Dim oColumns As SAPbouiCOM.Columns
    Dim oEditTxtCol, oEditTxtColItDesc, oEditTxtColIsbn As SAPbouiCOM.EditTextColumn
    Dim dtTemp As SAPbouiCOM.DataTable
    Dim boollItFlag As Boolean = False
    Dim strReportviewOption, strBPCondition, strSOBPCondition, strJournalDatecondition, strProjectCondition, strReportCurrency, strReportFilterdate, strDisplayOption, strAgeingCondition As String
    Dim strReconcilationCondition As String

    Dim dtAgingdate As Date
    Dim strBPChoice As String = ""
    Dim strBPCurrency, strLocalCurrency As String
    Dim strAcctType As String

    'Private rptAccountReport As New CrystalReport2
    Private rptaccountreport As New AcctStatement
    Dim cryRpt As New ReportDocument
    Private ds As New AccountBalance       '(dataset)

    Private blnFound As Boolean
    Private strContents(), str As String
    Private oReader As StreamReader
    Private intI As Integer = 0
    Private strArr(1) As String
    Private oDRow As DataRow
    Private blnProject As Boolean = False


#End Region

#Region "Methods"

    Public Sub New(ByVal objSBO As ClsSBO)
        objSBOAPI = objSBO
        objUtility = New clsUtilities(objSBOAPI)
    End Sub

#Region "LoadForm"
    Private Sub LoadForm()
        objForm = objSBOAPI.LoadForm(System.Windows.Forms.Application.StartupPath & "\xml\StatementofAccount.xml", "DABT_701")
        objForm.DataSources.UserDataSources.Add("FromBp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        objForm.DataSources.UserDataSources.Add("ToBp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        objForm.DataSources.UserDataSources.Add("FromDate", SAPbouiCOM.BoDataType.dt_DATE)
        objForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE)
        objForm.DataSources.UserDataSources.Add("Agedt", SAPbouiCOM.BoDataType.dt_DATE)
        objForm.DataSources.UserDataSources.Add("frmSlp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        objForm.DataSources.UserDataSources.Add("toSlp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        objForm.DataSources.UserDataSources.Add("Check", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        objForm.DataSources.UserDataSources.Add("intChoice", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
        objForm.DataSources.UserDataSources.Add("CustType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        AddChooseFromList()
        DataBind(objForm)
        objForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

    End Sub
#End Region

#Region "DataBind"
    Private Sub DataBind(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oCheck As SAPbouiCOM.CheckBox
            aForm.Freeze(True)
            oEdit = aForm.Items.Item("13").Specific
            oEdit.DataBind.SetBound(True, "", "FromBp")
            oEdit.ChooseFromListUID = "CFL1"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = aForm.Items.Item("15").Specific
            oEdit.DataBind.SetBound(True, "", "ToBp")
            oEdit.ChooseFromListUID = "CFL2"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = aForm.Items.Item("7").Specific
            oEdit.DataBind.SetBound(True, "", "FromDate")
            oEdit = aForm.Items.Item("9").Specific
            oEdit.DataBind.SetBound(True, "", "ToDate")
            oEdit = aForm.Items.Item("22").Specific
            oEdit.DataBind.SetBound(True, "", "Agedt")
            oEdit = aForm.Items.Item("29").Specific
            oEdit.DataBind.SetBound(True, "", "frmSlp")
            oEdit.ChooseFromListUID = "CFL7"
            oEdit.ChooseFromListAlias = "SlpName"
            oEdit = aForm.Items.Item("32").Specific
            oEdit.DataBind.SetBound(True, "", "toSlp")
            oEdit.ChooseFromListUID = "CFL8"
            oEdit.ChooseFromListAlias = "SlpName"
            'oCheck = aForm.Items.Item("43").Specific
            'oCheck.DataBind.SetBound(True, "", "Check")
            LoadComboBox(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            objUtility.ShowErrorMessage(ex.Message)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Load Group Combo"
    Private Sub LoadGroupCombo(ByVal aForm As SAPbouiCOM.Form)
        Dim oCombo, oCombo1 As SAPbouiCOM.ComboBox
        Dim strSQL As String
        Dim oTemprec As SAPbobsCOM.Recordset
        oCombo = aForm.Items.Item("17").Specific
        oCombo1 = aForm.Items.Item("18").Specific
        For intLoop As Integer = oCombo1.ValidValues.Count - 1 To 0 Step -1
            oCombo1.ValidValues.Remove(intLoop, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oTemprec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'SELECT groupCode,GroupName,GroupType,locked  FROM OCRG T0
        strSQL = ""
        Select Case oCombo.Selected.Value
            Case "C"
                strSQL = "SELECT groupCode,GroupName,GroupType,locked  FROM OCRG T0 where grouptype='C' order by GroupCode"
            Case "S"
                strSQL = "SELECT groupCode,GroupName,GroupType,locked  FROM OCRG T0 where grouptype='S' order by GroupCode"
        End Select
        oCombo1.ValidValues.Add("", "")
        If strSQL <> "" Then
            oTemprec.DoQuery(strSQL)
            For intRow As Integer = 0 To oTemprec.RecordCount - 1
                oCombo1.ValidValues.Add(oTemprec.Fields.Item(0).Value, oTemprec.Fields.Item(1).Value)
                oTemprec.MoveNext()
            Next
        End If
        oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("18").DisplayDesc = True
    End Sub
#End Region

#Region "Load Combo Box Values"
    Private Sub LoadComboBox(ByVal aForm As SAPbouiCOM.Form)
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = aForm.Items.Item("5").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("P", "Posting Date")
        oCombo.ValidValues.Add("D", "Document Date")
        oCombo.ValidValues.Add("DU", "Due Date")
        oCombo.Select("DU", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("5").DisplayDesc = True

        oCombo = aForm.Items.Item("11").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("C", "Customer")
        'oCombo.ValidValues.Add("S", "Supplier")
        'oCombo.ValidValues.Add("P", "Project")
        oCombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("11").DisplayDesc = True

        oCombo = aForm.Items.Item("17").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("C", "Customer")
        '  oCombo.ValidValues.Add("S", "Supplier")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("17").DisplayDesc = True

        oCombo = aForm.Items.Item("25").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("L", "Local Currency")
        'oCombo.ValidValues.Add("S", "System Currency")
        'oCombo.ValidValues.Add("B", "BP Currency")
        oCombo.Select("L", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("25").DisplayDesc = True

        oCombo = aForm.Items.Item("20").Specific
        oCombo.ValidValues.Add("All", "All Posting")
        oCombo.ValidValues.Add("U", "Unreconciled Externally")
        oCombo.ValidValues.Add("R", "Reconciled Externally")
        oCombo.ValidValues.Add("N", "Not Fully Reconciled ")
        oCombo.ValidValues.Add("F", "Fully Unreconciled")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("20").DisplayDesc = True

        oCombo = aForm.Items.Item("27").Specific
        'oCombo.ValidValues.Add("P", "PDF")
        oCombo.ValidValues.Add("W", "Window")
        oCombo.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("27").DisplayDesc = True

        oCombo = aForm.Items.Item("34").Specific
        oCombo.DataBind.SetBound(True, "", "intChoice")
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("0", "Regular")
        oCombo.ValidValues.Add("2", "With PDC")
        oCombo.ValidValues.Add("1", "Month End Batch")
        oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("34").DisplayDesc = True

        oCombo = aForm.Items.Item("35").Specific
        oCombo.DataBind.SetBound(True, "", "CustType")
        oCombo.ValidValues.Add("All", "All")
        oCombo.ValidValues.Add("Main", "Main")
        oCombo.ValidValues.Add("Branch", "Branch")
        oCombo.Select("All", SAPbouiCOM.BoSearchKey.psk_ByValue)
        aForm.Items.Item("35").DisplayDesc = True



    End Sub
#End Region

#Region "Change CFL"
    Private Sub changeBPCFL(ByVal aForm As SAPbouiCOM.Form)
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim strBPChoice As String
        oCombo = aForm.Items.Item("11").Specific
        strBPChoice = oCombo.Selected.Value
        If strBPChoice = "C" Then
            oEdit = aForm.Items.Item("13").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL1"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = aForm.Items.Item("15").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL2"
            oEdit.ChooseFromListAlias = "CardCode"
        ElseIf strBPChoice = "S" Then
            oEdit = aForm.Items.Item("13").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL3"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = aForm.Items.Item("15").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL4"
            oEdit.ChooseFromListAlias = "CardCode"
        ElseIf strBPChoice = "P" Then
            oEdit = aForm.Items.Item("13").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL5"
            oEdit.ChooseFromListAlias = "prjcode"
            oEdit = aForm.Items.Item("15").Specific
            oEdit.String = ""
            oEdit.ChooseFromListUID = "CFL6"
            oEdit.ChooseFromListAlias = "prjcode"
        Else

        End If

    End Sub
#End Region

#Region "AddChooseFromList"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddChooseFromList
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV 4
    'Created Date       : 
    'Last Modified By   : 
    'Purpose            : To Add The ChooseFromList
    '*****************************************************************
    Private Sub AddChooseFromList()
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = objSBOAPI.SBO_Appln.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()

            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL3"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL4"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.ObjectType = "63"
            oCFLCreationParams.UniqueID = "CFL5"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "63"
            oCFLCreationParams.UniqueID = "CFL6"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "53"
            oCFLCreationParams.UniqueID = "CFL7"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "53"
            oCFLCreationParams.UniqueID = "CFL8"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
#End Region


#Region "Get BP Details against Project"
    Private Function GetBPDetails(ByVal strProject As String) As String
        Dim strBPDetails As String
        Dim oTempProjectRec As SAPbobsCOM.Recordset
        oTempProjectRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempProjectRec.DoQuery("Select " & strShortname & ",count(*) from JDT1 where Project='" & strProject & "' group by shortname")
        strBPDetails = ""
        For intProjectRow As Integer = 0 To oTempProjectRec.RecordCount - 1
            If strBPDetails <> "" Then
                strBPDetails = strBPDetails & ",'" & oTempProjectRec.Fields.Item(0).Value & "'"
            Else
                strBPDetails = "'" & oTempProjectRec.Fields.Item(0).Value & "'"
            End If
            oTempProjectRec.MoveNext()
        Next
        Return strBPDetails
    End Function
#End Region

#Region "Statement of Account"
    Private Function StatementofAccount(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
        Dim strfrom, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer

        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double

        Try

            oRecBP = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBalanceRs = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            dtFrom = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("FromDate").Value)
            strfrom = objSBOAPI.GetSBODateString(dtFrom)

            dtTo = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("ToDate").Value)
            strto = objSBOAPI.GetSBODateString(dtTo)
            If strto = "" Then
                dtTo = Now.Date
            End If
            strLocalCurrency = objUtility.GetLocalCurrency()

            strFromBP = objSBOAPI.GetEditText(aForm, "13")
            strToBP = objSBOAPI.GetEditText(aForm, "15")
            If strfrom <> "" Then
                dtFrom = objSBOAPI.GetDateTimeValue(strfrom)
            End If
            If strto <> "" Then
                dtTo = objSBOAPI.GetDateTimeValue(strto)
            End If

            Dim strCond, strCond1, strBPSQL, strAddressSQL, strBP, strBalanceSQL, strPDCCustomer, strPDCDate As String
            If strFromBP <> "" And strToBP <> "" Then
                strCond = " and  " & strShortname & " BETWEEN '" & strFromBP.Replace("'", "''") & "' AND '" & strToBP.Replace("'", "''") & "'"
                strPDCCustomer = " and  " & strShortname & " BETWEEN '" & strFromBP.Replace("'", "''") & "' AND '" & strToBP.Replace("'", "''") & "'"

            ElseIf strFromBP <> "" And strToBP = "" Then
                strCond = " and  " & strShortname & " >= '" & strFromBP.Replace("'", "''") & "' "
                strPDCCustomer = " and  " & strShortname & " >= '" & strFromBP.Replace("'", "''") & "' "
            ElseIf strFromBP = "" And strToBP <> "" Then
                strCond = " and  " & strShortname & " <= '" & strToBP.Replace("'", "''") & "' "
                strPDCCustomer = " and  " & strShortname & " <= '" & strToBP.Replace("'", "''") & "' "
            Else
                strCond = " and 1=1"
                strPDCCustomer = " and 1=1"
            End If

            Dim strdate As String = ""

            Select Case strReportFilterdate
                Case "RefDate"
                    strdate = "DocDate"
                    ' strReportFilterdate = "RefDate"
                Case "taxDate"
                    strdate = "TaxDate"
                    strReportFilterdate = "taxDate"
                Case "DueDate"
                    strdate = "DocDueDate"
                    strReportFilterdate = "DueDate"
            End Select

            If strfrom <> "" And strto <> "" Then
                strCond1 = " and REFDATE BETWEEN '" & dtFrom.ToString("yyyy-MM-dd") & "' AND '" & dtTo.ToString("yyyy-MM-dd") & "'"
                strPDCDate = " and " & strdate & " BETWEEN '" & dtFrom.ToString("yyyy-MM-dd") & "' AND '" & dtTo.ToString("yyyy-MM-dd") & "'"
            ElseIf strfrom <> "" And strto = "" Then
                strCond1 = " and REFDATE >= '" & dtFrom.ToString("yyyy-MM-dd") & "' "
                strPDCDate = " and " & strdate & " >= '" & dtFrom.ToString("yyyy-MM-dd") & "' "
            ElseIf strfrom = "" And strto <> "" Then
                strCond1 = " and REFDATE <= '" & dtTo.ToString("yyyy-MM-dd") & "' "
                strPDCDate = " and " & strdate & " <= '" & dtTo.ToString("yyyy-MM-dd") & "' "
            Else
                strCond1 = " and 1=1"
                strPDCDate = " and 1=1"
            End If

            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = aForm.Items.Item("34").Specific
            intReportChoice = CInt(oCombo.Selected.Value)

            Dim strCustType As String = ""
            oCombo = aForm.Items.Item("35").Specific
            strCustType = oCombo.Selected.Value


            '   strBPSQL = "Select CardCode,CardName from ocrd where " & strBPCondition & " and  cardcode in ( select shortname from jdt1 where 1=1 and  " & strJournalDatecondition & " and " & strProjectCondition & " ) order by cardcode"
            ' strBPSQL = "Select CardCode,CardName,slpCode from ocrd where " & strBPCondition & " and  cardcode in ( select  " & strShortname & " from jdt1 where 1=1 and  " & strProjectCondition & " ) order by cardcode"
            strBPSQL = "Select CardCode,CardName,slpCode from ocrd where " & strBPCondition & " order by cardcode"
            oRecBP.DoQuery(strBPSQL)
            ds.Clear()
            oRecBP.DoQuery(strBPSQL)
            ds.Clear()
            Dim oTempBPRecSet As SAPbobsCOM.Recordset
            oTempBPRecSet = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oTestRS1 As SAPbobsCOM.Recordset
            oTestRS1 = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For intRow As Integer = 0 To oRecBP.RecordCount - 1
                strBP = oRecBP.Fields.Item(0).Value

                If strCustType = "All" Then
                    strBP = strBP
                ElseIf strCustType = "Main" Then
                    oTestRS1.DoQuery("Select * from OCRD where isnull(FatherCard,'')='' and CardCode='" & strBP & "'")
                    If oTestRS1.RecordCount > 0 Then
                        strBP = strBP
                    Else
                        strBP = ""
                    End If
                ElseIf strCustType = "Branch" Then
                    oTestRS1.DoQuery("Select * from OCRD where FatherCard='" & strBP & "' and FatherType='P'")
                    If oTestRS1.RecordCount > 0 Then
                        strBP = ""
                    Else
                        strBP = strBP
                    End If
                Else
                    strBP = ""
                End If

                strBPCurrency = objUtility.getBPCurrency(strBP)
                strSlpCode = oRecBP.Fields.Item(2).Value
                objUtility.ShowWarningMessage("Processing CardCode : " & strBP)
                'strAddressSQL = "Select CardCode, CardName, Block, City, BillToDef, ZipCode, Address, County, Phone1, Fax, CntctPrsn, Notes From OCRD, OCRY where OCRD.Country = OCRY.Code and OCRD.CardCode='" & strBP.Replace("'", "''") & "'"
                strAddressSQL = "Select CardCode, CardName, Block, City, BillToDef, ZipCode, Address,  Phone1, Fax, CntctPrsn, Notes,Cellular,FatherType,FatherCard From OCRD where OCRD.CardCode='" & strBP.Replace("'", "''") & "'"
                oRecTemp.DoQuery(strAddressSQL)
                If oRecTemp.RecordCount > 0 Then
                    If oRecTemp.Fields.Item("FatherType").Value = "P" Then
                        strBranch = oRecTemp.Fields.Item("FatherCard").Value
                        Dim oBPtest As SAPbobsCOM.Recordset
                        oBPtest = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oBPtest.DoQuery("Select isnull(Cardname,'') from OCRD where CardCode='" & strBranch & "'")
                        strBranch = oBPtest.Fields.Item(0).Value
                    Else
                        strBranch = ""
                    End If
                    If strReportCurrency = "L" Then
                        strBPCurrency = objUtility.GetLocalCurrency()
                        If strBPCurrency = "##" Then
                            strBPCurrency = objUtility.GetLocalCurrency()
                        Else
                            strBPCurrency = strBPCurrency
                        End If
                    ElseIf strReportCurrency = "B" Then
                        If strBPCurrency = "##" Then
                            strBPCurrency = objUtility.GetSystemCurrency()
                        Else
                            strBPCurrency = strBPCurrency
                        End If
                    ElseIf strReportCurrency = "S" Then
                        If strBPCurrency = "##" Then
                            strBPCurrency = objUtility.GetSystemCurrency()
                        Else
                            strBPCurrency = strBPCurrency
                        End If
                        strBPCurrency = objUtility.GetSystemCurrency()
                    Else
                        strBPCurrency = strBPCurrency
                    End If
                    If oRecTemp.RecordCount > 0 Then
                        oDRow = ds.Tables("Header").NewRow()
                        If strBPChoice = "C" Then
                            oDRow.Item("BPType") = "Y"
                        Else
                            oDRow.Item("BPType") = "N"
                        End If
                        oDRow.Item("CardCode") = strBP
                        oDRow.Item("Choice") = intReportChoice
                        oDRow.Item("SMNo") = CInt(strSlpCode)
                        Dim oSLPRS As SAPbobsCOM.Recordset
                        oSLPRS = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oSLPRS.DoQuery("Select isnull(slpName,'') from OSLP where SlpCode=" & CInt(strSlpCode))
                        oDRow.Item("SMName") = oSLPRS.Fields.Item(0).Value
                        oDRow.Item("CardName") = oRecTemp.Fields.Item(1).Value
                        oDRow.Item("Block") = oRecTemp.Fields.Item(2).Value
                        oDRow.Item("City") = oRecTemp.Fields.Item(3).Value
                        oDRow.Item("BilltoDef") = oRecTemp.Fields.Item(4).Value
                        oDRow.Item("Zipcode") = oRecTemp.Fields.Item(5).Value
                        oDRow.Item("Address") = oRecTemp.Fields.Item(6).Value
                        Dim strTempSql As String
                        strTempSql = "Select CardCode, CardName, Block, City, BillToDef, ZipCode, Address, County, Phone1, Fax, CntctPrsn, Notes,Cellular From OCRD, OCRY where OCRD.Country = OCRY.Code and OCRD.CardCode='" & strBP.Replace("'", "''") & "'"
                        oTempBPRecSet.DoQuery(strTempSql)
                        If oTempBPRecSet.RecordCount > 0 Then
                            oDRow.Item("County") = oTempBPRecSet.Fields.Item(7).Value
                        Else
                            oDRow.Item("County") = ""
                        End If
                        oDRow.Item("Phone1") = oRecTemp.Fields.Item(7).Value
                        oDRow.Item("Fax") = oRecTemp.Fields.Item(8).Value
                        oDRow.Item("CntctPrsn") = oRecTemp.Fields.Item(9).Value
                        oDRow.Item("Notes") = oRecTemp.Fields.Item(10).Value
                        oDRow.Item("Mobile") = oRecTemp.Fields.Item(11).Value
                        If strfrom <> "" Then
                            oDRow.Item("dtFrom") = dtFrom
                        End If
                        oDRow.Item("Currency") = strBPCurrency
                        If strto <> "" Then
                            oDRow.Item("dtto") = dtTo
                        Else
                            oDRow.Item("dtto") = Now.Date
                        End If
                        oDRow.Item("Ageingdate") = dtAgingdate
                        oDRow.Item("Title") = strDisplayOption
                        Dim st As String
                        st = "select isnull(sum(CheckSum),0) from RCT1 where duedate>getdate() and  DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "'" & strPDCDate & ")"
                        Dim totRS As SAPbobsCOM.Recordset
                        totRS = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        totRS.DoQuery(st)
                        If intReportChoice = 2 Then
                            oDRow.Item("TotPDC") = totRS.Fields.Item(0).Value
                        Else
                            oDRow.Item("TotPDC") = 0
                        End If
                        st = "Select U_Name from OUSR where User_code='" & objSBOAPI.oCompany.UserName & "'"
                        totRS.DoQuery(st)
                        oDRow.Item("Fax") = totRS.Fields.Item(0).Value

                        If strAcctType.ToUpper() = "ALL" Then
                            If strLocalCurrency <> strBPCurrency Then
                                oDRow.Item("W00_30") = GetFCAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetFCAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetFCAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetFCAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetFCAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetFCAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetFCAgingDetails(strBP, strAgeingCondition, "181")
                            Else
                                oDRow.Item("W00_30") = GetAgingDetails_Full(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetAgingDetails_Full(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetAgingDetails_Full(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetAgingDetails_Full(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetAgingDetails_Full(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetAgingDetails_Full(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetAgingDetails_Full(strBP, strAgeingCondition, "181")
                            End If
                        Else
                            If strLocalCurrency <> strBPCurrency Then
                                oDRow.Item("W00_30") = GetFCAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetFCAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetFCAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetFCAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetFCAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetFCAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetFCAgingDetails(strBP, strAgeingCondition, "181")
                            Else
                                oDRow.Item("W00_30") = GetAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetAgingDetails(strBP, strAgeingCondition, "181")
                            End If
                        End If


                        ds.Tables("Header").Rows.Add(oDRow)
                    Else
                        oDRow = ds.Tables("Header").NewRow()
                        oDRow.Item("CardCode") = strBP
                        oDRow.Item("Currency") = strBPCurrency
                        oDRow.Item("Title") = strDisplayOption
                        oDRow.Item("Ageingdate") = dtAgingdate
                        If strfrom <> "" Then
                            oDRow.Item("dtFrom") = dtFrom
                        End If
                        oDRow.Item("Currency") = strBPCurrency
                        If strto <> "" Then
                            oDRow.Item("dtto") = dtTo
                        Else
                            oDRow.Item("dtto") = Now.Date
                        End If
                        If strAcctType.ToUpper() = "ALL" Then
                            If strLocalCurrency <> strBPCurrency Then
                                oDRow.Item("W00_30") = GetFCAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetFCAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetFCAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetFCAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetFCAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetFCAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetFCAgingDetails(strBP, strAgeingCondition, "181")
                            Else
                                oDRow.Item("W00_30") = GetAgingDetails_Full(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetAgingDetails_Full(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetAgingDetails_Full(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetAgingDetails_Full(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetAgingDetails_Full(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetAgingDetails_Full(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetAgingDetails_Full(strBP, strAgeingCondition, "181")
                            End If
                        Else
                            If strLocalCurrency <> strBPCurrency Then
                                oDRow.Item("W00_30") = GetFCAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetFCAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetFCAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetFCAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetFCAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetFCAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetFCAgingDetails(strBP, strAgeingCondition, "181")
                            Else
                                oDRow.Item("W00_30") = GetAgingDetails(strBP, strAgeingCondition, "30")
                                oDRow.Item("W31_60") = GetAgingDetails(strBP, strAgeingCondition, "60")
                                oDRow.Item("W61_90") = GetAgingDetails(strBP, strAgeingCondition, "90")
                                oDRow.Item("W91_120") = GetAgingDetails(strBP, strAgeingCondition, "120")
                                oDRow.Item("W121_150") = GetAgingDetails(strBP, strAgeingCondition, "150")
                                oDRow.Item("W151_180") = GetAgingDetails(strBP, strAgeingCondition, "180")
                                oDRow.Item("W180_Above") = GetAgingDetails(strBP, strAgeingCondition, "181")
                            End If
                        End If
                        Dim st As String
                        st = "select isnull(sum(CheckSum),0) from RCT1 where duedate> '" & dtAgingdate.ToString("yyyy-MM-dd") & "' and DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "')"
                        Dim totRS As SAPbobsCOM.Recordset
                        totRS = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        totRS.DoQuery(st)
                        oDRow.Item("TotPDC") = totRS.Fields.Item(0).Value
                        st = "Select U_Name from OUSR where User_code='" & objSBOAPI.oCompany.UserName & "'"
                        totRS.DoQuery(st)
                        oDRow.Item("Fax") = totRS.Fields.Item(0).Value

                        ds.Tables("Header").Rows.Add(oDRow)
                    End If
                    Dim otemp As SAPbobsCOM.Recordset
                    otemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        otemp.DoQuery("Exec DropTable")
                    Catch ex As Exception
                    End Try
                    Dim dblBalan As Decimal = 0
                    Dim dblSOOpening As Double = 0
                    Dim dtPosting, dtDue As Date
                    Dim strsql1, strSystemCurrecny As String
                    strSystemCurrecny = objUtility.GetSystemCurrency()
                    If strReportCurrency = "L" Or strReportCurrency = "B" Then
                        If strLocalCurrency <> strBPCurrency Then
                            If strfrom <> "" Then
                                If strBPChoice = "C" Then
                                    strsql1 = "select  " & strShortname & ",Sum(A.FCDebit-A.FCcredit) from jdt1 A,ojdt B where a.TransId=B.TransId and (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & " < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & ""

                                    'strsql1 = "drop table " & strusercode & " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
                                    'strsql1 = " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
                                    'strsql1 = strsql1 & " (case when iscredit = 'D' then SUM(reconsum) else 0 end) as 'reconDeb'"
                                    'strsql1 = strsql1 & " into " & strusercode & " from OITR a inner join itr1 b on a.ReconNum =b.ReconNum inner join OCRD c on c.CardCode=b.ShortName "
                                    'strsql1 = strsql1 & " where IsCard='C'and (shortname = '" & strBP & "' or FatherCard='" & strBP & "') and recondate<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    'strsql1 = strsql1 & " group by transid,transrowid, shortname,IsCredit"
                                    'strsql1 = strsql1 & " select sum(x.Val1) from ( select SUM(debit -reconDeb)-SUM(credit - reconcred) 'Val1' from " & strusercode & " b "
                                    'strsql1 = strsql1 & " left outer  join jdt1 a on a.TransId =b.TransId and a.Line_ID=b.TransRowId"
                                    'strsql1 = strsql1 & " where " & strReconcilationCondition & " and A." & strReportFilterdate & "<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    'strsql1 = strsql1 & " and (a.U_CardCode='" & strBP & "' or  a." & strShortname & " = '" & strBP & "')" ' group by a.transid,DueDate having SUM(debit -reconDeb)<>0 or SUM(credit - reconcred) <>0"
                                    'strsql1 = strsql1 & " union all"
                                    ''  strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a where TransId not in (select TransId from " & strusercode & " )"
                                    ''strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a where TransId not in (select TransId from " & strusercode & " )"
                                    'strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a  where (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,Line_ID) )not in (select (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,TransRowId) ) from " & strusercode & " )"

                                    'strsql1 = strsql1 & "  and " & strReconcilationCondition & " and (U_CardCode='" & strBP & "' or " & strShortname & " = '" & strBP & "')" & " and " & strReportFilterdate & "<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    ''   strsql1 = strsql1 & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  " & strAgeingField
                                    'strsql1 = strsql1 & " ) x"
                                    oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Try
                                        oRecTemp.DoQuery("Exec DropTable")
                                    Catch ex As Exception

                                    End Try
                                    ' oTemp.DoQuery(strsql1)

                                ElseIf strBPChoice = "S" Then
                                    strsql1 = "select " & strShortname & ",Sum(A.FCDebit-A.FCcredit) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & " < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & ""
                                Else
                                    strsql1 = "select " & strShortname & ",Sum(A.FCDebit-A.FCcredit) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & "< '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & ""
                                End If
                                Try
                                    oRecTemp.DoQuery("Exec DropTable")
                                Catch ex As Exception

                                End Try
                                oRecTemp.DoQuery(strsql1)
                                dblCumulative = 0
                                For intRow11 As Integer = 0 To oRecTemp.RecordCount - 1
                                    dblCumulative = oRecTemp.Fields.Item(1).Value + dblCumulative
                                    oRecTemp.MoveNext()
                                Next

                                dblSOOpening = GetSOOB(strBP, dtFrom.ToString("yyyy-MM-dd"))
                                dblCumulative = dblCumulative + dblSOOpening
                            Else
                                dblCumulative = 0
                            End If
                        Else
                            If strfrom <> "" Then
                                If strBPChoice = "C" Then
                                    strsql1 = "select " & strShortname & ",Sum(A.Debit-A.credit) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & " < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "

                                    'strsql1 = "drop table " & strusercode & " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
                                    'strsql1 = " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
                                    'strsql1 = strsql1 & " (case when iscredit = 'D' then SUM(reconsum) else 0 end) as 'reconDeb'"
                                    'strsql1 = strsql1 & " into " & strusercode & " from OITR a inner join itr1 b on a.ReconNum =b.ReconNum inner join OCRD c on c.CardCode=b.ShortName "
                                    'strsql1 = strsql1 & " where IsCard='C'and (shortname = '" & strBP & "' or FatherCard='" & strBP & "') and recondate<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    'strsql1 = strsql1 & " group by transid,transrowid, shortname,IsCredit"
                                    'strsql1 = strsql1 & " select sum(x.Val1) from ( select SUM(debit -reconDeb)-SUM(credit - reconcred) 'Val1' from " & strusercode & " b "
                                    'strsql1 = strsql1 & " left outer  join jdt1 a on a.TransId =b.TransId and a.Line_ID=b.TransRowId"
                                    'strsql1 = strsql1 & " where " & strReconcilationCondition & " and A." & strReportFilterdate & "<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    'strsql1 = strsql1 & " and (a.U_CardCode='" & strBP & "' or  a." & strShortname & " = '" & strBP & "')" ' group by a.transid,DueDate having SUM(debit -reconDeb)<>0 or SUM(credit - reconcred) <>0"
                                    'strsql1 = strsql1 & " union all"
                                    ''strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a where TransId not in (select TransId from " & strusercode & " )"
                                    'strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a  where (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,Line_ID) )not in (select (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,TransRowId) ) from " & strusercode & " )"
                                    'strsql1 = strsql1 & "  and " & strReconcilationCondition & " and (U_CardCode='" & strBP & "' or " & strShortname & " = '" & strBP & "')" & " and " & strReportFilterdate & "<'" & dtFrom.ToString("yyyy-MM-dd") & "'"
                                    ''   strsql1 = strsql1 & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  " & strAgeingField
                                    'strsql1 = strsql1 & " ) x"
                                    oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Try
                                        oRecTemp.DoQuery("Exec DropTable")
                                    Catch ex As Exception

                                    End Try
                                    ' oTemp.DoQuery(strsql1)

                                ElseIf strBPChoice = "S" Then
                                    'strsql1 = "select shortname,Sum(A.credit-A.Debit) from jdt1 A,ojdt B where a.TransId=B.TransId and A.Shortname='" & strBP.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  group by shortname "
                                    strsql1 = "select " & strShortname & ",Sum(A.Debit-A.credit) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & " < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "
                                Else
                                    strsql1 = "select " & strShortname & ",Sum(A.Debit-A.credit) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A." & strReportFilterdate & " < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "
                                End If
                                Try
                                    oRecTemp.DoQuery("Exec DropTable")
                                Catch ex As Exception

                                End Try

                                oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Try
                                    oRecTemp.DoQuery("Exec DropTable")
                                Catch ex As Exception

                                End Try
                                oRecTemp.DoQuery(strsql1)
                                dblCumulative = 0
                                For intRow11 As Integer = 0 To oRecTemp.RecordCount - 1
                                    dblCumulative = oRecTemp.Fields.Item(1).Value + dblCumulative
                                    oRecTemp.MoveNext()
                                Next

                                dblSOOpening = GetSOOB(strBP, dtFrom.ToString("yyyy-MM-dd"))
                                dblCumulative = dblCumulative + dblSOOpening
                            Else
                                dblCumulative = 0
                            End If
                        End If
                    Else
                        If strSystemCurrecny <> strBPCurrency Then
                            If strfrom <> "" Then
                                If strBPChoice = "C" Then
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "   group by " & strShortname & ""
                                ElseIf strBPChoice = "S" Then
                                    'strsql1 = "select shortname,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A.Shortname='" & strBP.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by shortname "
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & ""
                                Else
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & ""
                                End If
                                oRecTemp.DoQuery(strsql1)
                                dblCumulative = 0
                                For intRow11 As Integer = 0 To oRecTemp.RecordCount - 1
                                    dblCumulative = oRecTemp.Fields.Item(1).Value + dblCumulative
                                    oRecTemp.MoveNext()
                                Next

                                dblSOOpening = GetSOOB(strBP, dtFrom.ToString("yyyy-MM-dd"))
                                dblCumulative = dblCumulative + dblSOOpening
                            Else
                                dblCumulative = 0
                            End If
                        Else
                            If strfrom <> "" Then
                                If strBPChoice = "C" Then
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "
                                ElseIf strBPChoice = "S" Then
                                    'strsql1 = "select shortname,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A.Shortname='" & strBP.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  group by shortname "
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "
                                Else
                                    strsql1 = "select " & strShortname & ",Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and  (A.U_Cardcode='" & strBP & "' or  A." & strShortname & "='" & strBP.Replace("'", "''") & "') and A.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by " & strShortname & " "
                                End If
                                oRecTemp.DoQuery(strsql1)
                                dblCumulative = 0
                                For intRow11 As Integer = 0 To oRecTemp.RecordCount - 1
                                    dblCumulative = oRecTemp.Fields.Item(1).Value + dblCumulative
                                    oRecTemp.MoveNext()
                                Next

                                dblSOOpening = GetSOOB(strBP, dtFrom.ToString("yyyy-MM-dd"))
                                dblCumulative = dblCumulative + dblSOOpening
                            Else
                                dblCumulative = 0
                            End If
                        End If
                    End If

                    dblOpenBalance = dblCumulative



                    strsql = "Select TransId, RefDate, DueDate, TaxDate, TransType, BaseRef, Ref1, Ref2, Ref3Line, LineMemo, ContraAct, AcctName,"
                    strsql = strsql & " isnull(Debit,0), isnull(Credit,0), JDT1.Project, OPRJ.PrjName, " & strShortname & ", isnull(FCDebit,0), isnull(FCCredit,0), isnull(SYSDeb,0), isnull(SYSCred,0), FCCurrency,isnull(U_SMNo,'') 'SMNO' , isnull(U_Cardcode,'') 'Branch' ,isnull(JDT1.TransCode,'') 'TranCode'"
                    strsql = strsql & " from JDT1 left outer join OPRJ on OPRJ.PrjCode=JDT1.PRoject  left outer join OACT   ON AcctCode = Account"
                    strsql = strsql & " where (jdt1.U_CardCode='" & strBP.Replace("'", "''") & "' or  jdt1." & strShortname & "='" & strBP.Replace("'", "''") & "') and " & strJournalDatecondition & " and " & strProjectCondition
                    strsql = strsql & " ORDER BY REFDATE,TRANSID"
                    Dim strLines As String


                    'strLines = "select T0.DocEntry,T0.DocNum,T0.DocDate,T0.DocDueDate,T0.[DocTotal], T0.[DocTotalFC], T0.[DocTotalSy],T0.JrnlMemo,T0.SlpCode 'SMNO' ,T0.CardCode 'BPCODE' from ORDR T0 inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                    ''  strLines = strLines & strSOBPCondition & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and  T0.DocStatus='O' and (T0.CardCode='" & strBP & "' or (T0.FatherCard='" & strBP & "' and T0.fatherType='P'))" & strPDCDate & " and T0.CardCode='" & strBP & "'"
                    'strLines = strLines & " 1=1 where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and  T0.DocStatus='O' and (T0.CardCode='" & strBP & "' or (T1.FatherCard='" & strBP & "' and T1.fatherType='P'))" & strPDCDate


                    strLines = "select T0.DocEntry,T0.DocDate,T0.DocDueDate,T0.TaxDate,17,T0.Docentry,convert(Varchar,T0.DocNum),'','',T0.JrnlMemo,'','',T0.[DocTotal],T0.DocTotal,T0.Project,T0.Project,T0.CardCode, T0.[DocTotalFC], T0.[DocTotalFC], T0.[DocTotalSy], T0.[DocTotalSy],T0.DocCur,T0.SlpCode 'SMNO' ,T0.CardCode 'BPCODE','',T0.NumAtCard from ORDR T0 inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                    ' strLines = strLines & strSOBPCondition & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and  T0.DocStatus='O' and (T0.CardCode='" & strBP & "' or (T0.FatherCard='" & strBP & "' and T0.fatherType='P'))" & strPDCDate & " and T0.CardCode='" & strBP & "'"
                    strLines = strLines & " 1=1 where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and ( T3.Seriesname<>'SR' and T3.Seriesname<>'GFC')) and  T0.DocStatus='O' and (T0.CardCode='" & strBP & "' or (T1.FatherCard='" & strBP & "' and T1.fatherType='P'))" & strPDCDate

                    Dim strSorting As String
                    strSorting = "x.RefDate"
                    Select Case strReportFilterdate
                        Case "RefDate"
                            strSorting = "x.RefDate"
                            ' strReportFilterdate = "RefDate"
                        Case "taxDate"
                            strSorting = "x.TaxDate"
                            'strReportFilterdate = "taxDate"
                        Case "DueDate"
                            strSorting = "x.DueDate"
                            ' strReportFilterdate = "DueDate"
                    End Select




                    strsql = "Select TransId, RefDate, DueDate, TaxDate, TransType, BaseRef, Ref1, Ref2, Ref3Line, LineMemo, ContraAct, AcctName,"
                    strsql = strsql & "  isnull(Debit,0) 'Debit', isnull(Credit,0) 'Credit', JDT1.Project, OPRJ.PrjName, ShortName, isnull(FCDebit,0) 'FCDebit', isnull(FCCredit,0) 'FCCredit', isnull(SYSDeb,0) 'SystDebit', isnull(SYSCred,0) 'SysCredit', FCCurrency,isnull(U_SMNo,'') 'SMNO' , isnull(U_Cardcode,'') 'Branch' ,isnull(JDT1.TransCode,'') 'TranCode'"
                    strsql = strsql & ", ' ' 'NumAtCard' from JDT1 left outer join OPRJ on OPRJ.PrjCode=JDT1.PRoject  left outer join OACT   ON AcctCode = Account"
                    strsql = strsql & " where   (jdt1.U_CardCode='" & strBP.Replace("'", "''") & "' or  jdt1." & strShortname & "='" & strBP.Replace("'", "''") & "') and " & strJournalDatecondition & " and " & strProjectCondition
                    ' strsql = strsql & " ORDER BY REFDATE,TRANSID"

                    strsql = "Select  * from  ( " & strsql & " Union All " & strLines & ") x order by  " & strSorting & " ,x.TransType,x.Ref1,X.Ref2 "

                    oRec.DoQuery(strsql)

                    'MsgBox(dblOpenBalance)
                    Dim dtTax As Date

                    For inti As Integer = 0 To oRec.RecordCount - 1
                        oStatic = aForm.Items.Item("23").Specific
                        oStatic.Caption = "Processing CardCode : " & strBP
                        ' objUtility.ShowWarningMessage("Processing CardCode : " & strBP)
                        dtPosting = oRec.Fields.Item(1).Value
                        dtDue = oRec.Fields.Item(2).Value
                        dtTax = oRec.Fields.Item(3).Value
                        If strLocalCurrency <> strBPCurrency Then
                            If strReportCurrency = "L" Then
                                dblDebit = oRec.Fields.Item(12).Value
                                dblCredit = oRec.Fields.Item(13).Value
                            ElseIf strReportCurrency = "B" Then
                                dblDebit = oRec.Fields.Item(17).Value
                                dblCredit = oRec.Fields.Item(18).Value
                            Else
                                dblDebit = oRec.Fields.Item(19).Value
                                dblCredit = oRec.Fields.Item(20).Value
                            End If
                        Else
                            If strReportCurrency = "L" Then
                                dblDebit = oRec.Fields.Item(12).Value
                                dblCredit = oRec.Fields.Item(13).Value
                            ElseIf strReportCurrency = "B" Then
                                dblDebit = oRec.Fields.Item(12).Value
                                dblCredit = oRec.Fields.Item(13).Value
                            Else
                                dblDebit = oRec.Fields.Item(19).Value
                                dblCredit = oRec.Fields.Item(20).Value
                            End If
                        End If
                        If oRec.Fields.Item(4).Value = "17" Then
                            dblCredit = 0
                            dblDebit = GetSOPendingtotal(oRec.Fields.Item(6).Value)
                        Else
                            dblCredit = dblCredit
                        End If
                        dblCumulative = dblCumulative + dblDebit - dblCredit
                        Dim strtransid, strtrnsName, strTran As String
                        strtransid = oRec.Fields.Item("TranCode").Value
                        If strtransid = "" Then
                            strTran = ""
                            strtrnsName = ""
                        Else
                            Dim oRS5 As SAPbobsCOM.Recordset
                            oRS5 = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS5.DoQuery("Select * from OTRC where TrnsCode='" & strtransid & "'")
                            strTran = oRS5.Fields.Item("TrnsCode").Value
                            strtrnsName = oRS5.Fields.Item("TrnsCodDsc").Value
                        End If
                        '  If dblCredit <> 0 Or dblDebit <> 0 Then
                        If 1 = 1 Then
                            oDRow = ds.Tables("AccountBalance").NewRow()
                            oDRow.Item("TrnsCode") = strTran
                            oDRow.Item("TrnsName") = strtrnsName
                            oDRow.Item("TransId") = oRec.Fields.Item(0).Value
                            oDRow.Item("RefDate") = oRec.Fields.Item(1).Value
                            oDRow.Item("DueDate") = dtDue
                            oDRow.Item("TaxDate") = dtTax
                            oDRow.Item("Transtype") = oRec.Fields.Item(4).Value
                            oDRow.Item("BaseRef") = oRec.Fields.Item(5).Value
                            oDRow.Item("Ref1") = oRec.Fields.Item(6).Value
                            oDRow.Item("Ref2") = oRec.Fields.Item(7).Value
                            oDRow.Item("Ref3Line") = oRec.Fields.Item(8).Value
                            oDRow.Item("LineMemo") = oRec.Fields.Item(9).Value
                            oDRow.Item("ContraAct") = oRec.Fields.Item(10).Value
                            oDRow.Item("AcctName") = oRec.Fields.Item(11).Value
                            oDRow.Item("Debit") = dblDebit

                            oDRow.Item("Credit") = dblCredit
                            oDRow.Item("Project") = oRec.Fields.Item(14).Value
                            oDRow.Item("ProjectName") = oRec.Fields.Item(15).Value
                            oDRow.Item("Cummulative") = dblCumulative
                            oDRow.Item("CardCode") = strBP
                            strBranch = oRec.Fields.Item("Branch").Value

                            oDRow.Item("Branch") = strBranch
                            strSlpCode = ""
                            strSlpCode = oRec.Fields.Item("SMNO").Value
                            Dim oSlpRec As SAPbobsCOM.Recordset
                            Dim strSlpName1 As String

                            strSlpName1 = ""
                            oSlpRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If strSlpCode <> "" Then
                                oSlpRec.DoQuery("Select * from OSLP where slpcode=" & CInt(strSlpCode))
                                oDRow.Item("SMNo") = strSlpCode 'oSlpRec.Fields.Item("SlpCode").Value
                                strSlpName1 = oSlpRec.Fields.Item("SlpName").Value
                                oDRow.Item("SMName") = strSlpName1 ' oSlpRec.Fields.Item("SlpName").Value
                            Else
                                oDRow.Item("SMNo") = strSlpCode 'oSlpRec.Fields.Item("SlpCode").Value
                                oDRow.Item("SMName") = strSlpName1 'oSlpRec.Fields.Item("SlpName").Value
                            End If

                            oDRow.Item("OB") = dblOpenBalance
                            'oDRow.Item("CurrentBalance") = dblBalan
                            oDRow.Item("OBBalance") = dblOpenBalance
                            oDRow.Item("NumAtCard") = oRec.Fields.Item("NumAtCard").Value
                            ds.Tables("AccountBalance").Rows.Add(oDRow)
                        End If
                        oRec.MoveNext()
                    Next


                    Dim oPDCRec As SAPbobsCOM.Recordset
                    Dim oTestRs As SAPbobsCOM.Recordset
                    oPDCRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTestRs = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim stPDCQuery As String
                    Dim dblPDCAmount As Double = 0
                    Dim dblStatementAmount, dblPDCTotaAmount As Double
                    dblStatementAmount = dblCumulative
                    If intReportChoice = 2 Then
                        'stPDCQuery = "select DueDate,DocNum,CheckNum,CheckSum from RCT1 where duedate>getdate() and  DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "'" & strPDCDate & ") order by DueDate,DocNum"
                        'stPDCQuery = "select DueDate,DocNum,CheckNum,CheckSum from RCT1 where duedate> '" & dtAgingdate.ToString("yyy-MM-dd") & "' and  DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "'" & strPDCDate & ") order by DueDate,DocNum"
                        stPDCQuery = "select DueDate,DocNum,CheckNum,CheckSum from RCT1 where duedate > '" & dtAgingdate.ToString("yyy-MM-dd") & "' and  DocNum in (Select Docentry from ORCT where CardCode='" & strBP.Replace("'", "''") & "' ) order by DueDate,DocNum"
                        oPDCRec.DoQuery(stPDCQuery)
                        For intPDC As Integer = 0 To oPDCRec.RecordCount - 1
                            oStatic = aForm.Items.Item("23").Specific
                            oStatic.Caption = "Processing CardCode : " & strBP
                            oDRow = ds.Tables("PDC").NewRow()
                            dblPDCAmount = dblPDCAmount + oPDCRec.Fields.Item(3).Value
                            oDRow.Item("CardCode") = strBP
                            oDRow.Item("DocDate") = oPDCRec.Fields.Item(0).Value
                            oTestRs.DoQuery("Select isnull(CounterRef,'') from ORCT where DocNum=" & oPDCRec.Fields.Item(1).Value)
                            oDRow.Item("DocNum") = oTestRs.Fields.Item(0).Value ' oPDCRec.Fields.Item(1).Value
                            oDRow.Item("SMNo") = ""
                            oDRow.Item("CheckNum") = oPDCRec.Fields.Item(2).Value
                            oDRow.Item("Amount") = oPDCRec.Fields.Item(3).Value
                            oDRow.Item("TotalPDC") = dblPDCAmount
                            dblPDCTotaAmount = dblCumulative - dblPDCAmount
                            oDRow.Item("NetAmount") = dblPDCTotaAmount
                            ds.Tables("PDC").Rows.Add(oDRow)
                            oPDCRec.MoveNext()
                        Next
                    End If
                End If
                oRecBP.MoveNext()
            Next
            If ds.Tables("AccountBalance").Rows.Count <= 0 Then
                oDRow = ds.Tables("AccountBalance").NewRow()
                oDRow.Item("CardCode") = strBP
                oDRow.Item("Cummulative") = dblCumulative
                oDRow.Item("OB") = dblOpenBalance
                ds.Tables("AccountBalance").Rows.Add(oDRow)
            End If
            oStatic = aForm.Items.Item("23").Specific
            oStatic.Caption = "Report Generation processing..."
            addCrystal(ds, intReportChoice)
            oStatic = aForm.Items.Item("23").Specific
            oStatic.Caption = ""
            Return True
        Catch ex As Exception
            If ex.Message.Contains("deadlocked on lock resources with another") = False Then

                objUtility.ShowErrorMessage(ex.Message)
            Else
                Return False
            End If

        End Try
    End Function

    Private Function GetSOPendingtotal(ByVal aDocNum As String) As Double
        Dim strLines As String
        Dim dblAmt As Double
        Dim oTS As SAPbobsCOM.Recordset
        oTS = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblAmt = 0

        strsql = "  select sum(T2.OpenQty*T2.Price)-  sum(t2.openqty*t2.price * t0.discprcnt/100)  from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode  "
        strsql = strsql & " where T0.DocNum=" & aDocNum & " and T0.DocStatus='O'"
        oTS.DoQuery(strsql)
        dblAmt = oTS.Fields.Item(0).Value
        strsql = "  select  sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode  "
        strsql = strsql & " where T0.DocNum=" & aDocNum & " and T0.DocStatus='O'"
        oTS.DoQuery(strsql)
        dblAmt = dblAmt + oTS.Fields.Item(0).Value
        Return dblAmt

    End Function

    Private Function GetSOOB(ByVal aCardCode As String, ByVal aFromDate As String) As Double
        Dim strLines As String
        Dim dblAmt As Double
        Dim oTS As SAPbobsCOM.Recordset
        oTS = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblAmt = 0
        strsql = "  select sum(T2.OpenQty*T2.Price) - sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode  "
        strsql = strsql & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and (T0.CardCode='" & aCardCode & "' or (T1.FatherCard='" & aCardCode & "' and T1.fatherType='P')) and T0.DocStatus='O' and T0.DocDate<'" & aFromDate & "'"
        '  strsql = strsql & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and (T0.CardCode='" & aCardCode & "') and T0.DocStatus='O' and T0.DocDate<'" & aFromDate & "'"
        oTS.DoQuery(strsql)
        dblAmt = oTS.Fields.Item(0).Value
        strsql = "  select  sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode  "
        strsql = strsql & " where  T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and (T0.CardCode='" & aCardCode & "' or (T1.FatherCard='" & aCardCode & "' and T1.fatherType='P')) and T0.DocStatus='O' and T0.DocDate<'" & aFromDate & "'"
        'strsql = strsql & " where  T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and (T0.CardCode='" & aCardCode & "') and T0.DocStatus='O' and T0.DocDate<'" & aFromDate & "'"
        oTS.DoQuery(strsql)
        dblAmt = dblAmt + oTS.Fields.Item(0).Value
        Return dblAmt

    End Function
    Private Function GetBPType(ByVal aCardCode As String) As String
        Dim oBPRec As SAPbobsCOM.Recordset
        oBPRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBPRec.DoQuery("Select cardType from ocrd where cardcode='" & aCardCode & "'")
        Return oBPRec.Fields.Item(0).Value
    End Function

    Private Sub StatementofAccount_Project(ByVal aForm As SAPbouiCOM.Form)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
        Dim strfrom, strto, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double

        oRecBP = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBalanceRs = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        dtFrom = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("FromDate").Value)
        strfrom = objSBOAPI.GetSBODateString(dtFrom)

        dtTo = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("ToDate").Value)
        strto = objSBOAPI.GetSBODateString(dtTo)
        If strto = "" Then
            dtTo = Now.Date
        End If
        strLocalCurrency = objUtility.GetLocalCurrency()

        strFromBP = objSBOAPI.GetEditText(aForm, "13")
        strToBP = objSBOAPI.GetEditText(aForm, "15")
        If strfrom <> "" Then
            dtFrom = objSBOAPI.GetDateTimeValue(strfrom)
        End If
        If strto <> "" Then
            dtTo = objSBOAPI.GetDateTimeValue(strto)
        End If

        Dim strCond, strCond1, strBPSQL, strAddressSQL, strBP, strBalanceSQL As String

        strBPSQL = "select Project,count(*) from jdt1 where 1=1 and  " & strJournalDatecondition & " and " & strProjectCondition & " group by Project"
        oRecBP.DoQuery(strBPSQL)
        ds.Clear()
        oRecBP.DoQuery(strBPSQL)
        ds.Clear()
        Dim strProjectCode As String
        Dim strTempJournalcondition As String

        For intRow As Integer = 0 To oRecBP.RecordCount - 1
            strProjectCode = oRecBP.Fields.Item(0).Value
            strBP = GetBPDetails(oRecBP.Fields.Item(0).Value) 'oRecBP.Fields.Item(0).Value
            strBPCurrency = objUtility.getBPCurrency_Project(strBP)
            '   strAddressSQL = "Select CardCode, CardName, Block, City, BillToDef, ZipCode, Address, County, Phone1, Fax, CntctPrsn, Notes From OCRD, OCRY where OCRD.Country = OCRY.Code and OCRD.CardCode='" & strBP.Replace("'", "''") & "'"
            strAddressSQL = "Select PrjName,U_Address,U_Telephone,U_Fax,U_Contperson from OPRJ where prjcode='" & strProjectCode & "'"
            If strReportCurrency = "L" Then
                strBPCurrency = objUtility.GetLocalCurrency()
            ElseIf strReportCurrency = "B" Then
                If strBPCurrency = "##" Then
                    strBPCurrency = objUtility.GetSystemCurrency()
                Else
                    strBPCurrency = strBPCurrency
                End If
            ElseIf strReportCurrency = "S" Then
                strBPCurrency = objUtility.GetSystemCurrency()
            Else
                strBPCurrency = strBPCurrency
            End If
            oRecTemp.DoQuery(strAddressSQL)
            If oRecTemp.RecordCount > 0 Then
                oDRow = ds.Tables("Header").NewRow()
                oDRow.Item("CardCode") = strProjectCode 'oRecTemp.Fields.Item(1).Value  'strBP
                oDRow.Item("CardName") = oRecTemp.Fields.Item(0).Value
                oDRow.Item("Block") = oRecTemp.Fields.Item(1).Value
                oDRow.Item("City") = "" 'oRecTemp.Fields.Item(3).Value
                oDRow.Item("BilltoDef") = "" '= oRecTemp.Fields.Item(4).Value
                oDRow.Item("Zipcode") = "" 'oRecTemp.Fields.Item(5).Value
                oDRow.Item("Address") = "" 'oRecTemp.Fields.Item(6).Value
                oDRow.Item("County") = "" 'oRecTemp.Fields.Item(7).Value
                oDRow.Item("Phone1") = oRecTemp.Fields.Item(2).Value
                oDRow.Item("Fax") = oRecTemp.Fields.Item(3).Value
                oDRow.Item("CntctPrsn") = oRecTemp.Fields.Item(4).Value
                oDRow.Item("Notes") = "" 'oRecTemp.Fields.Item(11).Value
                If strfrom <> "" Then
                    oDRow.Item("dtFrom") = dtFrom
                End If
                oDRow.Item("Currency") = strBPCurrency
                If strto <> "" Then
                    oDRow.Item("dtto") = dtTo
                Else
                    oDRow.Item("dtto") = Now.Date
                End If
                oDRow.Item("Ageingdate") = dtAgingdate
                If strLocalCurrency <> strBPCurrency Then
                    oDRow.Item("W00_30") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "30")
                    oDRow.Item("W31_60") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "60")
                    oDRow.Item("W61_90") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "90")
                    oDRow.Item("W91_180") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "180")
                    oDRow.Item("W181_Plus") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "181")
                Else
                    oDRow.Item("W00_30") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "30")
                    oDRow.Item("W31_60") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "60")
                    oDRow.Item("W61_90") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "90")
                    oDRow.Item("W91_180") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "180")
                    oDRow.Item("W181_Plus") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "181")
                End If
                oDRow.Item("Title") = strDisplayOption
                ds.Tables("Header").Rows.Add(oDRow)

            Else
                oDRow = ds.Tables("Header").NewRow()
                oDRow.Item("CardCode") = strBP
                oDRow.Item("Currency") = strBPCurrency
                oDRow.Item("Ageingdate") = dtAgingdate
                If strto <> "" Then
                    oDRow.Item("dtto") = dtTo
                Else
                    oDRow.Item("dtto") = Now.Date
                End If
                If strfrom <> "" Then
                    oDRow.Item("dtFrom") = dtFrom
                End If
                If strLocalCurrency <> strBPCurrency Then
                    oDRow.Item("W00_30") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "30")
                    oDRow.Item("W31_60") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "60")
                    oDRow.Item("W61_90") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "90")
                    oDRow.Item("W91_180") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "180")
                    oDRow.Item("W181_Plus") = GetFCAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "181")
                Else
                    oDRow.Item("W00_30") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "30")
                    oDRow.Item("W31_60") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "60")
                    oDRow.Item("W61_90") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "90")
                    oDRow.Item("W91_180") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "180")
                    oDRow.Item("W181_Plus") = GetAgingDetails_Project(strProjectCode, strBP, strAgeingCondition, "181")
                End If
                oDRow.Item("Title") = strDisplayOption

                ds.Tables("Header").Rows.Add(oDRow)
            End If

            '  strJournalDatecondition = strTempJournalcondition

            strsql = "Select TransId, RefDate, DueDate, TaxDate, TransType, BaseRef, Ref1, Ref2, Ref3Line, LineMemo, ContraAct, AcctName,"
            strsql = strsql & " isnull(Debit,0), isnull(Credit,0), JDT1.Project, OPRJ.PrjName, ShortName, isnull(FCDebit,0), isnull(FCCredit,0), isnull(SYSDeb,0), isnull(SYSCred,0), FCCurrency"
            strsql = strsql & " from JDT1 left outer join OPRJ on OPRJ.PrjCode=JDT1.PRoject  left outer join OACT   ON AcctCode = Account"
            ' strsql = strsql & " where  jdt1.shortname in (" & strBP.Replace("'", "''") & ") and " & strJournalDatecondition & " and " & strProjectCondition
            strsql = strsql & " where  jdt1.project in ('" & strProjectCode.Replace("'", "''") & "') and " & strJournalDatecondition & " and " & strProjectCondition
            strsql = strsql & " ORDER BY REFDATE,TRANSID"
            oRec.DoQuery(strsql)


            Dim dblBalan As Decimal = 0
            Dim dtPosting, dtDue As Date
            Dim strsql1, strSystemCurrecny, strBPCardCode As String
            strSystemCurrecny = objUtility.GetSystemCurrency()
            Dim oRecOpening As SAPbobsCOM.Recordset
            oRecOpening = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecOpening.DoQuery("Select cardcode,cardtype from OCRD where cardcode in (" & strBP & ")")
            dblCumulative = 0
            For intOP As Integer = 0 To oRecOpening.RecordCount - 1
                strBPChoice = oRecOpening.Fields.Item(1).Value
                strBPCardCode = oRecOpening.Fields.Item(0).Value
                strBPCurrency = objUtility.getBPCurrency(strBPCardCode)
                If strReportCurrency = "L" Or strReportCurrency = "B" Then
                    If strLocalCurrency <> strBPCurrency Then
                        If strfrom <> "" Then
                            If strBPChoice = "C" Then
                                strsql1 = "select A.project,Sum(A.FCDebit-A.FCcredit),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and  A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.project,A." & strShortname & ""
                            ElseIf strBPChoice = "S" Then
                                'strsql1 = "select A.project,Sum(A.FCcredit-A.FCDebit) from jdt1 A,ojdt B where a.TransId=B.TransId and A.ShortName='" & strBPCardCode & "' and A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.project "
                                strsql1 = "select A.project,Sum(A.FCDebit-A.FCcredit) ,A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and  A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.project,A." & strShortname & ""
                            Else
                                strsql1 = "select A.project,Sum(A.FCDebit-A.FCcredit) ,A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.project,A." & strShortname & ""
                            End If
                            oRecTemp.DoQuery(strsql1)
                            dblCumulative = dblCumulative + oRecTemp.Fields.Item(1).Value
                        Else
                            If dblCumulative = 0 Then
                                dblCumulative = 0
                            End If

                        End If
                    Else
                        If strfrom <> "" Then
                            If strBPChoice = "C" Then
                                strsql1 = "select A.project,Sum(A.Debit-A.credit),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "   group by A.project,A." & strShortname & ""
                            ElseIf strBPChoice = "S" Then
                                '  strsql1 = "select A.project, Sum(A.credit-A.Debit) from jdt1 A,ojdt B where a.TransId=B.TransId and A.ShortName='" & strBPCardCode & "' and A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  group by A.project"
                                strsql1 = "select A.project,Sum(A.Debit-A.credit),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.project,A." & strShortname & ""
                            Else
                                strsql1 = "select A.project,Sum(A.Debit-A.credit),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and  A.project in ('" & strProjectCode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.project,A." & strShortname & ""
                            End If
                            oRecTemp.DoQuery(strsql1)
                            dblCumulative = dblCumulative + oRecTemp.Fields.Item(1).Value
                        Else
                            If dblCumulative = 0 Then
                                dblCumulative = 0
                            End If

                        End If
                    End If
                Else
                    If strSystemCurrecny <> strBPCurrency Then
                        If strfrom <> "" Then
                            If strBPChoice = "C" Then
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.Project,A." & strShortname & ""
                            ElseIf strBPChoice = "S" Then
                                'strsql1 = "select A.Project,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A.ShortName='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project "
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) ,A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.Project,A." & strShortname & ""
                            Else
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project in '" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  and " & strReconcilationCondition & " group by A.Project,A." & strShortname & ""
                            End If

                            oRecTemp.DoQuery(strsql1)
                            dblCumulative = dblCumulative + oRecTemp.Fields.Item(1).Value
                        Else
                            If dblCumulative = 0 Then
                                dblCumulative = 0
                            End If

                        End If
                    Else
                        If strfrom <> "" Then
                            If strBPChoice = "C" Then
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) ,A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.Project,A." & strShortname & ""
                            ElseIf strBPChoice = "S" Then
                                ' strsql1 = "select A.Project,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A.ShortName='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project "
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.Project,A." & strShortname & ""
                            Else
                                strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred),A." & strShortname & " from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & "='" & strBPCardCode & "' and A.Project='" & strProjectCode.Replace("'", "''") & "' and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' and " & strReconcilationCondition & "  group by A.Project,A." & strShortname & ""
                            End If
                            oRecTemp.DoQuery(strsql1)
                            dblCumulative = dblCumulative + oRecTemp.Fields.Item(1).Value
                        Else
                            If dblCumulative = 0 Then
                                dblCumulative = 0
                            End If

                        End If
                    End If
                End If
                oRecOpening.MoveNext()
            Next

            'dblCumulative = GetOpeningBalance_Project(dtFrom, strfrom, strProjectCode, strBP)
            dblOpenBalance = dblCumulative
            strBPCurrency = objUtility.getBPCurrency_Project(strBP)
            Dim dtTax As Date
            For inti As Integer = 0 To oRec.RecordCount - 1
                dtPosting = oRec.Fields.Item(1).Value
                dtDue = oRec.Fields.Item(2).Value
                dtTax = oRec.Fields.Item(3).Value
                If strLocalCurrency <> strBPCurrency Then
                    If strReportCurrency = "L" Then
                        dblDebit = oRec.Fields.Item(12).Value
                        dblCredit = oRec.Fields.Item(13).Value
                    ElseIf strReportCurrency = "B" Then
                        dblDebit = oRec.Fields.Item(17).Value
                        dblCredit = oRec.Fields.Item(18).Value
                    Else
                        dblDebit = oRec.Fields.Item(19).Value
                        dblCredit = oRec.Fields.Item(20).Value
                    End If
                Else
                    If strReportCurrency = "L" Then
                        dblDebit = oRec.Fields.Item(12).Value
                        dblCredit = oRec.Fields.Item(13).Value
                    ElseIf strReportCurrency = "B" Then
                        dblDebit = oRec.Fields.Item(12).Value
                        dblCredit = oRec.Fields.Item(13).Value
                    Else
                        dblDebit = oRec.Fields.Item(19).Value
                        dblCredit = oRec.Fields.Item(20).Value
                    End If
                End If

                strBPChoice = GetBPType(oRec.Fields.Item(16).Value)

                If strBPChoice = "C" Then
                    dblCumulative = dblCumulative + dblDebit - dblCredit
                Else
                    dblCumulative = dblCumulative + dblDebit - dblCredit
                End If
                '  dblCredit = dblCredit * -1
                If dblCredit <> 0 Or dblDebit <> 0 Then
                    oDRow = ds.Tables("AccountBalance").NewRow()
                    oDRow.Item("TransId") = oRec.Fields.Item(0).Value
                    oDRow.Item("RefDate") = oRec.Fields.Item(1).Value
                    oDRow.Item("DueDate") = dtDue
                    oDRow.Item("TaxDate") = dtTax
                    oDRow.Item("Transtype") = oRec.Fields.Item(4).Value
                    oDRow.Item("BaseRef") = oRec.Fields.Item(5).Value
                    oDRow.Item("Ref1") = oRec.Fields.Item(6).Value
                    oDRow.Item("Ref2") = oRec.Fields.Item(7).Value
                    oDRow.Item("Ref3Line") = oRec.Fields.Item(8).Value
                    oDRow.Item("LineMemo") = oRec.Fields.Item(9).Value
                    oDRow.Item("ContraAct") = oRec.Fields.Item(10).Value
                    oDRow.Item("AcctName") = oRec.Fields.Item(11).Value
                    oDRow.Item("Debit") = dblDebit
                    oDRow.Item("Credit") = dblCredit
                    oDRow.Item("Project") = oRec.Fields.Item(14).Value
                    oDRow.Item("ProjectName") = oRec.Fields.Item(15).Value
                    oDRow.Item("CardCode") = strProjectCode 'oRec.Fields.Item(16).Value
                    oDRow.Item("Cummulative") = dblCumulative
                    oDRow.Item("OB") = dblOpenBalance
                    ds.Tables("AccountBalance").Rows.Add(oDRow)
                End If
                oRec.MoveNext()
            Next
            oRecBP.MoveNext()
        Next
        addCrystal(ds, 1)
    End Sub
#End Region

#Region "Get Ageing Details"

    Private Function GetAgingDetails_SalesOrder(ByVal strCardcode As String, ByVal aBPCond As String, ByVal aDateCo As String, ByVal dtAgingDate As Date, ByVal strchoice As String, ByVal intSlpCode As Integer) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""

        Select Case strchoice
            Case "30"
                strsql = "  select sum(T2.OpenQty*T2.Price) - sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and  (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and  (T0.CardCode='" & strCardcode & "') and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T2.OpenQty*T2.Price) - sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P'))and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "')and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T2.OpenQty*T2.Price)- sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "') and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T2.OpenQty*T2.Price)- sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and     (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and     (T0.CardCode='" & strCardcode & "') and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T2.OpenQty*T2.Price)-  sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"
                ' strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   (T0.CardCode='" & strCardcode & "') and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T2.OpenQty*T2.Price) - sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   between 151 and 180"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "') and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')   >150"
            Case "181"
                strsql = "  select sum(T2.OpenQty*T2.Price) - sum(t2.openqty*t2.price * t0.discprcnt/100) from ORDR T0 inner Join  RDR1 T2 on T2.DocEntry=T0.DocEntry  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
                'strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "') and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        strsql = ""

        Select Case strchoice
            Case "30"
                strsql = "  select  sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                ' strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   T0.CardCode='" & strCardcode & "' and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "') between 0 and 30"
            Case "60"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and     (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P'))and   T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 31 and 60"
            Case "90"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0 inner Join   OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 61 and 90"

            Case "120"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and   (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 91 and 120"
            Case "150"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0  inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P'))and T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 121 and 150"

            Case "180"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode  and "
                strsql = strsql & aBPCond & " where T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and     (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O' and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  between 151 and 180"
            Case "181"
                strsql = "  select sum(T0.[VatSum]-T0.[VatPaid]) from ORDR T0   inner Join OCRD T1 on T1.CardCode=T0.Cardcode and "
                strsql = strsql & aBPCond & " where  T0.Series in (SELECT T3.Series  FROM NNM1 T3 where T3.ObjectCode=17 and T3.Seriesname<>'GFC') and    (T0.CardCode='" & strCardcode & "' or (T1.FatherCard='" & strCardcode & "' and T1.fatherType='P')) and  T0.DocStatus='O'  and DATEDIFF(D,COALESCE(T0.DocDate,'" & dtAgingDate.ToString("yyyy-MM-dd") & "'),'" & dtAgingDate.ToString("yyyy-MM-dd") & "')  > 180"
        End Select

        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = dblAmount + oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function

    Private Function GetAgingDetails(ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql, strsql1, strAgeingField As String
        strsql = ""
        strsql1 = ""
        'strsql1 = "drop table " & strusercode & " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
        'strsql1 = strsql1 & " (case when iscredit = 'D' then SUM(reconsum) else 0 end) as 'reconDeb'"
        'strsql1 = strsql1 & " into " & strusercode & " from OITR a inner join itr1 b on a.ReconNum =b.ReconNum where IsCard='C'"
        ''strsql1 = strsql1 & " and (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "')"
        ' ''strsql1 = strsql1 & " and " & strCondition
        'strsql1 = strsql1 & " group by transid,transrowid, shortname,IsCredit"
        'strsql1 = strsql1 & " select a.TransId,DueDate,SUM(debit -reconDeb),SUM(credit - reconcred) from tmpoitr b "
        'strsql1 = strsql1 & " left outer  join jdt1 a on a.TransId =b.TransId and a.Line_ID=b.TransRowId"
        'strsql1 = strsql1 & " where " & strCondition & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
        'strsql1 = strsql1 & " and (a.U_CardCode='" & strCardcode & "' or  a." & strShortname & " = '" & strCardcode & "') group by a.transid,DueDate having SUM(debit -reconDeb)<>0 or SUM(credit - reconcred) <>0"
        'strsql1 = strsql1 & " union all"
        'strsql1 = strsql1 & " select  TransId,DueDate,SUM(debit),SUM(credit) from JDT1 where TransId not in (select TransId from " & strusercode & " )"
        ''--and Line_ID not in (select transrowid from " & strusercode & ")
        'strsql1 = strsql1 & "  and " & strCondition & " and (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "')"
        'strsql1 = strsql1 & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30 group by transid,DueDate"
        'strsql1 = strsql1 & " having SUM(debit)<>0 or SUM(credit) <>0"
        'strsql1 = strsql1 & " order by duedate"
        Dim strusercode As String
        strusercode = objSBOAPI.oCompany.UserName
        strusercode = "[" & strusercode & "_SOA_tmpoitr]"

        If strBPChoice = "C" Or strBPChoice = "P" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred  <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    strAgeingField = "between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    strAgeingField = " between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    strAgeingField = "between 61 and 90"
                Case "120"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                    strAgeingField = "between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                    strAgeingField = "between 121 and 150"

                Case "180"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
                    strAgeingField = "between 151 and 180"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    strAgeingField = " >180"
            End Select
        ElseIf strBPChoice = "S" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"


                Case "120"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                Case "180"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  between 151 and 180"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
            End Select
        End If
        Dim oTemp As SAPbobsCOM.Recordset

        Try
            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql1 = "drop table " & strusercode & " select * from OCRD where CardCode='D'"
            oTemp.DoQuery(strsql1)
        Catch ex As Exception
        End Try
        If strsql <> "" Then

            strsql1 = "drop table " & strusercode & " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
            strsql1 = " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
            strsql1 = strsql1 & " (case when iscredit = 'D' then SUM(reconsum) else 0 end) as 'reconDeb'"
            strsql1 = strsql1 & " into " & strusercode & " from OITR a inner join itr1 b on a.ReconNum =b.ReconNum inner join OCRD C on b.shortname=c.CardCode where IsCard='C' and recondate<='" & dtAgingdate.ToString("yyyy-MM-dd") & "'"
            strsql1 = strsql1 & " and (ShortName='" & strCardcode & "' or FatherCard = '" & strCardcode & "')"
            ''strsql1 = strsql1 & " and " & strCondition
            strsql1 = strsql1 & " group by transid,transrowid, shortname,IsCredit"
            strsql1 = strsql1 & " select sum(x.Val1) from ( select SUM(debit -reconDeb)-SUM(credit - reconcred) 'Val1' from " & strusercode & " b "
            strsql1 = strsql1 & " left outer  join jdt1 a on a.TransId =b.TransId and a.Line_ID=b.TransRowId"
            strsql1 = strsql1 & " where " & strCondition & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') " & strAgeingField
            strsql1 = strsql1 & " and (a.U_CardCode='" & strCardcode & "' or  a." & strShortname & " = '" & strCardcode & "')" ' group by a.transid,DueDate having SUM(debit -reconDeb)<>0 or SUM(credit - reconcred) <>0"
            strsql1 = strsql1 & " union all"
            '  strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 where TransId not in (select TransId from " & strusercode & " )"
            '--and Line_ID not in (select transrowid from " & strusercode & ")
            'strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 a where TransId not in (select TransId from " & strusercode & " )"
            strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1   where (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,Line_ID) )not in (select (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,TransRowId) ) from " & strusercode & " )"

            strsql1 = strsql1 & "  and " & strCondition & " and (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "')"
            strsql1 = strsql1 & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  " & strAgeingField
            strsql1 = strsql1 & " ) x drop table " & strusercode
            'strsql1 = strsql1 & " having SUM(debit)<>0 or SUM(credit) <>0"
            'strsql1 = strsql1 & " order by duedate"


            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                '    oTemp.DoQuery("Exec DropTable")
            Catch ex As Exception

            End Try


            oTemp.DoQuery(strsql1)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        Try
            '  oTemp.DoQuery("Exec DropTable")
        Catch ex As Exception

        End Try

        oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select SlpCode from OCRD where CardCode='" & strCardcode & "'")
        dblAmount = dblAmount + GetAgingDetails_SalesOrder(strCardcode, "1=1", "2=2", dtAgingdate, strchoice, oTemp.Fields.Item(0).Value)
        oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ' oTemp.DoQuery("Exec DropTable")
        Catch ex As Exception

        End Try
        Return dblAmount
    End Function

    Private Function GetAgingDetails_Full(ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql, strsql1, strAgeingField As String
        strsql = ""
        strAgeingField = ""
        Dim strusercode As String
        strusercode = objSBOAPI.oCompany.UserName
        ' strusercode = strusercode & "_SOA_tmpoitr"
        strusercode = "[" & strusercode & "_SOA_tmpoitr]"
        If strBPChoice = "C" Or strBPChoice = "P" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    strAgeingField = "between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    strAgeingField = "between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    strAgeingField = "between 61 and 90"

                Case "120"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                    strAgeingField = "between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                    strAgeingField = "between 121 and 150"
                Case "180"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and   " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
                    strAgeingField = "between 151 and 180"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(Debit,0)),0) - isnull(Sum(isnull(Credit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and   " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    strAgeingField = " > 180"
            End Select
        ElseIf strBPChoice = "S" Then
            Select Case strchoice
                Case "30"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                Case "60"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                Case "90"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"

                Case "120"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 120"
                Case "150"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and  " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 121 and 150"
                Case "180"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and   " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
                Case "181"
                    strsql = " Select  isnull(Sum(isnull(Credit,0)),0) - isnull(Sum(isnull(Debit,0)),0) From JDT1 "
                    strsql = strsql & " where (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "') and   " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
            End Select
        End If
        Dim oTemp As SAPbobsCOM.Recordset
        Try
            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql1 = "drop table " & strusercode & " select * from OCRD where CardCode='D'"
            oTemp.DoQuery(strsql1)
        Catch ex As Exception
        End Try
        If strsql <> "" Then

            'oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oTemp.DoQuery(strsql)
            'dblAmount = oTemp.Fields.Item(0).Value
            strsql1 = "drop table " & strusercode & " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
            strsql1 = " select transid,transrowid,shortname, (case when iscredit = 'C' then SUM(reconsum) else 0 end) as 'reconcred',"
            strsql1 = strsql1 & " (case when iscredit = 'D' then SUM(reconsum) else 0 end) as 'reconDeb'"
            strsql1 = strsql1 & " into " & strusercode & "  from OITR a inner join itr1 b on a.ReconNum =b.ReconNum inner join OCRD C on b.shortname=c.CardCode where IsCard='C' and recondate<='" & dtAgingdate.ToString("yyyy-MM-dd") & "'"
            strsql1 = strsql1 & " and (ShortName='" & strCardcode & "' or FatherCard = '" & strCardcode & "')"
            ''strsql1 = strsql1 & " and " & strCondition
            strsql1 = strsql1 & " group by transid,transrowid, shortname,IsCredit"
            strsql1 = strsql1 & " select sum(x.Val1) from ( select SUM(debit -reconDeb)-SUM(credit - reconcred) 'Val1' from " & strusercode & " b "
            strsql1 = strsql1 & " left outer  join jdt1 a on a.TransId =b.TransId and a.Line_ID=b.TransRowId"
            strsql1 = strsql1 & " where " & strCondition & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') " & strAgeingField
            strsql1 = strsql1 & " and (a.U_CardCode='" & strCardcode & "' or  a." & strShortname & " = '" & strCardcode & "')" ' group by a.transid,DueDate having SUM(debit -reconDeb)<>0 or SUM(credit - reconcred) <>0"
            strsql1 = strsql1 & " union all"
            ' strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1 where TransId not in (select TransId from " & strusercode & " )"
            '--and Line_ID not in (select transrowid from " & strusercode & ")
            strsql1 = strsql1 & " select  SUM(debit)-SUM(credit) 'Val1'  from JDT1  where (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,Line_ID) )not in (select (convert(VARCHAR,TransId)+'-'+convert(VARCHAR,TransRowId) ) from " & strusercode & " )"
            strsql1 = strsql1 & "  and " & strCondition & " and (U_CardCode='" & strCardcode & "' or " & strShortname & " = '" & strCardcode & "')"
            strsql1 = strsql1 & " and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "')  " & strAgeingField
            strsql1 = strsql1 & " ) x drop table " & strusercode
            'strsql1 = strsql1 & " having SUM(debit)<>0 or SUM(credit) <>0"
            'strsql1 = strsql1 & " order by duedate"


            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                '  oTemp.DoQuery("Exec DropTable")
            Catch ex As Exception
            End Try
            oTemp.DoQuery(strsql1)
            dblAmount = oTemp.Fields.Item(0).Value
        End If

        oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select SlpCode from OCRD where CardCode='" & strCardcode & "'")
        dblAmount = dblAmount + GetAgingDetails_SalesOrder(strCardcode, "1=1", "2=2", dtAgingdate, strchoice, oTemp.Fields.Item(0).Value)
        Return dblAmount
    End Function

    Private Function GetFCAgingDetails(ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""
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
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.BalFCDeb<>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
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
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.BalFCDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
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
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.[BalScCred] <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
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
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <>0 or JDT1.[BalScDeb] <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 151 and 180"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull([BalFCCred],0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " = '" & strCardcode & "' and  (JDT1.BalFCCred <> 0 or JDT1.BalFCDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            End If
        End If
        If strsql <> "" Then
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strsql)
            dblAmount = oTemp.Fields.Item(0).Value
        End If
        Return dblAmount
    End Function
    Private Function GetOpeningBalance_Project(ByVal dtFrom As Date, ByVal strfrom As String, ByVal strproject As String, ByVal strCardcode As String) As Double
        Dim dblAmount, dblCumulative As Double
        Dim strsql As String
        strsql = ""
        Dim oTempAgeing, oRecTemp As SAPbobsCOM.Recordset
        oRecTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempAgeing = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempAgeing.DoQuery("Select CardCode,CardType from OCRD where CardCode in (" & strCardcode & ")")
        dblAmount = 0
        For intAgeingrow As Integer = 0 To oTempAgeing.RecordCount - 1
            If oTempAgeing.Fields.Item(1).Value = "S" Then
                strBPChoice = "S"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            ElseIf oTempAgeing.Fields.Item(1).Value = "C" Then
                strBPChoice = "C"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            End If
            Dim strSystemCurrecny, strsql1 As String
            strSystemCurrecny = objUtility.GetSystemCurrency()
            If strReportCurrency = "L" Or strReportCurrency = "B" Then
                If strLocalCurrency <> strBPCurrency Then
                    If strfrom <> "" Then
                        If strBPChoice = "C" Then
                            strsql1 = "select A.project,Sum(A.FCDebit-A.FCcredit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.project"
                        ElseIf strBPChoice = "S" Then
                            strsql1 = "select A.project,Sum(A.FCcredit-A.FCDebit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.project "
                        Else
                            strsql1 = "select A.project,Sum(A.FCDebit-A.FCcredit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.project"
                        End If
                        oRecTemp.DoQuery(strsql1)
                        dblCumulative = oRecTemp.Fields.Item(1).Value
                    Else
                        dblCumulative = 0
                    End If
                Else
                    If strfrom <> "" Then
                        If strBPChoice = "C" Then
                            strsql1 = "select A.project,Sum(A.Debit-A.credit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  group by A.project"
                        ElseIf strBPChoice = "S" Then
                            strsql1 = "select A.project, Sum(A.credit-A.Debit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "'  group by A.project"
                        Else
                            strsql1 = "select A.project,Sum(A.Debit-A.credit) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.project"
                        End If
                        oRecTemp.DoQuery(strsql1)
                        dblCumulative = oRecTemp.Fields.Item(1).Value
                    Else
                        dblCumulative = 0
                    End If
                End If
            Else
                If strSystemCurrecny <> strBPCurrency Then
                    If strfrom <> "" Then
                        If strBPChoice = "C" Then
                            strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project"
                        ElseIf strBPChoice = "S" Then
                            strsql1 = "select A.Project,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project "
                        Else
                            strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project"
                        End If

                        oRecTemp.DoQuery(strsql1)
                        dblCumulative = oRecTemp.Fields.Item(1).Value
                    Else
                        dblCumulative = 0
                    End If
                Else
                    If strfrom <> "" Then
                        If strBPChoice = "C" Then
                            strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project"
                        ElseIf strBPChoice = "S" Then
                            strsql1 = "select A.Project,Sum(A.SYSCred-A.SYSDeb) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project "
                        Else
                            strsql1 = "select A.Project,Sum(A.SYSDeb-A.SYSCred) from jdt1 A,ojdt B where a.TransId=B.TransId and A." & strShortname & " in ('" & strCardcode.Replace("'", "''") & "') and B.REFDATE < '" & dtFrom.ToString("yyyy-MM-dd") & "' group by A.Project"
                        End If
                        oRecTemp.DoQuery(strsql1)
                        dblCumulative = oRecTemp.Fields.Item(1).Value
                    Else
                        dblCumulative = 0
                    End If
                End If

            End If
            If strsql <> "" Then
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strsql)
                If strBPChoice = "S" Then
                    dblAmount = dblAmount - oTemp.Fields.Item(0).Value
                Else
                    dblAmount = dblAmount + oTemp.Fields.Item(0).Value
                End If
            End If
            oTempAgeing.MoveNext()
        Next
        Return dblAmount
    End Function

    Private Function GetAgingDetails_Project(ByVal strproject As String, ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""
        Dim oTempAgeing As SAPbobsCOM.Recordset
        oTempAgeing = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempAgeing.DoQuery("Select CardCode,CardType from OCRD where CardCode in (" & strCardcode & ")")
        dblAmount = 0
        For intAgeingrow As Integer = 0 To oTempAgeing.RecordCount - 1
            If oTempAgeing.Fields.Item(1).Value = "S" Then
                strBPChoice = "S"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            ElseIf oTempAgeing.Fields.Item(1).Value = "C" Then
                strBPChoice = "C"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            End If
            If strBPChoice = "C" Or strBPChoice = "P" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    Case "180"
                        strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull(BalDueDeb,0)),0) - isnull(Sum(isnull(BalDueCred,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            ElseIf strBPChoice = "S" Then
                Select Case strchoice
                    Case "30"
                        strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                    Case "60"
                        strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                    Case "90"
                        strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                    Case "180"
                        strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and  (JDT1.BalDueCred <>0 or JDT1.BalDueDeb <>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                    Case "181"
                        strsql = " Select  isnull(Sum(isnull(BalDueCred,0)),0) - isnull(Sum(isnull(BalDueDeb,0)),0) From JDT1 "
                        strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and  (JDT1.BalDueCred <> 0 or JDT1.BalDueDeb <> 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                End Select
            End If
            If strsql <> "" Then
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strsql)
                If strBPChoice = "S" Then
                    dblAmount = dblAmount - oTemp.Fields.Item(0).Value
                Else
                    dblAmount = dblAmount + oTemp.Fields.Item(0).Value
                End If
            End If
            oTempAgeing.MoveNext()
        Next
        Return dblAmount
    End Function

    Private Function GetFCAgingDetails_Project(ByVal strproject As String, ByVal strCardcode As String, ByVal strCondition As String, ByVal strchoice As String) As Double
        Dim dblAmount As Double
        Dim strsql As String
        strsql = ""
        Dim oTempAgeing As SAPbobsCOM.Recordset
        oTempAgeing = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempAgeing.DoQuery("Select CardCode,CardType from OCRD where CardCode in (" & strCardcode & ")")
        dblAmount = 0
        For intAgeingrow As Integer = 0 To oTempAgeing.RecordCount - 1
            If oTempAgeing.Fields.Item(1).Value = "S" Then
                strBPChoice = "S"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            ElseIf oTempAgeing.Fields.Item(1).Value = "C" Then
                strBPChoice = "C"
                strCardcode = "'" & oTempAgeing.Fields.Item(0).Value & "'"
            End If
            If strReportCurrency = "L" Or strReportCurrency = "B" Then
                If strBPChoice = "C" Or strBPChoice = "P" Then
                    Select Case strchoice
                        Case "30"
                            strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                        Case "60"
                            strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                        Case "90"
                            strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and  Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                        Case "180"
                            strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and  (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                        Case "181"
                            strsql = " Select  isnull(Sum(isnull(BalFCDeb,0)),0) - isnull(Sum(isnull(BalFCCred,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and  (JDT1.BalFCCred > 0 or JDT1.BalFCDeb > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    End Select
                ElseIf strBPChoice = "S" Then
                    Select Case strchoice
                        Case "30"
                            strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                        Case "60"
                            strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where  " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                        Case "90"
                            strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                        Case "180"
                            strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and  (JDT1.BalFCCred>0 or JDT1.BalFCDeb>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                        Case "181"
                            strsql = " Select  isnull(Sum(isnull(BalFCCred,0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "') and  (JDT1.BalFCCred > 0 or JDT1.BalFCDeb > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(" & strReportFilterdate & ",'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    End Select
                End If
            Else
                If strBPChoice = "C" Or strBPChoice = "P" Then
                    Select Case strchoice
                        Case "30"
                            strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.[BalScCred]>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                        Case "60"
                            strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.[BalScCred]>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                        Case "90"
                            strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.[BalScCred]>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                        Case "180"
                            strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and  (JDT1.[BalScCred]>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                        Case "181"
                            strsql = " Select  isnull(Sum(isnull([BalScDeb],0)),0) - isnull(Sum(isnull([BalScCred],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and  (JDT1.[BalScCred] > 0 or JDT1.[BalScDeb] > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    End Select
                ElseIf strBPChoice = "S" Then
                    Select Case strchoice
                        Case "30"
                            strsql = " Select  isnull(Sum(isnull([BalScCred],0)),0) - isnull(Sum(isnull([BalScDeb],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.BalFCCred > 0 or JDT1.BalScDeb > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 0 and 30"
                        Case "60"
                            strsql = " Select  isnull(Sum(isnull([BalScCred],0)),0) - isnull(Sum(isnull([BalScDeb],0)),0) From JDT1 "
                            strsql = strsql & " where  " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.BalFCCred > 0 or JDT1.[BalScDeb] > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 31 and 60"
                        Case "90"
                            strsql = " Select  isnull(Sum(isnull([BalScCred],0)),0) - isnull(Sum(isnull([BalScDeb],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and (JDT1.BalFCCred>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 61 and 90"
                        Case "180"
                            strsql = " Select  isnull(Sum(isnull([BalScCred],0)),0) - isnull(Sum(isnull([BalScDeb],0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')   and  (JDT1.BalFCCred>0 or JDT1.[BalScDeb]>0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') between 91 and 180"
                        Case "181"
                            strsql = " Select  isnull(Sum(isnull([BalScCred],0)),0) - isnull(Sum(isnull(BalFCDeb,0)),0) From JDT1 "
                            strsql = strsql & " where " & strShortname & " in (" & strCardcode & ") and Project in('" & strproject & "')  and  (JDT1.BalFCCred > 0 or JDT1.BalFCDeb > 0) and " & strCondition & "  and DATEDIFF(D,COALESCE(refdate,'" & dtAgingdate.ToString("yyyy-MM-dd") & "'),'" & dtAgingdate.ToString("yyyy-MM-dd") & "') > 180"
                    End Select
                End If
            End If
            If strsql <> "" Then
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = objSBOAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strsql)
                If strBPChoice = "S" Then
                    dblAmount = dblAmount - oTemp.Fields.Item(0).Value
                Else
                    dblAmount = dblAmount + oTemp.Fields.Item(0).Value
                End If
            End If
            oTempAgeing.MoveNext()
        Next
        Return dblAmount
    End Function
#End Region

#Region "Validations"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim strBP As String
        oCombo = aForm.Items.Item("5").Specific
        If oCombo.Selected.Value = "" Then
            objUtility.ShowErrorMessage("Date parameter is missing...")
            Return False
        End If

        oCombo = aForm.Items.Item("11").Specific
        strBP = oCombo.Selected.Value

        If strBP = "" Then
            objUtility.ShowErrorMessage("Select the BP Type")
            Return False
        End If

        oCombo = aForm.Items.Item("17").Specific
        If oCombo.Selected.Value <> "" Then
            If strBP = "C" And oCombo.Selected.Value = "S" Then
                objUtility.ShowErrorMessage("Select the correct group type")
                Return False
            ElseIf strBP = "S" And oCombo.Selected.Value = "C" Then
                objUtility.ShowErrorMessage("Select the correct group type")
                Return False
            ElseIf strBP = "P" And oCombo.Selected.Value <> "" Then
                objUtility.ShowErrorMessage("BP Group selection is not allowed for Project range parameters")
                Return False
            End If
        End If

        If objSBOAPI.GetEditText(aForm, "22") = "" Then
            objUtility.ShowErrorMessage("Ageing date is missing")
            Return False
        End If

        oCombo = aForm.Items.Item("25").Specific
        If oCombo.Selected.Value = "" Then
            objUtility.ShowErrorMessage("Select the report currency ...")
            Return False
        End If

        oCombo = aForm.Items.Item("34").Specific
        If oCombo.Selected.Value = "" Then
            objUtility.ShowErrorMessage("Report Type is missing")
            Return False
        End If
        Return True
    End Function
#End Region

#Region "Add Crystal Report"

    'Private Sub callreport(ByVal ds1 As DataSet, ByVal aChoice As Integer)
    '    Dim thread As Thread = New Thread(New ThreadStart(AddressOf addCrystal))
    '    thread.Start()
    'End Sub
    'Public Sub Test()

    'End Sub


    Private Sub addCrystal(ByVal ds1 As DataSet, ByVal aChoice As Integer)

        Dim strFilename, strCompanyName, stfilepath As String
        Dim strReportFileName As String
        If aChoice = 1 Then
            strReportFileName = "AcctStatement_SalesPerson.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\ActStatment_MonthEndBatch"
        Else
            strReportFileName = "AcctStatement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\AccountStatement"
        End If
        strCompanyName = objSBOAPI.SBO_Appln.Company.Name
        'strReportFileName = strCompanyName & "_" & strReportFileName
        strReportFileName = strReportFileName

        strFilename = strFilename & ".pdf"
        stfilepath = System.Windows.Forms.Application.StartupPath & "\Reports\" & strReportFileName
        If File.Exists(stfilepath) = False Then
            objSBOAPI.SBO_Appln.MessageBox("Report file does not exists")
            Exit Sub
        End If
        If ds1.Tables.Item("AccountBalance").Rows.Count > 0 Then
            'Thread.SetApartmentState(ApartmentState.STA)
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\Reports\" & strReportFileName)
            cryRpt.SetDataSource(ds1)
            If strReportviewOption = "W" Then
                Dim mythread As New System.Threading.Thread(AddressOf openFileDialog)
              
                mythread.SetApartmentState(ApartmentState.STA)
                mythread.Start()
                mythread.Join()
               
                'Dim objPL As New frmReportViewer
                'objPL.iniViewer = AddressOf objPL.GenerateReport
                'objPL.rptViewer.ReportSource = cryRpt
                'objPL.rptViewer.Refresh()
                'objPL.rptViewer.Refresh()
                'objPL.WindowState = FormWindowState.Maximized
                'objPL.ShowDialog()
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
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                objUtility.ShowSuccessMessage("Report exported into PDF File")
            End If

        Else
            objUtility.ShowWarningMessage("No data found")
        End If

    End Sub
#End Region

    Private Sub openFileDialog()
        Try
            Dim objPL As New frmReportViewer
            objPL.iniViewer = AddressOf objPL.GenerateReport
            objPL.rptViewer.ReportSource = cryRpt
            objPL.rptViewer.Refresh()
            objPL.WindowState = FormWindowState.Maximized
            objPL.ShowDialog()
            System.Threading.Thread.Sleep(10 * 60)
            '  System.Threading.Thread.Sleep(200 * 60) ';    //replace with your real working code  
            objSBOAPI.SBO_Appln.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
            System.Threading.Thread.CurrentThread.Abort()
        Catch ex As Exception
            System.Threading.Thread.CurrentThread.Abort()
            objUtility.ShowErrorMessage(ex.Message)
        End Try
        
    End Sub
#Region "Selection Cretria"
    Private Sub SelectionCretria(ByVal aForm As SAPbouiCOM.Form)
        Dim strfrmSlp, strtoSlp, strAgeing, strCustomerType, strCondition, strdate, strproject, strtype, strgroup, strdisplay, strFromDate, strTOdate, strFromBP, strToBP, strFromProject, strToProject As String
        Dim strslpCondition As String
        Dim dtFromdate, dtTodate As Date
        Dim oCombo As SAPbouiCOM.ComboBox
        strCondition = ""
        strBPChoice = ""
        oCombo = aForm.Items.Item("35").Specific
        strCustomerType = oCombo.Selected.Value

        If objSBOAPI.GetEditText(aForm, "7") <> "" Then
            dtFromdate = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("FromDate").Value)
            strFromDate = objSBOAPI.GetSBODateString(dtFromdate)
        Else
            strFromDate = ""
        End If
        If objSBOAPI.GetEditText(aForm, "9") <> "" Then
            dtTodate = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("ToDate").Value)
            strTOdate = objSBOAPI.GetSBODateString(dtTodate)
        Else
            strTOdate = ""
        End If

        If objSBOAPI.GetEditText(aForm, "22") <> "" Then
            dtAgingdate = objSBOAPI.GetDateTimeValue(aForm.DataSources.UserDataSources.Item("Agedt").Value)
            strAgeing = objSBOAPI.GetSBODateString(dtTodate)
        Else
            strAgeing = ""
        End If

        strFromBP = objSBOAPI.GetEditText(aForm, "13")
        strToBP = objSBOAPI.GetEditText(aForm, "15")

        strfrmSlp = objSBOAPI.GetEditText(aForm, "29")
        strtoSlp = objSBOAPI.GetEditText(aForm, "32")

        oCombo = aForm.Items.Item("5").Specific
        strdate = ""
        Select Case oCombo.Selected.Value
            Case "P"
                strdate = "JDT1.RefDate"
                strReportFilterdate = "RefDate"
            Case "D"
                strdate = "JDT1.TaxDate"
                strReportFilterdate = "taxDate"
            Case "DU"
                strdate = "JDT1.DueDate"
                strReportFilterdate = "DueDate"
        End Select

        If strFromDate <> "" And strTOdate <> "" Then
            strJournalDatecondition = strdate & " Between '" & dtFromdate.ToString("yyyy-MM-dd") & "' and '" & dtTodate.ToString("yyyy-MM-dd") & "'"
        ElseIf strFromDate <> "" And strTOdate = "" Then
            strJournalDatecondition = strdate & " >= '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
        ElseIf strFromDate = "" And strTOdate <> "" Then
            strJournalDatecondition = strdate & " <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
        Else
            strJournalDatecondition = "1=1"
        End If


        oCombo = aForm.Items.Item("11").Specific
        strtype = ""
        strproject = ""
        blnProject = False
        Select Case oCombo.Selected.Value
            Case "C"
                strtype = "OCRD.CardType='C'"
                strBPChoice = "C"
            Case "S"
                strtype = "OCRD.CardType='S'"
                strBPChoice = "S"
            Case "P"
                strproject = "Yes"
                strtype = "JDT1.Project"
                strBPChoice = "P"
                blnProject = True
        End Select

        If strproject <> "" Then
            strBPCondition = " 2=2"
            If strFromBP <> "" And strToBP <> "" Then
                strProjectCondition = strtype & " between '" & strFromBP & "' and '" & strToBP & "'"
            ElseIf strFromBP <> "" And strToBP = "" Then
                strProjectCondition = strtype & ">='" & strFromBP & "'"
            ElseIf strToBP <> "" And strToBP = "" Then
                strProjectCondition = strtype & "<='" & strToBP & "'"
            Else
                strProjectCondition = strtype & " in (Select PrjCode from OPRJ) "
            End If
        Else
            strProjectCondition = " 2=2"
            If strFromBP <> "" And strToBP <> "" Then
                strBPCondition = strtype & " and OCRD.CardCode between '" & strFromBP & "' and '" & strToBP & "'"
                strSOBPCondition = "T1.CardType='C' and T1.CardCode between '" & strFromBP & "' and '" & strToBP & "'"
            ElseIf strFromBP <> "" And strToBP = "" Then
                strBPCondition = strtype & " and OCRD.CardCode >='" & strFromBP & "'"
                strSOBPCondition = "T1.CardType='C' and T1.CardCode>='" & strFromBP & "'"
            ElseIf strToBP <> "" And strToBP = "" Then
                strBPCondition = strtype & " and OCRD.CardCode <='" & strToBP & "'"
                strSOBPCondition = "T1.CardType='C' and T1.CardCode <='" & strToBP & "'"
            Else
                strBPCondition = strtype
                strSOBPCondition = "T1.CardType='C'"
            End If
        End If

        If strfrmSlp <> "" And strtoSlp <> "" Then
            strslpCondition = "  (select slpcode from OSLP where slpName between '" & strfrmSlp & "' and '" & strtoSlp & "')"
        ElseIf strfrmSlp <> "" And strtoSlp = "" Then
            strslpCondition = "  (select slpcode from OSLP where slpName = '" & strfrmSlp & "')"
        ElseIf strfrmSlp = "" And strtoSlp <> "" Then
            strslpCondition = "  (select slpcode from OSLP where slpName =' " & strtoSlp & "')"
        Else
            strslpCondition = "  (select slpcode from OSLP )"
        End If
        If strBPCondition <> " 2=2" Then
            strBPCondition = strBPCondition & " and OCRD.SlpCode in " & strslpCondition
            strSOBPCondition = strSOBPCondition & " and T1.Slpcode in " & strslpCondition

        End If


        oCombo = aForm.Items.Item("17").Specific
        strgroup = ""
        Select Case oCombo.Selected.Value
            Case "C"
                strgroup = "OCRD.groupcode"
            Case "S"
                strgroup = "OCRD.groupCode"
        End Select

        oCombo = aForm.Items.Item("18").Specific
        If oCombo.Selected.Value <> "" Then
            strBPCondition = strBPCondition & " and ( " & strgroup & " = '" & oCombo.Selected.Value & "')"
        End If

        oCombo = aForm.Items.Item("20").Specific
        strgroup = ""
        strReconcilationCondition = ""
        Select Case oCombo.Selected.Value
            Case "U"
                '  strgroup = " ( JDT1.ExtrMatch IS NULL OR isnull(JDT1.ExtrMatch,0)=0)"
                strgroup = " (isnull(JDT1.ExtrMatch,0)=0)"
                'strReconcilationCondition = "( A.ExtrMatch IS NULL OR isnull(A.ExtrMatch,0)=0)"
                strReconcilationCondition = "(isnull(A.ExtrMatch,0)=0)"

            Case "R"
                ' strgroup = " ( JDT1.ExtrMatch IS NOT NULL OR isnull(JDT1.ExtrMatch,0)<>0) "
                strgroup = " (isnull(JDT1.ExtrMatch,0)<>0) "
                'strReconcilationCondition = " ( A.ExtrMatch IS NOT NULL OR isnull(A.ExtrMatch,0)<>0) "
                strReconcilationCondition = "(Isnull(A.ExtrMatch,0)<>0)"
            Case "N"
                strgroup = " ( (JDT1.BalDueDeb+ JDT1.BalDueCred)<>0 )"
                strReconcilationCondition = " ( (A.BalDueDeb+ A.BalDueCred)<>0 )"
            Case "F"
                strgroup = " ( ( JDT1.BalDueDeb+ JDT1.BalDueCred)=0 )"
                strReconcilationCondition = " ( ( A.BalDueDeb+ A.BalDueCred)=0 )"
            Case "All"
                strgroup = " (3=3)"
                strReconcilationCondition = " (3=3)"
        End Select

        strAcctType = oCombo.Selected.Value
        strDisplayOption = oCombo.Selected.Description
        If strgroup = "" Then
            strgroup = " 3=3 "
            strReconcilationCondition = " (3=3)"
        End If

        'strReconcilationCondition = strgroup
        strJournalDatecondition = strJournalDatecondition & " and " & strgroup
        strAgeingCondition = strgroup
        strAgeingCondition = "3=3"
        'If strAgeing = "" Then
        '    strAgeingCondition = "3=3"
        'Else
        '    strAgeingCondition = strAgeing
        'End If

        oCombo = aForm.Items.Item("25").Specific
        strReportCurrency = ""
        Select Case oCombo.Selected.Value
            Case "L"
                strReportCurrency = "L"
            Case "S"
                strReportCurrency = "S"
            Case "B"
                strReportCurrency = "B"
        End Select

        oCombo = aForm.Items.Item("27").Specific
        If oCombo.Selected.Value <> "" Then
            strReportviewOption = oCombo.Selected.Value
        Else
            strReportviewOption = ""
        End If
    End Sub
#End Region

#End Region

#Region "Events"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : SBO_Appln_ItemEvent
    'Parameter          : FormUID, ItemEvent,BubbleEvent,Form
    'Return Value       : 
    'Author             : Dev8 
    'Created Date       :
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Handle Item Events
    '******************************************************************

    Public Sub SBO_Appln_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal oform As SAPbouiCOM.Form)
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "3" Then
                            oform = objSBOAPI.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                            Dim dtFr, dtTo As Date
                            Dim stFr, stTo As String
                            If objSBOAPI.GetEditText(oform, "7") <> "" Then
                                dtFr = objSBOAPI.GetDateTimeValue(oform.DataSources.UserDataSources.Item("FromDate").Value)
                                stFr = objSBOAPI.GetSBODateString(dtFr)
                            Else
                                stFr = ""
                            End If
                            If objSBOAPI.GetEditText(oform, "9") <> "" Then
                                dtTo = objSBOAPI.GetDateTimeValue(oform.DataSources.UserDataSources.Item("ToDate").Value)
                                stTo = objSBOAPI.GetSBODateString(dtTo)
                            Else
                                stTo = ""
                            End If
                            oStatic = oform.Items.Item("23").Specific
                            If Validation(oform) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                            'Dim oCheckbox As SAPbouiCOM.CheckBox
                            'oCheckbox = oform.Items.Item("43").Specific
                            'If oCheckbox.Checked = True Then
                            '    strShortname = "U_ShortName"
                            'Else
                            '    strShortname = "ShortName"
                            'End If
                            strShortname = "ShortName"
                            SelectionCretria(oform)
                            Dim blnBoolean As Boolean = False
                            For intRow As Integer = 0 To 10
                                Try
                                    If blnBoolean = True Then
                                        Exit Sub
                                    Else
                                        If blnProject = False Then
                                            blnBoolean = StatementofAccount(oform)
                                        Else
                                            StatementofAccount_Project(oform)
                                        End If
                                    End If
                                Catch ex As Exception
                                    blnBoolean = False
                                End Try
                            Next
                        End If
                End Select

                'For BP Master
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        oform = objSBOAPI.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If pVal.ItemUID = "9" And pVal.CharPressed = 9 Then
                            Dim oEdit1 As SAPbouiCOM.EditText
                            oEdit1 = oform.Items.Item("9").Specific
                            oEdit = oform.Items.Item("22").Specific
                            oEdit.String = oEdit1.String
                            oform.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        oform = objSBOAPI.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If pVal.ItemUID = "11" Then
                            changeBPCFL(oform)
                        ElseIf pVal.ItemUID = "17" Then
                            LoadGroupCombo(oform)
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Try
                            objForm = objSBOAPI.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                            Dim val1 As String
                            objForm = oform
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal
                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID
                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = objForm.ChooseFromLists.Item(sCFL_ID)
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                Dim val As String
                                val = oDataTable.GetValue(0, 0)
                                val1 = oDataTable.GetValue(1, 0)
                                If (pVal.ItemUID = "13") Then
                                    objForm.DataSources.UserDataSources.Item("FromBp").ValueEx = val
                                ElseIf (pVal.ItemUID = "15") Then
                                    objForm.DataSources.UserDataSources.Item("ToBp").ValueEx = val
                                ElseIf (pVal.ItemUID = "29") Then
                                    objForm.DataSources.UserDataSources.Item("frmSlp").ValueEx = val1
                                ElseIf (pVal.ItemUID = "32") Then
                                    objForm.DataSources.UserDataSources.Item("toSlp").ValueEx = val1

                                End If
                            End If
                        Catch ex As Exception
                            Dim x As String = ex.Message
                        End Try
                End Select
            End If
        Catch ex As Exception
            oStatic = oform.Items.Item("23").Specific
            oStatic.Caption = "Error Occured"
            objUtility.ShowMessage(ex.Message)
        End Try
    End Sub
    '*****************************************************************
    'Type               : Procedure    
    'Name               : SBO_Appln_MenuEvent
    'Parameter          : MenuEvent, BubbelEven
    'Return Value       : 
    'Author             : Dev8 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Handle Menu Events
    '******************************************************************
    Public Sub SBO_Appln_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If (pVal.BeforeAction = False) Then
            If (pVal.MenuUID = "DABT_701" And pVal.BeforeAction = False) Then
                objForm = objSBOAPI.SBO_Appln.Forms.ActiveForm
                LoadForm()
            End If
        End If
    End Sub
#End Region


End Class


