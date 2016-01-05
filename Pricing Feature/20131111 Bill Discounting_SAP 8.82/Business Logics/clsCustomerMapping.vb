Public Class clsCustomerMapping
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
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

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
    Private Sub AddChooseFromList(ByVal aForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = aForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL1
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()

            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "C"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()

          
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
#End Region

    Public Sub databind(ByVal aCode As String)
        Dim otemprec As SAPbobsCOM.Recordset
        Try

        
            otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemprec.DoQuery("Select * from [@Z_ODIS] where U_Z_Dis_Code='" & aCode & "'")
            If otemprec.RecordCount > 0 Then
                oApplication.Utilities.LoadForm(xml_CustMapping, frm_Customermapping)
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                ' oForm.EnableMenu(mnu_ADD_ROW, True)
                ' oForm.EnableMenu(mnu_DELETE_ROW, True)
                AddChooseFromList(oForm)
                If oForm.TypeEx = frm_Customermapping Then
                    oForm.Freeze(True)
                    oApplication.Utilities.SetEditText(oForm, "4", aCode)
                    oApplication.Utilities.SetEditText(oForm, "7", otemprec.Fields.Item("U_Z_Dis_Name").Value)
                    oGrid = oForm.Items.Item("8").Specific
                    oGrid.DataTable.ExecuteQuery("Select Code,U_Z_CardCode,U_Z_CardName,U_Z_FromDate,U_Z_Todate from [@Z_Dis_Mapping] where U_Z_Dis_Code='" & aCode & "'")
                    FormatGrid(oGrid)
                    oForm.Freeze(False)
                End If

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub Bindata(ByVal aCode As String)
        Dim otemprec As SAPbobsCOM.Recordset
        otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemprec.DoQuery("Select * from [@Z_ODIS] where U_Z_Dis_Code='" & aCode & "'")
        If otemprec.RecordCount > 0 Then
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If oForm.TypeEx = frm_Customermapping Then
                oForm.Freeze(True)
                oApplication.Utilities.SetEditText(oForm, "4", aCode)
                oApplication.Utilities.SetEditText(oForm, "7", otemprec.Fields.Item("U_Z_Dis_Name").Value)
                oGrid = oForm.Items.Item("8").Specific
                oGrid.DataTable.ExecuteQuery("Select Code,U_Z_CardCode,U_Z_CardName,U_Z_FromDate,U_Z_Todate from [@Z_Dis_Mapping] where U_Z_Dis_Code='" & aCode & "'")
                FormatGrid(oGrid)
                oForm.Freeze(False)
            End If
        End If
    End Sub
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        aGrid.Columns.Item(0).Visible = False
        aGrid.Columns.Item(1).TitleObject.Caption = "Customer Code"
        oEditTextColumn = aGrid.Columns.Item(1)
        oEditTextColumn.ChooseFromListUID = "CFL2"
        oEditTextColumn.ChooseFromListAlias = "CardCode"
        oEditTextColumn.LinkedObjectType = "2"
        aGrid.Columns.Item(2).TitleObject.Caption = "Customer Name"
        aGrid.Columns.Item(2).Editable = False
        aGrid.Columns.Item(3).TitleObject.Caption = "From Date"
        aGrid.Columns.Item(4).TitleObject.Caption = "To Date"
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        oForm.Freeze(True)
        If aGrid.DataTable.Rows.Count - 1 < 0 Then
            aGrid.DataTable.Rows.Add()
            aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
        Else
            If aGrid.DataTable.GetValue(1, aGrid.Rows.Count - 1) <> "" Then
                aGrid.DataTable.Rows.Add()
                aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
            End If
        End If
        oForm.Freeze(False)
    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemprs As SAPbobsCOM.Recordset
        Dim strDiscode, strDisname, StrFromdt, strTodt, strCode As String
        Dim strCardCode As String = ""
        Dim strCardName As String = ""
        Dim strLineCode As String = ""
        Dim FromDate, ToDate As Date
        oGrid = aform.Items.Item("8").Specific
        If validation(oGrid) = False Then
            Return False
        End If
        strDiscode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strDisname = oApplication.Utilities.getEdittextvalue(aform, "7")
      
        otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strCode = oGrid.DataTable.GetValue(0, intRow)
            oUserTable = oApplication.Company.UserTables.Item("Z_DIS_MAPPING")
            If strCode = "" Then
                strCode = oApplication.Utilities.getMaxCode("@Z_DIS_MAPPING", "Code")
                strCardCode = oGrid.DataTable.GetValue(1, intRow)
                strCardName = oGrid.DataTable.GetValue(2, intRow)
                oUserTable.Code = strCode
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_Dis_Code").Value = strDiscode
                oUserTable.UserFields.Fields.Item("U_Z_Dis_Name").Value = strDisname
                FromDate = oGrid.DataTable.GetValue(3, intRow)
                ToDate = oGrid.DataTable.GetValue(4, intRow)
                If Year(FromDate) <> 1899 Then
                    oUserTable.UserFields.Fields.Item("U_Z_FromDate").Value = FromDate
                End If
                If Year(FromDate) <> 1899 Then
                    oUserTable.UserFields.Fields.Item("U_Z_ToDate").Value = ToDate
                End If
                oUserTable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                oUserTable.UserFields.Fields.Item("U_Z_CardName").Value = strCardName
                If oUserTable.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Else
                If oUserTable.GetByKey(strCode) Then
                    strCardCode = oGrid.DataTable.GetValue(1, intRow)
                    strCardName = oGrid.DataTable.GetValue(2, intRow)
                    ' oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dis_Code").Value = strDiscode
                    oUserTable.UserFields.Fields.Item("U_Z_Dis_Name").Value = strDisname
                    FromDate = oGrid.DataTable.GetValue(3, intRow)
                    ToDate = oGrid.DataTable.GetValue(4, intRow)
                    If Year(FromDate) <> 1899 Then
                        oUserTable.UserFields.Fields.Item("U_Z_FromDate").Value = FromDate
                    End If
                    If Year(FromDate) <> 1899 Then
                        oUserTable.UserFields.Fields.Item("U_Z_ToDate").Value = ToDate
                    End If
                    oUserTable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                    oUserTable.UserFields.Fields.Item("U_Z_CardName").Value = strCardName
                    If oUserTable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        otemprs.DoQuery("Delete from [@Z_Dis_Mapping] where name like '%D' and U_Z_Dis_Code='" & strDiscode & "'")

        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Bindata(strDiscode)
        Return True
        'aform.Close()
    End Function
#End Region

#Region "Validation"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim dtFromdate, dtTodate As Date
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue(1, intRow) <> "" Then
                dtFromdate = oGrid.DataTable.GetValue(3, intRow)
                If dtFromdate.Year = "1" Then
                    oApplication.Utilities.Message("From date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item(3).Click(intRow, , 1)
                    Return False
                Else
                    dtFromdate = oGrid.DataTable.GetValue(3, intRow)
                End If
                dtTodate = oGrid.DataTable.GetValue(4, intRow)
                If dtTodate.Year = "1" Then
                    oApplication.Utilities.Message("To date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item(4).Click(intRow, , 1)
                    Return False
                Else
                    dtTodate = oGrid.DataTable.GetValue(4, intRow)
                    Dim strdate As Date
                    strdate = oGrid.DataTable.GetValue(4, intRow)
                End If
                If dtFromdate > dtTodate Then
                    oApplication.Utilities.Message("To date should be greater than from date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item(4).Click(intRow, , 1)
                    Return False
                End If
            End If
        Next
        Return True
    End Function
#End Region
#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Dim strCode, strname As String
        Dim otemprec As SAPbobsCOM.Recordset
        For intRow = 0 To agrid.DataTable.Rows.Count - 1
            If agrid.Rows.IsSelected(intRow) Then
                strCode = agrid.DataTable.GetValue(0, intRow)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oApplication.Utilities.ExecuteSQL(otemprec, "update [@Z_Dis_Mapping] set  Name =Name +'D'  where Code='" & strCode & "'")
                agrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Customermapping Then
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
                                    If AddtoUDT1(oForm) = True Then
                                        'oForm.Close()
                                    End If
                                End If

                                If pVal.ItemUID = "9" Then
                                    oGrid = oForm.Items.Item("8").Specific
                                    AddEmptyRow(oGrid)
                                End If
                                If pVal.ItemUID = "10" Then
                                    oGrid = oForm.Items.Item("8").Specific
                                    RemoveRow(1, oGrid)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val, val1, val3, val2 As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        If pVal.ItemUID = "8" And pVal.ColUID = "U_Z_CardCode" Then
                                            val2 = oDataTable.GetValue("CardCode", 0)
                                            val3 = oDataTable.GetValue("CardName", 0)
                                            oGrid = oForm.Items.Item("8").Specific
                                            oGrid.DataTable.SetValue(1, pVal.Row, val2)
                                            oGrid.DataTable.SetValue(2, pVal.Row, val3)
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
                ' Case mnu_InvSO
                'Case mnu_ADD_ROW
                '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                '    oGrid = oForm.Items.Item("8").Specific
                '    If pVal.BeforeAction = False Then
                '        AddEmptyRow(oGrid)
                '    End If

                'Case mnu_DELETE_ROW
                '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                '    oGrid = oForm.Items.Item("8").Specific
                '    If pVal.BeforeAction = True Then
                '        RemoveRow(1, oGrid)
                '        BubbleEvent = False
                '        Exit Sub
                '    End If
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
