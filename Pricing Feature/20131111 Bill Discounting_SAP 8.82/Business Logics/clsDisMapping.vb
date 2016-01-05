Public Class clsDisMapping
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
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
    Private oMenuobject As Object
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_Dis_mapping, frm_mapping)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        Databind(oForm)
    End Sub
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("11").Specific
            'dtTemp = oGrid.DataTable
            'dtTemp.ExecuteQuery("Select U_Z_ItemCode, U_Z_ItemName from View_1 where U_Z_ItemCode='A00003'")
            'oGrid.DataTable = dtTemp
            dtTemp = oGrid.DataTable
            dtTemp.ExecuteQuery("Select CardCode,CardName from OCRD where 1=2")
            oGrid.DataTable = dtTemp
            Formatgrid(oGrid)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item(0).TitleObject.Caption = "CardCode"
        oEditTextColumn = oGrid.Columns.Item(0)
        oEditTextColumn.LinkedObjectType = "2"
        agrid.Columns.Item(1).TitleObject.Caption = "CardName"
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        oForm.Freeze(True)
        If aGrid.DataTable.Rows.Count - 1 < 0 Then
            aGrid.DataTable.Rows.Add()
            aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
        Else
            If aGrid.DataTable.GetValue(0, aGrid.Rows.Count - 1) <> "" Then
                'If aGrid.DataTable.GetValue("U_Z_CardCode", aGrid.Rows.Count - 1) <> "" Then
                aGrid.DataTable.Rows.Add()
                aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
            End If
        End If
        oForm.Freeze(False)
    End Sub
#End Region

#Region "Validate Grid details"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strECode, strECode1, strEname, strEname1 As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strECode = aGrid.DataTable.GetValue(0, intRow)
            strEname = aGrid.DataTable.GetValue(1, intRow)
            If strECode = "" Then
                oApplication.Utilities.Message("CardCode is missing .....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        Next
        Return True
    End Function

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

        oGrid = aform.Items.Item("11").Specific
        If validation(oGrid) = False Then
            Return False
        End If
        strDiscode = oApplication.Utilities.getEdittextvalue(aform, "4")
        strDisname = oApplication.Utilities.getEdittextvalue(aform, "6")
        StrFromdt = oApplication.Utilities.getEdittextvalue(aform, "8")
        FromDate = oApplication.Utilities.GetDateTimeValue(StrFromdt)
        strTodt = oApplication.Utilities.getEdittextvalue(aform, "10")
        ToDate = oApplication.Utilities.GetDateTimeValue(strTodt)
        otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue(0, intRow) <> "" Or oGrid.DataTable.GetValue(1, intRow) <> "" Then
                otemprs.DoQuery("select * from [@Z_DIS_MAPPING]  WHERE U_Z_CardCode='" & oGrid.DataTable.GetValue(0, intRow) & "' and  [U_Z_Dis_Code] ='" & strDiscode & "' and [U_Z_FromDate]='" & FromDate.ToString("yyyy-MM-dd") & "' and [U_Z_ToDate]='" & ToDate.ToString("yyyy-MM-dd") & "'")
                If otemprs.RecordCount > 0 Then
                    ' oApplication.Utilities.Message("Already Exists for this combination of data...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ' Exit Function
                Else
                    strCode = oApplication.Utilities.getMaxCode("@Z_DIS_MAPPING", "Code")
                    strCardCode = oGrid.DataTable.GetValue(0, intRow)
                    strCardName = oGrid.DataTable.GetValue(1, intRow)
                    oUserTable = oApplication.Company.UserTables.Item("Z_DIS_MAPPING")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dis_Code").Value = strDiscode
                    oUserTable.UserFields.Fields.Item("U_Z_Dis_Name").Value = strDisname
                    If StrFromdt <> "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_FromDate").Value = FromDate
                    End If
                    If strTodt <> "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_ToDate").Value = ToDate
                    End If
                    oUserTable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                    oUserTable.UserFields.Fields.Item("U_Z_CardName").Value = strCardName
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        aform.Close()
    End Function
#End Region

#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Dim strCode, strname As String
        Dim otemprec As SAPbobsCOM.Recordset
        For intRow = 0 To agrid.DataTable.Rows.Count - 1
            If agrid.Rows.IsSelected(intRow) Then
                strCode = agrid.DataTable.GetValue(0, intRow)
                strname = agrid.DataTable.GetValue(1, intRow)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'otemprec.DoQuery("Select * from [@Z_PAY_OCON] where Code='" & strCode & "' and Name='" & strname & "'")
                'If otemprec.RecordCount > 0 And strCode <> "" Then
                '    oApplication.Utilities.Message("Transaction already exists. Can not delete the Bin Details.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Exit Sub
                'End If
                'oApplication.Utilities.ExecuteSQL(oTemp, "update [@Z_PAY_OCON] set  Name =Name +'D'  where Code='" & strCode & "'")
                agrid.DataTable.Rows.Remove(intRow)
                Exit For
            End If
        Next
        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_mapping Then
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
                                If pVal.ItemUID = "15" Then
                                    AddtoUDT1(oForm)
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
                                        If pVal.ItemUID = "4" Then
                                            val2 = oDataTable.GetValue("U_Z_Dis_Code", 0)
                                            val3 = oDataTable.GetValue("U_Z_Dis_Name", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val3)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val2)
                                        End If
                                        If pVal.ItemUID = "11" Then
                                            oGrid = oForm.Items.Item("11").Specific
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            Try

                                                oGrid.DataTable.SetValue(1, pVal.Row, val1)
                                                oGrid.DataTable.SetValue(0, pVal.Row, val)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
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
                Case mnu_Dismap
                    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_mapping Then
                        oGrid = oForm.Items.Item("11").Specific
                        If pVal.BeforeAction = False Then
                            AddEmptyRow(oGrid)
                        End If
                    End If
                    

                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_mapping Then
                        oGrid = oForm.Items.Item("11").Specific
                        If pVal.BeforeAction = True Then
                            RemoveRow(1, oGrid)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
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
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_Dismap
                        oMenuobject = New clsDisMapping
                        oMenuobject.MenuEvent(pVal, BubbleEvent)


                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
