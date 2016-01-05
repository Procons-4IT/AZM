Public Class clsItemCFL
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
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_ItemCFL, frm_ItemCFL)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oCombobox = oForm.Items.Item("9").Specific
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select * from OITB order by Itmsgrpcod")
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To oTemp.RecordCount - 1
            oCombobox.ValidValues.Add(oTemp.Fields.Item("ItmsGrpCod").Value, oTemp.Fields.Item("ItmsGrpnam").Value)
            oTemp.MoveNext()
        Next
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("9").DisplayDesc = True

        oCombobox = oForm.Items.Item("13").Specific
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select * from OITG")
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To oTemp.RecordCount - 1
            oCombobox.ValidValues.Add(oTemp.Fields.Item(0).Value, oTemp.Fields.Item(1).Value)
            oTemp.MoveNext()
        Next
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

        oForm.Items.Item("13").DisplayDesc = True
        oForm.Freeze(False)
    End Sub
    Private Function BindSelectedItems(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strcondition, strFromitemcode, strToItemCode, strItemgroup, strProperty As String
        Try
            aForm.Freeze(True)
            strFromitemcode = oApplication.Utilities.GetEditText(aForm, "5")
            strToItemCode = oApplication.Utilities.GetEditText(aForm, "7")
            oCombobox = aForm.Items.Item("9").Specific
            strItemgroup = oCombobox.Selected.Value
            oCombobox = aForm.Items.Item("13").Specific
            strProperty = oCombobox.Selected.Value

            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strFromitemcode = "" Then
                strcondition = " 1=1"
            Else
                strcondition = " ItemCode >='" & strFromitemcode & "'"
            End If

            If strToItemCode = "" Then
                strcondition = strcondition & " and   1=1"
            Else
                strcondition = strcondition & " and  ItemCode <='" & strToItemCode & "'"
            End If

            If strItemgroup = "" Then
                strcondition = strcondition & " and   1=1"
            Else
                strcondition = strcondition & " and   ItmsGrpCod =" & CInt(strItemgroup)
            End If

            If strProperty = "" Then
                strcondition = strcondition & " and   1=1"
            Else
                strcondition = strcondition & " and   QryGroup" & CInt(strProperty) & "='Y'"
            End If
            oMatrix = frmSourceSpecialPriceForm.Items.Item("9").Specific
            Dim strExistingItem, strItem As String
            strExistingItem = ""
            For intRow As Integer = 1 To oMatrix.RowCount
                strItem = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                If strItem <> "" Then
                    If strExistingItem = "" Then
                        strExistingItem = "'" & strItem & "'"
                    Else
                        strExistingItem = strExistingItem & ",'" & strItem & "'"
                    End If
                End If
            Next
            If strExistingItem <> "" Then
                strcondition = strcondition & " and ItemCode not in (" & strExistingItem & ")"

            End If
            oTemp.DoQuery("Select * from OITM where " & strcondition)
            Try
                frmSourceSpecialPriceForm.Freeze(True)
                For introw As Integer = 0 To oTemp.RecordCount - 1
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oMatrix = frmSourceSpecialPriceForm.Items.Item("9").Specific
                    oDataSrc_Line = frmSourceSpecialPriceForm.DataSources.DBDataSources.Item("@Z_DIS1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", oMatrix.RowCount, "")
                            ' oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, oApplication.Utilities.GetEditText(aForm, "17"))
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
                            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    Catch ex As Exception
                    End Try
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, oApplication.Utilities.GetEditText(frmSourceSpecialPriceForm, "17"))
                    '  MsgBox(oTemp.Fields.Item(0).Value)
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, oTemp.Fields.Item(0).Value)
                    oTemp.MoveNext()
                Next
                oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oMatrix.FlushToDataSource()
                oMatrix = frmSourceSpecialPriceForm.Items.Item("9").Specific
                oDataSrc_Line = frmSourceSpecialPriceForm.DataSources.DBDataSources.Item("@Z_DIS1")
                For count As Integer = 1 To oDataSrc_Line.Size
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
                frmSourceSpecialPriceForm.Freeze(False)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                frmSourceSpecialPriceForm.Freeze(False)
            End Try
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try


    End Function

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
       

    End Sub


#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ItemCFL Then
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
                                    If BindSelectedItems(oForm) = True Then
                                        oForm.Close()
                                    End If
                                End If

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
                                        If pVal.ItemUID = "5" Or pVal.ItemUID = "7" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            oApplication.Utilities.SetEditText(oForm, pVal.ItemUID, val)



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
