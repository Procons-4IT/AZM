Public Class clsDisDefine
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private oColumn As SAPbouiCOM.Column
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private MatrixId As String
    Private RowtoDelete As Integer
    Private InvBaseDocNo As String
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_DisDefine, frm_SpecialPrice)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "4"
        AddMode(oForm)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
       
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        FillPricelist(oForm)

        databind(oForm)


        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Freeze(False)
    End Sub
#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String
        strCode = oApplication.Utilities.getMaxCode("@Z_ODIS", "DocEntry")
        oApplication.Utilities.setEdittextvalue(aForm, "4", strCode)

    End Sub
#End Region
#Region "Validations"
    Private Function Validations(ByVal aform As SAPbouiCOM.Form) As Boolean

        'oForm.Freeze(True)

       

        Return True
        'oForm.Freeze(False)
    End Function
#End Region
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("9").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("9").Specific
            oColumn = oMatrix.Columns.Item("V_0")
            oColumn.ChooseFromListUID = "CFL1"
            oColumn.ChooseFromListAlias = "ItemCode"

            oColumn = oMatrix.Columns.Item("V_4")
            'oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oColumn = oMatrix.Columns.Item("V_3")
            ' oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            'oColumn = oMatrix.Columns.Item("V_5")
            'Dim oComborec As SAPbobsCOM.Recordset
            'oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oComborec.DoQuery("SELECT T0.[ListNum], T0.[ListName] FROM OPLN T0 order by T0.[ListNum]  ")
            'For introw As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
            '    oColumn.ValidValues.Remove(introw)
            'Next
            'oColumn.ValidValues.Add("", "")
            'For introw As Integer = 0 To oComborec.RecordCount - 1
            '    Try
            '        oColumn.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            '    Catch ex As Exception
            '    End Try

            '    oComborec.MoveNext()
            'Next
            'oColumn.ValidValues.Add("0", "Without Price list")
            'oColumn.DisplayDesc = True

            oMatrix.AutoResizeColumns()


            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("9").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
            count = 0
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
                oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                Try
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", oMatrix.RowCount, "")
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, oApplication.Utilities.GetEditText(aForm, "17"))
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
                    'oCombobox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                Catch ex As Exception
                End Try
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
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, oApplication.Utilities.GetEditText(aForm, "17"))
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, "")
                    'oCombobox = oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific
                    'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                End If
            Catch ex As Exception
            End Try
            oMatrix.FlushToDataSource()
            oMatrix = aForm.Items.Item("9").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#End Region
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId = "9" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
            'Else
            '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        End If
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
        If intSelectedMatrixrow <= 0 Then
            Exit Sub
        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
        End If
    End Sub
#End Region

    Private Sub DeleteRow(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        oMatrix = aform.Items.Item("9").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_DIS1")
        Dim intRow As Integer
        For intRow = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                oMatrix.DeleteRow(intRow)
                AddRow(aform)
                aform.Freeze(False)
                Exit Sub
            End If
        Next
        AssignLineNo(aform)
        aform.Freeze(False)
    End Sub
#Region "Fill Project Code"
    Private Sub FillPricelist(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = oForm.Items.Item("9").Specific
        Try
            oCombobox = oMatrix.Columns.Item("V_5").Cells.Item(1).Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("SELECT T0.[ListNum], T0.[ListName] FROM OPLN T0 order by T0.[ListNum]")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("ListNum").Value, oTempRec.Fields.Item("ListName").Value)
                oTempRec.MoveNext()
            Next
            oCombobox.ValidValues.Add("0", "Without Price list")
        Catch ex As Exception

        End Try
       
    End Sub
#End Region



#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strActivity, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strProject = oApplication.Utilities.getEdittextvalue(aform, "6")
            If strProject = "" Then
                oApplication.Utilities.Message("Discount Code is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                oTemp.DoQuery("Select * from [@Z_ODIS] where U_Z_Dis_Code='" & strProject & "'")
                If oTemp.RecordCount > 0 Then
                    oApplication.Utilities.Message("Project code already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            If oApplication.Utilities.getEdittextvalue(aform, "8") = "" Then
                oApplication.Utilities.Message("Discount Name is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If

        'oTemp.DoQuery("Select * From [@Z_ODIS] where isnull(U_Z_Default,'N')='Y' and docEntry <> " & oApplication.Utilities.GetEditText(aform, "4"))
        'If oTemp.RecordCount > 0 Then
        '    oCombobox = aform.Items.Item("15").Specific
        '    Dim strValue As String
        '    Try
        '        strValue = oCombobox.Selected.Value
        '    Catch ex As Exception
        '        strValue = ""
        '    End Try
        '    If strValue = "Y" Then
        '        oApplication.Utilities.Message("Already another price list is marked as default", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '        Return False
        '    End If
        'End If

        oMatrix = aform.Items.Item("9").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Line details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oMatrix.Columns.Item("V_0").Cells.Item(1).Specific.value = "" Then
            oApplication.Utilities.Message("Line details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Dim stritemcode, stritemcode1, strUOM, strBasedOn, strdefaultpricelist As String
        strBasedOn = oApplication.Utilities.GetEditText(aform, "17")
        oCombobox = aform.Items.Item("15").Specific
        strdefaultpricelist = oCombobox.Selected.Value

        For intRow As Integer = 1 To oMatrix.RowCount
            stritemcode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
            oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If stritemcode <> "" Then
                If strdefaultpricelist = "N" Then
                    If oApplication.Utilities.getMatrixValues(oMatrix, "V_5", intRow) = "" Then
                        oApplication.Utilities.Message("Price list missingg...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.Columns.Item("V_5").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                        Return False
                    End If
                End If
                
                strUOM = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", intRow)
                If strUOM = "" Then
                    oApplication.Utilities.Message("Alternative Number of Pieces should be greater than zero : Line no: " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_4").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                ElseIf CDbl(strUOM) <= 0 Then
                    oApplication.Utilities.Message("Alternative Number of Pieces should be greater than zero : Line no: " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_4").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                End If
                strUOM = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intRow)
                If strUOM = "" Then
                    oApplication.Utilities.Message("Alt.Unit Price per Carton should be greater than zero : Line no: " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_3").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                ElseIf CDbl(strUOM) <= 0 Then
                    oApplication.Utilities.Message("Alt.Unit Price per Carton  should be greater than zero : Line no: " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_3").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                End If
                If oApplication.Utilities.getMatrixValues(oMatrix, "V_10", intRow) = "" Then
                    oApplication.Utilities.Message("Currency missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_10").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                End If
                'For intLoop As Integer = intRow To oMatrix.RowCount
                '    stritemcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intLoop)
                '    If stritemcode1 <> "" Then
                '        If stritemcode = stritemcode1 And intRow <> intLoop Then
                '            oApplication.Utilities.Message("Item Code already exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            oMatrix.Columns.Item("V_0").Cells.Item(intLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                '            Return False
                '        End If
                '    End If
                'Next
            End If

        Next

        ' MsgBox(oMatrix.Columns.Item("V_0").Cells.Item(1).Specific.value=""
        Return True
    End Function
#End Region

#Region "Get Price"
    Private Sub getPrice(ByVal aItemCode As String, ByVal aRow As Integer, ByVal aform As SAPbouiCOM.Form)
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblPrice, dblDefPack, dblFactor As Double
        Dim intpricelist As Integer
        Dim strPriceList, strHeaderPriceList, strfactor, strCurrency As String
        Try
            aform.Freeze(True)
            dblPrice = 0
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aform.Items.Item("9").Specific

            strHeaderPriceList = oApplication.Utilities.GetEditText(aform, "17")
            strfactor = oApplication.Utilities.GetEditText(aform, "19")

            strPriceList = oApplication.Utilities.getMatrixValues(oMatrix, "V_5", aRow)
            strCurrency = ""
            If strPriceList = "" Then
                oTemp.DoQuery("Select 0,isnull(SalPackUn,1),'' from OITM where ItemCode='" & aItemCode & "'")
            Else
                oTemp.DoQuery("Select isnull(U_Z_Price,0),U_Z_No_Pices,U_Z_Currency from [@Z_DIS1] where U_Z_Itemcode='" & aItemCode & "' and DocEntry =(Select DocEntry from [@Z_ODIS] where U_Z_DIS_CODE='" & strPriceList & "')")
            End If
            If oTemp.RecordCount <= 0 Then
                oTemp.DoQuery("Select 0,isnull(SalPackUn,1),'' from OITM where ItemCode='" & aItemCode & "'")
            Else
                strCurrency = oTemp.Fields.Item(2).Value
            End If
            strCurrency = oTemp.Fields.Item(2).Value
            dblPrice = oTemp.Fields.Item(0).Value
            dblDefPack = oTemp.Fields.Item(1).Value 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_8", aRow))
            'oMatrix.Columns.Item("V_10").Cells.Item(aRow).Specific.value = strCurrency.Trim
            oMatrix.Columns.Item("V_7").Cells.Item(aRow).Specific.value = dblPrice
            oMatrix.Columns.Item("V_8").Cells.Item(aRow).Specific.value = dblDefPack
            If strHeaderPriceList <> "" Then
                If strfactor = "" Then
                    dblFactor = 0
                Else
                    dblFactor = CDbl(strfactor)
                End If
            Else
                dblFactor = 1
            End If
            dblPrice = (dblPrice * dblFactor)
            oMatrix.Columns.Item("V_3").Cells.Item(aRow).Specific.value = dblPrice
            oMatrix.Columns.Item("V_10").Cells.Item(aRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            oMatrix.Columns.Item("V_4").Cells.Item(aRow).Specific.value = dblDefPack
            oMatrix.Columns.Item("V_5").Cells.Item(aRow).Specific.value = strPriceList
            Try
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", aRow, strCurrency)
            Catch ex As Exception

            End Try

            'oMatrix.Columns.Item("V_10").Cells.Item(aRow).Specific.value = strCurrency
            ' oMatrix.Columns.Item("V_4").Cells.Item(aRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
            'oMatrix.Columns.Item("V_7").Editable = False
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub PopulatePrice(ByVal aform As SAPbouiCOM.Form)
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblPrice, dblDefPack, dblFactor As Double
        Dim intpricelist As Integer
        Dim strPriceList, strHeaderPriceList, strfactor, aItemCode, aCurrency As String
        Try
            aform.Freeze(True)
            dblPrice = 0
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aform.Items.Item("9").Specific
            strHeaderPriceList = oApplication.Utilities.GetEditText(aform, "17")
            strfactor = oApplication.Utilities.GetEditText(aform, "19")

            '  strPriceList = oApplication.Utilities.getMatrixValues(oMatrix, "V_5", aRow)
            aCurrency = oApplication.Utilities.GetLocalCurrency()
            If strPriceList = "" Then

                oTemp.DoQuery("Select 0,isnull(SalPackUn,1) from OITM where ItemCode='" & aItemCode & "'")
            Else
                oTemp.DoQuery("Select isnull(U_Z_Price,0),U_Z_No_Pices,U_Z_Currency from [@Z_DIS1] where U_Z_Itemcode='" & aItemCode & "' and DocEntry =(Select DocEntry from [@Z_ODIS] where U_Z_DIS_CODE='" & strPriceList & "')")
            End If

            For aRow As Integer = 1 To oMatrix.RowCount
                aItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", aRow)
                If aItemCode <> "" Then
                    oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strPriceList = strHeaderPriceList 'oApplication.Utilities.getMatrixValues(oMatrix, "V_5", aRow)
                    If strPriceList = "" Then
                        oTemp.DoQuery("Select 0,isnull(SalPackUn,1),'' from OITM where ItemCode='" & aItemCode & "'")
                    Else
                        oTemp.DoQuery("Select isnull(U_Z_Price,0),U_Z_No_Pices,U_Z_Currency from [@Z_DIS1] where U_Z_Itemcode='" & aItemCode & "' and DocEntry =(Select DocEntry from [@Z_ODIS] where U_Z_DIS_CODE='" & strPriceList & "')")
                    End If
                    If oTemp.RecordCount <= 0 Then
                        oTemp.DoQuery("Select 0,isnull(SalPackUn,1),'' from OITM where ItemCode='" & aItemCode & "'")
                    Else
                        aCurrency = oTemp.Fields.Item(2).Value
                    End If
                    aCurrency = oTemp.Fields.Item(2).Value
                    'oTemp.DoQuery("Select isnull(U_Z_Price,0),U_Z_No_Pices from [@Z_DIS1] where U_Z_Itemcode='" & aItemCode & "' and DocEntry =(Select DocEntry from [@Z_ODIS] where U_Z_DIS_CODE='" & strPriceList & "')")
                    dblPrice = oTemp.Fields.Item(0).Value
                    dblDefPack = oTemp.Fields.Item(1).Value 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_8", aRow))
                    oMatrix.Columns.Item("V_7").Cells.Item(aRow).Specific.value = dblPrice
                    oMatrix.Columns.Item("V_8").Cells.Item(aRow).Specific.value = dblDefPack
                    If strHeaderPriceList <> "" Then
                        If strfactor = "" Then
                            dblFactor = 0
                        Else
                            dblFactor = CDbl(strfactor)
                        End If
                    Else
                        dblFactor = 1
                    End If
                    dblPrice = (dblPrice * dblFactor)
                    oMatrix.Columns.Item("V_3").Cells.Item(aRow).Specific.value = dblPrice
                    oMatrix.Columns.Item("V_4").Cells.Item(aRow).Specific.value = dblDefPack
                    oMatrix.Columns.Item("V_5").Cells.Item(aRow).Specific.value = strPriceList
                    oMatrix.Columns.Item("V_10").Cells.Item(aRow).Specific.value = aCurrency
                End If
            Next
            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub calculateUnitprice(ByVal aRow As Integer, ByVal aform As SAPbouiCOM.Form)
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblPrice, dblSellingPrice As Double
        Dim intpricelist As Integer
        Try
            aform.Freeze(True)
            dblPrice = 0
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aform.Items.Item("9").Specific


            dblSellingPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_7", aRow))
            dblPrice = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", aRow)
            dblPrice = dblSellingPrice / dblPrice

            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub UnitPriceCalculation(ByVal aform As SAPbouiCOM.Form)
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblPrice, dblSellingPrice As Double
        Dim intpricelist As Integer
        Dim aItemCode As String
        Try
            aform.Freeze(True)
            oMatrix = aform.Items.Item("9").Specific
            For arow As Integer = 1 To oMatrix.RowCount
                dblPrice = 0
                aItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", arow)
                If aItemCode <> "" Then
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oCombobox = oMatrix.Columns.Item("V_5").Cells.Item(arow).Specific
                    'Try
                    '    intpricelist = CInt(oCombobox.Selected.Value)
                    'Catch ex As Exception
                    '    intpricelist = 1
                    'End Try
                    'oTemp.DoQuery("Select isnull(Price,0) from ITM1 where Itemcode='" & aItemCode & "' and PriceList=" & intpricelist)
                    'dblPrice = oTemp.Fields.Item(0).Value
                    '' oMatrix.Columns.Item("V_7").Cells.Item(arow).Specific.value = dblPrice
                    'dblSellingPrice = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_7", arow))
                    'dblPrice = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", arow)
                    'dblPrice = dblSellingPrice / dblPrice
                    ' oMatrix.Columns.Item("V_7").Cells.Item(arow).Specific.value = dblPrice
                End If
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SpecialPrice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("9").Specific
                                If pVal.ItemUID = "9" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "9"
                                    frmSourceMatrix = oMatrix
                                ElseIf (pVal.ItemUID = "17" Or pVal.ItemUID = "19") Then
                                    oCombobox = oForm.Items.Item("15").Specific
                                    If oCombobox.Selected.Value = "N" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If pVal.ItemUID = "19" Then
                                            If oApplication.Utilities.GetEditText(oForm, "17") = "" Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If (pVal.ItemUID = "17" Or pVal.ItemUID = "19") And pVal.CharPressed <> 9 Then
                                    oCombobox = oForm.Items.Item("15").Specific
                                    If oCombobox.Selected.Value = "N" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    If pVal.ItemUID = "19" Then
                                        If oApplication.Utilities.GetEditText(oForm, "17") = "" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
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
                                oMatrix = oForm.Items.Item("9").Specific
                                If pVal.ItemUID = "9" And (pVal.ColUID = "V_7" Or pVal.ColUID = "V_4") And pVal.CharPressed = 9 Then
                                    calculateUnitprice(pVal.Row, oForm)
                                    If pVal.ColUID = "V_7" Then
                                        oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                    Else
                                        oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                    End If
                                End If
                                If pVal.ItemUID = "19" And pVal.CharPressed = 9 Then
                                    If oMatrix.RowCount > 0 Then
                                        Dim stPrice As String
                                        stPrice = oApplication.Utilities.GetEditText(oForm, "17")
                                        If stPrice <> "" Then
                                            stPrice = oApplication.Utilities.GetEditText(oForm, "19")
                                            If CDbl(stPrice) > 0 Then
                                                If oApplication.SBO_Application.MessageBox("Do you want to recalcualte the  Alt.price per Carton.?", , "Yes", "No") = 1 Then
                                                    PopulatePrice(oForm)
                                                End If
                                            End If
                                        End If
                                        
                                    End If
                                End If
                                If pVal.ItemUID = "9" And pVal.ColUID = "V_5" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strbank, strGirdValue As String
                                    strGirdValue = oApplication.Utilities.getMatrixValues(oMatrix, pVal.ColUID, pVal.Row)
                                    Dim otemp As SAPbobsCOM.Recordset
                                    otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otemp.DoQuery("Select * from [@Z_ODIS] where U_Z_DIS_CODE='" & strGirdValue & "'")
                                    If otemp.RecordCount > 0 Then
                                        oMatrix = oForm.Items.Item("9").Specific
                                        getPrice(oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row), pVal.Row, oForm)
                                        oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                                        Exit Sub
                                    Else
                                        strbank = oApplication.Utilities.GetEditText(oForm, "6")
                                    End If
                                    If strbank <> "" Then
                                        clsChooseFromList.ItemUID = pVal.ItemUID
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.CFLChoice = "ROW" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "PriceList"
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

                                If pVal.ItemUID = "17" And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strbank, strGirdValue As String
                                    strGirdValue = oApplication.Utilities.GetEditText(oForm, pVal.ItemUID)
                                    Dim otemp As SAPbobsCOM.Recordset
                                    otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otemp.DoQuery("Select * from [@Z_ODIS] where U_Z_DIS_CODE='" & strGirdValue & "'")
                                    If otemp.RecordCount > 0 Then
                                        Exit Sub
                                    Else
                                        strbank = oApplication.Utilities.GetEditText(oForm, "6")
                                    End If
                                    If strbank <> "" Then
                                        clsChooseFromList.ItemUID = pVal.ItemUID
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = 0
                                        clsChooseFromList.CFLChoice = "Header" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "PriceList"
                                        clsChooseFromList.ItemCode = strbank
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = 0 ' pVal.ColUID
                                        clsChooseFromList.sourcerowId = 0 'pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" Then
                                    oCombobox = oForm.Items.Item("15").Specific
                                    If oCombobox.Selected.Value = "N" Then
                                        oApplication.Utilities.SetEditText(oForm, "17", "")
                                        oApplication.Utilities.SetEditText(oForm, "19", "0")
                                    Else
                                        oApplication.Utilities.SetEditText(oForm, "17", "")
                                        oApplication.Utilities.SetEditText(oForm, "19", "1")
                                    End If
                                    'getPrice(oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row), pVal.Row, oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) Then
                                    Dim oBj As New clsCustomerMapping
                                    oBj.databind(oApplication.Utilities.GetEditText(oForm, "6"))
                                End If
                                If pVal.ItemUID = "11" Then
                                    AddRow(oForm)
                                End If
                                If pVal.ItemUID = "12" Then
                                    ' DeleteRow(oForm)
                                    oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End If
                                If pVal.ItemUID = "btnItem" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    frmSourceSpecialPriceForm = oForm
                                    Dim oObject As New clsItemCFL()
                                    oObject.LoadForm()
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
                                        If pVal.ItemUID = "9" And pVal.ColUID = "V_0" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            oMatrix = oForm.Items.Item("9").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", pVal.Row, val1)
                                            ' getPrice(val, pVal.Row, oForm)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", pVal.Row, oDataTable.GetValue("SalPackUn", 0))
                                            getPrice(val, pVal.Row, oForm)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        End If
                                        If pVal.ItemUID = "9" And pVal.ColUID = "V_10" Then
                                            val = oDataTable.GetValue("CurrCode", 0)
                                            oMatrix = oForm.Items.Item("9").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_10", pVal.Row, val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                    End If
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
                Case mnu_DisDefin
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    AddRow(oForm)
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        'DeleteRow(oForm)
                        RefereshDeleteRow(oForm)
                        BubbleEvent = False
                        Exit Sub
                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    AddMode(oForm)
                Case mnu_DELETE_ROW

                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("8").Enabled = True
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry As String
        Try
            aForm.Freeze(True)
            UnitPriceCalculation(aForm)
            If validation(aForm) = False Then
                aForm.Freeze(False)
                Return False
            End If
            AssignLineNo(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False

        End Try
        
        Return True
    End Function
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
