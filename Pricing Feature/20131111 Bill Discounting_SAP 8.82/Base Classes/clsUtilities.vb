Imports System.IO
Imports System.ComponentModel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient



Public Class clsUtilities

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

    Public Function GetLocalCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Maincurncy from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function GetSystemCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select SysCurrncy from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function getBPCurrency(ByVal strCardcode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Currency from OCRD where Cardcode='" & strCardcode & "'"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

#Region "Add to Import UDT"
    Public Sub AddtoExportUDT(ByVal strCode As String, ByVal strMastercode As String, ByVal strchoice As String, ByVal transType As String)
        Try
            Dim oUsertable As SAPbobsCOM.UserTable
            Dim strsql, sCode, strUpdateQuery As String
            Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("Select * from [@Z_DABT_EXPORT] where U_Z_DocType='" & strchoice & "' and U_Z_MasterCode='" & strCode & "' and U_Z_Exported='N'")
            If oRec.RecordCount <= 0 Then
                strsql = getMaxCode("@Z_DABT_EXPORT", "CODE")
                oUsertable = oApplication.Company.UserTables.Item("Z_DABT_EXPORT")
                oUsertable.Code = strsql
                oUsertable.Name = strsql & "M"
                oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = strchoice
                oUsertable.UserFields.Fields.Item("U_Z_MasterCode").Value = strCode
                oUsertable.UserFields.Fields.Item("U_Z_DocNum").Value = strMastercode
                oUsertable.UserFields.Fields.Item("U_Z_Action").Value = transType 'strAction '"A"
                oUsertable.UserFields.Fields.Item("U_Z_CreateDate").Value = Now.Date
                oUsertable.UserFields.Fields.Item("U_Z_CreateTime").Value = Now.ToShortTimeString.Replace(":", "")
                oUsertable.UserFields.Fields.Item("U_Z_Exported").Value = "N"
                If oUsertable.Add <> 0 Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub




#End Region

#Region "Add Controls"

    '*****************************************************************
    'Type               : Procedure   
    'Name               : addControls
    'Parameter          : StrCode
    'Return Value       : string
    'Author             : Senthil Kumar B
    'Created Date       : 03-07-2009
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create Controls in the SAP B1 Screens
    '*****************************************************************
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal Left As Integer = 0, Optional ByVal Top As Integer = 0, Optional ByVal width As Integer = 0)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "DOWN" Then
                    .Top = objOldItem.Top + objOldItem.Height + 1
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "TOP" Then
                    .Top = objOldItem.Top - objOldItem.Height - 3
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "LEFT" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left - objOldItem.Width - 20
                End If
            End If
            If Left <> 0 Then
                objNewItem.Left = objOldItem.Left
            End If
            If Top <> 0 Then
                objNewItem.Top = objOldItem.Top
            End If
            '.FromPane = fromPane
            '.ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            ' .ForeColor = 255
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            If ItemUID = "btnDisplay" Then
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 20
            Else
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            End If
            If width <> 0 Then
                objNewItem.Width = width
            End If
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption
        End If
    End Sub

    Public Sub addDisplayControl(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal TopID As String, ByVal LeftID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal Width As Integer = 0)
        Dim objNewItem, objOldItem, objLeftItem, objTopItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        objOldItem = objForm.Items.Item(SourceUID)
        objLeftItem = objForm.Items.Item(LeftID)
        objTopItem = objForm.Items.Item(TopID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "DOWN" Then
                    .Top = objOldItem.Top + objOldItem.Height + 1
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "TOP" Then
                    .Top = objOldItem.Top - objOldItem.Height - 3
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "LEFT" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left - objOldItem.Width - 20
                End If
            End If
            objNewItem.Top = objTopItem.Top
            objNewItem.Left = objLeftItem.Left
            '.FromPane = fromPane
            '.ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            ' .ForeColor = 255
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            If ItemUID = "btnDiscount" Then
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 20
            Else
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            End If
            If Width <> 0 Then
                objNewItem.Width = Width
            End If
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption
        End If
    End Sub
#End Region
#Region "Export Documents"
#Region "Check the Filepaths"
    Private Function ValidateFilePaths(aPath as String ) As Boolean
        Dim strMessage, strpath, strFilename, strErrorLogPath As String
        strErrorLogPath = aPath
        strpath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath
        If Directory.Exists(strpath) = False Then
            System.IO.Directory.CreateDirectory(strpath)
            Return False
        End If

        Return True
    End Function
#End Region
#Region "Write into ErrorLog File"
    Public Sub WriteErrorHeader(ByVal apath As String, ByVal strMessage As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = Now.Date.ToString("dd/MM/yyyy") & ":" & Now.ToShortTimeString.Replace(":", "") & " --> " & strMessage
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(apath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "Export Documents Details"
    Public Sub ExportSKU(ByVal aPath As String, ByVal aChoice As String)
        If aChoice <> "SKU" Then
            Exit Sub
        End If
        Dim strPath, strFilename, strMessage, strExportFilePaty, strErrorLog, strTime As String
        strPath = aPath ' System.Windows.Forms.Application.StartupPath
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename = Now.Date.ToString("ddMMyyyy")
        strFilename = strFilename & strTime
        Dim stHours, stMin As String

        strErrorLog = ""
        If aChoice = "SKU" Then
            strErrorLog = strPath & "\Logs\SKU Import"
            strPath = strPath & "\Export\SKU Import"
        End If
        strExportFilePaty = strPath
        If Directory.Exists(strPath) = False Then
            Directory.CreateDirectory(strPath)
        End If
        If Directory.Exists(strErrorLog) = False Then
            Directory.CreateDirectory(strErrorLog)
        End If
        strFilename = "Export SKU_" & strFilename
        strErrorLog = strErrorLog & "\" & strFilename & ".txt"
        Message("Processing SKU's Exporting...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        WriteErrorHeader(strErrorLog, "Export SKU's Starting..")
        If Directory.Exists(strExportFilePaty) = False Then
            Directory.CreateDirectory(strPath)
        End If
        Try
            Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim strRecquery, strdocnum As String
            Dim oCheckrs As SAPbobsCOM.Recordset
            oCheckrs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '  strRecquery = "SELECT isnull(T0.[U_StoreKey],''),T0.[ItemCode], T0.[ItemName], T1.[ItmsGrpNam], T0.[ItemType], T0.[SWeight1], T0.[SVolume], isnull(U_Expirable,'N'), T0.[CodeBars],isnull(T0.U_BxBarCode,''),Isnull(T0.U_CrBarCode,'') , Isnull(T0.U_PlBarCode,''),T0.SalUnitMsr FROM OITM T0  INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod  and T0.ItemCode in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='SKU' and U_Z_Exported='N')"
            strRecquery = "SELECT isnull(T0.[U_StoreKey],''),T0.[ItemCode], T0.[ItemName], T1.[ItmsGrpNam], T0.[ItemType], T0.[SWeight1], T0.[SVolume], isnull(U_Expirable,'N'), T0.[CodeBars],T0.U_Pack1,T0.NumInBuy , T0.PurPackUn FROM OITM T0  INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod  and T0.ItemCode in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='SKU' and U_Z_Exported='N' and U_Storekey='" & companyStorekey & "')"
            oCheckrs.DoQuery(strRecquery)
            oApplication.Utilities.Message("Exporting SKU's in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oCheckrs.RecordCount > 0 Then
                Dim otemprec As SAPbobsCOM.Recordset
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery(strRecquery)
                If 1 = 1 Then
                    s.Remove(0, s.Length)
                    Dim cols As Integer = 2 ' Me.DataSet1.SalesOrder.Columns.Count
                    strdocnum = ""

                    s.Remove(0, s.Length)
                    s.Append("StoreKey" + vbTab)
                    s.Append("SKU" + vbTab)
                    s.Append("Description" + vbTab)
                    s.Append("SKUGroup2" + vbTab)
                    s.Append("SKUSubGroup" + vbTab)
                    s.Append("Weight" + vbTab)
                    s.Append("Volume" + vbTab)
                    s.Append("RotareBy" + vbTab)
                    '  s.Append("DefaultRotataion" + vbTab)
                    s.Append("AltSKU" + vbTab)
                    ' s.Append("Displaybox" + vbTab)
                    ' s.Append("Case" + vbTab)
                    ' s.Append("Pallet" + vbTab)
                    ' s.Append("DefaultUOM" + vbCrLf)
                    s.Append("DisplayBoxQtyinEA" + vbTab)
                    s.Append("ExportCartonQtyinEA" + vbTab)
                    s.Append("QtyofDisplayBoxinPallet" + vbCrLf)



                    Dim strItem As String
                    strItem = ""
                    For intRow As Integer = 0 To otemprec.RecordCount - 1
                        Dim strQt, strStoreKey, strName, groupname, itemtype, weight, volume, expirable, codebars, packkey, defaultuom As String
                        strQt = CStr(otemprec.Fields.Item(1).Value)
                        If strItem = "" Then
                            strItem = "'" & otemprec.Fields.Item("ItemCode").Value & "'"
                        Else
                            strItem = strItem & ",'" & otemprec.Fields.Item("ItemCode").Value & "'"
                        End If
                        strName = otemprec.Fields.Item("ItemName").Value
                        strStoreKey = ""
                        expirable = ""
                        ' strRecquery = "SELECT T0.[U_StoreKey),T0.[ItemCode], T0.[ItemName], T1.[ItmsGrpNam], T0.[ItemType],
                        ' T0.[SWeight1], T0.[SVolume], isnull(U_Expirable,'N'), T0.[CodeBars],isnull(T0.U_BxBarCode,''),Isnull(T0.U_CrBarCode,'') , Isnull(T0.U_PlBarCode,''),T0.SalUnitMsr FROM OITM T0  INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod  and T0.ItemCode in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='SKU' and U_Z_Exported='N')"

                        s.Append(otemprec.Fields.Item(0).Value + vbTab)
                        s.Append(otemprec.Fields.Item(1).Value + vbTab)
                        s.Append(otemprec.Fields.Item(2).Value + vbTab)
                        s.Append(otemprec.Fields.Item(3).Value + vbTab)
                        s.Append(otemprec.Fields.Item(4).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(5).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(6).Value.ToString + vbTab)
                        expirable = otemprec.Fields.Item(7).Value
                        'If expirable = "N" Then
                        '    s.Append(strStoreKey + vbTab)
                        '    s.Append(strStoreKey + vbTab)
                        'Else
                        '    s.Append("Lottable05" + vbTab)
                        '    s.Append("1" + vbTab)
                        'End If
                        s.Append(expirable + vbTab)
                        s.Append(otemprec.Fields.Item(8).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(9).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(10).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(11).Value.ToString + vbCrLf)
                        ' s.Append(otemprec.Fields.Item(12).Value.ToString + vbCrLf)
                        otemprec.MoveNext()
                    Next
                    Dim filename As String
                    filename = Now.Date.ToString("ddMMyyyyhhmm")
                    filename = strExportFilePaty & "\SKU_" & strFilename & ".csv"
                    Dim strFilename1, strcode, strinsert As String
                    strFilename1 = filename
                    Try
                        My.Computer.FileSystem.WriteAllText(strFilename1, s.ToString, False)
                        If File.Exists("C:\Test123.txt") Then
                            File.Delete("C:\test123.txt")
                        End If
                    Catch ex As Exception
                        strMessage = "Export File name : " & strFilename1 & " failed . Check the ConnectionInfo.Ini /  Connection"
                        WriteErrorlog(strMessage, strErrorLog)
                        End
                    End Try
                    strMessage = strItem & "--> SKU's  Exported compleated: File Name : " & strFilename1
                    Dim oUpdateRS As SAPbobsCOM.Recordset
                    Dim strUpdate As String
                    oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strUpdate = "Update [@Z_DABT_Export] set U_Z_Exported='Y' ,U_Z_ExportMethod='M',U_Z_ExportFile='" & strFilename1 & "',U_Z_ExportDate=getdate() where U_Z_MasterCode in (" & strItem & ") and U_Z_DocType='SKU'"
                    oUpdateRS.DoQuery(strUpdate)
                    WriteErrorlog(strMessage, strErrorLog)
                End If
            Else
                strMessage = ("No new SKUs!")
                WriteErrorlog(strMessage, strErrorLog)
            End If
        Catch ex As Exception
            strMessage = ("An Error Occured. A log entry has been created." & ex.Message)
            WriteErrorlog(strMessage, strErrorLog)
        Finally
            strMessage = "Export process compleated"
            WriteErrorlog(strMessage, strErrorLog)
        End Try
        '   System.Windows.Forms.Application.Exit()
    End Sub

    Public Sub ExportSalesOrder(ByVal aPath As String, ByVal aChoice As String)
        Dim strPath, strFilename, strMessage, strExportFilePaty, strErrorLog, strTime As String
        If aChoice <> "SO" Then
            Exit Sub
        End If
        strPath = aPath ' System.Windows.Forms.Application.StartupPath
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename = Now.Date.ToString("ddMMyyyy")
        strFilename = strFilename & strTime
        Dim stHours, stMin As String

        strErrorLog = ""
        If aChoice = "SO" Then
            strErrorLog = strPath & "\Logs\SO Import"
            strPath = strPath & "\Export\SO Import"
        End If
        strExportFilePaty = strPath
        If Directory.Exists(strPath) = False Then
            Directory.CreateDirectory(strPath)
        End If
        If Directory.Exists(strErrorLog) = False Then
            Directory.CreateDirectory(strErrorLog)
        End If
        strFilename = "Export SO_" & strFilename
        strErrorLog = strErrorLog & "\" & strFilename & ".txt"
        Message("Processing Sales Order Exporting...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        WriteErrorHeader(strErrorLog, "Export SO's Starting..")
        If Directory.Exists(strExportFilePaty) = False Then
            Directory.CreateDirectory(strPath)
        End If
        Try
            Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim strRecquery, strdocnum, strString As String
            Dim oCheckrs As SAPbobsCOM.Recordset
            oCheckrs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.DocNum,'SO'e,T0.[DocNum], T1.[LineNum], T0.[CardCode], isnull(T1.[U_Shelflife],'0'),isnull(T0.[U_TrafLine],''),isnull(T0.[U_Cust_Class],''),T2.[SlpName],isnull(T0.[U_TrafLine],''), T0.[DocDueDate], T0.[DocDueDate],  '',T1.[ItemCode], convert(numeric,T1.[Quantity]), T0.[Comments], T0.[CardName],T0.[Address],T0.[DocNum] FROM [dbo].[ORDR]  T0 INNER JOIN [dbo].[RDR1]  T1 ON T0.DocEntry = T1.DocEntry and T1.LineStatus='O' INNER JOIN [dbo].[OSLP]  T2 ON T0.SlpCode = T2.SlpCode and T0.U_Storekey='" & companyStorekey & "'"
            strRecquery = strString & " and T0.DocStatus='O'  and T0.DocEntry in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='SO' and U_Z_Exported='N')"
            oCheckrs.DoQuery(strRecquery)
            oApplication.Utilities.Message("Exporting Sales Orders in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oCheckrs.RecordCount > 0 Then
                Dim otemprec As SAPbobsCOM.Recordset
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery(strRecquery)
                If 1 = 1 Then
                    s.Remove(0, s.Length)
                    Dim cols As Integer = 2 ' Me.DataSet1.SalesOrder.Columns.Count
                    strdocnum = ""

                    s.Remove(0, s.Length)
                    s.Append("WhsCode" + vbTab)
                    s.Append("StoreKey" + vbTab)
                    s.Append("ExternOrderKey" + vbTab)
                    s.Append("OrderType" + vbTab)
                    's.Append("OrderNumber" + vbTab)
                    s.Append("LineNum" + vbTab)
                    s.Append("ConsigneeKey" + vbTab)
                    s.Append("ShelfLfe" + vbTab)
                    s.Append("Route" + vbTab)
                    s.Append("SUser1" + vbTab)
                    s.Append("Sales Person" + vbTab)
                    s.Append("TraffiLine" + vbTab)
                    s.Append("RequestedDeliveryDate" + vbTab)
                    s.Append("ExpectedDelieryDate" + vbTab)
                    s.Append("CustomerTerritory" + vbTab)
                    s.Append("ItemCode" + vbTab)
                    s.Append("Quantity" + vbTab)
                    s.Append("OrderNote" + vbTab)
                    s.Append("CardName" + vbTab)
                    s.Append("CustomerAddress" + vbCrLf)
                    Dim strItem As String
                    strItem = ""
                    For intRow As Integer = 0 To otemprec.RecordCount - 1
                        oApplication.Utilities.Message("Exporting Sales Orders --> " & otemprec.Fields.Item("DocNum").Value & "  in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim strQt, strStoreKey, strDuedate, strDocDate, groupname, itemtype, weight, volume, expiryflag, codebars, packkey, defaultuom As String
                        Dim dtduedate, dtdocdate As Date
                        strQt = CStr(otemprec.Fields.Item(1).Value)
                        If strItem = "" Then
                            strItem = "'" & otemprec.Fields.Item("DocNum").Value & "'"
                        Else
                            strItem = strItem & ",'" & otemprec.Fields.Item("DocNum").Value & "'"
                        End If
                        'strName = otemprec.Fields.Item("ItemName").Value
                        strStoreKey = " "
                        expiryflag = " "
                        dtdocdate = otemprec.Fields.Item("DocDueDate").Value
                        dtduedate = otemprec.Fields.Item("DocDueDate").Value
                        strDocDate = dtdocdate.ToString("dd-MM-yyyy")
                        strDuedate = dtduedate.ToString("dd-MM-yyyy")
                        s.Append(otemprec.Fields.Item(1).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(2).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(3).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(4).Value.ToString + vbTab)
                        '                        s.Append(otemprec.Fields.Item(5).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(6).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(7).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(8).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(9).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(10).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(11).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(12).Value.ToString + vbTab)
                        s.Append(strDuedate + vbTab)
                        s.Append(strDuedate + vbTab)
                        's.Append(otemprec.Fields.Item(14).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(15).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(16).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(17).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(18).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(19).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(20).Value.ToString + vbCrLf)
                        otemprec.MoveNext()
                    Next
                    Dim filename As String
                    filename = Now.Date.ToString("ddMMyyyyhhmm")
                    filename = strExportFilePaty & "\SO_" & strFilename & ".csv"
                    Dim strFilename1, strcode, strinsert As String
                    strFilename1 = filename
                    Try
                        My.Computer.FileSystem.WriteAllText(strFilename1, s.ToString, False)
                        If File.Exists("C:\Test123.txt") Then
                            File.Delete("C:\test123.txt")
                        End If
                    Catch ex As Exception
                        strMessage = "Export File name : " & strFilename1 & " failed . Check the ConnectionInfo.Ini /  Connection"
                        WriteErrorlog(strMessage, strErrorLog)
                        End
                    End Try
                    strMessage = strItem & "--> SO's  Exported compleated: File Name : " & strFilename1
                    Dim oUpdateRS As SAPbobsCOM.Recordset
                    Dim strUpdate As String
                    oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strUpdate = "Update [@Z_DABT_Export] set U_Z_Exported='Y' ,U_Z_ExportMethod='M',U_Z_ExportFile='" & strFilename1 & "',U_Z_ExportDate=getdate() where U_Z_DocNum in (" & strItem & ") and U_Z_DocType='SO'"
                    oUpdateRS.DoQuery(strUpdate)
                    WriteErrorlog(strMessage, strErrorLog)
                End If
            Else
                strMessage = ("No new SO's!")
                WriteErrorlog(strMessage, strErrorLog)
            End If
        Catch ex As Exception
            strMessage = ("An Error Occured. A log entry has been created." & ex.Message)
            WriteErrorlog(strMessage, strErrorLog)
        Finally
            strMessage = "Export process compleated"
            WriteErrorlog(strMessage, strErrorLog)
        End Try
        '   System.Windows.Forms.Application.Exit()
    End Sub

    Public Sub ExportARCreditMemo(ByVal aPath As String, ByVal aChoice As String)
        Dim strPath, strFilename, strMessage, strExportFilePaty, strErrorLog, strTime As String
        If aChoice <> "ARCR" Then
            Exit Sub
        End If
        strPath = aPath ' System.Windows.Forms.Application.StartupPath
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename = Now.Date.ToString("ddMMyyyy")
        strFilename = strFilename & strTime
        Dim stHours, stMin As String

        strErrorLog = ""
        If aChoice = "ARCR" Then
            strErrorLog = strPath & "\Logs\ARCR Import"
            strPath = strPath & "\Export\ARCR Import"
        End If
        strExportFilePaty = strPath
        If Directory.Exists(strPath) = False Then
            Directory.CreateDirectory(strPath)
        End If
        If Directory.Exists(strErrorLog) = False Then
            Directory.CreateDirectory(strErrorLog)
        End If
        strFilename = "Export ARCR_" & strFilename
        strErrorLog = strErrorLog & "\" & strFilename & ".txt"
        Message("Processing Supplier Returns Exporting...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        WriteErrorHeader(strErrorLog, "Export Supplier returns Starting..")
        If Directory.Exists(strExportFilePaty) = False Then
            Directory.CreateDirectory(strPath)
        End If
        Try
            Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim strRecquery, strdocnum, strString As String
            Dim oCheckrs As SAPbobsCOM.Recordset
            oCheckrs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.DocNum,'OR',T0.[DocNum], T1.[LineNum], T0.[CardCode], isnull(T1.[U_Shelflife],'0'),isnull(T0.[U_TrafLine],''),isnull(T0.[U_Cust_Class],''),T2.[SlpName],isnull(T0.[U_TrafLine],''), T0.[DocDueDate], T0.[DocDueDate],  '',T1.[ItemCode], T1.[Quantity], T0.[Comments], T0.[CardName],T0.[Address],T0.[DocNum] FROM [dbo].[ODRF]  T0 INNER JOIN [dbo].[DRF1]  T1 ON T0.DocEntry = T1.DocEntry and T0.ObjType=16 INNER JOIN [dbo].[OSLP]  T2 ON T0.SlpCode = T2.SlpCode and T0.U_Storekey='" & companyStorekey & "'"
            strRecquery = strString & " and T0.DocStatus='O' and T0.DocEntry in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='ARCR' and U_Z_Exported='N')"
            oCheckrs.DoQuery(strRecquery)
            oApplication.Utilities.Message("Exporting Supplier return in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oCheckrs.RecordCount > 0 Then
                Dim otemprec As SAPbobsCOM.Recordset
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery(strRecquery)
                If 1 = 1 Then
                    s.Remove(0, s.Length)
                    Dim cols As Integer = 2 ' Me.DataSet1.SalesOrder.Columns.Count
                    strdocnum = ""

                    s.Remove(0, s.Length)
                    s.Append("WhsCode" + vbTab)
                    s.Append("StoreKey" + vbTab)
                    s.Append("ExternOrderKey" + vbTab)
                    s.Append("OrderType" + vbTab)
                    ' s.Append("OrderNumber" + vbTab)
                    s.Append("LineNum" + vbTab)
                    s.Append("ConsigneeKey" + vbTab)
                    s.Append("ShelfLfe" + vbTab)
                    s.Append("Route" + vbTab)
                    s.Append("SUser1" + vbTab)
                    s.Append("Sales Person" + vbTab)
                    s.Append("TraffiLine" + vbTab)
                    s.Append("RequestedDeliveryDate" + vbTab)
                    s.Append("ExpectedDelieryDate" + vbTab)
                    s.Append("CustomerTerritory" + vbTab)
                    s.Append("ItemCode" + vbTab)
                    s.Append("Quantity" + vbTab)
                    s.Append("OrderNote" + vbTab)
                    s.Append("CardName" + vbTab)
                    s.Append("CustomerAddress" + vbCrLf)
                    Dim strItem As String
                    strItem = ""
                    For intRow As Integer = 0 To otemprec.RecordCount - 1

                        oApplication.Utilities.Message("Exporting AR Credit Memo --> " & otemprec.Fields.Item("DocNum").Value & "  in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim strQt, strStoreKey, strDuedate, strDocDate, groupname, itemtype, weight, volume, expiryflag, codebars, packkey, defaultuom As String
                        Dim dtduedate, dtdocdate As Date
                        strQt = CStr(otemprec.Fields.Item(1).Value)
                        If strItem = "" Then
                            strItem = "'" & otemprec.Fields.Item("DocNum").Value & "'"
                        Else
                            strItem = strItem & ",'" & otemprec.Fields.Item("DocNum").Value & "'"
                        End If
                        'strName = otemprec.Fields.Item("ItemName").Value
                        strStoreKey = " "
                        expiryflag = " "
                        'strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.NumAtCard,T0.ObjectType,T0.[DocNum], T1.[LineNum],
                        'T0.[CardCode], isnull(T1.[U_Shelflife,'0'),isnull(T1.[U_TrafLine],''),
                        'isnull(T0.[U_Cust_Class],''),T2.[SlpName],,isnull(T1.[U_TrafLine],''), T0.[DocDueDate], T0.[DocDueDate],  '',T1.[ItemCode], T1.[Quantity], T0.[Comments], 
                        'T0.[CardName], T0.[Address]
                        ' FROM [dbo].[ORDR]  T0 INNER JOIN [dbo].[RDR1]  T1 ON T0.DocEntry = T1.DocEntry INNER JOIN [dbo].[OSLP]  T2 ON T0.SlpCode = T2.SlpCode"
                        dtdocdate = otemprec.Fields.Item("DocDueDate").Value
                        dtduedate = otemprec.Fields.Item("DocDueDate").Value
                        strDocDate = dtdocdate.ToString("dd-MM-yyyy")
                        strDuedate = dtduedate.ToString("dd-MM-yyyy")
                        s.Append(otemprec.Fields.Item(1).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(2).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(3).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(4).Value.ToString + vbTab)
                        '    s.Append(otemprec.Fields.Item(5).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(6).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(7).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(8).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(9).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(10).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(11).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(12).Value.ToString + vbTab)
                        s.Append(strDuedate + vbTab)
                        s.Append(strDuedate + vbTab)
                        ' s.Append(otemprec.Fields.Item(14).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(15).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(16).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(17).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(18).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(19).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(20).Value.ToString + vbCrLf)
                        otemprec.MoveNext()
                    Next

                    Dim filename As String
                    filename = Now.Date.ToString("ddMMyyyyhhmm")
                    filename = strExportFilePaty & "\ARCR_" & strFilename & ".csv"
                    Dim strFilename1, strcode, strinsert As String
                    strFilename1 = filename
                    Try
                        My.Computer.FileSystem.WriteAllText(strFilename1, s.ToString, False)
                        If File.Exists("C:\Test123.txt") Then
                            File.Delete("C:\test123.txt")
                        End If
                    Catch ex As Exception
                        strMessage = "Export File name : " & strFilename1 & " failed . Check the ConnectionInfo.Ini /  Connection"
                        WriteErrorlog(strMessage, strErrorLog)
                        End
                    End Try
                    strMessage = strItem & "--> Supplier returns Exported compleated: File Name : " & strFilename1
                    Dim oUpdateRS As SAPbobsCOM.Recordset
                    Dim strUpdate As String
                    oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strUpdate = "Update [@Z_DABT_Export] set U_Z_Exported='Y' ,U_Z_ExportMethod='M',U_Z_ExportFile='" & strFilename1 & "',U_Z_ExportDate=getdate() where U_Z_DocNum in (" & strItem & ") and U_Z_DocType='ARCR'"
                    oUpdateRS.DoQuery(strUpdate)
                    WriteErrorlog(strMessage, strErrorLog)
                End If
            Else
                strMessage = ("No new Supplier returns!")
                WriteErrorlog(strMessage, strErrorLog)
            End If
        Catch ex As Exception
            strMessage = ("An Error Occured. A log entry has been created." & ex.Message)
            WriteErrorlog(strMessage, strErrorLog)
        Finally
            strMessage = "Export process compleated"
            WriteErrorlog(strMessage, strErrorLog)
        End Try
        '   System.Windows.Forms.Application.Exit()
    End Sub

    Public Sub ExportPurchaseOrder(ByVal aPath As String, ByVal aChoice As String)
        Dim strPath, strFilename, strMessage, strExportFilePaty, strErrorLog, strTime As String
        If aChoice <> "PO" Then
            Exit Sub
        End If
        strPath = aPath ' System.Windows.Forms.Application.StartupPath
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename = Now.Date.ToString("ddMMyyyy")
        strFilename = strFilename & strTime
        strErrorLog = ""
        If aChoice = "PO" Then
            strErrorLog = strPath & "\Logs\ASN Import"
            strPath = strPath & "\Export\ASN Import"
        End If
        strExportFilePaty = strPath
        If Directory.Exists(strPath) = False Then
            Directory.CreateDirectory(strPath)
        End If
        If Directory.Exists(strErrorLog) = False Then
            Directory.CreateDirectory(strErrorLog)
        End If
        strFilename = "Export ASN_" & strFilename
        strErrorLog = strErrorLog & "\" & strFilename & ".txt"
        Message("Processing Purchase Order Exporting...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        WriteErrorHeader(strErrorLog, "Export PO's Starting..")
        If Directory.Exists(strExportFilePaty) = False Then
            Directory.CreateDirectory(strPath)
        End If
        Try
            Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim strRecquery, strdocnum, strString As String
            Dim oCheckrs As SAPbobsCOM.Recordset
            oCheckrs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.[DocType], T0.NumAtCard,T0.NumAtCard,T0.[DocNum],  T0.[CardCode], T0.[DocDate], T0.[DocDueDate], T1.[ItemCode], T1.[Quantity] FROM [dbo].[OPOR]  T0 INNER JOIN [dbo].[POR1]  T1 ON T0.DocEntry = T1.DocEntry "
            strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.[DocType], T0.NumAtCard,T0.NumAtCard,T0.[DocNum],  T0.[CardCode], T0.[DocDate], T0.[DocDueDate], T1.[ItemCode], T1.[Quantity], T1.Quantity * T2.NumInBuy * isnull(T2.U_Pack1,1) ,T1.LineNum FROM [dbo].[OPOR]  T0 INNER JOIN [dbo].[POR1]  T1 ON T0.DocEntry = T1.DocEntry  and T1.LineStatus='O' inner Join OITM T2 on T1.ItemCode=T2.ItemCode and T0.U_Storekey='" & companyStorekey & "'"
            strRecquery = strString & " and T0.DocStatus='O'   and T0.DocEntry in (Select U_Z_Mastercode from [@Z_DABT_EXPORT] where U_Z_DocType='PO' and U_Z_Exported='N')"
            oCheckrs.DoQuery(strRecquery)
            oApplication.Utilities.Message("Exporting Purchase Orders   in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oCheckrs.RecordCount > 0 Then
                Dim otemprec As SAPbobsCOM.Recordset
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery(strRecquery)
                If 1 = 1 Then
                    s.Remove(0, s.Length)
                    Dim cols As Integer = 2 ' Me.DataSet1.SalesOrder.Columns.Count
                    strdocnum = ""
                    s.Remove(0, s.Length)
                    s.Append("WhsCode" + vbTab)
                    s.Append("StoreKey" + vbTab)
                    s.Append("DocType" + vbTab)
                    s.Append("Buyerref" + vbTab)
                    s.Append("SellerRef" + vbTab)
                    s.Append("DocumentNumber" + vbTab)
                    s.Append("Supplier Code" + vbTab)
                    s.Append("Document Date" + vbTab)
                    s.Append("ExpectedreceiptDate" + vbTab)
                    s.Append("ItemCode" + vbTab)
                    s.Append("Quantity" + vbTab)
                    s.Append("LineNum" + vbCrLf)

                    Dim strItem As String
                    strItem = ""
                    For intRow As Integer = 0 To otemprec.RecordCount - 1
                        oApplication.Utilities.Message("Exporting Purchase Orders --> " & otemprec.Fields.Item("DocNum").Value & "  in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Dim strQt, strStoreKey, strDuedate, strDocDate, groupname, itemtype, weight, volume, expiryflag, codebars, packkey, defaultuom As String
                        Dim dtduedate, dtdocdate As Date
                        strQt = CStr(otemprec.Fields.Item(1).Value)
                        If strItem = "" Then
                            strItem = "'" & otemprec.Fields.Item("DocNum").Value & "'"
                        Else
                            strItem = strItem & ",'" & otemprec.Fields.Item("DocNum").Value & "'"
                        End If
                        'strName = otemprec.Fields.Item("ItemName").Value
                        strStoreKey = " "
                        expiryflag = " "
                        dtdocdate = otemprec.Fields.Item("DocDate").Value
                        dtduedate = otemprec.Fields.Item("DocDueDate").Value
                        strDocDate = dtdocdate.ToString("dd-MM-yyyy")
                        strDuedate = dtduedate.ToString("dd-MM-yyyy")
                        'strString = "SELECT T0.[DocEntry],T1.[WhsCode],isnull(T0.[U_StoreKey],''), T0.[DocType], T0.NumAtCard,T0.NumAtCard,T0.[DocNum],  T0.[CardCode], T0.[DocDate], T0.[DocDueDate], T1.[ItemCode], T1.[Quantity] FROM [dbo].[OPOR]  T0 INNER JOIN [dbo].[POR1]  T1 ON T0.DocEntry = T1.DocEntry "
                        s.Append(otemprec.Fields.Item(1).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(2).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(3).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(4).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(5).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(6).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(7).Value.ToString + vbTab)
                        s.Append(strDocDate + vbTab)
                        s.Append(strDuedate + vbTab)
                        s.Append(otemprec.Fields.Item(10).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(12).Value.ToString + vbTab)
                        s.Append(otemprec.Fields.Item(13).Value.ToString + vbCrLf)
                        otemprec.MoveNext()
                    Next
                    Dim filename As String
                    filename = Now.Date.ToString("ddMMyyyyhhmm")
                    filename = strExportFilePaty & "\ASN_" & strFilename & ".csv"
                    Dim strFilename1, strcode, strinsert As String
                    strFilename1 = filename
                    Try
                        My.Computer.FileSystem.WriteAllText(strFilename1, s.ToString, False)
                        If File.Exists("C:\Test123.txt") Then
                            File.Delete("C:\test123.txt")
                        End If
                    Catch ex As Exception
                        strMessage = "Export File name : " & strFilename1 & " failed . Check the ConnectionInfo.Ini /  Connection"
                        WriteErrorlog(strMessage, strErrorLog)
                        End
                    End Try
                    strMessage = strItem & "--> PO's  Exported compleated: File Name : " & strFilename1
                    Dim oUpdateRS As SAPbobsCOM.Recordset
                    Dim strUpdate As String
                    oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strUpdate = "Update [@Z_DABT_Export] set U_Z_Exported='Y' ,U_Z_ExportMethod='M',U_Z_ExportFile='" & strFilename1 & "',U_Z_ExportDate=getdate() where U_Z_DocNum in (" & strItem & ") and U_Z_DocType='PO'"
                    oUpdateRS.DoQuery(strUpdate)
                    WriteErrorlog(strMessage, strErrorLog)
                End If
            Else
                strMessage = ("No new PO's!")
                WriteErrorlog(strMessage, strErrorLog)
            End If
        Catch ex As Exception
            strMessage = ("An Error Occured. A log entry has been created." & ex.Message)
            WriteErrorlog(strMessage, strErrorLog)
        Finally
            strMessage = "Export process compleated"
            WriteErrorlog(strMessage, strErrorLog)
        End Try
        '   System.Windows.Forms.Application.Exit()
    End Sub
#End Region

#Region "Import Documents"

    Public Sub ImportASNFiles(ByVal apath As String)
        ImportASN_GRPOFiles(apath)
        ImportASNRETURNSFiles(apath)
        ImportASNARCRFiles(apath)
    End Sub

    Public Sub ImportASN_GRPOFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.Documents
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASN Import"
            strPath = strPath & "\Import\ASN Import"
            strDeg = strPath & "\Import\ASN Import\Success"
            strExportFilePaty = strPath
            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XASN_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"
            Message("Processing XASN-GRPO Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XASN Import GRPO Starting..")
            WriteErrorlog("XASN-GRPO Import starting...", strImportErrorLog)
            Dim stStore As String
            stStore = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_SAPDocKey,''),Count(*) from   [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N'  and U_Z_ImpDocType ='GRPO' group by U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_SAPDocKey,'')"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XASN-GRPO Import Completed...", strImportErrorLog)
                Exit Sub
            End If

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing XASN Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  isnull(U_Z_SAPDocKey,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                If strDocType = "GRPO" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oDocument.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
                    oTempLines.DoQuery(strSQL1)
                    blnLineExists = False
                    For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                        oSourceDocument.DoQuery("Select * from OPOR where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_SAPDockey").Value)
                        If oSourceDocument.RecordCount > 0 Then
                            oDocument.CardCode = oSourceDocument.Fields.Item("CardCode").Value
                            oDocument.DocDate = oTempLines.Fields.Item("U_Z_Receiptdate").Value
                            If intLoop > 0 Then
                                oDocument.Lines.Add()
                            End If
                            oDocument.Lines.SetCurrentLine(intLoop)
                            oDocument.Lines.BaseType = 22
                            oDocument.Lines.BaseEntry = oSourceDocument.Fields.Item("DocEntry").Value
                            oDocument.Lines.BaseLine = oTempLines.Fields.Item("U_Z_Lineno").Value
                            oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                            blnLineExists = True
                        Else
                            WriteErrorlog("Purchase order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strErrorLog)
                            WriteErrorlog("Purchase order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strImportErrorLog)
                        End If
                        oTempLines.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDocument.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                        Else
                            Dim strdocCode As String
                            oApplication.Company.GetNewObjectCode(strdocCode)
                            If oDocument.GetByKey(strdocCode) Then
                                otempLines1.DoQuery("Update [@Z_DABT_XASN] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & stStore & "' and  U_Z_SAPDocKey='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                WriteErrorlog("Draft-- GRPO Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                WriteErrorlog("Draft -- GRPO Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                            End If
                        End If
                    End If
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XASN-GRPO Import Completed..")
            WriteErrorlog("XASN-GRPO Import Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ImportASNARCRFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.Documents
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASN Import"
            strPath = strPath & "\Import\ASN Import"
            strDeg = strPath & "\Import\ASN Import\Success"
            strExportFilePaty = strPath
            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XASN_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            Message("Processing XASN-ARCR Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XASN-ARCR Import Starting..")
            WriteErrorlog("XASN-ARCR Import starting...", strImportErrorLog)
            Dim ststore As String = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_Susr,''),Count(*) from   [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' and U_Z_ImpDocType='ARCR' group by U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_Susr,'')"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XASN Import Completed...", strImportErrorLog)
                Exit Sub
            End If
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing XASN-ARCR Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  isnull(U_Z_Susr,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                If strDocType = "ARCR" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oDocument.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                    oTempLines.DoQuery(strSQL1)
                    blnLineExists = False
                    For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                        oSourceDocument.DoQuery("Select * from OINV where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_Susr").Value)
                        If oSourceDocument.RecordCount > 0 Then
                            oDocument.CardCode = oSourceDocument.Fields.Item("CardCode").Value
                            oDocument.DocDate = Now.Date
                            oDocument.DocDueDate = Now.Date
                            If intLoop > 0 Then
                                oDocument.Lines.Add()
                            End If
                            oDocument.Lines.SetCurrentLine(intLoop)
                            oDocument.Lines.BaseType = 13
                            oDocument.Lines.BaseEntry = oSourceDocument.Fields.Item("DocEntry").Value
                            oDocument.Lines.BaseLine = oTempLines.Fields.Item("U_Z_Lineno").Value
                            'oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                            oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                            'oDocument.Lines.WarehouseCode = oTempLines.Fields.Item("U_Z_Storekey").Value
                            blnLineExists = True
                        Else
                            WriteErrorlog("Sales Invoice does not exits : FileName =" & strFileName & " : Invoice No : " & oTempLines.Fields.Item("U_Z_Susr").Value, strErrorLog)
                            WriteErrorlog("Sales Invoice does not exits : FileName =" & strFileName & " : Invoice No : " & oTempLines.Fields.Item("U_Z_Susr").Value, strImportErrorLog)
                        End If
                        oTempLines.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDocument.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                        Else
                            Dim strdocCode As String
                            oApplication.Company.GetNewObjectCode(strdocCode)
                            If oDocument.GetByKey(strdocCode) Then
                                otempLines1.DoQuery("Update [@Z_DABT_XASN] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & ststore & "' and  isnull(U_Z_Susr,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                WriteErrorlog("Draft - AR-Credit Memo Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                WriteErrorlog("Draft -AR-Credit Memo Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                            End If
                        End If
                    End If
                ElseIf strDocEntry = "ST" Then

                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XASN-ARCR Import Completed..")
            WriteErrorlog("XASN-ARCR Import Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ImportASNRETURNSFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.Documents
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASN Import"
            strPath = strPath & "\Import\ASN Import"
            strDeg = strPath & "\Import\ASN Import\Success"
            strExportFilePaty = strPath

            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XASN_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            Message("Processing XASN-RETURNS Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XASN-RETURNS Import Starting..")
            WriteErrorlog("XASN-RETURNS Import starting...", strImportErrorLog)
            Dim stStore As String = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_Susr,''),Count(*) from   [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' and U_Z_ImpDocType='RETURNS' group by U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_Susr,'')"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XASN-RETURNS Completed...", strImportErrorLog)
                Exit Sub
            End If

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing XASN-Returns Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  isnull(U_Z_Susr,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                If strDocType = "RETURNS" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)

                    oTempLines.DoQuery(strSQL1)
                    blnLineExists = False
                    For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                        ' oSourceDocument.DoQuery("Select * from OPOR where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_SAPDockey").Value)
                        If 1 = 1 Then
                            oDocument.CardCode = oTempLines.Fields.Item("U_Z_Susr").Value
                            oDocument.DocDate = Now.Date ' oTempLines.Fields.Item("U_Z_Receiptdate").Value
                            If intLoop > 0 Then
                                oDocument.Lines.Add()
                            End If
                            oDocument.Lines.SetCurrentLine(intLoop)
                            oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                            oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                            Dim otemp As SAPbobsCOM.Recordset
                            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            otemp.DoQuery("Select T0.[DfltWhs] from OADM T0")
                            oDocument.Lines.WarehouseCode = otemp.Fields.Item(0).Value
                            blnLineExists = True
                        Else
                            WriteErrorlog("Purchase order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strErrorLog)
                        End If
                        oTempLines.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDocument.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                        Else
                            Dim strdocCode As String
                            oApplication.Company.GetNewObjectCode(strdocCode)
                            If oDocument.GetByKey(strdocCode) Then
                                otempLines1.DoQuery("Update [@Z_DABT_XASN] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & stStore & "' and  isnull(U_Z_Susr,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                WriteErrorlog("Return Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                WriteErrorlog("Return Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                            End If
                        End If
                    End If
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XASN -Returns Import Completed..")
            WriteErrorlog("XASN-Returns Import Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Public Sub ImportASNSTFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey, strFromWhs, strToWhs As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.StockTransfer
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASN Import"
            strPath = strPath & "\Import\ASN Import"
            strDeg = strPath & "\Import\ASN Import\Success"
            strExportFilePaty = strPath

            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XASN_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            Message("Processing XASN-ST Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XASN-ST Import Starting..")
            WriteErrorlog("XASN-ST Import starting...", strImportErrorLog)
            Dim stStore As String = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_Type,U_Z_ImpDocType,Count(*) from   [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' and U_Z_ImpDocType='ST' group by U_Z_FileName,U_Z_Type,U_Z_ImpDocType"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XASN-RETURNS Completed...", strImportErrorLog)
                Exit Sub
            End If
            Dim strType As String
            Dim owhsrec As SAPbobsCOM.Recordset
            owhsrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strType = oTempRec.Fields.Item(1).Value
                strDocType = oTempRec.Fields.Item(2).Value
                oWhsrec.DoQuery("Select * from [@Z_DABT_ST] where U_Z_Storekey='" & stStore & "' and U_Z_Type='" & strType & "'")
                If oWhsrec.RecordCount > 0 Then
                    strFromWhs = oWhsrec.Fields.Item("U_Z_FrmWhs").Value
                    strToWhs = owhsrec.Fields.Item("U_Z_ToWhs").Value
                
                    Message("Processing XASN-Returns Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                    WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                    strSQL1 = "Select * from [@Z_DABT_XASN] where U_Z_Storekey='" & stStore & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                    If strDocType = "ST" Then
                        oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        oTempLines.DoQuery(strSQL1)
                        blnLineExists = False
                        For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                            ' oSourceDocument.DoQuery("Select * from OPOR where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_SAPDockey").Value)
                            If 1 = 1 Then
                                oDocument.FromWarehouse = strFromWhs ' oTempLines.Fields.Item("U_Z_FrmWhs").Value
                                If intLoop > 0 Then
                                    oDocument.Lines.Add()
                                End If
                                oDocument.Lines.SetCurrentLine(intLoop)
                                oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                                oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                                oDocument.Lines.WarehouseCode = strToWhs 'oTempLines.Fields.Item("U_Z_ToWhs").Value
                                blnLineExists = True
                            Else
                                ' WriteErrorlog("DatabaseName name :  " & objRemoteCompany.CompanyDB & " Purchase order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strErrorLog)
                            End If
                            oTempLines.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oDocument.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                                WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                            Else
                                Dim strdocCode As String
                                oApplication.Company.GetNewObjectCode(strdocCode)
                                If oDocument.GetByKey(strdocCode) Then
                                    otempLines1.DoQuery("Update [@Z_DABT_XASN] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Type='" & strType & "' and  U_Z_Storekey='" & stStore & "'  and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                    WriteErrorlog("Stock transfer Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                    WriteErrorlog("Stock transfer Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                                End If
                            End If
                        End If
                    End If
                Else
                    WriteErrorlog("Warehouse details missing for the type : " & strType & " : storekey : " & stStore, strImportErrorLog)
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XASN -ST Import Completed..")
            WriteErrorlog("XASN-ST Import Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Public Sub ImportSOTFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey, strFromWhs, strToWhs As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.StockTransfer
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASO Import"
            strPath = strPath & "\Import\ASO Import"
            strDeg = strPath & "\Import\ASO Import\Success"
            strExportFilePaty = strPath

            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XASO" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            Message("Processing XASO-ST Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XASO-ST Import Starting..")
            WriteErrorlog("XASO-ST Import starting...", strImportErrorLog)
            Dim stStore As String = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_Type,U_Z_ImpDocType,Count(*) from   [@Z_DABT_XSO] where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' and U_Z_ImpDocType='INVTRN' group by U_Z_FileName,U_Z_Type,U_Z_ImpDocType"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XASN-RETURNS Completed...", strImportErrorLog)
                Exit Sub
            End If
            Dim strType As String
            Dim owhsrec As SAPbobsCOM.Recordset
            owhsrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strType = oTempRec.Fields.Item(1).Value
                strDocType = oTempRec.Fields.Item(2).Value
                owhsrec.DoQuery("Select * from [@Z_DABT_ST] where U_Z_Storekey='" & stStore & "' and U_Z_Type='" & strType & "'")
                If owhsrec.RecordCount > 0 Then
                    strFromWhs = owhsrec.Fields.Item("U_Z_FrmWhs").Value
                    strToWhs = owhsrec.Fields.Item("U_Z_ToWhs").Value
                    Message("Processing XASN-Returns Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                    WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                    strSQL1 = "Select * from [@Z_DABT_XSO] where U_Z_Storekey='" & stStore & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                    If strDocType = "INVTRN" Then
                        oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        oTempLines.DoQuery(strSQL1)
                        blnLineExists = False
                        For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                            ' oSourceDocument.DoQuery("Select * from OPOR where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_SAPDockey").Value)
                            If 1 = 1 Then
                                oDocument.FromWarehouse = strFromWhs ' oTempLines.Fields.Item("U_Z_FrmWhs").Value
                                ' oDocument.Comments = oTempLines.Fields.Item("U_Z_Remarks").Value
                                If intLoop > 0 Then
                                    oDocument.Lines.Add()
                                End If
                                oDocument.Lines.SetCurrentLine(intLoop)
                                oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                                oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                                oDocument.Lines.WarehouseCode = strToWhs 'oTempLines.Fields.Item("U_Z_ToWhs").Value
                                blnLineExists = True
                            Else
                                ' WriteErrorlog("DatabaseName name :  " & objRemoteCompany.CompanyDB & " Purchase order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strErrorLog)
                            End If
                            oTempLines.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oDocument.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                                WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                            Else
                                Dim strdocCode As String
                                oApplication.Company.GetNewObjectCode(strdocCode)
                                If oDocument.GetByKey(strdocCode) Then
                                    otempLines1.DoQuery("Update [@Z_DABT_XSO] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & stStore & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                    WriteErrorlog("Stock transfer Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                    WriteErrorlog("Stock transfer Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                                End If
                            End If
                        End If
                    End If
                Else
                    WriteErrorlog("Warehouse details missing for the type : " & strType & " : storekey : " & stStore, strImportErrorLog)
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XASO -ST Import Completed..")
            WriteErrorlog("XASO-ST Import Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ImportHOLDFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strDeg, strerrorfodler, strSQL, strSQL1, strDocKey As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.StockTransfer
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\HOLD Import"
            strPath = strPath & "\Import\HOLD Import"
            strDeg = strPath & "\Import\HOLD Import\Success"
            strExportFilePaty = strPath
            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XHOL_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            Message("Processing XHOLD Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XHOLD Import Starting..")
            WriteErrorlog("XHOLD starting...", strImportErrorLog)
            Dim strFrom, strTo As String
            strSQL = "Select DfltWhs from OADM"
            oTempRec.DoQuery(strSQL)
            strFrom = oTempRec.Fields.Item(0).Value
            strSQL = "Select WhsCode from OWHS where U_Damaged='Y'"
            oTempRec.DoQuery(strSQL)
            If oTempRec.RecordCount > 0 Then
                strTo = oTempRec.Fields.Item(0).Value
            Else
                WriteErrorlog("Damaged warehouse is not defined....", strErrorLog)
                WriteErrorlog("Damaged warehouse is not defined....", strImportErrorLog)
                Exit Sub
            End If
            strSQL = "select U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_FrmWhs,''),Count(*) from   [@Z_DABT_XHOL] where U_Z_Imported='N' and U_Z_ImpDocType='ST' group by U_Z_FileName,U_Z_ImpDocType,isnull(U_Z_FrmWhs,'')"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No Records to Import...", strImportErrorLog)
                WriteErrorlog("XHOLD Import Completed...", strImportErrorLog)
                Exit Sub
            End If
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing XHOLD Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & "  data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XHOL] where isnull(U_Z_FrmWhs,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                If strDocType = "ST" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                    oTempLines.DoQuery(strSQL1)
                    blnLineExists = False
                    For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                        ' oSourceDocument.DoQuery("Select * from OINV where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_Susr").Value)
                        If 1 = 1 Then
                            'oDocument.FromWarehouse = oTempLines.Fields.Item("U_Z_FrmWhs").Value
                            oDocument.FromWarehouse = strFrom
                            oDocument.Comments = oTempLines.Fields.Item("U_Z_Remarks").Value
                            If intLoop > 0 Then
                                oDocument.Lines.Add()
                            End If
                            oDocument.Lines.SetCurrentLine(intLoop)
                            Dim stItem As String
                            Dim dblqyt As Double
                            'stItem = oTempLines.Fields.Item("U_Z_SKU").Value
                            'dblqyt = CDbl(oTempLines.Fields.Item("U_Z_Quantity").Value)
                            oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                            oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                            'oDocument.Lines.WarehouseCode = oTempLines.Fields.Item("U_Z_ToWhs").Value
                            oDocument.Lines.WarehouseCode = strTo
                            blnLineExists = True
                          End If
                        oTempLines.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDocument.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                        Else
                            Dim strdocCode As String
                            oApplication.Company.GetNewObjectCode(strdocCode)
                            If oDocument.GetByKey(strdocCode) Then
                                otempLines1.DoQuery("Update [@Z_DABT_XHOL] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where isnull(U_Z_FrmWhs,'')='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                WriteErrorlog("Stock Transfer Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                WriteErrorlog("Stock Transfer Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                            End If
                        End If
                    End If
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XHOLD Import Completed..")
            WriteErrorlog("XHOLDImport Completed", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ImportADJFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strSQL, strSQL1, strDocKey, sPath As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.Documents
            Dim blnLineExists As Boolean
            Dim dblQuantity As Double
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"

            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\INV Import"
            strPath = strPath & "\Import\INV Import"
            strExportFilePaty = strPath
            'If Directory.Exists(strPath) = False Then
            '    Directory.CreateDirectory(strPath)
            'End If
            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XINV_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            'If Directory.Exists(strExportFilePaty) = False Then
            '    Directory.CreateDirectory(strPath)
            'End If
            Message("Processing Adjustment file Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "Adjustment files Import Starting..")
            WriteErrorlog("Import Inventory adjustment processing...", strImportErrorLog)
            Dim stStore As String = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,U_Z_ImpDocType, Count(*) from   [@Z_DABT_XADJ] where  U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' group by U_Z_FileName,U_Z_ImpDocType"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No records to Import", strImportErrorLog)
                Exit Sub
            End If

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing Adjustment files Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XADJ] where U_Z_Storekey='" & stStore & "' and  Convert(Numeric,isnull(U_Z_Adjkey,'0'))<>0 and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"
                If strDocType = "Goods Recipt" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                Else
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                End If
                oTempLines.DoQuery(strSQL1)
                blnLineExists = False
                For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                    If 1 = 1 Then
                        oDocument.DocDate = Now.Date
                        oDocument.Comments = oTempLines.Fields.Item("U_Z_Remarks").Value
                        If intLoop > 0 Then
                            oDocument.Lines.Add()
                        End If
                        dblQuantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                        If dblQuantity < 0 Then
                            dblQuantity = dblQuantity * -1
                        End If
                        oDocument.Lines.SetCurrentLine(intLoop)
                        oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                        oDocument.Lines.Quantity = dblQuantity
                        oDocument.Lines.WarehouseCode = oTempLines.Fields.Item("U_Z_Whs").Value
                        blnLineExists = True
                    End If
                    oTempLines.MoveNext()
                Next
                If blnLineExists = True Then
                    If oDocument.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                        WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                    Else
                        Dim strdocCode As String
                        oApplication.Company.GetNewObjectCode(strdocCode)
                        If oDocument.GetByKey(strdocCode) Then
                            otempLines1.DoQuery("Update [@Z_DABT_XADJ] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & stStore & "' and  U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                            WriteErrorlog(strDocType & " Document Created successfully. " & oDocument.DocNum, strErrorLog)
                            WriteErrorlog(strDocType & " Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                        End If
                    End If
                End If
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "Adjustment files Import Completed..")
            WriteErrorlog("Import Adjustment files completed...", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ImportSOFiles(ByVal apath As String)
        Try
            Dim oTempRec, oTempLines, otempLines1, oSourceDocument As SAPbobsCOM.Recordset
            Dim strFileName, strDocType, strSQL, strSQL1, strDocKey As String
            Dim strPath, strFilename1, strMessage, strExportFilePaty, strErrorLog, strTime As String
            Dim oDocument As SAPbobsCOM.Documents
            Dim blnLineExists As Boolean
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempLines = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otempLines1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSourceDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            strPath = apath ' System.Windows.Forms.Application.StartupPath
            strTime = Now.ToShortTimeString.Replace(":", "")
            strFilename1 = Now.Date.ToString("ddMMyyyy")
            strFilename1 = strFilename1 & strTime
            strErrorLog = ""

            strErrorLog = strPath & "\Logs\ASO Import"
            strPath = strPath & "\Import\ASO Import"
            strExportFilePaty = strPath
            'If Directory.Exists(strPath) = False Then
            '    Directory.CreateDirectory(strPath)
            'End If
            If Directory.Exists(strErrorLog) = False Then
                Directory.CreateDirectory(strErrorLog)
            End If
            strFilename1 = "Import XSO_" & strFilename1
            strErrorLog = strErrorLog & "\" & strFilename1 & ".txt"

            'If Directory.Exists(strExportFilePaty) = False Then
            '    Directory.CreateDirectory(strPath)
            'End If
            Message("Processing XSO Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteErrorHeader(strErrorLog, "XSOport Starting..")
            WriteErrorlog("Import XSO Processing...", strImportErrorLog)
            Dim stStore As String
            stStore = oApplication.Utilities.getStoreKey()
            strSQL = "select U_Z_FileName,isnull(U_Z_ImpDocType,'R'),U_Z_SAPDocKey,Count(*) from   [@Z_DABT_XSO] where U_Z_Imported='N' and U_Z_Storekey='" & stStore & "' group by U_Z_FileName,U_Z_ImpDocType,U_Z_SAPDocKey"
            oTempRec.DoQuery(strSQL)
            otempLines1.DoQuery(strSQL)
            If otempLines1.RecordCount <= 0 Then
                WriteErrorlog("No records to Import", strErrorLog)
                WriteErrorlog("No records to Import", strImportErrorLog)
                Exit Sub
            End If

            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                strFileName = oTempRec.Fields.Item(0).Value
                strDocType = oTempRec.Fields.Item(1).Value
                Message("Processing XSO Importing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data", strImportErrorLog)
                strSQL1 = "Select * from [@Z_DABT_XSO] where  U_Z_SAPDocKey='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Storekey='" & stStore & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'"

                If strDocType = "R" Then
                    oDocument = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    oTempLines.DoQuery(strSQL1)
                    blnLineExists = False
                    For intLoop As Integer = 0 To oTempLines.RecordCount - 1
                        oSourceDocument.DoQuery("Select * from ORDR where DocStatus='O' and DocNum=" & oTempLines.Fields.Item("U_Z_SAPDockey").Value)
                        If oSourceDocument.RecordCount > 0 Then
                            oDocument.CardCode = oSourceDocument.Fields.Item("CardCode").Value
                            oDocument.DocDate = oTempLines.Fields.Item("U_Z_Receiptdate").Value
                            If intLoop > 0 Then
                                oDocument.Lines.Add()
                            End If
                            oDocument.Lines.SetCurrentLine(intLoop)
                            oDocument.Lines.BaseType = 17
                            oDocument.Lines.BaseEntry = oSourceDocument.Fields.Item("DocEntry").Value
                            '  oDocument.Lines.ItemCode = oTempLines.Fields.Item("U_Z_SKU").Value
                            oDocument.Lines.Quantity = oTempLines.Fields.Item("U_Z_Quantity").Value
                            oDocument.Lines.BaseLine = oTempLines.Fields.Item("U_Z_Lineno").Value
                            ' oDocument.Lines.WarehouseCode = oTempLines.Fields.Item("U_Z_Storekey").Value
                            blnLineExists = True
                        Else
                            WriteErrorlog("Sales order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strErrorLog)
                            WriteErrorlog("Sales order does not exits : FileName =" & strFileName & " : Order No : " & oTempLines.Fields.Item("U_Z_SAPDockey").Value, strImportErrorLog)
                        End If
                        oTempLines.MoveNext()
                    Next
                    If blnLineExists = True Then
                        If oDocument.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strErrorLog)
                            WriteErrorlog(oApplication.Company.GetLastErrorDescription & " : FileName =" & strFileName, strImportErrorLog)
                        Else
                            Dim strdocCode As String
                            oApplication.Company.GetNewObjectCode(strdocCode)
                            If oDocument.GetByKey(strdocCode) Then
                                otempLines1.DoQuery("Update [@Z_DABT_XSO] set U_Z_Imported='Y',U_Z_SAPDocNum='" & oDocument.DocNum & "',U_Z_Impmethod='M' where U_Z_Storekey='" & stStore & "' and  U_Z_SAPDocKey='" & oTempRec.Fields.Item(2).Value & "' and U_Z_Imported='N' and U_Z_Filename='" & strFileName & "' and U_Z_ImpDocType='" & strDocType & "'")
                                WriteErrorlog("Invoice Document Created successfully. " & oDocument.DocNum, strErrorLog)
                                WriteErrorlog("Invoice Document Created successfully. " & oDocument.DocNum, strImportErrorLog)
                            End If
                        End If
                    End If
                End If

                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strErrorLog)
                WriteErrorlog("Importing FileName--> " & strFileName & " data completed", strImportErrorLog)
                oTempRec.MoveNext()
            Next
            WriteErrorHeader(strErrorLog, "XSO Import Completed..")
            WriteErrorlog("Import XSO completed...", strImportErrorLog)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
#End Region

#Region "Get StoreKey"
    Public Function getStoreKey() As String
        Dim stStorekey As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select isnull(U_Z_Storekey,'') from OADM")
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region
#End Region

#Region "Close Open Sales Order Lines"


    Public Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        Try
            If File.Exists(aPath) Then
            End If
            aSw = New StreamWriter(aPath, True)
            aMessage = Now.Date.ToString("dd/MM/yyyy") & ":" & Now.ToShortTimeString.Replace(":", "") & " --> " & aMessage
            aSw.WriteLine(aMessage)
            aSw.Flush()
            aSw.Close()
            aSw.Dispose()
        Catch ex As Exception
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Sub createARINvoice()
        Dim strCardcode, stritemcode As String
        Dim intbaseEntry, intbaserow As Integer
        Dim oInv As SAPbobsCOM.Documents
        strCardcode = "C20000"
        intbaseEntry = 66
        intbaserow = 1
        oInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        oInv.DocDate = Now.Date
        oInv.CardCode = strCardcode
        oInv.Lines.BaseType = 17
        oInv.Lines.BaseEntry = intbaseEntry
        oInv.Lines.BaseLine = intbaserow
        oInv.Lines.Quantity = 1
        If oInv.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Else
            oApplication.Utilities.Message("AR Invoice added", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End If

    End Sub
    Public Sub CloseOpenSOLines()
        Try
            Dim oDoc As SAPbobsCOM.Documents
            Dim oTemp As SAPbobsCOM.Recordset
            Dim strSQL, strSQL1, spath As String
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                File.Delete(spath)
            End If
            blnError = False
            ' oTemp.DoQuery("Select DocEntry,LineNum from RDR1 where isnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            '            oTemp.DoQuery("Select DocEntry,VisOrder,LineNum from RDR1 where isnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            oTemp.DoQuery("Select DocEntry,VisOrder,LineNum from RDR1 where   LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            oApplication.Utilities.Message("Processing closing Sales order Lines", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim numb As Integer
            For introw As Integer = 0 To oTemp.RecordCount - 1
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                numb = oTemp.Fields.Item(1).Value
                '  numb = oTemp.Fields.Item(2).Value
                If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                    oApplication.Utilities.Message("Processing Sales order :" & oDoc.DocNum, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oDoc.Comments = oDoc.Comments & "XXX1"
                    If oDoc.Update() <> 0 Then
                        WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                        blnError = True
                    Else
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                            Dim strcomments As String
                            strcomments = oDoc.Comments
                            strcomments = strcomments.Replace("XXX1", "")
                            oDoc.Comments = strcomments
                            oDoc.Lines.SetCurrentLine(numb)
                            '  MsgBox(oDoc.Lines.VisualOrder)
                            If oDoc.Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                oDoc.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                            End If
                            If oDoc.Update <> 0 Then
                                WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                                blnError = True
                                'oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                WriteErrorlog(" Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Closed successfully  ", spath)
                            End If
                        End If
                    End If

                End If
                oTemp.MoveNext()
            Next
            oApplication.Utilities.Message("Operation completed succesfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            blnError = True
            ' oApplication.SBO_Application.MessageBox("Error Occured...")\
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = spath
                System.Diagnostics.Process.Start(x)
                x = Nothing
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region



#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

    Public Sub SetEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aUID As String, ByVal aVal As String)
        Dim objedit As SAPbouiCOM.EditText
        objedit = aForm.Items.Item(aUID).Specific
        objedit.Value = aVal
    End Sub
    Public Function GetEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aUID As String) As String
        Dim objedit As SAPbouiCOM.EditText
        objedit = aForm.Items.Item(aUID).Specific
        Return objedit.String
    End Function

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Generate Bank DBF File"
    Public Function GenerateBankDBFFile(ByVal aBatchNumber As String, ByVal aFileName As String, ByVal aBankName As String) As Boolean
        Try
            If CreateBankDBFFile(aBatchNumber, aFileName, aBankName) = True Then
                Return True
            Else
                Return False
            End If
            'Dim oDRow As DataRow
            'Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
            'Dim strQuery, strJouranlQry, strfrom, strto, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
            'Dim dtFrom, dtTo, dtAging As Date
            'Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
            'oRecBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oBalanceRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            '    strQuery = "Select * from [@Z_Bill_Export] where U_Z_BatchNumber='" & aBatchNumber & "'"
            '    oRecTemp.DoQuery(strQuery)
            '    Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            '    If oRecTemp.RecordCount > 0 Then
            '        'strQuery = "Select isnull(T1.U_Z_KFHNO,''),x.CardCode,x.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y'   group by CardCode,Cardname "
            '        'strQuery = strQuery & " union select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y' group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname group by isnull(T1.U_Z_KFHNO,''),X.CardCode,X.CardName"
            '        strQuery = "  Select X.DocNum,X.DocTotal,'',X.BankrefNo,1,DocDate,'',X.Type,'122' from "
            '        strQuery = strQuery & "(select T0.CardCode,isnull(T1.U_Z_KFHNO,'') 'BankRefNo',DocNum,DocTotal,1'Type',DocDate  from OINV  T0 inner join OCRD T1 on T0.CardCode=t1.CardCode where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "'"
            '        strQuery = strQuery & "and isnull(U_Z_Exported,'N')='Y' union select T0.CardCode,isnull(T1.U_Z_KFHNO,'') 'BankRefNo',DocNum,DocTotal,2 'Type',DocDate  from ORIN  T0   inner join OCRD T1 on T0.CardCode=t1.CardCode where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "'"
            '        strQuery = strQuery & "and isnull(U_Z_Exported,'N')='Y'  )   x  Order by X.CardCode"
            '        oRecBP.DoQuery(strQuery)
            '        s.Remove(0, s.Length)
            '        s.Append("VOUCHER" + " ")
            '        s.Append("TOTAL " + " ")
            '        s.Append("BALANCE" + " ")
            '        s.Append("CO_OP" + " ")
            '        s.Append("BRANCH" + " ")
            '        s.Append("V_DATE" + " ")
            '        s.Append("VP_DATE" + " ")
            '        s.Append("TYPE" + " ")
            '        s.Append("COMPANY" + vbCrLf)
            '        Dim dtdate As Date
            '        For intRow As Integer = 0 To oRecBP.RecordCount - 1
            '            '  s.Remove(0, s.Length)
            '            dtdate = oRecBP.Fields.Item(5).Value
            '            s.Append(oRecBP.Fields.Item(0).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(1).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(2).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(3).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(4).Value.ToString + " ")
            '            s.Append(dtdate.ToString("yyyyddMM") + " ")
            '            s.Append(oRecBP.Fields.Item(6).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(7).Value.ToString + " ")
            '            s.Append(oRecBP.Fields.Item(8).Value.ToString + vbCrLf)
            '            oRecBP.MoveNext()
            '        Next
            '        Dim strfileaname As String
            '        strfileaname = aFileName & "\" & aBankName & "_BatchNumber_" & aBatchNumber & ".dbf"
            '        If File.Exists(strfileaname) Then
            '            File.Delete(strfileaname)
            '        End If
            '        My.Computer.FileSystem.WriteAllText(strfileaname, s.ToString, False)
            '    Else
            '    End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False

        End Try
    End Function

#Region "Get Price"
    Public Function getPrice(ByVal aPrice As String) As String
        Dim strCurrency As String
        Dim oCurrencyRs As SAPbobsCOM.Recordset
        oCurrencyRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCurrencyRs.DoQuery("Select * from OCRN")
        For intRow As Integer = 0 To oCurrencyRs.RecordCount - 1
            strCurrency = oCurrencyRs.Fields.Item("Currcode").Value
            aPrice = aPrice.Replace(strCurrency, "")
            oCurrencyRs.MoveNext()
        Next
        Return aPrice

    End Function



#Region "Get Price "
    Public Function GetB1Price(ByVal StrItem As String, ByVal strBP As String) As Double
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItems As SAPbobsCOM.Items
        Dim oRec, oREc1, oRecTemp, oRecDiscount As SAPbobsCOM.Recordset
        Dim strSQL, strSQL1, strDiscount, strBPCod As String
        Dim price, discount As Double
        Dim intFlag As Integer
        Dim intPriceList As Integer
        Dim blnDiscountflag As Boolean
        '  Dim oBP As SAPbobsCOM.BusinessPartners
        Dim objForm As SAPbouiCOM.Form
        ' Dim oRec As SAPbobsCOM.Recordset
        Dim oStatic As SAPbouiCOM.StaticText
        Dim oItem, oItem1 As SAPbouiCOM.Item
        price = 0
        discount = 0
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        oItems = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oREc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecDiscount = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        intFlag = 0
        blnDiscountflag = False
        If oBP.GetByKey(strBP) Then
            'Find discount in Special Price Table
            If 1 = 1 Then 'Take to price for BP Price List
                If oBP.PriceListNum = -1 Then
                    strSQL = "SELECT T1.[ItemCode], T1.[ItemName], isnull(T1.[LastPurPrc],0) FROM OITM  T1  where T1.Itemcode='" & StrItem & "'"
                ElseIf oBP.PriceListNum = -2 Then
                    strSQL = "SELECT T1.[ItemCode], T1.[ItemName], isnull(T1.[LstEvlPric],0) FROM OITM  T1  where T1.Itemcode='" & StrItem & "'"
                Else
                    strSQL = "SELECT T1.[ItemCode], T1.[PriceList], T1.[Price] FROM OPLN T0  INNER JOIN ITM1 T1 ON T0.ListNum = T1.PriceList where T1.Itemcode='" & StrItem & "' and T1.PriceList=" & oBP.PriceListNum
                End If

                oRec.DoQuery(strSQL)
                If oRec.RecordCount > 0 Then
                    price = Convert.ToDouble(oRec.Fields.Item(2).Value)
                End If
            End If
        End If
        ' amatrix.Columns.Item("14").Cells.Item(intRow).Specific.value = price
        oBP = Nothing
        oItem = Nothing
        Return price
    End Function
#End Region
#End Region

    Public Function CreateBankDBFFile(ByVal aBatchNumber As String, ByVal aFileName As String, ByVal aBankName As String) As Boolean
        Dim oDRow As DataRow
        Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
        Dim strQuery, strJouranlQry, strfrom, strto, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance, strCompany As Double
        oRecBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBalanceRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRec.DoQuery("Select isnull(U_BankId,'0') from ODSC where BankCode='" & aBankName & "'")
            strCompany = oRec.Fields.Item(0).Value


            Dim strfileaname, strTableName As String
            strfileaname = aFileName & "\" & aBankName & "_BatchNumber_" & aBatchNumber & ".dbf"
            If File.Exists(strfileaname) Then
                File.Delete(strfileaname)
            End If
            '  strfileaname = aFileName & "\" & aBankName & "_BatchNumber_" & aBatchNumber
            strfileaname = aFileName & "\" '& aBankName & "_" & aBatchNumber
            'Dim cn1 As New OleDbConnection("Provider=VFPOLEDB.1;Data Source=" & strfileaname & ";")
            Dim cn1 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;data source=" & strfileaname & ";Extended Properties='dBASE III';")
            cn1.Open()
            '   strTableName = aBankName & "_BatchNumber_" & aBatchNumber
            strTableName = aBankName & "_" & aBatchNumber
            Try
                Dim cmd_Delete_cgdd As New OleDbCommand("Drop Table " & strTableName, cn1)
                cmd_Delete_cgdd.ExecuteNonQuery()
                cn1.Close()
                cn1.Open()
            Catch ex As Exception

            End Try


            'Dim cmd1 As New OleDbCommand("Create Table " & strTableName & " (VOUCHER int,TOTAL numeric ,BALANCE char(1), CO_OP Char(50), BRANCH Char(1), V_DATE Char(10),VP_DATE Char(1),TYPE Char(1),COMPANY Char(20))", cn1)
            Dim cmd1 As New OleDbCommand("Create Table " & strTableName & " (VOUCHER Char(10),TOTAL Char(20) ,BALANCE char(12), CO_OP Char(3), BRANCH Char(2), V_DATE Char(10),VP_DATE Char(8),TYPE Char(1),COMPANY Char(3))", cn1)
            'Dim cmd1 As New OleDbCommand("Create Table " & strTableName & " (VOUCHER numeric(8,0),TOTAL numeric(12,3) ,BALANCE char(12), CO_OP Char(3), BRANCH Char(2), V_DATE Char(10),VP_DATE Char(8),TYPE Char(1),COMPANY Char(3))", cn1)
            '            Dim cmd1 As New OleDbCommand("Create Table cgid( _BillID char(20), _StaIndex numeric(5,0), _EndIndex numeric(5,0))", cn1)
            cmd1.ExecuteNonQuery()
            cn1.Close()
            strQuery = "Select * from [@Z_Bill_Export] where U_Z_BankCode='" & aBankName & "' and  U_Z_BatchNumber='" & aBatchNumber & "'"
            oRecTemp.DoQuery(strQuery)
            If oRecTemp.RecordCount > 0 Then
                Dim dtdate As Date
                strQuery = "  Select convert(varchar,X.DocNum),X.DocTotal,'',X.BankrefNo,1,DocDate,'',X.Type,'122' from "
                strQuery = strQuery & "(select T0.CardCode,isnull(T1.U_Z_KFHNO,'') 'BankRefNo',DocNum,DocTotal,1'Type',DocDate  from OINV  T0 inner join OCRD T1 on T0.CardCode=t1.CardCode and T1.HouseBank='" & aBankName & "' where  T0.DocStatus<>'C' and  isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "'"
                strQuery = strQuery & "and isnull(U_Z_Exported,'N')='Y' union all select T0.CardCode,isnull(T1.U_Z_KFHNO,'') 'BankRefNo',DocNum,DocTotal,2 'Type',DocDate  from ORIN  T0   inner join OCRD T1 on T0.CardCode=t1.CardCode  and T1.HouseBank='" & aBankName & "' where  T0.DocStatus<>'C' and  isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "'"
                strQuery = strQuery & "and isnull(U_Z_Exported,'N')='Y'  )   x  Order by X.CardCode"
                oRecBP.DoQuery(strQuery)
                oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                For intRow As Integer = 0 To oRecBP.RecordCount - 1
                    oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    dtdate = oRecBP.Fields.Item(5).Value
                    cn1.Open()
                    Dim stQuery As String
                    Dim dblTotal As Decimal
                    Dim docNum As Integer
                    docNum = oRecBP.Fields.Item(0).Value
                    dblTotal = oRecBP.Fields.Item(1).Value

                    'stQuery = "Insert Into " & strTableName & " Values ('" & oRecBP.Fields.Item(0).Value.ToString & "','" & oRecBP.Fields.Item(1).Value.ToString & "','" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "'," & oRecBP.Fields.Item(4).Value.ToString & ",'" & dtdate.ToString("yyyyddMM") & "','" & oRecBP.Fields.Item(6).Value.ToString & "'," & oRecBP.Fields.Item(7).Value.ToString & "," & oRecBP.Fields.Item(8).Value.ToString & ")"
                    '  stQuery = "Insert Into " & strTableName & " Values ('" & oRecBP.Fields.Item(0).Value.ToString & "'," & dblTotal.ToString.Replace(",", ".") & ",'" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "','" & dtdate.ToString("yyyy-dd-MM") & "','" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & oRecBP.Fields.Item(8).Value.ToString & ")"
                    '  stQuery = "Insert Into " & strTableName & " Values ('" & oRecBP.Fields.Item(0).Value.ToString & "'," & dblTotal.ToString.Replace(",", ".") & ",'" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "',{^" & dtdate.ToString("yyyy/dd/MM") & "},'" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & oRecBP.Fields.Item(8).Value.ToString & ")"

                    ' stQuery = "Insert Into " & strTableName & " Values (" & oRecBP.Fields.Item(0).Value.ToString & "," & dblTotal.ToString.Replace(",", ".") & ",'" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "',{^" & dtdate.ToString("yyyyddMM") & "},'" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & oRecBP.Fields.Item(8).Value.ToString & ")"
                    ' stQuery = "Insert Into " & strTableName & " Values (" & oRecBP.Fields.Item(0).Value.ToString & "," & dblTotal.ToString.Replace(",", ".") & ",'" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "'," & dtdate.ToString("yyyyMMdd") & ",'" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & oRecBP.Fields.Item(8).Value.ToString & ")"
                    ' stQuery = "Insert Into " & strTableName & " Values (" & oRecBP.Fields.Item(0).Value.ToString & ",'" & dblTotal.ToString(".000") & "','" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "','" & dtdate.ToString("yyyyMMdd") & "','" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & strCompany & ")"
                    stQuery = "Insert Into " & strTableName & " Values (" & docNum & ",'" & dblTotal.ToString(".000") & "','" & oRecBP.Fields.Item(2).Value.ToString & "','" & oRecBP.Fields.Item(3).Value.ToString & "','" & oRecBP.Fields.Item(4).Value.ToString & "','" & dtdate.ToString("dd-MM-yyyy") & "','" & oRecBP.Fields.Item(6).Value.ToString & "','" & oRecBP.Fields.Item(7).Value.ToString & "'," & strCompany & ")"
                    Dim cmd2 As New OleDbCommand(stQuery, cn1)
                    cmd2.ExecuteNonQuery()
                    cn1.Close()
                    oRecBP.MoveNext()

                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False

        End Try

    End Function
#End Region

#Region "Generate BillDiscount report"
    Public Function generateBillDiscountreport(ByVal aBatchNumber As String, ByVal aFileName As String, ByVal abankName As String) As Boolean
        Dim rptaccountreport As New BillDiscounting_RPT
        Dim ds As New BillDiscounting
        Dim oDRow As DataRow
        Dim oRec, oRecTemp, oRecBP, oBalanceRs As SAPbobsCOM.Recordset
        Dim strQuery, strbank, strJouranlQry, strBPCondition, strfrom, strto, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        Try
            oRecBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBalanceRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ds.Clear()
            ds.Clear()
            strQuery = "Select * from [@Z_Bill_Export] where U_Z_BankCode='" & abankName & "' and U_Z_BatchNumber='" & aBatchNumber & "'"
            oRecTemp.DoQuery(strQuery)
            strBPCondition = " 1=1"
            If oRecTemp.RecordCount > 0 Then
                oDRow = ds.Tables("Header").NewRow
                oDRow.Item("BatchNumber") = aBatchNumber
                oDRow.Item("BankCode") = oRecTemp.Fields.Item("U_Z_BankCode").Value
                strbank = oRecTemp.Fields.Item("U_Z_BankCode").Value
                oDRow.Item("BankName") = oRecTemp.Fields.Item("U_Z_BankName").Value
                oDRow.Item("DocDate") = oRecTemp.Fields.Item("U_Z_ExportDate").Value
                oDRow.Item("FromDate") = oRecTemp.Fields.Item("U_Z_DateFrom").Value
                oDRow.Item("ToDate") = oRecTemp.Fields.Item("U_Z_DateTo").Value
                oDRow.Item("DiscountAmount") = oRecTemp.Fields.Item("U_Z_DiscountAmount").Value
                If strbank = "" Then
                    oApplication.Utilities.Message("Select bank Code", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    strBPCondition = "Select CardCode from OCRD where HouseBank='" & strbank & "'"
                End If

                ds.Tables("Header").Rows.Add(oDRow)

                'strQuery = "Select isnull(T1.U_Z_KFHNO,''),x.CardCode,x.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from (select T0.CardCode,T0.CardName,sum(T0.DocTotal) 'INV',0 'RETU'  from OINV  T0 where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y'   group by CardCode,Cardname "
                ' strQuery = strQuery & " union select T0.CardCode,T0.CardName,0 'INV',sum(T0.DocTotal) 'RETU'  from ORIN  T0 where isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y' group by CardCode,Cardname )   x  inner join OCRD T1 on T1.CardName=x.Cardname and T1.HouseBank='" & abankName & "' group by isnull(T1.U_Z_KFHNO,''),X.CardCode,X.CardName"

                strQuery = "Select isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName,sum(x.Inv),sum(x.RETU),sum(x.INV)-sum(x.RETU) from "
                strQuery = strQuery & "(select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',sum(round(T0.DocTotal,3)) 'INV',0 'RETU'  from OINV  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode where  T0.DocStatus<>'C' and   T1.HouseBank='" & abankName & "' and isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y'   group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard "
                strQuery = strQuery & " union all select case T1.fatherType when 'P' then  isnull(T1.FatherCard,T0.CardCode) else T0.CardCode end 'CardCode',0 'INV',sum(round(T0.DocTotal,3)) 'RETU'  from ORIN  T0  inner Join OCRD T1 on T1.CardCode=T0.CardCode  where  T0.DocStatus<>'C' and  T1.HouseBank='" & abankName & "' and isnull(U_Z_BatchNumber,'')='" & aBatchNumber & "' and isnull(U_Z_Exported,'N')='Y' group by T0.CardCode,T0.Cardname,T1.FatherType,T1.FatherCard )   x  inner join OCRD T1 on T1.CardCode=x.CardCode and T1.HouseBank='" & abankName & "' group by isnull(T1.U_Z_KFHNO,''),X.CardCode,T1.CardName"


                oRecBP.DoQuery(strQuery)
                If oRecBP.RecordCount > 0 Then
                    For intRow As Integer = 0 To oRecBP.RecordCount - 1
                        oDRow = ds.Tables("Details").NewRow
                        oDRow.Item("RefNo") = oRecBP.Fields.Item(0).Value
                        oDRow.Item("CustomerName") = oRecBP.Fields.Item(2).Value
                        oDRow.Item("BatchNumber") = aBatchNumber
                        oDRow.Item("Sales") = oRecBP.Fields.Item(3).Value
                        oDRow.Item("Returns") = oRecBP.Fields.Item(4).Value
                        oDRow.Item("NetSales") = oRecBP.Fields.Item(5).Value
                        ds.Tables("Details").Rows.Add(oDRow)
                        oRecBP.MoveNext()
                    Next
                Else

                End If

                addCrystal(ds, aFileName, aBatchNumber, abankName)

            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#Region "Add Crystal Report"
    Private Sub addCrystal(ByVal ds1 As DataSet, ByVal aFileName As String, ByVal aBatchNumber As String, ByVal aBankName As String)
        Dim cryRpt As New ReportDocument
        Dim strFilename, strReportviewOption As String
        Dim strReportFileName As String = "BillDiscounting_RPT.rpt"
        strReportviewOption = "P"
        '    strReportFileName = "AcctStatement_old.rpt"
        'strFilename = System.Windows.Forms.Application.StartupPath & "\BillDiscounting"
        strFilename = aBankName & "_BatchNumber_" & aBatchNumber
        strFilename = aFileName & "\" & strFilename & ".pdf"
        Message("Report Generation processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'strFilename = strFilename & ".pdf"
        If ds1.Tables.Item("Details").Rows.Count > 0 Then
            'If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\" & strReportFileName)
            cryRpt.SetDataSource(ds1)
            Message("Report Generation processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
                Message("Report exported into PDF File", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If

        Else
            Message("No data found", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

    End Sub
#End Region

#End Region


#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTempQty As String
        strTemp = CompanyDecimalSeprator
        strTempQty = strQuantity
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTempQty)
        End Try

        ' dblQuant = Convert.ToDouble(strQuantity)
        Return dblQuant
    End Function
#End Region


#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        objEdit.String = newvalue
    End Sub
#End Region

#End Region

    Public Function GetCode(ByVal sTableName As String) As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim sQuery As String
        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sQuery = "SELECT Top 1 DocEntry FROM " & sTableName + " ORDER BY Convert(Int,DocEntry) desc"
        oRecSet.DoQuery(sQuery)
        If Not oRecSet.EoF Then
            GetCode = Convert.ToInt32(oRecSet.Fields.Item(0).Value.ToString()) + 1
        Else
            GetCode = "1"
        End If
    End Function

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        'Dim lRetCode As Integer
        'Dim sErrMsg As String
        'Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()


          
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Fill Combo"
    Public Sub FillComboBox(ByVal aCombo As SAPbouiCOM.ComboBox, ByVal aQuery As String)
        Dim oFillRS As SAPbobsCOM.Recordset
        oFillRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(intRow)
        Next
        aCombo.ValidValues.Add("", "")
        oFillRS.DoQuery(aQuery)
        For intRow As Integer = 0 To oFillRS.RecordCount - 1
            Try
                aCombo.ValidValues.Add(oFillRS.Fields.Item(0).Value, oFillRS.Fields.Item(1).Value)
            Catch ex As Exception

            End Try

            oFillRS.MoveNext()
        Next
        aCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function



End Class
