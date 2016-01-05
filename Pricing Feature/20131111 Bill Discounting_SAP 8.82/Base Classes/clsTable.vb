Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "ODSC" Or strTab = "OCRD" Or strTab = "ORCT" Or strTab = "INV1" Or strTab = "OJDT" Or strTab = "OINV" Or strTab = "ORDR") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddTables("Z_ODIS", "Discount Definition Document", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ODIS", "Z_Dis_Code", "Discount Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_ODIS", "Z_Dis_Name", "Discount Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_ODIS", "Z_Default", "Default", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_ODIS", "Z_Base", "Base Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_ODIS", "Z_Factor", "Factor", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddTables("Z_DIS1", "Discount Definition Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_DIS1", "Z_ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_DIS1", "Z_ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_DIS1", "Z_DefPack", "Default pack", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_DIS1", "Z_PriceList", "Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            'AddFields("Z_DIS1", "Z_DefPrice", "Default Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_DIS1", "Z_Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_DIS1", "Z_No_Pices", "Number of Pieces", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_DIS1", "Z_Price", "Prices", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_DIS1", "Z_Discount", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_DIS1", "Z_SellPrice", "Selling Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)


            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddTables("Z_Dis_Mapping", "Discount Mapping", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_Dis_Mapping", "Z_Dis_Code", "Discount Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_Dis_Mapping", "Z_Dis_Name", "Discount Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_Dis_Mapping", "Z_FromDate", "From date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_Dis_Mapping", "Z_ToDate", "To date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_Dis_Mapping", "Z_CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_Dis_Mapping", "Z_CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)



            AddFields("ODSC", "BatchNumber", "BatchNumber", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODSC", "CreditAc", "Credit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("ODSC", "DebitAc", "Debit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("ODSC", "CashAc", "Cash Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("ODSC", "DisRate", "Bill discount rate ", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("ODSC", "BankId", "Bank Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("ORCT", "BatchNumber", "BatchNumber", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("OJDT", "BatchNumber", "BatchNumber", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("OCRD", "Z_KFHNO", "Bank Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OINV", "Z_BatchNumber", "Bank Batch Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            addField("OINV", "Z_Exported", "Exported", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddFields("INV1", "Z_Carton", "No of Cartons", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("INV1", "Z_Pieces", "No of Pieces", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            addField("INV1", "Z_ItemType", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,F", "Standard,Free", "S")
            AddFields("INV1", "Z_Pack", "Pack", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("INV1", "Z_SPrice", "Selling Price per Carton", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("INV1", "Z_GPrice", "Gross Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            AddFields("OINV", "Z_GPrice", "Gross Price for Free Items", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("INV1", "Z_Discount", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("INV1", "Z_DisAmt", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("INV1", "Z_DisCode", "Special Price Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("INV1", "Z_DiscName", "Special Price Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)




            AddTables("Z_Bill_Export", "Export details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_Bill_Export", "Z_BankCode", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_Bill_Export", "Z_BankName", "Bank Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_Bill_Export", "Z_DateFrom", "Date From", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_Bill_Export", "Z_DateTo", "Date To", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_Bill_Export", "Z_BatchNumber", "Batch Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_Bill_Export", "Z_DiscountAmount", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_Bill_Export", "Z_ExportDate", "Exported Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_Bill_Export", "Z_Exported", "Exported", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_Bill_Export", "Z_Imported", "Converted Documents", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

         

            '---- User Defined Object's
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            CreateUDO()
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Update UFD1 set Descr='Standard' where Descr='Stanard'")

            oApplication.Utilities.Message("Initializing Database Completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("Z_ODIS", "Discount Definition-Details", "Z_ODIS", "U_Z_Dis_Code", "U_Z_Dis_Name", "Z_DIS1", SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
