Public Class ClsSBO

#Region "Declaration"
    Dim x As SAPbouiCOM.ApplicationClass

    Public WithEvents SBO_Appln As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company

    Private objrCalc As clsRoyaltyCalculation
    Private objRecommdation As clsRecommandation

    Public sSearchList As String
    Public LastErrorDescription As String
    Public LastErrorCode As Integer
    Public intRow As Integer
    Private strFormUID As String
    Public strItemcode, strSourcecolumn, strFormtype, strPono As String
    Public bolScanboxidchoice As Boolean = False
    Private SboGuiApi As SAPbouiCOM.SboGuiApi
    Private objApplication As SAPbouiCOM.Application
    Private objCompany As SAPbouiCOM.Application

    Private objform As SAPbouiCOM.Form
    Private objEdit As SAPbouiCOM.EditText
#End Region

#Region "Methods"


#Region "Application Initialization"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : SetApplication
    'Parameter          : 
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To set SBO Application
    '******************************************************************

    Private Sub SetApplication()
        Dim sConnectionString As String = Environment.GetCommandLineArgs.GetValue(1)
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        SboGuiApi.Connect(sConnectionString)
        SBO_Appln = SboGuiApi.GetApplication()

    End Sub
#End Region

#Region "Connect Company"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : ConnectCompany
    'Parameter          : 
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Connect SBO Company
    '******************************************************************
    Private Function ConnectCompany() As Boolean
        Dim connectstr As String
        Dim lngConnect As Long
        Dim rsCompany As SAPbobsCOM.Recordset
        oCompany = New SAPbobsCOM.Company
        Dim sUsers As SAPbobsCOM.Users
        Dim usr1 As SAPbobsCOM.Users
        Dim conectstr2 As String
        Dim ocookies As String
        Dim ocookiecontext As String
        Try
            oCompany = New SAPbobsCOM.Company
            ocookies = oCompany.GetContextCookie
            ocookiecontext = SBO_Appln.Company.GetConnectionContext(ocookies)
            oCompany.SetSboLoginContext(ocookiecontext)
            If oCompany.Connect <> 0 Then
                SBO_Appln.StatusBar.SetText("Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                ' SBO_Appln.StatusBar.SetText("Connection fails", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            SBO_Appln.MessageBox(ex.Message)
        End Try
        oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        Return True
    End Function

    Public Sub New()

    End Sub

    Public Function Connect() As Boolean
        If (Not initialiseApplication()) Then
            Return False
        End If
        If (Not ConnectCompany()) Then Return False
        createobjects()
        Return True
    End Function

    Public Function initialiseApplication() As Boolean
        Try
            Dim strConstr As String
            Dim objGUI As SAPbouiCOM.SboGuiApiClass
            objGUI = New SAPbouiCOM.SboGuiApiClass
            strConstr = System.Environment.GetCommandLineArgs(1)
            objGUI.Connect(strConstr)
            objApplication = objGUI.GetApplication()
            SBO_Appln = objApplication
        Catch ex As Exception
            LastErrorCode = -100001
            LastErrorDescription = ex.Message
            Return False
        End Try

        Return True
    End Function

    Public Function initialiseCompany() As Boolean
        Dim strCookie As String
        Dim strConStr As String
        Dim intReturnCode As Integer
        objCompany = New SAPbobsCOM.Company
        strCookie = objCompany.GetContextCookie()
        strConStr = objApplication.Company.GetConnectionContext(strCookie)
        objCompany.SetSboLoginContext(strConStr)
        intReturnCode = objCompany.Connect()
        If (intReturnCode <> 0) Then
            updateLastErrorDetails(-102)
            Return False
        End If

        Return True


    End Function

#End Region

#Region "Update Error Details"
    Private Sub updateLastErrorDetails(ByVal ErrorCode As Integer)
        LastErrorCode = ErrorCode
        LastErrorDescription = oCompany.GetLastErrorCode() & ":" & oCompany.GetLastErrorDescription()

    End Sub
#End Region

#Region "Filters"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : Filters
    'Parameter          : EventFilters
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To set Event Filters to the Application
    '******************************************************************

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        SBO_Appln.SetFilter(Filters)
    End Sub

    '*****************************************************************
    'Type               : Procedure    
    'Name               : Filters
    'Parameter          : 
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To set Event Filters to the Application
    '*****************************************************************
    Public Sub SetFilter()
        Dim objFilters As SAPbouiCOM.EventFilters
        Dim objFilter As SAPbouiCOM.EventFilter
        objFilters = New SAPbouiCOM.EventFilters

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        objFilter.AddEx("RCalculation")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        objFilter.AddEx("RCalculation")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        objFilter.AddEx("RCalculation")
        objFilter.AddEx("Rec")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        objFilter.AddEx("RCalculation")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        objFilter.AddEx("RCalculation")

        SetFilter(objFilters)

    End Sub

#End Region

#Region "Create Objects"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : createObjects
    'Parameter          : 
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instants for the Class
    '*****************************************************************
    Private Sub createobjects()
        objrCalc = New clsRoyaltyCalculation(Me)
        objRecommdation = New clsRecommandation(Me)
    End Sub

#End Region

#Region "Database Function"

#Region "Get Code"
    '*****************************************************************
    'Type               : Function    
    'Name               : GetCode
    'Parameter          : Tablename
    'Return Value       : Maximum Code value in String Format
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 16-08-2006
    'Purpose            : To Get Maximum Code field value for given Table
    '*****************************************************************
    Public Function GetCode(ByVal sTableName As String) As String
        Dim oRec As SAPbobsCOM.Recordset
        Dim sQuery As String
        Dim intCode As Integer
        Try
            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sQuery = "SELECT Top 1 Code FROM " & sTableName + " ORDER BY Convert(Int,Code) desc"
            oRecSet.DoQuery(sQuery)
            If Not oRecSet.EoF Then
                GetCode = Convert.ToInt32(oRecSet.Fields.Item(0).Value.ToString()) + 1
            Else
                GetCode = "1"
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try
    End Function

#End Region

#Region "Add Column"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddCol
    'Parameter          : Tablename,FieldName,TableDescription,FieldType,Size,Sub Field Type
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Field to Table
    '*****************************************************************
    Private Sub AddCol(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As Integer, Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As Integer = 0)

        Dim oUFields As SAPbobsCOM.UserFieldsMD
        Dim nError As Integer

        oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUFields.TableName = strTab
        oUFields.Name = strCol
        oUFields.Type = nType
        oUFields.SubType = nSubType
        oUFields.Description = strDesc
        oUFields.EditSize = nEditSize
        nError = oUFields.Add()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        If nError <> 0 Then
            'MsgBox(strCol & " table could not be added")
        End If
    End Sub

#End Region

#Region "Create Table"
    '*****************************************************************
    'Type               : Function    
    'Name               : CreateTable
    'Parameter          : Tablename,TableDescription,TableType
    'Return Value       : Boolean
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create New  Table
    '*****************************************************************

    Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim nError As Integer
        Dim strColname As String
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                End If
            Else
                '  oCompany.GetLastError(nError, strColname)
                '   MsgBox(strColname)
                Return False
            End If
        Catch ex As Exception
            SBO_Appln.MessageBox(ex.Message)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()

        End Try
    End Function

#End Region

#Region "Field Creations"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddAlphaField
    'Parameter          : Tablename,FieldName,TableDescription,Size
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Alphabet Field to Table
    '*****************************************************************

    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddAlphaMemoField
    'Parameter          : Tablename,FieldName,TableDescription,Size
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Alphabet Memo Field to Table
    '*****************************************************************


    Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)

        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddAlphaField
    'Parameter          : Tablename,FieldName,TableDescription,Size,Validvalues,Description,Default Value
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Alpha Field to Table and add Validvalues and set Default Values
    '*****************************************************************
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '*****************************************************************
    'Type               : Procedure    
    'Name               : addField
    'Parameter          : Tablename,FieldName,columnDescription,FieldType,Size,SubType,Validvalues,Description,Default Value
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Field to Table 
    '*****************************************************************

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

            objUserFieldMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not isColumnExist(TableName, ColumnName)) Then
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
                For intLoop = 1 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    updateLastErrorDetails(-104)
                End If
            Else
                'Dim objRecordset As SAPbobsCOM.Recordset
                'objRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRecordset.DoQuery("SELECT FieldID FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'")
                'If (Not objRecordset.EoF) Then
                '    objUserFieldMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                '    If (objUserFieldMD.GetByKey(TableName, Convert.ToInt64(objRecordset.Fields.Item(0).Value))) Then
                '        objUserFieldMD.Type = FieldType
                '        If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                '            objUserFieldMD.Size = Size
                '        Else
                '            objUserFieldMD.EditSize = Size
                '        End If
                '        objUserFieldMD.SubType = SubType
                '        objUserFieldMD.DefaultValue = SetValidValue
                '        For intLoop = 1 To strValue.GetLength(0) - 1
                '            objUserFieldMD.ValidValues.Value = strValue(intLoop)
                '            objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                '            objUserFieldMD.ValidValues.Add()
                '        Next
                '        If (objUserFieldMD.Update() <> 0) Then
                '            MsgBox(oCompany.GetLastErrorDescription)
                '        End If


                '    End If


                'End If


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddNumericField
    'Parameter          : Tablename,FieldName,columnDescription,Size
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add  Numeric Field to Table 
    '*****************************************************************

    Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    '******************************************************************************************************
    'Type               : Procedure    
    'Name               : AddNumericField
    'Parameter          : Tablename,FieldName,ColumnDescription,Size,Validvalues,Description,Default Value
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add Numeric Field to Table and add Validvalues and set Default Values
    '********************************************************************************************************

    Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal DefultValue As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, DefultValue)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddFloatField
    'Parameter          : Tablename,FieldName,columnDescription,SubType
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add  Float Field to Table 
    '*****************************************************************

    Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    '*****************************************************************
    'Type               : Procedure    
    'Name               : AddDateField
    'Parameter          : Tablename,FieldName,columnDescription,SubType
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Add  Date Field to Table 
    '*****************************************************************

    Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    '*****************************************************************
    'Type               : Function   
    'Name               : isColumnExist
    'Parameter          : Tablename,FieldName
    'Return Value       : Boolean
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Verify the Given Field already Exists or not
    '*****************************************************************

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'")
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
            'Return (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0)
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try

    End Function

#End Region

#End Region

#Region "Load Form"
    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file based on FormType
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & SBO_Appln.Forms.Count.ToString)
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
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
            objFormCreationParams = (objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return objApplication.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function

    '*****************************************************************
    'Type               : Procedure   
    'Name               : LoadForm
    'Parameter          : XmlFile
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load SBO Form
    '*****************************************************************

    Public Sub LoadFromXML(ByRef FileName As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        Dim sPath As String
        Try
            oXmlDoc.Load(FileName)
            SBO_Appln.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    '*****************************************************************
    'Type               : Procedure   
    'Name               : LoadMenu
    'Parameter          : XmlFile
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load Menu Item 
    '*****************************************************************

    Public Sub LoadMenu(ByVal XMLFile As String)
        Dim oXML As System.Xml.XmlDocument
        Dim strXML As String
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            strXML = oXML.InnerXml()
            objApplication.LoadBatchActions(strXML)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "DI /UI Methods"

#Region "GetDateTime"

    '*****************************************************************
    'Type               : Function   
    'Name               : GetDateTimeValue
    'Parameter          : DateString
    'Return Value       : DateFormate
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Convert given string into dateTime Format
    '*****************************************************************

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : GetSBODateString
    'Parameter          : DateTime
    'Return Value       : String
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Convert given  dateTime Format into string format
    '*****************************************************************
    Public Function GetSBODateString(ByVal DateVal As DateTime) As String
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
    End Function


#End Region

#Region "Business Objects"

    '*****************************************************************
    'Type               : Function   
    'Name               : GetBusinessObject
    'Parameter          : BOobjectTypes
    'Return Value       : Object
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instance to the give object
    '*****************************************************************
    Public Function GetBusinessObject(ByVal ObjectType As SAPbobsCOM.BoObjectTypes) As Object
        Return oCompany.GetBusinessObject(ObjectType)
    End Function


    '*****************************************************************
    'Type               : Function   
    'Name               : CreateUIObject
    'Parameter          : BOCreatableobjectType
    'Return Value       : Object
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instance to the give UIObject
    '*****************************************************************
    Public Function CreateUIObject(ByVal Type As SAPbouiCOM.BoCreatableObjectType) As Object
        Return objApplication.CreateObject(Type)
    End Function

#End Region

#Region "Form Objects"


    '*****************************************************************
    'Type               : Function   
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : Form
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Get SBOForm object for given FormUID
    '*****************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return SBO_Appln.Forms.Item(FormUID)
    End Function

    '************************************************************************
    'Type               : Function   
    'Name               : GetForm
    'Parameter          : FormType,Count
    'Return Value       : Form
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Get SBOForm object for given FormType,FormTypecount
    '****************************************************************************
    Public Function GetForm(ByVal Type As String, ByVal Count As Integer) As SAPbouiCOM.Form
        Return SBO_Appln.Forms.GetForm(Type, Count)
    End Function

#End Region

#Region "GetEditTextValue"
    '*****************************************************************
    'Type               : Function   
    'Name               : GetEditText
    'Parameter          : SBOForm,ItemUID / FormUID,ItemUID
    'Return Value       : String
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Return Edit Text Value
    '*****************************************************************
    Public Function GetEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aUID As String) As String
        objEdit = aForm.Items.Item(aUID).Specific
        Return Convert.ToString(objEdit.Value)
    End Function
    Public Function GetEditText(ByVal aFormUID As String, ByVal aUID As String) As String
        objform = SBO_Appln.Forms.Item(aFormUID)
        objEdit = objform.Items.Item(aUID).Specific
        Return Convert.ToString(objEdit.Value)
    End Function
#End Region

#Region "SetEditTextValue"
    '*****************************************************************
    'Type               : Procedure
    'Name               : SetEditText
    'Parameter          : SBOForm,ItemUID,Value / SBOFormUID,ItemUID,value
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To set Value to Edit Text Box
    '*****************************************************************

    Public Sub SetEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aUID As String, ByVal aVal As String)
        objEdit = aForm.Items.Item(aUID).Specific
        objEdit.Value = aVal
    End Sub
    Public Sub SetEditText(ByVal aFormUID As String, ByVal aUID As String, ByVal aVal As String)
        objform = SBO_Appln.Forms.Item(aFormUID)
        objEdit = objform.Items.Item(aUID).Specific
        objEdit.Value = aVal
    End Sub

#End Region

#Region "Get Tax Rate"
    '*****************************************************************
    'Type               : Function   
    'Name               : GetTaxRate
    'Parameter          : StrCode
    'Return Value       : string
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Get Tax Value for give Item code
    '*****************************************************************
    Public Function GetTaxRate(ByVal strCode As String) As String
        Dim rsCurr As SAPbobsCOM.Recordset
        Dim strsql, GetTaxRate1 As String
        strsql = ""
        GetTaxRate1 = ""
        strsql = "Select rate from OVTG where code='" & strCode + "'"
        rsCurr = GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsCurr.DoQuery(strsql)
        GetTaxRate1 = rsCurr.Fields.Item(0).Value
        Return GetTaxRate1
    End Function

#End Region

#End Region

#End Region

#Region "Events"

#Region "Item Event"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : SBO_Appln_ItemEvent
    'Parameter          : FormUID, ItemEvent, BubbleEvent
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Handle Item Level Events
    '******************************************************************

    Private Sub SBO_Appln_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Appln.ItemEvent
        BubbleEvent = True

        Select Case pVal.FormTypeEx
            Case "RCalculation"
                If Not (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    objform = GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    objrCalc.SBO_Appln_ItemEvent(FormUID, pVal, BubbleEvent, objform)
                End If
            Case "Rec"
                If Not (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) Then
                    objform = GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    objRecommdation.SBO_Appln_ItemEvent(FormUID, pVal, BubbleEvent, objform)
                End If

        End Select
        
    End Sub
#End Region

#Region "Menu Events"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : SBO_Appln_MenuEvent
    'Parameter          : MenuEvent, BubbelEven
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Handle Menu Events
    '******************************************************************

    Private Sub SBO_Appln_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Appln.MenuEvent
        Try
            If (pVal.BeforeAction = False) Then
                If (pVal.MenuUID = "RCalculation") Then
                    objrCalc.SBO_Appln_MenuEvent(pVal, BubbleEvent)
                ElseIf pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
                    If (SBO_Appln.Forms.ActiveForm.TypeEx = "RCalculation") Then
                        objrCalc.SBO_Appln_MenuEvent(pVal, BubbleEvent)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub
#End Region

#Region "Application Event"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : SBO_Appln_AppEvent
    'Parameter          : Application Event Type
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Handle Application Event
    '******************************************************************
    Private Sub SBO_Appln_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Appln.AppEvent
        If (EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown) Then
            LoadFromXML(System.Windows.Forms.Application.StartupPath & "\XML\B1st01_Remove Menus.xml")
            System.Windows.Forms.Application.Exit()
        End If
    End Sub

#End Region

#End Region

End Class
