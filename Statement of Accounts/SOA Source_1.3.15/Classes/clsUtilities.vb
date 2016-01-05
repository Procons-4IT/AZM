Public Class clsUtilities

#Region "Declartion"

    Private objCompany As SAPbobsCOM.Company
    Private objclsSBO As ClsSBO

#End Region

#Region "Methods"

#Region "Constructor"

    Public Sub New(ByVal objSBO As ClsSBO)
        objclsSBO = objSBO
    End Sub

#End Region

#Region "Show Messages"
    '*****************************************************************
    'Type               : Procedure   
    'Name               : ShowMessage
    'Parameter          : Message
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Show Message with in SBO
    '*****************************************************************
    Public Sub ShowMessage(ByVal strMessage As String)
        objclsSBO.SBO_Appln.MessageBox(strMessage)
    End Sub
    '*****************************************************************
    'Type               : Procedure   
    'Name               : ShowSuccessMessage
    'Parameter          : Message
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Show Success Message in Status Bar
    '*****************************************************************

    Public Sub ShowSuccessMessage(ByVal strMessage As String)
        objclsSBO.SBO_Appln.StatusBar.SetText(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    '*****************************************************************
    'Type               : Procedure   
    'Name               : ShowErrorMessage
    'Parameter          : Message
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Show Error Message in Status Bar
    '*****************************************************************

    Public Sub ShowErrorMessage(ByVal strMessage As String)
        Try
            objclsSBO.SBO_Appln.StatusBar.SetText(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Catch ex As Exception
        End Try

    End Sub

    '*****************************************************************
    'Type               : Procedure   
    'Name               : ShowWarningMessage
    'Parameter          : Message
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Show Warning Message in Status Bar
    '*****************************************************************

    Public Sub ShowWarningMessage(ByVal strMessage As String)
        objclsSBO.SBO_Appln.StatusBar.SetText(strMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region

    Public Function GetLocalCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Maincurncy from OADM"
        oTemp = objclsSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function GetSystemCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select SysCurrncy from OADM"
        oTemp = objclsSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function getBPCurrency(ByVal strCardcode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Currency from OCRD where Cardcode='" & strCardcode & "'"
        oTemp = objclsSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function getBPCurrency_Project(ByVal strCardcode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select Top 1 Currency from OCRD where Cardcode in (" & strCardcode & ")"
        oTemp = objclsSBO.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function


#Region "DateFormat"
    '****************************************************************************
    'Type	        	    :   Procedure     
    'Name               	:   GetFormat
    'Parameter          	:   Company,Type
    'Return Value       	:	
    'Author             	:	Senthil Kumar B
    'Created Date       	:	10-Apr-2006
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	Get the Dateformat for the Given Company
    '****************************************************************************

    Public Function GetFormat(ByVal objcompany As SAPbobsCOM.Company, ByVal oType As Integer) As String
        Dim strDateFormat, strSql As String
        Dim oTemprecordset As SAPbobsCOM.Recordset
        strSql = "Select DateFormat,DateSep from OADM"
        oTemprecordset = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemprecordset.DoQuery(strSql)
        Select Case oTemprecordset.Fields.Item(0).Value
            Case 0  'dd/mm/yy'
                strDateFormat = "dd" & oTemprecordset.Fields.Item(1).Value & "MM" & oTemprecordset.Fields.Item(1).Value & "yy"
                GetFormat = 3
            Case 1 'dd/mm/yyyy'
                strDateFormat = "dd" & oTemprecordset.Fields.Item(1).Value & "MM" & oTemprecordset.Fields.Item(1).Value & "yyyy"
                GetFormat = 103
            Case 2 'mm/dd/yyyy'
                strDateFormat = "MM" & oTemprecordset.Fields.Item(1).Value & "dd" & oTemprecordset.Fields.Item(1).Value & "yy"
                GetFormat = 1
            Case 3 'yyyy/dd/mm'
                strDateFormat = "MM" & oTemprecordset.Fields.Item(1).Value & "dd" & oTemprecordset.Fields.Item(1).Value & "yyyy"
                GetFormat = 120
            Case 4 'dd/month/yyyy'
                strDateFormat = "yyyy" & oTemprecordset.Fields.Item(1).Value & "dd" & oTemprecordset.Fields.Item(1).Value & "MM"
                GetFormat = 126
            Case 5 'dd/month/yyyy'
                strDateFormat = "dd" & oTemprecordset.Fields.Item(1).Value & "MMM" & oTemprecordset.Fields.Item(1).Value & "yyyy"
                GetFormat = 130
        End Select

        If oType = 1 Then
            GetFormat = strDateFormat
        Else
            GetFormat = GetFormat
        End If
    End Function
#End Region

#End Region

End Class
