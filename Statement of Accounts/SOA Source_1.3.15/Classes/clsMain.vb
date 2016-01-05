
Public Class clsMain

#Region "Declaration"
    Public objSBOAPI As ClsSBO
    'Private objImport As clsImport
    Private objutity As clsUtilities
    Private objForm As SAPbouiCOM.Form
#End Region

#Region "Methods"

    Public Sub New()
        objSBOAPI = New ClsSBO
        objutity = New clsUtilities(objSBOAPI)
    End Sub

#Region "Initialise"
    '*****************************************************************
    'Type               :  Function    
    'Name               : Initialise
    'Parameter          :
    'Return Value       : Boolean
    'Author             : DEV-4
    'Created Date       : 05-10-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Initialise the Application and Create Table
    '******************************************************************
    Public Function Initialise() As Boolean
        If (Not objSBOAPI.Connect()) Then Return False
        If (Not createtables()) Then Return False
        objSBOAPI.SetFilter()
        objSBOAPI.LoadMenu("xml\AccountBalance_Menus.xml")

        Return True
    End Function
#End Region

#Region "Create Table"



    '*****************************************************************
    'Type               : Function
    'Name               : Create Table
    'Parameter          :
    'Return Value       : Boolean
    'Author             : DEV-4
    'Created Date       : 05-10-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To add table and Fields
    '******************************************************************
    Private Function createtables() As Boolean
        Dim oProgressBar As SAPbouiCOM.ProgressBar

        Try
            '  objUtility.ShowSuccessMessage("Addon  connected successfully")
            objSBOAPI.AddAlphaField("OPRJ", "Address", "Address", 100)
            objSBOAPI.AddAlphaField("OPRJ", "Telephone", "Telephone", 100)
            objSBOAPI.AddAlphaField("OPRJ", "Fax", "Fax", 100)
            objSBOAPI.AddAlphaField("OPRJ", "ContPerson", "Contact Person", 100)
            objSBOAPI.AddAlphaField("JDT1", "cardcode", "Consolidate Shortname", 20)
            objSBOAPI.AddAlphaField("JDT1", "SMNo", "Sales man number", 10)

            Return True
        Catch ex As Exception

            MsgBox(ex.Message)
            Return False
        Finally
        End Try

        Return True
    End Function
#End Region

#End Region

End Class
