
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
        objSBOAPI.LoadMenu("xml\Royalty_Menus.xml")

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


        Return True
    End Function
#End Region

#End Region

End Class
