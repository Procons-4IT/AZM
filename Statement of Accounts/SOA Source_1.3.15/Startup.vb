Imports System.Threading
Imports System.Collections.Generic
Module SubMain

#Region "Declaration"
    Private objMain As ClsSBO
    Public strRefBatch As String
    Public strSplitdocentry, dbSplitReceipt As Double
    Public strCombinedocentry, dbCombineIssue As Double
    Public dtRundate As DateTime
    Public strShortname As String
    Public blnItemChoose As Boolean = False
#End Region

#Region "Main Method"
    '*****************************************************************************
    'Type               : Procedure   
    'Name               : main
    'Parameter          : 
    'Return Value       : 
    'Author             : Senthil Kumar B
    'Created Date       : 27-07-2006
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create Instance to MainClass and Initialize Applicaiton 

    '******************************************************************************
    <STAThread()> _
    Public Sub main()
        Try
            Dim objSCR As clsMain
            objSCR = New clsMain
            If (objSCR.Initialise()) Then
                objSCR.objSBOAPI.SBO_Appln.StatusBar.SetText("Addon connected successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Windows.Forms.Application.Run()
            Else
                objSCR.objSBOAPI.SBO_Appln.StatusBar.SetText("Error in Connection", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            System.Windows.Forms.Application.Exit()
        End Try
    End Sub
#End Region

#Region "Close Application"

    Public Sub CloseApp()
        System.Windows.Forms.Application.Exit()
    End Sub

#End Region

End Module
