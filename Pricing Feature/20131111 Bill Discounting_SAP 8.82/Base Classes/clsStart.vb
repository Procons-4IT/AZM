Public Class clsStart
    
    Shared Sub Main()
        Dim oRead As System.IO.StreamReader
        Dim LineIn, strUsr, strPwd As String
        Dim i As Integer
        Try
            Try
                oApplication = New clsListener
                oApplication.Utilities.Connect()
                ' oApplication.SetFilter()
                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            ' companyStorekey = oApplication.Utilities.getStoreKey()
            'If companyStorekey = "" Then
            '    oApplication.Utilities.Message("Define the storekey in the company details", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
            Dim omenu As SAPbouiCOM.MenuItem
            omenu = oApplication.SBO_Application.Menus.Item("Z_mnu_P0001")
            omenu.Image = Application.StartupPath & "\DataBest.bmp"
            oApplication.Utilities.Message("Pricing Feature Addon Connected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
