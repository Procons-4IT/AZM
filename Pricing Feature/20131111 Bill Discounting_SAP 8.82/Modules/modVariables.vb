Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String

    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String
    Public strBPChoice As String = "C"
    Public strReportFilterdate As String = "RefDate"
    Public dtAgingdate As Date
    Public strImportErrorLog As String = ""
    Public companyStorekey As String = ""
    Public intSelectedMatrixrow As Integer = 0
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmDiscountmatrix As SAPbouiCOM.Matrix
    Public strShortname As String = "ShortName"
    Public frmSourceSpecialPriceForm As SAPbouiCOM.Form
    Public strInvoiceSeriesNumber As String = ""

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62

    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_itemmaster As String = "150"
    Public Const frm_BPMaster As String = "134"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_ARCreditMemo As String = "179"
    Public Const frm_PurchaseOrder As String = "142"
    Public Const frm_GRPO As String = "143"
    Public Const frm_APInvoice As String = "141"
    Public Const frm_APCreditnote As String = "181"
    Public Const frm_Invoice As String = "133"
    Public Const frm_Purchasereturn As String = "182"


    Public Const frm_CopyToForm As String = "-9876"

    Public Const frm_Export As String = "frm_BillExport"
    'Public Const frm_Export As String = "frm_Bill_Wizard"
    'Public Const frm_BillDiscount As String = "frm_Discount"
    Public Const frm_BillDiscount As String = "frm_Discount"
    Public Const frm_Validations As String = "frm_InvValidation"
    Public Const frm_ChoosefromList As String = "frm_CFL"

    Public Const frm_SpecialPrice As String = "frm_SpecialPrice"
    Public Const frm_mapping As String = "frm_mapping"
    Public Const frm_SalesOdrAging As String = "frm_Aging"

    Public Const frm_Customermapping As String = "frm_CustMapping"
    Public Const frm_DiscMapping As String = "frm_DiscMapping"
    Public Const frm_ItemCFL As String = "frm_ItemCFL"

  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
   

    ' Public Const mnu_Import As String = "Z_mnu_D003"

    Public Const mnu_Export As String = "Z_mnu_P0002"
    Public Const mnu_BillDiscount As String = "Z_mnu_P0003"
    Public Const mnu_validations As String = "Z_mnu_P0004"


    Public Const mnu_DisDefin As String = "Z_mnu_P0005"
    Public Const mnu_Dismap As String = "Z_mnu_P0006"
    Public Const mnu_SalAgeRpt As String = "Z_mnu_P0007"
    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
   
    'Public Const xml_Export As String = "frm_BillExport.xml"
    Public Const xml_Export As String = "frm_Bill_Wizard.xml"
    '  Public Const xml_BillDiscount As String = "frm_BillingDiscount.xml"
    Public Const xml_BillDiscount As String = "frm_BillPayment_Wizard.xml"
    Public Const xml_InvoiceValidation As String = "frm_InvoiceValidation.xml"

    Public Const xml_Dis_mapping As String = "frm_DiscountMapping.xml"
    Public Const xml_DisDefine As String = "frm_SpecialPrice.xml"
    Public Const xml_SalesAgingRpt As String = "frm_SOAgeing.xml"

    Public xml_CustMapping As String = "frm_BPMapping.xml"
    Public xml_DiscMapping As String = "frm_DiscMapping.xml"
    Public xml_ItemCFL As String = "frm_ItemCFL.xml"
   

End Module
