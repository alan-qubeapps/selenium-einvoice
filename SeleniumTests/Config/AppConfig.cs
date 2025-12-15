using OpenQA.Selenium.DevTools.V136.Runtime;

namespace EInvoice.SeleniumTests.Config
{
    public static class AppConfig
    {


        //-----------------------------------------------------------------------------------------------------------------------------------------------//
        //                                                                Declaration                                                                    //
        //-----------------------------------------------------------------------------------------------------------------------------------------------//

        public static string TesterName = "Choo Yan Shen";
        public static string FEDeveloperName = "Fahmy";
        public static string BEDeveloperName = "Lucas";
        public static string ManagerName = "Alan Ong";
        public static string ClientName = "";
        public static string ChangeDesc = "";


        //-----------------------------------------------------------------------------------------------------------------------------------------------//
        //                                                                    Dev Env                                                                    //
        //-----------------------------------------------------------------------------------------------------------------------------------------------//

        //Dev Env
        //public static string BaseUrl => "https://test.einvoice-dev.qubeposcloud-uat.com";
        //public static string UserName => "test@einvoice.com";
        //public static string Password => "password";

        //Dev Env
        //public static string BaseUrl => "https://test.einvoice-dev.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password1234!";

        //Dev Env
        public static string BaseUrl => "https://test.einvoice-dev.qubeposcloud-uat.com";
        public static string UserName => "yanshen.choo@qubeapps.com";
        public static string Password => "Password123!";

        //-----------------------------------------------------------------------------------------------------------------------------------------------//
        //                                                                Staging Env                                                                    //
        //-----------------------------------------------------------------------------------------------------------------------------------------------//


        //Staging Env diy (All type of log)
        //public static string BaseUrl => "https://diy.einvoice-staging.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password1234!";

        //Staging Env qubeappstest1 super admin
        //public static string BaseUrl => "https://qubeappstest1.einvoice-staging.qubeposcloud-uat.com";
        //public static string UserName => "support@qubeappstest1.com";
        //public static string Password => @"]alQa)-$\A";

        //Staging Env qubeappstest1
        //public static string BaseUrl => "https://qubeappstest1.einvoice-staging.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password123!";


        //Staging Env qubeappstest1 (Reset Password)
        //public static string BaseUrl => "https://qubeappstest1.einvoice-staging.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password1234!";


        //-----------------------------------------------------------------------------------------------------------------------------------------------//
        //                                                            Diy Staging Env                                                                    //
        //-----------------------------------------------------------------------------------------------------------------------------------------------//


        //Diy Staging Env diy
        //public static string BaseUrl => "https://diy.einv-diy-staging.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password1234!";

        //Diy Staging Env diy
        //public static string BaseUrl => "https://diy.einv-diy-staging.qubeposcloud-uat.com";
        //public static string UserName => "yanshen.choo@qubeapps.com";
        //public static string Password => "Password123!";


        //Diy Staging Env diy (Main)
        //public static string BaseUrl => "https://diy.einv-diy-staging.qubeposcloud-uat.com";
        //public static string UserName => "test@einvoice.com";
        //public static string Password => "password";



        //-----------------------------------------------------------------------------------------------------------------------------------------------//
        //                                                                  File Path                                                                    //
        //-----------------------------------------------------------------------------------------------------------------------------------------------//


        //Recording File Path
        public static string BaseVideoFolder => @"C:\Users\ChooYanShen\Desktop\E-Invoice\E-Invoice Testing Video";

        //Exported Test Case File Path
        public static string CsvExportFolder => @"C:\Users\ChooYanShen\Desktop\E-Invoice\TestCase";

        //Test Case Template
        public static string TestCaseFile = @"D:\e-invoice\SeleniumTests\TestCaseTemplate.xlsx";

        //Test Data Template
        public static string TestDataFolder = @"D:\e-invoice\SeleniumTests\TestDataFolder";

        //Downloaded File Path
        public static string DownloadPath = @"C:\Users\ChooYanShen\Downloads";

        //Import Template Empty File
        //public static string ImportBECSVFile = @"D:\e-invoice\SeleniumTests\Import Template Without Data\supplier.csv";
        //public static string ImportStoreCSVFile = @"D:\e-invoice\SeleniumTests\Import Template Without Data\Store Excel Sheet-Template.csv";   
        //public static string ImportCustomerCSVFile = @"D:\e-invoice\SeleniumTests\Import Template Without Data\Customer Excel Sheet-Template.csv";


        //Import Template Data File
        // add action download
        public static string ImportBECSVFile = @"D:\e-invoice\SeleniumTests\Import Template\supplier.csv";
        public static string ImportStoreCSVFile = @"D:\e-invoice\SeleniumTests\Import Template\Store Excel Sheet-Template.csv";
        public static string ImportCustomerCSVFile = @"D:\e-invoice\SeleniumTests\Import Template\Customer Excel Sheet-Template.csv";


        //Import Transaction
        public static string ImportB2CTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\b2c_invoice_template.csv";
        public static string ImportB2CTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\b2c_invoice_template.csv";
        public static string ImportB2BInvoiceTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\invoice_template.csv";
        public static string ImportB2BInvoiceTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\invoice_template.csv";
        public static string ImportB2BRefundTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\refund_note_template.csv";
        public static string ImportB2BRefundTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\refund_note_template.csv";
        public static string ImportB2BCreditNoteTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\credit_note_template.csv";
        public static string ImportB2BCreditNoteTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\credit_note_template.csv";
        public static string ImportB2BDebitNoteTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\debit_note_template.csv";
        public static string ImportB2BDebitNoteTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\debit_note_template.csv";
        public static string ImportB2BSBInvoiceTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\self_billed_invoice_template.csv";
        public static string ImportB2BSBInvoiceTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\self_billed_invoice_template.csv";
        public static string ImportB2BSBRefundTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\self_billed_refund_note_template.csv";
        public static string ImportB2BSBRefundTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\self_billed_refund_note_template.csv";
        public static string ImportB2BSBCreditNoteTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\self_billed_credit_note_template.csv";
        public static string ImportB2BSBCreditNoteTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\self_billed_credit_note_template.csv";
        public static string ImportB2BSBDebitNoteTransactionNDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\Without Data\self_billed_debit_note_template.csv";
        public static string ImportB2BSBDebitNoteTransactionWDCSVFile = @"D:\e-invoice\SeleniumTests\Import Transaction\With Data\self_billed_debit_note_template.csv";

        //Image Path
        public static string SampleReceiptImage = @"D:\e-invoice\SeleniumTests\Image\SampleReceipt.jpg";
        public static string UserProfileImage = @"D:\e-invoice\SeleniumTests\Image\UserProfileImage.png";
        public static string BusinessEntityImage = @"D:\e-invoice\SeleniumTests\Image\BusinessEntity.jpg";


        // Should think how to cater negative scenario for export (eg. export the filter data only after apply the filter)

        
    }
}
