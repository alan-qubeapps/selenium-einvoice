namespace ERPPlus.IntegrationTests.Config
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
        // Version Number
        public static string FooterValue = "";

        // Base URL for your API
        //public static string BaseUrl => "https://qubeappstest1.einvoice-staging.qubeposcloud-uatapi.com";
        //public static string BaseUrl => "https://test.einvoice-dev.qubeposcloud-uatapi.com";
        public static string BaseUrl => "https://diy.api.qubeposcloud-uatapi.com";

        // Replace with your actual login endpoint
        public static string TokenUrl => $"{BaseUrl}/oauth/token";
        public static string InvoiceUrl => $"{BaseUrl}/api/transaction/invoice";

        // Replace with real credentials or load from ENV
        public static string Username => "yanshen.choo@qubeapps.com";
        public static string Password => "Password123!";
        public static string ClientId => "9ea009d4-4f12-4aeb-8616-17bd5b18ccf3";
        public static string ClientSecret => "kSIUMifGz4p50jt4PirEeJcL4gjY3OXYTEileIvl";


        // Authentication token (if needed)
        public static string AuthToken { get; set; }


        //Exported Test Case File Path
        public static string CsvExportFolder => @"C:\Users\ChooYanShen\Desktop\E-Invoice\TestCase";

        //Test Case Template
        public static string TestCaseFile = @"D:\e-invoice\SeleniumTests\TestCaseTemplate.xlsx";

        //Test Data Template
        public static string TestDataFolder = @"D:\e-invoice\SeleniumTests\TestDataFolder";
    }
}
