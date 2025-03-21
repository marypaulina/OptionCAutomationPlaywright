namespace OptionCSMSAutomationPlayWright.Model
{
    //The reason for creating these model classes(DashboardAmount and ACutisDashboardAmount) is to structure and manage financial data efficiently within automation framework.
    //Instead of managing multiple separate variables, these classes group related financial attributes together, making the data easier to handle, access, and manipulate.
    //Following OOP principles, these classes act as blueprints for creating objects that hold transaction data.
    //If multiple parts of your project need to store or process financial data, you can reuse these classes instead of redefining the same properties everywhere.
    // Dashboard Amounts Model
    public class DashboardAmount
    {
        public decimal? CCTotalAmount { get; set; }
        public decimal? eCheckTotalAmount { get; set; }
        public string TotalAmount { get; set; } = string.Empty;
    }


    // ACutis Dashboard Amounts Model
    public class ACutisDashboardAmount
    {
        public decimal ccAmount { get; set; }
        public decimal ccCount { get; set; }
        public decimal ccServiceFeeAmount { get; set; }
        public decimal ccServiceFeeCount { get; set; }
        public decimal eCheckAmount { get; set; }
        public decimal eCheckCount { get; set; }
        public decimal eCheckServiceFeeAmount { get; set; }
        public decimal eCheckServiceFeeCount { get; set; }
        public decimal TotalAmount { get; set; }
    }

    // Report Model
    public class Report
    {
        public string Report911 { get; set; } = string.Empty;
        public string Report901 { get; set; } = string.Empty;
        public string Report904 { get; set; } = string.Empty;
        public string Report902 { get; set; } = string.Empty;
        public string Report905 { get; set; } = string.Empty;
        public string ServiceFeeAmount { get; set; } = string.Empty;
        public string AccountStatus { get; set; } = string.Empty;
        public string TotalPrimaryAccount { get; set; } = string.Empty;
        public string TotalAccountCreated { get; set; } = string.Empty;
        public string CreditCardCount { get; set; } = string.Empty;
        public string ECheckCount { get; set; } = string.Empty;
        public string CompareResults { get; set; } = string.Empty;
    }

    // Report Data Model
    public class ReportData
    {
        public string ReportValue { get; set; } = string.Empty;
        public string ComparedResult { get; set; } = string.Empty;
    }

    // Ledger Payments Model
    public class LedgerPayments
    {
        public string CCPayment { get; set; } = string.Empty;
        public string eCheckPayment { get; set; } = string.Empty;
        public string TotalPayment { get; set; } = string.Empty;
        public string TotalCharges { get; set; } = string.Empty;
    }
}