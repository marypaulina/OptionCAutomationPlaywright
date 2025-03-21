using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OptionCSMSAutomationPlayWright.Model
{
    // Dashboard Amounts Model
   /* public class DashboardAmount
    {
        public decimal CCTotalAmount { get; set; }
        public decimal eCheckTotalAmount { get; set; }
        public string TotalAmount { get; set; } = string.Empty;
    }*/
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