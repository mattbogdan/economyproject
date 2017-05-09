namespace CostManagementProject.Models
{
    public class YearGrowth
    {
        public double Year { get; set; }
        public double NetProfit { get; set; }
        public double SalesNetIncome { get; set; }
        public double Cost { get; set; }
        public double AverageAssets { get; set; }
        public double AverageFixedAssets { get; set; }
        public double AverageCurrentAssets { get; set; }
        public double EmployeeCount { get; set; }
        public double DeviationSum { get; set; }
        public double SpiermanCoef { get; set; }

        public YearGrowth() { }

        public YearGrowth(double year, double netProfit, double salesNetIncome, double cost, double averageAssets, double averageFixedAssets, double averageCurrentAssets, double employeeCount)
        {
            Year = year;
            NetProfit = netProfit;
            SalesNetIncome = salesNetIncome;
            Cost = cost;
            AverageAssets = averageAssets;
            AverageFixedAssets = averageFixedAssets;
            AverageCurrentAssets = averageCurrentAssets;
            EmployeeCount = employeeCount;
        }

        public YearGrowth(double year, double netProfit, double salesNetIncome, double cost, double averageAssets, double averageFixedAssets, double averageCurrentAssets, double employeeCount, double spiermanCoef)
        {
            Year = year;
            NetProfit = netProfit;
            SalesNetIncome = salesNetIncome;
            Cost = cost;
            AverageAssets = averageAssets;
            AverageFixedAssets = averageFixedAssets;
            AverageCurrentAssets = averageCurrentAssets;
            EmployeeCount = employeeCount;
            DeviationSum = NetProfit + SalesNetIncome + Cost + AverageAssets + AverageFixedAssets + AverageCurrentAssets + EmployeeCount;
            SpiermanCoef = spiermanCoef;
        }
    }
}
