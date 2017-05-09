using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CostManagementProject.Models;

namespace CostManagementProject
{
    class TestModule
    {
        public List<YearGrowthCriteries> Run(List<YearGrowth> yearGrowths)
        {
            //List<YearGrowth> yearGrowths = GetYearGrowths();//table 1-3

            List<YearGrowthCriteries> yearsCriterieses = GetYearsGrowsCriterieses(yearGrowths);//table 3 formatted

            yearsCriterieses = GetYearsGrowthRateCriteries(yearsCriterieses);//table 4

            yearsCriterieses = getYearsRanks(yearsCriterieses);//table 5

            
            CalculateSpirman(yearsCriterieses);//table 6

           
            for (int i = 1; i < yearsCriterieses.Count; ++i)
            {
                var yearTable = yearsCriterieses[i];

                var criteriums = yearsCriterieses[i].Criteriums.OrderBy(x=>x.StandardRank).ToList();

                for (int j = 0; j < criteriums.Count; ++j)
                {
                    for (int k = j + 1; k < criteriums.Count; ++k)
                    {
                        FahnerPairCriterium fahnerPairCriterium = new FahnerPairCriterium();
                        fahnerPairCriterium.FirstCriterium = criteriums[j];
                        fahnerPairCriterium.SecondCriterium = criteriums[k];

                        fahnerPairCriterium.Value = fahnerPairCriterium.FirstCriterium.Rate.CompareTo(fahnerPairCriterium.SecondCriterium.Rate);
                        if (fahnerPairCriterium.Value == 0)
                            fahnerPairCriterium.Value = 1;

                        yearTable.FahnerPairCriteriums.Add(fahnerPairCriterium);
                    }
                }

                yearTable.FehnerSum = yearTable.FahnerPairCriteriums.Sum(x => x.Value);
                yearTable.FehnerKoef = yearTable.FehnerSum/(double)yearTable.FahnerPairCriteriums.Count;

                yearTable.ScaleLevel = ((1 + yearTable.SpirmanKoef)*(1 + yearTable.FehnerKoef))/(2*2);

            }

            return yearsCriterieses;
        }

        private static void CalculateSpirman(List<YearGrowthCriteries> yearsCriterieses)
        {
            for (var i = 1; i < yearsCriterieses.Count; ++i)
            {
                var yearCriterieses = yearsCriterieses[i];
                var n = yearCriterieses.Criteriums.Count;
                

                yearCriterieses.SpirmanKoef = 1 - yearCriterieses.GetCriteriumsDeviationsSum()* (6.0/(n*(n*n - 1)));
            }
        }

        private static List<YearGrowthCriteries> getYearsRanks(List<YearGrowthCriteries> yearsCriterieses)
        {
            for (int i = 1; i < yearsCriterieses.Count; ++i)
            {
                var yearCriteries = yearsCriterieses[i];
                var sortedRanks = yearCriteries.Criteriums.GroupBy(x=>x.Rate).Select(x=> x.Key).OrderByDescending(x => x).ToList();
                foreach (var criterium in yearCriteries.Criteriums)
                {
                    criterium.Rank = sortedRanks.FindIndex(x => x.Equals(criterium.Rate)) + 1;
                }
            }

            return yearsCriterieses;
        }


        private static List<YearGrowthCriteries> GetYearsGrowthRateCriteries(List<YearGrowthCriteries> yearsGrowsCriterieses)
        {
          
            for (int i = 1; i < yearsGrowsCriterieses.Count; ++i)
            {

                var yearGrowsCriterieses = yearsGrowsCriterieses[i];
                var yearGrowsCriteriesesPrev = yearsGrowsCriterieses[i - 1];

              
                for (int j = 0; j < yearGrowsCriterieses.Criteriums.Count; ++j)
                {
                    yearGrowsCriterieses.Criteriums[j].Rate =
                        (yearGrowsCriterieses.Criteriums[j].Value - yearGrowsCriteriesesPrev.Criteriums[j].Value)/
                        yearGrowsCriteriesesPrev.Criteriums[j].Value;
                }

                
            }

            return yearsGrowsCriterieses;
        }

        private static List<YearGrowthCriteries> GetYearsGrowsCriterieses(List<YearGrowth> yearGrowths)
        {
            List<YearGrowthCriteries> yearsGrowsCriterieses = new List<YearGrowthCriteries>();
            foreach (var yearGrowth in yearGrowths)
            {
                yearsGrowsCriterieses.Add(new YearGrowthCriteries()
                {
                    Year = yearGrowth.Year,
                    Criteriums = new List<Criterium>()
                    {
                        new Criterium() {Id = 1, Value = yearGrowth.NetProfit, StandardRank = 1},
                        new Criterium() {Id = 2, Value = yearGrowth.SalesNetIncome, StandardRank = 3},
                        new Criterium() {Id = 3, Value = yearGrowth.Cost, StandardRank = 6},
                        new Criterium() {Id = 4, Value = yearGrowth.AverageAssets, StandardRank = 4},
                        new Criterium() {Id = 5, Value = yearGrowth.AverageFixedAssets, StandardRank = 5},
                        new Criterium() {Id = 6, Value = yearGrowth.AverageCurrentAssets, StandardRank = 2},
                        new Criterium() {Id = 7, Value = yearGrowth.EmployeeCount, StandardRank = 7},
                    }
                });
            }
            return yearsGrowsCriterieses;
        }

        private static List<YearGrowth> GetYearGrowths()
        {
            return new List<YearGrowth>()
            {
                new YearGrowth()
                {
                    Year = 2012,
                    NetProfit = 3792,//"Чистий прибуток\збиток"
                    SalesNetIncome = 671554,//"Чистий дохід від реалізації"
                    Cost = 667762,//собівартість
                    AverageAssets = 318453.5,//"Середньорічна вартість активів"
                    AverageFixedAssets = 329855,//"Середньорічна вартість основних засобів"
                    AverageCurrentAssets = 201850.5,//"Середньорічна вартість оборотних активів"
                    EmployeeCount = 2093,//"Середньоспискова чисельність працівників"
                    
                },
                new YearGrowth()
                {
                    Year = 2013,
                    NetProfit = 3925,//"Чистий прибуток\збиток"
                    SalesNetIncome = 497620,//"Чистий дохід від реалізації"
                    Cost = 493695,//собівартість
                    AverageAssets = 830832.5,//"Середньорічна вартість активів"
                    AverageFixedAssets = 738413,//"Середньорічна вартість основних засобів"
                    AverageCurrentAssets = 92458.5,//"Середньорічна вартість оборотних активів"
                    EmployeeCount = 1457,//"Середньоспискова чисельність працівників"
                    
                },

                new YearGrowth()
                {
                    Year = 2014,
                    NetProfit = 34816,//"Чистий прибуток\збиток"
                    SalesNetIncome = 294354,//"Чистий дохід від реалізації"
                    Cost = 259538,//собівартість
                    AverageAssets = 800849,//"Середньорічна вартість активів"
                    AverageFixedAssets = 730892,//"Середньорічна вартість основних засобів"
                    AverageCurrentAssets = 69936,//"Середньорічна вартість оборотних активів"
                    EmployeeCount = 1383,//"Середньоспискова чисельність працівників"
                    
                },

                new YearGrowth()
                {
                    Year = 2015,
                    NetProfit = 11021,//"Чистий прибуток\збиток"
                    SalesNetIncome = 250516,//"Чистий дохід від реалізації"
                    Cost = 239495,//собівартість
                    AverageAssets = 744672,//"Середньорічна вартість активів"
                    AverageFixedAssets = 654011,//"Середньорічна вартість основних засобів"
                    AverageCurrentAssets = 90640,//"Середньорічна вартість оборотних активів"
                    EmployeeCount = 757,//"Середньоспискова чисельність працівників"
                    
                },

                new YearGrowth()
                {
                    Year = 2016,
                    NetProfit = 12905.59,//"Чистий прибуток\збиток"
                    SalesNetIncome = 298865.6,//"Чистий дохід від реалізації"
                    Cost = 220095.9,//собівартість
                    AverageAssets = 836266.7,//"Середньорічна вартість активів"
                    AverageFixedAssets = 654011,//"Середньорічна вартість основних засобів"
                    AverageCurrentAssets = 90640,//"Середньорічна вартість оборотних активів"
                    EmployeeCount = 757,//"Середньоспискова чисельність працівників"
                    
                },
            };
        }
    }
}
