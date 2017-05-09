using CostManagementProject.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using MahApps.Metro.Controls;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using Microsoft.Research.DynamicDataDisplay;
using Microsoft.Research.DynamicDataDisplay.DataSources;
using Microsoft.Research.DynamicDataDisplay.PointMarkers;
using CostManagementProject.Internal;

namespace CostManagementProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow 
    {
        public ObservableCollection<YearGrowth> YearStats = new ObservableCollection<YearGrowth>();
        public ObservableCollection<YearGrowth> YearGrowthCriteriaStats = new ObservableCollection<YearGrowth>();
        public ObservableCollection<YearGrowth> YearGrowthRateStats = new ObservableCollection<YearGrowth>();
        public ObservableCollection<YearGrowth> YearSpiermanStats = new ObservableCollection<YearGrowth>();
        public ObservableCollection<FehnerCompare> FehnerStats = new ObservableCollection<FehnerCompare>();
        public ObservableCollection<CompanyScale> ScaleStats = new ObservableCollection<CompanyScale>();
        private int yearsCount;
        public int YearsCount
        {
            get
            {
                return yearsCount;
            }
            set
            {
                yearsCount = value;
                ElementsCountLabel.Content = value;
            }
        }
        public MainWindow()
        {
            InitializeComponent();

            YearStats.Add(new YearGrowth()
            {
                Year = 2012,
                NetProfit = 3792,//"Чистий прибуток\збиток"
                SalesNetIncome = 671554,//"Чистий дохід від реалізації"
                Cost = 667762,//собівартість
                AverageAssets = 318453.5,//"Середньорічна вартість активів"
                AverageFixedAssets = 329855,//"Середньорічна вартість основних засобів"
                AverageCurrentAssets = 201850.5,//"Середньорічна вартість оборотних активів"
                EmployeeCount = 2093,//"Середньоспискова чисельність працівників"

            });

            YearStats.Add(new YearGrowth()
            {
                Year = 2013,
                NetProfit = 3925,//"Чистий прибуток\збиток"
                SalesNetIncome = 497620,//"Чистий дохід від реалізації"
                Cost = 493695,//собівартість
                AverageAssets = 830832.5,//"Середньорічна вартість активів"
                AverageFixedAssets = 738413,//"Середньорічна вартість основних засобів"
                AverageCurrentAssets = 92458.5,//"Середньорічна вартість оборотних активів"
                EmployeeCount = 1457,//"Середньоспискова чисельність працівників"

            });

            YearStats.Add(new YearGrowth()
            {
                Year = 2014,
                NetProfit = 34816,//"Чистий прибуток\збиток"
                SalesNetIncome = 294354,//"Чистий дохід від реалізації"
                Cost = 259538,//собівартість
                AverageAssets = 800849,//"Середньорічна вартість активів"
                AverageFixedAssets = 730892,//"Середньорічна вартість основних засобів"
                AverageCurrentAssets = 69936,//"Середньорічна вартість оборотних активів"
                EmployeeCount = 1383,//"Середньоспискова чисельність працівників"

            });

            YearStats.Add(new YearGrowth()
            {
                Year = 2015,
                NetProfit = 11021,//"Чистий прибуток\збиток"
                SalesNetIncome = 250516,//"Чистий дохід від реалізації"
                Cost = 239495,//собівартість
                AverageAssets = 744672,//"Середньорічна вартість активів"
                AverageFixedAssets = 654011,//"Середньорічна вартість основних засобів"
                AverageCurrentAssets = 90640,//"Середньорічна вартість оборотних активів"
                EmployeeCount = 757,//"Середньоспискова чисельність працівників"

            });

            YearStats.Add(new YearGrowth()
            {
                Year = 2016,
                NetProfit = 12905.59,//"Чистий прибуток\збиток"
                SalesNetIncome = 298865.6,//"Чистий дохід від реалізації"
                Cost = 220095.9,//собівартість
                AverageAssets = 836266.7,//"Середньорічна вартість активів"
                AverageFixedAssets = 654011,//"Середньорічна вартість основних засобів"
                AverageCurrentAssets = 90640,//"Середньорічна вартість оборотних активів"
                EmployeeCount = 757,//"Середньоспискова чисельність працівників"

            });

            YearsCount = YearStats.Count;
            YearGrowthGrid.ItemsSource = YearStats;
            YearGrowthCriteriaGrid.ItemsSource = YearGrowthCriteriaStats;
            YearGrowthRateGrid.ItemsSource = YearGrowthRateStats;
            YearSpirmanGrid.ItemsSource = YearSpiermanStats;
            FehnerGrid.ItemsSource = FehnerStats;
            ScaleRateGrid.ItemsSource = ScaleStats;
        }



        private void AddButtonClick(object sender, RoutedEventArgs e)
        {
            YearsCount++;
            YearStats.Add(new YearGrowth());
        }

        private void SubButtonClick(object sender, RoutedEventArgs e)
        {
            if (YearsCount == 1)
                return;

            YearsCount--;
            YearStats.RemoveAt(YearStats.Count - 1);
        }

        private void ClearButtonClick(object sender, RoutedEventArgs e)
        {
            ScaleStats.Clear();
            FehnerStats.Clear();
            YearSpiermanStats.Clear();
            YearGrowthRateStats.Clear();
            YearGrowthCriteriaStats.Clear();
            YearStats.Clear();
            YearsCount = YearStats.Count;
        }

        private void CalculateButtonClick(object sender, RoutedEventArgs e)
        {
            if (YearStats.ToList().Count == 0)
                return;

            var calcsList = new TestModule().Run(YearStats.ToList());
            calcsList.RemoveAt(0);

            GenerateCostGraph();

            ExportToExcelButton.IsEnabled = true;

            foreach (YearGrowthCriteries year in calcsList)
            {
                var tempRate = new YearGrowth(year.Year, 
                    Math.Round(year.Criteriums[0].Rate,3), 
                    Math.Round(year.Criteriums[1].Rate,3), 
                    Math.Round(year.Criteriums[2].Rate,3), 
                    Math.Round(year.Criteriums[3].Rate,3), 
                    Math.Round(year.Criteriums[4].Rate,3), 
                    Math.Round(year.Criteriums[5].Rate,3), 
                    Math.Round(year.Criteriums[6].Rate,3));
                YearGrowthCriteriaStats.Add(tempRate);

                var tempArr = year.Criteriums.OrderBy(x => x.StandardRank).ToArray();
                var tempRank = new YearGrowth(year.Year,
                    tempArr[0].Rank,
                    tempArr[1].Rank,
                    tempArr[2].Rank,
                    tempArr[3].Rank,
                    tempArr[4].Rank,
                    tempArr[5].Rank,
                    tempArr[6].Rank);
                YearGrowthRateStats.Add(tempRank);

                var tempDev = new YearGrowth(year.Year,
                    Math.Pow(1 - tempArr[0].Rank, 2),
                    Math.Pow(2 - tempArr[1].Rank, 2),
                    Math.Pow(3 - tempArr[2].Rank, 2),
                    Math.Pow(4 - tempArr[3].Rank, 2),
                    Math.Pow(5 - tempArr[4].Rank, 2),
                    Math.Pow(6 - tempArr[5].Rank, 2),
                    Math.Pow(7 - tempArr[6].Rank, 2),
                    Math.Round(year.SpirmanKoef,3));
                YearSpiermanStats.Add(tempDev);

                var tempFeh = new FehnerCompare(year.Year,
                    year.FahnerPairCriteriums[0].Value,
                    year.FahnerPairCriteriums[1].Value,
                    year.FahnerPairCriteriums[2].Value,
                    year.FahnerPairCriteriums[3].Value,
                    year.FahnerPairCriteriums[4].Value,
                    year.FahnerPairCriteriums[5].Value,
                    year.FahnerPairCriteriums[6].Value,
                    year.FahnerPairCriteriums[7].Value,
                    year.FahnerPairCriteriums[8].Value,
                    year.FahnerPairCriteriums[9].Value,
                    year.FahnerPairCriteriums[10].Value,
                    year.FahnerPairCriteriums[11].Value,
                    year.FahnerPairCriteriums[12].Value,
                    year.FahnerPairCriteriums[13].Value,
                    year.FahnerPairCriteriums[14].Value,
                    year.FahnerPairCriteriums[15].Value,
                    year.FahnerPairCriteriums[16].Value,
                    year.FahnerPairCriteriums[17].Value,
                    year.FahnerPairCriteriums[18].Value,
                    year.FahnerPairCriteriums[19].Value,
                    year.FahnerPairCriteriums[20].Value,
                    year.FehnerSum,
                    Math.Round(year.FehnerKoef,3));
                FehnerStats.Add(tempFeh);

                var tempScale = new CompanyScale(year.Year, 
                    Math.Round(year.SpirmanKoef, 3), 
                    Math.Round(year.FehnerKoef, 3), 
                    Math.Round(year.ScaleLevel, 3));
                ScaleStats.Add(tempScale);
            }

        }

        private void GenerateCostGraph()
        {
            List<double> y = new List<double>();
            foreach (var year in YearStats)
            {
                y.Add(year.Cost);
            }

            var xVal = Enumerable.Range((int)YearStats.First().Year, YearStats.Count).ToList();

            // Create data sources:
            var xDataSource = xVal.AsXDataSource();
            var yDataSource = y.AsYDataSource();

            yDataSource.SetYMapping(Y => Y);
            xDataSource.SetXMapping(X => X);

            yDataSource.AddMapping(ShapeElementPointMarker.ToolTipTextProperty,
               Y => string.Format("Собівартість - {0}", Y));

            // plotter.Viewport.Restrictions.Add(new PhysicalProportionsRestriction { ProportionRatio = 1 });


            //CompositeDataSource 
            CompositeDataSource compositeDataSource = new CompositeDataSource(xDataSource, yDataSource);
            //xDataSource.Join(yDataSource);
            // adding graph to plotter
            plotter.AddLineGraph(compositeDataSource,
                new Pen(Brushes.Goldenrod, 3),
                new SampleMarker(),
                new PenDescription("Чст."));

            plotter.Legend.Visibility = System.Windows.Visibility.Collapsed;
            // Force evertyhing plotted to be visible
            plotter.FitToView();
            plotter.InvalidateVisual();
        }

        private void ExportToExcelButtonClick(object sender, RoutedEventArgs e)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog()
            {
                Filter = "Excel Worksheets 2003 (*.xls)|*.xls|,Excel Worksheets 2007 (*.xlsx)|*.xlsx"
            };

            if (saveDialog.ShowDialog().GetValueOrDefault())
            {
                try
                {
                    ExcelExporter.Export(new ExcelExportData
                    {
                        YearGrowthData = YearStats,
                        YearGrowthCriteriaData = YearGrowthCriteriaStats,
                        YearGrowthRateData = YearGrowthRateStats,
                        FehnerData = FehnerStats,
                        ScaleRateData = ScaleStats,
                        YearSpirmanData = YearSpiermanStats
                    }, saveDialog.FileName);
                }
                catch
                {
                    MessageBox.Show("Неможливо записати файл, поки він відкритий. Закрийте файл та спробуйте знову");
                }
            }
        }
    }
}
