using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostManagementProject.Models
{
    public class YearGrowthCriteries
    {
        public YearGrowthCriteries()
        {
            Criteriums = new List<Criterium>();
            FahnerPairCriteriums = new List<FahnerPairCriterium>();
        }

        public double Year { get; set; }

        public List<Criterium> Criteriums { get; set; }

        public double GetCriteriumsDeviationsSum()//сума квадратів рангових відхилень
        {
            double sum = 0;

            foreach (var criterium in Criteriums)
            {
                sum += criterium.GetRankDeviation();
            }

            return sum;
        }

        public List<FahnerPairCriterium> FahnerPairCriteriums { get; set; }

        public int FehnerSum { get; set; }

        public double FehnerKoef { get; set; }

        public double SpirmanKoef { get; set; }

        public double ScaleLevel { get; set; }


    }

}
