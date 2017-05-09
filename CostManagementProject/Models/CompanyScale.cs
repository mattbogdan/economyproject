using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostManagementProject.Models
{
    public class CompanyScale
    {
        public double Year { get; set; }
        public double SpirmanCoef { get; set; }
        public double FehnerCoef { get; set; }
        public double ScaleRate { get; set; }

        public CompanyScale(double year, double spirmanCoef, double fehnerCoef, double scaleRate)
        {
            Year = year;
            SpirmanCoef = spirmanCoef;
            FehnerCoef = fehnerCoef;
            ScaleRate = scaleRate;
        }        
    }
}
