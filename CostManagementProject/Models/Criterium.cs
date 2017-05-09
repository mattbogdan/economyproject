using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostManagementProject.Models
{
    public class Criterium
    {
        public int Id { get; set; }

        public double Value { get; set; }

        public double Rate { get; set; }

        public int StandardRank { get; set; }

        public int Rank { get; set; }

        public double GetRankDeviation()//квадрат рангових відхилень
        {
            return Math.Pow(Rank - StandardRank, 2);
        }

    }
}
