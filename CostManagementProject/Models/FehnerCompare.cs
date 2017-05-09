using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostManagementProject.Models
{
    public class FehnerCompare
    {
        public double Year { get; set; }
        public int First { get; set; }
        public int Second { get; set; }
        public int Third { get; set; }
        public int Fourth { get; set; }
        public int Fifth { get; set; }
        public int Sixth { get; set; }
        public int Seventh { get; set; }
        public int Eighth { get; set; }
        public int Ninth { get; set; }
        public int Tenth { get; set; }
        public int Eleventh { get; set; }
        public int Twelfth { get; set; }
        public int Thirteenth { get; set; }
        public int Fourteenth { get; set; }
        public int Fifteenth { get; set; }
        public int Sixteenth { get; set; }
        public int Seventeenth { get; set; }
        public int Eighteenth { get; set; }
        public int Nineteenth { get; set; }
        public int Twentieth { get; set; }
        public int TwentyFirst { get; set; }
        public int RangSum { get; set; }
        public double FahnerCoef { get; set; }

        public FehnerCompare(double year ,int first, int second, int third, int fourth, int fifth, int sixth, int seventh, int eighth, int ninth, int tenth, int eleventh, int twelfth, int thirteenth, int fourteenth, int fifteenth, int sixteenth, int seventeenth, int eighteenth, int nineteenth, int twentieth, int twentyFirst, int rangSum, double fahnerCoef)
        {
            Year = year;
            First = first;
            Second = second;
            Third = third;
            Fourth = fourth;
            Fifth = fifth;
            Sixth = sixth;
            Seventh = seventh;
            Eighth = eighth;
            Ninth = ninth;
            Tenth = tenth;
            Eleventh = eleventh;
            Twelfth = twelfth;
            Thirteenth = thirteenth;
            Fourteenth = fourteenth;
            Fifteenth = fifteenth;
            Sixteenth = sixteenth;
            Seventeenth = seventeenth;
            Eighteenth = eighteenth;
            Nineteenth = nineteenth;
            Twentieth = twentieth;
            TwentyFirst = twentyFirst;
            RangSum = rangSum;
            FahnerCoef = fahnerCoef;
        }
    }
}
