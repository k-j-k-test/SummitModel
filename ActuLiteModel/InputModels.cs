using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
    public class Input_mp
    {
        public int ID { get; set; }
        public string ProductCode { get; set; }
        public string RiderCode { get; set; }
        public string Model { get; set; }
        public int Age { get; set; }
        public int n { get; set; }
        public int m { get; set; }
        public double SA { get; set; }
        public int Freq { get; set; }
        public int F1 { get; set; }
        public int F2 { get; set; }
        public int F3 { get; set; }
        public int F4 { get; set; }
        public int F5 { get; set; }
        public int F6 { get; set; }
        public int F7 { get; set; }
        public int F8 { get; set; }
        public int F9 { get; set; }
        public string A1 { get; set; }
        public string A2 { get; set; }
        public string A3 { get; set; }
        public string A4 { get; set; }
        public string A5 { get; set; }
        public string A6 { get; set; }
        public string A7 { get; set; }
        public string A8 { get; set; }
        public string A9 { get; set; }
        public DateTime D1 { get; set; }
        public DateTime D2 { get; set; }
        public int t0 { get; set; }
    }

    public class Input_cell
    {
        public string Model { get; set; }
        public string Cell { get; set; }
        public string Formula { get; set; }
        public string Description { get; set; }
    }

    public class Input_assum
    {
        public string Key1 { get; set; }
        public string Key2 { get; set; }
        public string Key3 { get; set; }
        public string Condition { get; set; }
        public List<double> Rates { get; set; }
    }

    public class Input_setting
    {
        public string Property { get; set; }
        public int Type { get; set; }
        public string Value { get; set; }
    }

    public class Input_table
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string Delimiter { get; set; }
        public string Key { get; set; }
        public string Type1 { get; set; }
        public string Type2 { get; set; }
        public string Type3 { get; set; }
        public string Type4 { get; set; }
        public string Type5 { get; set; }
        public string Item1 { get; set; }
        public string Item2 { get; set; }
        public string Item3 { get; set; }
        public string Item4 { get; set; }
        public string Item5 { get; set; }
    }
}
