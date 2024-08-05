using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActuLiteModel
{
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
