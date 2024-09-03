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

    public class Input_output
    {
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public string Value { get; set; }
        public string Range { get; set; }
        public string Format { get; set; }
    }
}
