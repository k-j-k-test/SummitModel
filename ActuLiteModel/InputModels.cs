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

    public class Input_exp
    {
        public string ProductCode { get; set; }
        public string RiderCode { get; set; }
        public string Condition { get; set; }
        public string Alpha_P { get; set; }
        public string Alpha_P2 { get; set; }
        public string Alpha_S { get; set; }
        public string Alpha_P20 { get; set; }
        public string Beta_P { get; set; }
        public string Beta_S { get; set; }
        public string Beta_Fix { get; set; }
        public string BetaPrime_P { get; set; }
        public string BetaPrime_S { get; set; }
        public string BetaPrime_Fix { get; set; }
        public string Gamma { get; set; }
        public string Refund_P { get; set; }
        public string Refund_S { get; set; }
        public string Etc1 { get; set; }
        public string Etc2 { get; set; }
        public string Etc3 { get; set; }
        public string Etc4 { get; set; }
    }

    public class Input_output
    {
        public string Table { get; set; }
        public string ProductCode { get; set; }
        public string RiderCode { get; set; }
        public string Value { get; set; }
        public string Position { get; set; }
        public string Range { get; set; }
        public string Format { get; set; }
    }
}
