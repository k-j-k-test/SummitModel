using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Caching;
using Flee.PublicTypes;
using System.Windows;

namespace ActuLiteModel
{
    public class Model : BaseModel
    {
        public string Name { get; set; }
        public ModelEngine Engine { get; set; }
        public Parameter Parameter { get; set; }
        public Dictionary<string, Sheet> Sheets { get; set; }
        public Dictionary<string, CompiledCell> CompiledCells { get; set; }

        public Model(string name, ModelEngine engine)
        {
            Name = name;
            Engine = engine;
            Parameter = new Parameter();
            Sheets = new Dictionary<string, Sheet>();
            CompiledCells = new Dictionary<string, CompiledCell>();
        }

        public void AddSheet(string sheetName)
        {
            if (!Sheets.ContainsKey(sheetName))
            {
                Sheets[sheetName] = new Sheet();

                foreach (var compiledCell in CompiledCells)
                {
                    Sheets[sheetName].RegisterMethod(compiledCell.Key, compiledCell.Value.CellFunc);
                }
            }
        }

        public void ResisterCell(string name, string formula, string description)
        {
            CompiledCell compiledCell = new CompiledCell(name, formula, description, this);
            CompiledCells[name] = compiledCell;
            //CompiledCells[name].GetTest();
        }

        public void Invoke(string cellName, int t)
        {
            AddSheet(Name);
            t = Math.Min(t, FleeFunc.Max_T);
            Engine.Context.Variables["t"] = t;
            Sheets[Name][cellName, t] = Sheets[Name].GetMethod(cellName)(t);
            Sheets[Name].ClearCalculationStack();
        }

    }

    public class BaseModel
    {
        public T If<T>(bool condition, T truevalue, T falsevalue)
        {
            if (condition) return truevalue;
            else return falsevalue;
        }
    }

    public class CompiledCell
    {
        public string Name { get; set; }
        public string Formula { get; set; }
        public string Description { get; set; }
        public Model Model { get; set; }
        public IGenericExpression<double> Expression { get; set; }
        public Func<int, double> CellFunc { get; set; }

        public bool IsCompiled { get; set; }
        public string CompileStatusMessage { get; set; }

        public CompiledCell(string name, string formula, string description, Model model)
        {
            Name = name;
            Formula = formula;
            Description = description;
            Model = model;
            Compile();         
        }

        private void Compile()
        {
            try
            {
                string transFormedFormula = ModelEngine.TransformText(Formula, Model.Name);
                Expression = Model.Engine.Context.CompileGeneric<double>(transFormedFormula);
                CellFunc = t => Expression.Evaluate();
                IsCompiled = true;
                CompileStatusMessage = "Successfully Compiled";
            }
            catch(Exception ex)
            {
                IsCompiled = false;
                CompileStatusMessage = ex.Message;
            }
        }

        public void GetTest()
        {
            try
            {
                Model.Sheets.Clear();
                Model.Invoke(Name, 0);
                Model.Sheets.Clear();
                Model.Invoke(Name, 99);
                Model.Sheets.Clear();
            }
            catch (Exception ex)
            {
                this.IsCompiled = false;
                this.CompileStatusMessage = ex.Message;
            }
        }
    }

}
