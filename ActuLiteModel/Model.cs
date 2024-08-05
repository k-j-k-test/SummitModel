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

        // 매우 큰 값의 임계점을 정의
        private const double MaxAllowedValue = 99999999999999;

        public void Invoke(string cellName, int t)
        {
            try
            {
                AddSheet(Name);
                t = Math.Min(t, Sheet.MaxT);
                Engine.Context.Variables["t"] = t;

                if (!Sheets.TryGetValue(Name, out var sheet))
                {
                    throw new InvalidOperationException($"시트 '{Name}'을 찾을 수 없습니다.");
                }

                var method = sheet.GetMethod(cellName);
                if (method == null)
                {
                    throw new ArgumentException($"셀 '{cellName}'에 대한 메서드를 찾을 수 없습니다.");
                }

                double result = method(t);

                // 결과 값이 무한대이거나 허용 범위를 초과하는 경우 OverflowException 발생
                if (double.IsInfinity(result) || Math.Abs(result) > MaxAllowedValue)
                {
                    throw new OverflowException($"계산 결과가 허용 범위를 초과했습니다: 셀 {cellName}, t={t}, 결과={result}");
                }

                sheet[cellName, t] = result;

                foreach(var sh in Sheets)
                {
                    sh.Value.ChangeCellOrder(Sheet.SortOption.FirstCalculationTime);
                }
                
            }
            catch (CircularReferenceException ex)
            {
                throw new CircularReferenceException($"순환 참조 감지: 셀 {cellName}, t={t}. 경로: {ex.Message}");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new ArgumentOutOfRangeException($"범위 초과 오류: 셀 {cellName}, t={t}. 오류: {ex.Message}", ex);
            }
            catch (OverflowException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"예상치 못한 오류 발생: 셀 {cellName}, t={t}. 오류: {ex.Message}", ex);
            }
        }

        public void Clear()
        {
            Parameter = new Parameter();
            Sheets = new Dictionary<string, Sheet>();
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
